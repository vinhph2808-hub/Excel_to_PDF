# -*- coding: utf-8 -*-
"""
Công Cụ Chuyển Đổi Excel Sang PDF Chuyên Nghiệp (Excel To PDF Converter Pro)
Tích hợp:
- Giao diện Tkinter hiện đại, tinh tế (Phong cách Emerald Office).
- Đọc danh sách sheet động, hỗ trợ: Toàn bộ workbook, Từng sheet riêng, hoặc Chọn sheet cụ thể.
- Tùy chọn in ấn chuyên nghiệp (Professional Page Setup):
    + Tự động vừa vặn 1 trang ngang (FitToPagesWide = 1) chống tràn cột.
    + Căn giữa trang theo chiều ngang (CenterHorizontally).
    + Khổ giấy chuẩn (A4 / Letter / Giữ nguyên).
    + Hướng giấy linh hoạt (Dọc / Ngang / Tự động).
    + Căn lề chuẩn (Normal / Narrow / Tùy chỉnh).
- Độ ổn định cao nhất (kế thừa từ excel_to_pdf_windows):
    + Copy sang file tạm, không can thiệp hay lưu đè file gốc.
    + Mở an toàn file bị chặn Protected View / Trust Center.
    + Bọc an toàn thuộc tính Excel COM, chạy ngầm tuyệt đối.
    + Tự động đóng tiến trình Excel sạch sẽ sau khi hoàn tất.
"""

from __future__ import annotations

import datetime as dt
import gc
import os
import queue
import re
import shutil
import tempfile
import threading
import traceback
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Định dạng file hỗ trợ
SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}

# Chế độ xuất
MODE_WORKBOOK = "workbook"
MODE_EACH_SHEET = "each_sheet"
MODE_SELECTED_SHEETS = "selected_sheets"

# Hằng số Excel COM
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_SHEET_VISIBLE = -1
XL_CALCULATION_MANUAL = -4135
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# Hằng số PageSetup
XL_PORTRAIT = 1
XL_LANDSCAPE = 2
XL_PAPER_LETTER = 1
XL_PAPER_A4 = 9

LogFunc = Callable[[str], None]
ProgressFunc = Callable[[int, int], None]

INVALID_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}


@dataclass(frozen=True)
class SheetInfo:
    name: str
    visible: bool


@dataclass
class PageSetupOptions:
    """Các tùy chọn trang in giúp file PDF xuất ra chỉnh chu, chuyên nghiệp."""
    enabled: bool = False
    fit_to_width: bool = True           # Khớp 1 trang chiều ngang, không ngắt cột
    center_horizontally: bool = True     # Căn giữa bảng tính trên trang giấy
    paper_size: str = "A4"               # "A4", "Letter", "Keep"
    orientation: str = "Keep"            # "Keep", "Portrait", "Landscape"
    margin_mode: str = "Keep"            # "Keep", "Normal", "Narrow"


def safe_filename(value: str, fallback: str = "output", max_length: int = 120) -> str:
    """Tạo tên file an toàn cho Windows nhưng vẫn giữ nguyên tiếng Việt có dấu."""
    text = str(value or "").strip()
    text = INVALID_FILENAME_RE.sub("_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")

    if not text:
        text = fallback

    base_name = text.split(".")[0].upper()
    if base_name in WINDOWS_RESERVED_NAMES:
        text = f"_{text}"

    if len(text) > max_length:
        text = text[:max_length].rstrip(" .")

    return text or fallback


def unique_path(path: Path) -> Path:
    """Tránh ghi đè file có sẵn bằng cách thêm _2, _3... vào tên file."""
    path = Path(path)
    if not path.exists():
        return path

    for index in range(2, 10000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate

    raise RuntimeError(f"Không thể tạo tên file không trùng cho: {path}")


def format_error(exc: BaseException) -> str:
    """Rút gọn thông báo lỗi COM/Python để hiển thị rõ ràng trên log."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple):
        description = args[2][2]
        if description:
            return str(description).strip()

    text = str(exc).strip()
    return text or exc.__class__.__name__


def clear_mark_of_the_web(path: Path) -> None:
    """Xóa nhãn Mark-of-the-Web (Zone.Identifier) trên file tạm tải từ Zalo/Web."""
    if os.name != "nt":
        return
    try:
        os.remove(f"{path}:Zone.Identifier")
    except OSError:
        pass


def copy_file_to_temp(source: Path, temp_root: Path) -> Path:
    """Sao chép Excel sang thư mục tạm để không can thiệp hoặc khóa file gốc."""
    safe_stem = safe_filename(source.stem, fallback="excel", max_length=80)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    destination = Path(temp_root) / f"{safe_stem}_{timestamp}{source.suffix.lower()}"
    shutil.copy2(str(source), str(destination))
    clear_mark_of_the_web(destination)
    return destination


def import_pywin32_modules():
    try:
        import pythoncom  # type: ignore
        import win32com.client as win32_client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Chưa cài đặt pywin32. Vui lòng chạy lệnh: pip install pywin32"
        ) from exc
    return pythoncom, win32_client


class ExcelSession:
    """Quản lý phiên làm việc Excel COM an toàn, cô lập và tự dọn dẹp sạch sẽ."""
    def __init__(self, log: LogFunc):
        self.log = log
        self.pythoncom = None
        self.excel = None

    def __enter__(self):
        self.pythoncom, win32_client = import_pywin32_modules()
        self.pythoncom.CoInitialize()

        try:
            self.excel = win32_client.DispatchEx("Excel.Application")
            self._optimize_excel()
            return self
        except Exception:
            self.pythoncom.CoUninitialize()
            raise

    @property
    def app(self):
        if self.excel is None:
            raise RuntimeError("Phiên Excel COM chưa sẵn sàng.")
        return self.excel

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            self._close_leftover_workbooks()
        finally:
            if self.excel is not None:
                try:
                    self.excel.ScreenUpdating = True
                except Exception:
                    pass
                try:
                    self.excel.Interactive = True
                except Exception:
                    pass
                try:
                    self.excel.Quit()
                except Exception as quit_error:
                    self.log(f"Cảnh báo khi đóng Excel: {format_error(quit_error)}")

            self.excel = None
            gc.collect()

            if self.pythoncom is not None:
                self.pythoncom.CoUninitialize()

    def _optimize_excel(self) -> None:
        """Thiết lập Excel chạy ẩn an toàn (bọc try-except đề phòng lỗi trên một số máy)."""
        attrs = [
            ("Visible", False),
            ("DisplayAlerts", False),
            ("ScreenUpdating", False),
            ("EnableEvents", False),
            ("AskToUpdateLinks", False),
            ("Interactive", False),
            ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
        ]
        for name, value in attrs:
            try:
                setattr(self.excel, name, value)
            except Exception:
                pass

    def _close_leftover_workbooks(self) -> None:
        if self.excel is None:
            return
        try:
            while self.excel.Workbooks.Count > 0:
                self.excel.Workbooks(1).Close(SaveChanges=False)
        except Exception as exc:
            self.log(f"Cảnh báo dọn dẹp workbook: {format_error(exc)}")


def close_workbook(workbook, log: LogFunc) -> None:
    if workbook is None:
        return
    try:
        workbook.Close(SaveChanges=False)
    except Exception as exc:
        log(f"Cảnh báo khi đóng workbook: {format_error(exc)}")


def open_workbook_safely(excel, temp_path: Path, log: LogFunc):
    """Mở workbook với cơ chế fallback Protected View -> Edit an toàn."""
    file_name = str(temp_path)

    try:
        return excel.Workbooks.Open(
            Filename=file_name,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Notify=False,
        )
    except Exception as direct_error:
        direct_message = format_error(direct_error)
        log("Không mở trực tiếp được, đang thử qua ProtectedViewWindows.Open() rồi Edit()...")

    protected_view_window = None
    try:
        protected_view_window = excel.ProtectedViewWindows.Open(file_name)
        workbook = protected_view_window.Edit()
        log("Đã mở thành công qua chế độ Protected View.")
        return workbook
    except Exception as protected_error:
        if protected_view_window is not None:
            try:
                protected_view_window.Close()
            except Exception:
                pass

        raise RuntimeError(
            "Không thể mở file Excel. File có thể bị khóa mật khẩu, lỗi định dạng hoặc chặn bởi chính sách hệ thống. "
            f"Chi tiết: {direct_message} | Protected View: {format_error(protected_error)}"
        ) from protected_error


def set_manual_calculation_if_possible(excel) -> None:
    try:
        excel.Calculation = XL_CALCULATION_MANUAL
    except Exception:
        pass


def is_sheet_visible(sheet) -> bool:
    try:
        return int(sheet.Visible) == XL_SHEET_VISIBLE
    except Exception:
        return False


def get_sheet_infos(workbook) -> List[SheetInfo]:
    infos: List[SheetInfo] = []
    for index in range(1, workbook.Sheets.Count + 1):
        sheet = workbook.Sheets(index)
        infos.append(SheetInfo(name=str(sheet.Name), visible=is_sheet_visible(sheet)))
    return infos


def read_sheet_catalog(file_paths: Iterable[str], log: LogFunc) -> List[SheetInfo]:
    """Đọc danh mục sheet từ các file được chọn bằng COM."""
    catalog: "OrderedDict[str, SheetInfo]" = OrderedDict()

    with tempfile.TemporaryDirectory(prefix="ExcelToPDF_sheets_") as temp_root:
        with ExcelSession(log) as excel_session:
            for file_path in file_paths:
                source = Path(file_path)
                if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue

                workbook = None
                try:
                    log(f"Đang đọc danh sách sheet từ: {source.name}")
                    temp_copy = copy_file_to_temp(source, Path(temp_root))
                    workbook = open_workbook_safely(excel_session.app, temp_copy, log)

                    for info in get_sheet_infos(workbook):
                        if info.name not in catalog:
                            catalog[info.name] = info
                        else:
                            old_info = catalog[info.name]
                            catalog[info.name] = SheetInfo(
                                name=info.name,
                                visible=old_info.visible or info.visible,
                            )
                except Exception as exc:
                    log(f"Lỗi khi đọc sheet từ {source.name}: {format_error(exc)}")
                finally:
                    close_workbook(workbook, log)
                    workbook = None
                    gc.collect()

    return list(catalog.values())


def apply_professional_page_setup(sheet, excel_app, options: PageSetupOptions, log: LogFunc) -> None:
    """
    Áp dụng các thiết lập trang in chuẩn chỉnh cho từng sheet:
    - Vừa vặn 1 trang ngang (không bị đứt cột)
    - Căn giữa trang
    - Khổ giấy A4 / Letter
    - Hướng giấy Dọc / Ngang
    - Căn lề
    """
    if not options.enabled:
        return

    try:
        ps = sheet.PageSetup

        # 1. Hướng giấy
        if options.orientation == "Portrait":
            try:
                ps.Orientation = XL_PORTRAIT
            except Exception:
                pass
        elif options.orientation == "Landscape":
            try:
                ps.Orientation = XL_LANDSCAPE
            except Exception:
                pass

        # 2. Khổ giấy
        if options.paper_size == "A4":
            try:
                ps.PaperSize = XL_PAPER_A4
            except Exception:
                pass
        elif options.paper_size == "Letter":
            try:
                ps.PaperSize = XL_PAPER_LETTER
            except Exception:
                pass

        # 3. Vừa khít 1 trang ngang (FitToPagesWide = 1)
        if options.fit_to_width:
            try:
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False  # Chiều dọc tự động giãn theo dữ liệu
            except Exception as e:
                log(f"   [!] Lưu ý căn trang Fit: {format_error(e)}")

        # 4. Căn giữa trang theo chiều ngang
        if options.center_horizontally:
            try:
                ps.CenterHorizontally = True
            except Exception:
                pass

        # 5. Căn lề
        if options.margin_mode == "Normal":
            try:
                # Lề tiêu chuẩn (Trái 1.8cm, Phải 1.8cm, Trên 2cm, Dưới 2cm)
                ps.LeftMargin = excel_app.CentimetersToPoints(1.8)
                ps.RightMargin = excel_app.CentimetersToPoints(1.8)
                ps.TopMargin = excel_app.CentimetersToPoints(2.0)
                ps.BottomMargin = excel_app.CentimetersToPoints(2.0)
            except Exception:
                pass
        elif options.margin_mode == "Narrow":
            try:
                # Lề hẹp (Trái 0.8cm, Phải 0.8cm, Trên 1.0cm, Dưới 1.0cm)
                ps.LeftMargin = excel_app.CentimetersToPoints(0.8)
                ps.RightMargin = excel_app.CentimetersToPoints(0.8)
                ps.TopMargin = excel_app.CentimetersToPoints(1.0)
                ps.BottomMargin = excel_app.CentimetersToPoints(1.0)
            except Exception:
                pass

    except Exception as exc:
        log(f"   [!] Không thể thiết lập trang in cho sheet '{sheet.Name}': {format_error(exc)}")


def export_as_pdf(com_object, pdf_path: Path) -> None:
    com_object.ExportAsFixedFormat(
        Type=XL_TYPE_PDF,
        Filename=str(pdf_path),
        Quality=XL_QUALITY_STANDARD,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )


def workbook_pdf_path(output_dir: Path, excel_path: Path) -> Path:
    workbook_name = safe_filename(excel_path.stem, fallback="Workbook", max_length=150)
    return unique_path(output_dir / f"{workbook_name}.pdf")


def sheet_pdf_path(output_dir: Path, excel_path: Path, sheet_name: str) -> Path:
    workbook_name = safe_filename(excel_path.stem, fallback="Workbook", max_length=70)
    safe_sheet = safe_filename(sheet_name, fallback="Sheet", max_length=70)
    return unique_path(output_dir / f"{workbook_name} - {safe_sheet}.pdf")


def export_workbook_to_pdf(
    workbook,
    excel_app,
    excel_path: Path,
    output_dir: Path,
    options: PageSetupOptions,
    log: LogFunc,
) -> List[Path]:
    # Áp dụng page setup cho toàn bộ sheet hiển thị nếu có bật
    if options.enabled:
        for sheet in workbook.Worksheets:
            if is_sheet_visible(sheet):
                apply_professional_page_setup(sheet, excel_app, options, log)

    hidden_sheets = [info.name for info in get_sheet_infos(workbook) if not info.visible]
    if hidden_sheets:
        log("Sheet ẩn sẽ được Excel bỏ qua: " + ", ".join(hidden_sheets))

    pdf_path = workbook_pdf_path(output_dir, excel_path)
    log(f"Xuất toàn bộ workbook -> {pdf_path.name}")
    export_as_pdf(workbook, pdf_path)
    log(f"✓ Đã xuất PDF thành công: {pdf_path.name}")
    return [pdf_path]


def export_each_sheet_to_pdf(
    workbook,
    excel_app,
    excel_path: Path,
    output_dir: Path,
    options: PageSetupOptions,
    log: LogFunc,
) -> List[Path]:
    pdf_paths: List[Path] = []

    for index in range(1, workbook.Sheets.Count + 1):
        sheet = workbook.Sheets(index)
        sheet_name = str(sheet.Name)

        if not is_sheet_visible(sheet):
            log(f"Bỏ qua sheet ẩn: {sheet_name}")
            continue

        try:
            if options.enabled:
                apply_professional_page_setup(sheet, excel_app, options, log)

            pdf_path = sheet_pdf_path(output_dir, excel_path, sheet_name)
            export_as_pdf(sheet, pdf_path)
            pdf_paths.append(pdf_path)
            log(f"✓ Xuất sheet '{sheet_name}' -> {pdf_path.name}")
        except Exception as exc:
            log(f"Lỗi xuất sheet '{sheet_name}': {format_error(exc)}")

    return pdf_paths


def export_selected_sheets_to_pdf(
    workbook,
    excel_app,
    excel_path: Path,
    output_dir: Path,
    selected_sheet_names: Iterable[str],
    options: PageSetupOptions,
    log: LogFunc,
) -> List[Path]:
    pdf_paths: List[Path] = []
    selected_lookup = {name.casefold(): name for name in selected_sheet_names}
    seen_keys = set()

    for index in range(1, workbook.Sheets.Count + 1):
        sheet = workbook.Sheets(index)
        sheet_name = str(sheet.Name)
        sheet_key = sheet_name.casefold()

        if sheet_key not in selected_lookup:
            continue

        seen_keys.add(sheet_key)

        if not is_sheet_visible(sheet):
            log(f"Sheet '{sheet_name}' đang ẩn nên được bỏ qua.")
            continue

        try:
            if options.enabled:
                apply_professional_page_setup(sheet, excel_app, options, log)

            pdf_path = sheet_pdf_path(output_dir, excel_path, sheet_name)
            export_as_pdf(sheet, pdf_path)
            pdf_paths.append(pdf_path)
            log(f"✓ Xuất sheet '{sheet_name}' -> {pdf_path.name}")
        except Exception as exc:
            log(f"Lỗi xuất sheet '{sheet_name}': {format_error(exc)}")

    for selected_name in selected_sheet_names:
        if selected_name.casefold() not in seen_keys:
            log(f"Không tìm thấy sheet '{selected_name}' trong file này.")

    return pdf_paths


def compress_pdfs_to_zip(
    pdf_paths: Iterable[Path],
    output_dir: Path,
    excel_paths: Iterable[Path],
    log: LogFunc,
) -> Optional[Path]:
    pdf_list = [Path(path) for path in pdf_paths if Path(path).exists()]
    excel_list = list(excel_paths)

    if not pdf_list:
        log("Không có file PDF nào để nén ZIP.")
        return None

    if len(excel_list) == 1:
        zip_base = safe_filename(excel_list[0].stem, fallback="ExcelToPDF", max_length=120)
        zip_name = f"{zip_base}_PDF.zip"
    else:
        timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f"ExcelToPDF_{timestamp}.zip"

    zip_path = unique_path(output_dir / zip_name)
    log(f"Đang nén các file PDF vào tệp ZIP: {zip_path.name}...")

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as zip_file:
        for pdf_path in pdf_list:
            zip_file.write(pdf_path, arcname=pdf_path.name)

    log(f"✓ Đã tạo thành công file ZIP: {zip_path}")
    return zip_path


def convert_files(
    file_paths: Iterable[str],
    output_dir: Optional[Path],
    mode: str,
    selected_sheet_names: Iterable[str],
    page_options: PageSetupOptions,
    compress_after_export: bool,
    log: LogFunc,
    progress: ProgressFunc,
) -> Tuple[List[Path], Optional[Path]]:
    all_files = [Path(path) for path in file_paths]
    valid_files = [path for path in all_files if path.suffix.lower() in SUPPORTED_EXTENSIONS]

    for path in all_files:
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            log(f"Bỏ qua file không đúng định dạng Excel: {path}")

    if not valid_files:
        raise RuntimeError("Không có file Excel hợp lệ (.xlsx, .xls, .xlsm).")

    selected_sheet_list = list(selected_sheet_names)
    exported_pdfs: List[Path] = []
    explicit_output_dir = Path(output_dir) if output_dir is not None else None

    if explicit_output_dir is not None:
        explicit_output_dir.mkdir(parents=True, exist_ok=True)
    else:
        log("Lưu ý: PDF sẽ được lưu cùng thư mục với từng file Excel nguồn.")

    with tempfile.TemporaryDirectory(prefix="ExcelToPDF_") as temp_root:
        with ExcelSession(log) as excel_session:
            total = len(valid_files)

            for position, source in enumerate(valid_files, start=1):
                workbook = None
                progress(position - 1, total)

                try:
                    log(f"[{position}/{total}] Đang xử lý: {source.name}")
                    temp_copy = copy_file_to_temp(source, Path(temp_root))
                    workbook = open_workbook_safely(excel_session.app, temp_copy, log)
                    set_manual_calculation_if_possible(excel_session.app)
                    current_output_dir = explicit_output_dir or source.parent

                    if mode == MODE_WORKBOOK:
                        exported_pdfs.extend(
                            export_workbook_to_pdf(
                                workbook, excel_session.app, source, current_output_dir, page_options, log
                            )
                        )
                    elif mode == MODE_EACH_SHEET:
                        exported_pdfs.extend(
                            export_each_sheet_to_pdf(
                                workbook, excel_session.app, source, current_output_dir, page_options, log
                            )
                        )
                    elif mode == MODE_SELECTED_SHEETS:
                        if not selected_sheet_list:
                            raise RuntimeError("Bạn chưa chọn sheet cụ thể nào để xuất.")
                        exported_pdfs.extend(
                            export_selected_sheets_to_pdf(
                                workbook,
                                excel_session.app,
                                source,
                                current_output_dir,
                                selected_sheet_list,
                                page_options,
                                log,
                            )
                        )
                    else:
                        raise RuntimeError(f"Chế độ xuất không hợp lệ: {mode}")
                except Exception as exc:
                    log(f"✕ Lỗi khi xử lý {source.name}: {format_error(exc)}")
                    log("Chi tiết lỗi:\n" + traceback.format_exc(limit=5))
                finally:
                    close_workbook(workbook, log)
                    workbook = None
                    gc.collect()
                    progress(position, total)

    zip_path = None
    if compress_after_export and exported_pdfs:
        zip_output_dir = explicit_output_dir or valid_files[0].parent
        if explicit_output_dir is None and len({path.parent for path in valid_files}) > 1:
            log(f"File ZIP tổng được lưu tại thư mục của file đầu tiên: {zip_output_dir}")
        zip_path = compress_pdfs_to_zip(exported_pdfs, zip_output_dir, valid_files, log)

    return exported_pdfs, zip_path


# ==============================================================================
# GIAO DIỆN NGƯỜI DÙNG TKINTER HIỆN ĐẠI & CHUYÊN NGHIỆP
# ==============================================================================

class ExcelToPdfApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel To PDF Converter Pro - Chuyển Đổi Excel Sang PDF Chuyên Nghiệp")
        self.root.geometry("860x720")
        self.root.minsize(820, 650)

        # Màu sắc chủ đạo (Emerald Office Theme)
        self.COLOR_PRIMARY = "#0f7b52"       # Xanh ngọc lục bảo (Excel)
        self.COLOR_PRIMARY_HOVER = "#0b5f3f"
        self.COLOR_BG = "#f8f9fa"
        self.COLOR_CARD = "#ffffff"
        self.COLOR_TEXT = "#212529"
        self.COLOR_MUTED = "#6c757d"
        self.COLOR_BORDER = "#dee2e6"

        self.files: List[str] = []
        self.sheet_vars: Dict[str, tk.BooleanVar] = {}
        self.sheet_checkbuttons: List[ttk.Checkbutton] = []
        self.queue: "queue.Queue[Tuple[object, ...]]" = queue.Queue()

        # Biến điều khiển
        self.output_dir_var = tk.StringVar(value="")
        self.save_same_folder_var = tk.BooleanVar(value=True)
        self.mode_var = tk.StringVar(value=MODE_WORKBOOK)
        self.compress_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Sẵn sàng chuyển đổi.")

        # Biến PageSetup chuyên nghiệp
        self.ps_enabled_var = tk.BooleanVar(value=True)          # Mặc định bật tối ưu trang in
        self.ps_fit_width_var = tk.BooleanVar(value=True)        # Khớp 1 trang ngang
        self.ps_center_var = tk.BooleanVar(value=True)           # Căn giữa trang
        self.ps_paper_size_var = tk.StringVar(value="A4")         # Mặc định A4
        self.ps_orientation_var = tk.StringVar(value="Keep")     # Mặc định giữ nguyên theo file
        self.ps_margin_var = tk.StringVar(value="Keep")          # Mặc định giữ nguyên

        self.is_busy = False
        self.sheet_loading = False
        self.tracked_widgets: Dict[tk.Widget, str] = {}

        self._configure_styles()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.process_queue)

    def _configure_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg=self.COLOR_BG)

        # Thiết lập style ttk
        style.configure(".", font=("Segoe UI", 9), background=self.COLOR_BG, foreground=self.COLOR_TEXT)
        style.configure("Card.TFrame", background=self.COLOR_CARD, relief="flat")
        style.configure("TLabelframe", background=self.COLOR_CARD, bordercolor=self.COLOR_BORDER)
        style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground=self.COLOR_PRIMARY, background=self.COLOR_CARD)
        style.configure("TLabel", background=self.COLOR_CARD, foreground=self.COLOR_TEXT)
        style.configure("Muted.TLabel", font=("Segoe UI", 8, "italic"), foreground=self.COLOR_MUTED, background=self.COLOR_CARD)
        style.configure("HeaderTitle.TLabel", font=("Segoe UI", 14, "bold"), foreground=self.COLOR_PRIMARY, background=self.COLOR_BG)
        style.configure("HeaderSub.TLabel", font=("Segoe UI", 9), foreground=self.COLOR_MUTED, background=self.COLOR_BG)
        style.configure("TCheckbutton", background=self.COLOR_CARD)
        style.configure("TRadiobutton", background=self.COLOR_CARD)
        style.configure("TCombobox", selectbackground=self.COLOR_PRIMARY)

        # Progressbar xanh
        style.configure("Horizontal.TProgressbar", thickness=8, troughcolor="#e9ecef", background=self.COLOR_PRIMARY)

    def _track(self, widget: tk.Widget, normal_state: str = "normal") -> tk.Widget:
        self.tracked_widgets[widget] = normal_state
        return widget

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding="16 12 16 12")
        main.pack(fill="both", expand=True)

        # ----------------- HEADER -----------------
        header_frame = ttk.Frame(main)
        header_frame.pack(fill="x", pady=(0, 10))

        title_lbl = ttk.Label(header_frame, text="EXCEL TO PDF CONVERTER PRO", style="HeaderTitle.TLabel")
        title_lbl.pack(side="left")

        sub_lbl = ttk.Label(
            header_frame,
            text="• Chuyển đổi chuẩn xác qua Microsoft Excel COM • Bố cục in ấn chỉnh chu",
            style="HeaderSub.TLabel",
        )
        sub_lbl.pack(side="left", padx=(10, 0), pady=(4, 0))

        # ----------------- PHẦN 1: FILE EXCEL -----------------
        file_frame = ttk.LabelFrame(main, text=" 1. Danh Sách File Excel Cần Chuyển Đổi ", padding=10)
        file_frame.pack(fill="x", pady=(0, 8))

        btn_row = ttk.Frame(file_frame)
        btn_row.pack(fill="x", pady=(0, 6))

        self.btn_select_files = self._track(
            tk.Button(
                btn_row,
                text="📁 Chọn File Excel...",
                command=self.select_files,
                bg="#e8f5e9",
                fg=self.COLOR_PRIMARY,
                activebackground="#c8e6c9",
                activeforeground=self.COLOR_PRIMARY,
                font=("Segoe UI", 9, "bold"),
                relief="groove",
                padx=12,
                pady=4,
                cursor="hand2",
            )
        )
        self.btn_select_files.pack(side="left")

        self.btn_clear_files = self._track(
            ttk.Button(btn_row, text="Xóa tất cả", command=self.clear_all_files)
        )
        self.btn_clear_files.pack(side="left", padx=(8, 0))

        self.lbl_file_count = ttk.Label(btn_row, text="Chưa có file nào được chọn", style="Muted.TLabel")
        self.lbl_file_count.pack(side="right", pady=2)

        # Listbox danh sách file
        list_container = ttk.Frame(file_frame)
        list_container.pack(fill="x")

        self.file_listbox = tk.Listbox(
            list_container,
            height=3,
            selectmode=tk.EXTENDED,
            font=("Segoe UI", 9),
            relief="solid",
            bd=1,
            highlightthickness=0,
        )
        self._track(self.file_listbox)
        self.file_listbox.pack(side="left", fill="x", expand=True)

        file_sb = ttk.Scrollbar(list_container, orient="vertical", command=self.file_listbox.yview)
        file_sb.pack(side="right", fill="y")
        self.file_listbox.configure(yscrollcommand=file_sb.set)

        # ----------------- PHẦN 2: NƠI LƯU & CHẾ ĐỘ -----------------
        middle_grid = ttk.Frame(main)
        middle_grid.pack(fill="x", pady=(0, 8))

        # Khung Nơi lưu (Trái)
        dest_frame = ttk.LabelFrame(middle_grid, text=" 2. Thư Mục Lưu PDF ", padding=10)
        dest_frame.pack(side="left", fill="both", expand=True, padx=(0, 6))

        chk_same = ttk.Checkbutton(
            dest_frame,
            text="Lưu cùng thư mục với từng file Excel nguồn (Khuyên dùng)",
            variable=self.save_same_folder_var,
            command=self.toggle_output_dir,
        )
        chk_same.pack(anchor="w")

        self.custom_dir_row = ttk.Frame(dest_frame)
        self.custom_dir_row.pack(fill="x", pady=(6, 0))

        self.btn_browse_dir = self._track(
            ttk.Button(self.custom_dir_row, text="Chọn thư mục...", command=self.select_output_dir)
        )
        self.btn_browse_dir.pack(side="left", padx=(0, 6))

        self.lbl_dest_display = ttk.Label(self.custom_dir_row, text="Tự động theo từng file gốc", style="Muted.TLabel")
        self.lbl_dest_display.pack(side="left", fill="x", expand=True)

        self.toggle_output_dir()

        # Khung Chế độ (Phải)
        mode_frame = ttk.LabelFrame(middle_grid, text=" 3. Chế Độ Xuất PDF ", padding=10)
        mode_frame.pack(side="right", fill="both", expand=True, padx=(6, 0))

        self.rdo_wb = self._track(
            ttk.Radiobutton(
                mode_frame,
                text="Toàn bộ workbook (1 file PDF duy nhất)",
                value=MODE_WORKBOOK,
                variable=self.mode_var,
                command=self.on_mode_changed,
            )
        )
        self.rdo_wb.pack(anchor="w")

        self.rdo_each = self._track(
            ttk.Radiobutton(
                mode_frame,
                text="Từng sheet riêng lẻ (Mỗi sheet 1 file PDF)",
                value=MODE_EACH_SHEET,
                variable=self.mode_var,
                command=self.on_mode_changed,
            )
        )
        self.rdo_each.pack(anchor="w", pady=(3, 0))

        self.rdo_sel = self._track(
            ttk.Radiobutton(
                mode_frame,
                text="Chọn sheet cụ thể...",
                value=MODE_SELECTED_SHEETS,
                variable=self.mode_var,
                command=self.on_mode_changed,
            )
        )
        self.rdo_sel.pack(anchor="w", pady=(3, 0))

        # ----------------- PHẦN CHỌN SHEET ẨN/HIỆN -----------------
        self.sheet_frame = ttk.LabelFrame(main, text=" Chọn Các Sheet Cần Xuất PDF ", padding=8)
        self._build_sheet_picker()

        # ----------------- PHẦN 4: TỐI ƯU TRANG IN CHUYÊN NGHIỆP -----------------
        setup_frame = ttk.LabelFrame(main, text=" 4. Tối Ưu Bố Cục Trang In (Chuẩn Báo Cáo Chuyên Nghiệp) ", padding=10)
        setup_frame.pack(fill="x", pady=(0, 8))

        top_row = ttk.Frame(setup_frame)
        top_row.pack(fill="x")

        chk_enable_ps = ttk.Checkbutton(
            top_row,
            text="Bật bộ tối ưu trang in chuyên nghiệp (Bỏ tích nếu muốn giữ 100% Page Break thủ công của file Excel)",
            variable=self.ps_enabled_var,
            command=self.toggle_page_setup_ui,
        )
        chk_enable_ps.pack(side="left")

        self.ps_controls_frame = ttk.Frame(setup_frame)
        self.ps_controls_frame.pack(fill="x", pady=(8, 0))

        # Hàng 1 điều khiển: Fit to width, Center horizontally
        row1 = ttk.Frame(self.ps_controls_frame)
        row1.pack(fill="x", pady=(0, 4))

        self.chk_fit = ttk.Checkbutton(
            row1,
            text="Tự động vừa vặn 1 trang ngang (Fit 1 Page Wide - Không bị đứt cột)",
            variable=self.ps_fit_width_var,
        )
        self.chk_fit.pack(side="left")

        self.chk_center = ttk.Checkbutton(
            row1,
            text="Căn giữa trang theo chiều ngang (Center Horizontally)",
            variable=self.ps_center_var,
        )
        self.chk_center.pack(side="left", padx=(24, 0))

        # Hàng 2 điều khiển: Khổ giấy, Hướng giấy, Căn lề
        row2 = ttk.Frame(self.ps_controls_frame)
        row2.pack(fill="x", pady=(4, 0))

        ttk.Label(row2, text="Khổ giấy:").pack(side="left")
        self.cbo_paper = ttk.Combobox(
            row2,
            textvariable=self.ps_paper_size_var,
            values=["A4", "Letter", "Keep (Giữ nguyên)"],
            state="readonly",
            width=14,
        )
        self.cbo_paper.pack(side="left", padx=(6, 20))

        ttk.Label(row2, text="Hướng giấy:").pack(side="left")
        self.cbo_orient = ttk.Combobox(
            row2,
            textvariable=self.ps_orientation_var,
            values=["Keep (Giữ nguyên)", "Portrait (Dọc)", "Landscape (Ngang)"],
            state="readonly",
            width=16,
        )
        self.cbo_orient.pack(side="left", padx=(6, 20))

        ttk.Label(row2, text="Căn lề:").pack(side="left")
        self.cbo_margin = ttk.Combobox(
            row2,
            textvariable=self.ps_margin_var,
            values=["Keep (Giữ nguyên)", "Normal (Chuẩn)", "Narrow (Hẹp)"],
            state="readonly",
            width=16,
        )
        self.cbo_margin.pack(side="left", padx=(6, 0))

        self.toggle_page_setup_ui()

        # ----------------- NÚT BẮT ĐẦU & TÙY CHỌN NÉN -----------------
        action_row = ttk.Frame(main)
        action_row.pack(fill="x", pady=(2, 8))

        self.chk_compress = ttk.Checkbutton(
            action_row,
            text="Nén tất cả file PDF thành 1 tệp .ZIP sau khi xuất",
            variable=self.compress_var,
        )
        self.chk_compress.pack(side="left", pady=8)

        self.btn_start = self._track(
            tk.Button(
                action_row,
                text="🚀 BẮT ĐẦU CHUYỂN ĐỔI SANG PDF",
                command=self.start_conversion,
                bg=self.COLOR_PRIMARY,
                fg="#ffffff",
                activebackground=self.COLOR_PRIMARY_HOVER,
                activeforeground="#ffffff",
                disabledforeground="#c8e6c9",
                font=("Segoe UI", 11, "bold"),
                padx=24,
                pady=10,
                relief="flat",
                cursor="hand2",
            )
        )
        self.btn_start.pack(side="right")

        # ----------------- TIẾN ĐỘ & LOG -----------------
        prog_row = ttk.Frame(main)
        prog_row.pack(fill="x", pady=(0, 6))

        self.progress_bar = ttk.Progressbar(prog_row, style="Horizontal.TProgressbar", mode="determinate")
        self.progress_bar.pack(side="left", fill="x", expand=True)

        self.lbl_status = ttk.Label(prog_row, textvariable=self.status_var, font=("Segoe UI", 9, "bold"), foreground=self.COLOR_PRIMARY)
        self.lbl_status.pack(side="right", padx=(12, 0))

        # Log box
        log_frame = ttk.LabelFrame(main, text=" Nhật Ký Xử Lý (Log) ", padding=8)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            log_frame,
            height=6,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
            bg="#ffffff",
            fg="#333333",
            relief="solid",
            bd=1,
            highlightthickness=0,
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        log_sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_sb.set)

    def _build_sheet_picker(self) -> None:
        toolbar = ttk.Frame(self.sheet_frame)
        toolbar.pack(fill="x", pady=(0, 4))

        self.btn_reload_sheets = self._track(
            ttk.Button(toolbar, text="🔄 Tải lại sheet", command=self.load_sheets_async)
        )
        self.btn_reload_sheets.pack(side="left")

        self.btn_sel_all_sheets = self._track(
            ttk.Button(toolbar, text="Chọn tất cả", command=self.select_all_sheets)
        )
        self.btn_sel_all_sheets.pack(side="left", padx=(6, 0))

        self.btn_clear_sheets = self._track(
            ttk.Button(toolbar, text="Bỏ chọn", command=self.clear_sheet_selection)
        )
        self.btn_clear_sheets.pack(side="left", padx=(6, 0))

        canvas_row = ttk.Frame(self.sheet_frame)
        canvas_row.pack(fill="both", expand=True)

        self.sheet_canvas = tk.Canvas(canvas_row, height=100, highlightthickness=0, bg=self.COLOR_CARD)
        self.sheet_canvas.pack(side="left", fill="both", expand=True)

        sheet_sb = ttk.Scrollbar(canvas_row, orient="vertical", command=self.sheet_canvas.yview)
        sheet_sb.pack(side="right", fill="y")
        self.sheet_canvas.configure(yscrollcommand=sheet_sb.set)

        self.sheet_inner = ttk.Frame(self.sheet_canvas, style="Card.TFrame")
        self.sheet_window = self.sheet_canvas.create_window((0, 0), window=self.sheet_inner, anchor="nw")

        self.sheet_inner.bind(
            "<Configure>",
            lambda event: self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox("all")),
        )
        self.sheet_canvas.bind(
            "<Configure>",
            lambda event: self.sheet_canvas.itemconfigure(self.sheet_window, width=event.width),
        )

    def toggle_output_dir(self) -> None:
        if self.save_same_folder_var.get():
            self.btn_browse_dir.configure(state="disabled")
            self.lbl_dest_display.configure(text="Tự động lưu cùng thư mục với từng file Excel nguồn", foreground=self.COLOR_MUTED)
            self.output_dir_var.set("")
        else:
            self.btn_browse_dir.configure(state="normal")
            if self.output_dir_var.get():
                self.lbl_dest_display.configure(text=self.output_dir_var.get(), foreground=self.COLOR_TEXT)
            else:
                self.lbl_dest_display.configure(text="Vui lòng nhấn 'Chọn thư mục...'", foreground="#dc3545")

    def toggle_page_setup_ui(self) -> None:
        state = "normal" if self.ps_enabled_var.get() else "disabled"
        cbo_state = "readonly" if self.ps_enabled_var.get() else "disabled"

        self.chk_fit.configure(state=state)
        self.chk_center.configure(state=state)
        self.cbo_paper.configure(state=cbo_state)
        self.cbo_orient.configure(state=cbo_state)
        self.cbo_margin.configure(state=cbo_state)

    def select_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Chọn file Excel",
            filetypes=[
                ("File Excel (*.xlsx; *.xls; *.xlsm)", "*.xlsx *.xls *.xlsm"),
                ("Tất cả các file (*.*)", "*.*"),
            ],
        )
        if not selected:
            return

        for path in selected:
            norm = str(Path(path))
            if norm not in self.files:
                self.files.append(norm)

        self.refresh_file_list()
        self.clear_sheet_list()
        self.write_log(f"Đã thêm {len(selected)} file Excel vào danh sách.")

        if self.mode_var.get() == MODE_SELECTED_SHEETS:
            self.load_sheets_async()

    def clear_all_files(self) -> None:
        self.files.clear()
        self.file_listbox.delete(0, tk.END)
        self.lbl_file_count.configure(text="Chưa có file nào được chọn")
        self.clear_sheet_list()
        self.write_log("Đã làm sạch danh sách file.")

    def select_output_dir(self) -> None:
        directory = filedialog.askdirectory(title="Chọn thư mục lưu PDF")
        if directory:
            self.output_dir_var.set(str(Path(directory)))
            self.lbl_dest_display.configure(text=self.output_dir_var.get(), foreground=self.COLOR_TEXT)
            self.write_log(f"Thư mục lưu PDF được chọn: {directory}")

    def refresh_file_list(self) -> None:
        self.file_listbox.delete(0, tk.END)
        for path in self.files:
            self.file_listbox.insert(tk.END, Path(path).name)
        self.lbl_file_count.configure(text=f"Đã chọn {len(self.files)} file")

    def on_mode_changed(self) -> None:
        if self.mode_var.get() == MODE_SELECTED_SHEETS:
            if not self.sheet_frame.winfo_ismapped():
                self.sheet_frame.pack(fill="x", pady=(0, 8), before=self.chk_compress.master)
            if self.files and not self.sheet_vars and not self.sheet_loading:
                self.load_sheets_async()
        else:
            if self.sheet_frame.winfo_ismapped():
                self.sheet_frame.pack_forget()

    def clear_sheet_list(self) -> None:
        for checkbox in self.sheet_checkbuttons:
            self.tracked_widgets.pop(checkbox, None)

        for child in self.sheet_inner.winfo_children():
            child.destroy()

        self.sheet_vars.clear()
        self.sheet_checkbuttons.clear()

    def populate_sheet_list(self, sheet_infos: List[SheetInfo]) -> None:
        self.clear_sheet_list()

        if not sheet_infos:
            ttk.Label(self.sheet_inner, text="Không tìm thấy sheet nào.", style="Muted.TLabel").pack(anchor="w", padx=4, pady=4)
            return

        for info in sheet_infos:
            var = tk.BooleanVar(value=info.visible)
            text = info.name if info.visible else f"{info.name} (Sheet ẩn)"
            checkbox = ttk.Checkbutton(self.sheet_inner, text=text, variable=var)
            checkbox.pack(anchor="w", padx=4, pady=2)

            self.sheet_vars[info.name] = var
            self.sheet_checkbuttons.append(checkbox)

            if info.visible:
                self._track(checkbox)
            else:
                checkbox.configure(state="disabled")
                self._track(checkbox, normal_state="disabled")

    def select_all_sheets(self) -> None:
        for checkbox, (sheet_name, var) in zip(self.sheet_checkbuttons, self.sheet_vars.items()):
            if str(checkbox.cget("state")) != "disabled":
                var.set(True)

    def clear_sheet_selection(self) -> None:
        for var in self.sheet_vars.values():
            var.set(False)

    def load_sheets_async(self) -> None:
        if self.is_busy or self.sheet_loading:
            return

        if not self.files:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất một file Excel trước.")
            return

        self.sheet_loading = True
        self.status_var.set("Đang đọc danh sách sheet...")
        self.write_log("Đang kết nối Excel COM để nạp danh sách sheet...")

        worker = threading.Thread(target=self._sheet_loader_worker, daemon=True)
        worker.start()

    def _sheet_loader_worker(self) -> None:
        try:
            sheet_infos = read_sheet_catalog(self.files, self.queue_log)
            self.queue.put(("sheets", sheet_infos))
            self.queue.put(("status", f"Đã nạp {len(sheet_infos)} sheet."))
        except Exception as exc:
            self.queue_log(f"Lỗi đọc sheet: {format_error(exc)}")
            self.queue.put(("status", "Không đọc được sheet."))
        finally:
            self.queue.put(("sheet_loading", False))

    def get_page_setup_options(self) -> PageSetupOptions:
        enabled = self.ps_enabled_var.get()
        if not enabled:
            return PageSetupOptions(enabled=False)

        orient_raw = self.ps_orientation_var.get()
        orientation = "Portrait" if "Portrait" in orient_raw else ("Landscape" if "Landscape" in orient_raw else "Keep")

        paper_raw = self.ps_paper_size_var.get()
        paper_size = "A4" if "A4" in paper_raw else ("Letter" if "Letter" in paper_raw else "Keep")

        margin_raw = self.ps_margin_var.get()
        margin_mode = "Normal" if "Normal" in margin_raw else ("Narrow" if "Narrow" in margin_raw else "Keep")

        return PageSetupOptions(
            enabled=True,
            fit_to_width=self.ps_fit_width_var.get(),
            center_horizontally=self.ps_center_var.get(),
            paper_size=paper_size,
            orientation=orientation,
            margin_mode=margin_mode,
        )

    def start_conversion(self) -> None:
        if self.is_busy:
            return

        if not self.files:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất một file Excel cần chuyển đổi.")
            return

        mode = self.mode_var.get()
        selected_sheets: List[str] = []

        if mode == MODE_SELECTED_SHEETS:
            if self.sheet_loading:
                messagebox.showwarning("Đang tải", "Vui lòng chờ đọc xong danh sách sheet.")
                return

            selected_sheets = [name for name, var in self.sheet_vars.items() if var.get()]
            if not selected_sheets:
                messagebox.showwarning("Chưa chọn sheet", "Hãy tích chọn ít nhất 01 sheet cần xuất sang PDF.")
                return

        out_dir = None
        if not self.save_same_folder_var.get():
            if not self.output_dir_var.get():
                messagebox.showwarning("Chưa chọn nơi lưu", "Vui lòng chọn thư mục lưu file PDF.")
                return
            out_dir = Path(self.output_dir_var.get())

        page_opts = self.get_page_setup_options()

        self.progress_bar.configure(value=0, maximum=max(1, len(self.files)))
        self.set_busy(True)
        self.status_var.set("Đang chuyển đổi sang PDF...")
        self.write_log("=" * 60)
        self.write_log("BẮT ĐẦU QUÁ TRÌNH CHUYỂN ĐỔI SANG PDF")
        if page_opts.enabled:
            self.write_log(
                f"Tùy chọn in: FitWidth={page_opts.fit_to_width}, Center={page_opts.center_horizontally}, "
                f"Khổ={page_opts.paper_size}, Hướng={page_opts.orientation}"
            )
        else:
            self.write_log("Tùy chọn in: Giữ nguyên 100% bố cục trang in gốc của Excel.")

        worker = threading.Thread(
            target=self._conversion_worker,
            args=(
                list(self.files),
                out_dir,
                mode,
                selected_sheets,
                page_opts,
                bool(self.compress_var.get()),
            ),
            daemon=True,
        )
        worker.start()

    def _conversion_worker(
        self,
        files: List[str],
        output_dir: Optional[Path],
        mode: str,
        selected_sheets: List[str],
        page_options: PageSetupOptions,
        compress_after_export: bool,
    ) -> None:
        try:
            pdf_paths, zip_path = convert_files(
                file_paths=files,
                output_dir=output_dir,
                mode=mode,
                selected_sheet_names=selected_sheets,
                page_options=page_options,
                compress_after_export=compress_after_export,
                log=self.queue_log,
                progress=lambda done, total: self.queue.put(("progress", done, total)),
            )
            self.queue.put(("done", len(pdf_paths), str(zip_path) if zip_path else ""))
        except Exception as exc:
            self.queue_log(f"Lỗi chuyển đổi: {format_error(exc)}")
            self.queue.put(("failed", format_error(exc)))
        finally:
            self.queue.put(("busy", False))

    def set_busy(self, busy: bool) -> None:
        self.is_busy = busy
        for widget, normal_state in list(self.tracked_widgets.items()):
            try:
                widget.configure(state="disabled" if busy else normal_state)
            except tk.TclError:
                pass
        if not busy:
            self.toggle_output_dir()
            self.toggle_page_setup_ui()

    def queue_log(self, message: str) -> None:
        self.queue.put(("log", f"[{self.now()}] {message}"))

    def write_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def process_queue(self) -> None:
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]

                if kind == "log":
                    self.write_log(str(item[1]))
                elif kind == "status":
                    self.status_var.set(str(item[1]))
                elif kind == "progress":
                    done = int(item[1])
                    total = max(1, int(item[2]))
                    self.progress_bar.configure(maximum=total, value=done)
                elif kind == "sheets":
                    self.populate_sheet_list(item[1])
                elif kind == "sheet_loading":
                    self.sheet_loading = bool(item[1])
                elif kind == "busy":
                    self.set_busy(bool(item[1]))
                elif kind == "done":
                    pdf_count = int(item[1])
                    zip_path = str(item[2])
                    self.status_var.set("Hoàn tất thành công!")
                    msg = f"Đã chuyển đổi thành công {pdf_count} file PDF."
                    if zip_path:
                        msg += f"\nFile ZIP nén: {zip_path}"
                    self.write_log("=" * 60)
                    self.write_log(f"KẾT QUẢ: {msg}")
                    messagebox.showinfo("Thành công", msg)
                elif kind == "failed":
                    self.status_var.set("Có lỗi phát sinh.")
                    messagebox.showerror("Lỗi chuyển đổi", str(item[1]))
        except queue.Empty:
            pass

        self.root.after(100, self.process_queue)

    def on_close(self) -> None:
        if self.is_busy:
            messagebox.showwarning(
                "Đang xử lý",
                "Chương trình đang trong quá trình xuất PDF. Vui lòng đợi hoàn tất để đóng Excel sạch sẽ.",
            )
            return
        self.root.destroy()

    @staticmethod
    def now() -> str:
        return dt.datetime.now().strftime("%H:%M:%S")


def main() -> None:
    root = tk.Tk()
    ExcelToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
