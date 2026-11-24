# -*- coding: utf-8 -*-
"""
Excel to PDF Converter Pro - Ứng dụng Web Streamlit Chuyên Nghiệp.
Hỗ trợ:
- Chạy trực tiếp trên Web Browser với giao diện hiện đại, trực quan.
- Động cơ Microsoft Excel COM Automation (chất lượng in ấn văn phòng cao nhất).
- Hỗ trợ chế độ dự phòng LibreOffice / Cloud khi triển khai máy chủ.
- Xem trước trực tiếp file PDF (Interactive PDF Viewer & Thư viện ảnh từng trang) ngay trên trình duyệt mà không cần tải về trước.
- Đọc danh sách sheet động và cho phép chọn cụ thể sheet cần xuất.
- Bộ tối ưu trang in chuyên nghiệp (Fit 1 trang ngang, Căn giữa, Khổ giấy A4/Letter, Lề chuẩn/hẹp).
- Tải về từng file PDF lẻ hoặc đóng gói toàn bộ vào tệp ZIP một chạm.
"""

from __future__ import annotations

import base64
import gc
import io
import os
import re
import shutil
import subprocess
import tempfile
import time
import traceback
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import streamlit as st

# Kiểm tra pypdfium2 để kết xuất ảnh trang PDF xem trước
try:
    import pypdfium2
    HAS_PYPDFIUM2 = True
except ImportError:
    pypdfium2 = None
    HAS_PYPDFIUM2 = False

# Kiểm tra thư viện đọc sheet nhanh
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    openpyxl = None
    HAS_OPENPYXL = False

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    xlrd = None
    HAS_XLRD = False

# Kiểm tra hỗ trợ Excel COM trên Windows
try:
    import pythoncom
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    pythoncom = None
    win32com = None
    HAS_WIN32COM = False

# ==============================================================================
# HẰNG SỐ EXCEL VÀ HỆ THỐNG
# ==============================================================================
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_SHEET_VISIBLE = -1
XL_PORTRAIT = 1
XL_LANDSCAPE = 2
XL_PAPER_LETTER = 1
XL_PAPER_A4 = 9
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

INVALID_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}


def safe_filename(value: str, fallback: str = "output", max_length: int = 120) -> str:
    """Tạo tên file an toàn cho hệ thống nhưng vẫn giữ nguyên tiếng Việt có dấu."""
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
    """Đảm bảo đường dẫn file không bị trùng lặp bằng cách thêm hậu tố số."""
    path = Path(path)
    if not path.exists():
        return path
    for index in range(2, 10000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    return path


def clear_mark_of_the_web(path: Path) -> None:
    """Xóa nhãn bảo vệ Zone.Identifier khi file được tải về từ Web/Zalo."""
    if os.name != "nt":
        return
    try:
        os.remove(f"{path}:Zone.Identifier")
    except OSError:
        pass


def find_libreoffice_executable() -> Optional[str]:
    """Tìm đường dẫn thực thi của LibreOffice nếu có."""
    cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if cmd:
        return cmd
    common_win_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in common_win_paths:
        if os.path.exists(p):
            return p
    return None


def get_available_engine() -> Tuple[str, str]:
    """Xác định engine chuyển đổi khả dụng."""
    if HAS_WIN32COM and os.name == "nt":
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            version = str(getattr(excel, "Version", ""))
            excel.Quit()
            pythoncom.CoUninitialize()
            return "excel_com", f"Microsoft Excel COM (Bản quyền Office {version} - Chuẩn in ấn cao nhất)"
        except Exception:
            pass

    lo_path = find_libreoffice_executable()
    if lo_path:
        return "libreoffice", "LibreOffice Headless (Cross-platform)"

    return "none", "Không tìm thấy Microsoft Excel hoặc LibreOffice"


def extract_sheet_names(uploaded_file) -> List[str]:
    """Đọc danh sách sheet từ file Excel tải lên một cách nhanh chóng."""
    name_lower = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    # 1. Với file .xlsx, .xlsm dùng openpyxl
    if (name_lower.endswith(".xlsx") or name_lower.endswith(".xlsm")) and HAS_OPENPYXL:
        try:
            buf = io.BytesIO(file_bytes)
            wb = openpyxl.load_workbook(buf, read_only=True, keep_links=False)
            names = list(wb.sheetnames)
            wb.close()
            if names:
                return names
        except Exception:
            pass

    # 2. Với file .xls dùng xlrd
    if name_lower.endswith(".xls") and HAS_XLRD:
        try:
            book = xlrd.open_workbook(file_contents=file_bytes, on_demand=True)
            names = list(book.sheet_names())
            return names
        except Exception:
            pass

    return []


# ==============================================================================
# BỘ XỬ LÝ CHUYỂN ĐỔI BẰNG EXCEL COM
# ==============================================================================
def open_workbook_safely(excel, temp_path: Path):
    """Mở file Excel an toàn, tự động xử lý khi file bị kẹt Protected View."""
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
    except Exception:
        pass

    protected_view_window = None
    try:
        protected_view_window = excel.ProtectedViewWindows.Open(file_name)
        return protected_view_window.Edit()
    except Exception as exc:
        if protected_view_window is not None:
            try:
                protected_view_window.Close()
            except Exception:
                pass
        raise exc


def apply_professional_page_setup(sheet, excel_app, options: dict) -> None:
    """Áp dụng các tùy chọn trang in văn phòng chuyên nghiệp."""
    if not options.get("enabled", False):
        return
    try:
        ps = sheet.PageSetup

        # Hướng giấy
        if options.get("orientation") == "Portrait":
            try:
                ps.Orientation = XL_PORTRAIT
            except Exception:
                pass
        elif options.get("orientation") == "Landscape":
            try:
                ps.Orientation = XL_LANDSCAPE
            except Exception:
                pass

        # Khổ giấy
        if options.get("paper_size") == "A4":
            try:
                ps.PaperSize = XL_PAPER_A4
            except Exception:
                pass
        elif options.get("paper_size") == "Letter":
            try:
                ps.PaperSize = XL_PAPER_LETTER
            except Exception:
                pass

        # Vừa vặn 1 trang ngang (FitToPagesWide = 1)
        if options.get("fit_to_width", True):
            try:
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False  # Chiều dọc tự giãn theo nội dung dữ liệu
            except Exception:
                pass

        # Căn giữa trang theo chiều ngang
        if options.get("center_horizontally", True):
            try:
                ps.CenterHorizontally = True
            except Exception:
                pass

        # Căn lề
        margin_mode = options.get("margin_mode", "Keep")
        if margin_mode == "Normal":
            try:
                ps.LeftMargin = excel_app.CentimetersToPoints(1.8)
                ps.RightMargin = excel_app.CentimetersToPoints(1.8)
                ps.TopMargin = excel_app.CentimetersToPoints(2.0)
                ps.BottomMargin = excel_app.CentimetersToPoints(2.0)
            except Exception:
                pass
        elif margin_mode == "Narrow":
            try:
                ps.LeftMargin = excel_app.CentimetersToPoints(0.8)
                ps.RightMargin = excel_app.CentimetersToPoints(0.8)
                ps.TopMargin = excel_app.CentimetersToPoints(1.0)
                ps.BottomMargin = excel_app.CentimetersToPoints(1.0)
            except Exception:
                pass

    except Exception:
        pass


def convert_process_excel_com(
    saved_file_paths: List[Tuple[Path, str]],
    convert_mode: str,
    selected_sheets_map: Dict[str, List[str]],
    page_options: dict,
    progress_callback,
    status_callback,
) -> Tuple[List[Path], List[Tuple[str, str]], Path]:
    """Chuyển đổi file sử dụng Microsoft Excel COM trên Windows."""
    pythoncom.CoInitialize()

    excel = None
    success_files: List[Path] = []
    failed_files: List[Tuple[str, str]] = []
    temp_out_dir = Path(tempfile.mkdtemp(prefix="Streamlit_PDF_Out_"))

    try:
        status_callback("Đang khởi tạo phiên làm việc ngầm Microsoft Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")

        # Cấu hình Excel chạy ngầm an toàn tuyệt đối
        for attr, val in [
            ("Visible", False),
            ("DisplayAlerts", False),
            ("ScreenUpdating", False),
            ("Interactive", False),
            ("EnableEvents", False),
            ("AskToUpdateLinks", False),
            ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
        ]:
            try:
                setattr(excel, attr, val)
            except Exception:
                pass

        total_files = len(saved_file_paths)
        for idx, (file_path, original_name) in enumerate(saved_file_paths):
            status_callback(f"Đang xử lý file ({idx+1}/{total_files}): {original_name}...")
            progress_callback(idx / total_files)

            wb = None
            try:
                wb = open_workbook_safely(excel, file_path)

                if convert_mode == "all_sheets":
                    # Xuất toàn bộ workbook thành 1 file PDF
                    if page_options.get("enabled", False):
                        for ws in wb.Worksheets:
                            if int(ws.Visible) == XL_SHEET_VISIBLE:
                                apply_professional_page_setup(ws, excel, page_options)

                    pdf_name = f"{Path(original_name).stem}.pdf"
                    target_pdf = unique_path(temp_out_dir / pdf_name)

                    wb.ExportAsFixedFormat(
                        Type=XL_TYPE_PDF,
                        Filename=str(target_pdf),
                        Quality=XL_QUALITY_STANDARD,
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False,
                    )
                    success_files.append(target_pdf)

                elif convert_mode == "each_sheet":
                    # Xuất từng sheet hiển thị thành 1 file PDF riêng
                    for ws in wb.Worksheets:
                        if int(ws.Visible) == XL_SHEET_VISIBLE:
                            if page_options.get("enabled", False):
                                apply_professional_page_setup(ws, excel, page_options)

                            safe_sheet = safe_filename(str(ws.Name), fallback="Sheet")
                            pdf_name = f"{Path(original_name).stem} - {safe_sheet}.pdf"
                            target_pdf = unique_path(temp_out_dir / pdf_name)

                            ws.ExportAsFixedFormat(
                                Type=XL_TYPE_PDF,
                                Filename=str(target_pdf),
                                Quality=XL_QUALITY_STANDARD,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )
                            success_files.append(target_pdf)

                elif convert_mode == "selected_sheets":
                    # Xuất các sheet được người dùng tích chọn
                    target_sheets = selected_sheets_map.get(original_name, [])
                    target_sheets_lower = {s.lower() for s in target_sheets}

                    exported_count = 0
                    for ws in wb.Worksheets:
                        if int(ws.Visible) == XL_SHEET_VISIBLE:
                            if str(ws.Name).lower() in target_sheets_lower:
                                if page_options.get("enabled", False):
                                    apply_professional_page_setup(ws, excel, page_options)

                                safe_sheet = safe_filename(str(ws.Name), fallback="Sheet")
                                pdf_name = f"{Path(original_name).stem} - {safe_sheet}.pdf"
                                target_pdf = unique_path(temp_out_dir / pdf_name)

                                ws.ExportAsFixedFormat(
                                    Type=XL_TYPE_PDF,
                                    Filename=str(target_pdf),
                                    Quality=XL_QUALITY_STANDARD,
                                    IncludeDocProperties=True,
                                    IgnorePrintAreas=False,
                                    OpenAfterPublish=False,
                                )
                                success_files.append(target_pdf)
                                exported_count += 1

                    if exported_count == 0:
                        failed_files.append((original_name, "Không có sheet nào được chọn hoặc sheet bị ẩn."))

            except Exception as exc:
                failed_files.append((original_name, str(exc)))
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                    del wb
                    gc.collect()

        progress_callback(1.0)
        status_callback("Quá trình xuất PDF hoàn tất!")

    finally:
        if excel is not None:
            try:
                excel.ScreenUpdating = True
            except Exception:
                pass
            try:
                excel.Interactive = True
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass
            del excel

        pythoncom.CoUninitialize()

    return success_files, failed_files, temp_out_dir


def convert_process_libreoffice(
    saved_file_paths: List[Tuple[Path, str]],
    soffice_path: str,
    convert_mode: str,
    selected_sheets_map: Dict[str, List[str]],
    progress_callback,
    status_callback,
) -> Tuple[List[Path], List[Tuple[str, str]], Path]:
    """Chuyển đổi file sử dụng LibreOffice Headless (chế độ tối ưu cho Linux / Streamlit Community Cloud)."""
    success_files: List[Path] = []
    failed_files: List[Tuple[str, str]] = []
    temp_out_dir = Path(tempfile.mkdtemp(prefix="Streamlit_PDF_LO_Out_"))

    total_files = len(saved_file_paths)
    for idx, (file_path, original_name) in enumerate(saved_file_paths):
        status_callback(f"Đang xử lý qua LibreOffice ({idx+1}/{total_files}): {original_name}...")
        progress_callback(idx / total_files)

        try:
            if convert_mode == "all_sheets":
                cmd = [
                    soffice_path,
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(temp_out_dir),
                    str(file_path),
                ]
                res = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
                if res.returncode != 0:
                    failed_files.append((original_name, res.stderr or "Lỗi LibreOffice khi chuyển đổi"))
                    continue

                expected_pdf = temp_out_dir / f"{file_path.stem}.pdf"
                if expected_pdf.exists():
                    final_pdf = unique_path(temp_out_dir / f"{Path(original_name).stem}.pdf")
                    if expected_pdf != final_pdf:
                        shutil.move(str(expected_pdf), str(final_pdf))
                    success_files.append(final_pdf)
                else:
                    failed_files.append((original_name, "Không tìm thấy file PDF được tạo ra."))
            else:
                # Tách từng sheet hoặc xuất các sheet được chọn
                ext = file_path.suffix.lower()
                if ext in [".xlsx", ".xlsm"] and HAS_OPENPYXL:
                    wb = openpyxl.load_workbook(file_path, data_only=False)
                    all_sheets = wb.sheetnames
                    wb.close()

                    if convert_mode == "selected_sheets":
                        selected = selected_sheets_map.get(original_name, [])
                        target_sheets = [s for s in all_sheets if s in selected]
                    else:
                        target_sheets = all_sheets

                    if not target_sheets:
                        failed_files.append((original_name, "Không có sheet nào được chọn."))
                        continue

                    for sname in target_sheets:
                        single_wb = openpyxl.load_workbook(file_path, data_only=False)
                        for other in list(single_wb.sheetnames):
                            if other != sname:
                                del single_wb[other]
                        temp_sheet_file = file_path.parent / f"{file_path.stem}_{safe_filename(sname)}{ext}"
                        single_wb.save(temp_sheet_file)
                        single_wb.close()

                        cmd = [
                            soffice_path,
                            "--headless",
                            "--convert-to", "pdf",
                            "--outdir", str(temp_out_dir),
                            str(temp_sheet_file),
                        ]
                        subprocess.run(cmd, capture_output=True, text=True, timeout=180)
                        try:
                            os.remove(temp_sheet_file)
                        except Exception:
                            pass

                        expected_pdf = temp_out_dir / f"{temp_sheet_file.stem}.pdf"
                        if expected_pdf.exists():
                            final_name = f"{Path(original_name).stem} - {safe_filename(sname)}.pdf"
                            final_pdf = unique_path(temp_out_dir / final_name)
                            if expected_pdf != final_pdf:
                                shutil.move(str(expected_pdf), str(final_pdf))
                            success_files.append(final_pdf)
                else:
                    cmd = [
                        soffice_path,
                        "--headless",
                        "--convert-to", "pdf",
                        "--outdir", str(temp_out_dir),
                        str(file_path),
                    ]
                    res = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
                    expected_pdf = temp_out_dir / f"{file_path.stem}.pdf"
                    if expected_pdf.exists():
                        final_pdf = unique_path(temp_out_dir / f"{Path(original_name).stem}.pdf")
                        if expected_pdf != final_pdf:
                            shutil.move(str(expected_pdf), str(final_pdf))
                        success_files.append(final_pdf)
        except Exception as exc:
            failed_files.append((original_name, str(exc)))

    progress_callback(1.0)
    status_callback("Quá trình chuyển đổi bằng LibreOffice hoàn tất!")
    return success_files, failed_files, temp_out_dir


# ==============================================================================
# HÀM HIỂN THỊ FILE PDF TRỰC TIẾP TRÊN WEB
# ==============================================================================
def render_pdf_to_images(pdf_bytes: bytes, max_pages: int = 20, scale: float = 1.8):
    """Kết xuất các trang PDF thành ảnh PIL để người dùng xem trực tiếp trên trình duyệt."""
    if not HAS_PYPDFIUM2:
        return [], 0
    try:
        doc = pypdfium2.PdfDocument(pdf_bytes)
        total_pages = len(doc)
        images = []
        for i in range(min(total_pages, max_pages)):
            page = doc[i]
            pil_img = page.render(scale=scale).to_pil()
            images.append((i + 1, pil_img))
        return images, total_pages
    except Exception:
        return [], 0


def get_pdf_iframe_html(pdf_bytes: bytes, height: int = 700) -> str:
    """Tạo mã nhúng iframe PDF với base64 để người dùng tương tác trực tiếp."""
    b64_data = base64.b64encode(pdf_bytes).decode("utf-8")
    return (
        f'<iframe src="data:application/pdf;base64,{b64_data}#toolbar=1&navpanes=1&scrollbar=1" '
        f'width="100%" height="{height}px" type="application/pdf" '
        f'style="border: 1px solid #c7d2fe; border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">'
        f'</iframe>'
    )


# ==============================================================================
# GIAO DIỆN STREAMLIT CHUYÊN NGHIỆP
# ==============================================================================
st.set_page_config(
    page_title="Excel to PDF Converter Pro - Trình Chuyển Đổi Trên Web",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling (Emerald Office theme)
st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #065f46;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #4b5563;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
    }
    .badge-card {
        padding: 0.75rem 1rem;
        border-radius: 8px;
        background-color: #ecfdf5;
        border-left: 4px solid #059669;
        margin-bottom: 1rem;
    }
    .stat-box {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        text-align: center;
    }
    .stDownloadButton button {
        background-color: #059669 !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
    }
    .stDownloadButton button:hover {
        background-color: #047857 !important;
        color: white !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

engine_type, engine_desc = get_available_engine()

# --- SIDEBAR: CẤU HÌNH & THÔNG TIN ---
with st.sidebar:
    st.markdown("### ⚙️ Trạng Thái Hệ Thống")
    if engine_type == "excel_com":
        st.success("🟢 **Động cơ:** Microsoft Excel COM\n\n*Chuẩn in ấn văn phòng gốc, giữ nguyên công thức và đồ họa.*")
    elif engine_type == "libreoffice":
        st.info(f"🟡 **Động cơ:** {engine_desc}")
    else:
        st.error("🔴 **Cảnh báo:** Chưa tìm thấy Excel COM hoặc LibreOffice để chuyển đổi.")

    st.markdown("---")
    st.markdown("### 📐 Tối Ưu Trang In Chuyên Nghiệp")
    enable_ps = st.checkbox(
        "Bật bộ tối ưu trang in (Khuyên dùng)",
        value=True,
        help="Tự động căn chỉnh bảng tính vừa vặn 1 trang ngang, căn giữa và chuẩn hóa khổ giấy A4.",
    )

    page_options = {"enabled": enable_ps}

    if enable_ps:
        fit_width = st.checkbox(
            "Khớp 1 trang chiều ngang (Fit 1 page wide)",
            value=True,
            help="Tự động co giãn theo chiều rộng để không bị đứt bảng, tràn cột sang trang sau.",
        )
        center_h = st.checkbox("Căn giữa trang theo chiều ngang", value=True)
        page_options["fit_to_width"] = fit_width
        page_options["center_horizontally"] = center_h

        paper = st.selectbox(
            "Khổ giấy xuất:",
            options=["A4 (Tiêu chuẩn Việt Nam)", "Letter (Tiêu chuẩn Mỹ)", "Giữ nguyên thiết lập gốc"],
            index=0,
        )
        page_options["paper_size"] = "A4" if "A4" in paper else ("Letter" if "Letter" in paper else "Keep")

        orient = st.selectbox(
            "Hướng giấy:",
            options=["Giữ nguyên thiết lập gốc", "Dọc (Portrait)", "Ngang (Landscape)"],
            index=0,
        )
        page_options["orientation"] = "Portrait" if "Portrait" in orient else ("Landscape" if "Landscape" in orient else "Keep")

        margin = st.selectbox(
            "Căn lề trang:",
            options=["Giữ nguyên thiết lập gốc", "Tiêu chuẩn (Normal - Lề 1.8cm)", "Hẹp (Narrow - Lề 0.8cm)"],
            index=0,
        )
        page_options["margin_mode"] = "Normal" if "Normal" in margin else ("Narrow" if "Narrow" in margin else "Keep")

    st.markdown("---")
    st.markdown("### 💡 Mẹo Sử Dụng Trực Tiếp Trên Web")
    st.info(
        """
        - Bạn có thể **kéo thả nhiều file Excel cùng lúc**.
        - Sau khi chuyển đổi, bạn có thể **xem trực tiếp bản PDF ngay trên web** trước khi tải về.
        - Chế độ **Nén ZIP** sẽ tự động kích hoạt khi có từ 2 file PDF trở lên.
        """
    )


# --- NỘI DUNG CHÍNH ---
st.markdown('<div class="main-title">Excel to PDF Converter Pro 📑</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Chuyển đổi file Excel sang PDF trực tiếp trên trình duyệt web, đảm bảo bố cục chuẩn đẹp và xem trước trực quan.</div>',
    unsafe_allow_html=True,
)

# Khởi tạo session state lưu trữ kết quả chuyển đổi
if "converted_results" not in st.session_state:
    st.session_state.converted_results = None
if "active_preview_index" not in st.session_state:
    st.session_state.active_preview_index = 0

# BƯỚC 1: TẢI FILE EXCEL
col_upload, col_mode = st.columns([1.4, 1.0])

with col_upload:
    st.subheader("1. Tải lên file Excel cần chuyển")
    uploaded_files = st.file_uploader(
        "Kéo thả hoặc chọn các file .xlsx, .xls, .xlsm:",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        help="Hỗ trợ cả file Excel thông thường, file macro (.xlsm) và file cũ (.xls)",
    )

with col_mode:
    st.subheader("2. Chế độ xuất PDF")
    convert_mode_option = st.radio(
        "Chọn quy cách xuất file:",
        options=[
            ("📁 Xuất toàn bộ Workbook (Mỗi file Excel thành 1 PDF)", "all_sheets"),
            ("📑 Xuất từng Sheet riêng lẻ (Mỗi sheet thành 1 PDF)", "each_sheet"),
            ("🎯 Tùy chọn chỉ xuất các Sheet cụ thể", "selected_sheets"),
        ],
        format_func=lambda x: x[0],
    )
    convert_mode = convert_mode_option[1]

# BƯỚC 2: QUẢN LÝ VÀ CHỌN SHEET (NẾU CÓ)
selected_sheets_map: Dict[str, List[str]] = {}

if uploaded_files:
    total_size_mb = sum(f.size for f in uploaded_files) / (1024 * 1024)
    st.markdown(
        f'<div class="badge-card">✓ Đã tải lên <b>{len(uploaded_files)} file</b> | Tổng dung lượng: <b>{total_size_mb:.2f} MB</b></div>',
        unsafe_allow_html=True,
    )

    if convert_mode == "selected_sheets":
        st.subheader("2.1. Chọn các Sheet bạn muốn xuất PDF:")
        with st.expander("🔍 Danh sách các Sheet trong từng file Excel", expanded=True):
            for uf in uploaded_files:
                detected_sheets = extract_sheet_names(uf)
                if detected_sheets:
                    selected = st.multiselect(
                        f"File: **{uf.name}**",
                        options=detected_sheets,
                        default=detected_sheets,
                        key=f"sheets_{uf.name}",
                    )
                    selected_sheets_map[uf.name] = selected
                else:
                    st.warning(f"File **{uf.name}**: Không phát hiện danh sách sheet trước. Sẽ xuất toàn bộ sheet hiển thị.")
                    selected_sheets_map[uf.name] = []

    # BƯỚC 3: NÚT THỰC HIỆN
    st.subheader("3. Bắt đầu chuyển đổi")
    col_btn_start, col_btn_clear = st.columns([3, 1])

    with col_btn_start:
        start_btn = st.button("🚀 BẮT ĐẦU CHUYỂN ĐỔI SANG PDF", use_container_width=True, type="primary")

    with col_btn_clear:
        if st.button("🔄 Đặt lại / Làm mới", use_container_width=True):
            st.session_state.converted_results = None
            st.rerun()

    if start_btn:
        temp_in_dir = Path(tempfile.mkdtemp(prefix="Streamlit_In_"))
        saved_file_paths: List[Tuple[Path, str]] = []

        try:
            for uf in uploaded_files:
                dest = temp_in_dir / safe_filename(uf.name, fallback="upload.xlsx")
                with open(dest, "wb") as f:
                    f.write(uf.getvalue())
                clear_mark_of_the_web(dest)
                saved_file_paths.append((dest, uf.name))

            progress_bar = st.progress(0)
            status_text = st.empty()

            start_time = time.time()

            if engine_type == "excel_com":
                success_paths, failed_files, temp_out_dir = convert_process_excel_com(
                    saved_file_paths=saved_file_paths,
                    convert_mode=convert_mode,
                    selected_sheets_map=selected_sheets_map,
                    page_options=page_options,
                    progress_callback=progress_bar.progress,
                    status_callback=status_text.text,
                )
            elif engine_type == "libreoffice":
                lo_exe = find_libreoffice_executable() or "soffice"
                success_paths, failed_files, temp_out_dir = convert_process_libreoffice(
                    saved_file_paths=saved_file_paths,
                    soffice_path=lo_exe,
                    convert_mode=convert_mode,
                    selected_sheets_map=selected_sheets_map,
                    progress_callback=progress_bar.progress,
                    status_callback=status_text.text,
                )
            else:
                st.error("❌ Không tìm thấy công cụ chuyển đổi tương thích trên máy chủ.")
                success_paths, failed_files, temp_out_dir = [], [(f.name, "Thiếu engine chuyển đổi") for f in uploaded_files], None

            elapsed = time.time() - start_time

            # Đọc sẵn dữ liệu bytes vào bộ nhớ để session state lưu trữ và xóa temp_out_dir an toàn
            converted_data = []
            if success_paths and temp_out_dir:
                for p in success_paths:
                    with open(p, "rb") as f:
                        data = f.read()
                    converted_data.append({"name": p.name, "bytes": data, "size": len(data)})
                shutil.rmtree(temp_out_dir, ignore_errors=True)

            st.session_state.converted_results = {
                "files": converted_data,
                "failed": failed_files,
                "elapsed": elapsed,
            }

        except Exception as main_exc:
            st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý: {str(main_exc)}")
            st.code(traceback.format_exc())
        finally:
            shutil.rmtree(temp_in_dir, ignore_errors=True)

# BƯỚC 4: KẾT QUẢ VÀ TRÌNH XEM TRỰC TIẾP TRÊN WEB
if st.session_state.converted_results is not None:
    results = st.session_state.converted_results
    pdf_list = results.get("files", [])
    failed_list = results.get("failed", [])
    elapsed_sec = results.get("elapsed", 0)

    st.markdown("---")
    st.subheader("4. Kết Quả Chuyển Đổi & Xem Trực Tiếp Trên Web")

    # Bảng số liệu tóm tắt
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    with col_stat1:
        st.metric("Số file PDF tạo thành công", f"{len(pdf_list)} file")
    with col_stat2:
        total_pdf_size = sum(item["size"] for item in pdf_list) / (1024 * 1024)
        st.metric("Tổng dung lượng PDF", f"{total_pdf_size:.2f} MB")
    with col_stat3:
        st.metric("Thời gian xử lý", f"{elapsed_sec:.1f} giây")

    if failed_list:
        with st.expander("⚠️ Một số file gặp sự cố trong quá trình xuất:", expanded=True):
            for fname, err in failed_list:
                st.write(f"- **{fname}**: `{err}`")

    if pdf_list:
        st.success("🎉 Tất cả các file PDF đã được tạo hoàn tất! Bạn có thể tải về hoặc xem trực tiếp bên dưới:")

        # Khu vực Tải Về
        col_down_all, col_down_select = st.columns([1.5, 2.5])

        with col_down_all:
            if len(pdf_list) == 1:
                item = pdf_list[0]
                st.download_button(
                    label=f"📥 Tải Về File PDF: {item['name']}",
                    data=item["bytes"],
                    file_name=item["name"],
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                # Tạo file zip trong bộ nhớ
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
                    for item in pdf_list:
                        zipf.writestr(item["name"], item["bytes"])
                zip_bytes = zip_buffer.getvalue()

                st.download_button(
                    label=f"📦 Tải Về Toàn Bộ {len(pdf_list)} File (.ZIP)",
                    data=zip_bytes,
                    file_name="Excel_To_PDF_Results.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

        with col_down_select:
            selected_pdf_idx = st.selectbox(
                "Chọn file PDF cần xem trước hoặc tải lẻ:",
                options=list(range(len(pdf_list))),
                format_func=lambda idx: f"{pdf_list[idx]['name']} ({(pdf_list[idx]['size']/1024):.1f} KB)",
            )

        current_pdf = pdf_list[selected_pdf_idx]

        # Tùy chọn tải riêng file đang chọn
        st.download_button(
            label=f"📥 Tải xuống riêng file '{current_pdf['name']}'",
            data=current_pdf["bytes"],
            file_name=current_pdf["name"],
            mime="application/pdf",
        )

        # Tabs hiển thị: Trình xem PDF Tương tác vs Thư viện ảnh trang
        tab_interactive, tab_pages = st.tabs(["👁️ Trình Xem PDF Tương Tác", "🖼️ Xem Toàn Bộ Trang (Gallery)"])

        with tab_interactive:
            st.markdown("##### Bản xem trước PDF trực tiếp (Có hỗ trợ cuộn, phóng to, in):")
            iframe_html = get_pdf_iframe_html(current_pdf["bytes"], height=750)
            st.components.v1.html(iframe_html, height=760, scrolling=True)

        with tab_pages:
            st.markdown("##### Xem từng trang đã được render với độ phân giải cao:")
            page_images, total_p = render_pdf_to_images(current_pdf["bytes"], max_pages=20)
            if page_images:
                st.caption(f"Đang hiển thị {len(page_images)} / {total_p} trang của file PDF.")
                for page_num, img in page_images:
                    st.markdown(f"**Trang {page_num} / {total_p}:**")
                    st.image(img, use_container_width=True)
            else:
                st.info("Trình render ảnh trang phụ thuộc vào thư viện `pypdfium2`. Bạn có thể sử dụng tab 'Trình Xem PDF Tương Tác' ở bên cạnh.")

else:
    if not uploaded_files:
        st.info("👆 Vui lòng kéo thả hoặc chọn file Excel ở khung trên để bắt đầu chuyển đổi trực tiếp trên web.")
