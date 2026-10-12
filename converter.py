import os
import sys
import time
import uuid
try:
    import pymupdf
except ImportError:
    pymupdf = None
import openpyxl
from typing import Dict, Any, List, Optional

# Optional Windows COM import (graceful fallback on Vercel/Linux)
try:
    import pythoncom
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    pythoncom = None
    win32com = None
    COM_AVAILABLE = False

# Excel Page Setup Constants
XL_TYPE_PDF = 0
XL_PORTRAIT = 1
XL_LANDSCAPE = 2

XL_PAPER_LETTER = 1
XL_PAPER_LEGAL = 5
XL_PAPER_A4 = 9
XL_PAPER_A3 = 8
XL_PAPER_EXECUTIVE = 7

PAPER_SIZES = {
    "A4": XL_PAPER_A4,
    "Letter": XL_PAPER_LETTER,
    "Legal": XL_PAPER_LEGAL,
    "A3": XL_PAPER_A3,
    "Executive": XL_PAPER_EXECUTIVE,
}

# Unit conversion helpers
# 1 inch = 72 points
# 1 cm = 72 / 2.54 points ≈ 28.3464567 points
def cm_to_points(cm: float) -> float:
    try:
        return float(cm) * 72.0 / 2.54
    except (ValueError, TypeError):
        return 1.91 * 72.0 / 2.54

def inch_to_points(inch: float) -> float:
    try:
        return float(inch) * 72.0
    except (ValueError, TypeError):
        return 0.75 * 72.0

def points_to_cm(pt: float) -> float:
    return float(pt) * 2.54 / 72.0

def points_to_inch(pt: float) -> float:
    return float(pt) / 72.0


def check_excel_availability() -> bool:
    """Checks if Microsoft Excel COM is available on the machine."""
    if not COM_AVAILABLE or pythoncom is None or win32com is None:
        return False
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Quit()
        return True
    except Exception:
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def inspect_excel_file(file_path: str) -> Dict[str, Any]:
    """
    Reads workbook structure, sheets, row/col counts using openpyxl.
    """
    file_path = os.path.abspath(file_path)
    filename = os.path.basename(file_path)
    file_size = os.path.getsize(file_path)

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        details = []
        for name in sheets:
            try:
                ws = wb[name]
                details.append({
                    "name": name,
                    "max_row": ws.max_row or 0,
                    "max_column": ws.max_column or 0
                })
            except Exception:
                details.append({"name": name, "max_row": 0, "max_column": 0})
        wb.close()
        return {
            "success": True,
            "sheets": sheets,
            "details": details,
            "filename": filename,
            "filesize": file_size,
            "filesize_formatted": format_filesize(file_size)
        }
    except Exception as e:
        # Fallback for non-standard formats (e.g. .xls)
        return {
            "success": True,
            "sheets": ["Semua Lembar Kerja (Workbook)"],
            "details": [{"name": "Default", "max_row": 0, "max_column": 0}],
            "filename": filename,
            "filesize": file_size,
            "filesize_formatted": format_filesize(file_size),
            "note": str(e)
        }


def format_filesize(num_bytes: int) -> str:
    for unit in ['B', 'KB', 'MB', 'GB']:
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} TB"


def convert_excel_to_pdf(
    excel_path: str,
    output_pdf_path: str,
    top_margin_pt: float,
    bottom_margin_pt: float,
    left_margin_pt: float,
    right_margin_pt: float,
    header_margin_pt: Optional[float] = None,
    footer_margin_pt: Optional[float] = None,
    orientation: str = "portrait",
    paper_size: str = "A4",
    fit_mode: str = "fit_width",
    center_horizontally: bool = True,
    center_vertically: bool = False,
    sheet_name: str = "all",
    show_gridlines: bool = False,
    show_headings: bool = False
) -> Dict[str, Any]:
    """
    Main conversion entry point. Uses Native Microsoft Excel COM for 100% fidelity,
    and falls back to ReportLab table generator if COM encounters issues.
    """
    excel_path = os.path.abspath(excel_path)
    output_pdf_path = os.path.abspath(output_pdf_path)
    os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)

    start_time = time.time()
    engine_used = "Microsoft Excel COM (Native)"
    error_msg = None

    if COM_AVAILABLE:
        try:
            _convert_via_com(
                excel_path=excel_path,
                output_pdf_path=output_pdf_path,
                top_margin_pt=top_margin_pt,
                bottom_margin_pt=bottom_margin_pt,
                left_margin_pt=left_margin_pt,
                right_margin_pt=right_margin_pt,
                header_margin_pt=header_margin_pt,
                footer_margin_pt=footer_margin_pt,
                orientation=orientation,
                paper_size=paper_size,
                fit_mode=fit_mode,
                center_horizontally=center_horizontally,
                center_vertically=center_vertically,
                sheet_name=sheet_name,
                show_gridlines=show_gridlines,
                show_headings=show_headings
            )
        except Exception as exc:
            error_msg = str(exc)
            print(f"[Converter] Excel COM error: {error_msg}. Attempting fallback...", file=sys.stderr)
            engine_used = "ReportLab Engine"
            _convert_via_reportlab_fallback(
                excel_path=excel_path,
                output_pdf_path=output_pdf_path,
                top_margin_pt=top_margin_pt,
                bottom_margin_pt=bottom_margin_pt,
                left_margin_pt=left_margin_pt,
                right_margin_pt=right_margin_pt,
                orientation=orientation,
                paper_size=paper_size,
                sheet_name=sheet_name
            )
    else:
        engine_used = "ReportLab Engine (Cloud/Vercel)"
        _convert_via_reportlab_fallback(
            excel_path=excel_path,
            output_pdf_path=output_pdf_path,
            top_margin_pt=top_margin_pt,
            bottom_margin_pt=bottom_margin_pt,
            left_margin_pt=left_margin_pt,
            right_margin_pt=right_margin_pt,
            orientation=orientation,
            paper_size=paper_size,
            sheet_name=sheet_name
        )

    duration = round(time.time() - start_time, 2)
    pdf_size = os.path.getsize(output_pdf_path)

    return {
        "success": True,
        "engine": engine_used,
        "duration_seconds": duration,
        "pdf_size": pdf_size,
        "pdf_size_formatted": format_filesize(pdf_size),
        "output_path": output_pdf_path
    }


def _convert_via_com(
    excel_path: str,
    output_pdf_path: str,
    top_margin_pt: float,
    bottom_margin_pt: float,
    left_margin_pt: float,
    right_margin_pt: float,
    header_margin_pt: Optional[float],
    footer_margin_pt: Optional[float],
    orientation: str,
    paper_size: str,
    fit_mode: str,
    center_horizontally: bool,
    center_vertically: bool,
    sheet_name: str,
    show_gridlines: bool,
    show_headings: bool
):
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        workbook = excel.Workbooks.Open(
            excel_path,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True
        )

        sheets_to_process = []
        if sheet_name and sheet_name.lower() != "all":
            try:
                found_sheet = workbook.Worksheets(sheet_name)
                sheets_to_process.append(found_sheet)
            except Exception:
                sheets_to_process.append(workbook.ActiveSheet)
        else:
            for i in range(1, workbook.Worksheets.Count + 1):
                ws = workbook.Worksheets(i)
                if ws.Visible == -1:  # xlSheetVisible
                    sheets_to_process.append(ws)

        if not sheets_to_process:
            sheets_to_process.append(workbook.ActiveSheet)

        for ws in sheets_to_process:
            ps = ws.PageSetup

            # 1. Orientation must be applied first before margins & scaling
            try:
                if str(orientation).strip().lower() == "landscape":
                    ps.Orientation = XL_LANDSCAPE
                else:
                    ps.Orientation = XL_PORTRAIT
            except Exception as e:
                print(f"[Converter] Orientation error: {e}", file=sys.stderr)

            # 2. Paper Size
            ps_code = PAPER_SIZES.get(str(paper_size).strip(), XL_PAPER_A4)
            try:
                ps.PaperSize = ps_code
            except Exception:
                pass

            # 3. Custom Margins
            try:
                ps.TopMargin = top_margin_pt
                ps.BottomMargin = bottom_margin_pt
                ps.LeftMargin = left_margin_pt
                ps.RightMargin = right_margin_pt

                if header_margin_pt is not None:
                    ps.HeaderMargin = header_margin_pt
                if footer_margin_pt is not None:
                    ps.FooterMargin = footer_margin_pt
            except Exception as e:
                print(f"[Converter] Margin error: {e}", file=sys.stderr)

            try:
                ps.CenterHorizontally = bool(center_horizontally)
                ps.CenterVertically = bool(center_vertically)
            except Exception:
                pass

            try:
                ps.PrintGridlines = bool(show_gridlines)
                ps.PrintHeadings = bool(show_headings)
            except Exception:
                pass

            # Scaling / Fit to page setup
            try:
                if fit_mode == "fit_width":
                    ps.Zoom = False
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False
                elif fit_mode == "fit_page":
                    ps.Zoom = False
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = 1
                elif fit_mode == "fit_height":
                    ps.Zoom = False
                    ps.FitToPagesWide = False
                    ps.FitToPagesTall = 1
                else:
                    ps.Zoom = 100
            except Exception:
                pass

        if os.path.exists(output_pdf_path):
            os.remove(output_pdf_path)

        if sheet_name and sheet_name.lower() != "all" and len(sheets_to_process) == 1:
            sheets_to_process[0].ExportAsFixedFormat(XL_TYPE_PDF, output_pdf_path)
        else:
            workbook.ExportAsFixedFormat(XL_TYPE_PDF, output_pdf_path)

        if not os.path.exists(output_pdf_path) or os.path.getsize(output_pdf_path) == 0:
            raise RuntimeError("PDF file was not created or empty.")

    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _convert_via_reportlab_fallback(
    excel_path: str,
    output_pdf_path: str,
    top_margin_pt: float,
    bottom_margin_pt: float,
    left_margin_pt: float,
    right_margin_pt: float,
    orientation: str,
    paper_size: str,
    sheet_name: str
):
    from reportlab.lib.pagesizes import A4, LETTER, LEGAL, A3, landscape, portrait
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    size_map = {
        "A4": A4,
        "Letter": LETTER,
        "Legal": LEGAL,
        "A3": A3
    }
    base_size = size_map.get(paper_size, A4)
    pagesize = landscape(base_size) if orientation.lower() == "landscape" else portrait(base_size)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if sheet_name and sheet_name.lower() != "all" and sheet_name in wb.sheetnames:
        sheets = [wb[sheet_name]]
    else:
        sheets = [wb[s] for s in wb.sheetnames]

    doc = SimpleDocTemplate(
        output_pdf_path,
        pagesize=pagesize,
        leftMargin=left_margin_pt,
        rightMargin=right_margin_pt,
        topMargin=top_margin_pt,
        bottomMargin=bottom_margin_pt
    )

    story = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'SheetTitle',
        parent=styles['Heading2'],
        textColor=colors.HexColor('#1E293B'),
        spaceAfter=10
    )

    for ws in sheets:
        if len(sheets) > 1:
            story.append(Paragraph(f"<b>Sheet: {ws.title}</b>", title_style))
            story.append(Spacer(1, 8))

        data = []
        for row in ws.iter_rows(values_only=True):
            cleaned_row = [str(cell) if cell is not None else "" for cell in row]
            if any(cleaned_row):
                data.append(cleaned_row)

        if not data:
            data = [["(Lembar kerja kosong)"]]

        # Calculate max columns
        max_cols = max(len(r) for r in data)
        padded_data = [r + [""] * (max_cols - len(r)) for r in data]

        t = Table(padded_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0F172A')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ]))
        story.append(t)
        story.append(Spacer(1, 15))

    doc.build(story)
    wb.close()


def generate_pdf_previews(pdf_path: str, output_dir: str, max_pages: int = 10) -> Dict[str, Any]:
    """
    Renders preview images of the PDF using PyMuPDF.
    """
    os.makedirs(output_dir, exist_ok=True)
    if not pymupdf:
        return {"total_pages": 1, "previews": []}
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)
    preview_files = []

    pages_to_render = min(total_pages, max_pages)
    for page_idx in range(pages_to_render):
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=140)
        img_name = f"page_{page_idx + 1}.png"
        img_path = os.path.join(output_dir, img_name)
        pix.save(img_path)
        preview_files.append({
            "page": page_idx + 1,
            "filename": img_name,
            "width": pix.width,
            "height": pix.height
        })

    doc.close()
    return {
        "total_pages": total_pages,
        "previews": preview_files
    }
