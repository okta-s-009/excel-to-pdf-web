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
    import re
    from reportlab.lib.pagesizes import A4, LETTER, LEGAL, A3, landscape, portrait
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib import colors

    size_map = {
        "A4": A4,
        "Letter": LETTER,
        "Legal": LEGAL,
        "A3": A3
    }
    base_size = size_map.get(paper_size, A4)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if sheet_name and sheet_name.lower() != "all" and sheet_name in wb.sheetnames:
        sheets = [wb[sheet_name]]
    else:
        sheets = [wb[s] for s in wb.sheetnames]

    # Pre-scan sheets to find global max columns for auto-landscape
    global_max_cols = 0
    for ws in sheets:
        for row in ws.iter_rows(values_only=True):
            non_empty = [c for c in row if c is not None and str(c).strip() != ""]
            if len(non_empty) > global_max_cols:
                global_max_cols = len(non_empty)

    final_orientation = orientation.lower()
    if final_orientation == "portrait" and global_max_cols > 7:
        final_orientation = "landscape"

    pagesize = landscape(base_size) if final_orientation == "landscape" else portrait(base_size)
    available_width = max(pagesize[0] - left_margin_pt - right_margin_pt, 200)

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
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_CENTER,
        spaceAfter=3
    )

    subtitle_style = ParagraphStyle(
        'DocSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#475569'),
        alignment=TA_CENTER,
        spaceAfter=6
    )

    sheet_heading_style = ParagraphStyle(
        'SheetHeading',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#334155'),
        alignment=TA_LEFT,
        spaceAfter=6
    )

    for sheet_idx, ws in enumerate(sheets):
        if sheet_idx > 0:
            story.append(PageBreak())

        raw_rows = []
        for row in ws.iter_rows(values_only=True):
            cleaned = [str(c).strip() if c is not None else "" for c in row]
            if any(cleaned):
                raw_rows.append(cleaned)

        if not raw_rows:
            story.append(Paragraph(f"<b>Sheet: {ws.title}</b> (Lembar kerja kosong)", sheet_heading_style))
            continue

        max_cols = max(len(r) for r in raw_rows)
        padded_rows = [r + [""] * (max_cols - len(r)) for r in raw_rows]

        # Smart title detection
        title_lines = []
        data_start_idx = 0
        for ri in range(min(len(padded_rows), 5)):
            r = padded_rows[ri]
            filled = [c for c in r if c != ""]
            if len(filled) == 0:
                data_start_idx = ri + 1
                continue
            longest = max((len(c) for c in filled), default=0)
            if (len(filled) == 1 and max_cols >= 2) or (len(filled) <= 2 and longest > 15 and max_cols > 3):
                title_lines.append(" ".join(filled))
                data_start_idx = ri + 1
            else:
                break

        # Render titles above table
        if title_lines:
            for ti, t_text in enumerate(title_lines):
                st = title_style if ti == 0 else subtitle_style
                story.append(Paragraph(t_text, st))
            story.append(Spacer(1, 4))
        elif len(sheets) > 1:
            story.append(Paragraph(f"<b>Sheet: {ws.title}</b>", sheet_heading_style))

        data_rows = padded_rows[data_start_idx:]
        if not data_rows:
            continue

        header_row = data_rows[0]
        body_rows = data_rows[1:]

        # Column weights & alignments
        col_weights = []
        col_aligns = []
        for c in range(max_cols):
            h_val = header_row[c] if c < len(header_row) else ""
            h_len = len(h_val)

            num_count = 0
            sample_count = 0
            max_len = h_len
            for r in body_rows[:30]:
                val = r[c] if c < len(r) else ""
                if not val:
                    continue
                sample_count += 1
                if len(val) > max_len:
                    max_len = len(val)
                clean_num = re.sub(r'[Rp.,\s]', '', val)
                if clean_num.isdigit() or (clean_num.startswith('-') and clean_num[1:].isdigit()):
                    num_count += 1

            is_numeric = sample_count > 0 and (num_count / sample_count) > 0.6
            is_index = is_numeric and max_len <= 4

            # Ensure header has enough width so single words like 'Qty' or 'No' never wrap vertically
            min_header_w = max(h_len * 1.1, 4.0)

            if is_index:
                col_aligns.append(TA_CENTER)
                col_weights.append(max(min_header_w, 3.5))
            elif is_numeric:
                col_aligns.append(TA_RIGHT)
                col_weights.append(max(max_len * 0.85, min_header_w, 5.0))
            else:
                col_aligns.append(TA_LEFT)
                col_weights.append(max(max_len * 0.9, min_header_w, 6.0))

        total_weight = sum(col_weights) or 1
        col_widths = [(w / total_weight) * available_width for w in col_weights]

        # Dynamic font size
        font_size = 7.5
        leading = 9.5
        if max_cols > 15:
            font_size = 5.5
            leading = 7.0
        elif max_cols > 10:
            font_size = 6.5
            leading = 8.5

        th_style = ParagraphStyle(
            'TH',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=font_size,
            leading=leading,
            textColor=colors.HexColor('#0F172A'),
            alignment=TA_CENTER
        )

        td_styles = [
            ParagraphStyle(
                f'TD_{sheet_idx}_{c}',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=font_size,
                leading=leading,
                textColor=colors.HexColor('#1E293B'),
                alignment=col_aligns[c]
            ) for c in range(max_cols)
        ]

        table_matrix = []
        th_cells = [Paragraph(f"<b>{header_row[c]}</b>", th_style) for c in range(max_cols)]
        table_matrix.append(th_cells)

        table_style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
            ('LINEBELOW', (0, 0), (-1, 0), 1.0, colors.HexColor('#94A3B8')),
            ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#CBD5E1')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ]

        for ri, r in enumerate(body_rows):
            row_cells = []
            is_total_row = any(re.search(r'\b(total|jumlah|sub\s*total|grand\s*total)\b', c, re.I) for c in r)
            for c in range(max_cols):
                val = r[c] if c < len(r) else ""
                if is_total_row and val:
                    cell_p = Paragraph(f"<b>{val}</b>", td_styles[c])
                else:
                    cell_p = Paragraph(val or "", td_styles[c])
                row_cells.append(cell_p)
            table_matrix.append(row_cells)

            matrix_row_idx = ri + 1
            if is_total_row:
                table_style_commands.append(('BACKGROUND', (0, matrix_row_idx), (-1, matrix_row_idx), colors.HexColor('#F1F5F9')))
                table_style_commands.append(('LINEABOVE', (0, matrix_row_idx), (-1, matrix_row_idx), 0.8, colors.HexColor('#94A3B8')))
                table_style_commands.append(('LINEBELOW', (0, matrix_row_idx), (-1, matrix_row_idx), 1.2, colors.HexColor('#334155')))
            else:
                bg = colors.HexColor('#F8FAFC') if matrix_row_idx % 2 == 1 else colors.white
                table_style_commands.append(('BACKGROUND', (0, matrix_row_idx), (-1, matrix_row_idx), bg))

        table = Table(table_matrix, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle(table_style_commands))
        story.append(table)
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
