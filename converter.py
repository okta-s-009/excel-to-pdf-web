import os
import sys
import time
import uuid
import re
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


def detect_col_type(text: str) -> Optional[str]:
    """
    Intelligently identifies the semantic type of a column based on its header text.
    Returns: 'jumlah', 'harga', 'total', 'potongan', 'neutral', or None.
    """
    if not text:
        return None
    import re
    t = str(text).strip().lower()
    if re.search(r'\b(jumlah|grand\s*total|total\s*harga|total\s*gaji|total\s*biaya|netto)\b', t):
        return 'jumlah'
    elif re.search(r'\b(harga|tarif|price|rate|cost|upah)\b', t):
        return 'harga'
    elif re.search(r'\b(total|sub\s*total|subtotal)\b', t):
        return 'total'
    elif re.search(r'\b(kasbon|potongan|minus|denda|pph)\b', t):
        return 'potongan'
    elif re.search(r'\b(no|satuan|unit)\b', t):
        return 'neutral'
    return None



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
    Reads workbook structure, sheets, row/col counts, orientation, and page breaks using openpyxl.
    """
    file_path = os.path.abspath(file_path)
    filename = os.path.basename(file_path)
    file_size = os.path.getsize(file_path)

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        details = []
        overall_max_cols = 0
        overall_has_multipage = False
        suggested_orientation = "portrait"

        for name in sheets:
            try:
                ws = wb[name]
                max_r = ws.max_row or 0
                max_c = ws.max_column or 0
                if max_c > overall_max_cols:
                    overall_max_cols = max_c

                # Check orientation & paper setup from openpyxl if available
                ws_orientation = getattr(getattr(ws, "page_setup", None), "orientation", None) or "portrait"
                if ws_orientation == "landscape":
                    suggested_orientation = "landscape"

                # Check if multi-page structure is detected
                is_multipage = max_c > 9 or max_r > 50
                if is_multipage:
                    overall_has_multipage = True

                # Estimate pages count per sheet
                est_pages = 1
                if max_c > 9:
                    est_pages = (max_c + 8) // 9
                if max_r > 45:
                    est_pages *= ((max_r + 44) // 45)

                details.append({
                    "name": name,
                    "max_row": max_r,
                    "max_column": max_c,
                    "orientation": ws_orientation,
                    "has_multipage": is_multipage,
                    "estimated_pages": est_pages
                })
            except Exception:
                details.append({"name": name, "max_row": 0, "max_column": 0, "orientation": "portrait", "has_multipage": False, "estimated_pages": 1})
        wb.close()

        if overall_max_cols > 7 and suggested_orientation != "landscape":
            suggested_orientation = "landscape"

        return {
            "success": True,
            "sheets": sheets,
            "details": details,
            "filename": filename,
            "filesize": file_size,
            "filesize_formatted": format_filesize(file_size),
            "suggested_orientation": suggested_orientation,
            "has_multipage": overall_has_multipage,
            "max_columns": overall_max_cols
        }
    except Exception as e:
        # Fallback for non-standard formats (e.g. .xls)
        return {
            "success": True,
            "sheets": ["Semua Lembar Kerja (Workbook)"],
            "details": [{"name": "Default", "max_row": 0, "max_column": 0, "orientation": "portrait", "has_multipage": False, "estimated_pages": 1}],
            "filename": filename,
            "filesize": file_size,
            "filesize_formatted": format_filesize(file_size),
            "suggested_orientation": "portrait",
            "has_multipage": False,
            "max_columns": 0,
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
    fit_mode: str = "auto",
    center_horizontally: bool = True,
    center_vertically: bool = False,
    sheet_name: str = "all",
    show_gridlines: bool = True,
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
                ps.PrintGridlines = True if show_gridlines is None else bool(show_gridlines)
                ps.BlackAndWhite = False
                ps.Draft = False
                ps.PrintHeadings = bool(show_headings)
            except Exception:
                pass

            # Scaling / Fit to page setup
            try:
                if fit_mode == "auto":
                    # Smart Auto: reads the conditions in Excel so multi-page sheets are neatly separated!
                    v_breaks_count = 0
                    h_breaks_count = 0
                    try:
                        v_breaks_count = ws.VPageBreaks.Count
                    except Exception:
                        pass
                    try:
                        h_breaks_count = ws.HPageBreaks.Count
                    except Exception:
                        pass

                    pages_count = 1
                    try:
                        pages_count = ps.Pages.Count
                    except Exception:
                        pass

                    col_count = 0
                    try:
                        col_count = ws.UsedRange.Columns.Count
                    except Exception:
                        pass

                    orig_zoom = None
                    try:
                        orig_zoom = ps.Zoom
                    except Exception:
                        pass

                    # SMART FIT: Always fit all columns across 1 page width (never slice columns horizontally)
                    ps.Zoom = False
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False

                elif fit_mode == "fit_width":
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
                elif fit_mode == "original":
                    ps.Zoom = 100
                    ps.FitToPagesWide = False
                    ps.FitToPagesTall = False
                else:
                    # Default: smart auto
                    ps.Zoom = 100
                    ps.FitToPagesWide = False
                    ps.FitToPagesTall = False
            except Exception as e:
                print(f"[Converter] Scaling setup error: {e}", file=sys.stderr)

            # Smart formatting: enhance borders, separate tables, and apply smart column colors
            try:
                used_rng = ws.UsedRange
                if used_rng:
                    num_r = used_rng.Rows.Count
                    num_c = used_rng.Columns.Count

                    # 1. Detect all table blocks in the worksheet
                    # Identify real table headers (filtering out total rows and summary lines)
                    header_rows = []
                    for r_idx in range(1, num_r + 1):
                        text_cells = []
                        num_cells = 0
                        kw_count = 0
                        first_text = None
                        for c_idx in range(1, min(num_c + 1, 20)):
                            v = ws.Cells(r_idx, c_idx).Value
                            if v is None:
                                continue
                            v_str = str(v).strip()
                            if not v_str:
                                continue
                            clean = re.sub(r'[Rp.,\s\-]', '', v_str)
                            if clean.isdigit() and len(clean) > 3:
                                num_cells += 1
                            else:
                                text_cells.append(v_str)
                                if first_text is None:
                                    first_text = v_str.lower()
                                if detect_col_type(v_str):
                                    kw_count += 1

                        # Reject total/summary rows
                        if first_text and re.search(r'\b(total|sub\s*total|subtotal|grand\s*total|terbilang|saldo)\b', first_text):
                            continue
                        # Reject data rows that are mostly numbers
                        if num_cells > 0 and num_cells >= len(text_cells):
                            continue
                        if len(text_cells) >= 3 and kw_count >= 1:
                            header_rows.append(r_idx)

                    # If no keyword header found, fallback to first row with >= 2 cells
                    if not header_rows:
                        for r_idx in range(1, min(num_r + 1, 6)):
                            filled_cnt = sum(1 for c_idx in range(1, min(num_c + 1, 15)) if ws.Cells(r_idx, c_idx).Value is not None and str(ws.Cells(r_idx, c_idx).Value).strip() != "")
                            if filled_cnt >= 2:
                                header_rows.append(r_idx)
                                break

                    if not header_rows:
                        header_rows = [1]

                    # Define table blocks: (header_r, end_r)
                    table_blocks = []
                    for h_idx, h_r in enumerate(header_rows):
                        next_h = header_rows[h_idx + 1] if h_idx + 1 < len(header_rows) else num_r + 1
                        # Find end row for this table (stop at 2 consecutive blank rows or before next header)
                        t_end = h_r
                        blank_run = 0
                        for r_scan in range(h_r + 1, next_h):
                            r_has_val = any(ws.Cells(r_scan, c).Value is not None and str(ws.Cells(r_scan, c).Value).strip() != "" for c in range(1, min(num_c + 1, 10)))
                            if not r_has_val:
                                blank_run += 1
                                if blank_run >= 2:
                                    break
                            else:
                                blank_run = 0
                                t_end = r_scan
                        table_blocks.append((h_r, max(t_end, h_r)))

                    # 2. Smart vertical table separation:
                    # Separate distinct tables onto new pages
                    if fit_mode == "auto" and len(table_blocks) > 1:
                        for b_idx in range(1, len(table_blocks)):
                            curr_h = table_blocks[b_idx][0]
                            # Check if there is a table title row above the header (e.g. curr_h - 1 or curr_h - 2)
                            break_at = curr_h
                            for t_above in range(curr_h - 1, max(table_blocks[b_idx - 1][1], curr_h - 3), -1):
                                val_above = ws.Cells(t_above, 1).Value
                                if val_above is not None and str(val_above).strip() != "":
                                    break_at = t_above
                                    break
                            try:
                                ws.HPageBreaks.Add(Before=ws.Rows(break_at))
                            except Exception:
                                pass

                    COLOR_BGR_MAP = {
                        'total': 0xE7FCDC,   # Soft Mint Green (#DCFCE7)
                        'harga': 0xC7F3FE,   # Soft Amber Yellow (#FEF3C7)
                        'jumlah': 0xD5EDFF,  # Soft Warm Peach (#FFEDD5)
                        'potongan': 0xE2E2FE,# Soft Rose Red (#FEE2E2)
                        'neutral': 0xF9F5F1  # Soft Slate Neutral (#F1F5F9)
                    }

                    # 3. Apply smart column colors and crisp borders for each table block
                    for h_r, t_end in table_blocks:
                        # Color columns within this table
                        for c_idx in range(1, num_c + 1):
                            head_val = ws.Cells(h_r, c_idx).Value
                            col_type = detect_col_type(head_val)
                            if col_type:
                                tint_bgr = COLOR_BGR_MAP[col_type]
                                for r_fill in range(h_r + 1, t_end + 1):
                                    target_cell = ws.Cells(r_fill, c_idx)
                                    try:
                                        c_idx_val = target_cell.Interior.ColorIndex
                                        c_col_val = target_cell.Interior.Color
                                        # Only tint cells that do not already have a non-white custom fill
                                        if c_idx_val == -4142 or c_col_val == 16777215:
                                            target_cell.Interior.Color = tint_bgr
                                    except Exception:
                                        pass

                        # Ensure crisp table borders for this table block
                        try:
                            tbl_rng = ws.Range(ws.Cells(h_r, 1), ws.Cells(t_end, num_c))
                            for border_id in [7, 8, 9, 10, 11, 12]:
                                try:
                                    b = tbl_rng.Borders(border_id)
                                    b.LineStyle = 1      # xlContinuous
                                    b.Weight = 2         # xlThin
                                    b.Color = 0xA0A0A0   # Slate Grey border
                                except Exception:
                                    pass
                        except Exception:
                            pass
            except Exception as e:
                print(f"[Converter] Border and Table enhancement error: {e}", file=sys.stderr)

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


def _detect_page_sections(all_rows: List[List[str]], ws, max_cols: int) -> List[List[List[str]]]:
    """
    Smart page/section detection: analyzes an Excel sheet's rows to find
    logical page boundaries. Returns a list of sections, where each section
    is a list of rows that should appear on its own PDF page.

    Detection signals:
    1. Multiple consecutive empty rows (2+ blank rows = section gap)
    2. Column structure change (sudden shift in filled column count)
    3. Title-like row after a total/summary row (new section starting)
    4. Excel horizontal page breaks (from openpyxl sheet properties)
    5. Pattern: Total row → empty rows → Title/Header row = new page
    """
    import re

    if not all_rows:
        return []

    # --- Gather Excel page breaks (row indices, 1-based in openpyxl) ---
    excel_page_break_rows = set()
    try:
        if hasattr(ws, 'row_breaks') and ws.row_breaks:
            for brk in ws.row_breaks.brk:
                if hasattr(brk, 'id'):
                    excel_page_break_rows.add(int(brk.id))
        if hasattr(ws, 'page_breaks') and ws.page_breaks:
            for brk in ws.page_breaks:
                if hasattr(brk, 'id'):
                    excel_page_break_rows.add(int(brk.id))
    except Exception:
        pass

    # --- Build row metadata for analysis ---
    row_meta = []
    for idx, row in enumerate(all_rows):
        filled = [c for c in row if c.strip() != ""]
        filled_count = len(filled)
        is_empty = filled_count == 0
        longest_cell = max((len(c) for c in filled), default=0) if filled else 0

        # Detect if it's a "total" row
        is_total = False
        if filled_count > 0:
            row_text = " ".join(filled).lower()
            is_total = bool(re.search(
                r'\b(total|jumlah|sub\s*total|grand\s*total|amount|sum|saldo|balance)\b',
                row_text, re.I
            ))

        # Detect if it looks like a title (1-2 cells filled, long text, in a wide table)
        is_title_like = False
        if filled_count <= 2 and longest_cell > 15 and max_cols > 3 and not is_empty:
            is_title_like = True
        elif filled_count == 1 and max_cols >= 2 and not is_empty:
            is_title_like = True

        # Detect if it looks like a header row (many cells filled, short text, mostly text)
        is_header_like = False
        if filled_count >= 3 and longest_cell < 40:
            non_numeric_count = 0
            for c in filled:
                clean = re.sub(r'[Rp.,\s\-]', '', c)
                if not clean.isdigit():
                    non_numeric_count += 1
            if non_numeric_count / filled_count > 0.6:
                is_header_like = True

        row_meta.append({
            'idx': idx,
            'filled_count': filled_count,
            'is_empty': is_empty,
            'is_total': is_total,
            'is_title_like': is_title_like,
            'is_header_like': is_header_like,
            'longest_cell': longest_cell
        })

    # --- Find section boundaries ---
    boundaries = set()  # indices where a NEW section starts

    n = len(row_meta)
    for i in range(n):
        # Signal 1: Excel explicit page breaks
        # all_rows is already filtered (no fully-empty rows from original),
        # so we track original row numbers via ws
        if i in excel_page_break_rows:
            boundaries.add(i)

        # Signal 2: Gap of 2+ consecutive empty rows
        if i < n - 1 and row_meta[i]['is_empty']:
            # Count consecutive empties
            empty_run = 1
            j = i + 1
            while j < n and row_meta[j]['is_empty']:
                empty_run += 1
                j += 1
            if empty_run >= 2 and j < n:
                boundaries.add(j)

        # Signal 3: After a total row, if followed by empty then title/header = new section
        if row_meta[i]['is_total'] and i + 1 < n:
            # Look ahead past empty rows
            k = i + 1
            while k < n and row_meta[k]['is_empty']:
                k += 1
            if k < n and (row_meta[k]['is_title_like'] or row_meta[k]['is_header_like']):
                boundaries.add(k)

        # Signal 4: Column structure change (significant shift in filled columns)
        if i > 0 and not row_meta[i]['is_empty'] and not row_meta[i - 1]['is_empty']:
            prev_fc = row_meta[i - 1]['filled_count']
            curr_fc = row_meta[i]['filled_count']
            if prev_fc > 0 and curr_fc > 0:
                ratio = curr_fc / prev_fc if prev_fc > 0 else 1
                # Big structural change + the new row looks like a title or header
                if (ratio < 0.4 or ratio > 2.5) and (row_meta[i]['is_title_like'] or row_meta[i]['is_header_like']):
                    boundaries.add(i)

    # --- Split rows into sections ---
    if not boundaries:
        return [all_rows]

    sorted_boundaries = sorted(boundaries)
    # Filter out boundaries that are empty rows or too close together
    filtered_boundaries = []
    for b in sorted_boundaries:
        if b >= n:
            continue
        if row_meta[b]['is_empty']:
            continue
        if filtered_boundaries and b - filtered_boundaries[-1] < 3:
            continue
        filtered_boundaries.append(b)

    if not filtered_boundaries:
        return [all_rows]

    sections = []
    prev = 0
    for b in filtered_boundaries:
        if b > prev:
            section = [r for r in all_rows[prev:b] if any(c.strip() for c in r)]
            if section:
                sections.append(section)
        prev = b
    # Last section
    if prev < n:
        section = [r for r in all_rows[prev:] if any(c.strip() for c in r)]
        if section:
            sections.append(section)

    # If detection produced only 1 section, return as-is
    if len(sections) <= 1:
        return [all_rows]

    return sections


def _split_section_columns(section_rows: List[List[str]], max_cols_per_page: int = 10) -> List[List[List[str]]]:
    """
    If a table section has too many columns (e.g. 35 cols spanning across 4 pages in 1 sheet),
    splits the table horizontally into column chunks so each chunk fits neatly on a PDF page.
    Repeats the first 1-2 identifying columns (e.g. No, Uraian Pekerjaan) on subsequent chunks.
    """
    if not section_rows:
        return [section_rows]
    total_cols = max(len(r) for r in section_rows)
    if total_cols <= max_cols_per_page:
        return [section_rows]

    # Identify how many fixed header columns to repeat (1 or 2)
    repeat_cols = 1 if total_cols > 10 else 0
    if total_cols > 15:
        repeat_cols = min(2, total_cols)

    chunks = []
    # First chunk: columns 0 to max_cols_per_page
    first_chunk = [r[:max_cols_per_page] for r in section_rows]
    chunks.append(first_chunk)

    col_idx = max_cols_per_page
    step = max(max_cols_per_page - repeat_cols, 4)
    while col_idx < total_cols:
        next_end = min(col_idx + step, total_cols)
        chunk = []
        for r in section_rows:
            repeated = r[:repeat_cols] if repeat_cols > 0 else []
            slice_data = r[col_idx:next_end]
            chunk.append(repeated + slice_data)
        chunks.append(chunk)
        col_idx = next_end

    return chunks


def _render_section_to_story(
    section_rows: List[List[str]],
    story: list,
    styles,
    sheet_idx: int,
    section_idx: int,
    available_width: float,
    sheet_title: str = "",
    is_multi_section: bool = False
):
    """
    Renders a single logical section (detected page) into ReportLab story elements.
    Includes smart title detection, column analysis, and professional table formatting.
    """
    import re
    from reportlab.platypus import Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib import colors

    if not section_rows:
        return

    max_cols = max(len(r) for r in section_rows)
    padded_rows = [r + [""] * (max_cols - len(r)) for r in section_rows]

    title_style = ParagraphStyle(
        f'DocTitle_{sheet_idx}_{section_idx}',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_CENTER,
        spaceAfter=3
    )

    subtitle_style = ParagraphStyle(
        f'DocSubtitle_{sheet_idx}_{section_idx}',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#475569'),
        alignment=TA_CENTER,
        spaceAfter=6
    )

    sheet_heading_style = ParagraphStyle(
        f'SheetHeading_{sheet_idx}_{section_idx}',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#334155'),
        alignment=TA_LEFT,
        spaceAfter=6
    )

    # Smart title detection (scan first 5 rows for title-like patterns)
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
    elif is_multi_section and (section_idx > 0 if isinstance(section_idx, int) else section_idx not in ("0", "0_0")):
        # For subsequent sections without titles, add a subtle label
        pass
    elif sheet_title:
        story.append(Paragraph(f"<b>Sheet: {sheet_title}</b>", sheet_heading_style))

    data_rows = padded_rows[data_start_idx:]
    if not data_rows:
        return

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

    # Dynamic font size based on column count
    font_size = 7.5
    leading = 9.5
    if max_cols > 30:
        font_size = 4.0
        leading = 5.0
    elif max_cols > 20:
        font_size = 4.8
        leading = 6.0
    elif max_cols > 15:
        font_size = 5.5
        leading = 7.0
    elif max_cols > 10:
        font_size = 6.5
        leading = 8.5

    th_style = ParagraphStyle(
        f'TH_{sheet_idx}_{section_idx}',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=font_size,
        leading=leading,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_CENTER
    )

    td_styles = [
        ParagraphStyle(
            f'TD_{sheet_idx}_{section_idx}_{c}',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=font_size,
            leading=leading,
            textColor=colors.HexColor('#1E293B'),
            alignment=col_aligns[c]
        ) for c in range(max_cols)
    ]

    # Build table matrix
    table_matrix = []
    th_cells = [Paragraph(f"<b>{header_row[c]}</b>", th_style) for c in range(max_cols)]
    table_matrix.append(th_cells)

    # Smart column color detection
    COL_PALETTE = {
        'total': {'head': colors.HexColor('#BBF7D0'), 'body': colors.HexColor('#DCFCE7')},
        'harga': {'head': colors.HexColor('#FDE68A'), 'body': colors.HexColor('#FEF3C7')},
        'jumlah': {'head': colors.HexColor('#FED7AA'), 'body': colors.HexColor('#FFEDD5')},
        'potongan': {'head': colors.HexColor('#FECDD3'), 'body': colors.HexColor('#FEE2E2')},
        'neutral': {'head': colors.HexColor('#E2E8F0'), 'body': colors.HexColor('#F1F5F9')}
    }
    col_types = [detect_col_type(header_row[c] if c < len(header_row) else "") for c in range(max_cols)]

    pad = 1.0 if max_cols > 30 else (1.5 if max_cols > 15 else 3.0)
    table_style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
        ('BOX', (0, 0), (-1, -1), 0.8, colors.HexColor('#475569')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#94A3B8')),
        ('LINEBELOW', (0, 0), (-1, 0), 1.2, colors.HexColor('#334155')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), pad),
        ('BOTTOMPADDING', (0, 0), (-1, -1), pad),
        ('LEFTPADDING', (0, 0), (-1, -1), pad),
        ('RIGHTPADDING', (0, 0), (-1, -1), pad),
    ]

    # Apply smart header colors
    for c in range(max_cols):
        ctype = col_types[c]
        if ctype and ctype in COL_PALETTE:
            table_style_commands.append(('BACKGROUND', (c, 0), (c, 0), COL_PALETTE[ctype]['head']))

    for ri, r in enumerate(body_rows):
        row_cells = []
        is_total_row = any(
            re.search(r'\b(total|jumlah|sub\s*total|grand\s*total|amount|saldo|balance)\b', c, re.I)
            for c in r if c
        )
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
            # Base alternating row background
            bg = colors.HexColor('#F8FAFC') if matrix_row_idx % 2 == 1 else colors.white
            table_style_commands.append(('BACKGROUND', (0, matrix_row_idx), (-1, matrix_row_idx), bg))
            # Overlay smart column color tints
            for c in range(max_cols):
                ctype = col_types[c]
                if ctype and ctype in COL_PALETTE:
                    table_style_commands.append(('BACKGROUND', (c, matrix_row_idx), (c, matrix_row_idx), COL_PALETTE[ctype]['body']))

    table = Table(table_matrix, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(table_style_commands))
    story.append(table)
    story.append(Spacer(1, 15))


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
    """
    ReportLab-based Excel→PDF converter with smart multi-page detection.
    Automatically detects logical page boundaries within a single sheet and
    splits them into separate PDF pages for clean, professional output.
    """
    import re
    from reportlab.lib.pagesizes import A4, LETTER, LEGAL, A3, landscape, portrait
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
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

    # Pre-scan all sheets for global max columns (for auto-landscape detection)
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

    sheet_heading_style = ParagraphStyle(
        'SheetHeadingMain',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#334155'),
        alignment=TA_LEFT,
        spaceAfter=6
    )

    is_first_page = True

    for sheet_idx, ws in enumerate(sheets):
        # Read ALL rows from worksheet (including empty ones for gap detection)
        all_raw_rows = []
        for row in ws.iter_rows(values_only=True):
            cleaned = [str(c).strip() if c is not None else "" for c in row]
            all_raw_rows.append(cleaned)

        # Remove trailing empty rows
        while all_raw_rows and not any(c.strip() for c in all_raw_rows[-1]):
            all_raw_rows.pop()

        if not all_raw_rows or not any(any(c.strip() for c in r) for r in all_raw_rows):
            if not is_first_page:
                story.append(PageBreak())
            story.append(Paragraph(f"<b>Sheet: {ws.title}</b> (Lembar kerja kosong)", sheet_heading_style))
            is_first_page = False
            continue

        max_cols = max(len(r) for r in all_raw_rows) if all_raw_rows else 0

        # ═══════════════════════════════════════════════════════════════
        # SMART PAGE DETECTION: Split single sheet into logical sections
        # ═══════════════════════════════════════════════════════════════
        sections = _detect_page_sections(all_raw_rows, ws, max_cols)

        num_sections = len(sections)
        is_multi_section = num_sections > 1

        if is_multi_section:
            print(f"[Converter] Smart detection: Sheet '{ws.title}' -> {num_sections} logical pages detected", flush=True)

        for sec_idx, section_rows in enumerate(sections):
            if not is_first_page:
                story.append(PageBreak())
            is_first_page = False

            show_sheet_title = (len(sheets) > 1 and sec_idx == 0)

            _render_section_to_story(
                section_rows=section_rows,
                story=story,
                styles=styles,
                sheet_idx=sheet_idx,
                section_idx=sec_idx,
                available_width=available_width,
                sheet_title=ws.title if show_sheet_title else "",
                is_multi_section=is_multi_section
            )

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
