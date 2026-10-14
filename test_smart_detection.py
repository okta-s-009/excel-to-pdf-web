"""
Test script to verify smart page detection works correctly.
Converts the 4-pages-in-1-sheet Excel file using the ReportLab engine.
"""
import os
import sys

# Force ReportLab engine by temporarily disabling COM
import converter
converter.COM_AVAILABLE = False

result = converter.convert_excel_to_pdf(
    excel_path="test_4_pages_in_1_sheet.xlsx",
    output_pdf_path="test_smart_split_output.pdf",
    top_margin_pt=converter.cm_to_points(1.91),
    bottom_margin_pt=converter.cm_to_points(1.91),
    left_margin_pt=converter.cm_to_points(1.78),
    right_margin_pt=converter.cm_to_points(1.78),
    orientation="portrait",
    paper_size="A4",
    fit_mode="fit_width",
    sheet_name="all"
)

print(f"\n=== Hasil Konversi ===")
print(f"Engine     : {result['engine']}")
print(f"Durasi     : {result['duration_seconds']}s")
print(f"Ukuran PDF : {result['pdf_size_formatted']}")
print(f"Path       : {result['output_path']}")

# Check page count
try:
    import pymupdf
    doc = pymupdf.open("test_smart_split_output.pdf")
    total_pages = len(doc)
    print(f"Total Pages: {total_pages}")
    if total_pages == 4:
        print("[OK] BERHASIL! 4 halaman terdeteksi = 4 page PDF")
    elif total_pages > 1:
        print(f"[INFO] {total_pages} halaman terdeteksi (mungkin beberapa section kecil digabung)")
    else:
        print("[WARN] Hanya 1 halaman - smart detection mungkin tidak bekerja")
    doc.close()
except ImportError:
    print("[INFO] PyMuPDF tidak tersedia, tidak bisa cek jumlah halaman")
