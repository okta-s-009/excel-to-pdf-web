"""
Creates a test Excel file with 4 logical pages/sections in a SINGLE sheet.
This simulates a common scenario where users put multiple reports/tables
into one worksheet separated by empty rows.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Laporan Gabungan"

thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)
header_fill = PatternFill(start_color="10B981", end_color="10B981", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
title_font = Font(name="Calibri", size=14, bold=True, color="1E293B")
subtitle_font = Font(name="Calibri", size=10, italic=True, color="64748B")
total_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
total_font = Font(name="Calibri", size=11, bold=True, color="1E293B")

row_cursor = 1

# ═══════════════════════════════════════════════════════
# PAGE 1: Laporan Penjualan Q1
# ═══════════════════════════════════════════════════════
ws.merge_cells(f"A{row_cursor}:F{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="PT NUSANTARA JAYA ABADI - LAPORAN PENJUALAN")
c.font = title_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

ws.merge_cells(f"A{row_cursor}:F{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="Periode: Kuartal 1 (Januari - Maret 2026)")
c.font = subtitle_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

row_cursor += 1  # blank row

headers1 = ["No", "Kode", "Nama Produk", "Qty", "Harga Satuan", "Total"]
for ci, h in enumerate(headers1, 1):
    cell = ws.cell(row=row_cursor, column=ci, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border
row_cursor += 1

data1 = [
    [1, "PRD-001", "Cloud Server Enterprise 64GB", 15, 4500000, 67500000],
    [2, "PRD-002", "Lisensi Software ERP Core v4.2", 30, 2250000, 67500000],
    [3, "PRD-003", "Integrasi API Gateway FinTech", 8, 12000000, 96000000],
    [4, "PRD-004", "Maintenance & SLA Support", 24, 3500000, 84000000],
    [5, "PRD-005", "Keamanan Siber & Firewall", 5, 18500000, 92500000],
]
for rd in data1:
    for ci, val in enumerate(rd, 1):
        cell = ws.cell(row=row_cursor, column=ci, value=val)
        cell.border = thin_border
        if ci >= 4:
            cell.number_format = '#,##0'
    row_cursor += 1

# Total row
ws.merge_cells(f"A{row_cursor}:E{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="TOTAL PENJUALAN Q1")
c.font = total_font
c.fill = total_fill
c.alignment = Alignment(horizontal="right")
c = ws.cell(row=row_cursor, column=6, value=407500000)
c.font = Font(name="Calibri", size=11, bold=True, color="059669")
c.fill = total_fill
c.number_format = '#,##0'
row_cursor += 1

# ═══════════════════════════════════════════════════════
# GAP: 3 empty rows (signals a section break)
# ═══════════════════════════════════════════════════════
row_cursor += 3

# ═══════════════════════════════════════════════════════
# PAGE 2: Laporan Biaya Operasional
# ═══════════════════════════════════════════════════════
ws.merge_cells(f"A{row_cursor}:E{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="REKAPITULASI BIAYA OPERASIONAL PERUSAHAAN")
c.font = title_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

ws.merge_cells(f"A{row_cursor}:E{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="Periode: Semester 1 Tahun 2026")
c.font = subtitle_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

row_cursor += 1  # blank row

headers2 = ["No", "Item Biaya", "Anggaran (Rp)", "Realisasi (Rp)", "Persentase"]
for ci, h in enumerate(headers2, 1):
    cell = ws.cell(row=row_cursor, column=ci, value=h)
    cell.fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border
row_cursor += 1

data2 = [
    [1, "Gaji & Tunjangan Karyawan", 250000000, 245000000, "98%"],
    [2, "Operasional Server & Cloud", 85000000, 81000000, "95%"],
    [3, "Pemasaran & Kampanye Digital", 45000000, 43500000, "96%"],
    [4, "Sewa Kantor & Utilitas", 120000000, 118500000, "99%"],
    [5, "Transportasi & Perjalanan Dinas", 35000000, 28000000, "80%"],
    [6, "Pelatihan & Pengembangan SDM", 30000000, 27500000, "92%"],
]
for rd in data2:
    for ci, val in enumerate(rd, 1):
        cell = ws.cell(row=row_cursor, column=ci, value=val)
        cell.border = thin_border
        if ci in [3, 4]:
            cell.number_format = '#,##0'
    row_cursor += 1

ws.merge_cells(f"A{row_cursor}:D{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="TOTAL BIAYA OPERASIONAL")
c.font = total_font
c.fill = total_fill
c.alignment = Alignment(horizontal="right")
c = ws.cell(row=row_cursor, column=5, value="94%")
c.font = total_font
c.fill = total_fill
row_cursor += 1

# ═══════════════════════════════════════════════════════
# GAP: 3 empty rows
# ═══════════════════════════════════════════════════════
row_cursor += 3

# ═══════════════════════════════════════════════════════
# PAGE 3: Daftar Inventaris Aset
# ═══════════════════════════════════════════════════════
ws.merge_cells(f"A{row_cursor}:G{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="DAFTAR INVENTARIS ASET PERUSAHAAN")
c.font = title_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

ws.merge_cells(f"A{row_cursor}:G{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="Per Tanggal: 30 Juni 2026")
c.font = subtitle_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

row_cursor += 1  # blank

headers3 = ["No", "Kode Aset", "Nama Aset", "Lokasi", "Kondisi", "Nilai Perolehan", "Nilai Buku"]
for ci, h in enumerate(headers3, 1):
    cell = ws.cell(row=row_cursor, column=ci, value=h)
    cell.fill = PatternFill(start_color="8B5CF6", end_color="8B5CF6", fill_type="solid")
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border
row_cursor += 1

data3 = [
    [1, "AST-001", "Server Rack Dell R750", "Data Center Lt.3", "Baik", 250000000, 200000000],
    [2, "AST-002", "Laptop Lenovo ThinkPad X1", "Divisi IT", "Baik", 25000000, 18000000],
    [3, "AST-003", "UPS APC 10kVA", "Data Center Lt.3", "Baik", 45000000, 32000000],
    [4, "AST-004", "Printer HP LaserJet Pro", "Administrasi", "Perlu Service", 8500000, 4000000],
    [5, "AST-005", "Proyektor Epson EB-X51", "Ruang Meeting", "Baik", 12000000, 8500000],
    [6, "AST-006", "AC Daikin Split 2PK", "Kantor Utama", "Baik", 15000000, 10000000],
    [7, "AST-007", "Meja Kerja Ergonomic", "Semua Divisi", "Baik", 5000000, 3500000],
]
for rd in data3:
    for ci, val in enumerate(rd, 1):
        cell = ws.cell(row=row_cursor, column=ci, value=val)
        cell.border = thin_border
        if ci in [6, 7]:
            cell.number_format = '#,##0'
    row_cursor += 1

ws.merge_cells(f"A{row_cursor}:F{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="TOTAL NILAI ASET")
c.font = total_font
c.fill = total_fill
c.alignment = Alignment(horizontal="right")
c = ws.cell(row=row_cursor, column=7, value=276000000)
c.font = Font(name="Calibri", size=11, bold=True, color="7C3AED")
c.fill = total_fill
c.number_format = '#,##0'
row_cursor += 1

# ═══════════════════════════════════════════════════════
# GAP: 3 empty rows
# ═══════════════════════════════════════════════════════
row_cursor += 3

# ═══════════════════════════════════════════════════════
# PAGE 4: Data Karyawan
# ═══════════════════════════════════════════════════════
ws.merge_cells(f"A{row_cursor}:F{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="DATA KARYAWAN AKTIF PT NUSANTARA JAYA ABADI")
c.font = title_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

ws.merge_cells(f"A{row_cursor}:F{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="Update Terakhir: Juli 2026")
c.font = subtitle_font
c.alignment = Alignment(horizontal="center")
row_cursor += 1

row_cursor += 1  # blank

headers4 = ["No", "NIP", "Nama Lengkap", "Jabatan", "Divisi", "Masa Kerja"]
for ci, h in enumerate(headers4, 1):
    cell = ws.cell(row=row_cursor, column=ci, value=h)
    cell.fill = PatternFill(start_color="F59E0B", end_color="F59E0B", fill_type="solid")
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border
row_cursor += 1

data4 = [
    [1, "NIP-2020-001", "Budi Santoso, S.Kom", "IT Manager", "Teknologi Informasi", "6 Tahun"],
    [2, "NIP-2021-015", "Dewi Lestari, S.E", "Finance Lead", "Keuangan", "5 Tahun"],
    [3, "NIP-2022-008", "Ahmad Hidayat", "DevOps Engineer", "Teknologi Informasi", "4 Tahun"],
    [4, "NIP-2023-003", "Sari Mulyani, M.M", "HR Manager", "SDM", "3 Tahun"],
    [5, "NIP-2024-001", "Rizki Pratama", "Junior Developer", "Teknologi Informasi", "2 Tahun"],
    [6, "NIP-2024-012", "Nadia Putri", "Marketing Specialist", "Pemasaran", "2 Tahun"],
]
for rd in data4:
    for ci, val in enumerate(rd, 1):
        cell = ws.cell(row=row_cursor, column=ci, value=val)
        cell.border = thin_border
    row_cursor += 1

ws.merge_cells(f"A{row_cursor}:E{row_cursor}")
c = ws.cell(row=row_cursor, column=1, value="JUMLAH KARYAWAN AKTIF")
c.font = total_font
c.fill = total_fill
c.alignment = Alignment(horizontal="right")
c = ws.cell(row=row_cursor, column=6, value="6 Orang")
c.font = total_font
c.fill = total_fill

# Auto column widths
from openpyxl.utils import get_column_letter
for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

wb.save("test_4_pages_in_1_sheet.xlsx")
print("[OK] File 'test_4_pages_in_1_sheet.xlsx' berhasil dibuat!")
print("   -> Berisi 4 laporan/tabel berbeda dalam 1 sheet tunggal")
print("   -> Dipisahkan oleh baris kosong dan judul berbeda")
