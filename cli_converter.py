import argparse
import os
import sys
import converter

def main():
    parser = argparse.ArgumentParser(
        description="Konversi file Excel (.xlsx, .xls) ke PDF dengan pengaturan margin kustom."
    )
    parser.add_argument("input", help="Jalur file Excel sumber (contoh: data.xlsx)")
    parser.add_argument("output", nargs="?", default=None, help="Jalur file PDF tujuan (opsional, default: nama_file.pdf)")
    parser.add_argument("-o", "--output-flag", dest="output_flag", default=None, help="Jalur file PDF tujuan alternatif")
    parser.add_argument("--unit", choices=["cm", "inch"], default="cm", help="Satuan ukuran margin (default: cm)")
    parser.add_argument("--preset", choices=["normal", "narrow", "wide"], default=None, help="Preset margin cepat")
    parser.add_argument("--top", type=float, default=None, help="Margin atas (top)")
    parser.add_argument("--bottom", type=float, default=None, help="Margin bawah (bottom)")
    parser.add_argument("--left", type=float, default=None, help="Margin kiri (left)")
    parser.add_argument("--right", type=float, default=None, help="Margin kanan (right)")
    parser.add_argument("--orientation", choices=["portrait", "landscape"], default="portrait", help="Orientasi kertas")
    parser.add_argument("--paper", choices=["A4", "Letter", "Legal", "A3"], default="A4", help="Ukuran kertas")
    parser.add_argument("--fit", choices=["fit_width", "fit_page", "original"], default="fit_width", help="Penskalaan halaman")
    parser.add_argument("--sheet", default="all", help="Nama sheet tertentu atau 'all' untuk semua sheet")

    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"[ERROR] File input tidak ditemukan: {input_path}", file=sys.stderr)
        sys.exit(1)

    target_out = args.output_flag or args.output
    if not target_out:
        base_name = os.path.splitext(input_path)[0]
        output_path = f"{base_name}.pdf"
    else:
        output_path = os.path.abspath(target_out)

    # Determine margin values
    unit = args.unit.lower()
    if args.preset == "narrow":
        top = 1.27 if unit == "cm" else 0.50
        bot = 1.27 if unit == "cm" else 0.50
        left = 0.64 if unit == "cm" else 0.25
        right = 0.64 if unit == "cm" else 0.25
    elif args.preset == "wide":
        top = 2.54 if unit == "cm" else 1.00
        bot = 2.54 if unit == "cm" else 1.00
        left = 2.54 if unit == "cm" else 1.00
        right = 2.54 if unit == "cm" else 1.00
    else:
        # Default or Normal
        top = 1.91 if unit == "cm" else 0.75
        bot = 1.91 if unit == "cm" else 0.75
        left = 1.78 if unit == "cm" else 0.70
        right = 1.78 if unit == "cm" else 0.70

    if args.top is not None: top = args.top
    if args.bottom is not None: bot = args.bottom
    if args.left is not None: left = args.left
    if args.right is not None: right = args.right

    if unit == "inch":
        top_pt = converter.inch_to_points(top)
        bot_pt = converter.inch_to_points(bot)
        left_pt = converter.inch_to_points(left)
        right_pt = converter.inch_to_points(right)
    else:
        top_pt = converter.cm_to_points(top)
        bot_pt = converter.cm_to_points(bot)
        left_pt = converter.cm_to_points(left)
        right_pt = converter.cm_to_points(right)

    print("=" * 60)
    print("           EXCEL TO PDF CONVERTER (CLI)")
    print("=" * 60)
    print(f"File Sumber  : {input_path}")
    print(f"File Tujuan  : {output_path}")
    print(f"Margin       : Atas={top}{unit}, Bawah={bot}{unit}, Kiri={left}{unit}, Kanan={right}{unit}")
    print(f"Orientasi    : {args.orientation.upper()}")
    print(f"Ukuran Kertas: {args.paper}")
    print(f"Penskalaan   : {args.fit}")
    print(f"Sheet        : {args.sheet}")
    print("-" * 60)
    print("[*] Memulai konversi...")

    try:
        # Set utf-8 stdout if possible on Windows
        if sys.platform == "win32":
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    try:
        res = converter.convert_excel_to_pdf(
            excel_path=input_path,
            output_pdf_path=output_path,
            top_margin_pt=top_pt,
            bottom_margin_pt=bot_pt,
            left_margin_pt=left_pt,
            right_margin_pt=right_pt,
            orientation=args.orientation,
            paper_size=args.paper,
            fit_mode=args.fit,
            sheet_name=args.sheet
        )
        print(f"[SUCCESS] Berhasil dikonversi dalam {res['duration_seconds']} detik!")
        print(f"[INFO] Ukuran PDF: {res['pdf_size_formatted']}")
        print(f"[INFO] Engine    : {res['engine']}")
        print(f"[INFO] File PDF  : {output_path}")
    except Exception as e:
        print(f"[ERROR] Terjadi kesalahan: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
