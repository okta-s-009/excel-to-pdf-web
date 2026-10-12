import os
import sys
import socket
import webbrowser
import threading
import time
import uvicorn

def get_local_ip():
    """Mendapatkan alamat IP lokal komputer di jaringan Wi-Fi/LAN."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"

def open_browser(url, delay=1.2):
    time.sleep(delay)
    try:
        webbrowser.open(url)
    except Exception:
        pass

def main():
    host = "0.0.0.0"
    port = 8000
    local_ip = get_local_ip()

    local_url = f"http://localhost:{port}"
    mobile_url = f"http://{local_ip}:{port}"

    print("=" * 68)
    print("      EXCEL2PDF STUDIO - WEB APP (KOMPUTER, HP & TABLET)")
    print("=" * 68)
    print("  [1] AKSES DI KOMPUTER / LAPTOP INI:")
    print(f"      👉  {local_url}")
    print()
    print("  [2] AKSES DARI HP ATAU TABLET (ANDROID / iPAD / IPHONE):")
    print(f"      👉  {mobile_url}")
    print()
    print("      * Petunjuk HP/Tablet:")
    print("        1. Pastikan HP terhubung ke Wi-Fi yang sama dengan PC ini.")
    print(f"        2. Buka browser (Chrome / Safari) di HP, ketik: {mobile_url}")
    print("        3. Bisa klik menu browser 'Tambahkan ke Layar Utama'")
    print("           (Add to Home screen) agar menjadi aplikasi mandiri di HP!")
    print("=" * 68)
    print("  [*] Tekan CTRL + C untuk menghentikan server.")
    print("=" * 68)
    print()

    # Buka browser otomatis di PC
    threading.Thread(target=open_browser, args=(local_url,), daemon=True).start()

    # Jalankan server uvicorn pada semua interface (0.0.0.0)
    uvicorn.run("app:app", host=host, port=port, log_level="info", reload=False)

if __name__ == "__main__":
    main()
