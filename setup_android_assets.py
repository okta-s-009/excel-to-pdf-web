import os
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
ANDROID_DIR = os.path.join(BASE_DIR, "android")
APP_DIR = os.path.join(ANDROID_DIR, "app")
MAIN_DIR = os.path.join(APP_DIR, "src", "main")
ASSETS_WWW = os.path.join(MAIN_DIR, "assets", "www")

# Ensure target directories exist
for p in [
    os.path.join(ANDROID_DIR, "gradle", "wrapper"),
    os.path.join(MAIN_DIR, "java", "com", "excel2pdf", "studio"),
    os.path.join(MAIN_DIR, "res", "values"),
    os.path.join(MAIN_DIR, "res", "layout"),
    os.path.join(MAIN_DIR, "res", "xml"),
    os.path.join(MAIN_DIR, "res", "drawable"),
    os.path.join(MAIN_DIR, "res", "mipmap-hdpi"),
    os.path.join(MAIN_DIR, "res", "mipmap-xhdpi"),
    os.path.join(MAIN_DIR, "res", "mipmap-xxhdpi"),
    os.path.join(MAIN_DIR, "res", "mipmap-xxxhdpi"),
    ASSETS_WWW
]:
    os.makedirs(p, exist_ok=True)

# Copy static assets to assets/www
if os.path.exists(ASSETS_WWW):
    shutil.rmtree(ASSETS_WWW)
shutil.copytree(STATIC_DIR, ASSETS_WWW)

# Copy icons to mipmap
icon_path = os.path.join(STATIC_DIR, "icon-192.png")
if os.path.exists(icon_path):
    for m in ["mipmap-hdpi", "mipmap-xhdpi", "mipmap-xxhdpi", "mipmap-xxxhdpi"]:
        shutil.copyfile(icon_path, os.path.join(MAIN_DIR, "res", m, "ic_launcher.png"))
        shutil.copyfile(icon_path, os.path.join(MAIN_DIR, "res", m, "ic_launcher_round.png"))

print("Assets copied successfully to Android project!")
