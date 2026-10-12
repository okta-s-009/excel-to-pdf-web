import os
import sys

# Tambahkan root directory ke sys.path agar app.py dan converter.py dapat di-import
current_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = os.path.dirname(current_dir)
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

# Import FastAPI app dari root app.py
from app import app
