import sys
import os

# Menambahkan direktori saat ini ke path Python agar file app.py bisa dibaca
sys.path.insert(0, os.path.dirname(__file__))

# Mengimpor aplikasi Flask (variabel 'app') dari file 'app.py' dan mengubah namanya menjadi 'application'
from app import app as application
