import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, colorchooser, filedialog
import tkinter.font as tkfont
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import shutil
import os
import json
import re
import urllib.parse
import urllib.request
import urllib.error
import ssl
import base64
import uuid
import webbrowser
import sys
import subprocess
import tempfile
import threading
import time
from datetime import datetime

# --- Dawny, wlasny backend Freyonatora (Supabase) - uzywany TYLKO do
# sprawdzania nowych wersji programu (patrz _sprawdz_aktualizacje). Program
# NIE jest juz aktywowany kodem licencyjnym przez ten projekt - patrz
# "Aktywacja programu przez konto KOS" nizej.
FREYA_SUPABASE_URL = "https://dlffxncncnnyufdsgnxs.supabase.co"
FREYA_SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRsZmZ4bmNuY25ueXVmZHNnbnhzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODg5NzM0NjIsImV4cCI6MjEwNDU0OTQ2Mn0.mZVeHOQJvIpgpesqcmYqjBPrz6X9FEQf3sfJBMj1Rxo"
APP_WERSJA = "1.0.0"

PLACEHOLDER_PART1_ONLY = True
