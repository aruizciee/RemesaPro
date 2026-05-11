import pandas as pd
import pypdf
import os
import re
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
import threading
import sys
import json
import unicodedata
import platform
import subprocess
from urllib import request as urllib_request
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent
import ssl

# Pre-compiled regex patterns
_RE_DECIMAL_AMOUNT = re.compile(r"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})")
_RE_THOUSANDS_NODEC = re.compile(r"\b(\d{1,3}(?:\.\d{3})+)\b")  # e.g. "1.256" → 1256
_RE_WHOLE_EURO = re.compile(r"(\d+)\s*€")
_RE_NOMBRE = re.compile(r"[Nn]ombre:\s*(.+)")

APP_VERSION = 7  # Matches GitHub build number

# Configuration defaults
DEFAULT_DB_FILE = "Base datos IBAN proveedores.xlsx"
TEMPLATE_FILE = "FA25_REMESA PAGOS SANTANDER_.xlsx"
OUTPUT_PREFIX = "REMESA_GENERADA_"
CONFIG_FILE = "remesa_config.json"
SESSION_FILE = "remesa_session.json"
LOGO_FILE = "ciee logo.png"
ICON_ICO  = "icon.ico"
ICON_PNG  = "icon.png"

# SEPA debtor fields (values loaded from local remesa_config.json, never from code)
SEPA_DEFAULTS = {
    "sepa_nombre": "",
    "sepa_cif": "",
    "sepa_iban": "",
    "sepa_bic": "",
    "sepa_direccion": "",
    "sepa_cp": "",
    "sepa_ciudad": "",
    "sepa_provincia": "",
    "sepa_pais": "ES",
}