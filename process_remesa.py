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
# Importe con decimales: "1.234,56" / "1,234.56" / "28,92".
# Los límites (?<![\d.,]) y (?!\d) evitan cortar un número por la mitad:
# sin ellos "1.998 €" (mil novecientos noventa y ocho) se leía como "1.99".
_RE_DECIMAL_AMOUNT = re.compile(r"(?<![\d.,])(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})(?!\d)")
# Solo miles, sin parte decimal: "1.256" → 1256. El (?![.,]?\d) evita que
# "1.234,56" genere además un candidato erróneo de 1234.
_RE_THOUSANDS_NODEC = re.compile(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+)(?![.,]?\d)")
_RE_WHOLE_EURO = re.compile(r"(\d+)\s*€")
_RE_NOMBRE = re.compile(r"[Nn]ombre:\s*(.+)")

APP_VERSION = 7  # Matches GitHub build number

# Configuration defaults
DEFAULT_DB_FILE = "Base datos IBAN proveedores.xlsx"
TEMPLATE_FILE = "FA25_REMESA PAGOS SANTANDER_.xlsx"
OUTPUT_PREFIX = "REMESA_GENERADA_"
DEFAULT_CONCEPT = "NOTA DE GASTOS"  # concepto que ve el proveedor si la BD no tiene uno
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

def parse_amount(value):
    """
    Parse a numeric value from Excel cell or PDF text string.
    Handles all separator combinations:
      - 28.92      → 28.92   (punto = decimal, 2 dígitos)
      - 28,92      → 28.92   (coma = decimal)
      - 1.256      → 1256.0  (punto = miles, exactamente 3 dígitos tras él)
      - 1.234,56   → 1234.56 (punto = miles, coma = decimal)
      - 1,234.56   → 1234.56 (coma = miles, punto = decimal)
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(' ', '').replace('€', '').replace('$', '')
    if not s:
        return 0.0
    has_comma = ',' in s
    has_dot   = '.' in s
    if has_comma and has_dot:
        # Both present: rightmost separator is the decimal
        if s.rfind(',') > s.rfind('.'):
            # European: 1.234,56
            return float(s.replace('.', '').replace(',', '.'))
        else:
            # US: 1,234.56
            return float(s.replace(',', ''))
    elif has_comma:
        # Only comma: decimal if ≤2 digits follow it, thousands otherwise
        after_comma = s.split(',')[-1]
        if len(after_comma) <= 2:
            return float(s.replace(',', '.'))   # 28,92 → 28.92
        else:
            return float(s.replace(',', ''))    # 1,234 → 1234
    elif has_dot:
        # Only dot: thousands separator if ALL groups after dots are exactly 3 digits
        # e.g. "1.256" → 1256, "1.256.789" → 1256789, but "28.92" → 28.92
        parts = s.split('.')
        if all(len(p) == 3 for p in parts[1:]):
            return float(s.replace('.', ''))    # 1.256 → 1256
        return float(s)                         # 28.92 → 28.92
    else:
        return float(s)


def normalize_text(text):
    """
    Normalize text:
    1. Lowercase
    2. Strip whitespace
    3. Remove accents (NFD normalization)
    """
    if not isinstance(text, str): return ""
    text = text.lower().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')

# ── Validación de IBAN ────────────────────────────────────────────────────────
# Longitud oficial del IBAN por país (registro ISO 13616).
IBAN_COUNTRY_LENGTHS = {
    "AD": 24, "AE": 23, "AL": 28, "AT": 20, "AZ": 28, "BA": 20, "BE": 16, "BG": 22,
    "BH": 22, "BR": 29, "BY": 28, "CH": 21, "CR": 22, "CY": 28, "CZ": 24, "DE": 22,
    "DK": 18, "DO": 28, "EE": 20, "EG": 29, "ES": 24, "FI": 18, "FO": 18, "FR": 27,
    "GB": 22, "GE": 22, "GI": 23, "GL": 18, "GR": 27, "GT": 28, "HR": 21, "HU": 28,
    "IE": 22, "IL": 23, "IQ": 23, "IS": 26, "IT": 27, "JO": 30, "KW": 30, "KZ": 20,
    "LB": 28, "LC": 32, "LI": 21, "LT": 20, "LU": 20, "LV": 21, "LY": 25, "MC": 27,
    "MD": 24, "ME": 22, "MK": 19, "MR": 27, "MT": 31, "MU": 30, "NL": 18, "NO": 15,
    "PK": 24, "PL": 28, "PS": 29, "PT": 25, "QA": 29, "RO": 24, "RS": 22, "SA": 24,
    "SC": 31, "SE": 24, "SI": 19, "SK": 24, "SM": 27, "ST": 25, "SV": 28, "TL": 23,
    "TN": 24, "TR": 26, "UA": 29, "VA": 22, "VG": 24, "XK": 20,
}

# Marcadores que la aplicación escribe en la columna IBAN cuando no hay uno real
IBAN_PLACEHOLDERS = ("", "NO ENCONTRADO", "AMBIGUO")

_RE_IBAN_FORMAT = re.compile(r"^[A-Z]{2}\d{2}[A-Z0-9]{10,30}$")


def normalize_iban(value):
    """Deja el IBAN en mayúsculas y sin espacios ni separadores."""
    if value is None:
        return ""
    return re.sub(r"[\s.\-]", "", str(value)).upper()


def iban_is_valid(value):
    """Valida un IBAN: formato, longitud del país y dígito de control (mod 97)."""
    iban = normalize_iban(value)
    if not iban or iban in IBAN_PLACEHOLDERS or not _RE_IBAN_FORMAT.match(iban):
        return False
    expected_length = IBAN_COUNTRY_LENGTHS.get(iban[:2])
    if expected_length is not None and len(iban) != expected_length:
        return False
    # Se mueven los 4 primeros caracteres al final y cada letra pasa a número (A=10…Z=35)
    rearranged = iban[4:] + iban[:4]
    try:
        digits = "".join(str(int(char, 36)) for char in rearranged)
    except ValueError:
        return False
    return int(digits) % 97 == 1


def iban_error(value):
    """Devuelve el motivo por el que un IBAN no es válido, o None si lo es."""
    iban = normalize_iban(value)
    if not iban or iban in IBAN_PLACEHOLDERS:
        return "falta el IBAN"
    if not _RE_IBAN_FORMAT.match(iban):
        return "el formato no es un IBAN (debe ser 2 letras + 2 dígitos + cuenta)"
    expected_length = IBAN_COUNTRY_LENGTHS.get(iban[:2])
    if expected_length is None:
        return f"el código de país '{iban[:2]}' no existe"
    if len(iban) != expected_length:
        return f"un IBAN de {iban[:2]} tiene {expected_length} caracteres y este tiene {len(iban)}"
    if not iban_is_valid(iban):
        return "el dígito de control no cuadra (suele ser una errata)"
    return None


# ── Juego de caracteres admitido por SEPA ─────────────────────────────────────
# Los bancos rechazan caracteres fuera del subconjunto latino básico de SEPA.
SEPA_ALLOWED_CHARS = set(
    "abcdefghijklmnopqrstuvwxyz"
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "0123456789"
    "/-?:().,'+ "
)
SEPA_CHAR_REPLACEMENTS = {
    "&": "y", "€": "EUR", "$": "USD", "%": "por ciento", "@": "(at)",
    "_": "-", "\\": "/", "|": "/", "*": "-", "#": "-", "=": "-",
    "\"": "'", "«": "'", "»": "'", "“": "'", "”": "'", "‘": "'", "’": "'",
    "–": "-", "—": "-", "º": "o", "ª": "a", "ß": "ss", "æ": "ae", "Æ": "AE",
    "ø": "o", "Ø": "O", "å": "a", "Å": "A", "œ": "oe", "Œ": "OE",
}


def sepa_text(value, max_length, fallback=""):
    """Adapta un texto al juego de caracteres SEPA (sin acentos ni símbolos raros)."""
    text = "" if value is None else str(value)
    # Se separan los acentos y se descartan: á→a, ñ→n, ç→c…
    text = "".join(
        char for char in unicodedata.normalize('NFD', text)
        if unicodedata.category(char) != 'Mn'
    )
    converted = []
    for char in text:
        if char in SEPA_ALLOWED_CHARS:
            converted.append(char)
        elif char in SEPA_CHAR_REPLACEMENTS:
            converted.append(SEPA_CHAR_REPLACEMENTS[char])
        else:
            converted.append(" ")
    cleaned = re.sub(r"\s+", " ", "".join(converted)).strip()
    if not cleaned:
        cleaned = fallback
    return cleaned[:max_length].strip()


# ── Estado de cada línea de la remesa ─────────────────────────────────────────
# (estado, color de la fila, es un problema, motivo)
def payment_issue(record):
    """Clasifica una línea: OK, o el problema que impide pagarla."""
    name = str(record.get('NOMBRE', ''))
    iban = str(record.get('IBAN', ''))
    try:
        amount = float(record.get('IMPORTE', 0) or 0)
    except (TypeError, ValueError):
        amount = 0.0

    if "ERROR" in name or normalize_iban(iban) == "NOENCONTRADO" or not iban.strip():
        return ("ERROR", "error", True, "no se encontró el proveedor en la base de datos")
    if iban.strip().upper() == "AMBIGUO":
        return ("AMBIGUO", "warn", True, "hay varios proveedores posibles: elige uno")

    reason = iban_error(iban)
    if reason:
        return ("IBAN NO VÁLIDO", "error", True, f"IBAN incorrecto: {reason}")
    if amount <= 0:
        return ("SIN IMPORTE", "warn", True, "el importe es 0 € o negativo")
    return ("OK", "ok", False, None)


def is_payable(record):
    """True si la línea puede incluirse en el fichero SEPA."""
    return not payment_issue(record)[2]


def find_duplicate_payments(results):
    """Agrupa pagos que se repiten (mismo IBAN e importe) para avisar de dobles pagos."""
    groups = {}
    for record in results:
        if not is_payable(record):
            continue
        key = (normalize_iban(record.get('IBAN')), round(float(record.get('IMPORTE', 0)), 2))
        groups.setdefault(key, []).append(record)
    return [(key, rows) for key, rows in groups.items() if len(rows) > 1]


# ── Datos del ordenante ───────────────────────────────────────────────────────
SEPA_REQUIRED_FIELDS = {
    "sepa_nombre": "Nombre de la empresa",
    "sepa_cif": "CIF/NIF",
    "sepa_iban": "IBAN de la empresa",
}
_RE_BIC = re.compile(r"^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$")


def sepa_config_errors(config):
    """Comprueba los datos del ordenante. Devuelve la lista de problemas encontrados."""
    cfg = {**SEPA_DEFAULTS, **{k: v for k, v in (config or {}).items() if k.startswith("sepa_")}}
    errors = []
    for key, label in SEPA_REQUIRED_FIELDS.items():
        if not str(cfg.get(key, "")).strip():
            errors.append(f"Falta «{label}».")
    debtor_iban = cfg.get("sepa_iban", "")
    if str(debtor_iban).strip():
        reason = iban_error(debtor_iban)
        if reason:
            errors.append(f"El IBAN de la empresa no es válido: {reason}.")
    bic = normalize_iban(cfg.get("sepa_bic", ""))
    if bic and not _RE_BIC.match(bic):
        errors.append("El BIC/SWIFT debe tener 8 u 11 caracteres (p. ej. BSCHESMMXXX).")
    pais = str(cfg.get("sepa_pais", "")).strip().upper()
    if pais and len(pais) != 2:
        errors.append("El país debe ser el código ISO de 2 letras (p. ej. ES).")
    return errors


def clean_db_value(value):
    """Convierte un valor de la BD en texto limpio ('' si está vacío o es NaN)."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in ("nan", "none", "nat"):
        return ""
    return text


def find_db_row(db_df, name):
    """Busca una fila de la BD por nombre exacto (sin distinguir mayúsculas ni acentos)."""
    if db_df is None or 'NOMBRE' not in getattr(db_df, 'columns', []):
        return None
    target = normalize_text(name)
    if not target:
        return None
    for _, row in db_df.iterrows():
        if normalize_text(str(row.get('NOMBRE', ''))) == target:
            return row
    return None


class AutocompleteEntry(tk.Entry):
    """Entry con lista desplegable de proveedores.

    Al escribir se muestran los nombres de la base de datos: primero los que
    empiezan por el texto escrito y después los que lo contienen, ignorando
    mayúsculas y acentos. Cuantas más letras se escriben, menos candidatos
    quedan. Flechas ↑/↓ para navegar, Enter para elegir y Esc para cerrar.
    """

    MAX_VISIBLE_ROWS = 10
    _IGNORED_KEYS = {
        "Up", "Down", "Left", "Right", "Return", "KP_Enter", "Escape", "Tab",
        "ISO_Left_Tab", "Shift_L", "Shift_R", "Control_L", "Control_R",
        "Alt_L", "Alt_R", "Caps_Lock", "Home", "End", "Prior", "Next",
    }

    def __init__(self, master, values=None, on_select=None, on_commit=None,
                 on_cancel=None, on_tab=None, on_focus_out=None,
                 max_results=300, **kwargs):
        super().__init__(master, **kwargs)
        self.on_select = on_select        # se eligió un proveedor de la lista
        self.on_commit = on_commit        # Enter sin lista desplegada
        self.on_cancel = on_cancel        # Esc sin lista desplegada
        self.on_tab = on_tab              # Tab (recibe shift: True/False)
        self.on_focus_out = on_focus_out  # el foco salió del campo
        self.max_results = max_results
        self._popup = None
        self._listbox = None
        self.set_values(values or [])

        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<Down>", self._on_down)
        self.bind("<Up>", self._on_up)
        self.bind("<Return>", self._on_return)
        self.bind("<KP_Enter>", self._on_return)
        self.bind("<Tab>", lambda e: self._on_tab(e, shift=False))
        self.bind("<Shift-Tab>", lambda e: self._on_tab(e, shift=True))
        self.bind("<ISO_Left_Tab>", lambda e: self._on_tab(e, shift=True))
        self.bind("<Escape>", self._on_escape)
        self.bind("<FocusOut>", self._on_focus_out_event)
        self.bind("<Destroy>", lambda e: self._destroy_popup())

    # ── Datos ────────────────────────────────────────────────────────────
    def set_values(self, values):
        seen = set()
        clean = []
        for value in values or []:
            text = str(value).strip()
            key = normalize_text(text)
            if text and key not in seen:
                seen.add(key)
                clean.append(text)
        clean.sort(key=normalize_text)
        self._values = clean
        self._normalized = [(normalize_text(v), v) for v in clean]

    def matches(self, text):
        """Coincidencias ordenadas: por prefijo, por inicio de palabra y por contenido."""
        target = normalize_text(text)
        if not target:
            return self._values[:self.max_results]
        starts, word_starts, contains = [], [], []
        for norm, original in self._normalized:
            if norm.startswith(target):
                starts.append(original)
            elif any(word.startswith(target) for word in norm.split()):
                word_starts.append(original)
            elif target in norm:
                contains.append(original)
        if len(target) < 2:
            # Con una sola letra se listan solo los que empiezan por ella
            # (nombre o apellido); si no, saldría casi toda la base de datos.
            return (starts + word_starts)[:self.max_results]
        return (starts + word_starts + contains)[:self.max_results]

    # ── Lista desplegable ────────────────────────────────────────────────
    def _popup_alive(self):
        return self._popup is not None and self._popup.winfo_exists()

    def _list_visible(self):
        return self._popup_alive() and self._popup.winfo_viewable()

    def _ensure_popup(self):
        if self._popup_alive():
            return
        self._popup = tk.Toplevel(self)
        self._popup.wm_overrideredirect(True)
        try:
            self._popup.attributes("-topmost", True)
        except tk.TclError:
            pass
        border = tk.Frame(self._popup, bd=1, relief=tk.SOLID, bg="#9e9e9e")
        border.pack(fill=tk.BOTH, expand=True)
        self._listbox = tk.Listbox(
            border, activestyle="none", exportselection=False,
            highlightthickness=0, bd=0,
            selectbackground="#2c7be5", selectforeground="white",
        )
        scrollbar = ttk.Scrollbar(border, orient="vertical", command=self._listbox.yview)
        self._listbox.configure(yscrollcommand=scrollbar.set)
        self._listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._listbox.bind("<ButtonRelease-1>", lambda e: self._accept_selection())
        self._listbox.bind("<Motion>", self._on_list_motion)

    def show_list(self, matches=None):
        if matches is None:
            matches = self.matches(self.get())
        if not matches:
            self.hide_list()
            return
        self._ensure_popup()
        self._listbox.delete(0, tk.END)
        for name in matches:
            self._listbox.insert(tk.END, name)
        self._listbox.selection_clear(0, tk.END)
        self._listbox.selection_set(0)
        self._listbox.activate(0)
        self._listbox.see(0)
        self._listbox.configure(height=min(len(matches), self.MAX_VISIBLE_ROWS))

        self._popup.update_idletasks()
        width = max(self.winfo_width(), 280)
        height = self._popup.winfo_reqheight()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        if y + height > self.winfo_screenheight() and self.winfo_rooty() - height > 0:
            y = self.winfo_rooty() - height
        self._popup.wm_geometry(f"{width}x{height}+{x}+{y}")
        self._popup.deiconify()
        self._popup.lift()

    def hide_list(self):
        if self._popup_alive():
            self._popup.withdraw()

    def _destroy_popup(self):
        if self._popup_alive():
            try:
                self._popup.destroy()
            except tk.TclError:
                pass
        self._popup = None
        self._listbox = None

    def _move_selection(self, delta):
        if not self._list_visible():
            self.show_list()
            return
        size = self._listbox.size()
        if not size:
            return
        current = self._listbox.curselection()
        index = (current[0] if current else 0) + delta
        index = max(0, min(size - 1, index))
        self._listbox.selection_clear(0, tk.END)
        self._listbox.selection_set(index)
        self._listbox.activate(index)
        self._listbox.see(index)

    def _accept_selection(self):
        if not self._list_visible():
            return "break"
        selection = self._listbox.curselection()
        if not selection:
            return "break"
        value = self._listbox.get(selection[0])
        self.delete(0, tk.END)
        self.insert(0, value)
        self.icursor(tk.END)
        self.hide_list()
        if self.on_select:
            self.on_select(value)
        return "break"

    # ── Eventos ──────────────────────────────────────────────────────────
    def _on_key_release(self, event):
        if event.keysym in self._IGNORED_KEYS:
            return
        if event.state & 0x4 and len(event.keysym) == 1:   # Ctrl+letra
            return
        self.show_list()

    def _on_list_motion(self, event):
        index = self._listbox.nearest(event.y)
        if index >= 0:
            self._listbox.selection_clear(0, tk.END)
            self._listbox.selection_set(index)
            self._listbox.activate(index)

    def _on_down(self, event):
        self._move_selection(1)
        return "break"

    def _on_up(self, event):
        self._move_selection(-1)
        return "break"

    def _on_return(self, event):
        if self._list_visible():
            return self._accept_selection()
        if self.on_commit:
            self.on_commit()
            return "break"
        return None

    def _on_tab(self, event, shift=False):
        if self._list_visible():
            self._accept_selection()
            return "break"
        if self.on_tab:
            self.on_tab(shift)
            return "break"
        return None

    def _on_escape(self, event):
        if self._list_visible():
            self.hide_list()
            return "break"
        if self.on_cancel:
            self.on_cancel()
            return "break"
        return None

    def _on_focus_out_event(self, event):
        # Se retrasa para que un clic sobre la lista llegue a procesarse.
        self.after(150, self._close_after_focus_out)

    def _close_after_focus_out(self):
        if self._pointer_over_popup():
            return
        self.hide_list()
        if self.on_focus_out:
            self.on_focus_out()

    def _pointer_over_popup(self):
        if not self._list_visible():
            return False
        try:
            x, y = self._popup.winfo_pointerxy()
            left, top = self._popup.winfo_rootx(), self._popup.winfo_rooty()
            return (left <= x <= left + self._popup.winfo_width()
                    and top <= y <= top + self._popup.winfo_height())
        except tk.TclError:
            return False


class AmbiguityResolverDialog(tk.Toplevel):
    def __init__(self, parent, candidates_with_ibans, callback, manual_edit_callback=None):
        super().__init__(parent)
        self.title("Resolver Ambigüedad")
        self.geometry("600x450")
        self.callback = callback
        self.manual_edit_callback = manual_edit_callback
        self.selected_name = None
        self.selected_iban = None
        
        # Header
        tk.Label(self, text="Se encontraron múltiples coincidencias.", 
                 font=("Helvetica", 12, "bold")).pack(pady=10)
        tk.Label(self, text="Selecciona el registro correcto:", 
                 font=("Helvetica", 10)).pack(pady=5)
        
        # Listbox with candidates
        list_frame = tk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, 
                                   font=("Consolas", 10), height=10)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # Populate listbox
        self.candidates = candidates_with_ibans
        for name, iban in candidates_with_ibans:
            display = f"{name:<40} → {iban}"
            self.listbox.insert(tk.END, display)
        
        # Add "None of these" option
        self.listbox.insert(tk.END, "")  # Separator
        self.listbox.insert(tk.END, "❌ Ninguna de estas (Editar manualmente)")
        
        # Buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="✓ Seleccionar", command=self.select, 
                  bg="#c8e6c9", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Cancelar", command=self.destroy, 
                  font=("Helvetica", 10)).pack(side=tk.LEFT, padx=5)
        
        # Double-click to select
        self.listbox.bind("<Double-Button-1>", lambda e: self.select())
        
        # Select first by default
        if candidates_with_ibans:
            self.listbox.selection_set(0)
    
    def select(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecciona un registro.")
            return
        
        idx = selection[0]
        
        # Check if "None of these" was selected
        if idx >= len(self.candidates):
            # User wants to edit manually
            self.destroy()
            if self.manual_edit_callback:
                self.manual_edit_callback()
            return
        
        self.selected_name, self.selected_iban = self.candidates[idx]
        
        if self.callback:
            self.callback(self.selected_name, self.selected_iban)
        
        self.destroy()

class SepaConfigDialog(tk.Toplevel):
    """Dialog to configure SEPA debtor (ordenante) details."""
    def __init__(self, parent, config, save_callback):
        super().__init__(parent)
        self.title("Configuración SEPA - Datos del Ordenante")
        self.geometry("550x400")
        self.save_callback = save_callback

        tk.Label(self, text="Datos del Ordenante (Empresa)", font=("Helvetica", 12, "bold")).pack(pady=10)

        form = tk.Frame(self)
        form.pack(fill=tk.X, padx=20, pady=5)

        fields = [
            ("Nombre empresa:", "sepa_nombre"),
            ("CIF/NIF:", "sepa_cif"),
            ("IBAN:", "sepa_iban"),
            ("BIC/SWIFT:", "sepa_bic"),
            ("Dirección:", "sepa_direccion"),
            ("Código Postal:", "sepa_cp"),
            ("Ciudad:", "sepa_ciudad"),
            ("Provincia:", "sepa_provincia"),
            ("País (ISO):", "sepa_pais"),
        ]

        self.vars = {}
        for i, (label, key) in enumerate(fields):
            tk.Label(form, text=label, anchor="w").grid(row=i, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=config.get(key, SEPA_DEFAULTS.get(key, "")))
            width = 50 if key in ("sepa_nombre", "sepa_direccion", "sepa_iban") else 30
            tk.Entry(form, textvariable=var, width=width).grid(row=i, column=1, pady=3, padx=5)
            self.vars[key] = var

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="💾 Guardar", command=self.save, bg="#c8e6c9",
                  font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side=tk.LEFT, padx=5)

    def save(self):
        result = {key: var.get() for key, var in self.vars.items()}
        if self.save_callback:
            self.save_callback(result)
        self.destroy()


class SepaConfigError(ValueError):
    """Los datos del ordenante están incompletos o son incorrectos."""


def generate_sepa_xml(results, config, output_path=None, exec_date=None):
    """Generate SEPA Credit Transfer XML (pain.001.001.03) from remesa results."""
    # Sin unos datos del ordenante correctos el banco rechaza el fichero entero
    config_errors = sepa_config_errors(config)
    if config_errors:
        raise SepaConfigError("\n".join(config_errors))

    # Solo se pagan las líneas con IBAN válido (dígito de control) e importe positivo
    valid = [r for r in results if is_payable(r)]

    if not valid:
        return None

    now = datetime.now()
    msg_id = now.strftime("%Y%m%d%H%M%S")
    nb_txs = str(len(valid))
    ctrl_sum = f"{sum(r['IMPORTE'] for r in valid):.2f}"

    # Get config values with defaults
    cfg = {**SEPA_DEFAULTS, **{k: v for k, v in config.items() if k.startswith("sepa_")}}

    ns = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
    doc = Element("Document", xmlns=ns)
    root = SubElement(doc, "CstmrCdtTrfInitn")

    # --- Group Header ---
    grp = SubElement(root, "GrpHdr")
    SubElement(grp, "MsgId").text = msg_id
    SubElement(grp, "CreDtTm").text = now.strftime("%Y-%m-%dT%H:%M:%S")
    SubElement(grp, "NbOfTxs").text = nb_txs
    SubElement(grp, "CtrlSum").text = ctrl_sum
    initg = SubElement(grp, "InitgPty")
    SubElement(initg, "Nm").text = sepa_text(cfg["sepa_nombre"], 70)
    org_id = SubElement(SubElement(SubElement(initg, "Id"), "OrgId"), "Othr")
    SubElement(org_id, "Id").text = cfg["sepa_cif"]

    # --- Payment Information ---
    pmt = SubElement(root, "PmtInf")
    SubElement(pmt, "PmtInfId").text = f"{msg_id}-1"
    SubElement(pmt, "PmtMtd").text = "TRF"
    SubElement(pmt, "BtchBookg").text = "false"
    SubElement(pmt, "NbOfTxs").text = nb_txs
    SubElement(pmt, "CtrlSum").text = ctrl_sum

    svc = SubElement(SubElement(pmt, "PmtTpInf"), "SvcLvl")
    SubElement(svc, "Cd").text = "SEPA"

    SubElement(pmt, "ReqdExctnDt").text = exec_date or now.strftime("%Y-%m-%d")

    # Debtor
    dbtr = SubElement(pmt, "Dbtr")
    SubElement(dbtr, "Nm").text = sepa_text(cfg["sepa_nombre"], 70)
    addr = SubElement(dbtr, "PstlAdr")
    SubElement(addr, "PstCd").text = sepa_text(cfg["sepa_cp"], 16)
    SubElement(addr, "TwnNm").text = sepa_text(cfg["sepa_ciudad"], 35)
    SubElement(addr, "CtrySubDvsn").text = sepa_text(cfg["sepa_provincia"], 35)
    SubElement(addr, "Ctry").text = cfg["sepa_pais"].strip().upper()
    SubElement(addr, "AdrLine").text = sepa_text(cfg["sepa_direccion"], 70)
    dbtr_org = SubElement(SubElement(SubElement(dbtr, "Id"), "OrgId"), "Othr")
    SubElement(dbtr_org, "Id").text = cfg["sepa_cif"]

    # Debtor Account
    dbtr_acct = SubElement(pmt, "DbtrAcct")
    SubElement(SubElement(dbtr_acct, "Id"), "IBAN").text = normalize_iban(cfg["sepa_iban"])
    SubElement(dbtr_acct, "Ccy").text = "EUR"

    # Debtor Agent (Bank)
    dbtr_agt = SubElement(pmt, "DbtrAgt")
    SubElement(SubElement(dbtr_agt, "FinInstnId"), "BIC").text = normalize_iban(cfg["sepa_bic"])

    SubElement(pmt, "ChrgBr").text = "SLEV"

    # --- Credit Transfer Transactions ---
    for i, r in enumerate(valid, 1):
        tx = SubElement(pmt, "CdtTrfTxInf")

        pmt_id = SubElement(tx, "PmtId")
        end2end = f"{msg_id}{i:02d}"
        SubElement(pmt_id, "InstrId").text = end2end
        SubElement(pmt_id, "EndToEndId").text = end2end

        amt = SubElement(tx, "Amt")
        instd = SubElement(amt, "InstdAmt", Ccy="EUR")
        instd.text = f"{r['IMPORTE']:.2f}"

        cdtr = SubElement(tx, "Cdtr")
        # Clean name: remove prefixes like "REVISAR: AMBIGUO: ..."
        clean_name = r['NOMBRE']
        for prefix in ("REVISAR: ", "AMBIGUO: "):
            if clean_name.startswith(prefix):
                clean_name = clean_name[len(prefix):]
        # Máximo 70 caracteres y solo el juego de caracteres admitido por SEPA
        SubElement(cdtr, "Nm").text = sepa_text(clean_name, 70, fallback="BENEFICIARIO")

        cdtr_addr = SubElement(cdtr, "PstlAdr")
        # Derive country from IBAN prefix (first 2 chars)
        iban = normalize_iban(r['IBAN'])
        country = iban[:2] if len(iban) >= 2 else cfg["sepa_pais"]
        SubElement(cdtr_addr, "Ctry").text = country

        cdtr_acct = SubElement(tx, "CdtrAcct")
        SubElement(SubElement(cdtr_acct, "Id"), "IBAN").text = iban

        rmt = SubElement(tx, "RmtInf")
        concept = clean_db_value(r.get('CONCEPTO_NORMA')) or DEFAULT_CONCEPT
        # Máximo 140 caracteres, nunca vacío y sin caracteres que el banco rechace
        SubElement(rmt, "Ustrd").text = sepa_text(concept, 140, fallback=DEFAULT_CONCEPT)

    # Write XML
    if output_path is None:
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        output_path = f"REMESA_SEPA_{timestamp}.xml"

    tree = ElementTree(doc)
    indent(tree, space="  ")
    tree.write(output_path, encoding="UTF-8", xml_declaration=True)

    # Add standalone="no" attribute (standard SEPA requirement)
    with open(output_path, 'r', encoding='utf-8') as f:
        content = f.read()
    content = content.replace("<?xml version='1.0' encoding='UTF-8'?>",
                              '<?xml version="1.0" encoding="UTF-8" standalone="no"?>')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

    return output_path


class EditDialog(tk.Toplevel):
    def __init__(self, parent, result_data, db_df, save_callback):
        super().__init__(parent)
        self.title("Editar Detalle")
        self.geometry("500x400")
        self.result_data = result_data
        self.db_df = db_df
        self.save_callback = save_callback
        
        # Styles
        PADDING = 10
        
        # Current Info
        tk.Label(self, text=f"Archivo: {result_data['FILENAME']}", font=("bold", 10)).pack(pady=PADDING)
        
        # Form
        input_frame = tk.Frame(self)
        input_frame.pack(fill=tk.X, padx=PADDING)
        
        tk.Label(input_frame, text="Nombre:").grid(row=0, column=0, sticky="w")
        self.name_var = tk.StringVar(value=result_data['NOMBRE'])
        self.name_entry = AutocompleteEntry(
            input_frame, values=self._db_names(), textvariable=self.name_var,
            on_select=self._on_provider_selected, width=40,
        )
        self.name_entry.grid(row=0, column=1, pady=5)
        tk.Label(input_frame, text="(escribe y elige de la lista)",
                 fg="gray").grid(row=0, column=2, sticky="w", padx=5)
        
        tk.Label(input_frame, text="IBAN:").grid(row=1, column=0, sticky="w")
        self.iban_var = tk.StringVar(value=result_data['IBAN'])
        tk.Entry(input_frame, textvariable=self.iban_var, width=40).grid(row=1, column=1, pady=5)
        
        tk.Label(input_frame, text="Importe:").grid(row=2, column=0, sticky="w")
        self.amount_var = tk.StringVar(value=str(result_data['IMPORTE']))
        tk.Entry(input_frame, textvariable=self.amount_var, width=20).grid(row=2, column=1, pady=5, sticky="w")

        tk.Label(input_frame, text="Concepto:").grid(row=3, column=0, sticky="w")
        self.concepto_var = tk.StringVar(value=result_data.get('CONCEPTO_NORMA', ''))
        tk.Entry(input_frame, textvariable=self.concepto_var, width=40).grid(row=3, column=1, pady=5)

        # Actions
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=20, padx=PADDING)
        
        # 1. Open PDF
        tk.Button(btn_frame, text="📄 Abrir PDF Original", command=self.open_pdf, bg="#e1f5fe").pack(fill=tk.X, pady=5)
        
        # 2. Add to DB Checkbox
        self.add_db_var = tk.BooleanVar(value=False)
        self.chk_db = tk.Checkbutton(btn_frame, text="Añadir/Actualizar este Nombre e IBAN a la Base de Datos", variable=self.add_db_var)
        self.chk_db.pack(fill=tk.X, pady=5)
        
        # Save Buttons
        tk.Button(btn_frame, text="💾 Guardar Cambios", command=self.save, bg="#c8e6c9").pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)

    def _db_names(self):
        """Nombres de proveedor disponibles en la base de datos."""
        if self.db_df is None or 'NOMBRE' not in getattr(self.db_df, 'columns', []):
            return []
        return [str(n).strip() for n in self.db_df['NOMBRE'].dropna().tolist() if str(n).strip()]

    def _on_provider_selected(self, name):
        """Al elegir un proveedor se rellenan IBAN y concepto desde la BD."""
        row = find_db_row(self.db_df, name)
        if row is None:
            return
        iban = clean_db_value(row.get('IBAN'))
        if iban:
            self.iban_var.set(iban)
        concepto = clean_db_value(row.get('CONCEPTO_NORMA'))
        if concepto:
            self.concepto_var.set(concepto)

    def open_pdf(self):
        try:
            filepath = self.result_data.get('FULLPATH')
            if not filepath or not os.path.exists(filepath):
                messagebox.showerror("Error", "No se encuentra el archivo PDF.")
                return
            system = platform.system()
            if system == "Windows":
                os.startfile(filepath)
            elif system == "Darwin":
                subprocess.Popen(["open", filepath])
            else:
                subprocess.Popen(["xdg-open", filepath])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir PDF: {e}")

    def save(self):
        # Update result data
        self.result_data['NOMBRE'] = self.name_var.get()
        self.result_data['IBAN'] = self.iban_var.get()
        try:
            self.result_data['IMPORTE'] = float(self.amount_var.get().replace(',','.'))
        except (ValueError, TypeError):
            pass
        self.result_data['CONCEPTO_NORMA'] = self.concepto_var.get()
        
        # Callback to update Treeview
        if self.save_callback:
            self.save_callback(self.result_data, self.add_db_var.get())
        
        self.destroy()


# ── Auto-update ───────────────────────────────────────────────────────────────────────────
GITHUB_REPO = "aruizciee/RemesaPro"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def _get_ssl_context():
    """Get SSL context — handles macOS PyInstaller certificate issues."""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        pass
    # Try default context first
    ctx = ssl.create_default_context()
    try:
        urllib_request.urlopen(
            urllib_request.Request("https://api.github.com", headers={"User-Agent": "test"}),
            timeout=5, context=ctx
        )
        return ctx
    except ssl.SSLError:
        # Fallback: unverified context (safe for read-only public API)
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return ctx


def check_for_updates():
    """Check GitHub for a newer release. Returns (new_version, download_url, asset_name) or error string."""
    try:
        ctx = _get_ssl_context()
        req = urllib_request.Request(GITHUB_API_LATEST, headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "RemesaPro-Updater"
        })
        with urllib_request.urlopen(req, timeout=10, context=ctx) as resp:
            data = json.loads(resp.read().decode())
        tag = data.get("tag_name", "")  # e.g. "build-8"
        remote_version = int(tag.replace("build-", "")) if tag.startswith("build-") else 0
        print(f"[Updater] Local: v{APP_VERSION}, Remote: v{remote_version} (tag: {tag})")
        if remote_version <= APP_VERSION:
            return None
        # Pick the right asset for this OS
        is_mac = platform.system() == "Darwin"
        suffix = "macOS.zip" if is_mac else ".exe"
        for asset in data.get("assets", []):
            if asset["name"].endswith(suffix):
                return (remote_version, asset["browser_download_url"], asset["name"])
        return f"ERROR: No se encontró asset para {'macOS' if is_mac else 'Windows'}"
    except Exception as e:
        return f"ERROR: {e}"


def download_and_apply_update(download_url, asset_name, status_callback=None):
    """Download the new version and replace the current executable."""
    try:
        if status_callback:
            status_callback("Descargando actualización...")

        # Download to temp location
        import tempfile
        tmp_dir = tempfile.mkdtemp()
        tmp_file = os.path.join(tmp_dir, asset_name)
        ctx = _get_ssl_context()
        req = urllib_request.Request(download_url, headers={"User-Agent": "RemesaPro-Updater"})
        with urllib_request.urlopen(req, timeout=60, context=ctx) as resp:
            with open(tmp_file, 'wb') as f:
                f.write(resp.read())

        current_exe = sys.executable  # Path of the running .exe / binary
        is_mac = platform.system() == "Darwin"

        if is_mac:
            # macOS: unzip and replace the .app or binary
            import zipfile
            with zipfile.ZipFile(tmp_file, 'r') as zf:
                zf.extractall(tmp_dir)
            # Find the extracted binary
            extracted = os.path.join(tmp_dir, "RemesaPro")
            if not os.path.exists(extracted):
                # Look for it inside .app bundle
                app_binary = os.path.join(tmp_dir, "RemesaPro.app", "Contents", "MacOS", "RemesaPro")
                if os.path.exists(app_binary):
                    extracted = app_binary
            if os.path.exists(extracted):
                os.chmod(extracted, 0o755)
                backup = current_exe + ".old"
                if os.path.exists(backup):
                    os.remove(backup)
                os.rename(current_exe, backup)
                import shutil
                shutil.copy2(extracted, current_exe)
                os.chmod(current_exe, 0o755)
        else:
            # Windows: rename current exe, move new one in place
            backup = current_exe + ".old"
            if os.path.exists(backup):
                os.remove(backup)
            os.rename(current_exe, backup)
            import shutil
            shutil.copy2(tmp_file, current_exe)

        # Clean up temp
        import shutil
        shutil.rmtree(tmp_dir, ignore_errors=True)

        if status_callback:
            status_callback("Actualización completada")
        return True
    except Exception as e:
        if status_callback:
            status_callback(f"Error al actualizar: {e}")
        return False


class SepaPreviewDialog(tk.Toplevel):
    """Read-only preview of the SEPA XML before saving."""
    def __init__(self, parent, xml_content):
        super().__init__(parent)
        self.title("Vista previa SEPA XML")
        self.geometry("800x600")

        tk.Label(self, text="Vista previa del XML SEPA (solo lectura)",
                 font=("Helvetica", 11, "bold")).pack(pady=8)

        frame = tk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        vsb = tk.Scrollbar(frame, orient="vertical")
        hsb = tk.Scrollbar(frame, orient="horizontal")
        text_widget = tk.Text(frame, wrap="none", font=("Consolas", 9),
                              yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                              state="normal")
        vsb.config(command=text_widget.yview)
        hsb.config(command=text_widget.xview)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        text_widget.insert("1.0", xml_content)
        text_widget.config(state="disabled")

        tk.Button(self, text="Cerrar", command=self.destroy,
                  font=("Helvetica", 10)).pack(pady=8)


class RemesaApp:
    # Columnas que se pueden editar directamente en la tabla
    EDITABLE_COLUMNS = ("nombre_db", "iban", "importe", "concepto")

    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Remesas - CIEE Pro")
        self.root.geometry("1100x750")

        # App icon
        self.base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        try:
            if platform.system() == "Windows":
                ico_path = os.path.join(self.base_path, ICON_ICO)
                if not os.path.exists(ico_path):
                    ico_path = ICON_ICO
                if os.path.exists(ico_path):
                    self.root.iconbitmap(ico_path)
            else:
                png_path = os.path.join(self.base_path, ICON_PNG)
                if not os.path.exists(png_path):
                    png_path = ICON_PNG
                if os.path.exists(png_path):
                    img = tk.PhotoImage(file=png_path)
                    self.root.iconphoto(True, img)
        except Exception:
            pass


        # Load Config
        self.config = self.load_config()

        # Styles
        style = ttk.Style()
        style.configure("TButton", font=("Helvetica", 10))
        style.configure("Header.TLabel", font=("Helvetica", 14, "bold"), foreground="#2c3e50")
        
        # Header Frame
        header_frame = ttk.Frame(root, padding="10")
        header_frame.pack(fill=tk.X)
        
        try:
            logo_path = os.path.join(self.base_path, LOGO_FILE)
            if not os.path.exists(logo_path): logo_path = LOGO_FILE
            
            self.logo_img = tk.PhotoImage(file=logo_path)
            h = self.logo_img.height()
            if h > 80:
                factor = int(h / 80)
                if factor < 1: factor = 1
                self.logo_img = self.logo_img.subsample(factor, factor)
                
            lbl_logo = ttk.Label(header_frame, image=self.logo_img)
            lbl_logo.pack(side=tk.LEFT, padx=10)
        except Exception: pass

        ttk.Label(header_frame, text="RemesaPro - Generador de Remesas", style="Header.TLabel").pack(side=tk.LEFT, padx=10)

        # Version label + update button in header
        self.version_label = ttk.Label(header_frame, text=f"v{APP_VERSION}", font=("Helvetica", 8), foreground="gray")
        self.version_label.pack(side=tk.RIGHT, padx=5)
        self.btn_update = ttk.Button(header_frame, text="🔄 Buscar actualizaciones", command=self.check_updates_manual)
        self.btn_update.pack(side=tk.RIGHT, padx=5)

        # Main Container
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Inputs
        input_frame = ttk.LabelFrame(main_frame, text="Configuración", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(input_frame, text="Carpeta de PDFs:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.folder_var = tk.StringVar(value=self.config.get("last_folder", ""))
        ttk.Entry(input_frame, textvariable=self.folder_var, width=80).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Examinar", command=self.select_folder).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(input_frame, text="Base de Datos:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.db_var = tk.StringVar(value=self.config.get("last_db", ""))
        ttk.Entry(input_frame, textvariable=self.db_var, width=80).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Examinar", command=self.select_db).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(input_frame, text="Fecha ejecución SEPA:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.sepa_date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ttk.Entry(input_frame, textvariable=self.sepa_date_var, width=15).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(input_frame, text="(DD/MM/AAAA)", foreground="gray").grid(row=2, column=2, padx=5, pady=5, sticky="w")


        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.btn_process = ttk.Button(btn_frame, text="🔍 1. Analizar PDFs", command=self.start_processing_thread)
        self.btn_process.pack(side=tk.LEFT, padx=5)
        
        self.btn_save = ttk.Button(btn_frame, text="💾 2. Guardar Excel", command=self.save_results, state="disabled")
        self.btn_save.pack(side=tk.LEFT, padx=5)

        self.btn_sepa = ttk.Button(btn_frame, text="🏦 3. Generar SEPA XML", command=self.generate_sepa, state="disabled")
        self.btn_sepa.pack(side=tk.LEFT, padx=5)

        self.btn_sepa_preview = ttk.Button(btn_frame, text="🔍 Vista previa XML", command=self.preview_sepa, state="disabled")
        self.btn_sepa_preview.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="⚙ SEPA Config", command=self.open_sepa_config).pack(side=tk.LEFT, padx=5)

        self.lbl_status = ttk.Label(btn_frame, text="Listo", font=("Helvetica", 9, "italic"))
        self.lbl_status.pack(side=tk.LEFT, padx=15)
        
        # Filter checkbox
        self.filter_var = tk.BooleanVar(value=False)
        self.chk_filter = ttk.Checkbutton(btn_frame, text="Mostrar solo problemas (Ambiguos + Errores)", 
                                          variable=self.filter_var, command=self.refresh_table)
        self.chk_filter.pack(side=tk.RIGHT, padx=10)
        
        ttk.Label(btn_frame, text="(Doble clic en una celda para editarla)", foreground="gray").pack(side=tk.RIGHT)

        # Progress bar
        self.progress_var = tk.IntVar(value=0)
        self.progressbar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progressbar.pack(fill=tk.X, pady=(0, 5))

        # Búsqueda rápida sobre la tabla
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X)

        ttk.Label(search_frame, text="🔎 Buscar:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=35)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_var.trace_add("write", lambda *_: self.refresh_table(persist=False))
        ttk.Button(search_frame, text="✖", width=3,
                   command=lambda: self.search_var.set("")).pack(side=tk.LEFT)
        self.HINT_TEXT = "Ctrl+C/V: celdas · Ctrl+Mayús+C/V: proveedor · F2: editar"
        self.lbl_hint = ttk.Label(search_frame, text=self.HINT_TEXT, foreground="gray")
        self.lbl_hint.pack(side=tk.LEFT, padx=15)

        # Treeview
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Add hidden index column for proper mapping when filtered
        columns = ("idx", "archivo", "nombre_db", "iban", "importe", "concepto", "estado")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")

        # Hide the index column
        self.tree.column("idx", width=0, stretch=False)
        self.tree.heading("idx", text="")

        self.tree.heading("archivo", text="Archivo PDF", command=lambda: self._sort_table("archivo"))
        self.tree.heading("nombre_db", text="Nombre Detectado", command=lambda: self._sort_table("nombre_db"))
        self.tree.heading("iban", text="IBAN", command=lambda: self._sort_table("iban"))
        self.tree.heading("importe", text="Importe (€)", command=lambda: self._sort_table("importe"))
        self.tree.heading("concepto", text="Concepto", command=lambda: self._sort_table("concepto"))
        self.tree.heading("estado", text="Estado", command=lambda: self._sort_table("estado"))

        self.tree.column("archivo", width=200)
        self.tree.column("nombre_db", width=200)
        self.tree.column("iban", width=200)
        self.tree.column("importe", width=80, anchor="e")
        self.tree.column("concepto", width=180)
        self.tree.column("estado", width=120)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(column=0, row=0, sticky='nsew')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        self.tree.tag_configure("ok", background="#d4edda")
        self.tree.tag_configure("error", background="#f8d7da")
        self.tree.tag_configure("warn", background="#fff3cd")
        
        # Doble clic: edita la celda; en Archivo/Estado abre la ventana de detalle
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Button-1>", self._on_tree_click, add="+")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select, add="+")

        # Menú contextual y teclas
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Delete>", lambda e: self._delete_selected_row())
        self.tree.bind("<F2>", lambda e: (self._edit_focused_cell(), "break")[1])
        self.tree.bind("<Return>", lambda e: (self._edit_focused_cell(), "break")[1])
        self.tree.bind("<Control-c>", self._copy_cells)
        self.tree.bind("<Control-v>", self._paste_cells)
        self.tree.bind("<Control-C>", self._copy_provider)
        self.tree.bind("<Control-V>", self._paste_provider)
        self.tree.bind("<Control-Shift-C>", self._copy_provider)
        self.tree.bind("<Control-Shift-V>", self._paste_provider)

        self._context_menu = tk.Menu(self.root, tearoff=0)
        self._context_menu.add_command(label="✏️  Editar celda (F2)", command=self._edit_focused_cell)
        self._context_menu.add_command(label="🗔  Editar en ventana…", command=self._edit_selected_row)
        self._context_menu.add_separator()
        self._context_menu.add_command(label="📋  Copiar proveedor (Ctrl+Mayús+C)", command=self._copy_provider)
        self._context_menu.add_command(label="📥  Pegar proveedor (Ctrl+Mayús+V)", command=self._paste_provider)
        self._context_menu.add_separator()
        self._context_menu.add_command(label="🗑️  Eliminar de la remesa", command=self._delete_selected_row)

        self.current_results = []
        self.loaded_db_df = None
        self._sort_col = None
        self._sort_reverse = False

        # Estado de edición en línea y portapapeles
        self._item_by_index = {}
        self._focus_column = "nombre_db"
        self._provider_clipboard = None
        self._editor = None
        self._editor_item = None
        self._editor_column = None
        self._editor_index = None
        self._committing = False

        # Restore previous session if available
        self.root.after(200, self.load_session)

        # Check for updates on startup (in background thread)
        threading.Thread(target=self._auto_check_updates, daemon=True).start()

    def _sort_table(self, col):
        col_key = {
            "archivo":    lambda r: r['FILENAME'].lower(),
            "nombre_db":  lambda r: r['NOMBRE'].lower(),
            "iban":       lambda r: r['IBAN'].lower(),
            "importe":    lambda r: r['IMPORTE'],
            "concepto":   lambda r: clean_db_value(r.get('CONCEPTO_NORMA')).lower(),
            "estado":     lambda r: ("OK", "SIN IMPORTE", "AMBIGUO", "IBAN NO VÁLIDO", "ERROR").index(payment_issue(r)[0]),
        }
        col_labels = {
            "archivo": "Archivo PDF", "nombre_db": "Nombre Detectado",
            "iban": "IBAN", "importe": "Importe (€)", "concepto": "Concepto", "estado": "Estado",
        }
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col
            self._sort_reverse = False

        self.current_results.sort(key=col_key[col], reverse=self._sort_reverse)

        for c, label in col_labels.items():
            arrow = (" ▼" if self._sort_reverse else " ▲") if c == col else ""
            self.tree.heading(c, text=label + arrow, command=lambda _c=c: self._sort_table(_c))

        self.refresh_table()

    def _auto_check_updates(self):
        """Background check on startup — non-intrusive."""
        result = check_for_updates()
        if result and isinstance(result, tuple):
            new_ver, url, name = result
            self.root.after(0, lambda: self._prompt_update(new_ver, url, name))

    def _prompt_update(self, new_ver, url, name):
        """Show update dialog."""
        self.version_label.config(text=f"v{APP_VERSION} (nueva: v{new_ver})", foreground="red")
        resp = messagebox.askyesno(
            "Actualización disponible",
            f"Hay una nueva versión de RemesaPro (v{new_ver}).\n"
            f"Tu versión actual es v{APP_VERSION}.\n\n"
            f"¿Deseas actualizar ahora?",
            parent=self.root
        )
        if resp:
            self._do_update(url, name, new_ver)

    def check_updates_manual(self):
        """Manual check triggered by button click."""
        self.lbl_status.config(text="Comprobando actualizaciones...")
        self.root.update()
        try:
            result = check_for_updates()
            if isinstance(result, str) and result.startswith("ERROR"):
                # Error message returned
                self.lbl_status.config(text="Error al comprobar")
                messagebox.showerror("Error de actualización",
                                     f"{result}\n\nAPI: {GITHUB_API_LATEST}",
                                     parent=self.root)
            elif result and isinstance(result, tuple):
                new_ver, url, name = result
                self._prompt_update(new_ver, url, name)
            else:
                self.lbl_status.config(text="Listo")
                messagebox.showinfo("Sin actualizaciones",
                                    f"Ya tienes la última versión (v{APP_VERSION}).",
                                    parent=self.root)
        except Exception as e:
            self.lbl_status.config(text="Error al comprobar")
            messagebox.showerror("Error", f"No se pudo comprobar actualizaciones:\n{e}",
                                 parent=self.root)

    def _do_update(self, url, name, new_ver):
        """Download and apply the update."""
        def status_cb(msg):
            self.root.after(0, lambda: self.lbl_status.config(text=msg))

        def run():
            success = download_and_apply_update(url, name, status_cb)
            if success:
                self.root.after(0, lambda: self._restart_after_update(new_ver))
            else:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", "No se pudo actualizar. Inténtalo de nuevo.", parent=self.root))

        threading.Thread(target=run, daemon=True).start()

    def _restart_after_update(self, new_ver):
        """Prompt user to restart the app after successful update."""
        self.version_label.config(text=f"v{new_ver} ✓", foreground="green")
        resp = messagebox.askyesno(
            "Actualización completada",
            f"RemesaPro se ha actualizado a v{new_ver}.\n"
            f"¿Reiniciar ahora?",
            parent=self.root
        )
        if resp:
            # Restart the application
            exe = sys.executable
            if getattr(sys, 'frozen', False):
                # PyInstaller frozen app
                os.execv(exe, [exe])
            else:
                os.execv(sys.executable, [sys.executable] + sys.argv)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f: return json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        return {}

    def save_config(self):
        self.config["last_folder"] = self.folder_var.get()
        self.config["last_db"] = self.db_var.get()
        try:
            with open(CONFIG_FILE, 'w') as f: json.dump(self.config, f)
        except IOError:
            pass

    def save_session(self):
        if not self.current_results:
            return
        try:
            serializable = []
            for r in self.current_results:
                entry = dict(r)
                # Convert AMBIGUOUS_CANDIDATES tuples to lists for JSON
                if entry.get('AMBIGUOUS_CANDIDATES'):
                    entry['AMBIGUOUS_CANDIDATES'] = [list(c) for c in entry['AMBIGUOUS_CANDIDATES']]
                serializable.append(entry)
            session = {
                'saved_at': datetime.now().strftime("%d/%m/%Y %H:%M"),
                'folder': self.folder_var.get(),
                'db': self.db_var.get(),
                'results': serializable,
            }
            with open(SESSION_FILE, 'w', encoding='utf-8') as f:
                json.dump(session, f, ensure_ascii=False, indent=2)
        except IOError:
            pass

    def load_session(self):
        if not os.path.exists(SESSION_FILE):
            return
        try:
            with open(SESSION_FILE, 'r', encoding='utf-8') as f:
                session = json.load(f)
        except (json.JSONDecodeError, IOError):
            return

        results = session.get('results', [])
        if not results:
            return

        saved_at = session.get('saved_at', 'desconocido')
        folder = session.get('folder', '')
        resp = messagebox.askyesno(
            "Sesión anterior encontrada",
            f"Se encontró una sesión guardada el {saved_at}.\n"
            f"Carpeta: {os.path.basename(folder) or folder}\n"
            f"Registros: {len(results)}\n\n"
            "¿Restaurar la sesión anterior?",
            parent=self.root
        )
        if not resp:
            return

        # Restore folder/db fields
        if folder:
            self.folder_var.set(folder)
        db = session.get('db', '')
        if db:
            self.db_var.set(db)

        # Restore results (convert AMBIGUOUS_CANDIDATES back to list of tuples)
        for r in results:
            if r.get('AMBIGUOUS_CANDIDATES'):
                r['AMBIGUOUS_CANDIDATES'] = [tuple(c) for c in r['AMBIGUOUS_CANDIDATES']]
        self.current_results = results

        # Try to reload the DB so edits/additions still work
        if db and os.path.exists(db):
            self.loaded_db_df = load_database(db)

        self.refresh_table()

    def select_folder(self):
        f = filedialog.askdirectory(title="Selecciona Carpeta de PDFs", initialdir=self.config.get("last_folder", "."))
        if f: self.folder_var.set(f)

    def select_db(self):
        f = filedialog.askopenfilename(title="Selecciona Base de Datos", filetypes=[("Excel Files", "*.xlsx")], initialdir=os.path.dirname(self.config.get("last_db", ".")))
        if f: self.db_var.set(f)

    def _show_context_menu(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        if item_id not in self.tree.selection():
            self.tree.selection_set(item_id)
        self.tree.focus(item_id)
        column_name = self._column_name(self.tree.identify_column(event.x))
        if column_name in self.EDITABLE_COLUMNS:
            self._focus_column = column_name
        self._context_menu.tk_popup(event.x_root, event.y_root)

    def _edit_selected_row(self):
        items = self._selected_items()
        index = self._row_index(items[0]) if items else None
        if index is None:
            return
        self._open_row_dialog(index)

    def _open_row_dialog(self, index):
        result_item = self.current_results[index]
        if result_item.get('AMBIGUOUS_CANDIDATES'):
            self.show_ambiguity_resolver(result_item)
        else:
            EditDialog(self.root, result_item, self.loaded_db_df, self.on_edit_save)

    def _delete_selected_row(self):
        indices = sorted(set(self._selected_indices()), reverse=True)
        if not indices:
            return
        if len(indices) == 1:
            question = (f"¿Eliminar '{self.current_results[indices[0]]['FILENAME']}' de la remesa?"
                        "\n\nNo se borrará el archivo original.")
        else:
            question = (f"¿Eliminar {len(indices)} registros de la remesa?"
                        "\n\nNo se borrarán los archivos originales.")
        if not messagebox.askyesno("Eliminar registros", question, parent=self.root):
            return
        for index in indices:
            del self.current_results[index]
        self.refresh_table()

    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return "break"
        column_name = self._column_name(self.tree.identify_column(event.x))

        # Doble clic en Nombre/IBAN/Importe/Concepto → edición directa en la tabla
        if column_name in self.EDITABLE_COLUMNS:
            self.tree.selection_set(item_id)
            self.tree.focus(item_id)
            self._focus_column = column_name
            self._begin_edit(item_id, column_name)
            return "break"

        # Doble clic en Archivo/Estado → ventana de detalle (o resolver ambigüedad)
        index = self._row_index(item_id)
        if index is not None:
            self._open_row_dialog(index)
        return "break"

    
    def show_ambiguity_resolver(self, result_item):
        candidates = result_item['AMBIGUOUS_CANDIDATES']
        
        def on_select(name, iban):
            # Update the result item
            result_item['NOMBRE'] = name
            result_item['IBAN'] = iban
            result_item['AMBIGUOUS_CANDIDATES'] = None  # Clear ambiguity
            
            # Also update concept from DB
            if self.loaded_db_df is not None:
                match = self.loaded_db_df[self.loaded_db_df['NOMBRE'] == name]
                if not match.empty:
                    result_item['CONCEPTO_NORMA'] = (clean_db_value(match.iloc[0].get('CONCEPTO_NORMA'))
                                                     or result_item['CONCEPTO_NORMA'])
            
            # Refresh table
            self.refresh_table()
        
        def on_manual_edit():
            # Open the manual edit dialog instead
            EditDialog(self.root, result_item, self.loaded_db_df, self.on_edit_save)
        
        AmbiguityResolverDialog(self.root, candidates, on_select, on_manual_edit)

    def on_edit_save(self, updated_item, add_to_db):
        if add_to_db and self.loaded_db_df is not None:
            self.save_new_db_entry(updated_item['NOMBRE'], updated_item['IBAN'])

        # Refresh GUI
        self.refresh_table()

    def _warn_invalid_db_ibans(self):
        """Avisa (sin bloquear) de los IBAN incorrectos que haya en la base de datos."""
        if self.loaded_db_df is None:
            return
        wrong = []
        for _, row in self.loaded_db_df.iterrows():
            iban = clean_db_value(row.get('IBAN'))
            name = clean_db_value(row.get('NOMBRE'))
            if not name:
                continue
            reason = iban_error(iban)
            if reason:
                wrong.append(f"• {name}: {reason}")
        if not wrong:
            return
        shown = "\n".join(wrong[:12])
        if len(wrong) > 12:
            shown += f"\n… y {len(wrong) - 12} más"
        messagebox.showwarning(
            "IBAN incorrectos en la base de datos",
            f"{len(wrong)} proveedor(es) de la base de datos tienen un IBAN que el banco "
            f"rechazaría:\n\n{shown}\n\n"
            "Sus líneas aparecerán marcadas y no se incluirán en el fichero SEPA.",
            parent=self.root)

    def save_new_db_entry(self, name, iban):
        reason = iban_error(iban)
        if reason and not messagebox.askyesno(
            "IBAN no válido",
            f"El IBAN de «{name}» no es válido: {reason}.\n\n"
            "¿Guardarlo igualmente en la base de datos?",
            parent=self.root):
            return
        try:
            # Add to memory DF
            new_row = {"NOMBRE": name, "IBAN": iban, "CONCEPTO_NORMA": "Añadido Manualmente"}
            self.loaded_db_df = pd.concat([self.loaded_db_df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Save to File
            db_path = self.db_var.get()
            try:
                self.loaded_db_df.to_excel(db_path, index=False)
                messagebox.showinfo("Base de Datos", f"Se ha añadido '{name}' a la base de datos.")
            except PermissionError:
                messagebox.showwarning("Aviso", "No se pudo guardar en el Excel de Base de Datos porque está abierto. Se ha actualizado en memoria para esta sesión, pero no se guardará en el disco.")
            except Exception as e:
                messagebox.showerror("Error DB", f"Error al guardar en BD: {e}")

        except Exception as e:
            print(f"Error saving DB entry: {e}")

    def start_processing_thread(self):
        self.btn_process.config(state="disabled")
        self.btn_save.config(state="disabled")
        self.btn_sepa.config(state="disabled")
        self.lbl_status.config(text="Procesando...")
        self.tree.delete(*self.tree.get_children())
        self.current_results = []
        
        t = threading.Thread(target=self.run_process)
        t.start()

    def _update_progress(self, current, total):
        pct = int(current / total * 100) if total else 0
        self.progress_var.set(pct)
        self.lbl_status.config(text=f"Procesando... {current}/{total}")

    def run_process(self):
        try:
            folder = self.folder_var.get()
            db_file = self.db_var.get()

            if not folder or not os.path.isdir(folder):
                messagebox.showerror("Error", "Carpeta inválida.")
                return

            if not db_file or not os.path.exists(db_file):
                messagebox.showerror("Error", "Base inválida.")
                return

            self.loaded_db_df = load_database(db_file)
            if self.loaded_db_df is None:
                messagebox.showerror("Error", "Error cargando BD.")
                return
            self.root.after(0, self._warn_invalid_db_ibans)

            def progress_cb(current, total):
                self.root.after(0, lambda c=current, t=total: self._update_progress(c, t))

            self.current_results = generate_remesa_data(folder, self.loaded_db_df, progress_cb)
            self.root.after(0, self.refresh_table)

        except Exception as e:
            self.root.after(0, lambda err=e: messagebox.showerror("Error", f"Error al procesar: {err}"))
        finally:
            self.root.after(0, lambda: self.progress_var.set(0))
            self.root.after(0, lambda: self.btn_process.config(state="normal"))

    def refresh_table(self, persist=True):
        previous_selection = self._selected_indices()
        self._close_editor()
        self.tree.delete(*self.tree.get_children())
        self._item_by_index = {}
        if not self.current_results:
            self.lbl_status.config(text="Sin resultados.")
            return

        filter_problems = self.filter_var.get()
        search = normalize_text(self.search_var.get())
        visible_count = 0
        problem_count = 0

        for idx, r in enumerate(self.current_results):
            status_text, tag, is_problem, _reason = payment_issue(r)
            if is_problem:
                problem_count += 1
            
            # Skip OK entries if filter is active
            if filter_problems and not is_problem:
                continue

            # Caja de búsqueda: archivo, proveedor, IBAN, concepto o importe
            if search:
                haystack = normalize_text(" ".join([
                    r['FILENAME'], r['NOMBRE'], r['IBAN'],
                    str(r.get('CONCEPTO_NORMA', '')), f"{r['IMPORTE']:.2f}",
                ]))
                if search not in haystack:
                    continue
            
            display_name = r['NOMBRE']
            if display_name.startswith("AMBIGUO:") or display_name.startswith("REVISAR:"):
                pass

            # Include actual index as first (hidden) value
            item_id = self.tree.insert("", "end", values=(
                idx,  # Hidden index for proper mapping
                r['FILENAME'],
                display_name,
                r['IBAN'],
                f"{r['IMPORTE']:.2f}",
                r.get('CONCEPTO_NORMA', ''),
                status_text
            ), tags=(tag,))
            self._item_by_index[idx] = item_id
            visible_count += 1

        # Mantener seleccionadas las mismas filas tras redibujar
        self._select_indices(previous_selection)
        
        self.btn_save.config(state="normal")
        self.btn_sepa.config(state="normal")
        self.btn_sepa_preview.config(state="normal")

        total_amount = sum(r['IMPORTE'] for r in self.current_results)
        total_str = f"{total_amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # Update status with counts
        if search:
            self.lbl_status.config(
                text=f"Buscando '{self.search_var.get()}': {visible_count} de {len(self.current_results)} registros.")
        elif filter_problems:
            self.lbl_status.config(text=f"Mostrando {visible_count} problemas de {len(self.current_results)} archivos.")
        else:
            ok_count = len(self.current_results) - problem_count
            self.lbl_status.config(text=f"Procesados {len(self.current_results)} archivos · ✅ {ok_count} | ⚠️ {problem_count} | Total: {total_str} €")
        
        if persist:
            self.save_config()
            self.save_session()

    # ── Utilidades de tabla ──────────────────────────────────────────────
    def _column_name(self, column_id):
        """'#3' → 'nombre_db'."""
        try:
            position = int(str(column_id).replace("#", "")) - 1
        except (TypeError, ValueError):
            return None
        columns = list(self.tree["columns"])
        return columns[position] if 0 <= position < len(columns) else None

    def _column_id(self, column_name):
        """'nombre_db' → '#3'."""
        columns = list(self.tree["columns"])
        return f"#{columns.index(column_name) + 1}"

    def _row_index(self, item_id):
        """Índice real en current_results de una fila de la tabla."""
        if not item_id:
            return None
        values = self.tree.item(item_id, 'values')
        if not values:
            return None
        try:
            return int(values[0])
        except (TypeError, ValueError):
            return None

    def _selected_items(self):
        items = list(self.tree.selection())
        if not items:
            focused = self.tree.focus()
            items = [focused] if focused else []
        return sorted(items, key=self.tree.index)

    def _selected_indices(self):
        return [i for i in (self._row_index(item) for item in self._selected_items()) if i is not None]

    def _select_indices(self, indices):
        items = [self._item_by_index[i] for i in indices if i in self._item_by_index]
        if not items:
            return
        self.tree.selection_set(items)
        self.tree.focus(items[0])
        self.tree.see(items[0])

    def _on_tree_select(self, event=None):
        """Explica junto a la búsqueda por qué la línea seleccionada no se puede pagar."""
        indices = self._selected_indices()
        if len(indices) == 1:
            _status, _tag, is_problem, reason = payment_issue(self.current_results[indices[0]])
            if is_problem and reason:
                self.lbl_hint.config(text="⚠️  " + reason[0].upper() + reason[1:], foreground="#c0392b")
                return
        self.lbl_hint.config(text=self.HINT_TEXT, foreground="gray")

    def _on_tree_click(self, event):
        """Recuerda la columna pulsada para copiar/pegar y editar con F2."""
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        column_name = self._column_name(self.tree.identify_column(event.x))
        if column_name in self.EDITABLE_COLUMNS:
            self._focus_column = column_name

    # ── Datos de la base de datos ────────────────────────────────────────
    def _db_names(self):
        """Lista de proveedores de la BD para el autocompletado."""
        if self.loaded_db_df is None or 'NOMBRE' not in getattr(self.loaded_db_df, 'columns', []):
            return []
        return [str(n).strip() for n in self.loaded_db_df['NOMBRE'].dropna().tolist() if str(n).strip()]

    def _fill_from_db(self, record, name):
        """Rellena IBAN y concepto si el nombre existe tal cual en la BD."""
        row = find_db_row(self.loaded_db_df, name)
        if row is None:
            return False
        iban = clean_db_value(row.get('IBAN'))
        if iban:
            record['IBAN'] = iban
        concepto = clean_db_value(row.get('CONCEPTO_NORMA'))
        if concepto:
            record['CONCEPTO_NORMA'] = concepto
        return True

    # ── Edición en línea ─────────────────────────────────────────────────
    def _cell_text(self, record, column_name):
        if column_name == "nombre_db":
            return record['NOMBRE']
        if column_name == "iban":
            return record['IBAN']
        if column_name == "importe":
            return f"{record['IMPORTE']:.2f}"
        if column_name == "concepto":
            return record.get('CONCEPTO_NORMA', '')
        return ""

    def _apply_cell_value(self, index, column_name, value):
        """Escribe el valor de una celda en current_results. Devuelve True si cambió."""
        if not (0 <= index < len(self.current_results)):
            return False
        record = self.current_results[index]
        value = str(value).strip()

        if column_name == "nombre_db":
            if not value or value == record['NOMBRE']:
                return False
            record['NOMBRE'] = value
            record['AMBIGUOUS_CANDIDATES'] = None
            self._fill_from_db(record, value)
            return True

        if column_name == "iban":
            iban = value.upper().replace(" ", "")
            if iban == record['IBAN']:
                return False
            record['IBAN'] = iban
            return True

        if column_name == "importe":
            try:
                amount = parse_amount(value)
            except (ValueError, TypeError):
                self.lbl_status.config(text=f"Importe no válido: {value}")
                return False
            if amount == record['IMPORTE']:
                return False
            record['IMPORTE'] = amount
            return True

        if column_name == "concepto":
            if value == record.get('CONCEPTO_NORMA', ''):
                return False
            record['CONCEPTO_NORMA'] = value
            return True

        return False

    def _begin_edit(self, item_id, column_name):
        """Abre un editor encima de la celda, sin ventanas emergentes."""
        if not item_id or column_name not in self.EDITABLE_COLUMNS:
            return
        index = self._row_index(item_id)
        if index is None:
            return
        self._close_editor()

        self.tree.see(item_id)
        self.tree.update_idletasks()
        bbox = self.tree.bbox(item_id, self._column_id(column_name))
        if not bbox:
            return
        x, y, width, height = bbox
        record = self.current_results[index]

        if column_name == "nombre_db":
            editor = AutocompleteEntry(self.tree, values=self._db_names(), bd=1, relief=tk.SOLID)
            # Los callbacks llevan el editor concreto: si entretanto se abrió otro,
            # una llamada retrasada (p. ej. el FocusOut) no debe afectarle.
            editor.on_select = lambda name: self.root.after(1, lambda: self._commit_edit(source=editor))
            editor.on_commit = lambda: self._commit_edit(source=editor)
            editor.on_cancel = lambda: self._close_editor(source=editor)
            editor.on_tab = lambda shift: self._commit_edit(move=-1 if shift else 1, source=editor)
            editor.on_focus_out = lambda: self._commit_edit(source=editor)
        else:
            editor = tk.Entry(self.tree, bd=1, relief=tk.SOLID)

            def commit(move=0):
                self._commit_edit(move=move, source=editor)
                return "break"

            editor.bind("<Return>", lambda e: commit())
            editor.bind("<KP_Enter>", lambda e: commit())
            editor.bind("<Tab>", lambda e: commit(1))
            editor.bind("<Shift-Tab>", lambda e: commit(-1))
            editor.bind("<ISO_Left_Tab>", lambda e: commit(-1))
            editor.bind("<Escape>", lambda e: (self._close_editor(source=editor), "break")[1])
            editor.bind("<FocusOut>", lambda e: self._commit_edit(source=editor))

        editor.insert(0, self._cell_text(record, column_name))
        editor.select_range(0, tk.END)
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()

        self._editor = editor
        self._editor_item = item_id
        self._editor_column = column_name
        self._editor_index = index

    def _close_editor(self, source=None):
        if source is not None and source is not self._editor:
            return
        if self._editor is not None:
            try:
                self._editor.destroy()
            except tk.TclError:
                pass
        self._editor = None
        self._editor_item = None
        self._editor_column = None
        self._editor_index = None

    def _commit_edit(self, move=0, source=None):
        """Guarda la celda en edición y, si procede, salta a la siguiente."""
        if self._editor is None or self._committing:
            return
        if source is not None and source is not self._editor:
            return  # el editor ya se cerró o fue reemplazado por otro
        self._committing = True
        try:
            value = self._editor.get()
            column_name = self._editor_column
            index = self._editor_index
            self._close_editor()
            self._apply_cell_value(index, column_name, value)
            self.refresh_table()
            self._select_indices([index])
            if move:
                self._edit_neighbour(index, column_name, move)
        finally:
            self._committing = False

    def _edit_neighbour(self, index, column_name, step):
        """Salta a la celda editable anterior/siguiente de la misma fila."""
        columns = list(self.EDITABLE_COLUMNS)
        position = columns.index(column_name) + step
        if not (0 <= position < len(columns)):
            return
        item_id = self._item_by_index.get(index)
        if item_id:
            self._focus_column = columns[position]
            self.root.after(1, lambda: self._begin_edit(item_id, columns[position]))

    def _edit_focused_cell(self):
        """F2 / Enter: edita la celda activa de la fila seleccionada."""
        item_id = self.tree.focus() or (self._selected_items() or [None])[0]
        if item_id:
            self._begin_edit(item_id, self._focus_column)

    # ── Portapapeles ─────────────────────────────────────────────────────
    def _clipboard_set(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def _clipboard_get(self):
        try:
            return self.root.clipboard_get()
        except tk.TclError:
            return ""

    def _copy_cells(self, event=None):
        """Ctrl+C: copia la columna activa de las filas seleccionadas (compatible con Excel)."""
        items = self._selected_items()
        if not items:
            return "break"
        column_name = self._focus_column if self._focus_column in self.EDITABLE_COLUMNS else "nombre_db"
        position = list(self.tree["columns"]).index(column_name)
        lines = []
        for item in items:
            values = self.tree.item(item, 'values')
            if values:
                lines.append(str(values[position]))
        if not lines:
            return "break"
        self._clipboard_set("\n".join(lines))
        self.lbl_status.config(text=f"Copiado ({column_name}): {len(lines)} celda(s)")
        return "break"

    def _paste_cells(self, event=None):
        """Ctrl+V: pega en la columna activa; admite varias filas/columnas desde Excel."""
        text = self._clipboard_get()
        if not text:
            return "break"
        items = self._selected_items()
        if not items:
            return "break"

        lines = [l for l in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
        while lines and not lines[-1].strip():
            lines.pop()
        if not lines:
            return "break"

        # Un solo valor y varias filas seleccionadas → se replica en todas
        if len(lines) == 1 and "\t" not in lines[0] and len(items) > 1:
            lines = lines * len(items)

        targets = items if len(lines) <= len(items) else self._rows_from(items[0], len(lines))
        columns = list(self.tree["columns"])
        start = columns.index(self._focus_column if self._focus_column in self.EDITABLE_COLUMNS else "nombre_db")

        changed = 0
        for line, item in zip(lines, targets):
            index = self._row_index(item)
            if index is None:
                continue
            for offset, cell in enumerate(line.split("\t")):
                position = start + offset
                if position >= len(columns):
                    break
                column_name = columns[position]
                if column_name in self.EDITABLE_COLUMNS and self._apply_cell_value(index, column_name, cell):
                    changed += 1
        if changed:
            indices = [self._row_index(i) for i in targets]
            self.refresh_table()
            self._select_indices([i for i in indices if i is not None])
        self.lbl_status.config(text=f"Pegado: {changed} celda(s) actualizada(s)")
        return "break"

    def _rows_from(self, item_id, count):
        """Devuelve 'count' filas visibles a partir de item_id (incluida)."""
        children = list(self.tree.get_children())
        try:
            start = children.index(item_id)
        except ValueError:
            start = 0
        return children[start:start + count]

    def _copy_provider(self, event=None):
        """Ctrl+Mayús+C: copia el proveedor completo (nombre + IBAN + concepto)."""
        items = self._selected_items()
        index = self._row_index(items[0]) if items else None
        if index is None:
            return "break"
        record = self.current_results[index]
        self._provider_clipboard = {
            'NOMBRE': record['NOMBRE'],
            'IBAN': record['IBAN'],
            'CONCEPTO_NORMA': record.get('CONCEPTO_NORMA', ''),
        }
        self._clipboard_set("\t".join([
            self._provider_clipboard['NOMBRE'],
            self._provider_clipboard['IBAN'],
            self._provider_clipboard['CONCEPTO_NORMA'],
        ]))
        self.lbl_status.config(text=f"Proveedor copiado: {record['NOMBRE']}")
        return "break"

    def _paste_provider(self, event=None):
        """Ctrl+Mayús+V: aplica el proveedor copiado a todas las filas seleccionadas."""
        provider = self._provider_clipboard
        if not provider:
            parts = self._clipboard_get().split("\t")
            if len(parts) >= 2:
                provider = {
                    'NOMBRE': parts[0].strip(),
                    'IBAN': parts[1].strip(),
                    'CONCEPTO_NORMA': parts[2].strip() if len(parts) > 2 else '',
                }
        if not provider or not provider.get('NOMBRE'):
            self.lbl_status.config(text="No hay ningún proveedor copiado (usa Ctrl+Mayús+C).")
            return "break"

        indices = self._selected_indices()
        if not indices:
            return "break"
        for index in indices:
            record = self.current_results[index]
            record['NOMBRE'] = provider['NOMBRE']
            record['IBAN'] = provider['IBAN']
            if provider.get('CONCEPTO_NORMA'):
                record['CONCEPTO_NORMA'] = provider['CONCEPTO_NORMA']
            record['AMBIGUOUS_CANDIDATES'] = None
        self.refresh_table()
        self._select_indices(indices)
        self.lbl_status.config(text=f"Proveedor '{provider['NOMBRE']}' aplicado a {len(indices)} fila(s).")
        return "break"

    def save_results(self):
        if not self.current_results: return
        try:
            output_file = save_to_excel(self.current_results, TEMPLATE_FILE, OUTPUT_PREFIX)
            if output_file:
                messagebox.showinfo("Éxito", f"Guardado:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _check_sepa_config(self):
        """Bloquea la generación si los datos del ordenante no son correctos."""
        errors = sepa_config_errors(self.config)
        if not errors:
            return True
        messagebox.showerror(
            "Datos del ordenante incompletos",
            "No se puede generar el fichero SEPA porque los datos de la empresa "
            "no son correctos:\n\n• " + "\n• ".join(errors) +
            "\n\nRevísalos en «⚙ SEPA Config».",
            parent=self.root)
        self.open_sepa_config()
        return False

    def _confirm_excluded_rows(self):
        """Resume las líneas que se quedan fuera del fichero y pide confirmación."""
        problems = {}
        for r in self.current_results:
            status_text, _tag, is_problem, _reason = payment_issue(r)
            if is_problem:
                problems.setdefault(status_text, []).append(r['FILENAME'])
        if not problems:
            return True

        detail = "\n".join(
            f"• {status}: {len(files)} → {', '.join(files[:3])}"
            + (f" y {len(files) - 3} más" if len(files) > 3 else "")
            for status, files in sorted(problems.items())
        )
        total = sum(len(files) for files in problems.values())
        return messagebox.askyesno(
            "Líneas que se van a omitir",
            f"{total} línea(s) no se pueden pagar y quedarán fuera del fichero:\n\n{detail}\n\n"
            "¿Generar el SEPA XML solo con el resto?",
            parent=self.root)

    def _confirm_duplicates(self):
        """Avisa de posibles pagos duplicados (mismo IBAN e importe)."""
        duplicates = find_duplicate_payments(self.current_results)
        if not duplicates:
            return True
        detail = "\n".join(
            f"• {rows[0]['NOMBRE']} — {amount:,.2f} € x{len(rows)} ({', '.join(r['FILENAME'] for r in rows)})"
            .replace(",", "X").replace(".", ",").replace("X", ".")
            for (_iban, amount), rows in duplicates[:8]
        )
        return messagebox.askyesno(
            "¿Pagos duplicados?",
            "Hay líneas con el mismo IBAN y el mismo importe, y se pagarían por separado:\n\n"
            f"{detail}\n\n¿Continuar de todas formas?",
            parent=self.root)

    def _sepa_checks_ok(self):
        return self._check_sepa_config() and self._confirm_excluded_rows() and self._confirm_duplicates()

    def generate_sepa(self):
        if not self.current_results: return

        if not self._sepa_checks_ok():
            return

        try:
            exec_date_str = self.sepa_date_var.get().strip()
            try:
                exec_date = datetime.strptime(exec_date_str, "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                exec_date = datetime.now().strftime("%Y-%m-%d")
            output_file = generate_sepa_xml(self.current_results, self.config, exec_date=exec_date)
            if output_file:
                paid = [r for r in self.current_results if is_payable(r)]
                total = f"{sum(r['IMPORTE'] for r in paid):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                messagebox.showinfo("SEPA XML Generado",
                    f"Archivo SEPA generado correctamente:\n{os.path.abspath(output_file)}\n\n"
                    f"Transferencias: {len(paid)}\n"
                    f"Importe total: {total} €\n\n"
                    "Comprueba que estas dos cifras cuadran al subir el fichero al banco.")
            else:
                messagebox.showwarning("Aviso", "No hay transacciones válidas para generar el XML.")
        except SepaConfigError as e:
            messagebox.showerror("Datos del ordenante incompletos", str(e))
        except Exception as e:
            messagebox.showerror("Error SEPA", f"Error generando XML: {e}")

    def preview_sepa(self):
        if not self.current_results:
            return
        if not self._check_sepa_config():
            return
        try:
            exec_date_str = self.sepa_date_var.get().strip()
            try:
                exec_date = datetime.strptime(exec_date_str, "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                exec_date = datetime.now().strftime("%Y-%m-%d")

            import tempfile, os as _os
            with tempfile.NamedTemporaryFile(suffix=".xml", delete=False, mode='w') as tmp:
                tmp_path = tmp.name

            output_path = generate_sepa_xml(self.current_results, self.config,
                                            output_path=tmp_path, exec_date=exec_date)
            if not output_path:
                messagebox.showwarning("Aviso", "No hay transacciones válidas para previsualizar.")
                return

            with open(tmp_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            try:
                _os.remove(tmp_path)
            except OSError:
                pass

            SepaPreviewDialog(self.root, xml_content)

        except SepaConfigError as e:
            messagebox.showerror("Datos del ordenante incompletos", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Error generando vista previa: {e}")

    def open_sepa_config(self):
        def on_save(sepa_data):
            self.config.update(sepa_data)
            self.save_config()
            messagebox.showinfo("Config SEPA", "Configuración SEPA guardada.")

        SepaConfigDialog(self.root, self.config, on_save)

# --- Logic ---

REQUIRED_DB_COLUMNS = {'NOMBRE', 'IBAN'}

def _validate_db_schema(df):
    """Returns missing required columns, or empty set if schema is valid."""
    return REQUIRED_DB_COLUMNS - set(df.columns)

def load_database(db_path):
    try:
        df = pd.read_excel(db_path, engine='openpyxl')
        df.columns = [c.strip() for c in df.columns]
        missing = _validate_db_schema(df)
        if missing:
            messagebox.showerror(
                "Error en Base de Datos",
                f"La base de datos no tiene las columnas requeridas: {', '.join(sorted(missing))}\n\n"
                f"Columnas encontradas: {', '.join(df.columns.tolist())}"
            )
            return None
        return df
    except PermissionError:
        import shutil
        import time
        temp_path = db_path + f".temp_{int(time.time())}.xlsx"
        try:
            print(f"Archivo bloqueado. Copiando...")
            shutil.copy2(db_path, temp_path)
            if os.path.exists(temp_path):
                df = pd.read_excel(temp_path, engine='openpyxl')
                df.columns = [c.strip() for c in df.columns]
                missing = _validate_db_schema(df)
                if missing:
                    messagebox.showerror(
                        "Error en Base de Datos",
                        f"Columnas requeridas no encontradas: {', '.join(sorted(missing))}"
                    )
                    return None
                return df
            return None
        except Exception:
            return None
        finally:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except OSError:
                pass
    except Exception:
        return None

def extract_info_from_excel(xlsx_path, db_df):
    """Extract provider name and total amount from Excel expense report template.
    Template structure: Name in C2, Grand Total in J56 ('Cantidad total' label in H56).
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws = wb.active

        # 1. Amount: Fixed cell J56 (standard template)
        amount = 0.0
        try:
            amount = parse_amount(ws['J56'].value)
        except (ValueError, TypeError):
            pass

        # Fallback: search for "Cantidad total" label and read adjacent cell to the right
        if amount == 0.0:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and 'cantidad total' in str(cell.value).lower():
                        for offset in range(1, 5):
                            adj = ws.cell(row=cell.row, column=cell.column + offset)
                            if adj.value is not None:
                                try:
                                    amount = parse_amount(adj.value)
                                    if amount != 0.0:
                                        break
                                except (ValueError, TypeError):
                                    pass
                        break

        # 2. Name: Cell C2 (standard template), fallback to filename
        name_from_cell = ws['C2'].value
        name_from_cell = str(name_from_cell).strip().upper() if name_from_cell else None

        filename = os.path.basename(xlsx_path)
        name_from_file = None
        parts = filename.split('_')
        if len(parts) >= 2:
            candidate = parts[1]
            if len(candidate) > 2 and not candidate.isdigit():
                name_from_file = candidate.replace('.', ' ').strip().upper()

        # Cell value takes priority over filename hint
        name_hint = name_from_cell or name_from_file

        db_names = db_df['NOMBRE'].dropna().astype(str).tolist()
        final_name, status, ambiguous_candidates = find_best_match(name_hint, db_names, db_df, "")

        return final_name, amount, status, ambiguous_candidates

    except Exception as e:
        print(f"Excel extraction error ({os.path.basename(xlsx_path)}): {e}")
        return None, 0.0, "ERROR", None


def extract_info_from_pdf(pdf_path, db_df):
    try:
        reader = pypdf.PdfReader(pdf_path)
        # Read all pages (not just first) to handle multi-page documents
        text = "\n".join(
            page.extract_text() or "" for page in reader.pages
        )

        # 1. Amount — use pre-compiled regex
        # Cada candidato es (posición, valor, longitud del texto capturado):
        # si dos patrones casan en la misma posición gana el más largo, que es
        # el que abarca el número completo.
        amount = 0.0
        candidates = []
        for m in _RE_DECIMAL_AMOUNT.finditer(text):
            try:
                candidates.append((m.start(), parse_amount(m.group(1)), len(m.group(1))))
            except (ValueError, TypeError):
                pass
        # Thousands-only format: "1.256" → 1256 (European, no decimal part)
        for m in _RE_THOUSANDS_NODEC.finditer(text):
            try:
                val = float(m.group(1).replace('.', ''))
                candidates.append((m.start(), val, len(m.group(1))))
            except (ValueError, TypeError):
                pass
        for m in _RE_WHOLE_EURO.finditer(text):
            try:
                candidates.append((m.start(), float(m.group(1)), len(m.group(1))))
            except (ValueError, TypeError):
                pass

        # Search for total label — support multiple formats
        total_labels = ["total gastos", "cantidad total", "total"]
        total_idx = -1
        for label in total_labels:
            idx = text.lower().find(label)
            if idx != -1:
                total_idx = idx
                break

        if total_idx != -1:
            closest_val = None
            min_dist = 1000
            best_len = 0
            for start, val, length in candidates:
                dist = start - total_idx
                if not (0 < dist < 200):
                    continue
                # Más cercano a la etiqueta; a igual distancia, la captura más larga
                if dist < min_dist or (dist == min_dist and length > best_len):
                    min_dist = dist
                    best_len = length
                    closest_val = val
            amount = closest_val if closest_val is not None else (max([c[1] for c in candidates]) if candidates else 0.0)
        else:
            amount = max([c[1] for c in candidates]) if candidates else 0.0

        # 2. Name - Extract from PDF content first (most reliable), then filename
        filename = os.path.basename(pdf_path)
        name_from_pdf = None
        name_from_file = None

        # Priority 1: "Nombre: XXX" inside the PDF (expense report format)
        name_match = _RE_NOMBRE.search(text)
        if name_match:
            extracted = name_match.group(1).strip().upper()
            extracted = re.split(r"\n|Fecha:|Semestre:|Programa", extracted)[0].strip()
            if len(extracted) > 2:
                name_from_pdf = extracted

        # Priority 2: Filename patterns
        parts = filename.replace('.pdf', '').replace('.PDF', '').split('_')
        if len(parts) >= 2:
            candidate = parts[1].strip()
            if len(candidate) > 2 and not candidate.isdigit():
                name_from_file = candidate.replace('.', ' ').strip().upper()

        name_from_file = name_from_pdf or name_from_file

        db_names = db_df['NOMBRE'].dropna().astype(str).tolist()
        final_name, status, ambiguous_candidates = find_best_match(name_from_file, db_names, db_df, text)

        return final_name, amount, status, ambiguous_candidates

    except Exception as e:
        print(f"PDF extraction error ({os.path.basename(pdf_path)}): {e}")
        return None, 0.0, "ERROR", None

def find_best_match(name_from_file, db_names, db_df=None, pdf_text=""):
        final_name = None
        status = "NO_ENCONTRADO"
        ambiguous_candidates = None
        
        if name_from_file:
            norm_target = normalize_text(name_from_file)
            
            scored_candidates = []
            for db_name in db_names:
                norm_db = normalize_text(db_name)
                if norm_target == norm_db:
                    score = 1.0
                elif norm_target in norm_db:
                    score = 0.95
                else:
                    score = difflib.SequenceMatcher(None, norm_target, norm_db).ratio()
                
                if score > 0.6:
                    scored_candidates.append( (score, db_name) )
            
            scored_candidates.sort(key=lambda x: x[0], reverse=True)
            
            if not scored_candidates:
                final_name = name_from_file
                status = "NO_ENCONTRADO"
            elif len(scored_candidates) == 1:
                final_name = scored_candidates[0][1]
                status = "OK"
            else:
                top1_score, top1_name = scored_candidates[0]
                top2_score, top2_name = scored_candidates[1]
                
                if top1_score > 0.9 or (top1_score - top2_score > 0.15):
                    final_name = top1_name
                    status = "OK"
                else:
                    ambiguous_set = [n for s, n in scored_candidates if top1_score - s < 0.05]
                    
                    if db_df is not None:
                        ibans = db_df[db_df['NOMBRE'].isin(ambiguous_set)]['IBAN'].unique()
                        if len(ibans) == 1:
                            final_name = top1_name
                            status = "OK"
                        else:
                            final_name = f"AMBIGUO: {', '.join(ambiguous_set[:3])}"
                            status = "AMBIGUO"
                            ambiguous_candidates = ambiguous_set
                    else:
                        final_name = f"AMBIGUO: {', '.join(ambiguous_set[:3])}"
                        status = "AMBIGUO"
                        ambiguous_candidates = ambiguous_set

        else:
             status = "NO_ENCONTRADO"
             for n in sorted(db_names, key=len, reverse=True):
                 if normalize_text(n) in normalize_text(pdf_text):
                     final_name = n
                     status = "OK_TEXT"
                     break
        return final_name, status, ambiguous_candidates

def generate_remesa_data(folder_path, db_df, progress_callback=None):
    results = []
    files = [f for f in os.listdir(folder_path)
             if f.lower().endswith('.pdf') or f.lower().endswith('.xlsx')]
    if 'NOMBRE' not in db_df.columns: return []

    total = len(files)
    for i, filename in enumerate(files, 1):
        if progress_callback:
            progress_callback(i, total)
        filepath = os.path.join(folder_path, filename)
        if filename.lower().endswith('.xlsx'):
            extracted_name, amount, status, ambiguous_candidates = extract_info_from_excel(filepath, db_df)
        else:
            extracted_name, amount, status, ambiguous_candidates = extract_info_from_pdf(filepath, db_df)
        
        iban = ""
        concept = f"{DEFAULT_CONCEPT} {os.path.splitext(filename)[0]}"
        candidates_list = None
        
        if status.startswith("OK"):
            row = db_df[db_df['NOMBRE'] == extracted_name].iloc[0]
            # clean_db_value evita que un NaN de pandas acabe como "nan" en el banco
            iban = clean_db_value(row.get('IBAN'))
            concept = clean_db_value(row.get('CONCEPTO_NORMA')) or concept
        elif status == "AMBIGUO":
            extracted_name = "REVISAR: " + extracted_name
            iban = "AMBIGUO"
            if ambiguous_candidates:
                candidates_list = [(name, db_df[db_df['NOMBRE'] == name].iloc[0]['IBAN']) 
                                   for name in ambiguous_candidates if not db_df[db_df['NOMBRE'] == name].empty]
        else:
            extracted_name = extracted_name or f"NO NAME ({filename})"
            iban = "NO ENCONTRADO"

        results.append({
            'FILENAME': filename,
            'FULLPATH': filepath,
            'NOMBRE': extracted_name,
            'IBAN': iban,
            'IMPORTE': amount,
            'CONCEPTO_NORMA': concept,
            'AMBIGUOUS_CANDIDATES': candidates_list
        })
    return results

def save_to_excel(results, template_path, output_prefix):
    if not results: return None
    df_out = pd.DataFrame(results)
    save_cols = ['NOMBRE', 'IBAN', 'IMPORTE', 'CONCEPTO_NORMA']
    
    try:
        template_cols = save_cols
        if os.path.exists(template_path):
             try:
                 tdf = pd.read_excel(template_path, engine='openpyxl')
                 template_cols = tdf.columns
             except PermissionError: pass
        
        final_df = pd.DataFrame(columns=template_cols)
        for c in final_df.columns:
            if c in df_out.columns: final_df[c] = df_out[c]
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{output_prefix}{timestamp}.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Remesa')
            ws = writer.sheets['Remesa']
            
            red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
            
            iban_idx = None
            for idx, col in enumerate(final_df.columns):
                if col == 'IBAN': iban_idx = idx + 1
            
            for row in range(2, ws.max_row + 1):
                if iban_idx:
                    cell = ws.cell(row=row, column=iban_idx)
                    val = str(cell.value).strip().upper()
                    if "NO ENCONTRADO" in val: cell.fill = red_fill
                    elif "AMBIGUO" in val: cell.fill = yellow_fill
            
            for col in ws.columns:
                mx = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[get_column_letter(col[0].column)].width = mx + 2
                
        return output_file
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    root = tk.Tk()
    app = RemesaApp(root)
    root.mainloop()
