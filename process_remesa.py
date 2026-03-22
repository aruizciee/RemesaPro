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

APP_VERSION = 7  # Matches GitHub build number

# Configuration defaults
DEFAULT_DB_FILE = "Base datos IBAN proveedores.xlsx"
TEMPLATE_FILE = "FA25_REMESA PAGOS SANTANDER_.xlsx"
OUTPUT_PREFIX = "REMESA_GENERADA_"
CONFIG_FILE = "remesa_config.json"
LOGO_FILE = "ciee logo.png"

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
      - 28.92      → 28.92  (punto = decimal)
      - 28,92      → 28.92  (coma = decimal)
      - 1.234,56   → 1234.56 (punto = miles, coma = decimal)
      - 1,234.56   → 1234.56 (coma = miles, punto = decimal)
    Rule: when both separators are present, the rightmost one is the decimal.
    """
    if value is None:
        return 0.0
    # Already a numeric type (e.g. openpyxl float) — use directly
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
    else:
        # Only dot or no separator — standard float
        return float(s)                         # 28.92 → 28.92


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


def generate_sepa_xml(results, config, output_path=None):
    """Generate SEPA Credit Transfer XML (pain.001.001.03) from remesa results."""
    # Filter only valid transactions (with IBAN)
    valid = [r for r in results
             if r['IBAN'] and r['IBAN'] not in ('NO ENCONTRADO', 'AMBIGUO', '')]

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
    SubElement(initg, "Nm").text = cfg["sepa_nombre"]
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

    SubElement(pmt, "ReqdExctnDt").text = now.strftime("%Y-%m-%d")

    # Debtor
    dbtr = SubElement(pmt, "Dbtr")
    SubElement(dbtr, "Nm").text = cfg["sepa_nombre"]
    addr = SubElement(dbtr, "PstlAdr")
    SubElement(addr, "PstCd").text = cfg["sepa_cp"]
    SubElement(addr, "TwnNm").text = cfg["sepa_ciudad"]
    SubElement(addr, "CtrySubDvsn").text = cfg["sepa_provincia"]
    SubElement(addr, "Ctry").text = cfg["sepa_pais"]
    SubElement(addr, "AdrLine").text = cfg["sepa_direccion"]
    dbtr_org = SubElement(SubElement(SubElement(dbtr, "Id"), "OrgId"), "Othr")
    SubElement(dbtr_org, "Id").text = cfg["sepa_cif"]

    # Debtor Account
    dbtr_acct = SubElement(pmt, "DbtrAcct")
    SubElement(SubElement(dbtr_acct, "Id"), "IBAN").text = cfg["sepa_iban"]
    SubElement(dbtr_acct, "Ccy").text = "EUR"

    # Debtor Agent (Bank)
    dbtr_agt = SubElement(pmt, "DbtrAgt")
    SubElement(SubElement(dbtr_agt, "FinInstnId"), "BIC").text = cfg["sepa_bic"]

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
        SubElement(cdtr, "Nm").text = clean_name[:70]  # SEPA max 70 chars

        cdtr_addr = SubElement(cdtr, "PstlAdr")
        # Derive country from IBAN prefix (first 2 chars)
        iban = r['IBAN'].replace(" ", "")
        country = iban[:2].upper() if len(iban) >= 2 else cfg["sepa_pais"]
        SubElement(cdtr_addr, "Ctry").text = country

        cdtr_acct = SubElement(tx, "CdtrAcct")
        SubElement(SubElement(cdtr_acct, "Id"), "IBAN").text = iban

        rmt = SubElement(tx, "RmtInf")
        concept = r.get('CONCEPTO_NORMA', f"Pago-CIEE")
        SubElement(rmt, "Ustrd").text = concept[:140]  # SEPA max 140 chars

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
        tk.Entry(input_frame, textvariable=self.name_var, width=40).grid(row=0, column=1, pady=5)
        
        tk.Label(input_frame, text="IBAN:").grid(row=1, column=0, sticky="w")
        self.iban_var = tk.StringVar(value=result_data['IBAN'])
        tk.Entry(input_frame, textvariable=self.iban_var, width=40).grid(row=1, column=1, pady=5)
        
        tk.Label(input_frame, text="Importe:").grid(row=2, column=0, sticky="w")
        self.amount_var = tk.StringVar(value=str(result_data['IMPORTE']))
        tk.Entry(input_frame, textvariable=self.amount_var, width=20).grid(row=2, column=1, pady=5, sticky="w")

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

    def open_pdf(self):
        try:
            # We need to know full path. result_data only has filename usually? 
            # Wait, result_data needs full path or we reconstruct it.
            # Let's verify what result_data has.
            filepath = self.result_data.get('FULLPATH')
            if filepath and os.path.exists(filepath):
                os.startfile(filepath)
            else:
                messagebox.showerror("Error", "No se encuentra el archivo PDF.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir PDF: {e}")

    def save(self):
        # Update result data
        self.result_data['NOMBRE'] = self.name_var.get()
        self.result_data['IBAN'] = self.iban_var.get()
        try:
            self.result_data['IMPORTE'] = float(self.amount_var.get().replace(',','.'))
        except: pass
        
        # Callback to update Treeview
        if self.save_callback:
            self.save_callback(self.result_data, self.add_db_var.get())
        
        self.destroy()


# ── Auto-update ───────────────────────────────────────────────────────────
GITHUB_REPO = "aruizciee/RemesaPro"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def check_for_updates():
    """Check GitHub for a newer release. Returns (new_version, download_url) or None."""
    try:
        req = urllib_request.Request(GITHUB_API_LATEST, headers={"Accept": "application/vnd.github+json"})
        with urllib_request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode())
        tag = data.get("tag_name", "")  # e.g. "build-8"
        remote_version = int(tag.replace("build-", "")) if tag.startswith("build-") else 0
        if remote_version <= APP_VERSION:
            return None
        # Pick the right asset for this OS
        is_mac = platform.system() == "Darwin"
        suffix = "macOS.zip" if is_mac else ".exe"
        for asset in data.get("assets", []):
            if asset["name"].endswith(suffix):
                return (remote_version, asset["browser_download_url"], asset["name"])
        return None
    except Exception:
        return None


def download_and_apply_update(download_url, asset_name, status_callback=None):
    """Download the new version and replace the current executable."""
    try:
        if status_callback:
            status_callback("Descargando actualización...")

        # Download to temp location
        import tempfile
        tmp_dir = tempfile.mkdtemp()
        tmp_file = os.path.join(tmp_dir, asset_name)
        urllib_request.urlretrieve(download_url, tmp_file)

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


class RemesaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Remesas - CIEE Pro")
        self.root.geometry("1100x750")
        
        # Determine internal resource path for PyInstaller
        self.base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        
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

        ttk.Label(header_frame, text="Generador de Remesas", style="Header.TLabel").pack(side=tk.LEFT, padx=10)

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

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.btn_process = ttk.Button(btn_frame, text="🔍 1. Analizar PDFs", command=self.start_processing_thread)
        self.btn_process.pack(side=tk.LEFT, padx=5)
        
        self.btn_save = ttk.Button(btn_frame, text="💾 2. Guardar Excel", command=self.save_results, state="disabled")
        self.btn_save.pack(side=tk.LEFT, padx=5)

        self.btn_sepa = ttk.Button(btn_frame, text="🏦 3. Generar SEPA XML", command=self.generate_sepa, state="disabled")
        self.btn_sepa.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="⚙ SEPA Config", command=self.open_sepa_config).pack(side=tk.LEFT, padx=5)

        self.lbl_status = ttk.Label(btn_frame, text="Listo", font=("Helvetica", 9, "italic"))
        self.lbl_status.pack(side=tk.LEFT, padx=15)
        
        # Filter checkbox
        self.filter_var = tk.BooleanVar(value=False)
        self.chk_filter = ttk.Checkbutton(btn_frame, text="Mostrar solo problemas (Ambiguos + Errores)", 
                                          variable=self.filter_var, command=self.refresh_table)
        self.chk_filter.pack(side=tk.RIGHT, padx=10)
        
        ttk.Label(btn_frame, text="(Doble clic en fila para Editar/Abrir PDF)", foreground="gray").pack(side=tk.RIGHT)

        # Treeview
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Add hidden index column for proper mapping when filtered
        columns = ("idx", "archivo", "nombre_db", "iban", "importe", "estado")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        # Hide the index column
        self.tree.column("idx", width=0, stretch=False)
        self.tree.heading("idx", text="")
        
        self.tree.heading("archivo", text="Archivo PDF")
        self.tree.heading("nombre_db", text="Nombre Detectado")
        self.tree.heading("iban", text="IBAN")
        self.tree.heading("importe", text="Importe (€)")
        self.tree.heading("estado", text="Estado")
        
        self.tree.column("archivo", width=250)
        self.tree.column("nombre_db", width=250)
        self.tree.column("iban", width=250)
        self.tree.column("importe", width=80, anchor="e")
        self.tree.column("estado", width=100)
        
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
        
        # Bind Double Click
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        
        self.current_results = []
        self.loaded_db_df = None # Store loaded DB in memory

        # Check for updates on startup (in background thread)
        threading.Thread(target=self._auto_check_updates, daemon=True).start()

    def _auto_check_updates(self):
        """Background check on startup — non-intrusive."""
        result = check_for_updates()
        if result:
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
        result = check_for_updates()
        if result:
            new_ver, url, name = result
            self._prompt_update(new_ver, url, name)
        else:
            self.lbl_status.config(text="Listo")
            messagebox.showinfo("Sin actualizaciones",
                                f"Ya tienes la última versión (v{APP_VERSION}).",
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
            except: pass
        return {}

    def save_config(self):
        self.config["last_folder"] = self.folder_var.get()
        self.config["last_db"] = self.db_var.get()
        try:
            with open(CONFIG_FILE, 'w') as f: json.dump(self.config, f)
        except: pass

    def select_folder(self):
        f = filedialog.askdirectory(title="Selecciona Carpeta de PDFs", initialdir=self.config.get("last_folder", "."))
        if f: self.folder_var.set(f)

    def select_db(self):
        f = filedialog.askopenfilename(title="Selecciona Base de Datos", filetypes=[("Excel Files", "*.xlsx")], initialdir=os.path.dirname(self.config.get("last_db", ".")))
        if f: self.db_var.set(f)

    def on_tree_double_click(self, event):
        item_id = self.tree.selection()
        if not item_id: return
        
        # Get the actual index from the tree item's values (first hidden column)
        values = self.tree.item(item_id, 'values')
        if not values: return
        
        actual_idx = int(values[0])  # First value is the hidden index
        result_item = self.current_results[actual_idx]
        
        # Check if it's ambiguous and has candidates
        if result_item.get('AMBIGUOUS_CANDIDATES'):
            self.show_ambiguity_resolver(result_item)
        else:
            EditDialog(self.root, result_item, self.loaded_db_df, self.on_edit_save)
    
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
                    result_item['CONCEPTO_NORMA'] = match.iloc[0].get('CONCEPTO_NORMA', result_item['CONCEPTO_NORMA'])
            
            # Refresh table
            self.refresh_table()
        
        def on_manual_edit():
            # Open the manual edit dialog instead
            EditDialog(self.root, result_item, self.loaded_db_df, self.on_edit_save)
        
        AmbiguityResolverDialog(self.root, candidates, on_select, on_manual_edit)

    def on_edit_save(self, updated_item, add_to_db):
        # Update internal list
        # updated_item ref is already same object in list, but let's be safe
        
        # Reflect changes in UI?
        # Re-render table later? Or update specifically this row?
        # Let's just update valid display logic
        
        if add_to_db and self.loaded_db_df is not None:
            self.save_new_db_entry(updated_item['NOMBRE'], updated_item['IBAN'])

        # Refresh GUI
        self.refresh_table()

    def save_new_db_entry(self, name, iban):
        try:
            # Add to memory DF
            new_row = {"NOMBRE": name, "IBAN": iban, "CONCEPTO_NORMA": "Añadido Manualmente"}
            self.loaded_db_df = pd.concat([self.loaded_db_df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Save to File
            db_path = self.db_var.get()
            
            # Use append mode or rewrite? Pandas writes generic xlsx.
            # Ideally load, append, save.
            # Using locking safe logic from before?
            # Creating a helper for saving DB
            try:
                with pd.ExcelWriter(db_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    # Actually appending to an existing excel without destroying formats is hard with pandas.
                    # Best effort: Append to bottom.
                    # Or simpler: Just re-save the whole thing if structure is simple.
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
            
            self.current_results = generate_remesa_data(folder, self.loaded_db_df)
            self.root.after(0, self.refresh_table)
            
        except Exception as e:
            print(e)
        finally:
            self.root.after(0, lambda: self.btn_process.config(state="normal"))

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        if not self.current_results:
            self.lbl_status.config(text="Sin resultados.")
            return

        filter_problems = self.filter_var.get()
        visible_count = 0
        problem_count = 0

        for idx, r in enumerate(self.current_results):
            tag = "ok"
            status_text = "OK"
            is_problem = False
            
            if "NO ENCONTRADO" in r['IBAN'] or "ERROR" in r['NOMBRE']:
                tag = "error"
                status_text = "ERROR"
                is_problem = True
                problem_count += 1
            elif "AMBIGUO" in r['IBAN']:
                tag = "warn"
                status_text = "AMBIGUO"
                is_problem = True
                problem_count += 1
            
            # Skip OK entries if filter is active
            if filter_problems and not is_problem:
                continue
            
            display_name = r['NOMBRE']
            if display_name.startswith("AMBIGUO:") or display_name.startswith("REVISAR:"):
                # Clean up for display? Keep for awareness
                pass

            # Include actual index as first (hidden) value
            self.tree.insert("", "end", values=(
                idx,  # Hidden index for proper mapping
                r['FILENAME'],
                display_name,
                r['IBAN'],
                f"{r['IMPORTE']:.2f}",
                status_text
            ), tags=(tag,))
            visible_count += 1
        
        self.btn_save.config(state="normal")
        self.btn_sepa.config(state="normal")
        
        # Update status with counts
        if filter_problems:
            self.lbl_status.config(text=f"Mostrando {visible_count} problemas de {len(self.current_results)} archivos.")
        else:
            ok_count = len(self.current_results) - problem_count
            self.lbl_status.config(text=f"Procesados {len(self.current_results)} archivos (✅ {ok_count} | ⚠️ {problem_count}).")
        
        self.save_config()

    def save_results(self):
        if not self.current_results: return
        try:
            output_file = save_to_excel(self.current_results, TEMPLATE_FILE, OUTPUT_PREFIX)
            if output_file:
                messagebox.showinfo("Éxito", f"Guardado:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def generate_sepa(self):
        if not self.current_results: return

        # Check for problems
        problems = [r for r in self.current_results
                    if r['IBAN'] in ('NO ENCONTRADO', 'AMBIGUO', '')]
        if problems:
            resp = messagebox.askyesno(
                "Atención",
                f"Hay {len(problems)} registro(s) sin IBAN válido que se omitirán del XML.\n\n"
                "¿Deseas continuar generando el SEPA XML solo con los registros válidos?")
            if not resp:
                return

        try:
            output_file = generate_sepa_xml(self.current_results, self.config)
            if output_file:
                messagebox.showinfo("SEPA XML Generado",
                    f"Archivo SEPA generado correctamente:\n{output_file}\n\n"
                    f"Transacciones incluidas: {len([r for r in self.current_results if r['IBAN'] not in ('NO ENCONTRADO', 'AMBIGUO', '')])}")
            else:
                messagebox.showwarning("Aviso", "No hay transacciones válidas para generar el XML.")
        except Exception as e:
            messagebox.showerror("Error SEPA", f"Error generando XML: {e}")

    def open_sepa_config(self):
        def on_save(sepa_data):
            self.config.update(sepa_data)
            self.save_config()
            messagebox.showinfo("Config SEPA", "Configuración SEPA guardada.")

        SepaConfigDialog(self.root, self.config, on_save)

# --- Logic ---

def load_database(db_path):
    try:
        df = pd.read_excel(db_path, engine='openpyxl')
        df.columns = [c.strip() for c in df.columns]
        return df
    except PermissionError:
        import shutil
        import time
        temp_path = db_path + f".temp_{int(time.time())}.xlsx"
        try:
            print(f"🔒 Archivo Bloqueado. Copiando...")
            os.system(f'copy "{db_path}" "{temp_path}"')
            if os.path.exists(temp_path):
                df = pd.read_excel(temp_path, engine='openpyxl')
                df.columns = [c.strip() for c in df.columns]
                return df
            return None
        except: return None
        finally:
             try:
                if os.path.exists(temp_path): os.remove(temp_path)
             except: pass
    except: return None

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
        except:
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
                                except:
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
        first_page = reader.pages[0]
        text = first_page.extract_text()
        
        # 1. Amount
        # Regex captures amounts with/without decimals: 1.234,56 | 28,92 | 28.92 | 23 €
        amount = 0.0
        # Pattern for amounts with decimals
        decimal_matches = re.finditer(r"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})", text)
        candidates = []
        for m in decimal_matches:
            try: candidates.append((m.start(), parse_amount(m.group(1))))
            except: pass
        # Pattern for whole-number amounts followed by € (e.g. "23 €")
        whole_matches = re.finditer(r"(\d+)\s*€", text)
        for m in whole_matches:
            try: candidates.append((m.start(), float(m.group(1))))
            except: pass

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
            for start, val in candidates:
                dist = start - total_idx
                if 0 < dist < 200 and dist < min_dist:
                    min_dist = dist
                    closest_val = val
            amount = closest_val if closest_val is not None else (max([c[1] for c in candidates]) if candidates else 0.0)
        else:
             amount = max([c[1] for c in candidates]) if candidates else 0.0

        # 2. Name - Extract from filename and/or PDF content
        filename = os.path.basename(pdf_path)
        name_from_file = None

        # Try filename first: "SP26_ALBERTO RUIZ_compras en supersol.pdf"
        parts = filename.split('_')
        if len(parts) >= 2:
            candidate = parts[1]
            if len(candidate) > 2 and not candidate.isdigit():
                name_from_file = candidate.replace('.', ' ').strip().upper()

        # Fallback: extract from "Nombre: XXX" inside the PDF (new expense report format)
        if not name_from_file:
            name_match = re.search(r"[Nn]ombre:\s*(.+)", text)
            if name_match:
                extracted = name_match.group(1).strip().upper()
                # Clean up: stop at newline or next field label
                extracted = re.split(r"\n|Fecha:|Semestre:", extracted)[0].strip()
                if len(extracted) > 2:
                    name_from_file = extracted

        db_names = db_df['NOMBRE'].dropna().astype(str).tolist()
        
        final_name, status, ambiguous_candidates = find_best_match(name_from_file, db_names, db_df, text)
        
        return final_name, amount, status, ambiguous_candidates

    except Exception as e:
        print(e)
        return None, 0.0, "ERROR", None

def find_best_match(name_from_file, db_names, db_df=None, pdf_text=""):
        final_name = None
        status = "NO_ENCONTRADO"
        ambiguous_candidates = None
        
        if name_from_file:
            norm_target = normalize_text(name_from_file)
            
            # Calculate Scores for ALL DB names
            # Using SequenceMatcher to get ratio
            scored_candidates = []
            for db_name in db_names:
                norm_db = normalize_text(db_name)
                # Exact match normalized?
                if norm_target == norm_db:
                    score = 1.0
                elif norm_target in norm_db: # Substring normalized
                    score = 0.95
                else:
                    score = difflib.SequenceMatcher(None, norm_target, norm_db).ratio()
                
                if score > 0.6: # Threshold
                    scored_candidates.append( (score, db_name) )
            
            # Sort by score descending
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
                
                # Winner takes all check
                # If top1 is perfect (1.0) or significantly better than top2 (gap > 0.15)
                # Example: Angela Jimenez (0.95 match to Ángela Jiménez) vs Alicia Jimenez (0.7 match)
                if top1_score > 0.9 or (top1_score - top2_score > 0.15):
                    final_name = top1_name
                    status = "OK"
                else:
                    # Ambiguity Check with IBANs
                    ambiguous_set = [n for s, n in scored_candidates if top1_score - s < 0.05]
                    
                    if db_df is not None:
                        ibans = db_df[db_df['NOMBRE'].isin(ambiguous_set)]['IBAN'].unique()
                        if len(ibans) == 1:
                            final_name = top1_name # Same IBAN, pick best text match
                            status = "OK"
                        else:
                            final_name = f"AMBIGUO: {', '.join(ambiguous_set[:3])}"
                            status = "AMBIGUO"
                            ambiguous_candidates = ambiguous_set  # Return candidates for dialog
                    else:
                        final_name = f"AMBIGUO: {', '.join(ambiguous_set[:3])}"
                        status = "AMBIGUO"
                        ambiguous_candidates = ambiguous_set

        else:
             # Fallback text search
             status = "NO_ENCONTRADO"
             for n in sorted(db_names, key=len, reverse=True):
                 if normalize_text(n) in normalize_text(pdf_text):
                     final_name = n
                     status = "OK_TEXT"
                     break
        return final_name, status, ambiguous_candidates

def generate_remesa_data(folder_path, db_df):
    results = []
    files = [f for f in os.listdir(folder_path)
             if f.lower().endswith('.pdf') or f.lower().endswith('.xlsx')]
    if 'NOMBRE' not in db_df.columns: return []

    for filename in files:
        filepath = os.path.join(folder_path, filename)
        if filename.lower().endswith('.xlsx'):
            extracted_name, amount, status, ambiguous_candidates = extract_info_from_excel(filepath, db_df)
        else:
            extracted_name, amount, status, ambiguous_candidates = extract_info_from_pdf(filepath, db_df)
        
        iban = ""
        concept = f"Pago {filename[:20]}..."
        candidates_list = None
        
        if status.startswith("OK"):
            row = db_df[db_df['NOMBRE'] == extracted_name].iloc[0]
            iban = row.get('IBAN', '')
            concept = row.get('CONCEPTO_NORMA', concept)
        elif status == "AMBIGUO":
            extracted_name = "REVISAR: " + extracted_name
            iban = "AMBIGUO"
            # Store candidates with their IBANs
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
            'AMBIGUOUS_CANDIDATES': candidates_list  # Store for resolver dialog
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
