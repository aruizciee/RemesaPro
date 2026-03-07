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

# Configuration defaults
DEFAULT_DB_FILE = "Base datos IBAN proveedores.xlsx"
TEMPLATE_FILE = "FA25_REMESA PAGOS SANTANDER_.xlsx"
OUTPUT_PREFIX = "REMESA_GENERADA_"
CONFIG_FILE = "remesa_config.json"
LOGO_FILE = "ciee logo.png"

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
        total_cell = ws['J56'].value
        if total_cell is not None:
            try:
                amount = float(str(total_cell).replace('.', '').replace(',', '.'))
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
                                    amount = float(str(adj.value).replace('.', '').replace(',', '.'))
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
        # Regex handles European format (1.234,56) and simple format (234,56 or 234.56)
        amount = 0.0
        matches = re.finditer(r"(\d+(?:\.\d{3})*,\d{2}|\d+\.\d{2})", text)
        candidates = []
        for m in matches:
            val_str = m.group(1).replace('.', '').replace(',', '.')
            try: candidates.append((m.start(), float(val_str)))
            except: pass
        
        total_idx = text.lower().find("cantidad total")
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

        # 2. Name - New Logic (Accent Insensitive)
        filename = os.path.basename(pdf_path)
        name_from_file = None
        parts = filename.split('_')
        if len(parts) >= 2:
            candidate = parts[1]
            if len(candidate) > 2 and not candidate.isdigit():
                name_from_file = candidate.replace('.', ' ').strip().upper()

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
