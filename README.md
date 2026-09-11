# RemesaPro

<p align="center">
  <img src="ciee logo.png" alt="CIEE Logo" width="200"/>
</p>

<p align="center">
  <strong>Automated expense report processing and SEPA payment generation for CIEE</strong><br/>
  <em>Procesamiento automatizado de notas de gasto y generación de remesas SEPA para CIEE</em>
</p>

---

## English

### What is RemesaPro?

RemesaPro is a desktop application that automates the processing of expense reports (PDFs and Excel files) and generates SEPA XML payment files and Excel remittances ready for bank import. It was built for CIEE (Centro Internacional de Estudios para el Español) to streamline their accounts payable workflow.

### Features

- **Multi-format processing** – Reads expense reports from PDF and Excel files
- **Intelligent provider matching** – Fuzzy-matches extracted names against a provider database using normalized text comparison
- **IBAN lookup** – Automatically retrieves each provider's IBAN from the database
- **SEPA XML generation** – Produces valid `pain.001.001.03` credit transfer files ready for bank import
- **Payment validation** – Every IBAN is checked (format, country length and mod-97 check digit), amounts of 0 € are held back, debtor details are verified before anything is written, accented and special characters are converted to the SEPA character set, and repeated payments (same IBAN and amount) are flagged
- **Excel remittance output** – Timestamped Excel file with color-coded status (green = OK, yellow = ambiguous, red = error)
- **Interactive disambiguation** – GUI dialogs to manually resolve ambiguous matches or edit data
- **Provider autocomplete** – Start typing a name and the matching providers from the database appear; picking one fills in the IBAN and concepto
- **In-place table editing** – Double-click any cell to edit name, IBAN, amount or concepto without opening a dialog
- **Row search and clipboard** – Filter the table by file, provider, IBAN or concepto, and copy/paste cells or whole providers between rows
- **Auto-update** – Checks GitHub releases for new versions and applies updates automatically
- **Persistent configuration** – Saves last-used paths and SEPA debtor info to a local JSON file

### Requirements

**To run from source:**
- Python 3.13+
- `pandas`, `pypdf`, `openpyxl`

**Required files (in the same folder as the executable or script):**
| File | Purpose |
|------|---------|
| `Base datos IBAN proveedores.xlsx` | Provider database (columns: `NOMBRE`, `IBAN`, `CONCEPTO_NORMA`) |
| `FA25_REMESA PAGOS SANTANDER_.xlsx` | Output template |
| `ciee logo.png` | Application logo |

### Installation

**Option 1 – Pre-built binary (recommended)**

Download the latest release for your platform from [GitHub Releases](https://github.com/aruizciee/RemesaPro/releases):
- **Windows**: `RemesaPro.exe`
- **macOS**: `RemesaPro-macOS.zip` → unzip and run `RemesaPro.app`

**Option 2 – Run from source**

```bash
# Install dependencies
pip install pandas pypdf openpyxl

# Run the application
python process_remesa.py
```

### Configuration

On first run, the application creates `remesa_config.json` in its directory. You can edit this file directly or use the built-in settings dialog:

```json
{
  "last_folder": "/path/to/expense/reports",
  "last_db": "/path/to/Base datos IBAN proveedores.xlsx",
  "sepa_nombre": "Your Company Name",
  "sepa_cif": "ES12345678A",
  "sepa_iban": "ES9121000418450200051332",
  "sepa_bic": "BVAFESBB",
  "sepa_direccion": "Street Address",
  "sepa_cp": "28001",
  "sepa_ciudad": "Madrid",
  "sepa_provincia": "Madrid",
  "sepa_pais": "ES"
}
```

### How It Works

```
PDF / Excel expense reports
         ↓
  Extract name & amount
         ↓
  Load provider database
         ↓
  Fuzzy-match name → IBAN
         ↓
  Resolve ambiguous cases (GUI)
         ↓
  Excel output  +  SEPA XML
         ↓
      Bank import
```

### Editing the results table

| Action | Shortcut |
| --- | --- |
| Edit a cell in place (name, IBAN, amount, concepto) | Double-click the cell, or `F2` / `Enter` |
| Provider autocomplete | Start typing in the name cell: providers beginning with those letters appear first and the list narrows with every letter. `↑` `↓` to move, `Enter` to pick, `Esc` to close |
| Move to the next / previous cell | `Tab` / `Shift+Tab` |
| Copy the selected cells | `Ctrl+C` (tab/newline separated, so it pastes into Excel) |
| Paste into the selected cells | `Ctrl+V` (a single value fills every selected row; several lines fill consecutive rows) |
| Copy / paste a whole provider (name + IBAN + concepto) | `Ctrl+Shift+C` / `Ctrl+Shift+V` |
| Remove rows from the remittance | `Del` (works on a multiple selection) |
| Open the detail window | Double-click the *File* or *Status* column, or use the right-click menu |

Picking a provider from the autocomplete list fills in its IBAN and concepto from the database.

Each row shows why it cannot be paid, and only `OK` rows reach the SEPA file:

| Status | Meaning |
| --- | --- |
| `OK` | Valid IBAN and a positive amount — it will be paid |
| `ERROR` | The provider was not found in the database |
| `AMBIGUO` | Several providers match: pick one |
| `IBAN NO VÁLIDO` | The IBAN fails the check digit, the country length or the format |
| `SIN IMPORTE` | The amount is 0 € or negative |

### Building from Source

The project uses PyInstaller to create standalone executables. GitHub Actions builds binaries automatically on every push to `main` that modifies `process_remesa.py` or `RemesaPro.spec`.

```bash
# Build manually
pip install pyinstaller pandas pypdf openpyxl
pyinstaller RemesaPro.spec
# Output: dist/RemesaPro.exe (Windows) or dist/RemesaPro (macOS)
```

### Provider Database Format

The file `Base datos IBAN proveedores.xlsx` must contain these columns:

| Column | Description |
|--------|-------------|
| `NOMBRE` | Full provider name |
| `IBAN` | Provider's bank IBAN |
| `CONCEPTO_NORMA` | Payment description/concept |

---

## Español

### ¿Qué es RemesaPro?

RemesaPro es una aplicación de escritorio que automatiza el procesamiento de notas de gasto (PDFs y archivos Excel) y genera ficheros XML SEPA y remesas Excel listas para importar en el banco. Fue desarrollada para CIEE (Centro Internacional de Estudios para el Español) con el fin de agilizar el flujo de trabajo de cuentas a pagar.

### Funcionalidades

- **Procesamiento multiformato** – Lee notas de gasto desde PDFs y archivos Excel
- **Búsqueda inteligente de proveedores** – Coincidencia aproximada de nombres contra la base de datos mediante comparación de texto normalizado
- **Búsqueda de IBAN** – Recupera automáticamente el IBAN de cada proveedor desde la base de datos
- **Generación de XML SEPA** – Produce ficheros de transferencia de crédito `pain.001.001.03` válidos para importar en el banco
- **Validación de los pagos** – Comprueba cada IBAN (formato, longitud del país y dígito de control mod 97), retiene los importes de 0 €, verifica los datos del ordenante antes de escribir nada, adapta acentos y símbolos al juego de caracteres SEPA y avisa de posibles pagos duplicados (mismo IBAN e importe)
- **Remesa Excel de salida** – Archivo Excel con marca de tiempo y estado codificado por colores (verde = OK, amarillo = ambiguo, rojo = error)
- **Desambiguación interactiva** – Diálogos GUI para resolver manualmente coincidencias ambiguas o editar datos
- **Autocompletado de proveedores** – Al escribir un nombre aparecen los proveedores de la base de datos que coinciden; al elegir uno se rellenan el IBAN y el concepto
- **Edición directa en la tabla** – Doble clic en cualquier celda para editar nombre, IBAN, importe o concepto sin abrir ninguna ventana
- **Búsqueda y portapapeles** – Filtra la tabla por archivo, proveedor, IBAN o concepto, y copia/pega celdas o proveedores completos entre filas
- **Actualización automática** – Comprueba las versiones en GitHub Releases y aplica actualizaciones automáticamente
- **Configuración persistente** – Guarda las rutas utilizadas y los datos del ordenante SEPA en un archivo JSON local

### Requisitos

**Para ejecutar desde el código fuente:**
- Python 3.13+
- `pandas`, `pypdf`, `openpyxl`

**Archivos necesarios (en la misma carpeta que el ejecutable o el script):**
| Archivo | Propósito |
|---------|-----------|
| `Base datos IBAN proveedores.xlsx` | Base de datos de proveedores (columnas: `NOMBRE`, `IBAN`, `CONCEPTO_NORMA`) |
| `FA25_REMESA PAGOS SANTANDER_.xlsx` | Plantilla de salida |
| `ciee logo.png` | Logo de la aplicación |

### Instalación

**Opción 1 – Ejecutable pre-compilado (recomendado)**

Descarga la última versión para tu plataforma desde [GitHub Releases](https://github.com/aruizciee/RemesaPro/releases):
- **Windows**: `RemesaPro.exe`
- **macOS**: `RemesaPro-macOS.zip` → descomprime y ejecuta `RemesaPro.app`

**Opción 2 – Ejecutar desde el código fuente**

```bash
# Instalar dependencias
pip install pandas pypdf openpyxl

# Ejecutar la aplicación
python process_remesa.py
```

### Configuración

En el primer arranque, la aplicación crea `remesa_config.json` en su directorio. Puedes editar este archivo directamente o usar el diálogo de configuración integrado:

```json
{
  "last_folder": "/ruta/a/notas/de/gasto",
  "last_db": "/ruta/a/Base datos IBAN proveedores.xlsx",
  "sepa_nombre": "Nombre de tu empresa",
  "sepa_cif": "ES12345678A",
  "sepa_iban": "ES9121000418450200051332",
  "sepa_bic": "BVAFESBB",
  "sepa_direccion": "Dirección postal",
  "sepa_cp": "28001",
  "sepa_ciudad": "Madrid",
  "sepa_provincia": "Madrid",
  "sepa_pais": "ES"
}
```

### Cómo funciona

```
PDFs / Excel con notas de gasto
              ↓
    Extracción de nombre e importe
              ↓
    Carga de base de datos de proveedores
              ↓
    Coincidencia aproximada nombre → IBAN
              ↓
    Resolución de casos ambiguos (GUI)
              ↓
    Salida Excel  +  XML SEPA
              ↓
         Importación bancaria
```

### Edición de la tabla de resultados

| Acción | Atajo |
| --- | --- |
| Editar una celda en la propia tabla (nombre, IBAN, importe, concepto) | Doble clic en la celda, o `F2` / `Intro` |
| Autocompletado de proveedores | Al escribir en la celda de nombre aparecen primero los proveedores que empiezan por esas letras y la lista se reduce con cada letra. `↑` `↓` para moverse, `Intro` para elegir y `Esc` para cerrar |
| Ir a la celda siguiente / anterior | `Tab` / `Mayús+Tab` |
| Copiar las celdas seleccionadas | `Ctrl+C` (separado por tabuladores y saltos de línea, se pega en Excel) |
| Pegar en las celdas seleccionadas | `Ctrl+V` (un solo valor se replica en todas las filas seleccionadas; varias líneas rellenan filas consecutivas) |
| Copiar / pegar el proveedor completo (nombre + IBAN + concepto) | `Ctrl+Mayús+C` / `Ctrl+Mayús+V` |
| Quitar filas de la remesa | `Supr` (admite selección múltiple) |
| Abrir la ventana de detalle | Doble clic en la columna *Archivo* o *Estado*, o menú del botón derecho |

Al elegir un proveedor en la lista de autocompletado se rellenan su IBAN y su concepto desde la base de datos.

Cada fila indica por qué no se puede pagar, y solo las que están en `OK` llegan al fichero SEPA:

| Estado | Significado |
| --- | --- |
| `OK` | IBAN válido e importe positivo: se paga |
| `ERROR` | No se encontró el proveedor en la base de datos |
| `AMBIGUO` | Hay varios proveedores posibles: elige uno |
| `IBAN NO VÁLIDO` | El IBAN falla el dígito de control, la longitud del país o el formato |
| `SIN IMPORTE` | El importe es 0 € o negativo |

### Compilación desde el código fuente

El proyecto usa PyInstaller para crear ejecutables independientes. GitHub Actions compila los binarios automáticamente en cada push a `main` que modifique `process_remesa.py` o `RemesaPro.spec`.

```bash
# Compilar manualmente
pip install pyinstaller pandas pypdf openpyxl
pyinstaller RemesaPro.spec
# Resultado: dist/RemesaPro.exe (Windows) o dist/RemesaPro (macOS)
```

### Formato de la base de datos de proveedores

El archivo `Base datos IBAN proveedores.xlsx` debe contener estas columnas:

| Columna | Descripción |
|---------|-------------|
| `NOMBRE` | Nombre completo del proveedor |
| `IBAN` | IBAN bancario del proveedor |
| `CONCEPTO_NORMA` | Descripción / concepto del pago |

---

## License / Licencia

This project is proprietary software developed for CIEE internal use.
Este proyecto es software propietario desarrollado para uso interno de CIEE.
