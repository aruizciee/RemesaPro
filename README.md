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
- **Excel remittance output** – Timestamped Excel file with color-coded status (green = OK, yellow = ambiguous, red = error)
- **Interactive disambiguation** – GUI dialogs to manually resolve ambiguous matches or edit data
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
- **Remesa Excel de salida** – Archivo Excel con marca de tiempo y estado codificado por colores (verde = OK, amarillo = ambiguo, rojo = error)
- **Desambiguación interactiva** – Diálogos GUI para resolver manualmente coincidencias ambiguas o editar datos
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
