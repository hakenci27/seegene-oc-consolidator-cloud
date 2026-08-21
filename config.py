"""
config.py
Ubicaciones de columnas/hojas en los archivos maestros de Seegene Mexico.
Si el formato de alguno de estos archivos cambia, este es el unico lugar
que hay que actualizar.
"""

# ---------------------------------------------------------------------------
# Nombres de archivo esperados dentro de la carpeta de trabajo
# ---------------------------------------------------------------------------
FILE_SALES_PROGRESS = "Sales progress.xlsx"
FILE_CUSTOMER_MASTER = "customer master list.xlsx"
FILE_DAILY_REPORT = "daily report.xlsx"
FILE_PRICE_LIST = "190626 Official sales price list.xlsx"
FILE_CREDIT_REPORT = "sales collection report.xlsx"
FILE_PRODUCT_MASTER = "Product master list.xlsx"
OCS_SUBFOLDER = "OCs"
OUTPUT_FILENAME = "Nuevas OC - Control de Ordenes de Compra.xlsx"

# ---------------------------------------------------------------------------
# Sales progress.xlsx > hoja "Purchase Order"
# ---------------------------------------------------------------------------
SP_SHEET = "Purchase Order"
SP_HEADER_ROW = 18
SP_FIRST_DATA_ROW = 19
SP_COL_PO_NO = "D"          # No. de OC ya registradas (para no duplicar)

# Encabezados de la tabla de salida (mismo orden que Sales progress, columna A..U)
OUTPUT_HEADERS_BASE = [
    "PIC", "PO Date", "Reception Date", "PO No", "Customer Code", "Cliente",
    "Code", "Description", "Category Product", "Qty", " Unit Price (A) ",
    "Back order", "Amount", "Categoria", "Mes de facturacion",
    "Inventario Actual", "Comments  (SCM)", "Status ", "Quarter", "Semester", "Year",
]
# Columnas extendidas agregadas para este flujo (V..AB)
OUTPUT_HEADERS_EXT = [
    "Precio coincide", "Lista de precio", "Limite de Credito (Balance)",
    "Dentro de Limite", "Inv Disponible (MX Final Inventory)",
    "Caducidad < 6M", "Fecha de Caducidad mas proxima",
]
OUTPUT_HEADERS = OUTPUT_HEADERS_BASE + OUTPUT_HEADERS_EXT

# Columnas que el usuario llena a mano y que el programa NUNCA debe sobrescribir
# en filas que ya existian en una corrida anterior.
MANUAL_COLUMNS = {"Back order", "Categoria", "Mes de facturacion", "Status "}

# Sales progress.xlsx > hoja "Material info." (categoria de producto)
MI_SHEET = "Material info."
MI_HEADER_ROW = 3
MI_FIRST_DATA_ROW = 4
MI_COL_CODE = "Code"
MI_COL_CATEGORY = "Cat Product"

# ---------------------------------------------------------------------------
# customer master list.xlsx > "Sheet1"
# ---------------------------------------------------------------------------
CM_SHEET = "Sheet1"
CM_HEADER_ROW = 1
CM_FIRST_DATA_ROW = 2
CM_COL_CUSTOMER_CODE = "Customer Code"      # a veces con salto de linea en el header real
CM_COL_CUSTOMER_NAME = "Customer name"
CM_COL_CLASSIFICATION = "Classification (by customer account)"
CM_COL_RFC = "RFC"
CM_COL_PIC = "Sales Representative"

# Alias conocidos: nombres de cliente tal como pueden venir mal escritos u
# ordenados distinto en el documento de la OC (typos, orden de palabras
# distinto, sufijos como "SA DE CV" pegados) -> nombre EXACTO tal como esta
# en "customer name" de customer master list.xlsx. Se revisan antes que la
# coincidencia parcial generica en find_customer(), asi que estos casos
# especificos siempre resuelven correcto sin depender de heuristicas.
# Agregar aqui cualquier variante nueva que se detecte en el futuro.
CUSTOMER_NAME_ALIASES = {
    "DIAGNOSTICO MOLECULAR DE REFERENCIA Y SEVICIO": "DIAGNOSTICO MOLECULAR Y SERVICIO DE REFERENCIA",  # customer code 11
    "LABORATORIO DE ANALISIS CLINICOS Y CENTRO DE MAGEN SANTA MARIA SA DE CV": "LABORATORIO DE ANALISIS CLINICOS Y CENTRO DE IMAGEN SANTA MARIA",  # customer code 32
}

# ---------------------------------------------------------------------------
# daily report.xlsx
# ---------------------------------------------------------------------------
DR_SALES_YTD_SHEET = "Sales_YTD"
DR_SALES_YTD_HEADER_ROW = 5
DR_SALES_YTD_FIRST_DATA_ROW = 6
DR_SYTD_COL_CUSTOMER = "CUSTOMER NAME"
DR_SYTD_COL_CODE = "CODE2"
DR_SYTD_COL_PRICE = "PRODUCT PRICE"
DR_SYTD_COL_DATE = "DATE"

DR_INVENTORY_SHEET = "Inventory"
DR_INV_HEADER_ROW = 4
DR_INV_FIRST_DATA_ROW = 6
DR_INV_COL_CATNO = "Cat No."
DR_INV_COL_EXPIRATION = "EXPIRATION DATE"
DR_INV_COL_MX_FINAL = "MX FINAL INVENTORY"

# ---------------------------------------------------------------------------
# 190626 Official sales price list.xlsx > "Price list for contracts"
# ---------------------------------------------------------------------------
PL_SHEET = "Price list for contracts"
PL_HEADER_ROW = 4
PL_FIRST_DATA_ROW = 5
PL_COL_CATNO = "Cat No."
PL_COL_NEW_LIST = "New List price"
PL_COL_NEW_DIST_A = "New Price\nDistribuidor A"
PL_COL_NEW_DIST_B = "New Price\nDistribuidor B"

# ---------------------------------------------------------------------------
# sales collection report.xlsx > "CREDITO"
# ---------------------------------------------------------------------------
CR_SHEET = "CREDITO"
CR_HEADER_ROW = 1
CR_FIRST_DATA_ROW = 2
CR_COL_RFC = "Identificador"
CR_COL_NAME = "Nombre"
CR_COL_BALANCE = "BALANCE"

# ---------------------------------------------------------------------------
# Product master list.xlsx > "Sheet1"
# Cruce de codigo de producto -> codigo normalizado + descripcion + categoria.
# Cuando la OC trae un codigo combinado (ej. "187937 / SD7700X"), se separa
# por "/" (u otros separadores) y se busca cada parte contra "Cat No" -- si
# hay match, ese codigo/descripcion/categoria del maestro reemplaza lo
# extraido del documento.
# ---------------------------------------------------------------------------
PM_SHEET = "Sheet1"
PM_HEADER_ROW = 1
PM_FIRST_DATA_ROW = 2
PM_COL_CODE = "Cat No"
PM_COL_DESCRIPTION = "Description"
PM_COL_CATEGORY = "Category"

# ---------------------------------------------------------------------------
# Clasificaciones de cliente -> que columna de precio de distribuidor usar
# ---------------------------------------------------------------------------
DISTRIBUTOR_A_LABEL = "Distributor A"
DISTRIBUTOR_B_LABEL = "Distributor B"

# ---------------------------------------------------------------------------
# Modelo de Claude a usar para la extraccion de OC (vision + texto)
# ---------------------------------------------------------------------------
CLAUDE_MODEL = "claude-sonnet-5"
