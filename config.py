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
    "GRUPO ARH LABORATORIOS": "CENTRO DE DIAGNOSTICO ARISTA",  # confirmado por el usuario, revision Mapeo Agosto 2026
    "LABORATORIO JUAREZ S.A. DE C.V.": "LABORATORIOS JUAREZ",  # confirmado por el usuario, revision Mapeo Agosto 2026
    "LABORATORIO TELLEZ GIRON S.A. DE C.V.": "LABORATORIO TELLEZ GIRON",  # confirmado por el usuario, revision Mapeo Agosto 2026 -- cliente 261 (typo "LABORATIRIO" corregido en customer master list.xlsx)
    "LABORATORIO DORADO": "MARTHA DORADO DEL RIO",  # confirmado por el usuario, revision Mapeo Agosto 2026
    "MICHEL GUADALUPE GUTIÉRREZ LÓPEZ": "LABORATORIOS GALINDO",  # confirmado por el usuario, revision Mapeo Agosto 2026
    "LABORATORIO QUIM CLINIC AZTECA": "LABORATORIO QUIMICO CLINICO AZTECA",  # customer code 37 -- abreviado en la OC (QUIM/CLINIC en vez de QUIMICO/CLINICO)
    "BLANCA LIZBETH ESPINAL PEREZ (CONTED INNOVACIONES)": "BLANCA LIZBETH ESPINAL PEREZ",  # customer code 278 -- confirmado por el usuario como ejemplo
}

# Casos donde el nombre de cliente que se extrajo del documento en realidad
# es "SEEGENE MEXICO..." (el vendedor, no el comprador) porque el documento
# lo menciona de forma mas prominente que al cliente real -- CUSTOMER_NAME_ALIASES
# no sirve aqui porque el texto extraido ("SEEGENE MEXICO...") es el mismo en
# ambos casos pero el cliente real es distinto segun la OC. Se identifica por
# PO No. (unico por documento) en vez de por nombre. El cliente real es quien
# de verdad mando la orden de compra -- eso es lo que debe salir en Cliente,
# nunca el nombre de Seegene Mexico.
CUSTOMER_OVERRIDE_BY_PO_NUMBER = {
    "OC006846": "LAPI",  # LAPI 6846 SEEGENE MEXICO SAPI DE CV OK.pdf -- confirmado por el usuario
    "OC006988": "LAPI",  # LAPI 6988 SEEGENE MEXICO SAPI DE CV.pdf -- mismo patron, mismo cliente (LAPI)
    "OC-AGOSTO26-001": "INTRAGEN",  # INTRAGEN - UNAM OC-AGOSTO26-001*.pdf -- confirmado por el usuario
}

# Codigos de producto tal como los escribe el cliente en su propia OC (o su
# propio SKU/SAP interno) -> codigo real de Product master list.xlsx. Se
# revisan en find_product() antes de intentar cualquier otra heuristica.
# Igual que CUSTOMER_NAME_ALIASES: agregar aqui cualquier variante nueva que
# se detecte en el futuro (ya confirmada por el usuario, no adivinada).
PRODUCT_CODE_ALIASES = {
    "8205949": "MG10211Z",
    "8206260": "GI10202Z",
    "8206323": "GI10184Z",
    "8206324": "GI10201Z",
    "8206325": "GI10183Z",
    "DT0836": "GI9703X",
    "DT0837": "GI10191X",
    "DT0838": "MG10210Z",
    "EX000019P": "EX00009P",
    "EX000019T": "EX00009T",
    "S61701": "SG1701",
    "TB72OOX": "TB7200X",
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
# por "/" (u otros separadores) y se busca cada parte contra "Code" (columna B
# de Product master list.xlsx) -- si hay match, ese codigo/descripcion del
# maestro reemplaza lo extraido del documento, y la columna "Category Product"
# de la salida se llena con el valor de "Cat Product" (columna I).
# ---------------------------------------------------------------------------
PM_SHEET = "Sheet1"
PM_HEADER_ROW = 1
PM_FIRST_DATA_ROW = 2
PM_COL_CODE = "Code"
PM_COL_DESCRIPTION = "Description"
PM_COL_CATEGORY = "Cat Product"

# ---------------------------------------------------------------------------
# Clasificaciones de cliente -> que columna de precio de distribuidor usar
# ---------------------------------------------------------------------------
DISTRIBUTOR_A_LABEL = "Distributor A"
DISTRIBUTOR_B_LABEL = "Distributor B"

# ---------------------------------------------------------------------------
# Modelo de Claude a usar para la extraccion de OC (vision + texto)
# ---------------------------------------------------------------------------
CLAUDE_MODEL = "claude-sonnet-5"
