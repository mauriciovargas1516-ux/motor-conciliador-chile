import sys
import re
import logging
import warnings
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from thefuzz import process, fuzz
from openpyxl.styles import Font, PatternFill

warnings.filterwarnings('ignore')

# ========================================================
# 0. CONFIGURACIÓN DEL SISTEMA DE AUDITORÍA (FORENSE)
# ========================================================
# Esto crea un archivo .log y también imprime en la consola
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("auditoria_conciliacion.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logging.info("="*75)
logging.info("INICIANDO MOTOR CONCILIADOR V4 (STOPWORDS + AUDITORÍA)")
logging.info("="*75)

# ========================================================
# 1. EXTRACCIÓN CON INTERFAZ GRÁFICA (Selector de Archivos)
# ========================================================
root = tk.Tk()
root.withdraw() 

# --- SELECCIÓN 1: EL BANCO ---
logging.info("Esperando que el usuario seleccione la cartola bancaria...")
ruta_banco = filedialog.askopenfilename(
    title="Paso 1: Sube la Cartola Bancaria (XLSX o CSV)",
    filetypes=[("Archivos Excel / CSV", "*.xlsx *.xls *.csv"), ("Todos los archivos", "*.*")]
)

if not ruta_banco:
    logging.error("Operación cancelada. No se seleccionó cartola del banco.")
    sys.exit()

# --- SELECCIÓN 2: EL ERP ---
logging.info("Esperando que el usuario seleccione las facturas del ERP...")
ruta_erp = filedialog.askopenfilename(
    title="Paso 2: Sube el Excel de Cuentas por Cobrar (XLSX o CSV)",
    filetypes=[("Archivos Excel / CSV", "*.xlsx *.xls *.csv"), ("Todos los archivos", "*.*")]
)

if not ruta_erp:
    logging.error("Operación cancelada. No se seleccionó el listado del ERP.")
    sys.exit()
# ====================================================================
# FUNCIONES AUXILIARES: DETECTOR UNIVERSAL Y EXTRACTOR DE RUT
# ====================================================================
def detectar_cabecera_excel(ruta_archivo):
    """Escanea las primeras 35 filas y detecta automáticamente la cabecera real."""
    df_sample = pd.read_excel(ruta_archivo, header=None, nrows=35)
    for idx, row in df_sample.iterrows():
        celdas = [str(c).strip().upper() for c in row if pd.notna(c)]
        if len(celdas) < 3:
            continue
        tiene_fecha = any('FECHA' in c for c in celdas)
        tiene_desc = any(any(w in c for w in ['DESCRIP', 'GLOSA', 'DETALLE', 'PROVEEDOR', 'CLIENTE']) for c in celdas)
        tiene_monto = any(any(w in c for w in ['ABONO', 'CARGO', 'MONTO', 'TOTAL', 'DEBIDA']) for c in celdas)
        if tiene_fecha and (tiene_desc or tiene_monto):
            return idx
    return 0

def extraer_rut_texto(texto):
    """Extrae RUTs chilenos desde glosas (ej: formato Santander '0209530732' o estándar)."""
    if not isinstance(texto, str):
        return 'nan'
    match = re.search(r'\b0?(\d{7,8})[-]?([0-9kK])\b', texto)
    if match:
        return f"{match.group(1)}{match.group(2).upper()}"
    return 'nan'

# ====================================================================
# LECTURA INTELIGENTE (BANCO Y ERP)
# ====================================================================
if ruta_banco.endswith(('.xlsx', '.xls')):
    idx_b = detectar_cabecera_excel(ruta_banco)
    df_banco = pd.read_excel(ruta_banco, skiprows=idx_b)
else:
    df_banco = pd.read_csv(ruta_banco, sep=None, engine='python')

if ruta_erp.endswith(('.xlsx', '.xls')):
    idx_e = detectar_cabecera_excel(ruta_erp)
    df_erp = pd.read_excel(ruta_erp, skiprows=idx_e)
else:
    df_erp = pd.read_csv(ruta_erp, sep=None, engine='python')

# ====================================================================
# HOMOLOGACIÓN Y LIMPIEZA DE COLUMNAS (ADAPTADOR UNIVERSAL)
# ====================================================================
# 1. Homologación Banco
traductor_banco = {
    'Fecha': 'FECHA', 'FECHA': 'FECHA',
    'N° Operación': 'ID_TX', 'N° DOCUMENTO': 'ID_TX',
    'Descripción': 'GLOSA', 'DESCRIPCIÓN': 'GLOSA',
    'Abonos': 'MONTO', 'DEPOSITOS Y OTROS ABONOS': 'MONTO',
    'RUT': 'RUT_PAGADOR', 'RUT Pagador': 'RUT_PAGADOR'
}
df_banco.rename(columns=traductor_banco, inplace=True)

# Asegurar ID_TX único si viene vacío
if 'ID_TX' not in df_banco.columns or df_banco['ID_TX'].isna().all():
    df_banco['ID_TX'] = [f"TX_{i+1}" for i in range(len(df_banco))]

# Filtrar solo abonos válidos y eliminar mensajes de pie de página
df_banco = df_banco.dropna(subset=['FECHA']).copy()
if 'MONTO' in df_banco.columns:
    df_banco['MONTO'] = pd.to_numeric(df_banco['MONTO'].astype(str).str.replace(r'[\$\.]', '', regex=True), errors='coerce')
    df_banco = df_banco[df_banco['MONTO'] > 0].copy()

# ========================================================
# 2. HOMOLOGACIÓN ERP (Buscador Estadístico Difuso)
# ========================================================
def encontrar_columna(df, palabras_clave):
    """Busca columnas priorizando las palabras clave más exactas primero."""
    for palabra in palabras_clave:
        for col in df.columns.tolist():
            if palabra.lower() in str(col).lower().replace('\n', ''):
                return col
    return None

# Definimos el orden de prioridad de búsqueda (de más exacto a más general)
col_rut = encontrar_columna(df_erp, ['rut', 'identificador', 'provider'])
col_nombre = encontrar_columna(df_erp, ['nombre', 'razón', 'razon', 'cliente', 'proveedor', 'social'])
col_id = encontrar_columna(df_erp, ['factura', 'folio', 'número', 'numero', 'doc'])
# Fíjate que ponemos 'debida' y 'saldo' ANTES que 'total' para evitar la columna de total facturado
col_monto = encontrar_columna(df_erp, ['debida', 'saldo', 'deuda', 'monto', 'total'])

# Renombramos automáticamente lo que el algoritmo haya encontrado
if col_rut: df_erp.rename(columns={col_rut: 'RUT_CLIENTE'}, inplace=True)
if col_nombre: df_erp.rename(columns={col_nombre: 'NOMBRE_CLIENTE'}, inplace=True)
if col_id: df_erp.rename(columns={col_id: 'ID_FACTURA'}, inplace=True)
if col_monto: df_erp.rename(columns={col_monto: 'MONTO_FACTURA'}, inplace=True)

# Limpiamos las filas que no tengan un ID de factura válido
if 'ID_FACTURA' in df_erp.columns:
    df_erp = df_erp.dropna(subset=['ID_FACTURA']).copy()

if 'MONTO_FACTURA' in df_erp.columns:
    df_erp['MONTO_FACTURA'] = pd.to_numeric(df_erp['MONTO_FACTURA'].astype(str).str.replace(r'[\$\.]', '', regex=True), errors='coerce')

# ====================================================================
# LIMPIEZA FORENSE DE RUTS Y GLOSAS
# ====================================================================
# Limpieza de RUT en ERP
if 'RUT_CLIENTE' in df_erp.columns and not df_erp['RUT_CLIENTE'].isna().all():
    df_erp['RUT_LIMPIO'] = df_erp['RUT_CLIENTE'].astype(str).str.replace(r'[\.\-]', '', regex=True).str.strip()
else:
    df_erp['RUT_LIMPIO'] = 'nan'

# Limpieza / Extracción de RUT en Banco
if 'RUT_PAGADOR' in df_banco.columns and not df_banco['RUT_PAGADOR'].isna().all():
    df_banco['RUT_LIMPIO'] = df_banco['RUT_PAGADOR'].astype(str).str.replace(r'[\.\-]', '', regex=True).str.strip()
else:
    # Extracción automática desde la glosa (Santander, Itaú, etc.)
    df_banco['RUT_LIMPIO'] = df_banco['GLOSA'].apply(extraer_rut_texto)

# Limpieza de Stopwords en la Glosa
stopwords = r'\b(TRF|DEP|PAGO|FACTURA|FACTURAS|MULTIPLE|CANCELACION|ABONO|TERCEROS|DE|TRANSF|COMPRA|NACIONAL|MD)\b'
df_banco['GLOSA_LIMPIA'] = df_banco['GLOSA'].astype(str).str.upper().replace(stopwords, '', regex=True)
df_banco['GLOSA_LIMPIA'] = df_banco['GLOSA_LIMPIA'].replace(r'\s+', ' ', regex=True).str.strip()
tx_conciliadas = []
facturas_conciliadas = []
# ========================================================
# 2. EL CAJERO (AGRUPACIÓN MUCHOS A UNO)
# ========================================================
logging.info("[FASE 1] BUSCANDO PAGOS CONSOLIDADOS (Varios a Uno)...")
erp_agrupado = df_erp.groupby('RUT_LIMPIO')['MONTO_FACTURA'].sum().reset_index()

match_agrupado = pd.merge(
    df_banco[df_banco['RUT_LIMPIO'] != 'nan'], erp_agrupado, 
    left_on=['RUT_LIMPIO', 'MONTO'], right_on=['RUT_LIMPIO', 'MONTO_FACTURA'], how='inner'
)

for _, row in match_agrupado.iterrows():
    rut_match = row['RUT_LIMPIO']
    facturas_del_rut = df_erp[df_erp['RUT_LIMPIO'] == rut_match]
    
    # Registro forense
    logging.info(f"[CONSOLIDADO] ID_TX: {row['ID_TX']} | Depósito de ${row['MONTO']} liquidó {len(facturas_del_rut)} facturas de {facturas_del_rut.iloc[0]['NOMBRE_CLIENTE']}")
    
    tx_conciliadas.append(row['ID_TX'])
    facturas_conciliadas.extend(facturas_del_rut['ID_FACTURA'].tolist())

df_banco = df_banco[~df_banco['ID_TX'].isin(tx_conciliadas)]
df_erp = df_erp[~df_erp['ID_FACTURA'].isin(facturas_conciliadas)]

# ========================================================
# 3. EL FRANCOTIRADOR (UNO A UNO)
# ========================================================
logging.info("[FASE 2] BUSCANDO MATCH EXACTO (Uno a Uno)...")
match_exacto = pd.merge(
    df_banco[df_banco['RUT_LIMPIO'] != 'nan'], df_erp, 
    left_on=['RUT_LIMPIO', 'MONTO'], right_on=['RUT_LIMPIO', 'MONTO_FACTURA'], how='inner'
)

for _, row in match_exacto.iterrows():
    logging.info(f"[EXACTO] ID_TX: {row['ID_TX']} | Depósito de ${row['MONTO']} pagó la factura {row['ID_FACTURA']}")
    tx_conciliadas.append(row['ID_TX'])
    facturas_conciliadas.append(row['ID_FACTURA'])

df_banco = df_banco[~df_banco['ID_TX'].isin(tx_conciliadas)]
df_erp = df_erp[~df_erp['ID_FACTURA'].isin(facturas_conciliadas)]

# ========================================================
# 4. EL DETECTIVE (FUZZY MATCH POTENCIADO)
# ========================================================
logging.info("[FASE 3] BUSCANDO POR TEXTO DIFUSO CON GLOSAS LIMPIAS...")
for idx_banco, fila_banco in df_banco.iterrows():
    candidatos = df_erp[df_erp['MONTO_FACTURA'] == fila_banco['MONTO']]
    
    if not candidatos.empty:
        # Ahora usamos GLOSA_LIMPIA en lugar de la GLOSA sucia original
        mejor_match, puntaje = process.extractOne(fila_banco['GLOSA_LIMPIA'], candidatos['NOMBRE_CLIENTE'].tolist(), scorer=fuzz.WRatio)
        if puntaje >= 75: # Subimos la exigencia al 75% porque la data está más limpia
            factura = candidatos[candidatos['NOMBRE_CLIENTE'] == mejor_match].iloc[0]['ID_FACTURA']
            logging.info(f"[FUZZY {puntaje}%] ID_TX: {fila_banco['ID_TX']} | '{fila_banco['GLOSA_LIMPIA']}' emparejado con '{mejor_match}' (Factura {factura})")
            
            tx_conciliadas.append(fila_banco['ID_TX'])
            facturas_conciliadas.append(factura)

# ========================================================
# 5. REPORTE GERENCIAL (Excepciones)
# ========================================================
df_banco = df_banco[~df_banco['ID_TX'].isin(tx_conciliadas)]
df_erp = df_erp[~df_erp['ID_FACTURA'].isin(facturas_conciliadas)]

logging.warning(f"[FASE 4] ACTUALIZANDO REPORTE EXCEL (Hay {len(df_banco)} ingresos sin origen y {len(df_erp)} facturas impagas).")

with pd.ExcelWriter('Reporte_Excepciones_Gerencia.xlsx', engine='openpyxl') as writer:
    df_banco[['FECHA', 'GLOSA', 'MONTO']].rename(columns={'GLOSA': 'Descripción Original', 'MONTO': 'Sobrante ($)'}).to_excel(writer, sheet_name='Sobrantes', index=False)
    df_erp[['ID_FACTURA', 'NOMBRE_CLIENTE', 'MONTO_FACTURA']].rename(columns={'NOMBRE_CLIENTE': 'Cliente', 'MONTO_FACTURA': 'Deuda ($)'}).to_excel(writer, sheet_name='Impagas', index=False)

logging.info("PROCESO TERMINADO CON ÉXITO.")