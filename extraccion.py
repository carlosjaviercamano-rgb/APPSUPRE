import pandas as pd
import openpyxl
from datetime import datetime

# ─── Hojas por grupo ───────────────────────────────────────────────────────
HOJAS_BANCARIOS = [
    "BANCOLOMBIA SUPRECREDITO",
    "DAVIVIENDA SUPRECREDITO",
    "DAVIVIENDA SUPRECREDITO ",   # con espacio al final
]

HOJAS_RECAUDOS = [
    "OCCIDENTE SUPRECREDITO 2026",
    "RECORD"
]

# ─── Mapeo de columnas del libro fuente ────────────────────────────────────
# El libro tiene: A=FECHAINGRESO, B=ENTIDAD, C=NOMBRE, D=CEDULA,
#                 E=DOCUMENTODEAPROBACION, F=TIPODETRANSACIÓN,
#                 G=REFERENCIA, H=VALOR, I=FRA, J=RECIBOS/RECIBO

COL_FECHA    = 0   # A
COL_ENTIDAD  = 1   # B
COL_NOMBRE   = 2   # C
COL_CEDULA   = 3   # D
COL_DOC      = 4   # E
COL_TIPO     = 5   # F
COL_REF      = 6   # G
COL_VALOR    = 7   # H
COL_FRA      = 8   # I
COL_RECIBO   = 9   # J


def leer_hoja(archivo, nombre_hoja):
    """Lee una hoja del libro como DataFrame sin fórmulas."""
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    if nombre_hoja not in wb.sheetnames:
        wb.close()
        return None

    ws = wb[nombre_hoja]
    datos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(v is not None for v in row):
            datos.append(row)
    wb.close()

    if not datos:
        return pd.DataFrame()

    # Detectar número de columnas de la hoja
    max_cols = max(len(r) for r in datos)
    columnas_hoja = [f"COL_{i}" for i in range(max_cols)]

    df = pd.DataFrame(datos, columns=columnas_hoja)
    return df


def extraer_pagos_bancarios(archivo, archivo_corresponsal):
    """
    Extrae datos de las hojas de PAGOS BANCARIOS.
    Filtro: columna D (CEDULA) con valor Y columna J (RECIBOS) vacía.
    Copia columnas A, B, D, H, I → FECHA, ENTIDAD, CEDULA, VALOR, NUM_FACTURA
    De BANCOLOMBIA además copia columna F → T_TRANSACCION
    Proceso CORRESPONSAL: filtra solo REINCIDENTES si T_TRANSACCION = CONSIGNACION CORRESPONSAL CB
    """
    # Cargar cédulas corresponsales
    cedulas_corresponsal = cargar_corresponsal(archivo_corresponsal)

    frames = []

    for nombre_hoja in HOJAS_BANCARIOS:
        df = leer_hoja(archivo, nombre_hoja)
        if df is None or df.empty:
            continue

        # Renombrar columnas relevantes
        rename = {}
        if df.shape[1] > COL_FECHA:    rename[f"COL_{COL_FECHA}"]   = "FECHA"
        if df.shape[1] > COL_ENTIDAD:  rename[f"COL_{COL_ENTIDAD}"] = "ENTIDAD"
        if df.shape[1] > COL_CEDULA:   rename[f"COL_{COL_CEDULA}"]  = "CEDULA"
        if df.shape[1] > COL_TIPO:     rename[f"COL_{COL_TIPO}"]    = "T_TRANSACCION_SRC"
        if df.shape[1] > COL_VALOR:    rename[f"COL_{COL_VALOR}"]   = "VALOR"
        if df.shape[1] > COL_FRA:      rename[f"COL_{COL_FRA}"]     = "NUM_FACTURA"
        if df.shape[1] > COL_RECIBO:   rename[f"COL_{COL_RECIBO}"]  = "RECIBO_SRC"
        df = df.rename(columns=rename)

        # ── Filtro rosado: cedula con valor Y recibo vacío ──────────────
        tiene_cedula  = df["CEDULA"].notna() & (df["CEDULA"].astype(str).str.strip() != "")
        sin_recibo    = df["RECIBO_SRC"].isna() | (df["RECIBO_SRC"].astype(str).str.strip() == "")
        df = df[tiene_cedula & sin_recibo].copy()

        if df.empty:
            continue

        # ── Construir fila destino (orden AREA DE BANCO) ──────────────
        df_out = pd.DataFrame()
        df_out["FECHA"]           = df["FECHA"]   if "FECHA"   in df.columns else None
        df_out["ENTIDAD"]         = df["ENTIDAD"] if "ENTIDAD" in df.columns else nombre_hoja
        df_out["CEDULA"]          = df["CEDULA"]
        df_out["VALOR"]           = df["VALOR"]   if "VALOR"   in df.columns else None
        df_out["FRA"]             = df["NUM_FACTURA"] if "NUM_FACTURA" in df.columns else None
        df_out["RECIBOS"]         = None
        df_out["FECHA_DOCUMENTO"] = None
        df_out["T_TRANSACCION"]   = None   # interno, no visible

        # T_TRANSACCION: solo de BANCOLOMBIA (col F), resto vacío
        if "BANCOLOMBIA" in nombre_hoja.upper() and "T_TRANSACCION_SRC" in df.columns:
            df_out["T_TRANSACCION"] = df["T_TRANSACCION_SRC"]

        # REINCIDENTES_CB es la versión visual de T_TRANSACCION
        df_out["REINCIDENTES_CB"] = None
        df_out["COMPENSACION"]    = None

        # Alias interno para compatibilidad con alistar.py
        df_out["NUM_FACTURA"]     = df_out["FRA"]

        frames.append(df_out)

    if not frames:
        raise ValueError("No se encontraron datos con los filtros aplicados en las hojas de Pagos Bancarios.")

    df_final = pd.concat(frames, ignore_index=True)

    # ── Proceso CORRESPONSAL ────────────────────────────────────────────
    # 1. Limpiar T_TRANSACCION: dejar solo "CONSIGNACION CORRESPONSAL CB"
    mask_consig = df_final["T_TRANSACCION"].astype(str).str.upper().str.strip() == "CONSIGNACION CORRESPONSAL CB"
    df_final.loc[~mask_consig, "T_TRANSACCION"] = None

    # 2. Identificar PRIMERA VEZ: tiene CONSIGNACION pero cédula NO está en CORRESPONSAL
    mask_tiene_consig = df_final["T_TRANSACCION"].notna()
    cedulas_str = df_final["CEDULA"].astype(str).str.strip()
    mask_primera_vez = mask_tiene_consig & (~cedulas_str.isin(cedulas_corresponsal))

    # 3. Borrar T_TRANSACCION de los PRIMERA VEZ (quedan en blanco)
    df_final.loc[mask_primera_vez, "T_TRANSACCION"] = None

    # 4. Reflejar en REINCIDENTES_CB (columna visual)
    df_final["REINCIDENTES_CB"] = df_final["T_TRANSACCION"]

    # Calcular fecha documento = fecha más reciente de col FECHA
    fecha_doc = None
    fechas = pd.to_datetime(df_final["FECHA"], errors="coerce").dropna()
    if not fechas.empty:
        fecha_doc = fechas.max()
    df_final["FECHA_DOCUMENTO"] = fecha_doc

    n_total       = len(df_final)
    n_corresponsal = mask_tiene_consig.sum()
    n_primera_vez  = mask_primera_vez.sum()
    n_reincidente  = n_corresponsal - n_primera_vez

    resumen = (f"{n_total} registros extraídos. "
               f"Corresponsal: {n_corresponsal} ({n_reincidente} reincidentes, "
               f"{n_primera_vez} primera vez eliminados).")

    return df_final, resumen


def extraer_pagos_recaudos(archivo, fechas_filtro):
    """
    Extrae datos de las hojas de PAGOS POR RECAUDOS.
    Filtro: columna A (FECHAINGRESO) en lista de fechas seleccionadas.
    FECHA_DOCUMENTO = la misma fecha de cada registro (no la más reciente).
    """
    if not fechas_filtro:
        raise ValueError("Debes seleccionar al menos una fecha para los Pagos por Recaudos.")

    # Convertir fechas a timestamps normalizados
    fechas_dt = [pd.Timestamp(f).normalize() for f in fechas_filtro]
    frames = []

    for nombre_hoja in HOJAS_RECAUDOS:
        df = leer_hoja(archivo, nombre_hoja)
        if df is None or df.empty:
            continue

        rename = {}
        if df.shape[1] > COL_FECHA:   rename[f"COL_{COL_FECHA}"]   = "FECHA"
        if df.shape[1] > COL_ENTIDAD: rename[f"COL_{COL_ENTIDAD}"] = "ENTIDAD"
        if df.shape[1] > COL_CEDULA:  rename[f"COL_{COL_CEDULA}"]  = "CEDULA"
        if df.shape[1] > COL_VALOR:   rename[f"COL_{COL_VALOR}"]   = "VALOR"
        if df.shape[1] > COL_FRA:     rename[f"COL_{COL_FRA}"]     = "NUM_FACTURA"
        if df.shape[1] > COL_RECIBO:  rename[f"COL_{COL_RECIBO}"]  = "RECIBO_SRC"
        df = df.rename(columns=rename)

        # ── Filtro por múltiples fechas ─────────────────────────────────
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df = df[df["FECHA"].dt.normalize().isin(fechas_dt)].copy()

        if df.empty:
            continue

        df_out = pd.DataFrame()
        df_out["FECHA"]           = df["FECHA"]
        df_out["ENTIDAD"]         = df["ENTIDAD"] if "ENTIDAD" in df.columns else nombre_hoja
        df_out["CEDULA"]          = df["CEDULA"]  if "CEDULA"  in df.columns else None
        df_out["VALOR"]           = df["VALOR"]   if "VALOR"   in df.columns else None
        df_out["FRA"]             = df["NUM_FACTURA"] if "NUM_FACTURA" in df.columns else None
        df_out["RECIBOS"]         = None
        df_out["T_TRANSACCION"]   = None
        df_out["REINCIDENTES_CB"] = None
        df_out["COMPENSACION"]    = None

        # FECHA_DOCUMENTO = la misma fecha de cada registro (no la más reciente)
        df_out["FECHA_DOCUMENTO"] = df["FECHA"]

        # Alias interno
        df_out["NUM_FACTURA"] = df_out["FRA"]

        frames.append(df_out)

    if not frames:
        fechas_str = ", ".join(f.strftime("%d/%m/%Y") for f in fechas_filtro)
        raise ValueError(f"No se encontraron registros para las fechas {fechas_str} en las hojas de Recaudos.")

    df_final = pd.concat(frames, ignore_index=True)
    fechas_str = ", ".join(f.strftime("%d/%m/%Y") for f in fechas_filtro)
    resumen = f"{len(df_final)} registros extraídos para las fechas: {fechas_str}."
    return df_final, resumen


def cargar_corresponsal(archivo):
    """Carga las cédulas del archivo CORRESPONSAL como un set de strings."""
    try:
        df = pd.read_excel(archivo, header=None)
        cedulas = df.iloc[:, 0].dropna().astype(str).str.strip()
        return set(cedulas)
    except Exception:
        return set()
