"""
Lógica del módulo Obligaciones Financieras (sin dependencias de Streamlit,
para poder probarla por separado).

Flujo mensual:
  1. Se carga el reporte de AWS (csv/xlsx) con las obligaciones del mes.
  2. Se cruza contra el libro histórico por EMPRESA + FACTURA + IDENTIFICACIÓN
     (las tres juntas, porque un mismo número de factura puede repetirse entre
     empresas o proveedores):
        - existente  -> se actualizan los valores del mes
        - nueva      -> se agrega al final, previa clasificación del tipo
        - no viene   -> se marca "No activo" (crédito cancelado)
  3. Con el histórico ya actualizado se arma el libro de control compartido
     con tesorería:
        - PAGOS_POR_CAJA   : tipo Particular, UNA fila por identificación de
                             proveedor con la SUMA de sus cuotas del mes.
        - PAGOS_POR_BANCOS : tipo Bancario, una fila por obligación.
        - Las de tipo Gerencia solo viven en el histórico (manejo interno).
     Cada mes agrega dos columnas nuevas (valor_cuota_<mes>_<año> y
     estado_de_la_cuota_<mes>_<año>) y nunca pisa lo que ya escribieron
     tesorería / quienes legalizan.
  4. Las filas que dejan de estar activas se mueven a la hoja CANCELADAS,
     conservando su historial de meses anteriores.
"""
import io
import os
import re
import shutil
import unicodedata
from copy import copy
from datetime import datetime

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── Constantes ────────────────────────────────────────────────────────────
HOJA_HISTORICO  = "historico"
HOJA_CAJA       = "PAGOS_POR_CAJA"
HOJA_BANCOS     = "PAGOS_POR_BANCOS"
HOJA_CANCELADAS = "CANCELADAS"

TIPOS_OBLIGACION = ["Bancario", "Particular", "Gerencia"]
ESTADO_ACTIVO    = "Activo"
ESTADO_INACTIVO  = "No activo"

ESTADO_CUOTA_PENDIENTE = "PENDIENTE"
ESTADO_CUOTA_PAGADO    = "PAGADO"

MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Columnas del reporte de AWS (en el mismo orden del histórico)
COLUMNAS_REPORTE = [
    "empresa", "factura", "tipo_documento", "identificacion_proveedor",
    "nombre_proveedor", "monto_desembolsado", "tasa", "modalidad_tasa",
    "base_rate_type", "spread", "plazo", "valor_cuota", "capital", "interes",
    "iva", "seguro", "fondo_garantias", "saldo_real", "sistema_amortizacion",
    "fecha_pago", "estado_pago_mes",
]
COLUMNAS_NUMERICAS = {
    "monto_desembolsado", "tasa", "spread", "plazo", "valor_cuota", "capital",
    "interes", "iva", "seguro", "fondo_garantias", "saldo_real",
}
COLUMNAS_FECHA = {"fecha_pago"}
COLUMNAS_CLAVE = ["empresa", "factura", "identificacion_proveedor"]
COLUMNAS_REQUERIDAS_REPORTE = COLUMNAS_CLAVE + ["nombre_proveedor", "valor_cuota"]

FORMATO_MONEDA = '_-"$"\\ * #,##0.00_-;\\-"$"\\ * #,##0.00_-;_-"$"\\ * "-"??_-;_-@_-'

# Definición de las hojas del libro de control.
#  - "clave": columnas que identifican una fila de la hoja.
#  - "encabezados": columnas fijas; solo se escriben si la hoja está vacía.
#  - "metricas": columnas que se agregan CADA MES (con sufijo _<mes>_<año>).
METRICA_ESTADO = "estado_de_la_cuota"
ETIQUETA_NUEVO = "NUEVO"
ETIQUETA_EXISTENTE = "EXISTENTE"

SPEC_CAJA = {
    "hoja": HOJA_CAJA,
    "clave": ["id_proveedor"],
    "encabezados": ["id_proveedor", "nombre_proveedor", "Antigüedad_Crédito"],
    "metricas": ["valor_cuota", METRICA_ESTADO],
}
SPEC_BANCOS = {
    "hoja": HOJA_BANCOS,
    "clave": ["empresa", "factura", "identificacion_proveedor"],
    "encabezados": ["empresa", "factura", "tipo_documento", "identificacion_proveedor",
                    "nombre_proveedor", "Antigüedad_Crédito"],
    "metricas": ["valor_cuota", "capital", "interes", METRICA_ESTADO],
    # en bancos se agrupan y ocultan las columnas de los meses anteriores (se expanden con "+")
    "ocultar_meses_anteriores": True,
}


DIR_RESPALDOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "respaldos_obligaciones")
MAX_RESPALDOS_POR_ARCHIVO = 12


# ══════════════════════════════════════════════════════════════════════════
# NORMALIZACIÓN (claves de cruce)
# ══════════════════════════════════════════════════════════════════════════
def _sin_acentos(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")


def _texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def _arreglar_mojibake(s):
    """'AndrÃ©s' -> 'Andrés' (texto UTF-8 que alguien leyó como latin-1)."""
    if not isinstance(s, str) or ("Ã" not in s and "Â" not in s):
        return s
    for enc in ("latin-1", "cp1252"):
        try:
            return s.encode(enc).decode("utf-8")
        except Exception:
            continue
    return s


def _limpiar_nombre(s):
    return " ".join(_arreglar_mojibake(_texto(s)).split())


def _norm_header(v):
    return _sin_acentos(_texto(v).lower())


def _quitar_ceros_decimales(s):
    # Excel/pandas pueden dejar un número como '123.0'
    if re.fullmatch(r"-?\d+\.0", s):
        return s[:-2]
    return s


def _norm_empresa(v):
    return re.sub(r"\s+", "", _sin_acentos(_texto(v))).upper()


def _limpiar_factura(v):
    """Quita el envoltorio ="123" con el que AWS protege la factura en el csv."""
    s = _texto(v)
    m = re.fullmatch(r'=\s*"(.*)"', s)
    if m:
        s = m.group(1)
    return _quitar_ceros_decimales(s.strip())


def _norm_factura(v):
    return _limpiar_factura(v).upper()


def _norm_id(v):
    s = _quitar_ceros_decimales(_texto(v))
    if s and re.fullmatch(r"[\d.,\s]+", s):
        s = re.sub(r"[.,\s]", "", s)
    return s.upper()


def _norm_txt(v):
    return _texto(v).upper()


_NORMALIZADORES = {
    "empresa": _norm_empresa,
    "factura": _norm_factura,
    "id_proveedor": _norm_id,
    "identificacion_proveedor": _norm_id,
}


def _norm_campo(campo, valor):
    return _NORMALIZADORES.get(campo, _norm_txt)(valor)


def _clave_obligacion(empresa, factura, identificacion):
    return (_norm_empresa(empresa), _norm_factura(factura), _norm_id(identificacion))


# ══════════════════════════════════════════════════════════════════════════
# CONVERSIÓN DE VALORES
# ══════════════════════════════════════════════════════════════════════════
def _a_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None if (isinstance(v, float) and pd.isna(v)) else float(v)
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _a_fecha(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v
    if hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
        return datetime(v.year, v.month, v.day)
    s = _texto(v)
    if not s:
        return None
    iso = re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s)
    d = pd.to_datetime(s, errors="coerce", dayfirst=not iso)
    if pd.isna(d):
        return None
    return d.to_pydatetime().replace(hour=0, minute=0, second=0, microsecond=0)


def _id_para_celda(v):
    s = _norm_id(v)
    return int(s) if s.isdigit() else (_texto(v) or None)


def _valor_para_celda(columna, v):
    """Convierte un valor del reporte al tipo que lleva la celda del histórico."""
    if columna in COLUMNAS_FECHA:
        return _a_fecha(v)
    if columna in COLUMNAS_NUMERICAS:
        f = _a_float(v)
        if f is None:
            return None
        if columna == "plazo" and float(f).is_integer():
            return int(f)
        return f
    if columna == "identificacion_proveedor":
        return _id_para_celda(v)
    if columna == "factura":
        return _limpiar_factura(v) or None
    t = _arreglar_mojibake(_texto(v))
    return t if t != "" else None


# ══════════════════════════════════════════════════════════════════════════
# LECTURA DEL REPORTE
# ══════════════════════════════════════════════════════════════════════════
def leer_reporte(archivo):
    """
    Lee el reporte de AWS (csv o xlsx) y devuelve un DataFrame con las
    columnas ya convertidas y una columna '_clave' (empresa, factura, id).
    """
    nombre = str(getattr(archivo, "name", archivo)).lower()
    if hasattr(archivo, "seek"):
        archivo.seek(0)

    if nombre.endswith((".xlsx", ".xls")):
        df = pd.read_excel(archivo, dtype=str)
    else:
        raw = archivo.read() if hasattr(archivo, "read") else open(archivo, "rb").read()
        texto = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                texto = raw.decode(enc)
                break
            except Exception:
                continue
        df = pd.read_csv(io.StringIO(texto), dtype=str, keep_default_na=False)

    df.columns = [_texto(c).lower() for c in df.columns]
    faltantes = [c for c in COLUMNAS_REQUERIDAS_REPORTE if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en el reporte: {', '.join(faltantes)}")

    filas = []
    for _, r in df.iterrows():
        if not _texto(r["empresa"]) or not _texto(r["factura"]):
            continue
        fila = {c: _valor_para_celda(c, r[c]) for c in COLUMNAS_REPORTE if c in df.columns}
        fila["nombre_proveedor"] = _limpiar_nombre(r["nombre_proveedor"])
        fila["_clave"] = _clave_obligacion(r["empresa"], r["factura"], r["identificacion_proveedor"])
        filas.append(fila)

    if not filas:
        raise ValueError("El reporte no tiene filas con empresa y factura.")

    out = pd.DataFrame(filas)
    dup = out[out.duplicated("_clave", keep=False)]
    if not dup.empty:
        ej = ", ".join(f"{r.empresa}/{r.factura}/{r.identificacion_proveedor}"
                       for r in dup.drop_duplicates("_clave").head(5).itertuples())
        raise ValueError(
            f"El reporte trae {dup['_clave'].nunique()} obligación(es) repetida(s) con la misma "
            f"empresa + factura + identificación (ej.: {ej}). Revísalo antes de continuar."
        )
    return out


# ══════════════════════════════════════════════════════════════════════════
# HISTÓRICO
# ══════════════════════════════════════════════════════════════════════════
def _headers(ws):
    """Encabezados de la fila 1: {nombre normalizado: número de columna}."""
    h = {}
    for c in ws[1]:
        if _texto(c.value):
            h.setdefault(_norm_header(c.value), c.column)
    return h


def _ultima_col(ws):
    cols = [c.column for c in ws[1] if _texto(c.value)]
    return max(cols) if cols else 0


def _es_activo(valor_estado):
    return _norm_header(valor_estado) in ("activo", "")  # vacío = se asume activo


def _leer_filas_historico(ws):
    hdr = _headers(ws)
    faltan = [c for c in COLUMNAS_CLAVE + ["tipo_de_obligacion", "estado_obligacion"] if c not in hdr]
    if faltan:
        raise ValueError(f"En la hoja '{ws.title}' del histórico faltan las columnas: {', '.join(faltan)}")
    filas = []
    for r in range(2, ws.max_row + 1):
        emp = ws.cell(r, hdr["empresa"]).value
        fac = ws.cell(r, hdr["factura"]).value
        if not _texto(emp) or not _texto(fac):
            continue
        valores = {k: ws.cell(r, c).value for k, c in hdr.items()}
        filas.append({
            "_fila": r,
            "_clave": _clave_obligacion(emp, fac, valores["identificacion_proveedor"]),
            "valores": valores,
        })
    return hdr, filas


def _copiar_estilo(origen, destino):
    if origen.has_style:
        destino.font = copy(origen.font)
        destino.fill = copy(origen.fill)
        destino.border = copy(origen.border)
        destino.alignment = copy(origen.alignment)
        destino.number_format = origen.number_format
        destino.protection = copy(origen.protection)


def _resumen_fila(f):
    v = f["valores"]
    return {
        "empresa": _texto(v.get("empresa")),
        "factura": _limpiar_factura(v.get("factura")),
        "identificacion_proveedor": _norm_id(v.get("identificacion_proveedor")),
        "nombre_proveedor": _limpiar_nombre(v.get("nombre_proveedor")),
        "tipo_de_obligacion": _texto(v.get("tipo_de_obligacion")),
        "valor_cuota": _a_float(v.get("valor_cuota")) or 0.0,
    }


def analizar(df_rep, ruta_historico):
    """
    Cruza el reporte contra el histórico SIN modificar nada. Devuelve qué
    obligaciones son existentes, nuevas, canceladas (activas que ya no
    vienen) y reactivadas (marcadas 'No activo' que volvieron a venir).
    """
    wb = openpyxl.load_workbook(ruta_historico, data_only=True)
    if HOJA_HISTORICO not in wb.sheetnames:
        raise ValueError(f"El libro histórico no tiene una hoja llamada '{HOJA_HISTORICO}'.")
    _, filas = _leer_filas_historico(wb[HOJA_HISTORICO])
    por_clave = {f["_clave"]: f for f in filas}
    claves_rep = set(df_rep["_clave"])

    nuevas = df_rep[~df_rep["_clave"].isin(por_clave)].copy()
    canceladas = [f for f in filas
                  if f["_clave"] not in claves_rep and _es_activo(f["valores"].get("estado_obligacion"))]
    reactivadas = [f for f in filas
                   if f["_clave"] in claves_rep and not _es_activo(f["valores"].get("estado_obligacion"))]
    n_activas = sum(1 for f in filas if _es_activo(f["valores"].get("estado_obligacion")))

    return {
        "n_existentes": int(df_rep["_clave"].isin(por_clave).sum()),
        "n_activas_historico": n_activas,
        "nuevas": nuevas,
        "canceladas": [_resumen_fila(f) for f in canceladas],
        "reactivadas": [_resumen_fila(f) for f in reactivadas],
    }


def _aplicar_historico(wb, df_rep, tipos_nuevas):
    """Aplica en memoria el cruce sobre la hoja histórico. Devuelve (resumen, filas)."""
    ws = wb[HOJA_HISTORICO]
    hdr, filas = _leer_filas_historico(ws)
    por_clave = {f["_clave"]: f for f in filas}
    claves_rep = set(df_rep["_clave"])
    ultima = max([f["_fila"] for f in filas], default=1)

    sin_tipo = [c for c in df_rep["_clave"] if c not in por_clave and not tipos_nuevas.get(c)]
    if sin_tipo:
        raise ValueError(f"Faltan por clasificar {len(sin_tipo)} obligación(es) nueva(s).")

    # columnas del reporte que existen en el histórico (menos la clave, que no se toca)
    cols_actualizar = [c for c in COLUMNAS_REPORTE if c in hdr and c not in COLUMNAS_CLAVE]
    cols_escribir_nueva = [c for c in COLUMNAS_REPORTE if c in hdr]

    res = {"actualizadas": 0, "nuevas": [], "canceladas": [], "reactivadas": [],
           "claves_nuevas": []}

    for _, r in df_rep.iterrows():
        clave = r["_clave"]
        if clave in por_clave:
            fila = por_clave[clave]["_fila"]
            for c in cols_actualizar:
                ws.cell(fila, hdr[c]).value = r[c] if c in r and not _es_nulo(r[c]) else None
            est = ws.cell(fila, hdr["estado_obligacion"])
            if not _es_activo(est.value):
                est.value = ESTADO_ACTIVO
                res["reactivadas"].append(_resumen_fila(por_clave[clave]))
            res["actualizadas"] += 1
        else:
            ultima += 1
            for c in cols_escribir_nueva:
                celda = ws.cell(ultima, hdr[c])
                celda.value = r[c] if c in r and not _es_nulo(r[c]) else None
            ws.cell(ultima, hdr["tipo_de_obligacion"]).value = tipos_nuevas[clave]
            ws.cell(ultima, hdr["estado_obligacion"]).value = ESTADO_ACTIVO
            # estilo (formatos de número/fecha) igual al de la fila anterior
            for c in range(1, _ultima_col(ws) + 1):
                if ultima - 1 >= 2:
                    _copiar_estilo(ws.cell(ultima - 1, c), ws.cell(ultima, c))
            res["claves_nuevas"].append(clave)
            res["nuevas"].append({
                "empresa": _texto(r["empresa"]), "factura": _texto(r["factura"]),
                "identificacion_proveedor": _norm_id(r["identificacion_proveedor"]),
                "nombre_proveedor": _texto(r["nombre_proveedor"]),
                "tipo_de_obligacion": tipos_nuevas[clave],
                "valor_cuota": _a_float(r["valor_cuota"]) or 0.0,
            })

    for fila in filas:
        if fila["_clave"] not in claves_rep and _es_activo(fila["valores"].get("estado_obligacion")):
            ws.cell(fila["_fila"], hdr["estado_obligacion"]).value = ESTADO_INACTIVO
            res["canceladas"].append(_resumen_fila(fila))

    if ws.auto_filter and ws.auto_filter.ref:
        ws.auto_filter.ref = f"A1:{get_column_letter(_ultima_col(ws))}{max(ultima, 2)}"

    _, filas_proj = _leer_filas_historico(ws)
    return res, filas_proj


def _es_nulo(v):
    return v is None or (isinstance(v, float) and pd.isna(v))


# ══════════════════════════════════════════════════════════════════════════
# OBJETIVOS PARA EL LIBRO DE CONTROL
# ══════════════════════════════════════════════════════════════════════════
def construir_objetivos(filas_hist, claves_nuevas=None):
    """
    De las obligaciones ACTIVAS del histórico arma lo que va a cada hoja:
      - caja   : Particulares agrupadas por identificación (suma de cuotas)
      - bancos : Bancarias, una por obligación (con capital e interés)
    Las Gerencia no van a ninguna hoja.

    "antiguedad" es NUEVO si la obligación no estaba en el histórico antes de
    este cruce (en caja: si alguna de las obligaciones de ese proveedor es
    nueva) y EXISTENTE en caso contrario.
    """
    claves_nuevas = set(claves_nuevas or [])
    caja = {}
    bancos = []
    for f in filas_hist:
        v = f["valores"]
        if not _es_activo(v.get("estado_obligacion")):
            continue
        tipo = _norm_header(v.get("tipo_de_obligacion"))
        valor = _a_float(v.get("valor_cuota")) or 0.0
        id_prov = _norm_id(v.get("identificacion_proveedor"))
        nombre = _limpiar_nombre(v.get("nombre_proveedor"))
        es_nueva = f["_clave"] in claves_nuevas
        if tipo == "particular":
            if id_prov not in caja:
                caja[id_prov] = {"id_proveedor": id_prov, "identificacion_proveedor": id_prov,
                                 "nombre_proveedor": nombre, "nueva": False,
                                 "valores": {"valor_cuota": 0.0}}
            c = caja[id_prov]
            if not c["nombre_proveedor"] and nombre:
                c["nombre_proveedor"] = nombre
            c["nueva"] = c["nueva"] or es_nueva
            c["valores"]["valor_cuota"] += valor
        elif tipo == "bancario":
            bancos.append({
                "empresa": _texto(v.get("empresa")),
                "factura": _limpiar_factura(v.get("factura")),
                "tipo_documento": _limpiar_nombre(v.get("tipo_documento")),
                "id_proveedor": id_prov, "identificacion_proveedor": id_prov,
                "nombre_proveedor": nombre,
                "antiguedad": ETIQUETA_NUEVO if es_nueva else ETIQUETA_EXISTENTE,
                "valores": {
                    "valor_cuota": round(valor, 2),
                    "capital": round(_a_float(v.get("capital")) or 0.0, 2),
                    "interes": round(_a_float(v.get("interes")) or 0.0, 2),
                },
            })
    caja_lista = list(caja.values())
    for c in caja_lista:
        c["valores"]["valor_cuota"] = round(c["valores"]["valor_cuota"], 2)
        c["antiguedad"] = ETIQUETA_NUEVO if c["nueva"] else ETIQUETA_EXISTENTE
    return {"caja": caja_lista, "bancos": bancos}


# ══════════════════════════════════════════════════════════════════════════
# LIBRO DE CONTROL
# ══════════════════════════════════════════════════════════════════════════
def _ultima_fila_datos(ws, cols):
    for r in range(ws.max_row, 1, -1):
        if any(_texto(ws.cell(r, c).value) for c in cols):
            return r
    return 1


def _quitar_validaciones_columna(ws, letra):
    mantener = []
    for dv in ws.data_validations.dataValidation:
        rangos = str(dv.sqref).split()
        columnas = {re.match(r"([A-Z]+)", p).group(1) for p in rangos if re.match(r"([A-Z]+)", p)}
        if columnas != {letra}:
            mantener.append(dv)
    ws.data_validations.dataValidation = mantener


def _nombres_mes(spec, mes, anio):
    return {m: f"{m}_{mes}_{anio}" for m in spec["metricas"]}


def _agregar_columna(ws, nombre, metrica):
    """Agrega una columna al final copiando estilo y ancho de la del mes anterior."""
    h = _headers(ws)
    previos = [c for k, c in h.items() if k.startswith(f"{metrica}_")]
    ref = max(previos) if previos else _ultima_col(ws)
    nueva = _ultima_col(ws) + 1
    ws.cell(1, nueva).value = nombre
    ancho = None
    if ref:
        _copiar_estilo(ws.cell(1, ref), ws.cell(1, nueva))
        ancho = ws.column_dimensions[get_column_letter(ref)].width
    ws.column_dimensions[get_column_letter(nueva)].width = ancho or 24
    return nueva


def _asegurar_columnas_mes(ws, spec, mes, anio):
    """Devuelve {métrica: número de columna} del mes; crea las que no existan."""
    metricas = spec["metricas"]
    nombres = _nombres_mes(spec, mes, anio)
    h = _headers(ws)

    # ya existe (total o parcialmente) el bloque del mes
    if any(nombres[m] in h for m in metricas):
        cols = {}
        for m in metricas:
            cols[m] = h[nombres[m]] if nombres[m] in h else _agregar_columna(ws, nombres[m], m)
            h = _headers(ws)
        return cols

    # Columnas "viejas" sin año y contiguas (valor_cuota_<mes> o valor_cuota, seguidas de
    # las demás métricas): se renombran en lugar de duplicarlas.
    base = h.get(f"valor_cuota_{mes}")
    if base is None and not any(k.startswith("valor_cuota_") for k in h):
        base = h.get("valor_cuota")
    if base is not None:
        por_col = {c: k for k, c in h.items()}
        if all(por_col.get(base + i) in (m, f"{m}_{mes}")
               for i, m in enumerate(metricas) if i > 0):
            for i, m in enumerate(metricas):
                ws.cell(1, base + i).value = nombres[m]
            return {m: base + i for i, m in enumerate(metricas)}

    cols = {}
    for m in metricas:
        cols[m] = _agregar_columna(ws, nombres[m], m)
    return cols


def _archivar_fila(wb, hoja_origen, ws_origen, fila, headers, fecha_op):
    """Copia la fila (con todo su historial de meses) a la hoja CANCELADAS."""
    if HOJA_CANCELADAS in wb.sheetnames:
        wc = wb[HOJA_CANCELADAS]
    else:
        wc = wb.create_sheet(HOJA_CANCELADAS)
        wc.cell(1, 1).value = "hoja_origen"
        wc.cell(1, 2).value = "fecha_cancelacion"
    hc = _headers(wc)
    destino = max(_ultima_fila_datos(wc, [1, 2]), 1) + 1
    wc.cell(destino, 1).value = hoja_origen
    wc.cell(destino, 2).value = fecha_op
    wc.cell(destino, 2).number_format = "DD/MM/YYYY"
    for nombre_norm, col in headers.items():
        valor = ws_origen.cell(fila, col).value
        if nombre_norm not in hc:
            nueva = _ultima_col(wc) + 1
            wc.cell(1, nueva).value = _texto(ws_origen.cell(1, col).value)
            wc.column_dimensions[get_column_letter(nueva)].width = (
                ws_origen.column_dimensions[get_column_letter(col)].width or 20)
            hc[nombre_norm] = nueva
        celda = wc.cell(destino, hc[nombre_norm])
        celda.value = valor
        _copiar_estilo(ws_origen.cell(fila, col), celda)


def _valor_columna_nueva(nombre_norm, obj, spec, nombres):
    """Valor de cada columna al crear una fila nueva en la hoja de control."""
    if nombre_norm in ("id_proveedor", "identificacion_proveedor"):
        return _id_para_celda(obj["id_proveedor"])
    if nombre_norm in ("nombre_proveedor", "empresa", "factura", "tipo_documento"):
        return obj.get(nombre_norm) or None
    if nombre_norm == "antiguedad_credito":
        return obj.get("antiguedad")
    for m in spec["metricas"]:
        if nombre_norm == nombres[m]:
            return ESTADO_CUOTA_PENDIENTE if m == METRICA_ESTADO else obj["valores"].get(m, 0.0)
    for m in spec["metricas"]:
        if nombre_norm.startswith(f"{m}_"):          # meses anteriores a la aparición
            return None if m == METRICA_ESTADO else 0
    return None


def _vista_hoja(ws, spec, mes, anio, cols_clave):
    """Lo que quedará en la hoja: columnas fijas + las del mes en curso."""
    h = _headers(ws)
    del_mes = set(_nombres_mes(spec, mes, anio).values())
    prefijos = tuple(f"{m}_" for m in spec["metricas"])
    cols = [(c, _texto(ws.cell(1, c).value)) for k, c in h.items()
            if k in del_mes or not k.startswith(prefijos)]
    ultima = _ultima_fila_datos(ws, cols_clave)
    filas = [[ws.cell(r, c).value for c, _ in cols] for r in range(2, ultima + 1)]
    return pd.DataFrame(filas, columns=[n for _, n in cols])


def _expandir_columnas_agrupadas(ws):
    """openpyxl guarda anchos de columnas contiguas como un solo rango (min-max); se separan
    en una dimensión por columna para poder ocultar/agrupar cada una sin pisar a las demás."""
    for clave, dim in list(ws.column_dimensions.items()):
        if dim.min and dim.max and dim.max > dim.min:
            ancho, oculto, nivel = dim.width, dim.hidden, dim.outlineLevel
            desde, hasta = dim.min, dim.max
            del ws.column_dimensions[clave]
            for c in range(desde, hasta + 1):
                nueva = ws.column_dimensions[get_column_letter(c)]
                nueva.width = ancho
                nueva.hidden = oculto
                nueva.outlineLevel = nivel


def _ocultar_meses_anteriores(ws, spec, mes, anio):
    """Deja visibles las columnas del mes en curso y agrupa/oculta las de los demás meses."""
    _expandir_columnas_agrupadas(ws)
    del_mes = set(_nombres_mes(spec, mes, anio).values())
    prefijos = tuple(f"{m}_" for m in spec["metricas"])
    for nombre, col in _headers(ws).items():
        if not nombre.startswith(prefijos):
            continue
        dim = ws.column_dimensions[get_column_letter(col)]
        if nombre in del_mes:
            dim.hidden = False
            dim.outlineLevel = 0
        else:
            dim.hidden = True
            dim.outlineLevel = 1
    ws.sheet_properties.outlinePr.summaryRight = True


def _actualizar_hoja_control(wb, spec, objetivos, mes, anio, fecha_op):
    resumen = {"hoja": spec["hoja"], "actualizadas": [], "nuevas": [], "archivadas": [],
               "advertencias": [], "vista": None}
    ws = wb[spec["hoja"]] if spec["hoja"] in wb.sheetnames else wb.create_sheet(spec["hoja"])

    if not _headers(ws):
        for i, h in enumerate(spec["encabezados"], 1):
            ws.cell(1, i).value = h
            ws.column_dimensions[get_column_letter(i)].width = 24

    h = _headers(ws)
    campos_clave = spec["clave"]
    faltan = [c for c in campos_clave if _norm_header(c) not in h]
    if faltan:
        raise ValueError(f"En la hoja '{spec['hoja']}' faltan las columnas: {', '.join(faltan)}")

    cols_mes = _asegurar_columnas_mes(ws, spec, mes, anio)
    h = _headers(ws)
    nombres = _nombres_mes(spec, mes, anio)
    c_val, c_est = cols_mes["valor_cuota"], cols_mes[METRICA_ESTADO]
    c_ant = h.get("antiguedad_credito")
    metricas_num = [m for m in spec["metricas"] if m != METRICA_ESTADO]
    cols_clave = [h[_norm_header(c)] for c in campos_clave]

    def clave_obj(o):
        return tuple(_norm_campo(c, o.get(c)) for c in campos_clave)

    objetivos_por_clave = {}
    for o in objetivos:
        objetivos_por_clave.setdefault(clave_obj(o), o)

    ultima = _ultima_fila_datos(ws, cols_clave)
    # primera vez que se procesa este mes: la marca EXISTENTE/NUEVO se recalcula;
    # si el mes ya tenía datos (re-ejecución), se respeta lo que ya estaba escrito
    primera_corrida = not any(_texto(ws.cell(r, c_val).value) for r in range(2, ultima + 1))

    filas_hoja = {}
    for r in range(2, ultima + 1):
        k = tuple(_norm_campo(c, ws.cell(r, col).value) for c, col in zip(campos_clave, cols_clave))
        if any(k):
            filas_hoja.setdefault(k, r)

    # 1) filas existentes: solo los valores del mes; el estado únicamente si está vacío
    for k, o in objetivos_por_clave.items():
        if k not in filas_hoja:
            continue
        r = filas_hoja[k]
        anterior = _a_float(ws.cell(r, c_val).value)
        for m in metricas_num:
            celda = ws.cell(r, cols_mes[m])
            celda.value = o["valores"].get(m, 0.0)
            if celda.number_format == "General":
                celda.number_format = FORMATO_MONEDA
        estado = _texto(ws.cell(r, c_est).value)
        if not estado:
            ws.cell(r, c_est).value = ESTADO_CUOTA_PENDIENTE
        elif estado.upper() == ESTADO_CUOTA_PAGADO and anterior is not None \
                and abs(anterior - o["valores"]["valor_cuota"]) > 0.005:
            resumen["advertencias"].append(
                f"{_texto(o.get('nombre_proveedor'))} ({_texto(o.get('id_proveedor'))}): ya estaba "
                f"PAGADO con {anterior:,.2f} y el valor nuevo es {o['valores']['valor_cuota']:,.2f}.")
        if c_ant and primera_corrida:
            ws.cell(r, c_ant).value = o["antiguedad"]
        resumen["actualizadas"].append(o)

    # 2) filas que ya no están activas: se archivan en CANCELADAS y se eliminan
    a_archivar = sorted(((r, k) for k, r in filas_hoja.items() if k not in objetivos_por_clave),
                        reverse=True)
    for r, k in a_archivar:
        info = {c: _texto(ws.cell(r, col).value) for c, col in zip(campos_clave, cols_clave)}
        _archivar_fila(wb, spec["hoja"], ws, r, h, fecha_op)
        resumen["archivadas"].append(info)
    for r, _ in a_archivar:
        ws.delete_rows(r)

    # 3) obligaciones nuevas: al final, con 0 en los meses anteriores
    ultima = _ultima_fila_datos(ws, cols_clave)
    ultima_col = _ultima_col(ws)
    for k, o in objetivos_por_clave.items():
        if k in filas_hoja:
            continue
        ultima += 1
        for nombre_norm, col in h.items():
            ws.cell(ultima, col).value = _valor_columna_nueva(nombre_norm, o, spec, nombres)
        if ultima - 1 >= 2:
            for col in range(1, ultima_col + 1):
                _copiar_estilo(ws.cell(ultima - 1, col), ws.cell(ultima, col))
        for nombre_norm, col in h.items():
            if nombre_norm.startswith(tuple(f"{m}_" for m in metricas_num)) \
                    and ws.cell(ultima, col).number_format == "General":
                ws.cell(ultima, col).number_format = FORMATO_MONEDA
        resumen["nuevas"].append(o)

    # 4) lista desplegable PENDIENTE / PAGADO y autofiltro
    ultima_final = max(_ultima_fila_datos(ws, cols_clave), 2)
    letra = get_column_letter(c_est)
    _quitar_validaciones_columna(ws, letra)
    dv = DataValidation(type="list", formula1=f'"{ESTADO_CUOTA_PENDIENTE},{ESTADO_CUOTA_PAGADO}"',
                        allow_blank=True)
    dv.add(f"{letra}2:{letra}{ultima_final}")
    ws.add_data_validation(dv)
    ws.auto_filter.ref = f"A1:{get_column_letter(_ultima_col(ws))}{ultima_final}"

    if spec.get("ocultar_meses_anteriores"):
        _ocultar_meses_anteriores(ws, spec, mes, anio)

    resumen["vista"] = _vista_hoja(ws, spec, mes, anio, cols_clave)
    return resumen

# ══════════════════════════════════════════════════════════════════════════
# ORQUESTACIÓN Y GUARDADO
# ══════════════════════════════════════════════════════════════════════════
def _verificar_escritura(ruta):
    try:
        with open(ruta, "r+b"):
            pass
    except PermissionError:
        raise PermissionError(
            f"No se puede escribir en «{os.path.basename(ruta)}». Probablemente está abierto en "
            "Excel (tuyo o de otra persona). Ciérralo e inténtalo de nuevo; no se guardó nada.")
    except FileNotFoundError:
        raise FileNotFoundError(f"No se encontró el archivo: {ruta}")


def _respaldar(ruta, dir_respaldos):
    os.makedirs(dir_respaldos, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(ruta))
    destino = os.path.join(dir_respaldos, f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}")
    shutil.copy2(ruta, destino)
    previos = sorted(f for f in os.listdir(dir_respaldos) if f.startswith(base + "_") and f.endswith(ext))
    for viejo in previos[:-MAX_RESPALDOS_POR_ARCHIVO]:
        try:
            os.remove(os.path.join(dir_respaldos, viejo))
        except OSError:
            pass
    return destino


def ejecutar(df_rep, ruta_historico, ruta_control, mes_num, anio, tipos_nuevas,
             incluir_bancos=True, guardar=False, dir_respaldos=None):
    """
    Calcula (y si guardar=True, escribe) la actualización completa. Con
    guardar=False es un ensayo: trabaja en memoria y no toca ningún archivo.
    """
    mes = MESES_ES[mes_num - 1]
    fecha_op = datetime.now().replace(microsecond=0)

    wb_h = openpyxl.load_workbook(ruta_historico)
    if HOJA_HISTORICO not in wb_h.sheetnames:
        raise ValueError(f"El libro histórico no tiene una hoja llamada '{HOJA_HISTORICO}'.")
    res_h, filas_proj = _aplicar_historico(wb_h, df_rep, tipos_nuevas)
    objetivos = construir_objetivos(filas_proj, res_h["claves_nuevas"])

    wb_c = openpyxl.load_workbook(ruta_control)
    res_caja = _actualizar_hoja_control(wb_c, SPEC_CAJA, objetivos["caja"], mes, anio, fecha_op)
    res_bancos = None
    if incluir_bancos:
        res_bancos = _actualizar_hoja_control(wb_c, SPEC_BANCOS, objetivos["bancos"], mes, anio, fecha_op)

    resultado = {"historico": res_h, "caja": res_caja, "bancos": res_bancos,
                 "objetivos": objetivos, "respaldos": []}

    if guardar:
        dir_resp = dir_respaldos or DIR_RESPALDOS
        _verificar_escritura(ruta_historico)
        _verificar_escritura(ruta_control)
        b_h = _respaldar(ruta_historico, dir_resp)
        b_c = _respaldar(ruta_control, dir_resp)
        resultado["respaldos"] = [b_h, b_c]
        wb_h.save(ruta_historico)
        try:
            wb_c.save(ruta_control)
        except Exception as e:
            shutil.copy2(b_h, ruta_historico)   # deja el histórico como estaba
            raise RuntimeError(
                f"No se pudo guardar el libro de control ({e}). El histórico se restauró a su "
                "versión anterior; no quedó ningún cambio.")
    return resultado
