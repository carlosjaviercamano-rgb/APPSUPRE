import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

# ══════════════════════════════════════════════════════════════════════════
# CATÁLOGO DE EMPRESAS Y CUENTAS (leído de hoja CUENTAS)
# ══════════════════════════════════════════════════════════════════════════

EMPRESAS_CUENTAS = {
    "SUPRECREDITO": [
        ("BANCOLOMBIA AHORRO No 0605",  112005001),
        ("DAVIVIENDA CTA CTE No 7563",  111005005),
        ("OCCIDENTE CTA CTE No 5816",   111005002),
        ("BOGOTA CT CTE No 0165",       111005007),
    ],
    "SUPREMOTOS": [
        ("BANCOLOMBIA AHORRO CTA No 7140", 112005001),
        ("DAVIVIENDA CTA CTE No 7571",     111005005),
        ("OCCIDENTE CTA CTE No 9045",      111005002),
        ("BOGOTA CTA CTE No 9696",         111005007),
    ],
    "AmanecerDePascua": [
        ("AMANECER DE PASCUA CTA CTE 5004", 111005002),
    ],
    "AnochecerDePascua": [
        ("ANOCHECER DE PASCUA CTA CTE 4999", 111005002),
    ],
    "MañanaDePascua": [
        ("MAÑANA DE PASCUA CTA CTE 7481", 111005002),
    ],
    "GriGroup": [
        ("GRI CTA CTE 1525", 111005002),
    ],
    "Movicap": [
        ("MOVICAPSAS CTE 4395", 111005002),
        ("AHO MOVICAP 2050",    112005002),
    ],
    "ConfianzaGlobal": [
        ("OCCIDENTE AHORRO CTA No 9890", 112005001),
    ],
    "SeguroConfianzaGlobal": [
        ("AHO SEGUROS CONFIANZA 2449", 112005002),
    ],
    "Suprogreso": [
        ("BANCOLOMBIA CT AHO 9260", 111005001),
    ],
}

MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
         "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]


# ══════════════════════════════════════════════════════════════════════════
# PARSERS DE EXTRACTOS BANCARIOS
# ══════════════════════════════════════════════════════════════════════════

def detectar_y_parsear_extracto(archivo):
    """
    Detecta automáticamente el formato del extracto bancario y lo parsea.
    Retorna DataFrame con columnas: FECHA, DEBITO, CREDITO, CONCEPTO
    """
    nombre = archivo.name.lower()
    
    if nombre.endswith('.csv'):
        return _parsear_bancolombia_csv(archivo)
    elif nombre.endswith('.xlsx'):
        return _parsear_davivienda_xlsx(archivo)
    elif nombre.endswith('.xls') or nombre.endswith('.XLS'):
        return _parsear_xls(archivo)
    else:
        raise ValueError(f"Formato no reconocido: {archivo.name}")


def _parsear_bancolombia_csv(archivo):
    """BANCOLOMBIA CSV: col4=fecha, col6=valor(+cred/-deb), col8=concepto"""
    content = archivo.read().decode('utf-8', errors='replace')
    rows = []
    for line in content.strip().split('\n'):
        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 8:
            continue
        try:
            fecha_str = parts[3].strip()
            if len(fecha_str) == 8:
                fecha = datetime.strptime(fecha_str, '%Y%m%d')
            else:
                continue
            valor = float(parts[5].strip().replace(' ',''))
            concepto = parts[7].strip()
            debito  = abs(valor) if valor < 0 else 0.0
            credito = valor      if valor > 0 else 0.0
            rows.append({"FECHA": fecha, "DEBITO": round(debito,2),
                         "CREDITO": round(credito,2), "CONCEPTO": concepto})
        except Exception:
            continue
    return pd.DataFrame(rows)


def _parsear_davivienda_xlsx(archivo):
    """DAVIVIENDA XLSX: fila1=encabezados, colA=fecha, colC=concepto, colD=tipo, colH=valor"""
    df = pd.read_excel(archivo, sheet_name=0, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.iloc[0], dayfirst=True)
            concepto = str(row.iloc[2]).strip()
            tipo     = str(row.iloc[3]).strip().lower()
            val_str  = str(row.iloc[7]).replace('$','').replace('.','').replace(',','.').strip()
            valor    = float(val_str)
            debito  = round(valor, 2) if 'débito' in tipo or 'debito' in tipo else 0.0
            credito = round(valor, 2) if 'crédito' in tipo or 'credito' in tipo else 0.0
            rows.append({"FECHA": fecha, "DEBITO": debito,
                         "CREDITO": credito, "CONCEPTO": concepto})
        except Exception:
            continue
    return pd.DataFrame(rows)


def _parsear_xls(archivo):
    """
    XLS: detecta si es Occidente V1 (encabezados fila 7)
    o Bogotá/Occidente V2 (encabezados fila 27).
    """
    import xlrd
    content = archivo.read()
    wb = xlrd.open_workbook(file_contents=content)
    ws = wb.sheet_by_index(0)

    # Detectar versión buscando encabezados
    fila_datos = None
    tipo = None
    for i in range(min(30, ws.nrows)):
        row = ws.row_values(i)
        row_str = ' '.join(str(v) for v in row).lower()
        if 'fecha movimiento' in row_str or ('débitos' in row_str and 'créditos' in row_str and i < 10):
            fila_datos = i + 1  # datos empiezan en la siguiente fila
            tipo = 'v1'
            break
        if 'débitos' in row_str and 'créditos' in row_str and i >= 20:
            fila_datos = i + 1
            tipo = 'v2'
            break

    if fila_datos is None:
        raise ValueError("No se pudo detectar el formato del archivo XLS.")

    rows = []
    for i in range(fila_datos, ws.nrows):
        row = ws.row_values(i)
        try:
            if tipo == 'v1':
                # Col A=fecha, C=concepto, E=debito($), F=credito($)
                fecha   = _parse_fecha_str(str(row[0]))
                concepto = str(row[2]).strip()
                deb_str = str(row[4]).replace('$','').replace(',','').strip()
                cre_str = str(row[5]).replace('$','').replace(',','').strip()
                debito  = round(float(deb_str or 0), 2)
                credito = round(float(cre_str or 0), 2)
            else:
                # Col B=fecha, E=concepto, M=debito(num), N=credito(num)
                fecha    = _parse_fecha_str(str(row[1]))
                concepto = str(row[4]).strip()
                debito   = round(float(row[12] or 0), 2)
                credito  = round(float(row[13] or 0), 2)

            if not fecha or (debito == 0 and credito == 0):
                continue
            rows.append({"FECHA": fecha, "DEBITO": debito,
                         "CREDITO": credito, "CONCEPTO": concepto})
        except Exception:
            continue

    return pd.DataFrame(rows)


def _parse_fecha_str(s):
    """Parsea fechas en varios formatos."""
    s = s.strip()
    for fmt in ['%Y/%m/%d', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


# ══════════════════════════════════════════════════════════════════════════
# LÓGICA DE CONCILIACIÓN
# ══════════════════════════════════════════════════════════════════════════

def conciliar(df_banco, df_auxiliar):
    """
    Cruza banco vs auxiliar por valor exacto (2 decimales).
    Banco.debito  ↔ Auxiliar.credito
    Banco.credito ↔ Auxiliar.debito
    
    Retorna DataFrame con columnas:
    B_CONCEPTO, B_FECHA, B_DEBITO, B_CREDITO,
    A_FECHA, A_CONCEPTO, A_VALOR, ESTADO
    """
    # Normalizar valores a 2 decimales
    df_b = df_banco.copy()
    df_a = df_auxiliar.copy()
    df_b["DEBITO"]  = df_b["DEBITO"].fillna(0).round(2)
    df_b["CREDITO"] = df_b["CREDITO"].fillna(0).round(2)
    df_a["debito"]  = pd.to_numeric(df_a.get("debito",  0), errors="coerce").fillna(0).round(2)
    df_a["credito"] = pd.to_numeric(df_a.get("credito", 0), errors="coerce").fillna(0).round(2)
    df_a["valor"]   = pd.to_numeric(df_a.get("valor",   0), errors="coerce").fillna(0).round(2)

    usados_b = set()
    usados_a = set()
    resultado = []

    # PASO 1: Cruce banco.debito ↔ auxiliar.credito (valor exacto)
    for ib, rb in df_b.iterrows():
        if ib in usados_b: continue
        if rb["DEBITO"] <= 0: continue
        val_b = rb["DEBITO"]
        for ia, ra in df_a.iterrows():
            if ia in usados_a: continue
            if abs(ra["credito"] - val_b) < 0.01:
                resultado.append({
                    "B_CONCEPTO": rb["CONCEPTO"], "B_FECHA": rb["FECHA"],
                    "B_DEBITO":   val_b,          "B_CREDITO": 0,
                    "A_FECHA":    ra.get("fecha"), "A_CONCEPTO": ra.get("descripcion",""),
                    "A_DEBITO":   0,               "A_CREDITO": ra["credito"],
                    "ESTADO": "CONCILIADO"
                })
                usados_b.add(ib); usados_a.add(ia)
                break

    # PASO 2: Cruce banco.credito ↔ auxiliar.debito (valor exacto)
    for ib, rb in df_b.iterrows():
        if ib in usados_b: continue
        if rb["CREDITO"] <= 0: continue
        val_b = rb["CREDITO"]
        for ia, ra in df_a.iterrows():
            if ia in usados_a: continue
            if abs(ra["debito"] - val_b) < 0.01:
                resultado.append({
                    "B_CONCEPTO": rb["CONCEPTO"], "B_FECHA": rb["FECHA"],
                    "B_DEBITO":   0,              "B_CREDITO": val_b,
                    "A_FECHA":    ra.get("fecha"), "A_CONCEPTO": ra.get("descripcion",""),
                    "A_DEBITO":   ra["debito"],    "A_CREDITO": 0,
                    "ESTADO": "CONCILIADO"
                })
                usados_b.add(ib); usados_a.add(ia)
                break

    # PASO 3: Solo en banco (sin cruce)
    for ib, rb in df_b.iterrows():
        if ib in usados_b: continue
        resultado.append({
            "B_CONCEPTO": rb["CONCEPTO"], "B_FECHA": rb["FECHA"],
            "B_DEBITO":   rb["DEBITO"],   "B_CREDITO": rb["CREDITO"],
            "A_FECHA": None, "A_CONCEPTO": "", "A_DEBITO": 0, "A_CREDITO": 0,
            "ESTADO": "SOLO BANCO"
        })

    # PASO 4: Solo en auxiliar (sin cruce)
    for ia, ra in df_a.iterrows():
        if ia in usados_a: continue
        resultado.append({
            "B_CONCEPTO": "", "B_FECHA": None, "B_DEBITO": 0, "B_CREDITO": 0,
            "A_FECHA":    ra.get("fecha"), "A_CONCEPTO": ra.get("descripcion",""),
            "A_DEBITO":   ra["debito"],    "A_CREDITO": ra["credito"],
            "ESTADO": "SOLO AUXILIAR"
        })

    df_result = pd.DataFrame(resultado)

    # Calcular totales
    tot_b_deb = round(df_b["DEBITO"].sum(), 2)
    tot_b_cre = round(df_b["CREDITO"].sum(), 2)
    tot_a_deb = round(df_a["debito"].sum(), 2)
    tot_a_cre = round(df_a["credito"].sum(), 2)

    resumen = {
        "total_banco_debito":    tot_b_deb,
        "total_banco_credito":   tot_b_cre,
        "total_aux_debito":      tot_a_deb,
        "total_aux_credito":     tot_a_cre,
        "diferencia_debito":     round(tot_b_deb - tot_a_cre, 2),
        "diferencia_credito":    round(tot_b_cre - tot_a_deb, 2),
        "conciliados":           sum(1 for r in resultado if r["ESTADO"]=="CONCILIADO"),
        "solo_banco":            sum(1 for r in resultado if r["ESTADO"]=="SOLO BANCO"),
        "solo_auxiliar":         sum(1 for r in resultado if r["ESTADO"]=="SOLO AUXILIAR"),
    }

    return df_result, resumen
