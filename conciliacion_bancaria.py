import pandas as pd
import numpy as np
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from copy import copy

# ══════════════════════════════════════════════════════════════════════════
# CATÁLOGO DE EMPRESAS Y CUENTAS
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
    "AmanecerDePascua":     [("AMANECER DE PASCUA CTA CTE 5004",  111005002)],
    "AnochecerDePascua":    [("ANOCHECER DE PASCUA CTA CTE 4999", 111005002)],
    "MañanaDePascua":       [("MAÑANA DE PASCUA CTA CTE 7481",    111005002)],
    "GriGroup":             [("GRI CTA CTE 1525",                  111005002)],
    "Movicap":              [("MOVICAPSAS CTE 4395",               111005002),
                             ("AHO MOVICAP 2050",                  112005002)],
    "ConfianzaGlobal":      [("OCCIDENTE AHORRO CTA No 9890",      112005001)],
    "SeguroConfianzaGlobal":[("AHO SEGUROS CONFIANZA 2449",        112005002)],
    "Suprogreso":           [("BANCOLOMBIA CT AHO 9260",           111005001)],
}

MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
         "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]


# ══════════════════════════════════════════════════════════════════════════
# PARSERS DE EXTRACTOS BANCARIOS
# ══════════════════════════════════════════════════════════════════════════

def detectar_y_parsear_extracto(archivo):
    nombre = archivo.name.lower()
    if nombre.endswith('.csv'):
        return _parsear_bancolombia_csv(archivo)
    elif nombre.endswith('.xlsx'):
        return _parsear_davivienda_xlsx(archivo)
    elif nombre.endswith('.xls'):
        return _parsear_xls(archivo)
    else:
        raise ValueError(f"Formato no reconocido: {archivo.name}")


def _parsear_bancolombia_csv(archivo):
    content = archivo.read().decode('utf-8', errors='replace')
    rows = []
    for line in content.strip().split('\n'):
        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 8: continue
        try:
            fecha_str = parts[3].strip()
            if len(fecha_str) == 8:
                fecha = datetime.strptime(fecha_str, '%Y%m%d')
            else:
                continue
            valor   = float(parts[5].strip().replace(' ',''))
            concepto = parts[7].strip()
            debito  = round(abs(valor), 2) if valor < 0 else 0.0
            credito = round(valor, 2)      if valor > 0 else 0.0
            rows.append({"FECHA": fecha, "DEBITO": debito, "CREDITO": credito, "CONCEPTO": concepto})
        except Exception:
            continue
    return pd.DataFrame(rows)


def _parsear_davivienda_xlsx(archivo):
    df = pd.read_excel(archivo, sheet_name=0, header=0)
    rows = []
    for _, row in df.iterrows():
        try:
            fecha   = pd.to_datetime(row.iloc[0], dayfirst=True)
            concepto = str(row.iloc[2]).strip()
            tipo    = str(row.iloc[3]).strip().lower()
            val_str = str(row.iloc[7]).replace('$','').replace('.','').replace(',','.').strip()
            valor   = float(val_str)
            debito  = round(valor, 2) if 'débito' in tipo or 'debito' in tipo else 0.0
            credito = round(valor, 2) if 'crédito' in tipo or 'credito' in tipo else 0.0
            rows.append({"FECHA": fecha, "DEBITO": debito, "CREDITO": credito, "CONCEPTO": concepto})
        except Exception:
            continue
    return pd.DataFrame(rows)


def _parsear_xls(archivo):
    import xlrd
    content = archivo.read()
    wb = xlrd.open_workbook(file_contents=content)
    ws = wb.sheet_by_index(0)

    # Detectar versión
    fila_datos = None
    tipo = None
    for i in range(min(30, ws.nrows)):
        row_str = ' '.join(str(v) for v in ws.row_values(i)).lower()
        if 'fecha movimiento' in row_str or ('débitos' in row_str and i < 10):
            fila_datos = i + 1; tipo = 'v1'; break
        if 'débitos' in row_str and i >= 20:
            fila_datos = i + 1; tipo = 'v2'; break

    if fila_datos is None:
        raise ValueError("No se pudo detectar el formato XLS.")

    # Para V2: detectar si créditos están en col N(13) u O(14)
    # Bogotá tiene col N vacía, créditos en col O
    col_cre_v2 = 13  # default Occidente V2
    if tipo == 'v2' and ws.nrows > fila_datos:
        # Verificar encabezado: si col N es vacía y col O tiene "Crédito"
        hdr = ws.row_values(fila_datos - 1)
        if len(hdr) > 14:
            if str(hdr[13]).strip() == "" and "cr" in str(hdr[14]).lower():
                col_cre_v2 = 14  # Bogotá

    rows = []
    for i in range(fila_datos, ws.nrows):
        row = ws.row_values(i)
        try:
            if tipo == 'v1':
                fecha    = _parse_fecha_str(str(row[0]))
                concepto = str(row[2]).strip()
                deb_str  = str(row[4]).replace('$','').replace(',','').strip()
                cre_str  = str(row[5]).replace('$','').replace(',','').strip()
                debito   = round(float(deb_str or 0), 2)
                credito  = round(float(cre_str or 0), 2)
            else:
                fecha    = _parse_fecha_str(str(row[1]))
                concepto = str(row[4]).strip()
                debito   = round(float(row[12] or 0), 2)
                credito  = round(float(row[col_cre_v2] or 0), 2) if len(row) > col_cre_v2 else 0.0
            if not fecha or (debito == 0 and credito == 0): continue
            rows.append({"FECHA": fecha, "DEBITO": debito, "CREDITO": credito, "CONCEPTO": concepto})
        except Exception:
            continue
    return pd.DataFrame(rows)


def _parse_fecha_str(s):
    s = s.strip()
    for fmt in ['%Y/%m/%d','%d/%m/%Y','%Y-%m-%d','%d-%m-%Y']:
        try: return datetime.strptime(s, fmt)
        except: continue
    return None


# ══════════════════════════════════════════════════════════════════════════
# LÓGICA DE CONCILIACIÓN (replica la macro VBA)
# ══════════════════════════════════════════════════════════════════════════

def conciliar(df_banco, df_aux):
    """
    Replica exactamente la lógica de la macro ConciliacionBancaria:
    1. Cruce exacto: banco.debito ↔ aux.credito / banco.credito ↔ aux.debito
    2. Reclasificaciones auxiliar: mismo valor deb↔cre en filas distintas no cruzadas
    3. Banco no cruzado → falta en auxiliar
    4. Auxiliar no cruzado y no reclasificado → está de más
    """
    df_b = df_banco.copy()
    df_a = df_aux.copy()

    df_b["DEBITO"]  = pd.to_numeric(df_b["DEBITO"],  errors="coerce").fillna(0).round(2)
    df_b["CREDITO"] = pd.to_numeric(df_b["CREDITO"], errors="coerce").fillna(0).round(2)
    df_a["debito"]  = pd.to_numeric(df_a.get("debito",  0), errors="coerce").fillna(0).round(2)
    df_a["credito"] = pd.to_numeric(df_a.get("credito", 0), errors="coerce").fillna(0).round(2)

    nb = len(df_b)
    na = len(df_a)

    b_match   = [False] * nb
    a_match   = [False] * na
    a_reclass = [False] * na

    # PASO 1: Cruce exacto banco ↔ auxiliar
    for i in range(nb):
        for j in range(na):
            if b_match[i] or a_match[j]: continue
            bd = df_b.iloc[i]["DEBITO"]
            bc = df_b.iloc[i]["CREDITO"]
            ad = df_a.iloc[j]["debito"]
            ac = df_a.iloc[j]["credito"]
            if bd > 0 and ac > 0 and abs(bd - ac) < 0.01:
                b_match[i] = True; a_match[j] = True
            elif bc > 0 and ad > 0 and abs(bc - ad) < 0.01:
                b_match[i] = True; a_match[j] = True

    # PASO 2: Reclasificaciones en auxiliar no cruzado
    # mismo monto en débito de una fila y crédito de otra
    for i in range(na):
        if a_match[i] or a_reclass[i]: continue
        ad_i = df_a.iloc[i]["debito"]
        if ad_i <= 0: continue
        for k in range(i+1, na):
            if a_match[k] or a_reclass[k]: continue
            ac_k = df_a.iloc[k]["credito"]
            if ac_k > 0 and abs(ad_i - ac_k) < 0.01:
                a_reclass[i] = True
                a_reclass[k] = True
                break

    # Totales auxiliar
    saldo_aux_deb  = round(df_a["debito"].sum(), 2)
    saldo_aux_cred = round(df_a["credito"].sum(), 2)

    # Totales banco (espejo)
    saldo_banco_deb  = round(df_b["DEBITO"].sum(), 2)
    saldo_banco_cred = round(df_b["CREDITO"].sum(), 2)

    # Construir partidas conciliatorias
    partidas = []  # {tipo, concepto_aux, fecha_aux, deb_aux, cred_aux, concepto_banco, fecha_banco, deb_banco, cred_banco}

    # 2A: Reclasificaciones
    paired = [False] * na
    for i in range(na):
        if not a_reclass[i] or paired[i]: continue
        ad_i = df_a.iloc[i]["debito"]
        if ad_i <= 0: continue
        for k in range(i+1, na):
            if not a_reclass[k] or paired[k]: continue
            ac_k = df_a.iloc[k]["credito"]
            if ac_k > 0 and abs(ad_i - ac_k) < 0.01:
                paired[i] = True; paired[k] = True
                # Línea débito reclasificado (negativo)
                partidas.append({
                    "tipo": "RECLASIFICACION",
                    "conc_aux": df_a.iloc[i].get("descripcion",""),
                    "fecha_aux": df_a.iloc[i].get("fecha"),
                    "deb": -ad_i, "cred": 0,
                    "conc_banco": "", "fecha_banco": None,
                })
                # Línea crédito reclasificado (negativo)
                partidas.append({
                    "tipo": "RECLASIFICACION",
                    "conc_aux": "",
                    "fecha_aux": df_a.iloc[k].get("fecha"),
                    "deb": 0, "cred": -ac_k,
                    "conc_banco": df_a.iloc[k].get("descripcion",""),
                    "fecha_banco": df_a.iloc[k].get("fecha"),
                })
                break

    # 2B: Banco no cruzado (falta en auxiliar — positivo)
    for i in range(nb):
        if b_match[i]: continue
        bd = df_b.iloc[i]["DEBITO"]
        bc = df_b.iloc[i]["CREDITO"]
        conc = df_b.iloc[i]["CONCEPTO"]
        fecha = df_b.iloc[i]["FECHA"]
        if bd > 0:
            # banco débito = empresa crédito faltante → positivo en CRÉDITO
            partidas.append({"tipo": "SOLO_BANCO", "conc_aux": "", "fecha_aux": None,
                             "deb": 0, "cred": bd,
                             "conc_banco": conc, "fecha_banco": fecha})
        else:
            # banco crédito = empresa débito faltante → positivo en DÉBITO
            partidas.append({"tipo": "SOLO_BANCO", "conc_aux": conc, "fecha_aux": fecha,
                             "deb": bc, "cred": 0,
                             "conc_banco": "", "fecha_banco": None})

    # 2C: Auxiliar no cruzado y no reclasificado (está de más — negativo)
    for j in range(na):
        if a_match[j] or a_reclass[j]: continue
        ad = df_a.iloc[j]["debito"]
        ac = df_a.iloc[j]["credito"]
        conc  = df_a.iloc[j].get("descripcion","")
        fecha = df_a.iloc[j].get("fecha")
        if ad > 0:
            partidas.append({"tipo": "SOLO_AUX", "conc_aux": conc, "fecha_aux": fecha,
                             "deb": -ad, "cred": 0,
                             "conc_banco": "", "fecha_banco": None})
        else:
            partidas.append({"tipo": "SOLO_AUX", "conc_aux": "", "fecha_aux": None,
                             "deb": 0, "cred": -ac,
                             "conc_banco": conc, "fecha_banco": fecha})

    # Total auxiliar conciliado = saldo_aux + partidas
    total_conc_deb  = round(saldo_aux_deb  + sum(p["deb"]  for p in partidas), 2)
    total_conc_cred = round(saldo_aux_cred + sum(p["cred"] for p in partidas), 2)

    diferencia_deb  = round(total_conc_deb  - saldo_banco_cred, 2)
    diferencia_cred = round(total_conc_cred - saldo_banco_deb,  2)

    resumen = {
        "saldo_aux_deb":    saldo_aux_deb,
        "saldo_aux_cred":   saldo_aux_cred,
        "saldo_banco_deb":  saldo_banco_deb,
        "saldo_banco_cred": saldo_banco_cred,
        "total_conc_deb":   total_conc_deb,
        "total_conc_cred":  total_conc_cred,
        "diferencia_deb":   diferencia_deb,
        "diferencia_cred":  diferencia_cred,
        "n_reclasif":       sum(1 for p in partidas if p["tipo"]=="RECLASIFICACION"),
        "n_solo_banco":     sum(1 for p in partidas if p["tipo"]=="SOLO_BANCO"),
        "n_solo_aux":       sum(1 for p in partidas if p["tipo"]=="SOLO_AUX"),
        "conciliados_banco": sum(b_match),
        "conciliados_aux":   sum(a_match),
    }

    return partidas, df_b, df_a, resumen


# ══════════════════════════════════════════════════════════════════════════
# GENERACIÓN DEL EXCEL (3 hojas: CONCILIACION, BANCO, AUXILIAR)
# ══════════════════════════════════════════════════════════════════════════

def generar_excel(partidas, df_banco, df_aux, resumen, empresa, cuenta, mes, df_aux_original=None):
    wb = openpyxl.Workbook()

    _hoja_conciliacion(wb, partidas, resumen, empresa, cuenta, mes)
    _hoja_banco(wb, df_banco)
    _hoja_auxiliar(wb, df_aux, df_aux_original)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _hoja_conciliacion(wb, partidas, resumen, empresa, cuenta, mes):
    ws = wb.active
    ws.title = "CONCILIACION"

    # Colores
    azul     = PatternFill("solid", fgColor="BDD7EE")
    verde    = PatternFill("solid", fgColor="C6E0B4")
    amarillo = PatternFill("solid", fgColor="FFEB9C")
    rojo     = PatternFill("solid", fgColor="FFC7CE")
    bold     = Font(bold=True)
    fmt_num  = '#,##0.00;(#,##0.00);"-"'
    fmt_date = "DD/MM/YYYY"

    # Estructura de columnas (desde col A):
    # A=OBSERVACION | B=CONCEPTO | C=FECHA | D=DEBITO | E=CREDITO |
    # F=FECHA | G=CONCEPTO | H=OBSERVACION

    # Fila 1: encabezados tabla
    row = 1
    hdrs = ["OBSERVACION", "CONCEPTO", "FECHA", "DEBITO", "CREDITO",
            "FECHA", "CONCEPTO", "OBSERVACION"]
    for col, h in enumerate(hdrs, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = bold

    # Fila 2: SALDO AUXILIAR
    row = 2
    ws.cell(row=row, column=2, value="SALDO AUXILIAR")
    ws.cell(row=row, column=4, value=resumen["saldo_aux_deb"]).number_format  = fmt_num
    ws.cell(row=row, column=5, value=resumen["saldo_aux_cred"]).number_format = fmt_num
    ws.cell(row=row, column=7, value="SALDO AUXILIAR")
    for col in range(1, 9):
        ws.cell(row=row, column=col).fill = azul
        ws.cell(row=row, column=col).font = bold

    # Partidas conciliatorias
    row = 3
    for p in partidas:
        obs_tipo = {"RECLASIFICACION": "RECLASIFICACION",
                    "SOLO_BANCO":      "SOLO BANCO",
                    "SOLO_AUX":        "SOLO AUXILIAR"}.get(p["tipo"], "")

        # Lado auxiliar (cols A-E)
        if p["conc_aux"] or p["deb"] != 0:
            ws.cell(row=row, column=1, value=obs_tipo)
            ws.cell(row=row, column=2, value=p["conc_aux"])
            if p["fecha_aux"]:
                try:
                    ws.cell(row=row, column=3, value=pd.Timestamp(p["fecha_aux"]).to_pydatetime())
                    ws.cell(row=row, column=3).number_format = fmt_date
                except Exception:
                    pass
            if p["deb"] != 0:
                ws.cell(row=row, column=4, value=p["deb"]).number_format = fmt_num
            if p["cred"] != 0:
                ws.cell(row=row, column=5, value=p["cred"]).number_format = fmt_num

        # Lado banco (cols F-H)
        if p["conc_banco"] or p["fecha_banco"]:
            if p["fecha_banco"]:
                try:
                    ws.cell(row=row, column=6, value=pd.Timestamp(p["fecha_banco"]).to_pydatetime())
                    ws.cell(row=row, column=6).number_format = fmt_date
                except Exception:
                    pass
            ws.cell(row=row, column=7, value=p["conc_banco"])
            ws.cell(row=row, column=8, value=obs_tipo)
        row += 1

    # TOTAL AUXILIAR CONCILIADO
    ws.cell(row=row, column=2, value="TOTAL AUXILIAR CONCILIADO")
    ws.cell(row=row, column=4, value=resumen["total_conc_deb"]).number_format  = fmt_num
    ws.cell(row=row, column=5, value=resumen["total_conc_cred"]).number_format = fmt_num
    for col in range(1, 9):
        ws.cell(row=row, column=col).fill = verde
        ws.cell(row=row, column=col).font = bold
    row += 1

    # SALDO BANCO
    ws.cell(row=row, column=2, value="SALDO BANCO")
    ws.cell(row=row, column=4, value=resumen["saldo_banco_cred"]).number_format = fmt_num
    ws.cell(row=row, column=5, value=resumen["saldo_banco_deb"]).number_format  = fmt_num
    for col in range(1, 9):
        ws.cell(row=row, column=col).fill = amarillo
        ws.cell(row=row, column=col).font = bold
    row += 1

    # DIFERENCIA
    ws.cell(row=row, column=2, value="DIFERENCIA")
    ws.cell(row=row, column=4, value=resumen["diferencia_deb"]).number_format  = fmt_num
    ws.cell(row=row, column=5, value=resumen["diferencia_cred"]).number_format = fmt_num
    for col in range(1, 9):
        ws.cell(row=row, column=col).fill = rojo
        ws.cell(row=row, column=col).font = bold

    # Ancho columnas
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 50
    ws.column_dimensions["H"].width = 18


def _hoja_banco(wb, df_banco):
    ws = wb.create_sheet("BANCO")
    headers = ["FECHA","DEBITO","CREDITO","CONCEPTO"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h).font = Font(bold=True)
    for r, row in enumerate(df_banco.itertuples(index=False), 2):
        try: ws.cell(row=r, column=1, value=row.FECHA).number_format = "DD/MM/YYYY"
        except: ws.cell(row=r, column=1, value=str(row.FECHA))
        ws.cell(row=r, column=2, value=row.DEBITO)
        ws.cell(row=r, column=3, value=row.CREDITO)
        ws.cell(row=r, column=4, value=row.CONCEPTO)


def _hoja_auxiliar(wb, df_aux, df_aux_original=None):
    """Si se pasa df_aux_original, usa todas sus columnas; sino usa df_aux normalizado."""
    ws = wb.create_sheet("AUXILIAR")
    df = df_aux_original if df_aux_original is not None else df_aux
    cols = list(df.columns)
    bold = Font(bold=True)
    for col, h in enumerate(cols, 1):
        ws.cell(row=1, column=col, value=h).font = bold
    for r, row in enumerate(df.itertuples(index=False), 2):
        for col, val in enumerate(row, 1):
            try:
                ws.cell(row=r, column=col, value=val)
            except Exception:
                ws.cell(row=r, column=col, value=str(val))
