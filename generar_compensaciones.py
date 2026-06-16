import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import io
import zipfile
import streamlit as st

# NIT por entidad bancaria
NIT_ENTIDAD = {
    "BANCOLOMBIA": "890903938",
    "DAVIVIENDA":  "860034313",
    "PSE":         "830078512",
    "EFECTY":      "830131993",
    "RECORD":      "111111111",
    "OCCIDENTE":   "860002395",
}

CODIGO_CENTRO_COSTO  = "102"
CODIGO_CUENTA_DEBITO = "141299011"
CODIGO_CUENTA_CREDITO = "110505017"


def crear_compensaciones(df_area_banco, config):
    if df_area_banco is None or df_area_banco.empty:
        raise ValueError("No hay datos en Área de Banco para generar compensaciones.")

    df = df_area_banco.copy()

    # Normalizar fechas
    df["FECHA_NORM"]     = pd.to_datetime(df["FECHA"],           errors="coerce").dt.normalize()
    df["FECHA_DOC_NORM"] = pd.to_datetime(df["FECHA_DOCUMENTO"], errors="coerce").dt.normalize()

    # Agrupar por FECHA_DOCUMENTO
    fechas_doc = df["FECHA_DOC_NORM"].dropna().unique()

    if len(fechas_doc) == 0:
        raise ValueError("No se encontraron fechas de documento válidas.")

    hora_str = datetime.now().strftime("%H_%M_%S")
    archivos_generados = []

    for fecha_doc in sorted(fechas_doc):
        df_grupo = df[df["FECHA_DOC_NORM"] == fecha_doc].copy().reset_index(drop=True)
        if df_grupo.empty:
            continue

        fecha_str = pd.Timestamp(fecha_doc).strftime("%d_%m_%Y")
        nombre    = f"COMPENSACION_{fecha_str}_{hora_str}.xlsx"
        buffer    = _generar_excel_compensacion(df_grupo)
        archivos_generados.append({
            "nombre": nombre,
            "buffer": buffer,
            "fecha":  fecha_str
        })

    if not archivos_generados:
        return "No se generaron compensaciones."

    # ZIP con todos
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for arch in archivos_generados:
            arch["buffer"].seek(0)
            zf.writestr(arch["nombre"], arch["buffer"].read())
    zip_buffer.seek(0)

    hora_zip  = datetime.now().strftime("%H_%M_%S")
    zip_nombre = f"COMPENSACIONES_{hora_zip}.zip"

    st.markdown("#### 📥 Descargar compensaciones generadas:")
    st.download_button(
        label=f"⬇️  Descargar todas las compensaciones ({len(archivos_generados)} archivos)",
        data=zip_buffer,
        file_name=zip_nombre,
        mime="application/zip",
        key="dl_todos_comp",
        type="primary",
        use_container_width=True
    )

    with st.expander("📂 Descargar individualmente"):
        for arch in archivos_generados:
            arch["buffer"].seek(0)
            st.download_button(
                label=f"⬇️  Compensación {arch['fecha']}",
                data=arch["buffer"],
                file_name=arch["nombre"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_comp_{arch['fecha']}"
            )

    return f"{len(archivos_generados)} compensaciones generadas correctamente."


def _generar_excel_compensacion(df):
    """
    Genera el archivo Items con DÉBITOS seguidos de CRÉDITOS.
    Sin filas en blanco entre ellos.
    Consecutivo global no se reinicia.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Items"

    encabezados = [
        "Id", "codigoCentroCosto", "dniTercero", "codigoTipoDniTercero",
        "codigoCuenta", "valor", "factura", "fechaVencimiento",
        "codigoImpuesto", "valorBaseImpuesto", "porcentajeImpuesto", "detalle"
    ]

    # Encabezado con formato
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    for col_idx, col_name in enumerate(encabezados, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center")

    # ── Filtrar solo filas con cédula válida ────────────────────────────
    filas_validas = []
    for _, row in df.iterrows():
        cedula = str(row.get("CEDULA", "")).strip()
        if cedula and cedula != "nan":
            filas_validas.append(row)

    consecutivo = 1
    fila_excel  = 2

    # ── DÉBITOS ─────────────────────────────────────────────────────────
    for row in filas_validas:
        entidad     = str(row.get("ENTIDAD", "")).upper().strip()
        valor       = _num(row.get("VALOR"))
        num_factura = row.get("NUM_FACTURA", "")
        fecha_venc  = _fecha_dt(row.get("FECHA"))   # fecha del libro de banco

        # NIT del banco
        nit_banco = ""
        for clave, nit in NIT_ENTIDAD.items():
            if clave in entidad:
                nit_banco = nit
                break

        ws.cell(row=fila_excel, column=1,  value=consecutivo)
        ws.cell(row=fila_excel, column=2,  value=CODIGO_CENTRO_COSTO if nit_banco else "")
        ws.cell(row=fila_excel, column=3,  value=nit_banco)
        ws.cell(row=fila_excel, column=4,  value="NIT" if nit_banco else "")
        ws.cell(row=fila_excel, column=5,  value=CODIGO_CUENTA_DEBITO if nit_banco else "")
        ws.cell(row=fila_excel, column=6,  value=valor)
        ws.cell(row=fila_excel, column=7,  value=_factura(num_factura))

        # Fecha con formato
        if fecha_venc:
            cell_f = ws.cell(row=fila_excel, column=8, value=fecha_venc)
            cell_f.number_format = "DD/MM/YYYY"
        else:
            ws.cell(row=fila_excel, column=8, value="")

        consecutivo += 1
        fila_excel  += 1

    # ── CRÉDITOS (inmediatamente después, sin filas en blanco) ──────────
    for row in filas_validas:
        cedula  = str(row.get("CEDULA", "")).strip()
        valor   = _num(row.get("VALOR"))
        nit_banco = ""
        entidad = str(row.get("ENTIDAD", "")).upper().strip()
        for clave, nit in NIT_ENTIDAD.items():
            if clave in entidad:
                nit_banco = nit
                break

        ws.cell(row=fila_excel, column=1,  value=consecutivo)
        ws.cell(row=fila_excel, column=2,  value=CODIGO_CENTRO_COSTO if cedula else "")
        ws.cell(row=fila_excel, column=3,  value=cedula)
        ws.cell(row=fila_excel, column=4,  value="CC" if cedula else "")
        ws.cell(row=fila_excel, column=5,  value=CODIGO_CUENTA_CREDITO if cedula else "")
        ws.cell(row=fila_excel, column=6,  value=-abs(valor) if valor else "")
        ws.cell(row=fila_excel, column=7,  value="")   # factura vacía en crédito
        ws.cell(row=fila_excel, column=8,  value="")   # fecha vacía en crédito

        consecutivo += 1
        fila_excel  += 1

    # Ajustar anchos
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def _factura(val):
    """Formatea factura sin decimales .0"""
    if not val or str(val).strip() == "" or str(val) == "nan":
        return ""
    try:
        # Si es número entero como 66722.0 → "66722"
        f = float(str(val))
        if f == int(f):
            return str(int(f))
        return str(val)
    except Exception:
        return str(val).strip()


def _num(val):
    try:
        return float(str(val).replace(",", "") or 0)
    except Exception:
        return 0.0


def _fecha_dt(val):
    """Retorna datetime para formato Excel."""
    if val is None:
        return None
    try:
        return pd.Timestamp(val).to_pydatetime()
    except Exception:
        return None
