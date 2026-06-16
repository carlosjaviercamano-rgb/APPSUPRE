import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import io
import zipfile
import streamlit as st

# NIT por entidad bancaria
NIT_ENTIDAD = {
    "BANCOLOMBIA":  "890903938",
    "DAVIVIENDA":   "860034313",
    "PSE":          "830078512",
    "EFECTY":       "830131993",
    "RECORD":       "111111111",
    "OCCIDENTE":    "860002395",
}

# Cuenta contable por entidad (Contra Banco)
CUENTA_CONTRA_BANCO = {
    "890903938": "112005001",
    "860034313": "111005005",
}

CODIGO_CENTRO_COSTO = "102"
CODIGO_CUENTA_PENDIENTE = "141299011"


def crear_compensaciones(df_area_banco, config):
    if df_area_banco is None or df_area_banco.empty:
        raise ValueError("No hay datos en Área de Banco para generar compensaciones.")

    df = df_area_banco.copy()
    df["FECHA_NORM"] = pd.to_datetime(df["FECHA"], errors="coerce").dt.normalize()
    fechas_unicas = df["FECHA_NORM"].dropna().unique()

    if len(fechas_unicas) == 0:
        raise ValueError("No se encontraron fechas válidas en Área de Banco.")

    hora_str = datetime.now().strftime("%H_%M_%S")
    archivos_generados = []

    for fecha in sorted(fechas_unicas):
        df_fecha = df[df["FECHA_NORM"] == fecha].copy().reset_index(drop=True)
        if df_fecha.empty:
            continue

        fecha_str = pd.Timestamp(fecha).strftime("%d_%m_%Y")
        nombre    = f"COMPENSACION_{fecha_str}_{hora_str}.xlsx"
        buffer    = _generar_excel_compensacion(df_fecha, fecha)
        archivos_generados.append({
            "nombre": nombre,
            "buffer": buffer,
            "fecha":  fecha_str
        })

    if not archivos_generados:
        return "No se generaron compensaciones."

    # ZIP con todos los archivos
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for arch in archivos_generados:
            arch["buffer"].seek(0)
            zf.writestr(arch["nombre"], arch["buffer"].read())
    zip_buffer.seek(0)

    zip_nombre = f"COMPENSACIONES_{hora_str}.zip"

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


def _generar_excel_compensacion(df, fecha):
    """
    Genera el archivo de compensación para una fecha.
    Replica la lógica de Items_DEMO con modo 'Contra Pendiente'.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Items"

    encabezados = [
        "Id", "codigoCentroCosto", "dniTercero", "codigoTipoDniTercero",
        "codigoCuenta", "valor", "factura", "fechaVencimiento",
        "codigoImpuesto", "valorBaseImpuesto", "porcentajeImpuesto", "detalle"
    ]

    # Estilo encabezado
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=10)

    for col_idx, col_name in enumerate(encabezados, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center")

    # Construir filas según lógica Items_DEMO (modo Contra Pendiente)
    consecutivo = 1
    for _, row in df.iterrows():
        cedula = row.get("CEDULA")
        if not cedula or str(cedula).strip() == "" or str(cedula) == "nan":
            continue

        entidad     = str(row.get("ENTIDAD", "")).upper().strip()
        valor       = _num(row.get("VALOR"))
        num_factura = row.get("NUM_FACTURA", "")
        fecha_venc  = _fecha(row.get("FECHA"))

        # NIT del banco según entidad
        nit_tercero = ""
        for clave, nit in NIT_ENTIDAD.items():
            if clave in entidad:
                nit_tercero = nit
                break

        # Código cuenta: Contra Pendiente = 141299011
        codigo_cuenta = CODIGO_CUENTA_PENDIENTE if nit_tercero else ""

        fila = [
            consecutivo,
            CODIGO_CENTRO_COSTO,
            nit_tercero,
            "NIT" if nit_tercero else "",
            codigo_cuenta,
            valor,
            str(num_factura) if num_factura else "",
            fecha_venc,
            None,   # codigoImpuesto
            None,   # valorBaseImpuesto
            None,   # porcentajeImpuesto
            None,   # detalle
        ]
        ws.append(fila)

        # Formato fecha en col H
        cell_fecha = ws.cell(row=consecutivo + 1, column=8)
        if fecha_venc:
            try:
                cell_fecha.value         = pd.Timestamp(fecha_venc).to_pydatetime()
                cell_fecha.number_format = "DD/MM/YYYY"
            except Exception:
                pass

        consecutivo += 1

    # Total al final
    ws.append([])
    total = sum(_num(r.get("VALOR")) for _, r in df.iterrows())
    ws.append(["TOTAL", "", "", "", "", total])

    # Ajustar anchos
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def _num(val):
    try:
        return float(str(val).replace(",", "") or 0)
    except Exception:
        return 0.0


def _fecha(val):
    if val is None:
        return None
    try:
        return pd.Timestamp(val).to_pydatetime()
    except Exception:
        return None
