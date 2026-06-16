import pandas as pd
import openpyxl
from datetime import datetime
import io
import zipfile
import streamlit as st


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
        df_fecha = df[df["FECHA_NORM"] == fecha].copy()
        if df_fecha.empty:
            continue

        fecha_str = pd.Timestamp(fecha).strftime("%d_%m_%Y")
        nombre    = f"COMPENSACION_{fecha_str}_{hora_str}.xlsx"
        buffer    = _generar_excel_compensacion(df_fecha, fecha)
        archivos_generados.append({"nombre": nombre, "buffer": buffer, "fecha": fecha_str})

    if not archivos_generados:
        return "No se generaron compensaciones."

    # Mostrar botones de descarga
    st.markdown("#### 📥 Descargar compensaciones generadas:")
    for arch in archivos_generados:
        st.download_button(
            label=f"⬇️ Descargar compensación {arch['fecha']}",
            data=arch["buffer"],
            file_name=arch["nombre"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_comp_{arch['fecha']}"
        )

    return f"{len(archivos_generados)} compensaciones generadas correctamente."


def _generar_excel_compensacion(df, fecha):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compensacion"

    encabezados = [
        "ENTIDAD", "FECHA", "CEDULA", "VALOR", "T_TRANSACCION",
        "NUM_FACTURA", "CUOTA", "RECIBO", "DIFERENCIA",
        "VALOR_CB", "INMOVILIZACION", "OTROS_GASTOS",
        "OBSERVACION", "CORRESPONSAL", "FECHA_DOCUMENTO"
    ]
    ws.append(encabezados)
    for _, fila in df.iterrows():
        ws.append([fila.get(c) for c in encabezados if c in df.columns])

    # Total
    ws.append([])
    total = pd.to_numeric(df["VALOR"], errors="coerce").sum() if "VALOR" in df.columns else 0
    ws.append(["TOTAL", "", "", total])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
