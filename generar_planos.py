import pandas as pd
import openpyxl
from datetime import datetime
import io
import streamlit as st

EMPRESAS = {
    "Movicap":       "ruta_movicap",
    "Suprecartera":  "ruta_suprecartera",
    "Suprecreditos": "ruta_suprecredito",
    "TuCredito":     "ruta_tucredito",
}

def crear_planos(df_sheet1, config, df_area_banco, tipo_pago):
    if df_sheet1 is None or df_sheet1.empty:
        raise ValueError("No hay datos en Sheet1 para generar planos.")

    # Fecha documento
    fecha_doc = None
    if df_area_banco is not None and "FECHA" in df_area_banco.columns:
        fechas = pd.to_datetime(df_area_banco["FECHA"], errors="coerce").dropna()
        if not fechas.empty:
            fecha_doc = fechas.max()

    fecha_str = fecha_doc.strftime("%d_%m_%Y") if fecha_doc else datetime.now().strftime("%d_%m_%Y")
    hora_str  = datetime.now().strftime("%H_%M_%S")
    tipo_str  = "PAGOS_BANCARIOS" if tipo_pago == "bancarios" else "PAGOS_RECAUDOS"

    archivos_generados = []

    for empresa, ruta_key in EMPRESAS.items():
        df_emp = df_sheet1[
            df_sheet1["COMPANY"].astype(str).str.upper().str.contains(empresa.upper(), na=False)
        ].copy()

        if df_emp.empty:
            continue

        nombre = f"{tipo_str}_{empresa}_{fecha_str}_{hora_str}.xlsx"
        buffer = _generar_excel_plano(df_emp, empresa, fecha_doc)
        archivos_generados.append({"nombre": nombre, "buffer": buffer, "empresa": empresa})

    if not archivos_generados:
        return "No se encontraron datos para ninguna empresa."

    # Mostrar botones de descarga
    st.markdown("#### 📥 Descargar planos generados:")
    for arch in archivos_generados:
        st.download_button(
            label=f"⬇️ Descargar {arch['empresa']}",
            data=arch["buffer"],
            file_name=arch["nombre"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{arch['empresa']}"
        )

    return f"{len(archivos_generados)} planos generados correctamente."


def _generar_excel_plano(df, empresa, fecha_doc):
    wb = openpyxl.Workbook()

    # Hoja CashReceipt
    ws_cr = wb.active
    ws_cr.title = "CashReceipt"
    encabezados_cr = [
        "ENTIDAD", "FECHA", "COMPANY", "IDEN", "NUM_FACTURA",
        "CUOTA", "RECIBO", "DIFERENCIA", "VALOR_CB",
        "INMOVILIZACION", "OTROS_GASTOS", "OBSERVACION",
        "CORRESPONSAL", "FECHA_DOCUMENTO", "DETALLE"
    ]
    ws_cr.append(encabezados_cr)
    for _, fila in df.iterrows():
        ws_cr.append([fila.get(c) for c in encabezados_cr])

    # Hoja Services
    ws_sv = wb.create_sheet("Services")
    ws_sv.append(["COMPANY", "IDEN", "NUM_FACTURA", "CUOTA", "FECHA"])
    for _, fila in df.iterrows():
        ws_sv.append([fila.get("COMPANY"), fila.get("IDEN"),
                      fila.get("NUM_FACTURA"), fila.get("CUOTA"), fila.get("FECHA")])

    # Hoja PaymentMethod
    ws_pm = wb.create_sheet("PaymentMethod")
    ws_pm.append(["COMPANY", "IDEN", "RECIBO", "VALOR_CB", "FECHA"])
    for _, fila in df.iterrows():
        ws_pm.append([fila.get("COMPANY"), fila.get("IDEN"),
                      fila.get("RECIBO"), fila.get("VALOR_CB"), fila.get("FECHA")])

    # Guardar en buffer de memoria
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
