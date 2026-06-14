"""
generar_planos.py
Genera los archivos de planos por empresa (CashReceipt, Services, PaymentMethod).
Basado en la macro CREAR_PLANOS_2026_PRO del workbook original.
"""
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import os


EMPRESAS = {
    "Movicap":       "ruta_movicap",
    "Suprecartera":  "ruta_suprecartera",
    "Suprecreditos": "ruta_suprecredito",
    "TuCredito":     "ruta_tucredito",
}


def crear_planos(df_sheet1, config, df_area_banco, tipo_pago):
    """
    Filtra por empresa, genera los archivos de planos y los guarda.
    Nombre del archivo: TIPO_PAGO + FECHA_DOCUMENTO + HORA
    """
    if df_sheet1 is None or df_sheet1.empty:
        raise ValueError("No hay datos en Sheet1 para generar planos.")

    # Fecha documento = más reciente en AREA DE BANCO columna FECHA
    fecha_doc = None
    if df_area_banco is not None and "FECHA" in df_area_banco.columns:
        fechas = pd.to_datetime(df_area_banco["FECHA"], errors="coerce").dropna()
        if not fechas.empty:
            fecha_doc = fechas.max()

    fecha_str = fecha_doc.strftime("%d-%m-%Y") if fecha_doc else datetime.now().strftime("%d-%m-%Y")
    hora_str  = datetime.now().strftime("%H-%M-%S")
    tipo_str  = "PAGOS_BANCARIOS" if tipo_pago == "bancarios" else "PAGOS_RECAUDOS"

    archivos_creados = []

    for empresa, ruta_key in EMPRESAS.items():
        ruta = config.get(ruta_key, "")
        if not ruta:
            continue

        # Filtrar por empresa
        df_emp = df_sheet1[
            df_sheet1["COMPANY"].astype(str).str.upper().str.contains(empresa.upper(), na=False)
        ].copy()

        if df_emp.empty:
            continue

        # Nombre del archivo
        nombre = f"{tipo_str}_{empresa}_{fecha_str}_{hora_str}.xlsx"
        ruta_completa = os.path.join(ruta, nombre)

        # Crear el archivo
        _generar_excel_plano(df_emp, ruta_completa, empresa, fecha_doc)
        archivos_creados.append(nombre)

    if not archivos_creados:
        return "No se encontraron datos para ninguna empresa. Verifica que COMPANY esté correctamente cruzado."

    return f"Planos generados: {', '.join(archivos_creados)}"


def _generar_excel_plano(df, ruta, empresa, fecha_doc):
    """
    Genera el archivo Excel de plano con las hojas CashReceipt, Services, PaymentMethod.
    Mantiene la estructura y formatos del workbook original.
    """
    wb = openpyxl.Workbook()

    # ── Hoja CashReceipt ────────────────────────────────────────────────
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

    # ── Hoja Services ───────────────────────────────────────────────────
    ws_sv = wb.create_sheet("Services")
    encabezados_sv = ["COMPANY", "IDEN", "NUM_FACTURA", "CUOTA", "FECHA"]
    ws_sv.append(encabezados_sv)
    for _, fila in df.iterrows():
        ws_sv.append([fila.get("COMPANY"), fila.get("IDEN"),
                      fila.get("NUM_FACTURA"), fila.get("CUOTA"), fila.get("FECHA")])

    # ── Hoja PaymentMethod ──────────────────────────────────────────────
    ws_pm = wb.create_sheet("PaymentMethod")
    encabezados_pm = ["COMPANY", "IDEN", "RECIBO", "VALOR_CB", "FECHA"]
    ws_pm.append(encabezados_pm)
    for _, fila in df.iterrows():
        ws_pm.append([fila.get("COMPANY"), fila.get("IDEN"),
                      fila.get("RECIBO"), fila.get("VALOR_CB"), fila.get("FECHA")])

    os.makedirs(os.path.dirname(ruta), exist_ok=True) if os.path.dirname(ruta) else None
    wb.save(ruta)
