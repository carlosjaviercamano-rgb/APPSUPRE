"""
generar_compensaciones.py
Genera archivos de compensación por fecha única en AREA DE BANCO.
Basado en la macro CREAR_COMPENSACIONES_PRO del workbook original.
"""
import pandas as pd
import openpyxl
from datetime import datetime
import os


def crear_compensaciones(df_area_banco, config):
    """
    Agrupa AREA DE BANCO por fecha única y genera un archivo por fecha.
    """
    if df_area_banco is None or df_area_banco.empty:
        raise ValueError("No hay datos en Área de Banco para generar compensaciones.")

    ruta_comp = config.get("ruta_compensaciones", "")
    if not ruta_comp:
        raise ValueError("Configura la ruta de compensaciones en ⚙️ Configuración.")

    df = df_area_banco.copy()
    df["FECHA_NORM"] = pd.to_datetime(df["FECHA"], errors="coerce").dt.normalize()
    fechas_unicas = df["FECHA_NORM"].dropna().unique()

    if len(fechas_unicas) == 0:
        raise ValueError("No se encontraron fechas válidas en Área de Banco.")

    archivos_creados = []
    hora_str = datetime.now().strftime("%H-%M-%S")

    for fecha in sorted(fechas_unicas):
        df_fecha = df[df["FECHA_NORM"] == fecha].copy()
        if df_fecha.empty:
            continue

        fecha_str = pd.Timestamp(fecha).strftime("%d-%m-%Y")
        nombre    = f"COMPENSACION_{fecha_str}_{hora_str}.xlsx"
        ruta_completa = os.path.join(ruta_comp, nombre)

        _generar_excel_compensacion(df_fecha, ruta_completa, fecha)
        archivos_creados.append(nombre)

    return f"{len(archivos_creados)} compensaciones generadas: {', '.join(archivos_creados)}"


def _generar_excel_compensacion(df, ruta, fecha):
    """
    Genera el archivo Excel de compensación para una fecha.
    """
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
        ws.append([fila.get(c) for c in encabezados if c in df.columns or None])

    # Totales
    ws.append([])
    ws.append(["TOTAL", "", "", df["VALOR"].sum() if "VALOR" in df.columns else 0])

    os.makedirs(os.path.dirname(ruta), exist_ok=True) if os.path.dirname(ruta) else None
    wb.save(ruta)
