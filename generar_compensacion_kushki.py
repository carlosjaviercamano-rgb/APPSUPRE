import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import io
import os
import streamlit as st

# ── Constantes fijas de Kushki ───────────────────────────────────────────
DNI_TERCERO       = "901000330"
TIPO_TERCERO      = "NIT"
CENTRO_COSTO      = "102"

CUENTA_BOLSA      = "138095011"   # débito: total recaudado - comisión - iva
CUENTA_COMISION   = "530515001"   # débito: comisión (1250 por movimiento)
CUENTA_IVA        = "240805002"   # débito: iva de la comisión (19%)
CUENTA_CREDITO    = "110505017"   # crédito: uno a uno por ticket

VALOR_COMISION_UNITARIA = 1250
PORCENTAJE_IVA_COMISION = 0.19

DETALLE_DEBITO = "COMPENSACION PAGO PAGINA KUSHKI"


def _fecha_str(fecha):
    return pd.Timestamp(fecha).strftime("%d-%m-%Y")


def crear_compensacion_kushki(archivo_excel, config=None):
    """
    Genera un archivo de compensación (hoja Items) por cada fecha distinta
    encontrada en la columna 'created' del reporte de Kushki. Solo se
    compensa (no hay aplicación/planos), similar a PSE/Efecty/Record.
    """
    df = pd.read_excel(archivo_excel)
    df.columns = [str(c).strip() for c in df.columns]

    requeridas = ["ticket_number", "created", "approved_transaction_amount", "document_number"]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en el archivo: {', '.join(faltantes)}")

    df["created_dt"] = pd.to_datetime(df["created"], errors="coerce", utc=False)
    if df["created_dt"].isna().any():
        n_malas = int(df["created_dt"].isna().sum())
        st.warning(f"⚠️ {n_malas} fila(s) con fecha inválida en 'created' — se excluyeron.")
        df = df[df["created_dt"].notna()].copy()

    if df.empty:
        raise ValueError("No hay filas válidas para procesar.")

    df["fecha_grupo"] = df["created_dt"].dt.date
    df["approved_transaction_amount"] = pd.to_numeric(
        df["approved_transaction_amount"], errors="coerce"
    ).fillna(0)

    hora_str = datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
    archivos_generados = []

    for fecha_grupo, df_dia in df.groupby("fecha_grupo"):
        n_movimientos = len(df_dia)
        total_dia      = df_dia["approved_transaction_amount"].sum()
        comision_total = VALOR_COMISION_UNITARIA * n_movimientos
        iva_comision   = round(comision_total * PORCENTAJE_IVA_COMISION, 2)
        valor_bolsa    = round(total_dia - comision_total - iva_comision, 2)

        fecha_dt  = pd.Timestamp(fecha_grupo)
        fecha_txt = _fecha_str(fecha_dt)

        filas_debito = [
            {
                "codigoCentroCosto": CENTRO_COSTO, "dniTercero": DNI_TERCERO,
                "codigoTipoDniTercero": TIPO_TERCERO, "codigoCuenta": CUENTA_BOLSA,
                "valor": valor_bolsa, "factura": fecha_txt, "fechaVencimiento": fecha_dt,
                "codigoImpuesto": None, "valorBaseImpuesto": None, "porcentajeImpuesto": None,
                "detalle": DETALLE_DEBITO,
            },
            {
                "codigoCentroCosto": CENTRO_COSTO, "dniTercero": DNI_TERCERO,
                "codigoTipoDniTercero": TIPO_TERCERO, "codigoCuenta": CUENTA_COMISION,
                "valor": comision_total, "factura": "", "fechaVencimiento": None,
                "codigoImpuesto": None, "valorBaseImpuesto": None, "porcentajeImpuesto": None,
                "detalle": DETALLE_DEBITO,
            },
            {
                "codigoCentroCosto": CENTRO_COSTO, "dniTercero": DNI_TERCERO,
                "codigoTipoDniTercero": TIPO_TERCERO, "codigoCuenta": CUENTA_IVA,
                "valor": iva_comision, "factura": "", "fechaVencimiento": None,
                "codigoImpuesto": "01", "valorBaseImpuesto": comision_total,
                "porcentajeImpuesto": PORCENTAJE_IVA_COMISION,
                "detalle": DETALLE_DEBITO,
            },
        ]

        filas_credito = []
        for _, row in df_dia.iterrows():
            dni_credito = str(row["document_number"]).strip()
            if dni_credito.endswith(".0"):
                dni_credito = dni_credito[:-2]
            ticket = str(row["ticket_number"]).strip()
            if ticket.endswith(".0"):
                ticket = ticket[:-2]
            filas_credito.append({
                "codigoCentroCosto": CENTRO_COSTO, "dniTercero": dni_credito,
                "codigoTipoDniTercero": "CC", "codigoCuenta": CUENTA_CREDITO,
                "valor": -abs(row["approved_transaction_amount"]),
                "factura": "", "fechaVencimiento": None,
                "codigoImpuesto": None, "valorBaseImpuesto": None, "porcentajeImpuesto": None,
                "detalle": f"{DETALLE_DEBITO} {ticket}",
            })

        todas_las_filas = filas_debito + filas_credito
        for idx, fila in enumerate(todas_las_filas, start=1):
            fila["Id"] = idx

        nombre = f"COMPENSACION_KUSHKI_{fecha_dt.strftime('%d_%m_%Y')}_{hora_str}.xlsx"
        buffer = _generar_excel(todas_las_filas)

        ruta_auto = config.get("ruta_compensaciones", "") if config else ""
        if ruta_auto:
            try:
                os.makedirs(ruta_auto, exist_ok=True)
                ruta_completa = os.path.join(ruta_auto, nombre)
                buffer.seek(0)
                with open(ruta_completa, "wb") as f:
                    f.write(buffer.read())
                st.success(f"💾 {nombre}: guardado en {ruta_completa}")
            except Exception as e:
                st.warning(f"⚠️ {nombre}: no se pudo guardar automáticamente — {str(e)}")
            buffer.seek(0)

        archivos_generados.append({"nombre": nombre, "buffer": buffer})

    st.markdown("#### 📥 Descargar compensación(es) generada(s):")
    for arch in archivos_generados:
        arch["buffer"].seek(0)
        st.download_button(
            label=f"⬇️  Descargar {arch['nombre']}",
            data=arch["buffer"],
            file_name=arch["nombre"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_kushki_{arch['nombre']}",
        )

    return f"{len(df)} movimiento(s) procesados en {len(archivos_generados)} archivo(s) (uno por fecha)."


def _generar_excel(filas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Items"

    encabezados = [
        "Id", "codigoCentroCosto", "dniTercero", "codigoTipoDniTercero",
        "codigoCuenta", "valor", "factura", "fechaVencimiento",
        "codigoImpuesto", "valorBaseImpuesto", "porcentajeImpuesto", "detalle"
    ]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    for col_idx, col_name in enumerate(encabezados, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for fila_idx, fila in enumerate(filas, start=2):
        for col_idx, col_name in enumerate(encabezados, start=1):
            ws.cell(row=fila_idx, column=col_idx, value=fila.get(col_name))
        if fila.get("fechaVencimiento"):
            ws.cell(row=fila_idx, column=8).number_format = "DD/MM/YYYY"

    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
