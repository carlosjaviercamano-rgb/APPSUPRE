import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import io
import streamlit as st

EMPRESAS = {
    "Movicap":       "ruta_movicap",
    "Suprecartera":  "ruta_suprecartera",
    "Suprecreditos": "ruta_suprecredito",
    "TuCredito":     "ruta_tucredito",
}

# Código de método de pago por empresa
METODO_PAGO = {
    "Movicap":       "0041",
    "TuCredito":     "0041",
    "Suprecreditos": "0040",
    "Suprecartera":  "0040",
}


def crear_planos(df_sheet1, config, df_area_banco, tipo_pago):
    if df_sheet1 is None or df_sheet1.empty:
        raise ValueError("No hay datos en Sheet1 para generar planos.")

    # Fecha documento = más reciente en AREA DE BANCO columna FECHA
    fecha_doc = None
    if df_area_banco is not None and "FECHA" in df_area_banco.columns:
        fechas = pd.to_datetime(df_area_banco["FECHA"], errors="coerce").dropna()
        if not fechas.empty:
            fecha_doc = fechas.max()

    fecha_str = fecha_doc.strftime("%d_%m_%Y") if fecha_doc else datetime.now().strftime("%d_%m_%Y")
    hora_str  = datetime.now().strftime("%H_%M_%S")
    tipo_str  = "PAGOS_BANCARIOS" if tipo_pago == "bancarios" else "PAGOS_RECAUDOS"

    archivos_generados = []

    for empresa in EMPRESAS:
        # Filtrar por empresa (columna COMPANY)
        df_emp = df_sheet1[
            df_sheet1["COMPANY"].astype(str).str.strip().str.upper() == empresa.upper()
        ].copy().reset_index(drop=True)

        if df_emp.empty:
            continue

        # Construir las 3 hojas
        df_cash     = _construir_cash_receipt(df_emp, empresa)
        df_services = _construir_services(df_emp, df_cash)
        df_payment  = _construir_payment_method(df_emp, df_cash, empresa)

        nombre = f"{tipo_str}_{empresa}_{fecha_str}_{hora_str}.xlsx"
        buffer = _generar_excel(df_cash, df_services, df_payment)
        archivos_generados.append({
            "nombre":  nombre,
            "buffer":  buffer,
            "empresa": empresa
        })

    if not archivos_generados:
        return "No se encontraron datos para ninguna empresa."

    # Botones de descarga
    st.markdown("#### 📥 Descargar planos generados:")
    for arch in archivos_generados:
        st.download_button(
            label=f"⬇️  Descargar {arch['empresa']}",
            data=arch["buffer"],
            file_name=arch["nombre"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{arch['empresa']}"
        )

    return f"{len(archivos_generados)} planos generados correctamente."


# ══════════════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DE HOJAS
# ══════════════════════════════════════════════════════════════════════════

def _construir_cash_receipt(df, empresa):
    """
    Replica la lógica de CashReceipt_demo:
    A=id(1), B="DC", C=104, D=fechaDocumento, E=fechaPago,
    F=detalle, G="CC", H=dni, I=factura, J=valorPago, K=1, L=1, M=1
    """
    rows = []
    consecutivo = 1
    for idx, row in df.iterrows():
        cedula = row.get("IDEN")
        if not cedula or str(cedula).strip() == "" or str(cedula) == "nan":
            continue

        cuota      = _num(row.get("CUOTA"))
        diferencia = _num(row.get("DIFERENCIA"))
        valor_cb   = _num(row.get("VALOR_CB"))

        # Si tiene corresponsal (valor_cb > 0) → valorPago = DIFERENCIA
        # Si no tiene corresponsal → valorPago = CUOTA
        valor_pago = diferencia if valor_cb > 0 and diferencia > 0 else cuota

        rows.append({
            "id":                          consecutivo,
            "codigoTipoDocumento":         "DC",
            "codigoCentroCosto":           104,
            "fechaDocumento":              _fecha(row.get("FECHA_DOCUMENTO")),
            "fechaPago":                   _fecha(row.get("FECHA")),
            "detalle":                     row.get("DETALLE", ""),
            "codigoTipoDni":               "CC",
            "dni":                         str(cedula).strip(),
            "factura":                     row.get("NUM_FACTURA", ""),
            "valorPago":                   valor_pago,
            "aplicaInteresMoratorio":      1,
            "aplicaDescuentoProntoPago":   1,
            "aplicaGestionCobranza":       1,
        })
        consecutivo += 1

    return pd.DataFrame(rows)


def _construir_services(df, df_cash):
    """
    Replica la lógica de Services_demo:
    A=idDocumento, B=codigoServicio(según valor_cb/inmov/otros),
    C=valor, D=factura, F=dni, G="CC"
    """
    rows = []
    for (idx, row_emp), (_, row_cash) in zip(df.iterrows(), df_cash.iterrows()):
        cedula = row_cash.get("dni")
        if not cedula:
            continue

        valor_cb      = _num(row_emp.get("VALOR_CB"))
        inmovilizacion = _num(row_emp.get("INMOVILIZACION"))
        otros_gastos   = _num(row_emp.get("OTROS_GASTOS"))
        cuota          = _num(row_emp.get("CUOTA"))

        # Solo agrega fila si hay valor en VALOR_CB, INMOVILIZACION u OTROS_GASTOS
        if valor_cb > 0:
            codigo_servicio = 586325
            valor_servicio  = valor_cb   # VALOR_CB = $10.000 corresponsal
        elif inmovilizacion > 0:
            codigo_servicio = 19051
            valor_servicio  = inmovilizacion
        elif otros_gastos > 0:
            codigo_servicio = 11111
            valor_servicio  = otros_gastos
        else:
            continue  # Sin valores especiales → no aparece en Services

        rows.append({
            "idDocumento":      row_cash.get("id", 1),
            "codigoServicio":   codigo_servicio,
            "valor":            valor_servicio,
            "factura":          row_emp.get("NUM_FACTURA", ""),
            "fechaVencimiento": None,
            "dni":              cedula,
            "codigoTipoDni":    "CC",
        })

    return pd.DataFrame(rows)


def _construir_payment_method(df, df_cash, empresa):
    """
    Replica la lógica de PaymentMethod_demo:
    A=idDocumento, B=codigoMetodoPago(según empresa),
    C=valor(cuota), E=dni, F="CC"
    """
    metodo = METODO_PAGO.get(empresa, "0040")
    rows = []
    for (idx, row_emp), (_, row_cash) in zip(df.iterrows(), df_cash.iterrows()):
        cedula = row_cash.get("dni")
        if not cedula:
            continue

        rows.append({
            "idDocumento":      row_cash.get("id", 1),
            "codigoMetodoPago": metodo,
            "valor":            _num(row_emp.get("CUOTA")),  # Siempre el valor total
            "voucher":          None,
            "dni":              cedula,
            "codigoTipoDni":    "CC",
        })

    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════════════════
# GENERAR EXCEL CON FORMATO
# ══════════════════════════════════════════════════════════════════════════

def _generar_excel(df_cash, df_services, df_payment):
    wb = openpyxl.Workbook()

    _escribir_hoja(wb.active, "CashReceipt", df_cash)
    _escribir_hoja(wb.create_sheet("Services"), "Services", df_services)
    _escribir_hoja(wb.create_sheet("PaymentMethod"), "PaymentMethod", df_payment)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def _escribir_hoja(ws, titulo, df):
    ws.title = titulo

    # Estilo encabezado
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_align = Alignment(horizontal="center", vertical="center")

    # Escribir encabezados
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill   = header_fill
        cell.font   = header_font
        cell.alignment = header_align

    # Columnas que deben tener formato fecha
    cols_fecha = [col for col in df.columns if "fecha" in col.lower()]
    idx_fechas = [list(df.columns).index(c) + 1 for c in cols_fecha]

    # Escribir datos
    for row_idx, row in df.iterrows():
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=value)
            if col_idx in idx_fechas and value:
                try:
                    cell.value = pd.Timestamp(value).to_pydatetime()
                    cell.number_format = "DD/MM/YYYY"
                except Exception:
                    pass

    # Ajustar ancho de columnas
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)


# ══════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════════════════

def _num(val):
    """Convierte a float seguro."""
    try:
        return float(str(val).replace(",", "") or 0)
    except Exception:
        return 0.0


def _fecha(val):
    """Convierte a datetime para que Excel aplique formato de fecha."""
    if val is None:
        return None
    try:
        return pd.Timestamp(val).to_pydatetime()
    except Exception:
        return None
