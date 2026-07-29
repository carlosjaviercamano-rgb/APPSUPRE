import openpyxl
from datetime import datetime
import os
import streamlit as st

from generar_gastos_bancarios import (
    NIT_BANCOS,
    TIPO_TERCERO_POR_BANCO,
    obtener_centro_costo,
    obtener_cuenta_credito,
    _num,
    _parsear_fecha_mixta,
    _fecha_str,
    _generar_excel,
)

# Cuenta fija de crédito para rendimientos financieros (intereses de ahorro)
CUENTA_RENDIMIENTO_FINANCIERO = "421005001"
LIMITE_MOVIMIENTOS = 480


def crear_rendimiento_financiero(df_tabla, empresa, banco, config=None):
    """
    Genera el archivo de Rendimiento Financiero (hoja Items, débitos y
    créditos) a partir de la tabla pegada por el usuario
    (CONCEPTO, FECHA, VALOR).

    A diferencia de Gastos Bancarios, aquí no hay clasificación por
    concepto: el débito siempre es la cuenta del banco (empresa+banco,
    igual que en Gastos Bancarios) y el crédito siempre es la cuenta
    fija 421005001.
    """
    if df_tabla is None or df_tabla.empty:
        raise ValueError("No hay filas en la tabla para generar el archivo.")

    nit_banco = NIT_BANCOS.get(banco.upper())
    if not nit_banco:
        raise ValueError(f"No se encontró el NIT configurado para el banco '{banco}'.")
    tipo_tercero = TIPO_TERCERO_POR_BANCO.get(banco.upper(), "NIT")
    centro_costo = obtener_centro_costo(empresa)

    cuenta_debito = obtener_cuenta_credito(empresa, banco)
    if not cuenta_debito:
        raise ValueError(
            f"No se encontró una cuenta de '{banco}' configurada para la empresa '{empresa}'."
        )

    df = df_tabla.copy()
    df = df[
        df["FECHA"].astype(str).str.strip().ne("") |
        df["VALOR"].astype(str).str.strip().ne("") |
        df["CONCEPTO"].astype(str).str.strip().ne("")
    ].reset_index(drop=True)
    if df.empty:
        raise ValueError("No hay filas con datos para generar el archivo.")

    df["_FECHA_PARSED"] = _parsear_fecha_mixta(df["FECHA"])
    df["_VALOR_PARSED"] = df["VALOR"].apply(_num)

    filas_sin_fecha = int(df["_FECHA_PARSED"].isna().sum())
    if filas_sin_fecha:
        st.warning(f"⚠️ {filas_sin_fecha} fila(s) no tienen una fecha válida y se excluyeron.")

    df_validas = df[df["_FECHA_PARSED"].notna()].reset_index(drop=True)
    if df_validas.empty:
        raise ValueError("Ninguna fila válida pudo procesarse (revisa la fecha).")

    total = len(df_validas)
    n_bloques = max(1, -(-total // LIMITE_MOVIMIENTOS))  # ceil division
    hora_str = datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
    archivos_generados = []

    for bloque_idx in range(n_bloques):
        inicio = bloque_idx * LIMITE_MOVIMIENTOS
        fin = inicio + LIMITE_MOVIMIENTOS
        df_bloque = df_validas.iloc[inicio:fin].reset_index(drop=True)

        filas_debito  = []
        filas_credito = []

        for _, row in df_bloque.iterrows():
            valor        = row["_VALOR_PARSED"]
            concepto_src = str(row.get("CONCEPTO", "") or "").strip()
            fecha_str    = _fecha_str(row["_FECHA_PARSED"])
            detalle = f"{banco.upper()} {fecha_str} {concepto_src}".strip()

            filas_debito.append({
                "codigoCentroCosto": centro_costo,
                "dniTercero": nit_banco, "codigoTipoDniTercero": tipo_tercero,
                "codigoCuenta": cuenta_debito, "valor": valor,
                "factura": "", "fechaVencimiento": None,
                "codigoImpuesto": None, "valorBaseImpuesto": None,
                "porcentajeImpuesto": None, "detalle": detalle,
            })
            filas_credito.append({
                "codigoCentroCosto": centro_costo,
                "dniTercero": nit_banco, "codigoTipoDniTercero": tipo_tercero,
                "codigoCuenta": CUENTA_RENDIMIENTO_FINANCIERO, "valor": -abs(valor),
                "factura": "", "fechaVencimiento": None,
                "codigoImpuesto": None, "valorBaseImpuesto": None,
                "porcentajeImpuesto": None, "detalle": detalle,
            })

        todas_las_filas = filas_debito + filas_credito
        for idx, fila in enumerate(todas_las_filas, start=1):
            fila["Id"] = idx

        sufijo = f"_{bloque_idx + 1}" if n_bloques > 1 else ""
        nombre = f"RENDIMIENTO_FINANCIERO_{banco.upper()}_{empresa}_{hora_str}{sufijo}.xlsx"
        buffer = _generar_excel(todas_las_filas)

        # Misma ruta de guardado que usa Ingresos Bancarios
        ruta_auto = config.get("ruta_cargue_banco", "") if config else ""
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

    st.markdown("#### 📥 Descargar archivo(s) generado(s):")
    for arch in archivos_generados:
        arch["buffer"].seek(0)
        st.download_button(
            label=f"⬇️  Descargar {arch['nombre']}",
            data=arch["buffer"],
            file_name=arch["nombre"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_rendimiento_{arch['nombre']}",
        )

    return f"{total} movimiento(s) procesados en {len(archivos_generados)} archivo(s)."
