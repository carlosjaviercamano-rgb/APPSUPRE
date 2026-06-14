import pandas as pd
import streamlit as st


def alistar_informacion(df_area_banco, archivo_clientes):
    """
    Cruza df_area_banco con la base de clientes activos.
    Maneja los 4 escenarios y retorna (df_resultado, alertas).
    """
    # ── Cargar clientes activos ──────────────────────────────────────────
    df_clientes = pd.read_excel(archivo_clientes, sheet_name="sheet1")
    df_clientes.columns = [c.strip().upper() for c in df_clientes.columns]

    # Normalizar cédulas
    df_area_banco = df_area_banco.copy()
    df_area_banco["CEDULA_STR"] = df_area_banco["CEDULA"].astype(str).str.strip()
    df_clientes["IDEN_STR"]     = df_clientes["IDEN"].astype(str).str.strip()

    alertas = []
    filas_resultado = []

    # Agrupar área de banco por cédula para detectar duplicados
    grupos = df_area_banco.groupby("CEDULA_STR")

    for cedula_str, grupo in grupos:
        # Buscar en clientes activos
        clientes_cedula = df_clientes[df_clientes["IDEN_STR"] == cedula_str]

        if clientes_cedula.empty:
            # Cédula no encontrada en clientes — se incluye sin factura
            for _, fila_banco in grupo.iterrows():
                filas_resultado.append(_construir_fila(fila_banco, None, None, None))
            continue

        facturas = clientes_cedula["NUM_FACTURA"].tolist()
        companies = clientes_cedula["COMPANY"].tolist()

        # Clasificar escenario
        n_pagos   = len(grupo)
        n_facturas = len(facturas)

        if n_pagos == 1 and n_facturas == 1:
            # ── ESCENARIO 1 o caso simple ─────────────────────────────
            fila_banco = grupo.iloc[0]
            misma_fecha = True  # solo hay 1 fila
            filas_resultado.append(
                _construir_fila(fila_banco, companies[0], facturas[0], fila_banco["VALOR"])
            )

        elif n_pagos > 1 and n_facturas == 1:
            # Múltiples pagos, misma factura
            fechas = grupo["FECHA"].dt.normalize().unique() if hasattr(grupo["FECHA"].iloc[0], 'date') else group["FECHA"].unique()

            if len(fechas) == 1:
                # ── ESCENARIO 1: misma fecha → SUMAR ─────────────────
                total = pd.to_numeric(grupo["VALOR"], errors="coerce").sum()
                fila_base = grupo.iloc[0].copy()
                fila_base["VALOR"] = total
                filas_resultado.append(
                    _construir_fila(fila_base, companies[0], facturas[0], total)
                )
            else:
                # ── ESCENARIO 2: diferente fecha → fila separada ──────
                for _, fila_banco in grupo.iterrows():
                    filas_resultado.append(
                        _construir_fila(fila_banco, companies[0], facturas[0], fila_banco["VALOR"])
                    )

        elif n_facturas > 1:
            # ── ESCENARIOS 3 y 4: varias facturas ────────────────────
            alerta = {
                "cedula": cedula_str,
                "pagos": grupo.to_dict("records"),
                "facturas": facturas,
                "companies": companies,
                "n_pagos": n_pagos,
                "n_facturas": n_facturas,
            }
            alertas.append(alerta)
            # Las filas se agregarán cuando el usuario resuelva el popup

    df_resultado = pd.DataFrame(filas_resultado) if filas_resultado else pd.DataFrame()

    # ── Mostrar popups para escenarios 3 y 4 ────────────────────────────
    if alertas:
        filas_extra = _resolver_escenarios_multifactura(alertas)
        if filas_extra:
            df_extra = pd.DataFrame(filas_extra)
            df_resultado = pd.concat([df_resultado, df_extra], ignore_index=True)

    return df_resultado, alertas


def _construir_fila(fila_banco, company, num_factura, cuota):
    """Construye una fila con la estructura de Sheet1."""
    return {
        "ENTIDAD":         fila_banco.get("ENTIDAD"),
        "FECHA":           fila_banco.get("FECHA"),
        "COMPANY":         company,
        "IDEN":            fila_banco.get("CEDULA"),
        "NUM_FACTURA":     num_factura,
        "CUOTA":           cuota,
        "RECIBO":          fila_banco.get("RECIBO"),
        "DIFERENCIA":      fila_banco.get("DIFERENCIA"),
        "VALOR_CB":        fila_banco.get("VALOR_CB"),
        "INMOVILIZACION":  fila_banco.get("INMOVILIZACION"),
        "OTROS_GASTOS":    fila_banco.get("OTROS_GASTOS"),
        "OBSERVACION":     fila_banco.get("OBSERVACION"),
        "CORRESPONSAL":    fila_banco.get("T_TRANSACCION"),
        "FECHA_DOCUMENTO": fila_banco.get("FECHA_DOCUMENTO"),
        "DETALLE":         None,
    }


def _resolver_escenarios_multifactura(alertas):
    """
    Muestra un popup (expander) por cada cédula con múltiples facturas
    para que el usuario distribuya los valores manualmente.
    Retorna las filas resultantes.
    """
    filas_resultado = []

    st.markdown("---")
    st.markdown("### ⚠️ Cédulas con múltiples facturas")
    st.markdown("Distribuye los valores de pago entre las facturas disponibles. Solo se incluirán facturas con valor mayor a cero.")

    for i, alerta in enumerate(alertas):
        cedula    = alerta["cedula"]
        pagos     = alerta["pagos"]
        facturas  = alerta["facturas"]
        companies = alerta["companies"]
        n_pagos   = alerta["n_pagos"]

        with st.expander(f"📋 Cédula: {cedula} — {n_pagos} pago(s), {len(facturas)} factura(s)", expanded=True):

            # Mostrar pagos disponibles
            st.markdown("**Pagos encontrados:**")
            for j, pago in enumerate(pagos):
                fecha_str = str(pago.get("FECHA", ""))[:10]
                valor     = pago.get("VALOR", 0)
                st.info(f"Pago {j+1}: Fecha {fecha_str} — Valor ${valor:,.0f}")

            total_pagos = sum(
                float(str(p.get("VALOR", 0)).replace(",", "") or 0)
                for p in pagos
            )

            st.markdown(f"**Total a distribuir: ${total_pagos:,.0f}**")
            st.markdown("**Asigna el valor a cada factura:**")

            valores_por_factura = {}
            suma_ingresada = 0.0

            cols = st.columns(min(len(facturas), 3))
            for k, (factura, company) in enumerate(zip(facturas, companies)):
                with cols[k % len(cols)]:
                    val = st.number_input(
                        f"Factura {factura}\n({company})",
                        min_value=0.0,
                        value=0.0,
                        step=1000.0,
                        key=f"dist_{i}_{k}",
                        format="%.0f"
                    )
                    valores_por_factura[factura] = {"valor": val, "company": company}
                    suma_ingresada += val

            # Validación de suma
            diferencia = total_pagos - suma_ingresada
            if abs(diferencia) > 0.5:
                st.warning(f"⚠️ La suma asignada (${suma_ingresada:,.0f}) difiere del total (${total_pagos:,.0f}). Diferencia: ${diferencia:,.0f}")
            else:
                st.success(f"✅ Distribución correcta. Total asignado: ${suma_ingresada:,.0f}")

            if st.button(f"✅ Confirmar distribución — Cédula {cedula}", key=f"confirm_{i}"):
                # Construir filas solo para facturas con valor > 0
                pago_base = pagos[0]  # Usamos el primer pago como base de fecha/entidad
                for factura, info in valores_por_factura.items():
                    if info["valor"] > 0:
                        fila = {
                            "ENTIDAD":         pago_base.get("ENTIDAD"),
                            "FECHA":           pago_base.get("FECHA"),
                            "COMPANY":         info["company"],
                            "IDEN":            cedula,
                            "NUM_FACTURA":     factura,
                            "CUOTA":           info["valor"],
                            "RECIBO":          None,
                            "DIFERENCIA":      None,
                            "VALOR_CB":        None,
                            "INMOVILIZACION":  None,
                            "OTROS_GASTOS":    None,
                            "OBSERVACION":     None,
                            "CORRESPONSAL":    pago_base.get("T_TRANSACCION"),
                            "FECHA_DOCUMENTO": pago_base.get("FECHA_DOCUMENTO"),
                            "DETALLE":         None,
                        }
                        filas_resultado.append(fila)
                st.success(f"✅ Distribuido correctamente para cédula {cedula}.")

    return filas_resultado
