import pandas as pd
import streamlit as st


def alistar_informacion(df_area_banco, archivo_clientes):
    df_clientes = pd.read_excel(archivo_clientes, sheet_name="sheet1")
    df_clientes.columns = [c.strip().upper() for c in df_clientes.columns]

    df_area_banco = df_area_banco.copy()
    df_area_banco["CEDULA_STR"] = df_area_banco["CEDULA"].astype(str).str.strip()
    df_clientes["IDEN_STR"]     = df_clientes["IDEN"].astype(str).str.strip()

    alertas = []
    filas_simples = []

    grupos = df_area_banco.groupby("CEDULA_STR")

    for cedula_str, grupo in grupos:
        clientes_cedula = df_clientes[df_clientes["IDEN_STR"] == cedula_str]

        if clientes_cedula.empty:
            for _, fila_banco in grupo.iterrows():
                filas_simples.append(_construir_fila(fila_banco, None, None, fila_banco["VALOR"]))
            continue

        facturas  = clientes_cedula["NUM_FACTURA"].tolist()
        companies = clientes_cedula["COMPANY"].tolist()
        n_pagos    = len(grupo)
        n_facturas = len(facturas)

        if n_facturas == 1:
            # Escenario 1 o 2
            fechas_unicas = pd.to_datetime(grupo["FECHA"], errors="coerce").dt.normalize().nunique()
            if fechas_unicas <= 1 and n_pagos > 1:
                # Escenario 1: sumar
                total = pd.to_numeric(grupo["VALOR"], errors="coerce").sum()
                fila_base = grupo.iloc[0].copy()
                fila_base["VALOR"] = total
                filas_simples.append(_construir_fila(fila_base, companies[0], facturas[0], total))
            else:
                # Escenario 2: separadas
                for _, fila_banco in grupo.iterrows():
                    filas_simples.append(_construir_fila(fila_banco, companies[0], facturas[0], fila_banco["VALOR"]))
        else:
            # Escenarios 3 y 4: múltiples facturas
            alertas.append({
                "cedula":    cedula_str,
                "pagos":     grupo.to_dict("records"),
                "facturas":  facturas,
                "companies": companies,
                "n_pagos":   n_pagos,
            })

    df_simples = pd.DataFrame(filas_simples) if filas_simples else pd.DataFrame()

    # Manejar escenarios multifactura
    if alertas:
        df_extra = _resolver_escenarios_multifactura(alertas)
        if df_extra is not None and not df_extra.empty:
            df_simples = pd.concat([df_simples, df_extra], ignore_index=True)

    return df_simples, alertas


def _construir_fila(fila_banco, company, num_factura, cuota):
    entidad = fila_banco.get("ENTIDAD", "")
    fecha   = fila_banco.get("FECHA")
    try:
        fecha_fmt = pd.Timestamp(fecha).strftime("%d-%m-%Y") if fecha else ""
    except Exception:
        fecha_fmt = str(fecha)[:10] if fecha else ""
    detalle = f"{entidad} {fecha_fmt}".strip()

    return {
        "ENTIDAD":         entidad,
        "FECHA":           fecha,
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
        "DETALLE":         detalle,
    }


def _resolver_escenarios_multifactura(alertas):
    st.markdown("---")
    st.markdown("### ⚠️ Cédulas con múltiples facturas")
    st.markdown("Distribuye los valores en todas las cédulas y al final confirma todo en un solo botón.")

    # Inicializar estado de distribuciones
    if "distribuciones" not in st.session_state:
        st.session_state.distribuciones = {}

    todas_validas = True

    for i, alerta in enumerate(alertas):
        cedula    = alerta["cedula"]
        pagos     = alerta["pagos"]
        facturas  = alerta["facturas"]
        companies = alerta["companies"]

        total_pagos = sum(
            float(str(p.get("VALOR", 0)).replace(",", "") or 0)
            for p in pagos
        )

        with st.expander(
            f"📋 Cédula: {cedula} — {alerta['n_pagos']} pago(s), {len(facturas)} factura(s)",
            expanded=True
        ):
            st.markdown("**Pagos encontrados:**")
            for j, pago in enumerate(pagos):
                fecha_str = str(pago.get("FECHA", ""))[:10]
                valor     = pago.get("VALOR", 0)
                st.info(f"Pago {j+1}: Fecha {fecha_str} — Valor ${float(str(valor).replace(',','') or 0):,.0f}")

            st.markdown(f"**Total a distribuir: ${total_pagos:,.0f}**")
            st.markdown("**Asigna el valor a cada factura:**")

            cols = st.columns(min(len(facturas), 3))
            valores_factura = {}
            suma = 0.0

            for k, (factura, company) in enumerate(zip(facturas, companies)):
                with cols[k % len(cols)]:
                    val = st.number_input(
                        f"Factura {factura}\n({company})",
                        min_value=0.0,
                        value=st.session_state.distribuciones.get(f"{cedula}_{factura}", 0.0),
                        step=1000.0,
                        key=f"dist_{i}_{k}",
                        format="%.0f"
                    )
                    valores_factura[factura] = {"valor": val, "company": company}
                    suma += val
                    # Guardar en estado
                    st.session_state.distribuciones[f"{cedula}_{factura}"] = val

            diferencia = total_pagos - suma
            if abs(diferencia) > 0.5:
                st.warning(f"⚠️ Suma asignada ${suma:,.0f} — Diferencia: ${diferencia:,.0f}")
                todas_validas = False
            else:
                st.success(f"✅ Distribución correcta — Total: ${suma:,.0f}")

    # ── Un solo botón de confirmación al final ───────────────────────────
    st.markdown("---")
    if not todas_validas:
        st.warning("⚠️ Corrige las distribuciones antes de confirmar.")

    if st.button("✅ Confirmar todas las distribuciones", type="primary", use_container_width=True):
        filas = []
        for alerta in alertas:
            cedula    = alerta["cedula"]
            pagos     = alerta["pagos"]
            facturas  = alerta["facturas"]
            companies = alerta["companies"]
            pago_base = pagos[0]

            for factura, company in zip(facturas, companies):
                val = st.session_state.distribuciones.get(f"{cedula}_{factura}", 0.0)
                if val > 0:
                    entidad = pago_base.get("ENTIDAD", "")
                    fecha   = pago_base.get("FECHA")
                    try:
                        fecha_fmt = pd.Timestamp(fecha).strftime("%d-%m-%Y") if fecha else ""
                    except Exception:
                        fecha_fmt = str(fecha)[:10] if fecha else ""

                    filas.append({
                        "ENTIDAD":         entidad,
                        "FECHA":           fecha,
                        "COMPANY":         company,
                        "IDEN":            cedula,
                        "NUM_FACTURA":     factura,
                        "CUOTA":           val,
                        "RECIBO":          None,
                        "DIFERENCIA":      None,
                        "VALOR_CB":        None,
                        "INMOVILIZACION":  None,
                        "OTROS_GASTOS":    None,
                        "OBSERVACION":     None,
                        "CORRESPONSAL":    pago_base.get("T_TRANSACCION"),
                        "FECHA_DOCUMENTO": pago_base.get("FECHA_DOCUMENTO"),
                        "DETALLE":         f"{entidad} {fecha_fmt}".strip(),
                    })

        if filas:
            st.session_state.distribuciones_confirmadas = pd.DataFrame(filas)
            st.success(f"✅ {len(filas)} registros distribuidos correctamente.")
            st.rerun()

    # Retornar confirmadas si ya existen
    return st.session_state.get("distribuciones_confirmadas", None)
