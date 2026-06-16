import pandas as pd
import streamlit as st

VALOR_CORRESPONSAL = 10000


def alistar_informacion(df_area_banco, archivo_clientes):
    df_clientes = pd.read_excel(archivo_clientes, sheet_name="sheet1")
    df_clientes.columns = [c.strip().upper() for c in df_clientes.columns]

    df_area_banco = df_area_banco.copy()
    df_area_banco["CEDULA_STR"] = df_area_banco["CEDULA"].astype(str).str.strip()
    df_clientes["IDEN_STR"]     = df_clientes["IDEN"].astype(str).str.strip()

    alertas     = []
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

    if alertas:
        df_extra = _resolver_escenarios_multifactura(alertas)
        if df_extra is not None and not df_extra.empty:
            df_simples = pd.concat([df_simples, df_extra], ignore_index=True)

    return df_simples, alertas


def _construir_fila(fila_banco, company, num_factura, cuota):
    entidad    = fila_banco.get("ENTIDAD", "")
    fecha      = fila_banco.get("FECHA")
    corresponsal = fila_banco.get("T_TRANSACCION")
    cuota_num  = _num(cuota)

    try:
        fecha_fmt = pd.Timestamp(fecha).strftime("%d-%m-%Y") if fecha else ""
    except Exception:
        fecha_fmt = str(fecha)[:10] if fecha else ""

    detalle = f"{entidad} {fecha_fmt}".strip()

    # ── Lógica CORRESPONSAL ──────────────────────────────────────────────
    # Si tiene corresponsal (reincidente) → VALOR_CB=10000, DIFERENCIA=cuota-10000
    if corresponsal and str(corresponsal).strip() != "" and str(corresponsal) != "nan":
        valor_cb   = VALOR_CORRESPONSAL
        diferencia = cuota_num - VALOR_CORRESPONSAL
    else:
        valor_cb   = None
        diferencia = None

    return {
        "ENTIDAD":         entidad,
        "FECHA":           fecha,
        "COMPANY":         company,
        "IDEN":            fila_banco.get("CEDULA"),
        "NUM_FACTURA":     num_factura,
        "CUOTA":           cuota_num,
        "RECIBO":          fila_banco.get("RECIBO"),
        "DIFERENCIA":      diferencia,
        "VALOR_CB":        valor_cb,
        "INMOVILIZACION":  fila_banco.get("INMOVILIZACION"),
        "OTROS_GASTOS":    fila_banco.get("OTROS_GASTOS"),
        "OBSERVACION":     fila_banco.get("OBSERVACION"),
        "CORRESPONSAL":    corresponsal,
        "FECHA_DOCUMENTO": fila_banco.get("FECHA_DOCUMENTO"),
        "DETALLE":         detalle,
    }


def _resolver_escenarios_multifactura(alertas):
    st.markdown("---")
    st.markdown("### ⚠️ Cédulas con múltiples facturas")
    st.markdown("Completa la distribución de todas las cédulas y confirma al final.")

    # Si ya hay distribuciones confirmadas las retornamos directo
    if st.session_state.get("distribuciones_confirmadas") is not None:
        st.success("✅ Distribuciones ya confirmadas.")
        return st.session_state["distribuciones_confirmadas"]

    with st.form(key="form_distribuciones"):
        inputs = {}

        for i, alerta in enumerate(alertas):
            cedula    = alerta["cedula"]
            pagos     = alerta["pagos"]
            facturas  = alerta["facturas"]
            companies = alerta["companies"]

            total_pagos = sum(
                _num(p.get("VALOR", 0)) for p in pagos
            )

            st.markdown(f"**📋 Cédula: {cedula}** — {alerta['n_pagos']} pago(s), {len(facturas)} factura(s)")
            for j, pago in enumerate(pagos):
                fecha_str = str(pago.get("FECHA", ""))[:10]
                valor     = _num(pago.get("VALOR", 0))
                st.caption(f"Pago {j+1}: {fecha_str} — ${valor:,.0f}")

            st.markdown(f"**Total a distribuir: ${total_pagos:,.0f}**")

            inputs[cedula] = {"facturas": {}, "pagos": pagos}
            cols = st.columns(min(len(facturas), 3))
            for k, (factura, company) in enumerate(zip(facturas, companies)):
                with cols[k % len(cols)]:
                    val = st.number_input(
                        f"Factura {factura} ({company})",
                        min_value=0.0,
                        value=0.0,
                        step=1000.0,
                        key=f"form_dist_{i}_{k}",
                        format="%.0f"
                    )
                    inputs[cedula]["facturas"][factura] = {"valor": val, "company": company}

            st.markdown("---")

        submitted = st.form_submit_button(
            "✅ Confirmar todas las distribuciones",
            use_container_width=True,
            type="primary"
        )

    if submitted:
        # Validar sumas
        errores = []
        for alerta in alertas:
            cedula     = alerta["cedula"]
            total      = sum(_num(p.get("VALOR", 0)) for p in alerta["pagos"])
            asignado   = sum(v["valor"] for v in inputs[cedula]["facturas"].values())
            if abs(total - asignado) > 0.5:
                errores.append(f"Cédula {cedula}: total ${total:,.0f} ≠ asignado ${asignado:,.0f}")

        if errores:
            for e in errores:
                st.error(f"⚠️ {e}")
            return None

        # Construir filas confirmadas
        filas = []
        for alerta in alertas:
            cedula    = alerta["cedula"]
            pagos     = alerta["pagos"]
            pago_base = pagos[0]

            # Fila base temporal para reutilizar _construir_fila
            fila_base = pd.Series({
                "ENTIDAD":         pago_base.get("ENTIDAD", ""),
                "FECHA":           pago_base.get("FECHA"),
                "CEDULA":          cedula,
                "RECIBO":          pago_base.get("RECIBO"),
                "INMOVILIZACION":  pago_base.get("INMOVILIZACION"),
                "OTROS_GASTOS":    pago_base.get("OTROS_GASTOS"),
                "OBSERVACION":     pago_base.get("OBSERVACION"),
                "T_TRANSACCION":   pago_base.get("T_TRANSACCION"),
                "FECHA_DOCUMENTO": pago_base.get("FECHA_DOCUMENTO"),
                "VALOR":           0,
            })

            for factura, info in inputs[cedula]["facturas"].items():
                if info["valor"] > 0:
                    fila = _construir_fila(fila_base, info["company"], factura, info["valor"])
                    filas.append(fila)

        if filas:
            df_confirmado = pd.DataFrame(filas)
            st.session_state["distribuciones_confirmadas"] = df_confirmado
            st.success(f"✅ {len(filas)} registros distribuidos correctamente.")
            st.rerun()

    return None


def _num(val):
    try:
        return float(str(val).replace(",", "") or 0)
    except Exception:
        return 0.0
