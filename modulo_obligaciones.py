import os
from datetime import date

import pandas as pd
import streamlit as st

import generar_obligaciones as G


def _estilo_moneda(df, columnas):
    return df.style.format({c: "${:,.2f}" for c in columnas})


def _mostrar_vista(df, titulo_total):
    """Muestra la vista previa de una hoja de control con formato de moneda."""
    if df is None or df.empty:
        st.info("Sin filas para mostrar.")
        return
    monedas = [c for c in df.columns
               if str(c).lower().startswith(("valor_cuota", "capital", "interes"))]
    fmt = {c: "${:,.2f}" for c in monedas}
    st.dataframe(df.style.format(fmt, na_rep=""), use_container_width=True, hide_index=True)
    col_val = next((c for c in df.columns if str(c).lower().startswith("valor_cuota")), None)
    if col_val is not None:
        total = pd.to_numeric(df[col_val], errors="coerce").sum()
        st.caption(f"{titulo_total}: ${total:,.2f}")


def render(config=None):
    st.markdown("""
    <div class="module-header">
        <div class="module-icon">🧾</div>
        <div>
            <h1>Obligaciones Financieras</h1>
            <p>Actualización mensual del histórico y del libro de control de pagos</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    cfg = config or st.session_state.get("config", {})
    ruta_hist = (cfg.get("ruta_obligaciones_historico") or "").strip().strip('"')
    ruta_ctrl = (cfg.get("ruta_control_pagos_obligaciones") or "").strip().strip('"')

    # ── Verificación de rutas ─────────────────────────────────────────────
    problemas = []
    for etiqueta, ruta in (("Libro histórico", ruta_hist), ("Libro de control de pagos", ruta_ctrl)):
        if not ruta:
            problemas.append(f"{etiqueta}: falta configurar la ruta.")
        elif ruta.lower().endswith(".lnk"):
            problemas.append(f"{etiqueta}: la ruta apunta a un acceso directo (.lnk). Usa la ruta real del archivo .xlsx.")
        elif os.path.isdir(ruta):
            xlsx = sorted(f for f in os.listdir(ruta) if f.lower().endswith(".xlsx") and not f.startswith("~$"))
            ayuda = (f" En esa carpeta hay: {', '.join(xlsx)}." if xlsx else "")
            problemas.append(
                f"{etiqueta}: la ruta es una carpeta, falta el nombre del archivo al final "
                f"(por ejemplo «{ruta.rstrip(chr(92))}\\nombre_del_libro.xlsx»).{ayuda}")
        elif not os.path.isfile(ruta):
            problemas.append(f"{etiqueta}: no se encontró el archivo «{ruta}».")
    if problemas:
        for p in problemas:
            st.warning(f"⚠️ {p}")
        st.info("Configura las rutas en **⚙️ Configuración → Obligaciones Financieras**.")
        return

    st.caption(
        "Sube el reporte de AWS del mes. La app cruza cada obligación por **empresa + factura + "
        "identificación**, te deja clasificar las nuevas, marca como *No activo* las que ya no vienen "
        "y actualiza el histórico y el libro de control. Antes de guardar verás una vista previa y se "
        "hace un respaldo de ambos libros. **Cierra los dos libros en Excel antes de guardar.**"
    )

    # ── Paso 1: reporte + mes ─────────────────────────────────────────────
    st.markdown("**1. Reporte de obligaciones (AWS, .csv o .xlsx):**")
    archivo = st.file_uploader(
        "Reporte de obligaciones", type=["csv", "xlsx"],
        key="obl_up_reporte", label_visibility="collapsed"
    )

    hoy = date.today()
    col_mes, col_anio, _ = st.columns([2, 1, 3])
    with col_mes:
        mes_nombre = st.selectbox(
            "Mes a reportar", [m.capitalize() for m in G.MESES_ES],
            index=hoy.month - 1, key="obl_mes"
        )
    with col_anio:
        anio = st.number_input("Año", min_value=2024, max_value=2100, value=hoy.year,
                               step=1, key="obl_anio")
    mes_num = [m.capitalize() for m in G.MESES_ES].index(mes_nombre) + 1
    st.caption(f"Columnas del mes en el libro de control: "
               f"`valor_cuota_{G.MESES_ES[mes_num - 1]}_{int(anio)}` y "
               f"`estado_de_la_cuota_{G.MESES_ES[mes_num - 1]}_{int(anio)}` "
               f"(en bancos, además `capital_…` e `interes_…`).")

    col_a, col_l = st.columns([3, 1])
    with col_a:
        analizar = st.button("🔍  Analizar reporte", type="primary",
                             use_container_width=True, key="obl_btn_analizar")
    with col_l:
        if st.button("🔄  Limpiar", use_container_width=True, key="obl_btn_limpiar"):
            for k in ["obl_analisis", "obl_df_rep", "obl_version", "obl_ok_cancel"]:
                st.session_state.pop(k, None)
            st.rerun()

    if analizar:
        if archivo is None:
            st.error("❌ Debes cargar el reporte primero.")
        else:
            try:
                df_rep = G.leer_reporte(archivo)
                analisis = G.analizar(df_rep, ruta_hist)
                st.session_state["obl_df_rep"] = df_rep
                st.session_state["obl_analisis"] = analisis
                st.session_state["obl_version"] = st.session_state.get("obl_version", 0) + 1
                st.session_state.pop("obl_ok_cancel", None)
            except Exception as e:
                st.session_state.pop("obl_analisis", None)
                st.error(f"❌ {e}")

    analisis = st.session_state.get("obl_analisis")
    df_rep = st.session_state.get("obl_df_rep")
    if analisis is None or df_rep is None:
        return

    version = st.session_state.get("obl_version", 0)
    nuevas = analisis["nuevas"]
    canceladas = analisis["canceladas"]
    reactivadas = analisis["reactivadas"]

    st.markdown("---")
    st.markdown("#### 📊 Resultado del cruce")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("En el reporte", f"{len(df_rep):,}")
    c2.metric("✅ Existentes", f"{analisis['n_existentes']:,}")
    c3.metric("🆕 Nuevas", f"{len(nuevas):,}")
    c4.metric("🚫 Canceladas", f"{len(canceladas):,}")

    listo = True

    # ── Paso 2: clasificar nuevas ─────────────────────────────────────────
    tipos_nuevas = {}
    if len(nuevas) > 0:
        st.markdown("#### 🆕 Obligaciones nuevas — clasifícalas antes de guardar")
        st.caption("No se escribe nada en el histórico hasta que todas tengan tipo: "
                   "**Bancario** (se paga por bancos), **Particular** (se paga por caja) o "
                   "**Gerencia** (manejo interno, no va al libro de control).")
        base = pd.DataFrame({
            "empresa": nuevas["empresa"].values,
            "factura": nuevas["factura"].values,
            "identificacion_proveedor": nuevas["identificacion_proveedor"].astype(str).values,
            "nombre_proveedor": nuevas["nombre_proveedor"].values,
            "valor_cuota": nuevas["valor_cuota"].astype(float).values,
            "tipo_de_obligacion": [None] * len(nuevas),
        })
        editado = st.data_editor(
            base, num_rows="fixed", use_container_width=True, hide_index=True,
            key=f"obl_editor_nuevas_{version}",
            disabled=["empresa", "factura", "identificacion_proveedor", "nombre_proveedor", "valor_cuota"],
            column_config={
                "valor_cuota": st.column_config.NumberColumn("valor_cuota", format="$ %.2f"),
                "tipo_de_obligacion": st.column_config.SelectboxColumn(
                    "Tipo de obligación", options=G.TIPOS_OBLIGACION, required=True),
            },
        )
        for clave, tipo in zip(nuevas["_clave"].tolist(), editado["tipo_de_obligacion"].tolist()):
            if tipo:
                tipos_nuevas[clave] = tipo
        faltan = len(nuevas) - len(tipos_nuevas)
        if faltan:
            listo = False
            st.warning(f"⚠️ Faltan {faltan} obligación(es) por clasificar.")

    # ── Canceladas ────────────────────────────────────────────────────────
    if canceladas:
        st.markdown("#### 🚫 Obligaciones que ya no vienen en el reporte")
        st.caption("Se marcarán como **No activo** en el histórico y, si estaban en el libro de control, "
                   "se mueven a la hoja CANCELADAS conservando su historial.")
        st.dataframe(_estilo_moneda(pd.DataFrame(canceladas), ["valor_cuota"]),
                     use_container_width=True, hide_index=True)
        n_act = max(analisis["n_activas_historico"], 1)
        if len(canceladas) > max(5, 0.25 * n_act):
            st.error(f"⚠️ Son {len(canceladas)} de {n_act} obligaciones activas "
                     f"({len(canceladas) / n_act:.0%}). Si el reporte salió incompleto o filtrado, "
                     "se desactivarían por error. Revisa el archivo antes de confirmar.")
        confirma = st.checkbox(
            f"Confirmo que estas {len(canceladas)} obligación(es) están canceladas",
            key=f"obl_ok_cancel_{version}")
        if not confirma:
            listo = False

    if reactivadas:
        st.info(f"♻️ {len(reactivadas)} obligación(es) que estaban *No activo* volvieron a venir en el "
                "reporte y se reactivarán.")
        st.dataframe(_estilo_moneda(pd.DataFrame(reactivadas), ["valor_cuota"]),
                     use_container_width=True, hide_index=True)

    if not listo:
        st.caption("Completa la clasificación y la confirmación para ver la vista previa.")
        return

    # ── Paso 3: vista previa (ensayo en memoria) ──────────────────────────
    st.markdown("---")
    incluir_bancos = st.checkbox("Actualizar también la hoja PAGOS_POR_BANCOS", value=True,
                                 key="obl_incluir_bancos")
    try:
        ensayo = G.ejecutar(df_rep, ruta_hist, ruta_ctrl, mes_num, int(anio), tipos_nuevas,
                            incluir_bancos=incluir_bancos, guardar=False)
    except Exception as e:
        st.error(f"❌ {e}")
        return

    st.markdown("#### 👁️ Vista previa de lo que se va a guardar")
    h = ensayo["historico"]
    tab_h, tab_caja, tab_bancos = st.tabs(["Histórico", "PAGOS_POR_CAJA", "PAGOS_POR_BANCOS"])

    with tab_h:
        st.write(f"**{h['actualizadas']}** obligaciones actualizadas · **{len(h['nuevas'])}** nuevas · "
                 f"**{len(h['canceladas'])}** pasan a No activo · **{len(h['reactivadas'])}** reactivadas.")
        if h["nuevas"]:
            st.dataframe(_estilo_moneda(pd.DataFrame(h["nuevas"]), ["valor_cuota"]),
                         use_container_width=True, hide_index=True)

    with tab_caja:
        rc = ensayo["caja"]
        st.write(f"**{len(rc['actualizadas'])}** filas actualizadas · **{len(rc['nuevas'])}** nuevas · "
                 f"**{len(rc['archivadas'])}** pasan a CANCELADAS. Una fila por identificación, "
                 "con la suma de sus cuotas del mes (solo tipo Particular).")
        for adv in rc["advertencias"]:
            st.warning(f"⚠️ {adv}")
        st.caption("Antigüedad_Crédito: **EXISTENTE** si la obligación ya estaba en el histórico, "
                   "**NUEVO** si viene por primera vez en este reporte. El estado de la cuota "
                   "queda en PENDIENTE (con lista desplegable PENDIENTE / PAGADO).")
        _mostrar_vista(rc.get("vista"), "Total a pagar por caja")

    with tab_bancos:
        if ensayo["bancos"] is None:
            st.info("No se actualizará la hoja de bancos.")
        else:
            rb = ensayo["bancos"]
            st.write(f"**{len(rb['actualizadas'])}** filas actualizadas · **{len(rb['nuevas'])}** nuevas · "
                     f"**{len(rb['archivadas'])}** pasan a CANCELADAS. Una fila por obligación (solo tipo Bancario).")
            for adv in rb["advertencias"]:
                st.warning(f"⚠️ {adv}")
            _mostrar_vista(rb.get("vista"), "Total a pagar por bancos")

    # ── Paso 4: guardar ───────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾  Guardar y actualizar libros", type="primary",
                 use_container_width=True, key="obl_btn_guardar"):
        try:
            with st.spinner("Guardando..."):
                res = G.ejecutar(df_rep, ruta_hist, ruta_ctrl, mes_num, int(anio), tipos_nuevas,
                                 incluir_bancos=incluir_bancos, guardar=True)
            st.success("✅ Libros actualizados correctamente.")
            st.write(f"📘 Histórico: `{ruta_hist}`")
            st.write(f"📗 Control de pagos: `{ruta_ctrl}`")
            st.caption("Respaldos previos guardados en: " + " · ".join(f"`{r}`" for r in res["respaldos"]))
            for k in ["obl_analisis", "obl_df_rep"]:
                st.session_state.pop(k, None)
        except PermissionError as e:
            st.error(f"🔒 {e}")
        except Exception as e:
            st.error(f"❌ {e}")
