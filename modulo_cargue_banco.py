import streamlit as st
import pandas as pd
from datetime import date


def _formato_cop(valor):
    """Formatea un número como moneda colombiana: $5.000.025,25"""
    try:
        s = f"{float(valor):,.2f}"
    except Exception:
        return "$0,00"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"${s}"


def render():
    st.markdown("""
    <div class="module-header">
        <div class="module-icon">📤</div>
        <div>
            <h1>Cargue de Banco</h1>
            <p>Ingresos bancarios, gastos bancarios y rendimiento financiero</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if "submodulo_cargue" not in st.session_state:
        st.session_state.submodulo_cargue = None

    sub = st.session_state.submodulo_cargue

    if sub is None:
        _render_menu_submodulos()
    elif sub == "ingresos":
        _render_volver()
        st.markdown("### 💰 Ingresos Bancarios")
        tab1, tab2 = st.tabs(["📥 1. Extracción", "📁 2. Generar Archivo"])
        with tab1:
            render_extraccion()
        with tab2:
            render_generar()
    elif sub == "gastos":
        _render_volver()
        render_gastos_bancarios()
    elif sub == "rendimiento":
        _render_volver()
        render_rendimiento_financiero()


def _render_menu_submodulos():
    st.markdown("### Selecciona el tipo de cargue")
    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">💰</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem;font-size:1rem">
                Ingresos Bancarios</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Extracción del Libro de Banco</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_ingresos", use_container_width=True, type="primary"):
            st.session_state.submodulo_cargue = "ingresos"
            st.rerun()

    with col2:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">💸</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem;font-size:1rem">
                Gastos Bancarios</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Pegado manual desde Excel</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_gastos", use_container_width=True, type="primary"):
            st.session_state.submodulo_cargue = "gastos"
            st.rerun()

    with col3:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">📈</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem;font-size:1rem">
                Rendimiento Financiero</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Intereses de cuentas de ahorro</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_rendimiento", use_container_width=True, type="primary"):
            st.session_state.submodulo_cargue = "rendimiento"
            st.rerun()


def _render_volver():
    if st.button("← Volver a Cargue de Banco", key="btn_volver_cargue"):
        st.session_state.submodulo_cargue = None
        st.rerun()
    st.markdown("<br>", unsafe_allow_html=True)


def render_gastos_bancarios():
    from generar_gastos_bancarios import crear_gastos_bancarios, detectar_concepto_cargue
    from conciliacion_bancaria import EMPRESAS_CUENTAS

    st.markdown("### Gastos Bancarios")
    st.caption(
        "Pega desde Excel las columnas FECHA, VALOR y CONCEPTO tal cual vienen "
        "(se conservan exactamente como las pegues). El CONCEPTO CARGUE se calcula solo."
    )

    col_banco, col_empresa = st.columns(2)
    with col_banco:
        banco = st.selectbox(
            "Banco",
            ["BANCOLOMBIA", "DAVIVIENDA", "OCCIDENTE", "BOGOTA"],
            key="gastos_banco_sel"
        )
    with col_empresa:
        empresa = st.selectbox(
            "Empresa",
            sorted(EMPRESAS_CUENTAS.keys()),
            key="gastos_empresa_sel"
        )

    columnas_base = ["VALOR", "FECHA", "CONCEPTO"]

    # Placeholder reservado arriba; se rellena después de leer la tabla
    # (el valor depende de lo que haya en el data_editor, que se lee más abajo)
    metric_placeholder = st.empty()

    # Plantilla vacía estable: NO se retroalimenta con lo ya editado.
    # Streamlit conserva internamente lo pegado/editado a través de la
    # misma key del widget entre recargas; si le devolvemos su propia
    # salida como "value" en cada corrida, se confunde y puede perder
    # filas pegadas al recargar por un clic de botón.
    if "gastos_editor_version" not in st.session_state:
        st.session_state["gastos_editor_version"] = 0
    editor_key = f"editor_gastos_bancarios_{st.session_state['gastos_editor_version']}"

    plantilla_vacia = pd.DataFrame({c: pd.Series(dtype="str") for c in columnas_base})

    st.markdown("**⬆️ Pega aquí las filas copiadas desde Excel:**")
    df_editado = st.data_editor(
        plantilla_vacia,
        num_rows="dynamic",
        use_container_width=True,
        key=editor_key,
        column_config={
            "FECHA":    st.column_config.TextColumn("Fecha (tal cual la pegues)"),
            "VALOR":    st.column_config.TextColumn("Valor (tal cual lo pegues)"),
            "CONCEPTO": st.column_config.TextColumn("Concepto (banco)"),
        }
    )

    from generar_gastos_bancarios import _num
    total_valor = sum(_num(v) for v in df_editado["VALOR"] if str(v).strip() != "")
    metric_placeholder.markdown(f"""
    <div style="background-color:#0f2a1a;border:1px solid #2ecc71;border-radius:10px;
                padding:0.4rem 1.2rem;margin-bottom:1rem;width:fit-content;">
        <div style="color:#2ecc71;font-size:0.85rem;font-weight:600;letter-spacing:0.5px;">
            💰 VALOR GASTOS BANCARIOS
        </div>
        <div style="color:#2ecc71;font-size:1.4rem;font-weight:700;">
            {_formato_cop(total_valor)}
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_prev, col_gen, col_lim = st.columns([1, 2, 1])
    with col_prev:
        ver_preview = st.button("🔍  Ver Concepto Cargue", use_container_width=True, key="preview_gastos")
    with col_lim:
        if st.button("🔄  Limpiar tabla", use_container_width=True, key="limpiar_gastos"):
            st.session_state["gastos_editor_version"] += 1
            st.session_state.pop("preview_gastos_df", None)
            st.rerun()

    if ver_preview:
        df_preview = df_editado[columnas_base].copy()
        df_preview["CONCEPTO CARGUE"] = df_preview["CONCEPTO"].apply(
            lambda x: detectar_concepto_cargue(x) or "⚠️ No detectado"
        )
        st.session_state["preview_gastos_df"] = df_preview

    if st.session_state.get("preview_gastos_df") is not None:
        st.markdown("**Vista previa del Concepto Cargue detectado:**")
        st.dataframe(st.session_state["preview_gastos_df"], use_container_width=True)

    with col_gen:
        if st.button("📤  Generar Archivo de Gastos Bancarios", type="primary", use_container_width=True):
            with st.spinner("Generando archivo..."):
                try:
                    resultado = crear_gastos_bancarios(
                        df_editado[columnas_base], empresa, banco, st.session_state.config
                    )
                    st.success(f"✅ {resultado}")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())


def render_rendimiento_financiero():
    from generar_rendimiento_financiero import crear_rendimiento_financiero
    from generar_gastos_bancarios import _num
    from conciliacion_bancaria import EMPRESAS_CUENTAS

    st.markdown("### 📈 Rendimiento Financiero")
    st.caption(
        "Pega desde Excel las columnas CONCEPTO, FECHA y VALOR tal cual vienen "
        "(se conservan exactamente como las pegues). Son los intereses que paga "
        "el banco por el dinero en cuentas de ahorro."
    )

    col_banco, col_empresa = st.columns(2)
    with col_banco:
        banco = st.selectbox(
            "Banco",
            ["BANCOLOMBIA", "DAVIVIENDA", "OCCIDENTE", "BOGOTA"],
            key="rendimiento_banco_sel"
        )
    with col_empresa:
        empresa = st.selectbox(
            "Empresa",
            sorted(EMPRESAS_CUENTAS.keys()),
            key="rendimiento_empresa_sel"
        )

    columnas_base = ["CONCEPTO", "FECHA", "VALOR"]

    # Placeholder reservado arriba; se rellena después de leer la tabla
    metric_placeholder = st.empty()

    if "rendimiento_editor_version" not in st.session_state:
        st.session_state["rendimiento_editor_version"] = 0
    editor_key = f"editor_rendimiento_financiero_{st.session_state['rendimiento_editor_version']}"

    plantilla_vacia = pd.DataFrame({c: pd.Series(dtype="str") for c in columnas_base})

    st.markdown("**⬆️ Pega aquí las filas copiadas desde Excel:**")
    df_editado = st.data_editor(
        plantilla_vacia,
        num_rows="dynamic",
        use_container_width=True,
        key=editor_key,
        column_config={
            "CONCEPTO": st.column_config.TextColumn("Concepto (banco)"),
            "FECHA":    st.column_config.TextColumn("Fecha (tal cual la pegues)"),
            "VALOR":    st.column_config.TextColumn("Valor (tal cual lo pegues)"),
        }
    )

    total_valor = sum(_num(v) for v in df_editado["VALOR"] if str(v).strip() != "")
    metric_placeholder.markdown(f"""
    <div style="background-color:#0f2a1a;border:1px solid #2ecc71;border-radius:10px;
                padding:0.4rem 1.2rem;margin-bottom:1rem;width:fit-content;">
        <div style="color:#2ecc71;font-size:0.85rem;font-weight:600;letter-spacing:0.5px;">
            📈 VALOR RENDIMIENTO FINANCIERO
        </div>
        <div style="color:#2ecc71;font-size:1.4rem;font-weight:700;">
            {_formato_cop(total_valor)}
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_gen, col_lim = st.columns([3, 1])
    with col_lim:
        if st.button("🔄  Limpiar tabla", use_container_width=True, key="limpiar_rendimiento"):
            st.session_state["rendimiento_editor_version"] += 1
            st.rerun()
    with col_gen:
        if st.button("📤  Generar Archivo de Rendimiento Financiero", type="primary", use_container_width=True):
            with st.spinner("Generando archivo..."):
                try:
                    resultado = crear_rendimiento_financiero(
                        df_editado[columnas_base], empresa, banco, st.session_state.config
                    )
                    st.success(f"✅ {resultado}")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())


def render_extraccion():
    from extraccion import extraer_cargue_banco

    # Verificar libro cargado
    if not st.session_state.get("archivo_libro"):
        st.warning("⚠️ Primero carga el Libro de Banco en el módulo **Aplicación y Compensación → Carga de Archivos**.")
        return

    st.markdown("### Selecciona las fechas a trabajar")
    st.caption("Los datos se extraen de las hojas BANCOLOMBIA y DAVIVIENDA del libro de banco.")

    # ── Selector de fechas ───────────────────────────────────────────────
    if "fechas_cargue" not in st.session_state:
        st.session_state.fechas_cargue = [None]

    col_add, col_del = st.columns([1, 1])
    with col_add:
        if len(st.session_state.fechas_cargue) < 5:
            if st.button("➕ Agregar fecha", key="add_fecha_cargue"):
                st.session_state.fechas_cargue.append(None)
                st.rerun()
    with col_del:
        if len(st.session_state.fechas_cargue) > 1:
            if st.button("➖ Quitar fecha", key="del_fecha_cargue"):
                st.session_state.fechas_cargue.pop()
                st.rerun()

    fechas_cols = st.columns(len(st.session_state.fechas_cargue))
    for i, col in enumerate(fechas_cols):
        with col:
            val = st.session_state.fechas_cargue[i] or date.today()
            fecha_sel = st.date_input(
                f"Fecha {i+1}",
                value=val,
                key=f"fecha_cargue_{i}"
            )
            st.session_state.fechas_cargue[i] = fecha_sel

    fechas_validas = [f for f in st.session_state.fechas_cargue if f is not None]
    fechas_str = ", ".join(f.strftime("%d/%m/%Y") for f in fechas_validas)
    st.caption(f"Fechas seleccionadas: **{fechas_str}**")
    st.markdown("<br>", unsafe_allow_html=True)

    # ── Botones extracción y limpieza ───────────────────────────────────
    col_ext, col_lim = st.columns([3, 1])
    with col_lim:
        if st.button("🔄  Limpiar", use_container_width=True, key="limpiar_cargue"):
            st.session_state["df_cargue_banco"] = None
            st.session_state["fechas_cargue"]   = [None]
            st.rerun()
    with col_ext:
     if st.button("⬇️  Extraer movimientos bancarios", type="primary", use_container_width=True):
            with st.spinner("Extrayendo movimientos..."):
                try:
                    df, resumen = extraer_cargue_banco(
                        st.session_state.archivo_libro,
                        fechas_validas
                    )
                    st.session_state["df_cargue_banco"] = df
                    st.success(f"✅ {resumen}")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

    # ── Tabla interactiva ────────────────────────────────────────────────
    if st.session_state.get("df_cargue_banco") is not None:
        df = st.session_state["df_cargue_banco"]
        st.markdown("---")

        c1, c2, c3 = st.columns(3)
        c1.metric("Total movimientos", len(df))
        c2.metric("Entidades", df["ENTIDAD"].nunique() if "ENTIDAD" in df.columns else 0)
        if "VALOR" in df.columns:
            total = pd.to_numeric(df["VALOR"], errors="coerce").sum()
            c3.metric("Valor total", f"${total:,.0f}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**📊 Movimientos extraídos:**")

        df_show = df.copy()
        if "FECHA" in df_show.columns:
            df_show["FECHA"] = pd.to_datetime(df_show["FECHA"], errors="coerce")

        st.data_editor(
            df_show,
            use_container_width=True,
            num_rows="fixed",
            key="tabla_cargue_banco",
            column_config={
                "FECHA":  st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                "VALOR":  st.column_config.NumberColumn("Valor", format="$%d"),
            }
        )


def render_generar():
    from generar_cargue_banco import crear_cargue_banco

    if st.session_state.get("df_cargue_banco") is None:
        st.warning("⚠️ Primero extrae los movimientos en la pestaña **1. Extracción**.")
        return

    st.markdown("### Generar archivo de cargue banco")
    st.caption("Genera un archivo Items por cada fecha con débitos y créditos.")

    if st.button("📤  Crear Cargue Banco", type="primary", use_container_width=True):
        with st.spinner("Generando archivos..."):
            try:
                resultado = crear_cargue_banco(
                    st.session_state["df_cargue_banco"],
                    st.session_state.config
                )
                st.success(f"✅ {resultado}")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                import traceback
                st.code(traceback.format_exc())