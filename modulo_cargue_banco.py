import streamlit as st
import pandas as pd
from datetime import date


def render():
    st.markdown("""
    <div class="module-header">
        <div class="module-icon">📤</div>
        <div>
            <h1>Cargue Banco</h1>
            <p>Extracción y generación de archivos de cargue bancario</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📥 1. Extracción", "📁 2. Generar Archivo"])

    with tab1:
        render_extraccion()
    with tab2:
        render_generar()


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

    # ── Botón extracción ─────────────────────────────────────────────────
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
                    st.session_state["df_cargue_banco"]
                )
                st.success(f"✅ {resultado}")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
