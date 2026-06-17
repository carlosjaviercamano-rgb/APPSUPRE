import streamlit as st
import pandas as pd
import os

# ─── Estilos del módulo ────────────────────────────────────────────────────
STYLES = """
<style>
/* Selector de tipo de pago */
.tipo-pago-card {
    border: 2px solid #e2e8f0;
    border-radius: 10px;
    padding: 1.2rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.15s;
    background: white;
}
.tipo-pago-card:hover { border-color: #1e40af; background: #eff6ff; }
.tipo-pago-card.selected { border-color: #1e40af; background: #eff6ff; }
.tipo-pago-card .tp-icon { font-size: 2rem; }
.tipo-pago-card .tp-title { font-weight: 600; color: #1e293b; font-size: 0.95rem; margin-top: 0.4rem; }
.tipo-pago-card .tp-sub { color: #64748b; font-size: 0.78rem; margin-top: 0.2rem; }

/* Tabla interactiva */
.tabla-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 0.8rem;
}
.tabla-header h3 { margin: 0; font-size: 1rem; color: #1e293b; }
.badge-count {
    background: #dbeafe;
    color: #1e40af;
    font-size: 0.78rem;
    font-weight: 600;
    padding: 0.2rem 0.6rem;
    border-radius: 20px;
}

/* Pasos del proceso */
.paso-container {
    display: flex;
    align-items: flex-start;
    gap: 1rem;
    margin-bottom: 1.5rem;
}
.paso-num {
    background: #1e40af;
    color: white;
    width: 28px;
    height: 28px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.8rem;
    font-weight: 700;
    flex-shrink: 0;
    margin-top: 0.1rem;
}
.paso-content h4 { margin: 0 0 0.2rem 0; font-size: 0.92rem; color: #1e293b; }
.paso-content p  { margin: 0; font-size: 0.82rem; color: #64748b; }

/* Botones de acción */
.btn-accion {
    background: #1e40af;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.55rem 1.4rem;
    font-weight: 600;
    font-size: 0.88rem;
    cursor: pointer;
}
.btn-secundario {
    background: white;
    color: #1e40af;
    border: 1.5px solid #1e40af;
    border-radius: 8px;
    padding: 0.5rem 1.2rem;
    font-weight: 600;
    font-size: 0.88rem;
    cursor: pointer;
}

/* Alerta info */
.alerta-info {
    background: #eff6ff;
    border-left: 4px solid #1e40af;
    border-radius: 4px;
    padding: 0.8rem 1rem;
    font-size: 0.85rem;
    color: #1e40af;
    margin: 0.5rem 0;
}
</style>
"""

# ─── Columnas de AREA DE BANCO ─────────────────────────────────────────────
COLUMNAS_AREA_BANCO = [
    "ENTIDAD", "FECHA", "CEDULA", "VALOR", "T_TRANSACCION",
    "NUM_FACTURA", "CUOTA", "RECIBO", "DIFERENCIA",
    "VALOR_CB", "INMOVILIZACION", "OTROS_GASTOS",
    "OBSERVACION", "CORRESPONSAL", "FECHA_DOCUMENTO"
]

def init_estado():
    """Inicializa variables de sesión del módulo"""
    defaults = {
        "tipo_pago": None,          # "bancarios" | "recaudos"
        "fecha_recaudo": None,
        "df_area_banco": None,      # DataFrame tabla interactiva
        "df_sheet1": None,          # DataFrame cruzado (alistar información)
        "archivo_libro": None,
        "archivo_clientes": None,
        "archivo_corresponsal": None,
        "paso_actual": 1,           # 1=carga, 2=extracción, 3=alistar, 4=planos/comp
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def render():
    """Punto de entrada del módulo"""
    init_estado()
    st.markdown(STYLES, unsafe_allow_html=True)

    # Encabezado del módulo
    st.markdown("""
    <div class="module-header">
        <div class="module-icon">💳</div>
        <div>
            <h1>Aplicación y Compensación de Pagos</h1>
            <p>Extracción, cruce y generación de planos y compensaciones</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Tabs del módulo ──────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "📥 1. Carga de Archivos",
        "📊 2. Tabla de Pagos",
        "🔗 3. Alistar Información",
        "📁 4. Generar Archivos"
    ])

    with tab1:
        render_carga_archivos()

    with tab2:
        render_tabla_pagos()

    with tab3:
        render_alistar_informacion()

    with tab4:
        render_generar_archivos()


# ══════════════════════════════════════════════════════════════════════════
# TAB 1 — CARGA DE ARCHIVOS
# ══════════════════════════════════════════════════════════════════════════
def render_carga_archivos():
    st.markdown("### Archivos necesarios para el proceso")
    st.markdown('<div class="alerta-info">📌 Carga los tres archivos antes de continuar. El libro de banco se puede reemplazar cada vez que vayas a trabajar.</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**📂 Libro de Banco**")
        st.caption("Archivo descargado de SharePoint / OneDrive")
        archivo_libro = st.file_uploader(
            "Subir libro de banco",
            type=["xlsx", "xlsm"],
            key="up_libro",
            label_visibility="collapsed"
        )
        if archivo_libro:
            st.session_state.archivo_libro = archivo_libro
            st.success(f"✅ {archivo_libro.name}")

    with col2:
        st.markdown("**👥 Clientes Activos**")
        st.caption("Base de datos: company, cédula, factura")
        archivo_clientes = st.file_uploader(
            "Subir clientes activos",
            type=["xlsx"],
            key="up_clientes",
            label_visibility="collapsed"
        )
        if archivo_clientes:
            st.session_state.archivo_clientes = archivo_clientes
            st.success(f"✅ {archivo_clientes.name}")

    with col3:
        st.markdown("**🔑 Corresponsal**")
        st.caption("Historial de cédulas reincidentes")
        archivo_corresponsal = st.file_uploader(
            "Subir archivo corresponsal",
            type=["xlsx"],
            key="up_corresponsal",
            label_visibility="collapsed"
        )
        if archivo_corresponsal:
            st.session_state.archivo_corresponsal = archivo_corresponsal
            st.success(f"✅ {archivo_corresponsal.name}")

    # ── Estado de archivos cargados ──────────────────────────────────────
    st.markdown("---")
    libro_ok      = st.session_state.archivo_libro is not None
    clientes_ok   = st.session_state.archivo_clientes is not None
    corresponsal_ok = st.session_state.archivo_corresponsal is not None

    c1, c2, c3 = st.columns(3)
    c1.metric("Libro de Banco",   "✅ Cargado" if libro_ok      else "⚠️ Pendiente")
    c2.metric("Clientes Activos", "✅ Cargado" if clientes_ok   else "⚠️ Pendiente")
    c3.metric("Corresponsal",     "✅ Cargado" if corresponsal_ok else "⚠️ Pendiente")

    if libro_ok and clientes_ok and corresponsal_ok:
        st.success("🎉 Todos los archivos están cargados. Ve a la pestaña **2. Tabla de Pagos** para continuar.")


# ══════════════════════════════════════════════════════════════════════════
# TAB 2 — TABLA DE PAGOS (Extracción y tabla interactiva)
# ══════════════════════════════════════════════════════════════════════════
def render_tabla_pagos():
    from extraccion import extraer_pagos_bancarios, extraer_pagos_recaudos

    if not st.session_state.archivo_libro:
        st.warning("⚠️ Primero carga el Libro de Banco en la pestaña **1. Carga de Archivos**.")
        return

    st.markdown("### Selecciona el tipo de pago a trabajar")

    # ── Selector tipo de pago ────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        sel1 = st.session_state.tipo_pago == "bancarios"
        if st.button(
            "🏦  PAGOS BANCARIOS\nBancolombia · Davivienda · Suprecredito",
            key="sel_bancarios",
            use_container_width=True,
            type="primary" if sel1 else "secondary"
        ):
            st.session_state.tipo_pago = "bancarios"
            st.session_state.df_area_banco = None
            st.rerun()

    with col2:
        sel2 = st.session_state.tipo_pago == "recaudos"
        if st.button(
            "📋  PAGOS POR RECAUDOS\nOccidente · Suprecredito 2026 · Record",
            key="sel_recaudos",
            use_container_width=True,
            type="primary" if sel2 else "secondary"
        ):
            st.session_state.tipo_pago = "recaudos"
            st.session_state.df_area_banco = None
            st.rerun()

    if not st.session_state.tipo_pago:
        st.info("👆 Selecciona el tipo de pago para continuar.")
        return

    st.markdown("---")

    # ── Selector de fecha (solo para RECAUDOS) ───────────────────────────
    if st.session_state.tipo_pago == "recaudos":
        st.markdown("**📅 Fechas de trabajo** (máximo 5 fechas)")
        st.caption("Útil para fines de semana o días acumulados.")

        if "fechas_recaudo" not in st.session_state:
            st.session_state.fechas_recaudo = [None]

        # Agregar / quitar fechas
        col_add, col_del = st.columns([1, 1])
        with col_add:
            if len(st.session_state.fechas_recaudo) < 5:
                if st.button("➕ Agregar fecha", key="add_fecha"):
                    st.session_state.fechas_recaudo.append(None)
                    st.rerun()
        with col_del:
            if len(st.session_state.fechas_recaudo) > 1:
                if st.button("➖ Quitar fecha", key="del_fecha"):
                    st.session_state.fechas_recaudo.pop()
                    st.rerun()

        fechas_cols = st.columns(len(st.session_state.fechas_recaudo))
        for i, col in enumerate(fechas_cols):
            with col:
                from datetime import date
                val = st.session_state.fechas_recaudo[i] or date.today()
                fecha_sel = st.date_input(
                    f"Fecha {i+1}",
                    value=val,
                    key=f"fecha_recaudo_{i}"
                )
                st.session_state.fechas_recaudo[i] = fecha_sel

        fechas_validas = [f for f in st.session_state.fechas_recaudo if f is not None]
        fechas_str = ", ".join(f.strftime("%d/%m/%Y") for f in fechas_validas)
        st.caption(f"Fechas seleccionadas: **{fechas_str}**")
        st.markdown("<br>", unsafe_allow_html=True)

    # ── Botón de extracción ──────────────────────────────────────────────
    tipo_label = "Pagos Bancarios" if st.session_state.tipo_pago == "bancarios" else "Pagos por Recaudos"
    if st.button(f"⬇️  Extraer {tipo_label} del Libro de Banco", type="primary", use_container_width=True):
        with st.spinner("Leyendo el libro de banco y aplicando filtros..."):
            try:
                if st.session_state.tipo_pago == "bancarios":
                    df, resumen = extraer_pagos_bancarios(
                        st.session_state.archivo_libro,
                        st.session_state.archivo_corresponsal
                    )
                else:
                    fechas_sel = [f for f in st.session_state.get("fechas_recaudo", [None]) if f is not None]
                    df, resumen = extraer_pagos_recaudos(
                        st.session_state.archivo_libro,
                        fechas_sel
                    )
                st.session_state.df_area_banco = df
                st.success(f"✅ Extracción completada. {resumen}")
            except Exception as e:
                st.error(f"❌ Error en la extracción: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    # ── Tabla interactiva ────────────────────────────────────────────────
    if st.session_state.df_area_banco is not None:
        df = st.session_state.df_area_banco
        st.markdown("---")

        # Métricas rápidas
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total registros", len(df))
        c2.metric("Corresponsales", df["T_TRANSACCION"].notna().sum() if "T_TRANSACCION" in df.columns else 0)
        c3.metric("Entidades", df["ENTIDAD"].nunique() if "ENTIDAD" in df.columns else 0)
        if "VALOR" in df.columns:
            total = pd.to_numeric(df["VALOR"], errors="coerce").sum()
            c4.metric("Valor total", f"${total:,.0f}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="tabla-header"><h3>📊 Área de Banco</h3></div>', unsafe_allow_html=True)

        # Tabla editable
        # Columnas visibles en orden AREA DE BANCO
        cols_visibles = ["FECHA", "ENTIDAD", "CEDULA", "VALOR", "FRA",
                         "RECIBOS", "FECHA_DOCUMENTO", "REINCIDENTES_CB", "COMPENSACION"]
        cols_mostrar = [c for c in cols_visibles if c in df.columns]

        # Limpiar tipos para evitar conflictos en data_editor
        df_show = df[cols_mostrar].copy()
        for col in df_show.columns:
            if col in ["FECHA", "FECHA_DOCUMENTO"]:
                df_show[col] = pd.to_datetime(df_show[col], errors="coerce")
            elif col == "VALOR":
                df_show[col] = pd.to_numeric(df_show[col], errors="coerce")
            else:
                df_show[col] = df_show[col].astype(str).replace("None", "").replace("nan", "")

        df_editado = st.data_editor(
            df_show,
            use_container_width=True,
            num_rows="fixed",
            key="tabla_area_banco",
            column_config={
                "FECHA":           st.column_config.DateColumn("Fecha",           format="DD/MM/YYYY"),
                "FECHA_DOCUMENTO": st.column_config.DateColumn("Fecha Documento", format="DD/MM/YYYY"),
                "VALOR":           st.column_config.NumberColumn("Valor",         format="$%d"),
                "REINCIDENTES_CB": st.column_config.TextColumn("Reincidentes CB"),
                "COMPENSACION":    st.column_config.TextColumn("Compensación"),
            }
        )
        # Preservar columnas internas no visibles
        for col in df.columns:
            if col not in cols_mostrar:
                df_editado[col] = df[col].values
        st.session_state.df_area_banco = df_editado

        st.caption("Puedes editar celdas directamente en la tabla si necesitas hacer ajustes manuales.")


# ══════════════════════════════════════════════════════════════════════════
# TAB 3 — ALISTAR INFORMACIÓN
# ══════════════════════════════════════════════════════════════════════════
def render_alistar_informacion():
    from alistar import alistar_informacion

    if st.session_state.df_area_banco is None:
        st.warning("⚠️ Primero extrae los datos en la pestaña **2. Tabla de Pagos**.")
        return

    if not st.session_state.archivo_clientes:
        st.warning("⚠️ Primero carga el archivo de Clientes Activos en la pestaña **1. Carga de Archivos**.")
        return

    st.markdown("### Cruce con base de clientes activos")
    st.markdown("""
    <div class="alerta-info">
    🔗 Este proceso cruza las cédulas del Área de Banco con la base de clientes activos,
    generando la tabla final con entidad, fecha, company, cédula, factura, cuota y demás campos.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Botón para limpiar y volver a alistar
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        ejecutar = st.button("🔗  Alistar Información", type="primary", use_container_width=True)
    with col_btn2:
        if st.button("🔄  Limpiar", use_container_width=True):
            st.session_state.df_sheet1 = None
            st.session_state.df_sheet1_base = None
            st.session_state.df_sheet1_alertas = None
            st.session_state.distribuciones_confirmadas = None
            st.rerun()

    if ejecutar:
        with st.spinner("Cruzando información y procesando escenarios..."):
            try:
                # Limpiar marcas previas
                st.session_state["no_encontradas"] = []

                df_resultado, alertas = alistar_informacion(
                    st.session_state.df_area_banco,
                    st.session_state.archivo_clientes
                )
                # Guardar base y alertas en sesión
                st.session_state.df_sheet1_base    = df_resultado
                st.session_state.df_sheet1_alertas = alertas
                st.session_state.distribuciones_confirmadas = None

                # ── Marcar cédulas no encontradas en AREA DE BANCO ──────
                no_encontradas = st.session_state.get("no_encontradas", [])
                if no_encontradas:
                    df_banco = st.session_state.df_area_banco.copy()
                    for idx2 in no_encontradas:
                        if idx2 in df_banco.index:
                            df_banco.at[idx2, "RECIBOS"] = "NO EXISTE"
                            for col in ["COMPENSACION"]:
                                if col in df_banco.columns and (df_banco.at[idx2, col] is None or str(df_banco.at[idx2, col]) == "nan"):
                                    df_banco.at[idx2, col] = 0
                    st.session_state.df_area_banco = df_banco
                    st.warning(f"⚠️ {len(no_encontradas)} cédula(s) no encontradas en clientes activos. Marcadas como 'NO EXISTE' en la tabla.")

                if not alertas:
                    st.session_state.df_sheet1 = df_resultado
                    st.success(f"✅ Información alistada. {len(df_resultado)} registros procesados.")

            except Exception as e:
                st.error(f"❌ Error al alistar: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    # Si hay alertas pendientes mostrar el formulario de distribución
    if st.session_state.get("df_sheet1_alertas"):
        alertas = st.session_state.df_sheet1_alertas
        from alistar import _resolver_escenarios_multifactura
        df_extra = _resolver_escenarios_multifactura(alertas)

        # Cuando el usuario confirma, concatenar base + distribuidos
        if st.session_state.get("distribuciones_confirmadas") is not None:
            df_base  = st.session_state.get("df_sheet1_base", pd.DataFrame())
            df_extra = st.session_state["distribuciones_confirmadas"]
            df_final = pd.concat([df_base, df_extra], ignore_index=True)
            st.session_state.df_sheet1 = df_final
            st.success(f"✅ Información alistada. {len(df_final)} registros procesados.")

    # ── Tabla resultado ──────────────────────────────────────────────────
    if st.session_state.df_sheet1 is not None:
        df = st.session_state.df_sheet1
        st.markdown("---")

        # Cuadre de control
        df_banco      = st.session_state.df_area_banco
        total_extraido = pd.to_numeric(df_banco["VALOR"], errors="coerce").sum() if df_banco is not None and "VALOR" in df_banco.columns else 0

        # Total alistado (Sheet1)
        total_alistado = pd.to_numeric(df["CUOTA"], errors="coerce").sum() if "CUOTA" in df.columns else 0

        # Total no encontrado (filas con RECIBOS = "NO EXISTE")
        total_no_encontrado = 0
        if df_banco is not None and "RECIBOS" in df_banco.columns and "VALOR" in df_banco.columns:
            mask_no = df_banco["RECIBOS"].astype(str).str.strip() == "NO EXISTE"
            total_no_encontrado = pd.to_numeric(df_banco.loc[mask_no, "VALOR"], errors="coerce").sum()

        diferencia = total_extraido - total_alistado - total_no_encontrado

        st.markdown("#### 📊 Cuadre de Control")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total extraído",       f"${total_extraido:,.0f}")
        c2.metric("✅ Alistado",           f"${total_alistado:,.0f}")
        c3.metric("❌ No encontrado",      f"${total_no_encontrado:,.0f}")
        if abs(diferencia) < 1:
            c4.metric("Diferencia", "$0", delta="✅ Cuadra", delta_color="normal")
        else:
            c4.metric("Diferencia", f"${diferencia:,.0f}", delta="⚠️ Revisar", delta_color="inverse")

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros alistados", len(df))
        if "COMPANY" in df.columns:
            c2.metric("Empresas", df["COMPANY"].nunique())
        if "CUOTA" in df.columns:
            c3.metric("Total cuotas", f"${total_alistado:,.0f}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.data_editor(
            df,
            use_container_width=True,
            num_rows="fixed",
            key="tabla_sheet1",
            column_config={
                "FECHA": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                "FECHA_DOCUMENTO": st.column_config.DateColumn("Fecha Doc.", format="DD/MM/YYYY"),
                "CUOTA": st.column_config.NumberColumn("Cuota", format="$%d"),
            }
        )


# ══════════════════════════════════════════════════════════════════════════
# TAB 4 — GENERAR ARCHIVOS
# ══════════════════════════════════════════════════════════════════════════
def render_generar_archivos():
    if st.session_state.df_sheet1 is None:
        st.warning("⚠️ Primero completa el paso **3. Alistar Información**.")
        return

    st.markdown("### Generación de archivos finales")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 📄 Planos por empresa")
        st.caption("Genera CashReceipt, Services y PaymentMethod para cada empresa.")
        if st.button("📄  Crear Planos", type="primary", use_container_width=True, key="btn_planos"):
            from generar_planos import crear_planos
            with st.spinner("Generando planos..."):
                try:
                    resultado = crear_planos(
                        st.session_state.df_sheet1,
                        st.session_state.config,
                        st.session_state.df_area_banco,
                        st.session_state.tipo_pago
                    )
                    st.success(f"✅ {resultado}")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

    with col2:
        st.markdown("#### 📊 Compensaciones por fecha")
        st.caption("Genera un archivo de compensación por cada fecha única en Área de Banco.")
        if st.button("📊  Crear Compensaciones", type="primary", use_container_width=True, key="btn_comp"):
            from generar_compensaciones import crear_compensaciones
            with st.spinner("Generando compensaciones..."):
                try:
                    resultado = crear_compensaciones(
                        st.session_state.df_area_banco,
                        st.session_state.config,
                        st.session_state.tipo_pago
                    )
                    st.success(f"✅ {resultado}")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
