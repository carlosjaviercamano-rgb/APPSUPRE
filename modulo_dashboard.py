import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import io
from datetime import datetime

MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
         "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

def render():
    st.markdown("""
    <div class="module-header">
        <div class="module-icon">📊</div>
        <div>
            <h1>Dashboard</h1>
            <p>Informes y análisis gerencial</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Menú de informes
    if "submodulo_dashboard" not in st.session_state:
        st.session_state.submodulo_dashboard = None

    sub = st.session_state.submodulo_dashboard

    if sub is None:
        _render_menu()
    elif sub == "corresponsal":
        if st.button("← Volver", key="btn_volver_dash"):
            st.session_state.submodulo_dashboard = None
            st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
        render_corresponsal()
    elif sub == "cartera":
        if st.button("← Volver", key="btn_volver_dash2"):
            st.session_state.submodulo_dashboard = None
            st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
        st.info("🚧 Informe Planilla Cartera — próximamente.")


def _render_menu():
    st.markdown("### Selecciona el informe")
    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">🏦</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem">
                Informe Corresponsal</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Bancolombia — actualización mensual</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_corresponsal", use_container_width=True, type="primary"):
            st.session_state.submodulo_dashboard = "corresponsal"
            st.rerun()

    with col2:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">📋</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem">
                Informe Planilla Cartera</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Próximamente</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_cartera", use_container_width=True, type="primary"):
            st.session_state.submodulo_dashboard = "cartera"
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# INFORME CORRESPONSAL
# ══════════════════════════════════════════════════════════════════════════
def render_corresponsal():
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1e3a5f,#1e40af);border-radius:10px;
                padding:1rem 1.5rem;margin-bottom:1rem;">
        <h3 style="color:#fff;margin:0">🏦 Informe Corresponsal Bancolombia</h3>
        <p style="color:#93c5fd;margin:0.2rem 0 0 0;font-size:0.85rem">
            Actualización mensual del histórico de transferencias</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📂 1. Cargar y Procesar", "📊 2. Dashboard"])

    with tab1:
        _render_procesar_corresponsal()
    with tab2:
        _render_dashboard_corresponsal()


def _render_procesar_corresponsal():
    st.markdown("#### 📂 Archivos necesarios")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**📖 Libro de Banco**")
        libro = st.file_uploader("Libro de Banco", type=["xlsx","xlsm"],
                                  key="up_libro_corresponsal", label_visibility="collapsed")
        if libro:
            st.session_state["corr_libro"] = libro
            st.success(f"✅ {libro.name}")
        elif st.session_state.get("corr_libro"):
            st.success("✅ Ya cargado")

    with col2:
        st.markdown("**📊 Informe Corresponsal**")
        informe = st.file_uploader("Informe Corresponsal", type=["xlsx"],
                                    key="up_informe_corresponsal", label_visibility="collapsed")
        if informe:
            st.session_state["corr_informe"] = informe
            st.success(f"✅ {informe.name}")
        elif st.session_state.get("corr_informe"):
            st.success("✅ Ya cargado")

    st.markdown("---")
    st.markdown("#### 📅 Mes a trabajar")
    col_m, col_a = st.columns(2)
    with col_m:
        mes_sel = st.selectbox("Mes", MESES, key="corr_mes")
    with col_a:
        anio_sel = st.number_input("Año", min_value=2020, max_value=2035,
                                    value=datetime.now().year, key="corr_anio")

    st.markdown("<br>", unsafe_allow_html=True)

    # Botón procesar
    col_btn, col_lim = st.columns([3,1])
    with col_btn:
        procesar = st.button("⚙️ Procesar datos", type="primary",
                              use_container_width=True, key="btn_procesar_corr")
    with col_lim:
        if st.button("🔄 Limpiar", use_container_width=True, key="btn_limpiar_corr"):
            for k in ["corr_resultado","corr_stats","corr_libro","corr_informe"]:
                st.session_state.pop(k, None)
            st.rerun()

    if procesar:
        if not st.session_state.get("corr_libro") or not st.session_state.get("corr_informe"):
            st.error("❌ Debes cargar ambos archivos.")
        else:
            with st.spinner("Procesando..."):
                try:
                    stats = _procesar_corresponsal(
                        st.session_state["corr_libro"],
                        st.session_state["corr_informe"],
                        mes_sel, int(anio_sel)
                    )
                    st.session_state["corr_stats"]  = stats
                    st.session_state["corr_mes_sel"]  = mes_sel
                    st.session_state["corr_anio_sel"] = int(anio_sel)
                    st.success("✅ Datos procesados correctamente.")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

    # Mostrar resumen y campo comisión
    stats = st.session_state.get("corr_stats")
    if stats:
        st.markdown("---")
        st.markdown("#### 📊 Resumen del mes")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total transacciones", f"{stats['transacciones']:,}")
        c2.metric("Sin identificar",     f"{stats['sin_identificar']:,}")
        c3.metric("Identificadas",        f"{stats['identificadas']:,}")
        c4.metric("Reincidentes",         f"{stats['reincidentes']:,}")
        c5.metric("Nuevos",               f"{stats['nuevos']:,}")

        st.markdown("---")
        st.markdown("#### 💰 Valor de la comisión")
        comision = st.number_input(
            "Ingresa el valor de la comisión del mes ($):",
            min_value=0, value=0, step=10000,
            key="corr_comision", format="%d"
        )

        if st.button("✅ Generar Informe Excel", type="primary",
                     use_container_width=True, key="btn_generar_corr"):
            if comision <= 0:
                st.warning("⚠️ Ingresa un valor de comisión mayor a 0.")
            else:
                with st.spinner("Generando Excel..."):
                    try:
                        buf = _generar_excel_corresponsal(
                            st.session_state["corr_informe"],
                            stats, comision,
                            st.session_state["corr_mes_sel"],
                            st.session_state["corr_anio_sel"]
                        )
                        st.session_state["corr_excel"] = buf
                        st.session_state["corr_comision_val"] = comision
                        st.success("✅ Excel generado correctamente.")
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())

        if st.session_state.get("corr_excel"):
            mes = st.session_state.get("corr_mes_sel","")
            anio = st.session_state.get("corr_anio_sel","")
            st.download_button(
                label="⬇️  Descargar Informe Corresponsal actualizado",
                data=st.session_state["corr_excel"],
                file_name=f"INFORME_CORRESPONSAL_BANCOLOMBIA_{mes.upper()}_{anio}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_corr_excel"
            )


def _procesar_corresponsal(libro, informe, mes, anio):
    """Extrae movimientos del libro de banco y cruza con histórico."""
    # ── Leer libro de banco ──────────────────────────────────────────────
    df_banco = pd.read_excel(libro, sheet_name="BANCOLOMBIA SUPRECREDITO", header=0)
    df_banco.columns = [f"COL_{i}" for i in range(df_banco.shape[1])]

    # Col A = fecha, Col D = cédula, Col F = tipo transacción
    df_banco["COL_0"] = pd.to_datetime(df_banco["COL_0"], errors="coerce")

    # Filtrar por mes y año
    df_mes = df_banco[
        (df_banco["COL_0"].dt.month == MESES.index(mes) + 1) &
        (df_banco["COL_0"].dt.year  == anio)
    ].copy()

    # Filtrar por CONSIGNACION CORRESPONSAL CB en col F
    df_corr = df_mes[
        df_mes["COL_5"].astype(str).str.strip().str.upper() == "CONSIGNACION CORRESPONSAL CB"
    ].copy()

    total_transacciones = len(df_corr)

    # Sin identificar = col D vacía
    sin_iden = df_corr[
        df_corr["COL_3"].isna() | (df_corr["COL_3"].astype(str).str.strip().isin(["","nan"]))
    ]
    con_iden = df_corr[
        ~df_corr["COL_3"].isna() & (~df_corr["COL_3"].astype(str).str.strip().isin(["","nan"]))
    ]

    sin_identificar = len(sin_iden)
    identificadas   = len(con_iden)

    # Consolidar cédulas del mes (cuántos movimientos tuvo cada una)
    cedulas_mes = con_iden["COL_3"].astype(str).str.strip().value_counts().to_dict()

    # ── Leer histórico ───────────────────────────────────────────────────
    df_hist = pd.read_excel(informe, sheet_name="TRANSFERENCIAS CORRESPONSAL", header=0)
    df_hist.columns = ["CEDULA","HIST_ANT","TRANS_MES","HIST_ACT","OBS",
                       "VACIO"] + [f"X{i}" for i in range(df_hist.shape[1]-6)]
    df_hist["CEDULA"] = df_hist["CEDULA"].astype(str).str.strip()

    cedulas_historico = set(df_hist["CEDULA"].tolist())

    reincidentes = 0
    nuevos       = 0
    actualizaciones = {}  # cedula → {hist_ant, trans_mes, hist_act, obs}

    for ced, movs in cedulas_mes.items():
        if ced in cedulas_historico:
            reincidentes += 1
            fila = df_hist[df_hist["CEDULA"] == ced].iloc[0]
            hist_ant = int(fila["HIST_ACT"]) if pd.notna(fila["HIST_ACT"]) else 0
            actualizaciones[ced] = {
                "hist_ant":  hist_ant,
                "trans_mes": movs,
                "hist_act":  hist_ant + movs,
                "obs":       "REICIDENTE"
            }
        else:
            nuevos += 1
            actualizaciones[ced] = {
                "hist_ant":  0,
                "trans_mes": movs,
                "hist_act":  movs,
                "obs":       "NUEVO"
            }

    return {
        "transacciones":   total_transacciones,
        "sin_identificar": sin_identificar,
        "identificadas":   identificadas,
        "reincidentes":    reincidentes,
        "nuevos":          nuevos,
        "actualizaciones": actualizaciones,
        "cedulas_mes":     cedulas_mes,
        "mes":             mes,
        "anio":            anio,
    }


def _generar_excel_corresponsal(informe_file, stats, comision, mes, anio):
    """Actualiza el Excel del informe corresponsal."""
    # Leer el workbook original con data_only para obtener valores de fórmulas
    informe_file.seek(0)
    wb = openpyxl.load_workbook(informe_file, data_only=True)

    # ── HOJA TRANSFERENCIAS CORRESPONSAL ────────────────────────────────
    ws_trans = wb["TRANSFERENCIAS CORRESPONSAL"]

    # Actualizar encabezados fila 1 col C y D con nuevo mes
    mes_anterior = ws_trans["C1"].value or ""
    ws_trans["B1"] = mes_anterior  # El anterior pasa a B
    ws_trans["C1"] = f"TRANSACCIONES {mes.upper()} {anio}"
    ws_trans["D1"] = f"HISTORICO A {_ultimo_dia_mes(mes, anio)}"

    # Actualizar encabezado tabla resumen col G
    ws_trans["G1"] = f"TRANSACCIONES CORRESPONSAL MES DE {mes.upper()} {anio}"

    actualizaciones = stats["actualizaciones"]
    cedulas_en_hist = set()

    # Actualizar cédulas existentes
    for row in ws_trans.iter_rows(min_row=2):
        ced = str(row[0].value).strip() if row[0].value else ""
        if not ced or ced == "None":
            continue
        cedulas_en_hist.add(ced)

        if ced in actualizaciones:
            act = actualizaciones[ced]
            # B = historico anterior (lo que era C antes)
            row[1].value = act["hist_ant"]
            # C = transacciones del mes
            row[2].value = act["trans_mes"]
            # D = historico actual
            row[3].value = act["hist_act"]
            # E = observación
            row[4].value = act["obs"]
        else:
            # No tuvo movimientos este mes
            hist_act = row[3].value or 0
            row[1].value = hist_act  # B = lo que era D
            row[2].value = 0         # C = 0
            # D permanece igual
            row[4].value = None      # Limpiar observación

    # Agregar cédulas nuevas al final
    ultima_fila = ws_trans.max_row
    for ced, act in actualizaciones.items():
        if ced not in cedulas_en_hist:
            ultima_fila += 1
            ws_trans.cell(row=ultima_fila, column=1, value=int(ced) if ced.isdigit() else ced)
            ws_trans.cell(row=ultima_fila, column=2, value=act["hist_ant"])
            ws_trans.cell(row=ultima_fila, column=3, value=act["trans_mes"])
            ws_trans.cell(row=ultima_fila, column=4, value=act["hist_act"])
            ws_trans.cell(row=ultima_fila, column=5, value=act["obs"])

    # Actualizar tabla resumen fila 3 cols G-L
    ws_trans["G3"] = stats["transacciones"]
    ws_trans["H3"] = stats["sin_identificar"]
    ws_trans["I3"] = stats["identificadas"]
    ws_trans["J3"] = stats["reincidentes"]
    ws_trans["K3"] = stats["nuevos"]
    ws_trans["L3"] = comision

    # ── HOJA COMISIÓN CORRESPONSAL ───────────────────────────────────────
    ws_com = wb["COMISIÓN CORRESPONSAL"]

    # Encontrar última fila con datos
    ultima_com = 1
    for row in ws_com.iter_rows(min_row=2):
        if row[0].value and str(row[0].value).strip() not in ["","None","TOTAL "]:
            ultima_com = row[0].row

    # Calcular variaciones
    fila_ant = None
    for row in ws_com.iter_rows(min_row=2):
        if row[0].value and str(row[0].value).strip() not in ["","None","TOTAL "]:
            fila_ant = row

    var_mes_ant_val = 0
    var_mes_ant_pct = 0
    if fila_ant and fila_ant[3].value:
        try:
            com_ant = float(str(fila_ant[3].value).replace(",","").replace("$","").strip() or 0)
        except (ValueError, TypeError):
            com_ant = 0
        var_mes_ant_val = comision - com_ant
        var_mes_ant_pct = var_mes_ant_val / com_ant if com_ant else 0

    # Variación vs año anterior (mismo mes)
    mes_idx = MESES.index(mes) + 1
    var_anio_val = 0
    var_anio_pct = 0
    for row in ws_com.iter_rows(min_row=2):
        if (str(row[0].value).strip() == str(anio - 1) and
            str(row[1].value).strip().lower().startswith(mes.lower()[:3])):
            if row[3].value:
                com_anio_ant = float(str(row[3].value).replace(",","").replace("$","").strip() or 0)
                var_anio_val = comision - com_anio_ant
                var_anio_pct = var_anio_val / com_anio_ant if com_anio_ant else 0
            break

    # Insertar nueva fila antes del TOTAL
    nueva_fila = ultima_com + 1
    ws_com.insert_rows(nueva_fila)
    ws_com.cell(row=nueva_fila, column=1, value=anio)
    ws_com.cell(row=nueva_fila, column=2, value=mes)
    ws_com.cell(row=nueva_fila, column=3, value=stats["transacciones"])
    ws_com.cell(row=nueva_fila, column=4, value=comision)
    ws_com.cell(row=nueva_fila, column=5, value=var_mes_ant_val)
    ws_com.cell(row=nueva_fila, column=6, value=round(var_mes_ant_pct, 4))
    ws_com.cell(row=nueva_fila, column=7, value=var_anio_val)
    ws_com.cell(row=nueva_fila, column=8, value=round(var_anio_pct, 4))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _ultimo_dia_mes(mes, anio):
    """Retorna el último día del mes como string DD-MM-YYYY."""
    import calendar
    mes_idx = MESES.index(mes) + 1
    ultimo  = calendar.monthrange(anio, mes_idx)[1]
    return f"{ultimo:02d}-{mes_idx:02d}-{anio}"


# ══════════════════════════════════════════════════════════════════════════
# DASHBOARD VISUAL
# ══════════════════════════════════════════════════════════════════════════
def _render_dashboard_corresponsal():
    stats = st.session_state.get("corr_stats")
    if not stats:
        st.warning("⚠️ Primero procesa los datos en la pestaña **1. Cargar y Procesar**.")
        return

    mes  = stats["mes"]
    anio = stats["anio"]
    com  = st.session_state.get("corr_comision_val", 0)

    st.markdown(f"### Corresponsal Bancolombia — {mes} {anio}")

    # KPIs
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total transacciones", f"{stats['transacciones']:,}")
    c2.metric("Sin identificar",     f"{stats['sin_identificar']:,}")
    c3.metric("Identificadas",        f"{stats['identificadas']:,}")
    c4.metric("Reincidentes",         f"{stats['reincidentes']:,}")
    c5.metric("Nuevos",               f"{stats['nuevos']:,}")
    c6.metric("Comisión",             f"${com:,.0f}" if com else "—")

    st.markdown("---")

    # Tabla de cédulas del mes
    st.markdown("#### 📋 Cédulas del mes")
    actualizaciones = stats.get("actualizaciones", {})
    if actualizaciones:
        df_tabla = pd.DataFrame([
            {
                "Cédula":       ced,
                "Hist. anterior": act["hist_ant"],
                "Trans. mes":   act["trans_mes"],
                "Hist. actual": act["hist_act"],
                "Observación":  act["obs"]
            }
            for ced, act in actualizaciones.items()
        ])
        st.dataframe(df_tabla, use_container_width=True, height=400)
