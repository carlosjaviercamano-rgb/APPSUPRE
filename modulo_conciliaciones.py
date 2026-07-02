import streamlit as st
import pandas as pd
import io


def render():
    st.markdown("""
    <div class="module-header">
        <div class="module-icon">🔍</div>
        <div>
            <h1>Conciliaciones</h1>
            <p>Conciliación bancaria, QR & Credibanco y cuentas puentes/transitorias</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if "submodulo_conciliacion" not in st.session_state:
        st.session_state.submodulo_conciliacion = None

    sub = st.session_state.submodulo_conciliacion

    if sub is None:
        _render_menu_submodulos()
    elif sub == "bancaria":
        _render_volver()
        render_bancaria()
    elif sub == "qr_credibanco":
        _render_volver()
        render_qr_credibanco()
    elif sub == "puentes":
        _render_volver()
        render_puentes()


def _render_menu_submodulos():
    st.markdown("### Selecciona el tipo de conciliación")
    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">🏦</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem;font-size:1rem">
                Conciliación Bancaria</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Libro Auxiliar + Extracto bancario</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_bancaria", use_container_width=True, type="primary"):
            st.session_state.submodulo_conciliacion = "bancaria"
            st.rerun()

    with col2:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">📱</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem;font-size:1rem">
                Conciliación QR & Credibanco</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Libro Auxiliar + Libro de Banco</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_qr", use_container_width=True, type="primary"):
            st.session_state.submodulo_conciliacion = "qr_credibanco"
            st.rerun()

    with col3:
        st.markdown("""
        <div style="background:#1a1f2e;border:1px solid #2d3548;border-radius:12px;
                    padding:1.5rem;text-align:center;">
            <div style="font-size:2.5rem">🔄</div>
            <div style="font-weight:700;color:#fff;margin-top:0.5rem;font-size:1rem">
                Cuentas Puentes/Transitorias</div>
            <div style="color:#64748b;font-size:0.8rem;margin-top:0.3rem">
                Solo Libro Auxiliar</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_puentes", use_container_width=True, type="primary"):
            st.session_state.submodulo_conciliacion = "puentes"
            st.rerun()


def _render_volver():
    if st.button("← Volver a Conciliaciones", key="btn_volver_conciliacion"):
        st.session_state.submodulo_conciliacion = None
        st.rerun()
    st.markdown("<br>", unsafe_allow_html=True)


MESES_PUENTES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                 "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]


# ══════════════════════════════════════════════════════════════════════════
# SECCIÓN COMPARTIDA: LIBRO AUXILIAR
# ══════════════════════════════════════════════════════════════════════════

def render_cargue_auxiliares(key_prefix=""):
    st.markdown("#### 📂 Libro Auxiliar")
    st.caption("Sube hasta 15 auxiliares individuales. La app los unirá automáticamente.")

    key_archivos = f"{key_prefix}_auxiliares"
    if key_archivos not in st.session_state:
        st.session_state[key_archivos] = []

    archivos = st.file_uploader(
        "Selecciona los auxiliares (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"up_{key_prefix}_auxiliares",
        label_visibility="collapsed"
    )

    if archivos:
        if len(archivos) > 15:
            st.warning("⚠️ Máximo 15 auxiliares. Se tomarán los primeros 15.")
            archivos = archivos[:15]
        st.session_state[key_archivos] = archivos

    archivos_cargados = st.session_state.get(key_archivos, [])
    key_df = f"{key_prefix}_df_auxiliar"

    if archivos_cargados:
        st.success(f"✅ {len(archivos_cargados)} auxiliar(es) cargado(s):")
        for arch in archivos_cargados:
            st.caption(f"   📄 {arch.name}")

        if st.button("🔗 Unir Auxiliares", key=f"btn_unir_{key_prefix}",
                     type="primary", use_container_width=True):
            with st.spinner("Uniendo auxiliares..."):
                try:
                    df_completo, df_conciliar = _unir_auxiliares(archivos_cargados)
                    st.session_state[key_df]                      = df_conciliar
                    st.session_state[f"{key_prefix}_df_completo"] = df_completo
                    key_bytes = f"{key_prefix}_auxiliar_bytes"
                    if key_bytes in st.session_state:
                        del st.session_state[key_bytes]
                    st.success(f"✅ Libro Auxiliar creado: {len(df_completo):,} registros.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

        df_auxiliar = st.session_state.get(key_df)
        if df_auxiliar is not None:
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            df_completo_prev = st.session_state.get(f"{key_prefix}_df_completo", df_auxiliar)
            c1.metric("Total registros", f"{len(df_auxiliar):,}")
            c2.metric("Empresas", df_auxiliar["empresa"].nunique() if "empresa" in df_auxiliar.columns else 0)
            c3.metric("Auxiliares", df_auxiliar["Source.Name"].nunique() if "Source.Name" in df_auxiliar.columns else 0)

            with st.expander("👁️ Vista previa"):
                st.dataframe(df_completo_prev.head(20), use_container_width=True)

            key_bytes    = f"{key_prefix}_auxiliar_bytes"
            key_completo = f"{key_prefix}_df_completo"
            df_completo  = st.session_state.get(key_completo, df_auxiliar)

            if key_bytes not in st.session_state:
                cols_decimales = ["valor","baseimpuesto","saldoanterior","debito","credito","saldoactual"]
                df_descarga = df_completo.copy()
                for col in cols_decimales:
                    if col in df_descarga.columns:
                        df_descarga[col] = pd.to_numeric(df_descarga[col], errors="coerce").round(2)
                buf = io.BytesIO()
                df_descarga.to_excel(buf, index=False, sheet_name="AUXILIARES")
                st.session_state[key_bytes] = buf.getvalue()

            st.download_button(
                label="⬇️  Descargar Libro Auxiliar unificado",
                data=st.session_state[key_bytes],
                file_name="LIBRO_AUXILIAR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_auxiliar_{key_prefix}"
            )
            return df_auxiliar

    df_auxiliar = st.session_state.get(key_df)
    if df_auxiliar is not None:
        key_bytes = f"{key_prefix}_auxiliar_bytes"
        if key_bytes in st.session_state:
            st.markdown("---")
            st.info("📋 Libro Auxiliar previamente generado disponible.")
            st.download_button(
                label="⬇️  Descargar Libro Auxiliar unificado",
                data=st.session_state[key_bytes],
                file_name="LIBRO_AUXILIAR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_auxiliar_prev_{key_prefix}"
            )
        return df_auxiliar

    return None


def _unir_auxiliares(archivos):
    cols_conciliacion = {"id","empresa","codigocuenta","identificacion","valor"}
    frames_completos  = []
    frames_conciliar  = []

    for archivo in archivos:
        try:
            df = pd.read_excel(archivo, sheet_name=0)
            df.columns = [c.strip().lower() for c in df.columns]
            empresa = ""
            if "empresa" in df.columns and not df.empty:
                empresa = str(df["empresa"].iloc[0]).strip()
            df.insert(0, "Source.Name", empresa if empresa else archivo.name)
            frames_completos.append(df)
            cols_disp = [c for c in df.columns if c in cols_conciliacion or c == "Source.Name"]
            frames_conciliar.append(df[cols_disp].copy())
        except Exception as e:
            raise ValueError(f"Error leyendo {archivo.name}: {str(e)}")

    if not frames_completos:
        raise ValueError("No se pudieron leer los auxiliares.")

    return pd.concat(frames_completos, ignore_index=True), pd.concat(frames_conciliar, ignore_index=True)


# ══════════════════════════════════════════════════════════════════════════
# SUBMÓDULO 1: CONCILIACIÓN BANCARIA
# ══════════════════════════════════════════════════════════════════════════

def render_bancaria():
    from conciliacion_bancaria import EMPRESAS_CUENTAS, MESES as MESES_CB

    st.markdown("""
    <div style="background:linear-gradient(135deg,#1e3a5f,#1e40af);border-radius:10px;
                padding:1rem 1.5rem;margin-bottom:1rem;">
        <h3 style="color:#fff;margin:0">🏦 Conciliación Bancaria</h3>
        <p style="color:#93c5fd;margin:0.2rem 0 0 0;font-size:0.85rem">
            Libro Auxiliar + Extracto bancario</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📂 1. Cargar Archivos", "🔍 2. Conciliar"])

    with tab1:
        st.markdown("#### ⚙️ Configuración")
        col1, col2, col3 = st.columns(3)

        empresas = list(EMPRESAS_CUENTAS.keys())
        with col1:
            empresa_sel = st.selectbox("Empresa", empresas, key="banc_empresa")
        cuentas_emp  = EMPRESAS_CUENTAS.get(empresa_sel, [])
        nombres_ctas = [c[0] for c in cuentas_emp]
        with col2:
            cuenta_sel = st.selectbox("Cuenta bancaria", nombres_ctas, key="banc_cuenta")
        codigo_cuenta = next((c[1] for c in cuentas_emp if c[0] == cuenta_sel), None)
        with col3:
            mes_sel = st.selectbox("Mes", MESES_CB, key="banc_mes")
        st.caption(f"Código cuenta: **{codigo_cuenta}**")

        st.markdown("---")
        st.markdown("#### 📂 Auxiliares")
        st.caption("Sube hasta 15 auxiliares (opcional). Si no se encuentra el auxiliar, igual podrás conciliar y todo el extracto quedará como no registrado.")

        _, col_lim_aux = st.columns([3, 1])
        with col_lim_aux:
            if st.button("🔄 Limpiar archivos", key="limpiar_aux_bancaria", use_container_width=True):
                for k in ["banc_auxiliares","banc_extracto","banc_filtro_ok",
                          "banc_df_banco","banc_df_aux","banc_df_aux_orig",
                          "banc_resumen_filtro","banc_resumen","banc_excel"]:
                    st.session_state.pop(k, None)
                st.rerun()

        archivos_aux = st.file_uploader(
            "Selecciona los auxiliares (.xlsx)",
            type=["xlsx"], accept_multiple_files=True,
            key="up_banc_auxiliares", label_visibility="collapsed"
        )
        if archivos_aux:
            st.session_state["banc_auxiliares"] = archivos_aux[:15]

        cargados = st.session_state.get("banc_auxiliares", [])
        if cargados:
            st.success(f"✅ {len(cargados)} auxiliar(es) listos.")
            for a in cargados:
                st.caption(f"   📄 {a.name}")

        st.markdown("---")
        st.markdown("#### 🏦 Extracto Bancario")
        st.caption("Formato detectado automáticamente: Bancolombia CSV, Davivienda XLSX, Occidente/Bogotá XLS.")

        extracto = st.file_uploader(
            "Extracto bancario",
            type=["csv","xlsx","xls","XLS"],
            key="up_banc_extracto", label_visibility="collapsed"
        )
        if extracto:
            st.session_state["banc_extracto"] = extracto
            st.success(f"✅ {extracto.name}")
        elif st.session_state.get("banc_extracto"):
            st.success(f"✅ {st.session_state['banc_extracto'].name}")

        if st.session_state.get("banc_extracto"):
            st.info("👉 Ve a la pestaña **2. Conciliar** para filtrar y conciliar.")

    with tab2:
        _render_conciliar_bancaria()


def _render_conciliar_bancaria():
    from conciliacion_bancaria import filtrar_datos, conciliar, generar_excel
    from conciliacion_bancaria import MESES as MESES_CB, EMPRESAS_CUENTAS

    archivos_aux = st.session_state.get("banc_auxiliares", [])
    extracto     = st.session_state.get("banc_extracto")
    empresa      = st.session_state.get("banc_empresa", "")
    cuenta_nom   = st.session_state.get("banc_cuenta", "")
    mes          = st.session_state.get("banc_mes", "")
    cuentas_emp  = EMPRESAS_CUENTAS.get(empresa, [])
    cod_cuenta   = next((c[1] for c in cuentas_emp if c[0] == cuenta_nom), None)
    mes_idx      = MESES_CB.index(mes) if mes in MESES_CB else 0

    if not extracto:
        st.warning("⚠️ Primero carga el extracto bancario en la pestaña **1. Cargar Archivos**.")
        return

    st.markdown(f"### {empresa} | {cuenta_nom} | {mes}")

    col_f, col_c, col_l = st.columns([2, 2, 1])

    with col_f:
        btn_filtrar = st.button(
            "⚙️ Filtrar datos",
            type="primary",
            use_container_width=True,
            key="btn_filtrar_bancaria",
            disabled=bool(st.session_state.get("banc_filtro_ok"))
        )

    filtro_ok = st.session_state.get("banc_filtro_ok", False)

    with col_c:
        btn_conciliar = st.button(
            "🔍 Conciliar",
            type="primary",
            use_container_width=True,
            key="btn_conciliar_bancaria",
            disabled=not filtro_ok
        )

    with col_l:
        if st.button("🔄 Limpiar", use_container_width=True, key="btn_limpiar_bancaria"):
            for k in ["banc_filtro_ok","banc_df_banco","banc_df_aux",
                      "banc_df_aux_orig","banc_resumen_filtro","banc_resumen","banc_excel"]:
                st.session_state.pop(k, None)
            st.rerun()

    if btn_filtrar:
        with st.spinner("Filtrando datos... por favor espera."):
            try:
                for a in archivos_aux:
                    a.seek(0)
                extracto.seek(0)
                df_banco, df_aux, df_aux_orig, resumen_filtro = filtrar_datos(
                    archivos_aux, extracto, empresa, cod_cuenta, mes_idx
                )
                st.session_state["banc_df_banco"]       = df_banco
                st.session_state["banc_df_aux"]         = df_aux
                st.session_state["banc_df_aux_orig"]    = df_aux_orig
                st.session_state["banc_resumen_filtro"] = resumen_filtro
                st.session_state["banc_filtro_ok"]      = True
                st.session_state.pop("banc_resumen", None)
                st.session_state.pop("banc_excel",   None)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al filtrar: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    resumen_filtro = st.session_state.get("banc_resumen_filtro")
    if resumen_filtro:
        if resumen_filtro.get("sin_auxiliar"):
            st.warning(
                "⚠️ No se encontraron movimientos en el auxiliar para "
                f"**{empresa} / cuenta {cod_cuenta} / {mes}**. "
                "Puedes conciliar igual: todos los movimientos del extracto "
                "aparecerán como **no registrados en el auxiliar**."
            )
        else:
            st.success("✅ Datos filtrados correctamente. Ya puedes conciliar.")

        st.markdown("#### 📋 Resumen de datos filtrados")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Movimientos banco",  f"{resumen_filtro['n_banco']:,}")
        c2.metric("Movimientos aux",    f"{resumen_filtro['n_aux']:,}")
        c3.metric("Débito banco",       f"${resumen_filtro['total_deb_banco']:,.2f}")
        c4.metric("Crédito banco",      f"${resumen_filtro['total_cre_banco']:,.2f}")
        c5.metric("Débito aux",         f"${resumen_filtro['total_deb_aux']:,.2f}")
        c6.metric("Crédito aux",        f"${resumen_filtro['total_cre_aux']:,.2f}")

    if btn_conciliar:
        df_banco    = st.session_state.get("banc_df_banco")
        df_aux      = st.session_state.get("banc_df_aux")
        df_aux_orig = st.session_state.get("banc_df_aux_orig")
        if df_banco is None or df_aux is None:
            st.error("❌ No hay datos filtrados. Primero haz clic en Filtrar datos.")
            return
        with st.spinner("Conciliando... esto toma solo unos segundos."):
            try:
                partidas, df_b, df_a, resumen = conciliar(df_banco, df_aux)
                excel_bytes = generar_excel(
                    partidas, df_b, df_a, resumen,
                    empresa, cuenta_nom, mes, df_aux_orig
                )
                st.session_state["banc_resumen"] = resumen
                st.session_state["banc_excel"]   = excel_bytes
                st.success("✅ Conciliación completada.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al conciliar: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    resumen = st.session_state.get("banc_resumen")
    if resumen is not None:
        st.markdown("---")
        st.markdown("#### 📊 Resultado de Conciliación")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Saldo aux débito",    f"${resumen['saldo_aux_deb']:,.2f}")
        c2.metric("Saldo aux crédito",   f"${resumen['saldo_aux_cred']:,.2f}")
        c3.metric("Saldo banco débito",  f"${resumen['saldo_banco_deb']:,.2f}")
        c4.metric("Saldo banco crédito", f"${resumen['saldo_banco_cred']:,.2f}")
        c5.metric("✅ Cruzados banco",   f"{resumen['conciliados_banco']:,}")
        c6.metric("✅ Cruzados aux",     f"{resumen['conciliados_aux']:,}")

        dif_d = resumen["diferencia_deb"]
        dif_c = resumen["diferencia_cred"]
        if abs(dif_d) < 0.01 and abs(dif_c) < 0.01:
            st.success("✅ DIFERENCIA = $0 — Cuenta conciliada correctamente.")
        else:
            st.warning(f"⚠️ Diferencia déb: ${dif_d:,.2f} | Diferencia cré: ${dif_c:,.2f}")
            st.caption(f"Solo banco: {resumen['n_solo_banco']} | "
                       f"Solo aux: {resumen['n_solo_aux']} | "
                       f"Reclasif: {resumen['n_reclasif']}")

        if st.session_state.get("banc_excel"):
            from datetime import date, timedelta
            ayer = date.today() - timedelta(days=1)
            fecha_str = ayer.strftime("%d_%m_%Y")
            cuenta_limpia = cuenta_nom.replace("/", "-").replace("\\", "-").strip()
            nombre_archivo = f"CONCILIACION_{cuenta_limpia}_{fecha_str}.xlsx"

            cfg = st.session_state.get("config", {})
            nombre_lower = cuenta_nom.lower()
            if "bancolombia" in nombre_lower:
                ruta_auto = cfg.get("ruta_conc_bancolombia", "")
            elif "davivienda" in nombre_lower:
                ruta_auto = cfg.get("ruta_conc_davivienda", "")
            elif "occidente" in nombre_lower:
                ruta_auto = cfg.get("ruta_conc_occidente", "")
            elif "bogota" in nombre_lower or "bogotá" in nombre_lower:
                ruta_auto = cfg.get("ruta_conc_bogota", "")
            else:
                ruta_auto = ""

            if ruta_auto:
                try:
                    import os
                    os.makedirs(ruta_auto, exist_ok=True)
                    ruta_completa = os.path.join(ruta_auto, nombre_archivo)
                    with open(ruta_completa, "wb") as f:
                        f.write(st.session_state["banc_excel"])
                    st.success(f"💾 Guardado automáticamente en: {ruta_completa}")
                except Exception as e:
                    st.warning(f"⚠️ No se pudo guardar automáticamente: {str(e)}")

            st.download_button(
                label="⬇️  Descargar Excel (CONCILIACION + BANCO + AUXILIAR)",
                data=st.session_state["banc_excel"],
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_bancaria"
            )


# ══════════════════════════════════════════════════════════════════════════
# SUBMÓDULO 3: CUENTAS PUENTES / TRANSITORIAS
# ══════════════════════════════════════════════════════════════════════════

def _render_conciliar_puentes():
    archivos = st.session_state.get("puentes_auxiliares", [])
    if not archivos:
        st.warning("⚠️ Primero carga los auxiliares en la pestaña **1. Cargar Archivos**.")
        return

    st.markdown("### 🔍 Conciliación Cuentas Puentes / Transitorias")
    st.caption(f"📂 {len(archivos)} auxiliar(es) cargado(s). Ingresa la cuenta y el mes a conciliar.")

    # ── Fila: cuenta + mes + botones ─────────────────────────────────────
    col_inp, col_mes, col_f, col_c, col_lim = st.columns([2, 1, 1, 1, 1])

    with col_inp:
        codigo_cuenta = st.text_input(
            "Código de cuenta (codigocuenta):",
            value=st.session_state.get("puentes_cuenta", ""),
            placeholder="Ej: 141299011",
            key="input_cuenta_puentes"
        )

    with col_mes:
        mes_sel = st.selectbox(
            "Mes:",
            MESES_PUENTES,
            index=MESES_PUENTES.index(st.session_state.get("puentes_mes", MESES_PUENTES[0])),
            key="sel_mes_puentes"
        )
        st.session_state["puentes_mes"] = mes_sel

    filtro_ok = st.session_state.get("puentes_filtro_ok", False)

    with col_f:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_filtrar = st.button(
            "⚙️ Filtrar",
            type="primary",
            use_container_width=True,
            key="btn_filtrar_puentes",
            disabled=filtro_ok
        )

    with col_c:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_conciliar = st.button(
            "🔍 Conciliar",
            type="primary",
            use_container_width=True,
            key="btn_conciliar_puentes",
            disabled=not filtro_ok
        )

    with col_lim:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Limpiar", use_container_width=True, key="limpiar_conciliar_puentes"):
            for k in ["puentes_df_filtrado", "puentes_cuenta", "puentes_mes",
                      "puentes_resultado_bytes", "puentes_filtro_ok",
                      "puentes_resumen_filtro", "puentes_resultado", "puentes_resultado_df"]:
                st.session_state.pop(k, None)
            st.rerun()

    # ── PASO 1: FILTRAR ───────────────────────────────────────────────────
    if btn_filtrar:
        if not codigo_cuenta.strip():
            st.warning("⚠️ Ingresa un código de cuenta.")
        else:
            st.session_state["puentes_cuenta"] = codigo_cuenta.strip()
            mes_num = MESES_PUENTES.index(mes_sel) + 1
            with st.spinner("Leyendo auxiliares y filtrando por cuenta y mes..."):
                try:
                    df_filtrado = _leer_y_filtrar_por_cuenta(
                        archivos, codigo_cuenta.strip(), mes_num
                    )
                    if df_filtrado is None or df_filtrado.empty:
                        st.warning(
                            f"⚠️ No se encontraron registros para la cuenta "
                            f"**{codigo_cuenta}** en **{mes_sel}**."
                        )
                    else:
                        resumen_filtro = {
                            "n_registros": len(df_filtrado),
                            "cuenta":      codigo_cuenta.strip(),
                            "mes":         mes_sel,
                            "suma_total":  round(
                                pd.to_numeric(df_filtrado["valor"], errors="coerce").fillna(0).sum(), 2
                            ) if "valor" in df_filtrado.columns else 0,
                        }
                        st.session_state["puentes_df_filtrado"]    = df_filtrado
                        st.session_state["puentes_resumen_filtro"] = resumen_filtro
                        st.session_state["puentes_filtro_ok"]      = True
                        st.session_state.pop("puentes_resultado",    None)
                        st.session_state.pop("puentes_resultado_df", None)
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

    # ── Mostrar resumen del filtro ────────────────────────────────────────
    resumen_filtro = st.session_state.get("puentes_resumen_filtro")
    if resumen_filtro:
        st.success("✅ Datos filtrados correctamente. Ya puedes conciliar.")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Cuenta",     resumen_filtro["cuenta"])
        c2.metric("Mes",        resumen_filtro.get("mes", ""))
        c3.metric("Registros",  f"{resumen_filtro['n_registros']:,}")
        c4.metric("Suma total", f"${resumen_filtro['suma_total']:,.2f}")

    # ── PASO 2: CONCILIAR ─────────────────────────────────────────────────
    if btn_conciliar:
        df_filtrado   = st.session_state.get("puentes_df_filtrado")
        cuenta_actual = st.session_state.get("puentes_cuenta", "")
        if df_filtrado is None:
            st.error("❌ No hay datos filtrados. Primero haz clic en Filtrar.")
            return
        with st.spinner("Conciliando... esto toma solo unos segundos."):
            try:
                df_resultado = _ejecutar_conciliacion_puentes(df_filtrado, cuenta_actual)
                st.session_state["puentes_resultado_df"] = df_resultado
                st.session_state["puentes_resultado"]    = True
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al conciliar: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    # ── Mostrar resultado ─────────────────────────────────────────────────
    if st.session_state.get("puentes_resultado"):
        df_resultado  = st.session_state.get("puentes_resultado_df")
        cuenta_actual = st.session_state.get("puentes_cuenta", "")
        if df_resultado is not None:
            _mostrar_resultado_puentes(df_resultado, cuenta_actual)


def _leer_y_filtrar_por_cuenta(archivos, codigo_cuenta, mes_num=None):
    """
    Lee los auxiliares y filtra por cuenta. Si mes_num se proporciona,
    también filtra por ese mes (1=enero … 12=diciembre).
    """
    frames = []
    for archivo in archivos:
        try:
            df = pd.read_excel(archivo, sheet_name=0)
            df.columns = [c.strip().lower() for c in df.columns]
            col_cuenta = "codigocuenta"
            if col_cuenta not in df.columns:
                continue

            # Filtro por cuenta
            filtrado = df[df[col_cuenta].astype(str).str.strip() == str(codigo_cuenta).strip()].copy()
            if filtrado.empty:
                continue

            # Filtro por mes si se especificó
            if mes_num is not None:
                col_fec = next((c for c in filtrado.columns if c == "fecha"), None)
                if col_fec:
                    filtrado[col_fec] = pd.to_datetime(filtrado[col_fec], dayfirst=True, errors="coerce")
                    filtrado = filtrado[filtrado[col_fec].dt.month == mes_num].copy()
                if filtrado.empty:
                    continue

            empresa = str(filtrado["empresa"].iloc[0]).strip() if "empresa" in filtrado.columns else archivo.name
            filtrado.insert(0, "Source.Name", empresa)
            frames.append(filtrado)
            del df
        except Exception as e:
            raise ValueError(f"Error leyendo {archivo.name}: {str(e)}")

    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def _ejecutar_conciliacion_puentes(df_filtrado, codigo_cuenta):
    """Aplica la lógica de conciliación y retorna el DataFrame resultado."""
    from itertools import combinations

    col_valor = "valor"
    col_iden  = "identificacion"
    col_id    = "id"

    df = df_filtrado.copy()
    df[col_valor] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0)
    df[col_iden]  = df[col_iden].astype(str).str.strip().replace({"nan":"","None":"","NaN":""})

    total_movimientos = len(df)
    suma_total        = round(df[col_valor].sum(), 2)

    df["concilia_con_id"] = ""

    if abs(suma_total) < 0.01:
        df["concilia_con_id"] = "CONCILIADA"
        return df

    def _buscar_pares_internos(grupo):
        pares = set()
        pos    = grupo[grupo[col_valor] > 0]
        neg    = grupo[grupo[col_valor] < 0]
        usados = set()
        for ip, rp in pos.iterrows():
            if ip in pares: continue
            vp = round(rp[col_valor], 2)
            for inn, rn in neg.iterrows():
                if inn in usados: continue
                vn = round(rn[col_valor], 2)
                if abs(vp + vn) < 0.01:
                    pares.add(ip); pares.add(inn); usados.add(inn)
                    break
        return pares

    indices_sn_interno = set()
    for ced, grupo in df[df[col_iden] != ""].groupby(col_iden):
        pares = _buscar_pares_internos(grupo)
        indices_sn_interno.update(pares)
    df.loc[list(indices_sn_interno), "concilia_con_id"] = "SIN NOVEDAD"

    df_resto     = df[df["concilia_con_id"] == ""]
    saldos_resto = df_resto[df_resto[col_iden] != ""].groupby(col_iden)[col_valor].sum().round(2)

    ced_sin_nov2 = set(saldos_resto[abs(saldos_resto) < 0.01].index)
    mask_sn2 = (df[col_iden].isin(ced_sin_nov2)) & (df["concilia_con_id"] == "")
    df.loc[mask_sn2, "concilia_con_id"] = "SIN NOVEDAD"

    ced_con_sal = {ced: round(sal, 2) for ced, sal in saldos_resto.items()
                   if abs(sal) >= 0.01 and ced not in ced_sin_nov2}

    saldos_pend  = dict(ced_con_sal)
    concilia_map = {}

    encontrado = True
    while encontrado and len(saldos_pend) >= 2:
        encontrado = False
        items = list(saldos_pend.items())
        for r in range(2, min(len(items) + 1, 6)):
            if encontrado: break
            for combo in combinations(items, r):
                ceds = [c for c, _ in combo]
                vals = [v for _, v in combo]
                if abs(round(sum(vals), 2)) < 0.01:
                    for ced in ceds:
                        saldo_ced = saldos_pend[ced]
                        opuestas  = [c for c in ceds if c != ced and saldos_pend[c] * saldo_ced < 0]
                        if not opuestas:
                            opuestas = [c for c in ceds if c != ced]
                        ced_op   = opuestas[0]
                        saldo_op = saldos_pend[ced_op]
                        filas_op  = df[(df[col_iden] == ced_op) & (df["concilia_con_id"] == "")]
                        mov_exact = filas_op[abs(filas_op[col_valor] - saldo_op) < 0.01]
                        if not mov_exact.empty:
                            id_ref = int(mov_exact[col_id].iloc[0])
                        else:
                            mov_s  = filas_op[filas_op[col_valor] * saldo_ced < 0]
                            id_ref = int(mov_s[col_id].iloc[0]) if not mov_s.empty \
                                     else int(filas_op[col_id].iloc[0]) if not filas_op.empty \
                                     else ced_op
                        concilia_map[ced] = (f"Concilia con ID {id_ref}", saldo_ced)
                    for ced in ceds:
                        del saldos_pend[ced]
                    encontrado = True
                    break

    for ced, (label, saldo_ced) in concilia_map.items():
        filas_ced = df[(df[col_iden] == ced) & (df["concilia_con_id"] == "")]
        mov_exact = filas_ced[abs(filas_ced[col_valor] - saldo_ced) < 0.01]
        if not mov_exact.empty:
            df.loc[mov_exact.index, "concilia_con_id"] = label
        else:
            df.loc[filas_ced.index, "concilia_con_id"] = label

    df.loc[df["concilia_con_id"] == "", "concilia_con_id"] = "REVISAR"

    return df


def _mostrar_resultado_puentes(df, codigo_cuenta):
    col_valor = "valor"

    total_movimientos = len(df)
    suma_total = round(pd.to_numeric(df[col_valor], errors="coerce").fillna(0).sum(), 2) \
                 if col_valor in df.columns else 0

    n_sin_nov  = (df["concilia_con_id"] == "SIN NOVEDAD").sum()
    n_revisar  = (df["concilia_con_id"] == "REVISAR").sum()
    n_concilia = total_movimientos - n_sin_nov - n_revisar

    val_sin_nov  = round(df[df["concilia_con_id"] == "SIN NOVEDAD"][col_valor].sum(), 2) \
                   if col_valor in df.columns else 0
    val_concilia = round(df[df["concilia_con_id"].str.startswith("Concilia", na=False)][col_valor].sum(), 2) \
                   if col_valor in df.columns else 0
    val_revisar  = round(df[df["concilia_con_id"] == "REVISAR"][col_valor].sum(), 2) \
                   if col_valor in df.columns else 0

    st.markdown("---")
    st.markdown("#### 📊 Resumen de Conciliación")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total movimientos",           f"{total_movimientos:,}")
    c2.metric("✅ Sin novedad",              f"{n_sin_nov:,}",   f"${val_sin_nov:,.2f}")
    c3.metric("🔗 Concilia con otra cédula", f"{n_concilia:,}",  f"${val_concilia:,.2f}")
    c4.metric("⚠️ Por revisar",              f"{n_revisar:,}",   f"${val_revisar:,.2f} de ${suma_total:,.2f}")

    st.markdown("#### 📋 Tabla de conciliación")
    st.dataframe(df, use_container_width=True, height=400)

    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="CONCILIACION")
    buf.seek(0)
    st.session_state["puentes_resultado_bytes"] = buf.getvalue()

    from datetime import date, timedelta
    ayer = date.today() - timedelta(days=1)
    fecha_str = ayer.strftime("%d_%m_%Y")
    nombre_archivo_puentes = f"CONCILIACION_{codigo_cuenta}_{fecha_str}.xlsx"

    cfg = st.session_state.get("config", {})
    ruta_auto_puentes = cfg.get("ruta_conc_puentes", "")
    if ruta_auto_puentes:
        try:
            import os
            os.makedirs(ruta_auto_puentes, exist_ok=True)
            ruta_completa_puentes = os.path.join(ruta_auto_puentes, nombre_archivo_puentes)
            with open(ruta_completa_puentes, "wb") as f:
                f.write(st.session_state["puentes_resultado_bytes"])
            st.success(f"💾 Guardado automáticamente en: {ruta_completa_puentes}")
        except Exception as e:
            st.warning(f"⚠️ No se pudo guardar automáticamente: {str(e)}")

    st.download_button(
        label=f"⬇️  Descargar resultado conciliación cuenta {codigo_cuenta} ({len(df):,} registros)",
        data=st.session_state["puentes_resultado_bytes"],
        file_name=nombre_archivo_puentes,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_resultado_puentes"
    )


# ══════════════════════════════════════════════════════════════════════════
# SUBMÓDULO 2: QR & CREDIBANCO
# ══════════════════════════════════════════════════════════════════════════

def render_qr_credibanco():
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1e3a5f,#1e40af);border-radius:10px;
                padding:1rem 1.5rem;margin-bottom:1rem;">
        <h3 style="color:#fff;margin:0">📱 Conciliación QR & Credibanco</h3>
        <p style="color:#93c5fd;margin:0.2rem 0 0 0;font-size:0.85rem">
            Libro Auxiliar + Libro de Banco</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📂 1. Cargar Archivos", "🔍 2. Conciliar"])

    with tab1:
        df_aux = render_cargue_auxiliares("qr")
        if df_aux is not None:
            st.session_state["qr_df_auxiliar"] = df_aux

        st.markdown("---")
        st.markdown("#### 📖 Libro de Banco")
        if st.session_state.get("archivo_libro"):
            st.success("✅ Libro de Banco ya cargado desde Aplicación y Compensación.")
        else:
            libro = st.file_uploader(
                "Sube el Libro de Banco",
                type=["xlsx","xlsm"],
                key="up_libro_qr",
                label_visibility="collapsed"
            )
            if libro:
                st.session_state["qr_libro_banco"] = libro
                st.success(f"✅ {libro.name}")

    with tab2:
        _render_conciliar_puentes()


# ══════════════════════════════════════════════════════════════════════════
# SUBMÓDULO 3: CUENTAS PUENTES / TRANSITORIAS
# ══════════════════════════════════════════════════════════════════════════

def render_puentes():
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1e3a5f,#1e40af);border-radius:10px;
                padding:1rem 1.5rem;margin-bottom:1rem;">
        <h3 style="color:#fff;margin:0">🔄 Cuentas Puentes / Transitorias</h3>
        <p style="color:#93c5fd;margin:0.2rem 0 0 0;font-size:0.85rem">
            Solo Libro Auxiliar</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📂 1. Cargar Archivos", "🔍 2. Conciliar"])

    with tab1:
        _render_carga_auxiliares_puentes()

    with tab2:
        _render_conciliar_puentes()


def _render_carga_auxiliares_puentes():
    st.markdown("#### 📂 Auxiliares")
    st.caption("Sube hasta 15 auxiliares. La app filtrará solo la cuenta que necesitas.")

    col_up, col_lim = st.columns([3, 1])
    with col_lim:
        if st.button("🔄 Limpiar archivos", key="limpiar_aux_puentes", use_container_width=True):
            st.session_state["puentes_auxiliares"] = []
            st.rerun()

    archivos = st.file_uploader(
        "Selecciona los auxiliares (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="up_puentes_auxiliares",
        label_visibility="collapsed"
    )

    if archivos:
        if len(archivos) > 15:
            st.warning("⚠️ Máximo 15 auxiliares.")
            archivos = archivos[:15]
        st.session_state["puentes_auxiliares"] = archivos

    cargados = st.session_state.get("puentes_auxiliares", [])
    if cargados:
        st.success(f"✅ {len(cargados)} auxiliar(es) listos.")
        for a in cargados:
            st.caption(f"   📄 {a.name}")
        st.info("👉 Ve a la pestaña **2. Conciliar**, ingresa el código de cuenta, selecciona el mes y filtra.")