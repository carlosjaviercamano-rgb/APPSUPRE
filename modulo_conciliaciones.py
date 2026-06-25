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

    # ── Submódulos ───────────────────────────────────────────────────────
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
                    padding:1.5rem;text-align:center;cursor:pointer;">
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
                    padding:1.5rem;text-align:center;cursor:pointer;">
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
                    padding:1.5rem;text-align:center;cursor:pointer;">
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


# ══════════════════════════════════════════════════════════════════════════
# SECCIÓN COMPARTIDA: LIBRO AUXILIAR
# ══════════════════════════════════════════════════════════════════════════

COLUMNAS_AUXILIAR = [
    "Source.Name", "id", "empresa", "centrocosto", "fecha", "tipodocumento",
    "nrodocumento", "codigocuenta", "nombrecuenta", "descripcion", "factura",
    "fechavencimiento", "identificacion", "tercero", "valor", "baseimpuesto",
    "porcentajeimpuesto", "saldoanterior", "debito", "credito", "saldoactual",
    "usuariocreador", "usuarioactualizador"
]


def render_cargue_auxiliares(key_prefix=""):
    """
    Sección reutilizable para cargar y unir auxiliares.
    Retorna el DataFrame del libro auxiliar unificado o None.
    """
    st.markdown("#### 📂 Libro Auxiliar")
    st.caption("Sube hasta 15 auxiliares individuales. La app los unirá automáticamente.")

    key_archivos = f"{key_prefix}_auxiliares"
    if key_archivos not in st.session_state:
        st.session_state[key_archivos] = []

    # Uploader múltiple
    archivos = st.file_uploader(
        "Selecciona los auxiliares (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"up_{key_prefix}_auxiliares",
        label_visibility="collapsed"
    )

    if archivos:
        # Limitar a 15
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
            with st.spinner("Uniendo auxiliares... esto puede tomar unos segundos."):
                try:
                    df_completo, df_conciliar = _unir_auxiliares(archivos_cargados)
                    # Guardar ambas versiones en sesión
                    st.session_state[key_df]                      = df_conciliar  # para conciliar (ligero)
                    st.session_state[f"{key_prefix}_df_completo"] = df_completo   # para descarga (completo)
                    # Limpiar bytes anteriores para regenerar
                    key_bytes = f"{key_prefix}_auxiliar_bytes"
                    if key_bytes in st.session_state:
                        del st.session_state[key_bytes]
                    st.success(f"✅ Libro Auxiliar creado: {len(df_completo):,} registros de {len(archivos_cargados)} auxiliar(es).")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

        # Mostrar preview y descarga si ya está unido
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

            # Generar bytes del Excel completo solo una vez
            key_bytes    = f"{key_prefix}_auxiliar_bytes"
            key_completo = f"{key_prefix}_df_completo"
            df_completo  = st.session_state.get(key_completo, df_auxiliar)

            if key_bytes not in st.session_state:
                cols_decimales = ["valor", "baseimpuesto", "saldoanterior",
                                  "debito", "credito", "saldoactual"]
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

    # Si ya hay df en sesión aunque no haya archivos cargados, mostrar descarga
    df_auxiliar = st.session_state.get(key_df)
    if df_auxiliar is not None:
        key_bytes = f"{key_prefix}_auxiliar_bytes"
        if key_bytes in st.session_state:
            st.markdown("---")
            st.info("📋 Libro Auxiliar previamente generado disponible para descarga.")
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
    """
    Une múltiples auxiliares manteniendo estructura completa.
    Para descarga: todas las columnas.
    Para conciliación: solo columnas clave (optimización de memoria).
    """
    cols_conciliacion = {
        "id", "empresa", "codigocuenta", "identificacion", "valor"
    }

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

            # DataFrame completo para descarga
            frames_completos.append(df)

            # DataFrame reducido para conciliación (menos memoria)
            cols_disp = [c for c in df.columns if c in cols_conciliacion or c == "Source.Name"]
            frames_conciliar.append(df[cols_disp].copy())

        except Exception as e:
            raise ValueError(f"Error leyendo {archivo.name}: {str(e)}")

    if not frames_completos:
        raise ValueError("No se pudieron leer los auxiliares.")

    df_completo  = pd.concat(frames_completos,  ignore_index=True)
    df_conciliar = pd.concat(frames_conciliar,  ignore_index=True)

    return df_completo, df_conciliar


# ══════════════════════════════════════════════════════════════════════════
# SUBMÓDULO 1: CONCILIACIÓN BANCARIA
# ══════════════════════════════════════════════════════════════════════════
def render_bancaria():
    from conciliacion_bancaria import EMPRESAS_CUENTAS, MESES as MESES_CB
    from conciliacion_bancaria import detectar_y_parsear_extracto, conciliar

    st.markdown("""
    <div style="background:linear-gradient(135deg,#1e3a5f,#1e40af);border-radius:10px;
                padding:1rem 1.5rem;margin-bottom:1rem;">
        <h3 style="color:#fff;margin:0">\U0001f3e6 Conciliaci\u00f3n Bancaria</h3>
        <p style="color:#93c5fd;margin:0.2rem 0 0 0;font-size:0.85rem">
            Libro Auxiliar + Extracto bancario</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["\U0001f4c2 1. Cargar Archivos", "\U0001f50d 2. Conciliar"])

    with tab1:
        # Selectores
        st.markdown("#### \u2699\ufe0f Configuraci\u00f3n")
        col1, col2, col3 = st.columns(3)

        empresas = list(EMPRESAS_CUENTAS.keys())
        with col1:
            empresa_sel = st.selectbox("Empresa", empresas, key="banc_empresa")
        cuentas_emp   = EMPRESAS_CUENTAS.get(empresa_sel, [])
        nombres_ctas  = [c[0] for c in cuentas_emp]
        with col2:
            cuenta_sel = st.selectbox("Cuenta bancaria", nombres_ctas, key="banc_cuenta")
        codigo_cuenta = next((c[1] for c in cuentas_emp if c[0] == cuenta_sel), None)
        with col3:
            mes_sel = st.selectbox("Mes", MESES_CB, key="banc_mes")
        st.caption(f"C\u00f3digo cuenta: **{codigo_cuenta}**")

        st.markdown("---")

        # Auxiliares
        st.markdown("#### \U0001f4c2 Auxiliares")
        st.caption("Sube hasta 15 auxiliares. Se filtrar\u00e1n al conciliar.")

        _, col_lim_aux = st.columns([3,1])
        with col_lim_aux:
            if st.button("\U0001f504 Limpiar archivos", key="limpiar_aux_bancaria", use_container_width=True):
                st.session_state.pop("banc_auxiliares", None)
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
            st.success(f"\u2705 {len(cargados)} auxiliar(es) listos.")
            for a in cargados:
                st.caption(f"   \U0001f4c4 {a.name}")

        st.markdown("---")

        # Extracto
        st.markdown("#### \U0001f3e6 Extracto Bancario")
        st.caption("Formato detectado autom\u00e1ticamente: Bancolombia CSV, Davivienda XLSX, Occidente/Bogot\u00e1 XLS.")

        extracto = st.file_uploader(
            "Extracto bancario",
            type=["csv","xlsx","xls","XLS"],
            key="up_banc_extracto", label_visibility="collapsed"
        )
        if extracto:
            st.session_state["banc_extracto"] = extracto
            st.success(f"\u2705 {extracto.name}")
        elif st.session_state.get("banc_extracto"):
            st.success(f"\u2705 {st.session_state['banc_extracto'].name}")

        if cargados and st.session_state.get("banc_extracto"):
            st.info("\U0001f449 Ve a la pesta\u00f1a **2. Conciliar** para ejecutar.")

    with tab2:
        _render_conciliar_bancaria()


def _render_conciliar_bancaria():
    from conciliacion_bancaria import detectar_y_parsear_extracto, conciliar, MESES as MESES_CB

    archivos_aux = st.session_state.get("banc_auxiliares", [])
    extracto     = st.session_state.get("banc_extracto")
    empresa      = st.session_state.get("banc_empresa", "")
    cuenta_nom   = st.session_state.get("banc_cuenta", "")
    cod_cuenta   = None
    mes          = st.session_state.get("banc_mes", "")

    # Recalcular código cuenta
    from conciliacion_bancaria import EMPRESAS_CUENTAS
    cuentas_emp = EMPRESAS_CUENTAS.get(empresa, [])
    cod_cuenta  = next((c[1] for c in cuentas_emp if c[0] == cuenta_nom), None)

    if not archivos_aux:
        st.warning("\u26a0\ufe0f Primero carga los auxiliares en la pesta\u00f1a **1. Cargar Archivos**.")
        return
    if not extracto:
        st.warning("\u26a0\ufe0f Primero carga el extracto bancario en la pesta\u00f1a **1. Cargar Archivos**.")
        return

    st.markdown(f"### {empresa} | {cuenta_nom} | {mes}")

    col_btn, col_lim = st.columns([3,1])
    with col_btn:
        ejecutar = st.button("\U0001f50d Conciliar", type="primary",
                             use_container_width=True, key="btn_conciliar_bancaria")
    with col_lim:
        if st.button("\U0001f504 Limpiar", use_container_width=True, key="btn_limpiar_bancaria"):
            for k in ["banc_resultado","banc_resumen"]:
                st.session_state.pop(k, None)
            st.rerun()

    if ejecutar:
        with st.spinner("Procesando..."):
            try:
                # 1. Parsear extracto
                extracto.seek(0)
                df_banco = detectar_y_parsear_extracto(extracto)
                if df_banco.empty:
                    st.error("\u274c No se pudieron leer datos del extracto.")
                    return

                # 2. Leer y filtrar auxiliares por empresa+cuenta+mes
                mes_idx = MESES_CB.index(mes) + 1 if mes in MESES_CB else None
                frames = []
                for archivo in archivos_aux:
                    try:
                        archivo.seek(0)
                        df = pd.read_excel(archivo, sheet_name=0)
                        df.columns = [c.strip().lower() for c in df.columns]
                        # Limpiar columnas NaN
                        df = df.loc[:, df.columns.notna()]
                        df = df.loc[:, ~df.columns.astype(str).str.lower().isin(["nan","none",""])]

                        # Filtrar empresa (case-insensitive)
                        if "empresa" in df.columns:
                            df = df[df["empresa"].astype(str).str.strip().str.lower() == empresa.lower()]

                        # Filtrar cuenta — columna codigocuenta
                        col_c = next((c for c in df.columns if "codigocuenta" in c.lower() or "codigo_cuenta" in c.lower()), None)
                        if col_c:
                            df = df[pd.to_numeric(df[col_c], errors="coerce").fillna(0).astype(int).astype(str) == str(cod_cuenta).strip()]

                        # Filtrar mes (soporta DD/MM/YYYY y YYYY-MM-DD)
                        col_f = next((c for c in df.columns if c == "fecha"), None)
                        if col_f and mes_idx:
                            df[col_f] = pd.to_datetime(df[col_f], dayfirst=True, errors="coerce")
                            df = df[df[col_f].dt.month == mes_idx]

                        if not df.empty:
                            frames.append(df)
                        del df
                    except Exception:
                        continue

                if not frames:
                    st.error(f"\u274c No se encontraron datos para {empresa} / cuenta {cod_cuenta} / {mes}.")
                    return

                df_aux = pd.concat(frames, ignore_index=True)

                # Mapear columnas del auxiliar
                col_map = {}
                for col in df_aux.columns:
                    col_n = col.lower().strip()
                    if col_n == "fecha": col_map["fecha"] = col
                    if "descripci" in col_n: col_map["descripcion"] = col
                    if col_n in ["debito","débito"]: col_map["debito"] = col
                    if col_n in ["credito","crédito"]: col_map["credito"] = col
                    if col_n == "valor": col_map["valor"] = col
                st.caption(f"Columnas detectadas: {col_map}")

                df_aux_norm = pd.DataFrame()
                df_aux_norm["fecha"]       = pd.to_datetime(df_aux.get(col_map.get("fecha"), pd.Series(dtype=str)), errors="coerce")
                df_aux_norm["descripcion"] = df_aux.get(col_map.get("descripcion",""), pd.Series(dtype=str)).astype(str)
                df_aux_norm["debito"]      = pd.to_numeric(df_aux.get(col_map.get("debito",""), 0), errors="coerce").fillna(0).round(2)
                df_aux_norm["credito"]     = pd.to_numeric(df_aux.get(col_map.get("credito",""), 0), errors="coerce").fillna(0).round(2)
                df_aux_norm["valor"]       = pd.to_numeric(df_aux.get(col_map.get("valor",""), 0), errors="coerce").fillna(0).round(2)

                # 3. Conciliar
                df_result, resumen = conciliar(df_banco, df_aux_norm)
                st.session_state["banc_resultado"] = df_result
                st.session_state["banc_resumen"]   = resumen
                st.success("\u2705 Conciliaci\u00f3n completada.")
            except Exception as e:
                st.error(f"\u274c Error: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    # Mostrar resultado
    df_result = st.session_state.get("banc_resultado")
    resumen   = st.session_state.get("banc_resumen")

    if df_result is not None and resumen is not None:
        st.markdown("---")
        st.markdown("#### \U0001f4ca Resumen")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total banco d\u00e9bito",  f"${resumen['total_banco_debito']:,.2f}")
        c2.metric("Total banco cr\u00e9dito", f"${resumen['total_banco_credito']:,.2f}")
        c3.metric("\u2705 Conciliados",       f"{resumen['conciliados']:,}")
        c4.metric("\U0001f3e6 Solo banco",    f"{resumen['solo_banco']:,}")
        c5.metric("\U0001f4d2 Solo auxiliar", f"{resumen['solo_auxiliar']:,}")

        dif_deb = resumen["diferencia_debito"]
        dif_cre = resumen["diferencia_credito"]
        if abs(dif_deb) < 0.01 and abs(dif_cre) < 0.01:
            st.success("\u2705 DIFERENCIA = $0 \u2014 Cuenta conciliada correctamente.")
        else:
            st.warning(f"\u26a0\ufe0f Diferencia d\u00e9bito: ${dif_deb:,.2f} | Diferencia cr\u00e9dito: ${dif_cre:,.2f}")

        st.markdown("#### \U0001f4cb Detalle de conciliaci\u00f3n")
        st.dataframe(df_result, use_container_width=True, height=450)

        import io as _io
        buf = _io.BytesIO()
        df_result.to_excel(buf, index=False, sheet_name="CONCILIACION")
        buf.seek(0)
        st.download_button(
            label="\u2b07\ufe0f  Descargar resultado Excel",
            data=buf.getvalue(),
            file_name=f"CONCILIACION_{empresa}_{mes}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_bancaria"
        )


def _render_conciliar_puentes():
    archivos = st.session_state.get("puentes_auxiliares", [])
    if not archivos:
        st.warning("⚠️ Primero carga los auxiliares en la pestaña **1. Cargar Archivos**.")
        return

    st.markdown("### 🔍 Conciliación Cuentas Puentes / Transitorias")
    st.caption(f"📂 {len(archivos)} auxiliar(es) cargado(s). Ingresa la cuenta y filtra solo esos datos.")

    col_inp, col_btn, col_lim = st.columns([3, 1, 1])
    with col_inp:
        codigo_cuenta = st.text_input(
            "Código de cuenta a conciliar (codigocuenta):",
            value=st.session_state.get("puentes_cuenta", ""),
            placeholder="Ej: 141299011",
            key="input_cuenta_puentes"
        )
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        ejecutar = st.button("🔍 Filtrar y Conciliar", type="primary",
                             use_container_width=True, key="btn_conciliar_puentes")
    with col_lim:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Limpiar", use_container_width=True, key="limpiar_conciliar_puentes"):
            st.session_state["puentes_df_filtrado"] = None
            st.session_state["puentes_cuenta"]      = ""
            st.session_state["puentes_resultado_bytes"] = None
            st.rerun()

    if ejecutar and codigo_cuenta.strip():
        st.session_state["puentes_cuenta"] = codigo_cuenta.strip()
        with st.spinner("Leyendo auxiliares y filtrando por cuenta..."):
            try:
                df_filtrado = _leer_y_filtrar_por_cuenta(archivos, codigo_cuenta.strip())
                if df_filtrado is None or df_filtrado.empty:
                    st.warning(f"⚠️ No se encontraron registros para la cuenta **{codigo_cuenta}**.")
                else:
                    st.session_state["puentes_df_filtrado"] = df_filtrado
                    st.success(f"✅ {len(df_filtrado):,} registros encontrados para cuenta {codigo_cuenta}.")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
    elif ejecutar:
        st.warning("⚠️ Ingresa un código de cuenta.")

    # Mostrar resultado si ya está filtrado
    df_filtrado = st.session_state.get("puentes_df_filtrado")
    if df_filtrado is not None:
        cuenta_actual = st.session_state.get("puentes_cuenta", "")
        _ejecutar_conciliacion_puentes(df_filtrado, cuenta_actual)


def _leer_y_filtrar_por_cuenta(archivos, codigo_cuenta):
    """Lee cada auxiliar filtrando solo la cuenta indicada."""
    frames = []
    for archivo in archivos:
        try:
            # Leer el archivo completo (read_excel no soporta chunksize)
            # pero solo las columnas necesarias para ahorrar memoria
            df = pd.read_excel(archivo, sheet_name=0)
            df.columns = [c.strip().lower() for c in df.columns]

            # Columna G del auxiliar original = codigocuenta (sin Source.Name al inicio)
            col_cuenta = "codigocuenta"
            if col_cuenta not in df.columns:
                continue

            # Filtrar solo filas de la cuenta
            filtrado = df[
                df[col_cuenta].astype(str).str.strip() == str(codigo_cuenta).strip()
            ].copy()

            if filtrado.empty:
                continue

            # Agregar Source.Name con nombre de empresa
            empresa = str(filtrado["empresa"].iloc[0]).strip() if "empresa" in filtrado.columns else archivo.name
            filtrado.insert(0, "Source.Name", empresa)
            frames.append(filtrado)

            # Liberar memoria del df completo
            del df

        except Exception as e:
            raise ValueError(f"Error leyendo {archivo.name}: {str(e)}")

    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def _ejecutar_conciliacion_puentes(df_filtrado, codigo_cuenta):
    from itertools import combinations
    col_valor = "valor"
    col_iden  = "identificacion"
    col_id    = "id"

    df = df_filtrado.copy()
    df[col_valor] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0)
    df[col_iden]  = df[col_iden].astype(str).str.strip().replace({"nan":"","None":"","NaN":""})

    total_movimientos = len(df)
    suma_total        = round(df[col_valor].sum(), 2)

    st.markdown(f"**Cuenta:** `{codigo_cuenta}` | **Movimientos:** {total_movimientos:,} | **Suma total:** ${suma_total:,.2f}")

    if abs(suma_total) < 0.01:
        st.success("✅ CUENTA CONCILIADA — La suma de todos los movimientos es 0.")
        df["concilia_con_id"] = "CONCILIADA"
        _mostrar_resultado(df, codigo_cuenta, total_movimientos, 0, total_movimientos, 0)
        return

    st.warning(f"⚠️ Suma ≠ 0 (${suma_total:,.2f}). Aplicando lógica de conciliación...")

    df["concilia_con_id"] = ""

    # PASO 1A: Pares exactos dentro de la misma cédula → SIN NOVEDAD
    def _buscar_pares_internos(grupo):
        pares = set()
        pos = grupo[grupo[col_valor] > 0]
        neg = grupo[grupo[col_valor] < 0]
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

    # PASO 1B: Saldo neto de cédulas restantes
    df_resto = df[df["concilia_con_id"] == ""]
    saldos_resto = df_resto[df_resto[col_iden] != ""].groupby(col_iden)[col_valor].sum().round(2)

    # Cédulas cuyo saldo neto restante = 0 → SIN NOVEDAD
    ced_sin_nov2 = set(saldos_resto[abs(saldos_resto) < 0.01].index)
    mask_sn2 = (df[col_iden].isin(ced_sin_nov2)) & (df["concilia_con_id"] == "")
    df.loc[mask_sn2, "concilia_con_id"] = "SIN NOVEDAD"

    # Cédulas con saldo neto ≠ 0
    ced_con_sal = {ced: round(sal,2) for ced,sal in saldos_resto.items()
                   if abs(sal) >= 0.01 and ced not in ced_sin_nov2}

    # PASO 2: Buscar grupos de cédulas cuyo saldo neto sume 0
    saldos_pend  = dict(ced_con_sal)
    concilia_map = {}  # cedula → (label, indices a marcar)

    encontrado = True
    while encontrado and len(saldos_pend) >= 2:
        encontrado = False
        items = list(saldos_pend.items())
        for r in range(2, min(len(items)+1, 6)):
            if encontrado: break
            for combo in combinations(items, r):
                ceds = [c for c,_ in combo]
                vals = [v for _,v in combo]
                if abs(round(sum(vals), 2)) < 0.01:
                    for ced in ceds:
                        saldo_ced = saldos_pend[ced]
                        # Buscar cédula opuesta
                        opuestas = [c for c in ceds if c != ced and saldos_pend[c] * saldo_ced < 0]
                        if not opuestas:
                            opuestas = [c for c in ceds if c != ced]
                        ced_op   = opuestas[0]
                        saldo_op = saldos_pend[ced_op]
                        # Buscar movimiento de ced_op con valor = saldo_op
                        filas_op  = df[(df[col_iden] == ced_op) & (df["concilia_con_id"] == "")]
                        mov_exact = filas_op[abs(filas_op[col_valor] - saldo_op) < 0.01]
                        if not mov_exact.empty:
                            id_ref = int(mov_exact[col_id].iloc[0])
                        else:
                            mov_s = filas_op[filas_op[col_valor] * saldo_ced < 0]
                            id_ref = int(mov_s[col_id].iloc[0]) if not mov_s.empty else int(filas_op[col_id].iloc[0]) if not filas_op.empty else ced_op
                        concilia_map[ced] = (f"Concilia con ID {id_ref}", saldo_ced)
                    for ced in ceds:
                        del saldos_pend[ced]
                    encontrado = True
                    break

    # Aplicar label solo al movimiento cuyo valor = saldo neto de esa cédula
    for ced, (label, saldo_ced) in concilia_map.items():
        filas_ced = df[(df[col_iden] == ced) & (df["concilia_con_id"] == "")]
        # Buscar movimiento con valor = saldo neto
        mov_exact = filas_ced[abs(filas_ced[col_valor] - saldo_ced) < 0.01]
        if not mov_exact.empty:
            df.loc[mov_exact.index, "concilia_con_id"] = label
        else:
            # Si no hay exacto, marcar todos los pendientes
            df.loc[filas_ced.index, "concilia_con_id"] = label

    # PASO 3: Resto → REVISAR
    df.loc[df["concilia_con_id"] == "", "concilia_con_id"] = "REVISAR"

    n_sin_nov  = (df["concilia_con_id"] == "SIN NOVEDAD").sum()
    n_revisar  = (df["concilia_con_id"] == "REVISAR").sum()
    n_concilia = total_movimientos - n_sin_nov - n_revisar
    val_sin_nov  = round(df[df["concilia_con_id"] == "SIN NOVEDAD"][col_valor].sum(), 2)
    val_concilia = round(df[df["concilia_con_id"].str.startswith("Concilia")][col_valor].sum(), 2)
    val_revisar  = round(df[df["concilia_con_id"] == "REVISAR"][col_valor].sum(), 2)

    _mostrar_resultado(df, codigo_cuenta, total_movimientos, n_sin_nov, n_concilia, n_revisar,
                       val_sin_nov, val_concilia, val_revisar, suma_total)


def _mostrar_resultado(df, codigo_cuenta, total, n_sn, n_conc, n_rev,
                       val_sn=0, val_conc=0, val_rev=0, suma_total=0):
    st.markdown("---")
    st.markdown("#### 📊 Resumen de Conciliación")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total movimientos",            f"{total:,}")
    c2.metric("✅ Sin novedad",               f"{n_sn:,}",   f"${val_sn:,.2f}")
    c3.metric("🔗 Concilia con otra cédula",  f"{n_conc:,}", f"${val_conc:,.2f}")
    c4.metric("⚠️ Por revisar",               f"{n_rev:,}",  f"${val_rev:,.2f} de ${suma_total:,.2f}")

    st.markdown("#### 📋 Tabla de conciliación")
    st.dataframe(df, use_container_width=True, height=400)

    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="CONCILIACION")
    buf.seek(0)
    st.session_state["puentes_resultado_bytes"] = buf.getvalue()

    st.download_button(
        label=f"⬇️  Descargar resultado conciliación cuenta {codigo_cuenta} ({len(df):,} registros)",
        data=st.session_state["puentes_resultado_bytes"],
        file_name=f"CONCILIACION_{codigo_cuenta}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_resultado_puentes"
    )


# ══════════════════════════════════════════════════════════════════════════
# SUBMÓDULO 2: CONCILIACIÓN QR & CREDIBANCO
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
                type=["xlsx", "xlsm"],
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
        st.info("👉 Ve a la pestaña **2. Conciliar**, ingresa el código de cuenta y filtra.")
