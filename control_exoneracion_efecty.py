import streamlit as st

# ══════════════════════════════════════════════════════════════════════════
# CÓDIGOS CAJERO POR GRUPO Y LÍMITES DE EXONERACIÓN
# ══════════════════════════════════════════════════════════════════════════

CODIGOS_EFECTY_EXITO = {
    27943, 27950, 27952, 27953, 27954, 27955, 27956, 27957, 27958, 27962, 27963, 27964, 27965, 27966, 27967, 27969, 27971, 27972, 27975, 27978, 27980, 27981, 27982, 27986, 27987, 27988, 27994, 27998, 27999, 28000, 28004, 28006, 28007, 28011, 28014, 28016, 28018, 28019, 28020, 28022, 28024, 28025, 28027, 28028, 28032, 28035, 28036, 28038, 28039, 28041, 28042, 28043, 28046, 28048, 28052, 28057, 28058, 28059, 28060, 28062, 28064, 28066, 28067, 28070, 28072, 28073, 28074, 28075, 28077, 28079, 28080, 28081, 28084, 28085, 28088, 28089, 28090, 28091, 28092, 28093, 28094, 28096, 28097, 28131, 28133, 28136, 28137, 28138, 28139, 28142, 28144, 28145, 28147, 28150, 28152, 28157, 28158, 28159, 28160, 28161, 28162, 28164, 28165, 28166, 28167, 28168, 28170, 28172, 28178, 28179, 28181, 28184, 28185, 28186, 28187, 28189, 28191, 28192, 28195, 28196, 28197, 28199, 28231, 28232, 28233, 28235, 28236, 28237, 28238, 28239, 28240, 28241, 28242, 28243, 28245, 28246, 28247, 28248, 28251, 28252, 28253, 28254, 28255, 28256, 28257, 28259, 28260, 28262, 28263, 28265, 28400, 28402, 28403, 28405, 28407, 28408, 28409, 28411, 28412, 28413, 28414, 28416, 28417, 28418, 28419, 28420, 28424, 28425, 28426, 28427, 28429, 28430, 28431, 28433, 28434, 28435, 28436, 28437, 28438, 28439, 28440, 28441, 28442, 28443, 28444, 28445, 28446, 28447, 28448, 28449, 28450, 28451, 28453, 28454, 28455, 28456, 28457, 28458, 28459, 28460, 28461, 28462, 28464, 28466, 28468, 28469, 28470, 28472, 28474, 28477, 28700, 28701, 28702, 28703, 28704, 28705, 28706, 28707, 28709, 28711, 28712, 28713, 28714, 28715, 28716, 28717, 28718, 28719, 28720, 28721, 28723, 28727, 28728, 28729, 28731, 28733, 28734, 28735, 28736, 28741, 28742, 28743, 28744, 28745, 28746, 28747, 28748, 28749, 28751, 28752, 28755, 28756, 28758, 28760, 28762, 28773, 28775, 28776, 28777, 28778, 28780, 28781, 28782, 28783, 28784, 28785, 28787, 28789, 28797, 28798, 28800, 28801, 28802, 28803, 28804, 28805, 28806, 28807, 28808, 28809, 28810, 28811, 28812, 28813, 28814, 28815, 28816, 28817, 28818, 28819, 28820, 28821, 28822, 28824, 28825, 28827, 28828, 28830, 28831, 28832, 28833, 28834, 28835, 28837, 28844, 28846, 28847, 28848, 28849, 28850, 28851, 28852, 28853, 28854, 28855, 28856, 28858, 28859, 28860, 28863, 28864, 28871, 28875, 28876, 28881, 28884, 28890, 28892, 28894, 28900, 28901, 28902, 28903, 28904, 28905, 28906, 28907, 28908, 28909, 28910, 28911, 28912, 28913, 28914, 28916, 28917, 28918, 28919, 28921, 28922, 28923, 28928, 28929, 28930, 28932, 28934, 28935, 28936, 28937, 28939, 28940, 28942, 28945, 28946, 28947, 28948, 28949, 28951, 28952, 28954, 28955, 28956, 28957, 28958, 28959, 28960, 28961, 28962, 28963, 28964, 28965, 28966, 28968, 28969, 28970, 28971, 28972, 28973, 28974, 28975, 28976, 28977, 28978, 28979, 28980, 28983, 28984, 28985, 28986, 28989, 28990, 28991, 28993, 28994, 28995, 28996, 28997, 28999, 29000, 29001, 29002, 29003, 29004, 29005, 29006, 29007, 29008, 29009, 29010, 29011, 29012, 29013, 29014, 29015, 29017, 29018, 29019, 29020, 29021, 29022, 29023, 29027, 29028, 29031, 29035, 29038, 29039, 29040, 29041, 29042, 29043, 29044, 29045, 29047, 29049, 29050, 29051, 29052, 29053, 29054, 29055, 29057, 29058, 29060, 29062, 29064, 29066, 29067, 29068, 29072, 29080, 29081, 29084, 29085, 29097, 29098, 29404, 29406, 29419, 29422, 29424, 29427, 29428, 29429, 29430, 29431, 29432, 29433, 29434
}
LIMITE_EFECTY_EXITO = 50000

CODIGOS_SUPERGIROS_OTROS = {
    27905, 27906, 27912, 27915, 27920, 27926, 27930, 27936, 27938, 27939, 27940, 27944, 27945, 27946, 27947, 27948, 28500, 28501, 28504, 28505, 28506, 29730, 29732, 29737, 29742, 29748, 29752, 29759, 29771, 29786, 29805, 29814, 29816, 29836, 29842, 29857, 29894, 29912, 29914, 29915, 29936, 29941, 29964, 29997, 29999
}
LIMITE_SUPERGIROS_OTROS = 20000

GRUPOS = {
    "EFECTY Y EXITO":       {"codigos": CODIGOS_EFECTY_EXITO,     "limite": LIMITE_EFECTY_EXITO},
    "SUPERGIROS Y OTROS":   {"codigos": CODIGOS_SUPERGIROS_OTROS, "limite": LIMITE_SUPERGIROS_OTROS},
}

UMBRAL_ALERTA = 0.75  # 75%

NOMBRE_HOJA = "Control Exoneración Efecty"
NOMBRE_PESTANA = "Conteo"


def clasificar_codigo(codigo_cajero):
    """Devuelve a qué grupo pertenece un código CAJERO, o None si no está
    registrado en ninguno de los dos grupos."""
    try:
        codigo = int(str(codigo_cajero).strip())
    except (ValueError, TypeError):
        return None
    for nombre_grupo, info in GRUPOS.items():
        if codigo in info["codigos"]:
            return nombre_grupo
    return None


def contar_transacciones_por_grupo(codigos_cajero):
    """
    Recibe una lista de códigos CAJERO (uno por transacción, pueden repetirse)
    y devuelve un diccionario {grupo: cantidad} más la lista de códigos que
    no se pudieron clasificar en ningún grupo.
    """
    conteo = {nombre: 0 for nombre in GRUPOS.keys()}
    no_clasificados = []
    for codigo in codigos_cajero:
        grupo = clasificar_codigo(codigo)
        if grupo:
            conteo[grupo] += 1
        else:
            no_clasificados.append(codigo)
    return conteo, no_clasificados


# ══════════════════════════════════════════════════════════════════════════
# CONEXIÓN A GOOGLE SHEETS
# ══════════════════════════════════════════════════════════════════════════

@st.cache_resource(ttl=3600)
def _obtener_hoja():
    """Abre la pestaña de conteo en la hoja de Google Sheets, usando las
    credenciales de la cuenta de servicio guardadas en secrets.toml."""
    import gspread
    from google.oauth2.service_account import Credentials

    creds_dict = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    cliente = gspread.authorize(creds)

    hoja = cliente.open(NOMBRE_HOJA)
    try:
        pestana = hoja.worksheet(NOMBRE_PESTANA)
    except gspread.WorksheetNotFound:
        pestana = hoja.sheet1

    return pestana


def _asegurar_encabezados(pestana):
    """Si la pestaña está vacía, crea los encabezados y una fila por grupo."""
    valores = pestana.get_all_values()
    if not valores:
        pestana.update("A1", [["PUNTO/CAJERO", "CONTEO", "LIMITE", "FECHA_ULTIMA_ACTUALIZACION"]])
        filas = []
        for nombre_grupo, info in GRUPOS.items():
            filas.append([nombre_grupo, 0, info["limite"], ""])
        pestana.update(f"A2:D{1 + len(filas)}", filas)


def leer_estado():
    """
    Devuelve un diccionario {grupo: {conteo, limite, fecha}} con el estado
    actual guardado en Google Sheets.
    """
    pestana = _obtener_hoja()
    _asegurar_encabezados(pestana)
    registros = pestana.get_all_records()

    estado = {}
    for nombre_grupo, info in GRUPOS.items():
        estado[nombre_grupo] = {"conteo": 0, "limite": info["limite"], "fecha": ""}

    for fila in registros:
        nombre_grupo = str(fila.get("PUNTO/CAJERO", "")).strip()
        if nombre_grupo in estado:
            try:
                estado[nombre_grupo]["conteo"] = int(fila.get("CONTEO", 0) or 0)
            except (ValueError, TypeError):
                estado[nombre_grupo]["conteo"] = 0
            estado[nombre_grupo]["fecha"] = fila.get("FECHA_ULTIMA_ACTUALIZACION", "")

    return estado


def registrar_transacciones(conteo_nuevo_por_grupo):
    """
    Suma las transacciones nuevas (por grupo) al conteo ya guardado en
    Google Sheets, y guarda el resultado.
    """
    from datetime import datetime

    pestana = _obtener_hoja()
    _asegurar_encabezados(pestana)
    registros = pestana.get_all_records()

    conteo_actual = {nombre: 0 for nombre in GRUPOS.keys()}
    for fila in registros:
        nombre_grupo = str(fila.get("PUNTO/CAJERO", "")).strip()
        if nombre_grupo in conteo_actual:
            try:
                conteo_actual[nombre_grupo] = int(fila.get("CONTEO", 0) or 0)
            except (ValueError, TypeError):
                conteo_actual[nombre_grupo] = 0

    filas = []
    for nombre_grupo, info in GRUPOS.items():
        nuevo_total = conteo_actual[nombre_grupo] + conteo_nuevo_por_grupo.get(nombre_grupo, 0)
        filas.append([
            nombre_grupo,
            nuevo_total,
            info["limite"],
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        ])

    pestana.update(f"A2:D{1 + len(filas)}", filas)
    return {nombre: fila[1] for nombre, fila in zip(GRUPOS.keys(), filas)}


def reiniciar_conteo():
    """Pone el conteo de todos los grupos en cero (nueva exoneración)."""
    from datetime import datetime

    pestana = _obtener_hoja()
    _asegurar_encabezados(pestana)

    filas = []
    for nombre_grupo, info in GRUPOS.items():
        filas.append([nombre_grupo, 0, info["limite"], datetime.now().strftime("%d/%m/%Y %H:%M:%S")])

    pestana.update(f"A2:D{1 + len(filas)}", filas)


# ══════════════════════════════════════════════════════════════════════════
# INTERFAZ DE CONSULTA
# ══════════════════════════════════════════════════════════════════════════

def render_estado_exoneracion():
    st.markdown("#### 📊 Estado de Exoneración por Punto")

    try:
        estado = leer_estado()
    except Exception as e:
        st.error(f"❌ No se pudo conectar con Google Sheets: {str(e)}")
        return

    cols = st.columns(len(estado))
    for col, (nombre_grupo, datos) in zip(cols, estado.items()):
        conteo = datos["conteo"]
        limite = datos["limite"]
        porcentaje = (conteo / limite * 100) if limite else 0

        if porcentaje >= 100:
            color = "#e74c3c"
            icono = "🔴"
        elif porcentaje >= UMBRAL_ALERTA * 100:
            color = "#f39c12"
            icono = "🟡"
        else:
            color = "#2ecc71"
            icono = "🟢"

        with col:
            st.markdown(f"""
            <div style="background-color:#1a1f2e;border:1px solid {color};border-radius:10px;
                        padding:1rem;margin-bottom:0.5rem;">
                <div style="color:{color};font-size:0.85rem;font-weight:600;">
                    {icono} {nombre_grupo}
                </div>
                <div style="color:#ffffff;font-size:1.6rem;font-weight:700;margin-top:0.3rem;">
                    {conteo:,} / {limite:,}
                </div>
                <div style="color:{color};font-size:0.9rem;font-weight:600;margin-top:0.2rem;">
                    {porcentaje:.1f}%
                </div>
            </div>
            """, unsafe_allow_html=True)

            if datos["fecha"]:
                st.caption(f"Última actualización: {datos['fecha']}")

    if st.button("🔄  Reiniciar conteo (nueva exoneración)", key="btn_reiniciar_exoneracion"):
        st.session_state["confirmar_reinicio_exoneracion"] = True

    if st.session_state.get("confirmar_reinicio_exoneracion"):
        st.warning("⚠️ Esto pondrá el conteo de ambos grupos en 0. ¿Confirmas?")
        col_si, col_no = st.columns(2)
        with col_si:
            if st.button("✅ Sí, reiniciar", key="btn_confirmar_reinicio", use_container_width=True):
                try:
                    reiniciar_conteo()
                    st.session_state["confirmar_reinicio_exoneracion"] = False
                    st.success("✅ Conteo reiniciado correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ No se pudo reiniciar: {str(e)}")
        with col_no:
            if st.button("Cancelar", key="btn_cancelar_reinicio", use_container_width=True):
                st.session_state["confirmar_reinicio_exoneracion"] = False
                st.rerun()