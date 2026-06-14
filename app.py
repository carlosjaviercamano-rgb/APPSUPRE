import streamlit as st
import json
import os

# ─── Configuración de página ───────────────────────────────────────────────
st.set_page_config(
    page_title="Herramienta Financiera Supre",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Estilos ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fuente principal */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #1a1f2e;
        border-right: 1px solid #2d3548;
    }
    [data-testid="stSidebar"] * { color: #e2e8f0 !important; }

    /* Logo / título sidebar */
    .sidebar-logo {
        padding: 1.5rem 1rem 1rem 1rem;
        border-bottom: 1px solid #2d3548;
        margin-bottom: 1rem;
    }
    .sidebar-logo h2 {
        color: #ffffff !important;
        font-size: 1.1rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.02em;
    }
    .sidebar-logo p {
        color: #64748b !important;
        font-size: 0.75rem;
        margin: 0.2rem 0 0 0;
    }

    /* Botones del menú */
    .menu-btn {
        display: flex;
        align-items: center;
        gap: 0.6rem;
        padding: 0.65rem 1rem;
        margin: 0.15rem 0;
        border-radius: 8px;
        cursor: pointer;
        transition: background 0.15s;
        font-size: 0.88rem;
        font-weight: 500;
        color: #94a3b8 !important;
        text-decoration: none;
    }
    .menu-btn:hover { background: #2d3548; color: #e2e8f0 !important; }
    .menu-btn.active { background: #1e40af; color: #ffffff !important; }
    .menu-btn .icon { font-size: 1rem; width: 20px; text-align: center; }
    .menu-badge {
        margin-left: auto;
        font-size: 0.65rem;
        background: #374151;
        color: #9ca3af !important;
        padding: 0.1rem 0.4rem;
        border-radius: 10px;
    }
    .menu-badge.ready { background: #065f46; color: #6ee7b7 !important; }

    /* Encabezado de módulo */
    .module-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #1e40af 100%);
        border-radius: 12px;
        padding: 1.5rem 2rem;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .module-header h1 {
        color: #ffffff !important;
        font-size: 1.4rem;
        font-weight: 700;
        margin: 0;
    }
    .module-header p {
        color: #93c5fd !important;
        font-size: 0.85rem;
        margin: 0.2rem 0 0 0;
    }
    .module-icon { font-size: 2rem; }

    /* Cards de configuración */
    .config-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 1rem;
    }
    .config-card h4 {
        color: #1e293b;
        font-size: 0.9rem;
        font-weight: 600;
        margin: 0 0 0.8rem 0;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    /* Botón guardar */
    .stButton > button {
        background: #1e40af;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        font-size: 0.9rem;
        transition: background 0.15s;
    }
    .stButton > button:hover { background: #1d4ed8; }

    /* Próximamente */
    .coming-soon {
        text-align: center;
        padding: 4rem 2rem;
        color: #94a3b8;
    }
    .coming-soon .cs-icon { font-size: 3rem; margin-bottom: 1rem; }
    .coming-soon h3 { color: #64748b; font-size: 1.1rem; font-weight: 600; }
    .coming-soon p { font-size: 0.88rem; }

    /* Ocultar elementos de Streamlit */
    #MainMenu, footer, header { visibility: hidden; }
    [data-testid="stDecoration"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ─── Ruta del archivo de configuración ────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

def cargar_config():
    default = {
        "libro_nombre": "LIBRO BANCOS 2025",
        "sharepoint_url": "",
        "ruta_movicap": r"C:\Users\ASUS\Desktop\BANCOS\1.MOVICAP\\",
        "ruta_suprecartera": r"C:\Users\ASUS\Desktop\BANCOS\2.SUPRECARTERA\\",
        "ruta_suprecredito": r"C:\Users\ASUS\Desktop\BANCOS\3.SUPRECREDITO\\",
        "ruta_tucredito": r"C:\Users\ASUS\Desktop\BANCOS\4.TUCREDITO\\",
        "ruta_compensaciones": r"C:\Users\ASUS\Desktop\BANCOS\COMPENSACION 2026\\"
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
                default.update(saved)
        except Exception:
            pass
    return default

def guardar_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ─── Estado de sesión ──────────────────────────────────────────────────────
if "modulo" not in st.session_state:
    st.session_state.modulo = "pagos"
if "config" not in st.session_state:
    st.session_state.config = cargar_config()

# ─── Sidebar ───────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <h2>🏦 Supre Financiero</h2>
        <p>Herramienta de gestión bancaria</p>
    </div>
    """, unsafe_allow_html=True)

    def menu_item(key, icon, label, badge=None, badge_class=""):
        active = "active" if st.session_state.modulo == key else ""
        badge_html = f'<span class="menu-badge {badge_class}">{badge}</span>' if badge else ""
        if st.button(f"{icon}  {label}", key=f"btn_{key}", use_container_width=True):
            st.session_state.modulo = key
            st.rerun()

    st.markdown("**MÓDULOS**", help="Selecciona el módulo de trabajo")

    for key, icon, label, badge, bcls in [
        ("pagos",         "💳", "Aplicación y Compensación",  "Activo", "ready"),
        ("cargue",        "📤", "Cargue Banco",                "Próximo", ""),
        ("conciliacion",  "🔍", "Conciliación Bancaria",       "Próximo", ""),
    ]:
        is_active = st.session_state.modulo == key
        btn_style = "background:#1e40af;color:white;" if is_active else ""
        if st.button(
            f"{icon}  {label}",
            key=f"nav_{key}",
            use_container_width=True,
            type="primary" if is_active else "secondary"
        ):
            st.session_state.modulo = key
            st.rerun()

    st.markdown("---")
    st.markdown("**SISTEMA**")
    if st.button("⚙️  Configuración", key="nav_config", use_container_width=True,
                 type="primary" if st.session_state.modulo == "config" else "secondary"):
        st.session_state.modulo = "config"
        st.rerun()

# ─── MÓDULO: CONFIGURACIÓN ─────────────────────────────────────────────────
def modulo_configuracion():
    st.markdown("""
    <div class="module-header">
        <div class="module-icon">⚙️</div>
        <div>
            <h1>Configuración</h1>
            <p>Rutas de archivos y conexiones del sistema</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    cfg = st.session_state.config

    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="config-card"><h4>📖 Libro SharePoint</h4>', unsafe_allow_html=True)
        cfg["libro_nombre"] = st.text_input(
            "Nombre del libro",
            value=cfg["libro_nombre"],
            help="Nombre exacto del archivo en SharePoint"
        )
        cfg["sharepoint_url"] = st.text_input(
            "URL SharePoint / OneDrive",
            value=cfg["sharepoint_url"],
            help="Ruta completa del libro de bancos"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="config-card"><h4>📁 Rutas de Planos</h4>', unsafe_allow_html=True)
        cfg["ruta_movicap"] = st.text_input("Movicap", value=cfg["ruta_movicap"])
        cfg["ruta_suprecartera"] = st.text_input("Suprecartera", value=cfg["ruta_suprecartera"])
        cfg["ruta_suprecredito"] = st.text_input("Suprecredito", value=cfg["ruta_suprecredito"])
        cfg["ruta_tucredito"] = st.text_input("TuCredito", value=cfg["ruta_tucredito"])
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="config-card"><h4>📁 Ruta Compensaciones</h4>', unsafe_allow_html=True)
        cfg["ruta_compensaciones"] = st.text_input(
            "Carpeta de compensaciones",
            value=cfg["ruta_compensaciones"]
        )
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="config-card"><h4>ℹ️ Estado del sistema</h4>', unsafe_allow_html=True)
        libro_ok = bool(cfg.get("libro_nombre"))
        url_ok   = bool(cfg.get("sharepoint_url"))
        st.markdown(f"{'✅' if libro_ok else '⚠️'} Nombre del libro: {'configurado' if libro_ok else 'pendiente'}")
        st.markdown(f"{'✅' if url_ok else '⚠️'} URL SharePoint: {'configurada' if url_ok else 'pendiente'}")
        rutas = ["ruta_movicap","ruta_suprecartera","ruta_suprecredito","ruta_tucredito","ruta_compensaciones"]
        rutas_ok = all(cfg.get(r) for r in rutas)
        st.markdown(f"{'✅' if rutas_ok else '⚠️'} Rutas de salida: {'todas configuradas' if rutas_ok else 'revisar rutas'}")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾  Guardar configuración", type="primary"):
        guardar_config(cfg)
        st.session_state.config = cfg
        st.success("✅ Configuración guardada correctamente.")

# ─── MÓDULO: PRÓXIMAMENTE ──────────────────────────────────────────────────
def modulo_proximamente(nombre, icono):
    st.markdown(f"""
    <div class="module-header">
        <div class="module-icon">{icono}</div>
        <div>
            <h1>{nombre}</h1>
            <p>Módulo en desarrollo</p>
        </div>
    </div>
    <div class="coming-soon">
        <div class="cs-icon">🔧</div>
        <h3>Próximamente disponible</h3>
        <p>Este módulo se habilitará una vez que<br>
        el módulo de Aplicación y Compensación esté completo.</p>
    </div>
    """, unsafe_allow_html=True)

# ─── MÓDULO: APLICACIÓN Y COMPENSACIÓN ────────────────────────────────────
def modulo_pagos():
    # Se importa desde su propio archivo
    from modulo_pagos import render
    render()

# ─── ROUTER ────────────────────────────────────────────────────────────────
modulo = st.session_state.modulo

if modulo == "config":
    modulo_configuracion()
elif modulo == "pagos":
    modulo_pagos()
elif modulo == "cargue":
    modulo_proximamente("Cargue Banco", "📤")
elif modulo == "conciliacion":
    modulo_proximamente("Conciliación Bancaria", "🔍")
