"""
Módulo para conectar con SharePoint y descargar el Libro de Banco.
Usa autenticación OAuth con cuenta Microsoft (Device Flow).
El token se guarda localmente para no pedir credenciales cada vez.
"""
import os
import json
import io
import streamlit as st

# ── Configuración ──────────────────────────────────────────────────────────
SHAREPOINT_URL   = "https://supresas.sharepoint.com"
SITE_PATH        = "/sites/Departamentoadministrativoyfinanciero"
FILE_ID          = "019B1310-E23D-408C-B9CB-BDA242B72B31"
USER_EMAIL       = "ccamano@supre.com.co"
TOKEN_PATH       = os.path.join(os.path.dirname(__file__), ".sp_token.json")

# Client ID de la app Microsoft pública (Office365-REST-Python-Client default)
CLIENT_ID        = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
AUTHORITY        = "https://login.microsoftonline.com/common"
SCOPES           = ["https://graph.microsoft.com/Files.Read.All",
                    "https://graph.microsoft.com/Sites.Read.All"]


def _cargar_token():
    """Carga el token guardado si existe y no ha expirado."""
    if not os.path.exists(TOKEN_PATH):
        return None
    try:
        with open(TOKEN_PATH, "r") as f:
            return json.load(f)
    except Exception:
        return None


def _guardar_token(token_data):
    """Guarda el token en disco."""
    try:
        with open(TOKEN_PATH, "w") as f:
            json.dump(token_data, f)
    except Exception:
        pass


def _borrar_token():
    """Elimina el token guardado."""
    if os.path.exists(TOKEN_PATH):
        os.remove(TOKEN_PATH)


def obtener_token_device_flow():
    """
    Inicia el flujo Device Code para autenticar con Microsoft.
    Retorna el token o None si falla.
    """
    import msal

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )

    # Intentar con token cacheado primero
    cuentas = app.get_accounts(username=USER_EMAIL)
    if cuentas:
        resultado = app.acquire_token_silent(SCOPES, account=cuentas[0])
        if resultado and "access_token" in resultado:
            _guardar_token(resultado)
            return resultado

    # Iniciar Device Flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        return None

    return flow  # Retorna el flow para mostrar instrucciones al usuario


def completar_autenticacion(flow):
    """
    Completa la autenticación después de que el usuario ingresó el código.
    """
    import msal

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )

    resultado = app.acquire_token_by_device_flow(flow)
    if "access_token" in resultado:
        _guardar_token(resultado)
        return resultado
    return None


def descargar_libro_banco():
    """
    Descarga el Libro de Banco desde SharePoint.
    Retorna un BytesIO con el contenido del archivo o None si falla.
    """
    import requests

    token_data = _cargar_token()
    if not token_data or "access_token" not in token_data:
        return None, "Sin token de autenticación"

    access_token = token_data["access_token"]

    # URL de descarga directa via Graph API
    graph_url = (
        f"https://graph.microsoft.com/v1.0/sites/supresas.sharepoint.com:"
        f"{SITE_PATH}:/drive/items/{FILE_ID}/content"
    )

    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(graph_url, headers=headers, timeout=60)

    if response.status_code == 200:
        return io.BytesIO(response.content), None
    elif response.status_code == 401:
        _borrar_token()
        return None, "Token expirado — reconecta SharePoint"
    else:
        return None, f"Error {response.status_code}: {response.text[:200]}"


def tiene_token():
    """Verifica si hay un token guardado válido."""
    token = _cargar_token()
    return token is not None and "access_token" in token


def render_sharepoint_widget():
    """
    Renderiza el widget de conexión SharePoint en la UI de Streamlit.
    Retorna el archivo descargado (BytesIO) o None.
    """
    st.markdown("#### 🔗 Libro de Banco desde SharePoint")

    if tiene_token():
        col1, col2 = st.columns([3, 1])
        with col1:
            st.success("✅ Conectado a SharePoint como ccamano@supre.com.co")
        with col2:
            if st.button("🔓 Desconectar", key="btn_sp_desconectar", use_container_width=True):
                _borrar_token()
                st.rerun()

        if st.button("⬇️  Descargar Libro de Banco", type="primary",
                     use_container_width=True, key="btn_sp_descargar"):
            with st.spinner("Descargando desde SharePoint..."):
                archivo, error = descargar_libro_banco()
                if archivo:
                    archivo.name = "LIBRO BANCOS 2025.xlsx"
                    st.session_state.archivo_libro = archivo
                    st.success("✅ Libro de Banco descargado automáticamente.")
                    st.rerun()
                else:
                    st.error(f"❌ Error al descargar: {error}")
    else:
        st.info("📌 Conecta tu cuenta Microsoft para descargar el Libro de Banco automáticamente.")

        if "sp_flow" not in st.session_state:
            if st.button("🔑 Conectar con Microsoft", type="primary",
                         use_container_width=True, key="btn_sp_conectar"):
                with st.spinner("Iniciando autenticación..."):
                    flow = obtener_token_device_flow()
                    if flow and "user_code" in flow:
                        st.session_state["sp_flow"] = flow
                        st.rerun()
                    else:
                        st.error("❌ No se pudo iniciar la autenticación.")
        else:
            flow = st.session_state["sp_flow"]
            st.markdown(f"""
            **Pasos para conectar:**
            1. Abre 👉 [microsoft.com/devicelogin](https://microsoft.com/devicelogin)
            2. Ingresa el código: **`{flow['user_code']}`**
            3. Inicia sesión con `ccamano@supre.com.co`
            4. Haz clic en el botón de abajo cuando termines
            """)

            if st.button("✅ Ya ingresé el código", type="primary",
                         use_container_width=True, key="btn_sp_completar"):
                with st.spinner("Verificando autenticación..."):
                    resultado = completar_autenticacion(flow)
                    if resultado:
                        del st.session_state["sp_flow"]
                        st.success("✅ ¡Conectado exitosamente!")
                        st.rerun()
                    else:
                        st.error("❌ No se completó la autenticación. Intenta de nuevo.")
                        del st.session_state["sp_flow"]
                        st.rerun()

    return st.session_state.get("archivo_libro")