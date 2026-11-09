"""
Tipo de documento (CC / NIT) de los clientes para los planos y compensaciones.

El tipo sale de la columna `tipo_documento` del libro CLIENTES_ACTIVOS.xlsx (hoja sheet1),
que queda así: company | iden | num_factura | tipo_documento.

Reglas:
  - Solo se manejan CC y NIT.
  - Si el cliente no tiene tipo (celda vacía, columna inexistente o cédula que no está en el
    libro) se usa CC.
  - Si un mismo cliente aparece con tipos contradictorios (CC en una fila y NIT en otra) se usa CC
    y se avisa.
Los NIT de entidades (bancos, Efecty, etc.) siguen fijos en cada generador: no pasan por aquí.
"""
import hashlib
import io

import pandas as pd
import streamlit as st

TIPO_POR_DEFECTO = "CC"
TIPOS_VALIDOS = ("CC", "NIT")
GITHUB_RAW = "https://raw.githubusercontent.com/carlosjaviercamano-rgb/APPSUPRE/main"


def _norm_cedula(v):
    s = "" if v is None else str(v).strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s[:-2] if s.endswith(".0") else s


def _norm_tipo(v):
    """'cc', 'C.C.', ' NIT ' -> 'CC' / 'NIT'. Cualquier otra cosa -> '' (sin tipo) o el texto crudo."""
    s = "" if v is None else str(v).strip()
    if s.lower() in ("", "nan", "none"):
        return ""
    t = s.upper().replace(".", "").replace(" ", "")
    return t


class TiposDocumento:
    """Consulta el tipo por cédula y recuerda qué cédulas cayeron en el valor por defecto."""

    def __init__(self, mapa=None, conflictos=None, invalidos=None, columna_presente=False):
        self.mapa = mapa or {}
        self.conflictos = set(conflictos or [])
        self.invalidos = dict(invalidos or {})
        self.columna_presente = columna_presente
        self._sin_tipo = set()
        self._en_conflicto = set()
        self._invalidos_usados = set()

    def tipo(self, cedula):
        c = _norm_cedula(cedula)
        if not c:
            return TIPO_POR_DEFECTO
        if c in self.conflictos:
            self._en_conflicto.add(c)
            return TIPO_POR_DEFECTO
        t = self.mapa.get(c)
        if t is None:
            if c in self.invalidos:
                self._invalidos_usados.add(c)
            self._sin_tipo.add(c)
            return TIPO_POR_DEFECTO
        return t

    def avisar(self):
        """Avisos en pantalla sobre las cédulas que se consultaron."""
        if not self.columna_presente:
            st.info("ℹ️ El libro de Clientes Activos todavía no trae la columna «tipo_documento»: "
                    "se usó CC para todos.")
            return
        if self._en_conflicto:
            ej = ", ".join(sorted(self._en_conflicto)[:10])
            st.warning(f"⚠️ {len(self._en_conflicto)} cédula(s) con tipo de documento contradictorio "
                       f"en Clientes Activos (CC y NIT a la vez). Se usó CC: {ej}"
                       + ("…" if len(self._en_conflicto) > 10 else ""))
        if self._invalidos_usados:
            vals = sorted({self.invalidos[c] for c in self._invalidos_usados})
            st.warning(f"⚠️ {len(self._invalidos_usados)} cédula(s) con un tipo de documento que no es "
                       f"CC ni NIT ({', '.join(vals[:5])}). Se usó CC.")
        sin = self._sin_tipo - self._invalidos_usados
        if sin:
            ej = ", ".join(sorted(sin)[:10])
            st.info(f"ℹ️ {len(sin)} cédula(s) sin tipo de documento en Clientes Activos (vacío o no "
                    f"encontradas). Se usó CC: {ej}" + ("…" if len(sin) > 10 else ""))


def construir_tipos_documento(archivo_clientes):
    """Lee el libro de clientes activos y devuelve un TiposDocumento (nunca lanza error)."""
    try:
        if hasattr(archivo_clientes, "seek"):
            archivo_clientes.seek(0)
        df = pd.read_excel(archivo_clientes, sheet_name="sheet1", dtype=str)
        if hasattr(archivo_clientes, "seek"):
            archivo_clientes.seek(0)
    except Exception:
        return TiposDocumento()

    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    if "IDEN" not in df.columns or "TIPO_DOCUMENTO" not in df.columns:
        return TiposDocumento(columna_presente=False)

    tipos_por_cedula = {}
    invalidos = {}
    for ced, tip in zip(df["IDEN"], df["TIPO_DOCUMENTO"]):
        c = _norm_cedula(ced)
        t = _norm_tipo(tip)
        if not c or not t:
            continue
        if t in TIPOS_VALIDOS:
            tipos_por_cedula.setdefault(c, set()).add(t)
        else:
            invalidos[c] = str(tip).strip()

    mapa, conflictos = {}, set()
    for c, ts in tipos_por_cedula.items():
        if len(ts) == 1:
            mapa[c] = next(iter(ts))
        else:
            conflictos.add(c)
    # un valor no reconocido no cuenta si la misma cédula tiene un tipo válido en otra fila
    invalidos = {c: v for c, v in invalidos.items() if c not in mapa and c not in conflictos}
    return TiposDocumento(mapa, conflictos, invalidos, columna_presente=True)


def _bytes_de(archivo):
    if archivo is None:
        return None
    if hasattr(archivo, "getvalue"):
        return archivo.getvalue()
    if hasattr(archivo, "read"):
        if hasattr(archivo, "seek"):
            archivo.seek(0)
        data = archivo.read()
        if hasattr(archivo, "seek"):
            archivo.seek(0)
        return data
    return None


def obtener_tipos_documento(tipos=None):
    """
    Devuelve un TiposDocumento. Si ya viene uno se usa tal cual; si no, se toma el libro de
    clientes cargado en la sesión o, si no hay, se baja del repositorio (GitHub).
    """
    if isinstance(tipos, TiposDocumento):
        return tipos
    if isinstance(tipos, dict):
        return TiposDocumento({_norm_cedula(k): v for k, v in tipos.items()}, columna_presente=True)

    datos = _bytes_de(st.session_state.get("archivo_clientes"))
    if datos is None:
        try:
            import requests
            resp = requests.get(f"{GITHUB_RAW}/CLIENTES_ACTIVOS.xlsx", timeout=30)
            if resp.status_code == 200:
                datos = resp.content
        except Exception:
            datos = None
    if datos is None:
        return TiposDocumento()

    huella = hashlib.md5(datos).hexdigest()
    cache = st.session_state.get("_tipos_documento_cache")
    if cache and cache[0] == huella:
        base = cache[1]
    else:
        base = construir_tipos_documento(io.BytesIO(datos))
        st.session_state["_tipos_documento_cache"] = (huella, base)
    # copia ligera para que el registro de avisos sea propio de cada generación
    return TiposDocumento(base.mapa, base.conflictos, base.invalidos, base.columna_presente)
