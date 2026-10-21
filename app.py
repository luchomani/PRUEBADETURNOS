# -*- coding: utf-8 -*-
"""
Sistema de Gestión de Turnos y Control de Accesos
Zona Franca Santander — Cúcuta
=================================================
Portales: Conductor | Vigilante | Operador | Pantalla Pública

La lógica (datos, seguridad, validación, Excel) vive en zfc_core.py;
este archivo solo contiene la interfaz.
"""
import importlib.util
import json
import re
import time
from pathlib import Path

import streamlit as st

import zfc_core as core

# ==========================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ==========================================================================
st.set_page_config(
    page_title="Turnos y Accesos | Zona Franca Santander",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ESTILOS = """<style>
    :root {
        --zfc-green-dark:  #12402A;
        --zfc-green-main:  #1B4D3E;
        --zfc-green-mid:   #2C6B56;
        --zfc-bg:          #F5F7F6;
        --zfc-card:        #FFFFFF;
        --zfc-text:        #2C3E50;
        --zfc-text-soft:   #6B7280;
        --zfc-border:      #E1E8E5;
        --zfc-danger:      #9B1C1C;
    }

    .stApp { background-color: var(--zfc-bg); }

    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: var(--zfc-text) !important;
    }

    h1, h2, h3, h4 {
        color: var(--zfc-green-dark) !important;
        font-weight: 700;
        letter-spacing: -0.3px;
    }

    .stTextInput label,
    .stSelectbox label,
    .stTextArea label,
    .stNumberInput label {
        color: #2C3E50 !important;
        font-weight: 600 !important;
        font-size: 15px !important;
    }

    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stNumberInput > div > div > input {
        background-color: #FFFFFF !important;
        color: #1B1B1B !important;
        border: 1.5px solid #C9D1CC !important;
        border-radius: 8px !important;
        font-size: 16px !important;
        padding: 12px 14px !important;
    }

    .stTextInput > div > div > input::placeholder,
    .stTextArea > div > div > textarea::placeholder,
    .stNumberInput > div > div > input::placeholder {
        color: #9CA3AF !important;
        opacity: 1 !important;
    }

    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #1B4D3E !important;
        box-shadow: 0 0 0 2px rgba(27,77,62,0.15) !important;
    }

    .stSelectbox > div > div {
        background-color: #FFFFFF !important;
        border: 1.5px solid #C9D1CC !important;
        border-radius: 8px !important;
        color: #1B1B1B !important;
    }
    .stSelectbox > div > div > div { color: #1B1B1B !important; }

    div[data-baseweb="popover"] ul,
    div[data-baseweb="menu"] ul { background-color: #FFFFFF !important; }
    div[data-baseweb="popover"] li,
    div[data-baseweb="menu"] li {
        color: #1B1B1B !important;
        background-color: #FFFFFF !important;
    }
    div[data-baseweb="popover"] li:hover,
    div[data-baseweb="menu"] li:hover {
        background-color: #E6F0EB !important;
        color: #12402A !important;
    }

    .stButton>button {
        background-color: var(--zfc-green-main) !important;
        color: #FFFFFF !important;
        border-radius: 8px;
        border: none;
        font-weight: 700;
        font-size: 16px;
        padding: 0.8rem 1.2rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    .stButton>button:hover {
        background-color: var(--zfc-green-mid) !important;
        color: #FFFFFF !important;
    }

    .stFormSubmitButton > button {
        background-color: var(--zfc-green-main) !important;
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 17px;
        padding: 0.9rem 1.5rem;
        border-radius: 8px;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #EDF1EE;
        border-radius: 6px 6px 0 0;
        padding: 10px 18px;
        font-weight: 600;
        color: #2C3E50;
    }
    .stTabs [aria-selected="true"] {
        background-color: var(--zfc-green-main) !important;
        color: #FFFFFF !important;
    }

    .ticket-card {
        background: linear-gradient(135deg, #12402A 0%, #1B4D3E 100%);
        color: #FFFFFF;
        padding: 32px 24px;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 10px 25px rgba(18,64,42,0.25);
        margin: 18px 0;
        border: 1px solid #0D2F1E;
    }
    .ticket-label {
        font-size: 11px;
        letter-spacing: 3px;
        color: #C7D5CD;
        font-weight: 600;
        text-transform: uppercase;
    }
    .ticket-placa {
        color: #FFD966;
        font-size: 44px;
        font-weight: 900;
        letter-spacing: 4px;
        margin: 6px 0;
        font-family: 'Courier New', monospace;
    }
    .ticket-turno {
        color: #FFFFFF;
        font-size: 52px;
        font-weight: 900;
        margin: 4px 0;
        font-family: 'Courier New', monospace;
    }
    .ticket-info { font-size: 14px; color: #D8E3DC; margin: 4px 0; }
    .ticket-divider { border-color: rgba(255,255,255,0.18); margin: 14px 0; }
    .ticket-mensaje {
        background: rgba(255,255,255,0.10);
        padding: 12px;
        border-radius: 8px;
        margin-top: 14px;
        border-left: 3px solid #FFD966;
    }

    .badge {
        display: inline-block;
        padding: 6px 14px;
        border-radius: 4px;
        font-weight: 700;
        font-size: 12px;
        letter-spacing: 1px;
        text-transform: uppercase;
    }
    .badge-espera     { background: #FBF3D9; color: #8A6A00; border: 1px solid #E5D7A0; }
    .badge-ingresar   { background: #E6F0EB; color: #12402A; border: 1px solid #B8D4C5; }
    .badge-salir      { background: #FCE9E9; color: #9B1C1C; border: 1px solid #E8BFBF; }
    .badge-zf         { background: #E0E7FF; color: #233876; border: 1px solid #B4C0E8; }
    .badge-finalizado { background: #EDEDED; color: #4B5563; border: 1px solid #D5D5D5; }

    .col-header-ing {
        background: var(--zfc-green-main);
        color: #FFFFFF;
        padding: 14px 16px;
        border-radius: 8px 8px 0 0;
        text-align: center;
        font-weight: 700;
        font-size: 13px;
        letter-spacing: 2px;
        text-transform: uppercase;
    }
    .col-header-sal {
        background: var(--zfc-danger);
        color: #FFFFFF;
        padding: 14px 16px;
        border-radius: 8px 8px 0 0;
        text-align: center;
        font-weight: 700;
        font-size: 13px;
        letter-spacing: 2px;
        text-transform: uppercase;
    }
    .empty-state {
        background: #EFEFEF;
        padding: 22px;
        text-align: center;
        color: #6B7280;
        border-radius: 0 0 8px 8px;
        font-style: italic;
        font-size: 14px;
    }

    .card-ingreso {
        background: #FFFFFF;
        border: 1px solid #B8D4C5;
        border-left: 4px solid var(--zfc-green-main);
        padding: 16px;
        border-radius: 0 0 8px 8px;
        margin-bottom: 10px;
        text-align: center;
    }
    .card-salida {
        background: #FFFFFF;
        border: 1px solid #E8BFBF;
        border-left: 4px solid var(--zfc-danger);
        padding: 16px;
        border-radius: 0 0 8px 8px;
        margin-bottom: 10px;
        text-align: center;
    }
    .card-placa {
        font-size: 30px;
        font-weight: 900;
        letter-spacing: 3px;
        font-family: 'Courier New', monospace;
        color: var(--zfc-green-dark);
        margin: 4px 0;
    }
    .card-placa-sal {
        font-size: 30px;
        font-weight: 900;
        letter-spacing: 3px;
        font-family: 'Courier New', monospace;
        color: var(--zfc-danger);
        margin: 4px 0;
    }
    .card-ficho {
        font-size: 20px;
        font-weight: 800;
        color: var(--zfc-text-soft);
        letter-spacing: 2px;
        text-transform: uppercase;
    }
    .card-detalle { font-size: 13px; color: #333; margin-top: 4px; }
    .card-aviso {
        display: inline-block;
        font-size: 12px;
        background: #F5F7F6;
        padding: 4px 10px;
        border-radius: 4px;
        margin-top: 6px;
        color: #2C3E50;
        font-weight: 600;
    }

    .role-card {
        background: #FFFFFF;
        border: 1px solid var(--zfc-border);
        border-top: 4px solid var(--zfc-green-main);
        border-radius: 10px;
        padding: 28px 20px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        height: 100%;
    }
    .role-titulo {
        font-size: 18px;
        font-weight: 700;
        color: var(--zfc-green-dark);
        letter-spacing: 0.5px;
        margin-bottom: 6px;
    }
    .role-desc { font-size: 13px; color: var(--zfc-text-soft); line-height: 1.5; }

    .inst-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 2px solid var(--zfc-green-main);
        padding-bottom: 12px;
        margin-bottom: 20px;
    }
    .inst-title {
        color: var(--zfc-green-dark);
        font-size: 22px;
        font-weight: 700;
        margin: 0;
    }
    .inst-sub {
        color: var(--zfc-text-soft);
        font-size: 12px;
        margin-top: 2px;
        letter-spacing: 0.5px;
    }

    div[data-testid="stMetricValue"] { color: var(--zfc-green-dark) !important; font-weight: 700; }
    div[data-testid="stMetricLabel"] { color: #556B2F !important; font-weight: 600; }
    div[data-testid="stAlert"] { border-radius: 8px; }
</style>"""
st.markdown(ESTILOS, unsafe_allow_html=True)

APP_DIR = Path(__file__).resolve().parent
REFRESCO_PANTALLA = 4      # segundos
REFRESCO_CONDUCTOR = 6
REFRESCO_VIGILANTE = 6
HAS_GEO = importlib.util.find_spec("streamlit_geolocation") is not None
PARAM_TOKEN = {"operador": "op_tok", "vigilante": "vig_tok"}
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SIN_BODEGA = "Seleccione una bodega..."

esc = core.esc


# ==========================================================================
# UTILIDADES DE INTERFAZ
# ==========================================================================
def _stretch(fn, *args, **kwargs):
    """Compatibilidad entre versiones de Streamlit (use_container_width ↔ width='stretch')."""
    try:
        return fn(*args, use_container_width=True, **kwargs)
    except TypeError:
        return fn(*args, width="stretch", **kwargs)


def boton(*a, **k):
    return _stretch(st.button, *a, **k)


def boton_form(*a, **k):
    return _stretch(st.form_submit_button, *a, **k)


def boton_descarga(*a, **k):
    return _stretch(st.download_button, *a, **k)


def imagen(*a, **k):
    return _stretch(st.image, *a, **k)


def tabla(*a, **k):
    return _stretch(st.dataframe, *a, **k)


def md_esc(s) -> str:
    """Escapa caracteres de Markdown para mostrar texto de usuario con st.markdown/st.info."""
    return re.sub(r"([\\`*_{}\[\]()#+\-.!|>~$<])", r"\\\1", "" if s is None else str(s))


def init_session():
    for k, v in (("historial_turnos", []), ("auth_operador", None), ("auth_vigilante", None),
                 ("mi_turno_id", None), ("ultimo_registro", 0.0)):
        st.session_state.setdefault(k, v)


def flash(msg: str, tipo: str = "success"):
    """Mensaje que sobrevive al st.rerun() (los st.success previos a un rerun nunca se veían)."""
    st.session_state["_flash"] = (tipo, msg)


def mostrar_flash():
    f = st.session_state.pop("_flash", None)
    if f:
        {"success": st.success, "error": st.error, "warning": st.warning}.get(f[0], st.info)(f[1])


@st.cache_resource
def limitador():
    return core.LimiteIntentos(max_intentos=5, ventana_s=300)


def ip_cliente() -> str:
    try:
        xff = st.context.headers.get("X-Forwarded-For", "") or ""
        return xff.split(",")[0].strip() or "local"
    except Exception:
        return "local"


def _primer_logo(nombres):
    for n in nombres:
        p = APP_DIR / n
        if p.exists():
            return str(p)
    return None


def render_header():
    c1, c2, c3 = st.columns([1.5, 3, 1.5])
    with c1:
        logo = _primer_logo(["LOGO ZFS-ZFC.jpeg", "LOGO ZFS-ZFC.jpg", "LOGO ZFS-ZFC.png"])
        if logo:
            imagen(logo)
    with c2:
        st.markdown(
            '<div style="text-align:center;">'
            '<div class="inst-title">Sistema de Enturnamiento y Control de Accesos</div>'
            '<div class="inst-sub" style="font-size: 14px; font-weight: 600; color: #1B4D3E; margin-top: 4px;">'
            'Ingresos y salidas de Resto del mundo</div>'
            '<div class="inst-sub">ZONA FRANCA DE CÚCUTA · OPERADA POR ZONA FRANCA SANTANDER</div>'
            '</div>',
            unsafe_allow_html=True,
        )
    with c3:
        logo = _primer_logo(["LOGO SANTANDER.jpg", "LOGO SANTANDER.png", "logo.jpeg"])
        if logo:
            imagen(logo)


# ==========================================================================
# AUTENTICACIÓN (token firmado; ya no se acepta un timestamp arbitrario)
# ==========================================================================
def autenticado(rol: str, cfg: dict) -> bool:
    clave = f"auth_{rol}"
    ts = st.session_state.get(clave)
    if ts and time.time() - ts < core.SESION_TTL:
        return True
    st.session_state[clave] = None
    tok = st.query_params.get(PARAM_TOKEN[rol])
    if tok:
        ts_tok = core.validar_token(core.secret_de(cfg), rol, tok)
        if ts_tok:
            st.session_state[clave] = ts_tok
            return True
    return False


def pantalla_login(rol: str, titulo: str, cfg: dict):
    st.markdown(f"<h3 style='text-align:center;'>{esc(titulo)}</h3>", unsafe_allow_html=True)
    _, centro, _ = st.columns([1, 1.2, 1])
    with centro:
        with st.form(f"login_{rol}"):
            u = st.text_input("Usuario", placeholder="Ingresa tu usuario")
            p = st.text_input("Contraseña", type="password", placeholder="Ingresa tu contraseña")
            enviado = boton_form("Iniciar Sesión")
        if not enviado:
            return
        lim, clave = limitador(), (rol, ip_cliente())
        espera = lim.segundos_bloqueo(clave)
        if espera:
            st.error(f"Demasiados intentos fallidos. Intenta de nuevo en {max(1, espera // 60)} min.")
        elif core.verificar_credenciales(cfg, rol, u.strip(), p):
            lim.exito(clave)
            st.session_state[f"auth_{rol}"] = time.time()
            st.query_params[PARAM_TOKEN[rol]] = core.crear_token(core.secret_de(cfg), rol)
            core.registrar_auditoria(rol, "login", "", clave[1])
            st.rerun()
        else:
            lim.fallo(clave)
            core.registrar_auditoria(rol, "login_fallido", "", clave[1])
            time.sleep(1)
            st.error("Credenciales incorrectas.")


def cerrar_sesion(rol: str):
    st.session_state[f"auth_{rol}"] = None
    try:
        if PARAM_TOKEN[rol] in st.query_params:
            del st.query_params[PARAM_TOKEN[rol]]
    except Exception:
        pass
    st.rerun()


def barra_sesion(rol: str, titulo: str):
    c1, c2 = st.columns([4, 1])
    with c1:
        st.markdown(f"### {titulo}")
        st.caption(f"Sesión activa · {core.ahora_str()}")
    with c2:
        if boton("Cerrar Sesión", key=f"logout_{rol}"):
            cerrar_sesion(rol)
    st.markdown("---")


# ==========================================================================
# COMPONENTES COMPARTIDOS
# ==========================================================================
def tarjeta_html(t: dict, salida: bool = False) -> str:
    clase = "card-salida" if salida else "card-ingreso"
    clase_placa = "card-placa-sal" if salida else "card-placa"
    msg = esc(t.get("mensaje_admin")) or "Sin mensaje"
    return (
        f'<div class="{clase}">'
        f'<div class="card-ficho">Ficho {esc(t.get("id_turno"))}</div>'
        f'<div class="{clase_placa}">{esc(t.get("placa"))}</div>'
        f'<div class="card-aviso">{msg}</div>'
        f'</div>'
    )


def render_tablero(turnos: list, cfg: dict, vacio_ing: str, vacio_sal: str):
    visibles = [t for t in turnos if core.visible_en_pantalla(t, cfg)]
    ingresos = [t for t in visibles if core.es(t, core.AUT_ING)]
    salidas = [t for t in visibles if core.es(t, core.AUT_SAL)]

    col_i, col_s = st.columns(2)
    with col_i:
        st.markdown('<div class="col-header-ing">AUTORIZADOS PARA INGRESAR</div>', unsafe_allow_html=True)
        if not ingresos:
            st.markdown(f'<div class="empty-state">{esc(vacio_ing)}</div>', unsafe_allow_html=True)
        for t in ingresos:
            st.markdown(tarjeta_html(t), unsafe_allow_html=True)
    with col_s:
        st.markdown('<div class="col-header-sal">AUTORIZADOS PARA SALIR</div>', unsafe_allow_html=True)
        if not salidas:
            st.markdown(f'<div class="empty-state">{esc(vacio_sal)}</div>', unsafe_allow_html=True)
        for t in salidas:
            st.markdown(tarjeta_html(t, salida=True), unsafe_allow_html=True)


@st.cache_data(show_spinner=False, max_entries=3)
def _excel_cacheado(turnos_json: str) -> bytes:
    return core.generar_excel(core.df_exportacion(json.loads(turnos_json)))


def excel_de(turnos: list) -> bytes:
    """El Excel solo se regenera cuando cambian los datos (no en cada rerun)."""
    return _excel_cacheado(json.dumps(turnos, sort_keys=True, ensure_ascii=False, default=str))


def render_historial(prefijo: str, con_auditoria: bool = False):
    st.markdown("#### Historial General de Turnos")
    turnos = core.cargar_datos()
    if not turnos:
        st.info("No hay turnos registrados.")
        return
    r = core.resumen(turnos)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total", r["total"])
    m2.metric("En Espera", r["espera"])
    m3.metric("Autorizados", r["autorizados"])
    m4.metric("En Zona Franca", r["en_zf"])
    m5.metric("Finalizados", r["finalizados"])
    st.markdown("---")

    df = core.df_exportacion(turnos)
    tabla(df.iloc[::-1], hide_index=True, height=450)  # más recientes primero
    boton_descarga(
        "Descargar Excel",
        data=excel_de(turnos),
        file_name=f"turnos_zfc_{core.ahora().strftime('%Y%m%d_%H%M')}.xlsx",
        mime=EXCEL_MIME,
        key=f"dl_hist_{prefijo}",
    )

    if con_auditoria:
        with st.expander("Registro de auditoría (últimos 200 eventos)"):
            eventos = core.leer_auditoria(200)
            if eventos:
                tabla(eventos, hide_index=True, height=300)
            else:
                st.caption("Sin eventos registrados.")


# ==========================================================================
# PORTAL: PANTALLA TV
# ==========================================================================
def portal_pantalla():
    @st.fragment(run_every=REFRESCO_PANTALLA)
    def _vista():
        turnos = core.cargar_datos()
        cfg = core.cargar_config()

        c1, c2, c3 = st.columns([1, 3, 1])
        with c1:
            logo = _primer_logo(["LOGO ZFS-ZFC.jpeg", "LOGO ZFS-ZFC.png"])
            if logo:
                imagen(logo)
        with c2:
            st.markdown(
                '<div style="text-align:center; padding: 8px 0;">'
                '<div style="color:#12402A; font-size:32px; font-weight:900;">CONTROL DE ACCESOS EN TIEMPO REAL</div>'
                '<div style="color:#6B7280; font-size:14px; letter-spacing:3px; margin-top:6px;">'
                'ZONA FRANCA DE CÚCUTA · ZONA FRANCA SANTANDER</div></div>',
                unsafe_allow_html=True,
            )
        with c3:
            logo = _primer_logo(["LOGO SANTANDER.jpg", "LOGO SANTANDER.png"])
            if logo:
                imagen(logo)

        st.markdown("<br>", unsafe_allow_html=True)
        render_tablero(turnos, cfg, "No hay vehículos autorizados para ingresar.",
                       "No hay vehículos autorizados para salir.")
        st.markdown("---")
        st.markdown(
            f"<div style='text-align:center; color:#6B7280; font-size:12px;'>"
            f"Última actualización: {core.ahora_str()}</div>",
            unsafe_allow_html=True,
        )

    _vista()


# ==========================================================================
# PORTAL: CONDUCTOR
# ==========================================================================
def ticket_html(t: dict) -> str:
    msg = esc(t.get("mensaje_admin")) or "En espera de autorización"
    return (
        '<div class="ticket-card">'
        '<div class="ticket-label">Turno Asignado</div><hr class="ticket-divider">'
        '<div class="ticket-label">Placa del Vehículo</div>'
        f'<div class="ticket-placa">{esc(t.get("placa"))}</div>'
        '<div class="ticket-label" style="margin-top:18px;">Número de Ficho</div>'
        f'<div class="ticket-turno">{esc(t.get("id_turno"))}</div><hr class="ticket-divider">'
        f'<div class="ticket-info"><b>Destino:</b> {esc(t.get("bodega") or "—")}</div>'
        f'<div class="ticket-info"><b>Movimiento:</b> {esc(t.get("tipo_movimiento") or "Ingresar")}</div>'
        '<div class="ticket-info" style="font-size:12px; color:#A9BCB2; margin-top:6px;">'
        f'Registrado: {esc(core.fecha_legible(t.get("fecha_hora")))}</div>'
        f'<div style="margin-top:16px;">{core.badge_html(t.get("estado"))}</div>'
        '<div class="ticket-mensaje"><div class="ticket-label" style="font-size:10px;">Mensaje del Operador</div>'
        f'<div style="color:#FFD966; font-weight:700; font-size:15px; margin-top:4px;">{msg}</div></div>'
        '</div>'
    )


@st.fragment(run_every=REFRESCO_CONDUCTOR)
def ticket_vivo():
    actual = core.buscar_turno(core.cargar_datos(), st.session_state.mi_turno_id)
    if not actual:
        st.warning("Tu turno ya no está disponible.")
        return
    st.markdown(ticket_html(actual), unsafe_allow_html=True)

    if actual.get("observaciones_conductor"):
        st.info(f"**Tus observaciones:** {md_esc(actual['observaciones_conductor'])}")
    if actual.get("fecha_ingreso_zf"):
        st.markdown(
            "<div style='text-align:center; color:#1B4D3E; font-weight:600; padding:10px; "
            "background:#E6F0EB; border-radius:6px; margin-top:10px;'>"
            f"<b>Ingreso confirmado por vigilante:</b> {esc(core.fecha_legible(actual['fecha_ingreso_zf']))}</div>",
            unsafe_allow_html=True,
        )
    if actual.get("fecha_salida_zf"):
        st.markdown(
            "<div style='text-align:center; color:#9B1C1C; font-weight:600; padding:10px; "
            "background:#FCE9E9; border-radius:6px; margin-top:10px;'>"
            f"<b>Salida confirmada por vigilante:</b> {esc(core.fecha_legible(actual['fecha_salida_zf']))}</div>",
            unsafe_allow_html=True,
        )
    if core.es_final(actual):
        st.success("Tu turno ha finalizado. ¡Gracias! Puedes registrar otro vehículo.")


def verificar_geocerca(cfg: dict) -> bool:
    """Geocerca opcional (requiere el paquete streamlit-geolocation).
    OJO: la ubicación la envía el navegador, así que es una barrera disuasoria, no infalible."""
    if not cfg.get("geocerca_activa") or not HAS_GEO:
        return True
    if st.session_state.get("geo_ok"):
        return True
    from streamlit_geolocation import streamlit_geolocation
    st.info("📍 Para registrarte debes encontrarte en la Zona Franca. "
            "Pulsa el botón y permite el acceso a tu ubicación.")
    loc = streamlit_geolocation() or {}
    lat, lon = loc.get("latitude"), loc.get("longitude")
    if lat is None or lon is None:
        return False
    dentro, dist = core.evaluar_geocerca(lat, lon, cfg)
    if dentro:
        st.session_state["geo_ok"] = True
        return True
    st.error(f"Estás a {dist:.1f} km de la Zona Franca. El máximo permitido es {cfg['radio_geofence_km']} km.")
    return False


def formulario_registro(cfg: dict):
    if not verificar_geocerca(cfg):
        return
    st.markdown("Completa tus datos para asignarte un turno.")
    st.caption("Los campos con (*) son obligatorios. La bodega es opcional.")

    with st.form("form_conductor", clear_on_submit=False):
        placa = st.text_input("Placa del Vehículo *", placeholder="Ej: ABC123", max_chars=12)
        nombre = st.text_input("Nombre Completo del Conductor *", placeholder="Ej: Juan Pérez", max_chars=80)
        cedula = st.text_input("Número de Cédula *", placeholder="Ej: 1234567890", max_chars=20)
        telefono = st.text_input("Teléfono de Contacto *", placeholder="Ej: 3001234567", max_chars=20)
        bodega = st.selectbox("Bodega de Destino (Opcional)", [SIN_BODEGA] + list(cfg.get("bodegas", [])))
        observaciones = st.text_area(
            "Observaciones (Opcional)", max_chars=300,
            placeholder="Si tienes alguna nota sobre tu vehículo o carga, escríbela aquí...",
        )
        acepta = st.checkbox(
            "Autorizo el tratamiento de mis datos personales (nombre, cédula, teléfono y placa) "
            "para la gestión de turnos y el control de accesos de la Zona Franca, conforme a la Ley 1581 de 2012. *"
        )
        enviado = boton_form("Obtener Turno")

    if not enviado:
        return

    reg, errores = core.validar_registro(placa, nombre, cedula, telefono, observaciones)
    if not acepta:
        errores.append("Debes autorizar el tratamiento de datos personales para continuar.")
    if time.time() - st.session_state.ultimo_registro < 30:
        errores.append("Espera unos segundos antes de registrar otro vehículo.")
    if errores:
        st.error("⚠️ Corrige los siguientes errores:\n\n" + "\n".join(f"• {e}" for e in errores))
        return

    reg["bodega"] = "" if bodega == SIN_BODEGA else bodega
    estado, turno = core.crear_turno(reg)
    if estado == "duplicado":
        if turno.get("cedula") == reg["cedula"]:  # es el mismo conductor: se recupera su turno
            st.session_state.mi_turno_id = turno["uid"]
            st.query_params["t"] = turno["uid"]
            flash(f"La placa {turno['placa']} ya tenía el turno {turno['id_turno']}. Aquí está tu ticket.", "info")
            st.rerun()
        st.error("Esa placa ya tiene un turno activo. Si es tu vehículo, usa «Visualizar estado de mi placa» "
                 "con la cédula con la que te registraste.")
        return

    st.session_state.ultimo_registro = time.time()
    st.session_state.mi_turno_id = turno["uid"]
    st.query_params["t"] = turno["uid"]
    flash(f"Turno asignado: {turno['id_turno']}")
    st.rerun()


def consulta_por_placa():
    with st.expander("Visualizar estado de mi placa"):
        p = st.text_input("Placa", key="rec_placa", max_chars=12)
        c = st.text_input("Cédula del conductor registrado", key="rec_cedula", max_chars=20)
        if boton("Buscar", key="btn_rec"):
            lim, clave = limitador(), ("consulta", ip_cliente())
            if lim.segundos_bloqueo(clave):
                st.error("Demasiadas consultas. Intenta de nuevo en unos minutos.")
                return
            if not p.strip() or not c.strip():
                st.error("Ingresa la placa y la cédula.")
                return
            t = core.buscar_activo(p, c)
            if t:
                lim.exito(clave)
                st.session_state.mi_turno_id = t["uid"]
                st.query_params["t"] = t["uid"]
                st.rerun()
            lim.fallo(clave)
            st.error("No se encontró un turno activo para esa placa y cédula.")


def portal_conductor(cfg: dict):
    render_header()
    st.markdown("<h3 style='text-align:center;'>Portal del Conductor</h3>", unsafe_allow_html=True)
    st.markdown(
        '<div style="background: linear-gradient(135deg, #E6F0EB 0%, #D1E7DD 100%); border: 1px solid #B8D4C5; '
        'border-left: 5px solid #1B4D3E; padding: 16px 20px; border-radius: 8px; margin: 15px 0 25px 0; '
        'color: #12402A; box-shadow: 0 4px 6px rgba(0,0,0,0.04); font-size: 15px; font-weight: 600; '
        'text-align: center;">¡Bienvenido a la Zona Franca de Cúcuta! Si vienes de algún Puerto o de otra '
        'Zona Franca, te invitamos a que te enturnes.</div>',
        unsafe_allow_html=True,
    )

    _, centro, _ = st.columns([1, 2, 1])
    with centro:
        mostrar_flash()
        mi_turno = None
        if st.session_state.mi_turno_id:
            mi_turno = core.buscar_turno(core.cargar_datos(), st.session_state.mi_turno_id)
            if mi_turno is None:
                st.session_state.mi_turno_id = None

        if mi_turno:
            ticket_vivo()
            st.markdown("<br>", unsafe_allow_html=True)
            col_a, col_b = st.columns(2)
            with col_a:
                if boton("Actualizar Estado Manualmente"):
                    st.rerun()
            with col_b:
                if boton("Registrar otro vehículo"):
                    st.session_state.mi_turno_id = None
                    st.query_params.clear()
                    st.query_params["modo"] = "conductor"
                    st.rerun()
        else:
            formulario_registro(cfg)

        st.markdown("---")
        consulta_por_placa()


# ==========================================================================
# PORTAL: VIGILANTE
# ==========================================================================
def _resultado_confirmacion(res, placa: str, etiqueta: str):
    estado, _ = res
    if estado == "ok":
        flash(f"{etiqueta} autorizado para {placa} a las {core.ahora().strftime('%H:%M:%S')}.")
    elif estado == "estado_cambio":
        flash(f"El estado de {placa} cambió mientras lo revisabas (el operador lo modificó). "
              "Verifica la lista actualizada.", "warning")
    else:
        flash("El turno ya no existe.", "error")


def _expander_vehiculo(t: dict, tipo: str):
    uid = t["uid"]
    with st.expander(f"{t['placa']} · {t['id_turno']}", expanded=False):
        st.markdown(f"**Conductor:** {md_esc(t.get('nombre'))} — C.C. {md_esc(t.get('cedula'))}")
        st.markdown(f"**Teléfono:** {md_esc(t.get('telefono') or '—')}")
        if t.get("bodega"):
            st.markdown(f"**Bodega:** {md_esc(t['bodega'])}")
        if tipo == "ingreso":
            st.caption(f"Registrado: {core.fecha_legible(t.get('fecha_hora'))}")
        elif t.get("fecha_ingreso_zf"):
            st.caption(f"Ingresó: {core.fecha_legible(t['fecha_ingreso_zf'])}")
        if t.get("observaciones_conductor"):
            st.info(f"**Observación del Conductor:** {md_esc(t['observaciones_conductor'])}")
        if t.get("mensaje_admin"):
            st.markdown(f"**Mensaje Operador:** {md_esc(t['mensaje_admin'])}")

        obs_key = f"obs_{tipo}_{uid}"
        st.text_input("Observación del vigilante (opcional)", key=obs_key, max_chars=300)
        etiqueta, key = ("Ingreso Autorizado", f"btn_in_{uid}") if tipo == "ingreso" \
            else ("Salida Autorizada", f"btn_out_{uid}")
        if boton(etiqueta, key=key):
            res = core.confirmar_movimiento(uid, tipo, st.session_state.get(obs_key, ""))
            _resultado_confirmacion(res, t["placa"], "Ingreso" if tipo == "ingreso" else "Salida")
            st.rerun()


@st.fragment(run_every=REFRESCO_VIGILANTE)
def panel_vigilante():
    turnos = core.cargar_datos()
    pend_ing = [t for t in turnos if core.es(t, core.AUT_ING)]
    pend_sal = [t for t in turnos if core.es(t, core.AUT_SAL)]

    m1, m2, m3 = st.columns(3)
    m1.metric("Pendientes de Ingreso", len(pend_ing))
    m2.metric("Pendientes de Salida", len(pend_sal))
    m3.metric("Total en Sistema", len(turnos))

    buscar = core.limpiar_texto(st.text_input("🔎 Buscar por placa o ficho", key="vig_buscar"), 20).upper()

    def coincide(t):
        return (not buscar or buscar in core.norm_placa(t.get("placa"))
                or buscar in str(t.get("id_turno", "")).upper())

    st.markdown("---")
    col_i, col_s = st.columns(2)
    with col_i:
        st.markdown('<div class="col-header-ing">CHECK IN · INGRESOS</div>', unsafe_allow_html=True)
        lista = [t for t in pend_ing if coincide(t)]
        if not lista:
            st.markdown('<div class="empty-state">Sin vehículos pendientes de ingreso.</div>', unsafe_allow_html=True)
        for t in lista:
            _expander_vehiculo(t, "ingreso")
    with col_s:
        st.markdown('<div class="col-header-sal">CHECK OUT · SALIDAS</div>', unsafe_allow_html=True)
        lista = [t for t in pend_sal if coincide(t)]
        if not lista:
            st.markdown('<div class="empty-state">Sin vehículos pendientes de salida.</div>', unsafe_allow_html=True)
        for t in lista:
            _expander_vehiculo(t, "salida")


def portal_vigilante(cfg: dict):
    render_header()
    if not autenticado("vigilante", cfg):
        pantalla_login("vigilante", "Acceso Vigilante", cfg)
        return
    barra_sesion("vigilante", "Portal del Vigilante")
    mostrar_flash()
    if cfg.get("cambiar_pass_vigilante"):
        st.warning("Se está usando la contraseña por defecto del vigilante. Pide al operador que la cambie.")

    tab_control, tab_hist = st.tabs(["Control de Acceso", "Historial"])
    with tab_control:
        panel_vigilante()
    with tab_hist:
        render_historial("vig")


# ==========================================================================
# PORTAL: OPERADOR
# ==========================================================================
FILTROS = {
    "Todos": lambda t: True,
    "En Espera": lambda t: core.es(t, core.ESPERA),
    "Autorizados": lambda t: core.es(t, core.AUT_ING, core.AUT_SAL),
    "En Zona Franca": lambda t: core.es(t, core.EN_ZF),
    "Finalizados": core.es_final,
}


def _push_historial(entry: dict):
    hist = st.session_state.historial_turnos
    hist.append(entry)
    del hist[:-core.MAX_HISTORIAL_UNDO]


def _deshacer():
    entry = st.session_state.historial_turnos.pop()
    res = core.aplicar_deshacer(entry)
    mensajes = {
        "ok": ("Se deshizo la última acción.", "success"),
        "conflicto": ("No se pudo deshacer: el turno fue modificado después por otro usuario "
                      "(p. ej. el vigilante).", "error"),
        "no_existe": ("No se pudo deshacer: el turno ya no existe.", "error"),
        "ya_existe": ("Nada que restaurar: los registros ya existen.", "warning"),
    }
    msg, tipo = mensajes.get(res, ("Acción no reconocida.", "error"))
    flash(msg, tipo)
    st.rerun()


def _etiqueta_turno(t: dict) -> str:
    return f"{t['id_turno']} ({core.fecha_legible(t.get('fecha_hora'))}) — {t['placa']} — {t['estado']}"


def form_edicion(turno: dict, cfg: dict):
    uid = turno["uid"]
    bodegas = [SIN_BODEGA] + list(cfg.get("bodegas", []))
    actual = turno.get("bodega", "")
    if actual and actual not in bodegas:  # bodega retirada de la lista: no se pierde el dato
        bodegas.append(actual)
    idx_b = bodegas.index(actual) if actual in bodegas else 0
    idx_e = next((i for i, e in enumerate(core.ESTADOS) if e.upper() == core.estado(turno)), 0)

    with st.form(f"form_g_{uid}"):
        st.markdown("##### Editar Datos del Conductor")
        f_ficho = st.text_input("Número de Ficho (Tiquet)", value=turno.get("id_turno", ""), max_chars=12)
        f_placa = st.text_input("Placa del Vehículo", value=turno.get("placa", ""), max_chars=12)
        f_nombre = st.text_input("Nombre del Conductor", value=turno.get("nombre", ""), max_chars=80)
        f_cedula = st.text_input("Cédula", value=turno.get("cedula", ""), max_chars=20)
        f_tel = st.text_input("Teléfono", value=turno.get("telefono", ""), max_chars=20)
        f_bodega = st.selectbox("Bodega de Destino", bodegas, index=idx_b)
        st.text_area(
            "Observaciones del Conductor (solo lectura)",
            value=turno.get("observaciones_conductor", ""), disabled=True,
            help="Este campo solo puede ser modificado por el conductor al momento de registrarse.",
        )
        st.caption(f"**Registrado originalmente:** {core.fecha_legible(turno.get('fecha_hora'))}")

        st.markdown("---")
        st.markdown("##### Estado y Gestión")
        f_estado = st.selectbox("Estado", core.ESTADOS, index=idx_e)
        f_msg = st.text_input("Mensaje para el conductor (ej: Muelle 3)",
                              value=turno.get("mensaje_admin", ""), max_chars=120)
        f_mostrar = st.checkbox("Mostrar en pantalla pública", value=bool(turno.get("mostrar_en_pantalla", False)))
        enviado = boton_form("Guardar Cambios")

    if not enviado:
        return

    cambios, errores = {}, []
    for campo, validador, raw in (("id_turno", "ficho", f_ficho), ("placa", "placa", f_placa),
                                  ("nombre", "nombre", f_nombre), ("cedula", "cedula", f_cedula),
                                  ("telefono", "telefono", f_tel)):
        if str(raw).strip() == str(turno.get(campo, "")).strip():
            continue  # sin cambios: no se revalida (datos antiguos no bloquean la edición)
        limpio, err = core.validar_campo(validador, raw)
        if err:
            errores.append(err)
        else:
            cambios[campo] = limpio
    if errores:
        st.error("⚠️ " + " ".join(errores))
        return

    cambios.update({
        "bodega": "" if f_bodega == SIN_BODEGA else f_bodega,
        "estado": f_estado,
        "mensaje_admin": core.limpiar_texto(f_msg, 120),
        "mostrar_en_pantalla": bool(f_mostrar),
    })
    res = core.actualizar_turno(uid, cambios)
    mensajes = {
        "placa_duplicada": ("❌ Esa placa ya está registrada en otro turno activo.", "error"),
        "ficho_duplicado": ("❌ Ese Ficho ya está en uso en la jornada actual. Asigna uno diferente.", "error"),
        "estado_invalido": ("❌ Estado no válido.", "error"),
        "no_existe": ("❌ El turno ya no existe.", "error"),
    }
    if res["estado"] in mensajes:
        st.error(mensajes[res["estado"]][0])
        return
    _push_historial({"accion": "update", "antes": res["antes"], "despues": res["despues"]})
    flash("Cambios guardados exitosamente.")
    st.rerun()


def zona_eliminar(turno: dict, cfg: dict):
    uid = turno["uid"]
    with st.expander("🗑️ Eliminar este turno permanentemente"):
        st.warning("Esta acción elimina el registro. Podrás deshacerla desde «Deshacer última acción» "
                   "mientras no cierres esta sesión.")
        confirmar = st.checkbox("Confirmo que deseo eliminar este turno", key=f"del_chk_{uid}")
        pwd = st.text_input("Contraseña de Operador", type="password", key=f"del_pwd_{uid}")
        if boton("Eliminar turno", key=f"del_btn_{uid}", disabled=not confirmar):
            lim, clave = limitador(), ("eliminar", ip_cliente())
            if lim.segundos_bloqueo(clave):
                st.error("Demasiados intentos. Espera unos minutos.")
            elif not core.verificar_password(pwd, cfg.get("hash_admin", "")):
                lim.fallo(clave)
                st.error("❌ Contraseña incorrecta. No se eliminó el turno.")
            else:
                lim.exito(clave)
                antes = core.eliminar_turno(uid)
                if antes:
                    _push_historial({"accion": "delete", "antes": antes})
                    flash(f"Turno {antes.get('id_turno', '')} eliminado.")
                else:
                    flash("El turno ya no existía.", "warning")
                st.rerun()


def tab_gestion(cfg: dict):
    turnos = core.cargar_datos()

    if st.session_state.historial_turnos:  # visible aunque ya no queden turnos
        if st.button("Deshacer Última Acción", type="primary", key="btn_undo"):
            _deshacer()

    if not turnos:
        st.info("No hay turnos registrados.")
        return

    filtro = st.selectbox("Filtrar por estado:", list(FILTROS.keys()), key="filtro_op")
    lista = [t for t in turnos if FILTROS[filtro](t)]
    if not lista:
        st.info("No hay turnos que coincidan con el filtro.")
        return

    por_uid = {t["uid"]: t for t in lista}
    sel = st.selectbox("Selecciona un turno para gestionar:", list(por_uid.keys()),
                       format_func=lambda u: _etiqueta_turno(por_uid[u]), key="sel_op")
    turno = por_uid[sel]
    form_edicion(turno, cfg)
    zona_eliminar(turno, cfg)


def tab_pantalla(cfg: dict):
    st.markdown("#### Vista previa de Pantalla Pública")
    st.caption(f"Los vehículos autorizados se muestran por {cfg.get('tiempo_en_pantalla_min', 30)} "
               "minutos y luego se ocultan automáticamente.")
    render_tablero(core.cargar_datos(), cfg, "Sin vehículos visibles.", "Sin vehículos visibles.")


PORTALES_QR = [
    {"clave": "conductor", "titulo": "CONDUCTOR", "subtitulo": "Punto de enturnamiento",
     "descripcion": "El conductor escanea este QR y accede al formulario de registro.",
     "color_hex": "#1B4D3E", "archivo_png": "qr_conductor_zfc.png"},
    {"clave": "vigilante", "titulo": "VIGILANTE", "subtitulo": "Punto de control / garita",
     "descripcion": "El vigilante escanea este QR y accede a su portal de check-in/out.",
     "color_hex": "#0D4A56", "archivo_png": "qr_vigilante_zfc.png"},
    {"clave": "pantalla", "titulo": "PANTALLA TV", "subtitulo": "Televisor informativo",
     "descripcion": "Abre este enlace en las pantallas para mostrar vehículos autorizados.",
     "color_hex": "#12402A", "archivo_png": "qr_pantalla_zfc.png"},
]


def tab_qr(cfg: dict):
    st.markdown("#### Centro de Códigos QR")
    st.caption("Configura la URL base de tu aplicación y descarga los códigos QR de cada portal.")
    st.markdown("##### Paso 1 · Configura la URL base")

    with st.form("form_url_base_qr"):
        url_in = st.text_input("URL base de la aplicación", value=cfg.get("url_base", "").strip().rstrip("/"),
                               placeholder="https://tu-app.streamlit.app")
        enviado = boton_form("Guardar URL base")
    if enviado:
        url = url_in.strip().rstrip("/")
        if url and not re.match(r"^https?://[^\s]+$", url):
            st.error("La URL debe comenzar con http:// o https:// y no contener espacios.")
        else:
            nuevo = core.cargar_config()
            nuevo["url_base"] = url
            core.guardar_config(nuevo)
            flash("URL base actualizada.")
            st.rerun()

    base_url = cfg.get("url_base", "").strip().rstrip("/")
    st.markdown("---")
    if not base_url:
        st.warning("Configura primero la URL base (paso 1).")
        return
    if not core.HAS_QR:
        st.error("La librería **qrcode[pil]** no está instalada. Agrégala a `requirements.txt`.")
        return

    st.markdown("##### Paso 2 · Descarga e imprime los códigos QR")
    cols = st.columns(3)
    for idx, p in enumerate(PORTALES_QR):
        url_portal = f"{base_url}/?modo={p['clave']}"
        qr_bytes = core.generar_qr_png(url_portal)
        with cols[idx]:
            st.markdown(
                f'<div style="background:#FFFFFF; border: 1px solid #E1E8E5; border-top: 4px solid {p["color_hex"]}; '
                f'border-radius: 10px; padding: 18px 16px 8px 16px; text-align: center; margin-bottom: 12px;">'
                f'<div style="color:{p["color_hex"]}; font-weight:700; font-size:14px; letter-spacing:2px;">{p["titulo"]}</div>'
                f'<div style="color:#6B7280; font-size:11px; margin-top:2px;">{p["subtitulo"]}</div></div>',
                unsafe_allow_html=True,
            )
            if qr_bytes:
                imagen(qr_bytes)
                st.markdown("<div style='text-align:center; margin-top:-10px; margin-bottom:10px;'>"
                            "<b>O ingresa a este enlace:</b></div>", unsafe_allow_html=True)
                st.code(url_portal, language=None)
                st.caption(p["descripcion"])
                boton_descarga(f"Descargar QR · {p['titulo'].title()}", data=qr_bytes,
                               file_name=p["archivo_png"], mime="image/png", key=f"dl_qr_{p['clave']}")
            else:
                st.error("No se pudo generar el QR.")


def tab_config(cfg: dict):
    st.markdown("#### Configuración del Sistema")
    with st.form("form_cfg"):
        st.markdown("**Tiempo en pantalla**")
        tiempo = st.number_input("Minutos que un vehículo permanece visible tras su autorización",
                                 min_value=1, max_value=240, value=int(cfg.get("tiempo_en_pantalla_min", 30)))

        st.markdown("---")
        st.markdown("**Rango de geolocalización (opcional)**")
        geo_activa = st.checkbox(
            "Exigir ubicación dentro del radio para registrarse",
            value=bool(cfg.get("geocerca_activa", False)) and HAS_GEO,
            disabled=not HAS_GEO,
            help="Requiere instalar el paquete «streamlit-geolocation» (añádelo a requirements.txt).",
        )
        if not HAS_GEO:
            st.caption("Para activarla agrega `streamlit-geolocation` a requirements.txt.")
        radio_val = st.number_input("Radio máximo permitido en kilómetros desde la Zona Franca",
                                    min_value=0.5, max_value=200.0,
                                    value=float(cfg.get("radio_geofence_km", 3.0)), step=0.5)
        cg1, cg2 = st.columns(2)
        with cg1:
            lat_in = st.number_input("Latitud Zona Franca", value=float(cfg.get("zfc_lat", 7.90750)), format="%.5f")
        with cg2:
            lon_in = st.number_input("Longitud Zona Franca", value=float(cfg.get("zfc_lon", -72.50472)), format="%.5f")

        st.markdown("---")
        st.markdown("**URL base de la aplicación**")
        url_in = st.text_input("URL base (ej: https://tu-app.streamlit.app)", value=cfg.get("url_base", ""))

        st.markdown("---")
        st.markdown("**Credenciales Operador**")
        u_admin = st.text_input("Usuario Operador", value=cfg.get("usuario_admin", ""))
        p_admin = st.text_input("Nueva contraseña Operador (vacío = no cambiar)", type="password")
        p_admin2 = st.text_input("Confirmar nueva contraseña Operador", type="password")

        st.markdown("---")
        st.markdown("**Credenciales Vigilante**")
        u_vig = st.text_input("Usuario Vigilante", value=cfg.get("usuario_vigilante", ""))
        p_vig = st.text_input("Nueva contraseña Vigilante (vacío = no cambiar)", type="password")
        p_vig2 = st.text_input("Confirmar nueva contraseña Vigilante", type="password")

        st.markdown("---")
        st.markdown("**Bodegas disponibles** (una por línea)")
        bodegas_txt = st.text_area("Lista de bodegas", value="\n".join(cfg.get("bodegas", [])), height=250)
        enviado = boton_form("Guardar Configuración")

    if enviado:
        errores = []
        url = url_in.strip().rstrip("/")
        if url and not re.match(r"^https?://[^\s]+$", url):
            errores.append("La URL base debe comenzar con http:// o https:// y no contener espacios.")
        u_admin, u_vig = u_admin.strip(), u_vig.strip()
        if not u_admin or not u_vig:
            errores.append("Los usuarios no pueden estar vacíos.")
        elif u_admin == u_vig:
            errores.append("El usuario del operador y el del vigilante deben ser distintos.")

        nuevo = core.cargar_config()  # base fresca: conserva hashes y clave secreta
        cambio_pass = False
        for rol, clave_hash, flag, nueva, conf in (
            ("Operador", "hash_admin", "cambiar_pass_admin", p_admin, p_admin2),
            ("Vigilante", "hash_vigilante", "cambiar_pass_vigilante", p_vig, p_vig2),
        ):
            if not (nueva or conf):
                continue
            if nueva != conf:
                errores.append(f"Las contraseñas del {rol} no coinciden.")
                continue
            err = core.validar_password_nueva(nueva)
            if err:
                errores.append(f"{rol}: {err}")
                continue
            nuevo[clave_hash] = core.hash_password(nueva)
            nuevo[flag] = False
            cambio_pass = True

        if errores:
            st.error("⚠️ " + " ".join(errores))
        else:
            bodegas = list(dict.fromkeys(core.limpiar_texto(b, 80) for b in bodegas_txt.split("\n")
                                         if core.limpiar_texto(b, 80)))
            nuevo.update({
                "tiempo_en_pantalla_min": int(tiempo),
                "geocerca_activa": bool(geo_activa),
                "radio_geofence_km": float(radio_val),
                "zfc_lat": float(lat_in),
                "zfc_lon": float(lon_in),
                "url_base": url,
                "usuario_admin": u_admin,
                "usuario_vigilante": u_vig,
                "bodegas": bodegas,
            })
            if cambio_pass:
                core.rotar_secreto(nuevo)  # invalida las sesiones/enlaces emitidos con la clave anterior
            core.guardar_config(nuevo)
            core.registrar_auditoria("operador", "config", "", "cambio de contraseña" if cambio_pass else "ajustes")
            if cambio_pass:  # mantiene abierta la sesión de quien acaba de guardar
                st.session_state["auth_operador"] = time.time()
                st.query_params[PARAM_TOKEN["operador"]] = core.crear_token(core.secret_de(nuevo), "operador")
            flash("Configuración guardada." + (" Las demás sesiones deberán iniciar de nuevo." if cambio_pass else ""))
            st.rerun()

    st.markdown("---")
    st.markdown("#### Limpieza de Jornada")
    turnos = core.cargar_datos()
    st.metric("Total registros en historial", len(turnos))
    st.caption("Esta acción vacía el historial. Antes se guarda automáticamente una copia en la carpeta "
               "`archivo/` y puedes deshacerla desde «Gestión de Turnos». Aun así, descarga el respaldo en Excel.")
    confirmar = st.checkbox("Confirmo que deseo limpiar TODO el historial", key="clean_confirm")
    if boton("Limpiar todo el historial", key="clean_todos", disabled=not confirmar or not turnos):
        eliminado = core.limpiar_historial()
        if eliminado:
            _push_historial({"accion": "clear", "antes": eliminado})
        flash(f"Se archivaron y eliminaron {len(eliminado)} turnos. Sistema limpio.")
        st.rerun()

    st.markdown("---")
    st.markdown("#### Exportar antes de limpiar")
    if turnos:
        boton_descarga(
            "Descargar respaldo del historial completo (Excel)",
            data=excel_de(turnos),
            file_name=f"respaldo_turnos_{core.ahora().strftime('%Y%m%d_%H%M')}.xlsx",
            mime=EXCEL_MIME,
            key="dl_respaldo",
        )


def portal_operador(cfg: dict):
    render_header()
    if not autenticado("operador", cfg):
        pantalla_login("operador", "Acceso Operador", cfg)
        return
    barra_sesion("operador", "Panel del Operador")
    mostrar_flash()
    if cfg.get("cambiar_pass_admin") or cfg.get("cambiar_pass_vigilante"):
        st.warning("⚠️ Se están usando contraseñas por defecto. Cámbialas ahora en la pestaña «Configuración».")

    t_gestion, t_pantalla, t_qr, t_config, t_hist = st.tabs(
        ["Gestión de Turnos", "Pantalla Pública", "Código QR", "Configuración", "Historial"])
    with t_gestion:
        tab_gestion(cfg)
    with t_pantalla:
        tab_pantalla(cfg)
    with t_qr:
        tab_qr(cfg)
    with t_config:
        tab_config(cfg)
    with t_hist:
        render_historial("op", con_auditoria=True)


# ==========================================================================
# HOME
# ==========================================================================
def home():
    render_header()
    st.markdown("<h3 style='text-align:center; margin-top:10px;'>Selecciona tu Portal de Acceso</h3>",
                unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    tarjetas = [
        ("conductor", "CONDUCTOR", "Regístrate, obtén tu turno y consulta tu estado desde tu celular.",
         "Ingresar como Conductor", "home_btn_cond"),
        ("vigilante", "VIGILANTE", "Autoriza el ingreso y salida de vehículos con un solo clic.",
         "Ingresar como Vigilante", "home_btn_vig"),
        ("operador", "OPERADOR", "Autoriza turnos, gestiona la pantalla y administra el sistema.",
         "Ingresar como Operador", "home_btn_op"),
        ("pantalla", "PANTALLA TV", "Vista pública para televisores. Solo lectura, sin interacción.",
         "Abrir Pantalla Pública", "home_btn_pant"),
    ]
    for col, (modo, titulo, desc, etiqueta, key) in zip(st.columns(4), tarjetas):
        with col:
            st.markdown(f'<div class="role-card"><div class="role-titulo">{titulo}</div>'
                        f'<div class="role-desc">{desc}</div></div>', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            if boton(etiqueta, key=key):
                st.query_params["modo"] = modo
                st.rerun()

    st.markdown("---")
    st.caption("Zona Franca de Cúcuta · Operada por Zona Franca Santander · Sistema de Automatización Aduanera")


# ==========================================================================
# ENRUTADOR
# ==========================================================================
def main():
    init_session()
    cfg = core.cargar_config()
    modo = (st.query_params.get("modo") or "").strip().lower()

    if st.session_state.mi_turno_id is None and st.query_params.get("t"):
        st.session_state.mi_turno_id = st.query_params.get("t")

    if modo == "pantalla":
        portal_pantalla()
    elif modo == "conductor":
        portal_conductor(cfg)
    elif modo == "vigilante":
        portal_vigilante(cfg)
    elif modo == "operador":
        portal_operador(cfg)
    else:
        home()


main()
