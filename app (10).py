# -*- coding: utf-8 -*-
"""
Sistema de Gestión de Turnos y Control de Accesos
Zona Franca Santander — Cúcuta
=================================================
Portales: Conductor | Vigilante | Operador | Pantalla Pública
"""
import os
import json
import math
import time
import copy
import uuid
from datetime import datetime, timedelta, timezone
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import qrcode
    HAS_QR = True
except ImportError:
    HAS_QR = False

# ==========================================================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================================================
st.set_page_config(
    page_title="Turnos y Accesos | Zona Franca Santander",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==========================================================================
# CSS CORPORATIVO
# ==========================================================================
st.markdown("""
<style>
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
</style>
""", unsafe_allow_html=True)

# ==========================================================================
# CAPA DE DATOS
# ==========================================================================
DATA_FILE = "turnos_zfc.json"
CONFIG_FILE = "config_zfc.json"

DEFAULT_CONFIG = {
    "config_version": 2, 
    "tiempo_en_pantalla_min": 30,
    "usuario_admin": "ZONAFRANCA123-",
    "contrasena_admin": "ZONAFRANCA123-",
    "usuario_vigilante": "VIGILANTE123-",
    "contrasena_vigilante": "VIGILANTE123-",
    "url_base": "",
    "radio_geofence_km": 3.0,
    "zfc_lat": 7.90750,
    "zfc_lon": -72.50472,
    "bodegas": [
        "OTRA BODEGA",
        "ALFRANCA CUCUTA S.A.S.",
        "ALMACODER SAS",
        "ALMANORTE S.A.S",
        "BERZATEX S.A.S.",
        "COSMOGREEN SAS",
        "DEL BEN SAS",
        "GLEX S.A.S.",
        "GRUPO EMPRESARIAL SUAPEL S.A.S.",
        "HSQ GROUP INTERNATIONAL SAS",
        "INDULOGISTICS SAS",
        "INTERNATIONAL FRONTIER S.A.S.",
        "INVERSIONES CARGAMAX HYJ S.A.S.",
        "LOGISTICA ADUANERA J.H. LTDA",
        "LOGISTICA COEX SAS",
        "M.B. INDUSTRIAS Y SERVICIOS SAS",
        "MANUFEX GROUP S.A.S.",
        "SETRACOMEX S.A.S.",
        "TRENDS LOVERS S.A.S.",
        "VEHIMAQ ZONA FRANCA SAS"
    ],
}

def ahora_str():
    colombia_tz = timezone(timedelta(hours=-5))
    return datetime.now(colombia_tz).strftime("%Y-%m-%d %H:%M:%S")

def obtener_ultimo_corte():
    colombia_tz = timezone(timedelta(hours=-5))
    ahora = datetime.now(colombia_tz)
    if ahora.hour >= 18:
        return ahora.replace(hour=18, minute=0, second=0, microsecond=0)
    else:
        ayer = ahora - timedelta(days=1)
        return ayer.replace(hour=18, minute=0, second=0, microsecond=0)

def _migrar_campos(item):
    defaults = {
        "uid": str(uuid.uuid4()),
        "tipo_movimiento": "Ingresar",
        "fecha_hora": ahora_str(),
        "mostrar_en_pantalla": False,
        "fecha_autorizacion": None,
        "fecha_ingreso_zf": None,
        "fecha_salida_zf": None,
        "checklist_ingreso": {},
        "checklist_salida": {},
        "observaciones_vigilante": "",
        "observaciones_conductor": "",
        "telefono": "",
    }
    for k, v in defaults.items():
        item.setdefault(k, v)

def cargar_datos():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                datos = json.load(f)
                for item in datos:
                    _migrar_campos(item)
                return datos
        except Exception:
            return []
    return []

def guardar_datos(datos):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                
                if cfg.get("config_version", 1) < DEFAULT_CONFIG["config_version"]:
                    cfg["config_version"] = DEFAULT_CONFIG["config_version"]
                    cfg["contrasena_admin"] = DEFAULT_CONFIG["contrasena_admin"]
                    cfg["usuario_vigilante"] = DEFAULT_CONFIG["usuario_vigilante"]
                    cfg["contrasena_vigilante"] = DEFAULT_CONFIG["contrasena_vigilante"]
                    cfg["bodegas"] = DEFAULT_CONFIG["bodegas"]
                    guardar_config(cfg)
                    
                return {**DEFAULT_CONFIG, **cfg}
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def guardar_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ==========================================================================
# HELPERS Y EXPORTACIÓN A EXCEL
# ==========================================================================
def fecha_legible(s):
    if not s:
        return "—"
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return s

def parse_fecha(s):
    if not s:
        return None
    try:
        colombia_tz = timezone(timedelta(hours=-5))
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S").replace(tzinfo=colombia_tz)
    except Exception:
        return None

def debe_ocultar_por_tiempo(turno, cfg):
    if not turno.get("mostrar_en_pantalla"):
        return True
    fa = parse_fecha(turno.get("fecha_autorizacion"))
    if fa is None:
        return False
    colombia_tz = timezone(timedelta(hours=-5))
    ahora_col = datetime.now(colombia_tz)
    limite = timedelta(minutes=cfg.get("tiempo_en_pantalla_min", 30))
    return (ahora_col - fa) > limite

def badge_html(estado_raw):
    e = (estado_raw or "").upper()
    if "INGRESAR" in e and "ZONA" not in e and "FINALIZADO" not in e:
        return '<span class="badge badge-ingresar">Autorizado para Ingresar</span>'
    if "SALIR" in e and "ZONA" not in e and "FINALIZADO" not in e:
        return '<span class="badge badge-salir">Autorizado para Salir</span>'
    if "ZONA FRANCA" in e:
        return '<span class="badge badge-zf">En Zona Franca</span>'
    if "FINALIZADO" in e:
        return '<span class="badge badge-finalizado">Finalizado</span>'
    return '<span class="badge badge-espera">En Espera</span>'

def buscar_turno(turnos, uid):
    return next((t for t in turnos if t.get("uid") == uid), None)

def generar_qr_png(url: str):
    if not HAS_QR or not url:
        return None
    try:
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=15,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="#000000", back_color="#FFFFFF")
        buf = BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None

def generar_excel(df, filename="reporte.xlsx"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte')
        ws = writer.sheets['Reporte']
        
        header_fill = PatternFill(start_color="1B4D3E", end_color="1B4D3E", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )
        
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
            
            max_length = len(str(df.columns[col_idx - 1]))
            for val in df.iloc[:, col_idx - 1]:
                max_length = max(max_length, len(str(val)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 3, 40)

        for row_idx in range(2, len(df) + 2):
            for col_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                
        ws.freeze_panes = "A2"
    return buffer.getvalue()

def render_header_institucional():
    c1, c2, c3 = st.columns([1.5, 3, 1.5])
    with c1:
        for fn in ["LOGO ZFS-ZFC.jpeg", "LOGO ZFS-ZFC.jpg", "LOGO ZFS-ZFC.png"]:
            if os.path.exists(fn):
                st.image(fn, use_container_width=True)
                break
    with c2:
        st.markdown("""
        <div style="text-align:center;">
            <div class="inst-title">Sistema de Enturnamiento y Control de Accesos</div>
            <div class="inst-sub" style="font-size: 14px; font-weight: 600; color: #1B4D3E; margin-top: 4px;">
                Ingresos y salidas de Resto del mundo
            </div>
            <div class="inst-sub">ZONA FRANCA DE CÚCUTA · OPERADA POR ZONA FRANCA SANTANDER</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        for fn in ["LOGO SANTANDER.jpg", "LOGO SANTANDER.png", "logo.jpeg"]:
            if os.path.exists(fn):
                st.image(fn, use_container_width=True)
                break

# ==========================================================================
# SESSION STATE
# ==========================================================================
if "historial_turnos" not in st.session_state:
    st.session_state.historial_turnos = []

if "auth_operador" not in st.session_state:
    st.session_state.auth_operador = False
if "auth_vigilante" not in st.session_state:
    st.session_state.auth_vigilante = False
if "mi_turno_id" not in st.session_state:
    st.session_state.mi_turno_id = None

try:
    qp = st.query_params
    modo_qp = (qp.get("modo") or "").strip().lower()
    
    if st.session_state.mi_turno_id is None and "t" in qp:
        st.session_state.mi_turno_id = qp.get("t")
        
    ahora_ts = time.time()
    
    if not st.session_state.auth_operador and "op_ts" in qp:
        try:
            if ahora_ts - float(qp["op_ts"]) < 7200:
                st.session_state.auth_operador = True
        except ValueError:
            pass
            
    if not st.session_state.auth_vigilante and "vig_ts" in qp:
        try:
            if ahora_ts - float(qp["vig_ts"]) < 7200:
                st.session_state.auth_vigilante = True
        except ValueError:
            pass

except Exception:
    qp = {}
    modo_qp = ""

CONFIG = cargar_config()

# ==========================================================================
# MODO: PANTALLA TV
# ==========================================================================
if modo_qp == "pantalla":
    @st.fragment(run_every=4)
    def pantalla_publica():
        turnos = cargar_datos()
        cfg = cargar_config()

        def visible(t):
            if not t.get("mostrar_en_pantalla"):
                return False
            if debe_ocultar_por_tiempo(t, cfg):
                return False
            est = (t.get("estado") or "").upper()
            return ("INGRESAR" in est or "SALIR" in est) and "ZONA" not in est

        autorizados = [t for t in turnos if visible(t)]
        ingresos = [t for t in autorizados if "INGRESAR" in (t.get("estado") or "").upper()]
        salidas = [t for t in autorizados if "SALIR" in (t.get("estado") or "").upper()]

        c1, c2, c3 = st.columns([1, 3, 1])
        with c1:
            for fn in ["LOGO ZFS-ZFC.jpeg", "LOGO ZFS-ZFC.png"]:
                if os.path.exists(fn):
                    st.image(fn, use_container_width=True)
                    break
        with c2:
            st.markdown("""
            <div style="text-align:center; padding: 8px 0;">
                <div style="color:#12402A; font-size:32px; font-weight:900;">
                    CONTROL DE ACCESOS EN TIEMPO REAL
                </div>
                <div style="color:#6B7280; font-size:14px; letter-spacing:3px; margin-top:6px;">
                    ZONA FRANCA DE CÚCUTA · ZONA FRANCA SANTANDER
                </div>
            </div>
            """, unsafe_allow_html=True)
        with c3:
            for fn in ["LOGO SANTANDER.jpg", "LOGO SANTANDER.png"]:
                if os.path.exists(fn):
                    st.image(fn, use_container_width=True)
                    break

        st.markdown("<br>", unsafe_allow_html=True)
        col_i, col_s = st.columns(2)

        with col_i:
            st.markdown('<div class="col-header-ing">AUTORIZADOS PARA INGRESAR</div>', unsafe_allow_html=True)
            if not ingresos:
                st.markdown('<div class="empty-state">No hay vehículos autorizados para ingresar.</div>', unsafe_allow_html=True)
            else:
                for t in ingresos:
                    st.markdown(f"""
                    <div class="card-ingreso">
                        <div class="card-ficho">Ficho {t['id_turno']}</div>
                        <div class="card-placa">{t['placa']}</div>
                        <div class="card-aviso">{t.get('mensaje_admin','') or 'Sin mensaje'}</div>
                    </div>
                    """, unsafe_allow_html=True)

        with col_s:
            st.markdown('<div class="col-header-sal">AUTORIZADOS PARA SALIR</div>', unsafe_allow_html=True)
            if not salidas:
                st.markdown('<div class="empty-state">No hay vehículos autorizados para salir.</div>', unsafe_allow_html=True)
            else:
                for t in salidas:
                    st.markdown(f"""
                    <div class="card-salida">
                        <div class="card-ficho">Ficho {t['id_turno']}</div>
                        <div class="card-placa-sal">{t['placa']}</div>
                        <div class="card-aviso">{t.get('mensaje_admin','') or 'Sin mensaje'}</div>
                    </div>
                    """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(
            f"<div style='text-align:center; color:#6B7280; font-size:12px;'>"
            f"Última actualización: {ahora_str()}</div>",
            unsafe_allow_html=True,
        )

    pantalla_publica()
    st.stop()

# ==========================================================================
# MODO: CONDUCTOR
# ==========================================================================
if modo_qp == "conductor":
    render_header_institucional()
    st.markdown("<h3 style='text-align:center;'>Portal del Conductor</h3>", unsafe_allow_html=True)

    st.markdown("""
    <div style="
        background: linear-gradient(135deg, #E6F0EB 0%, #D1E7DD 100%);
        border: 1px solid #B8D4C5;
        border-left: 5px solid #1B4D3E;
        padding: 16px 20px;
        border-radius: 8px;
        margin: 15px 0 25px 0;
        color: #12402A;
        box-shadow: 0 4px 6px rgba(0,0,0,0.04);
        font-size: 15px;
        font-weight: 600;
        text-align: center;
    ">
        ¡Bienvenido a la Zona Franca de Cúcuta! Si vienes de algún Puerto o de otra Zona Franca, te invitamos a que te enturnes.
    </div>
    """, unsafe_allow_html=True)

    left, center, right = st.columns([1, 2, 1])
    with center:
        turnos = cargar_datos()

        mi_turno = None
        if st.session_state.mi_turno_id:
            mi_turno = buscar_turno(turnos, st.session_state.mi_turno_id)

        if mi_turno and "FINALIZADO" not in (mi_turno.get("estado") or "").upper():

            @st.fragment(run_every=5)
            def tiquet_vivo():
                tv = cargar_datos()
                actual = buscar_turno(tv, st.session_state.mi_turno_id)
                if not actual:
                    st.warning("Tu turno ya no está disponible.")
                    return
                msg = actual.get("mensaje_admin") or "En espera de autorización"

                st.markdown(f"""
                <div class="ticket-card">
                    <div class="ticket-label">Turno Asignado</div>
                    <hr class="ticket-divider">
                    <div class="ticket-label">Placa del Vehículo</div>
                    <div class="ticket-placa">{actual['placa']}</div>
                    <div class="ticket-label" style="margin-top:18px;">Número de Ficho</div>
                    <div class="ticket-turno">{actual['id_turno']}</div>
                    <hr class="ticket-divider">
                    <div class="ticket-info"><b>Destino:</b> {actual.get('bodega','—')}</div>
                    <div class="ticket-info"><b>Movimiento:</b> {actual.get('tipo_movimiento','Ingresar')}</div>
                    <div class="ticket-info" style="font-size:12px; color:#A9BCB2; margin-top:6px;">
                        Registrado: {fecha_legible(actual.get('fecha_hora',''))}
                    </div>
                    <div style="margin-top:16px;">{badge_html(actual.get('estado',''))}</div>
                    <div class="ticket-mensaje">
                        <div class="ticket-label" style="font-size:10px;">Mensaje del Operador</div>
                        <div style="color:#FFD966; font-weight:700; font-size:15px; margin-top:4px;">{msg}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                if actual.get("observaciones_conductor"):
                    st.info(f"**Tus observaciones:** {actual['observaciones_conductor']}")

                if actual.get("fecha_ingreso_zf"):
                    st.markdown(
                        f"<div style='text-align:center; color:#1B4D3E; font-weight:600; "
                        f"padding:10px; background:#E6F0EB; border-radius:6px; margin-top:10px;'>"
                        f"<b>Ingreso confirmado por vigilante:</b> {fecha_legible(actual['fecha_ingreso_zf'])}</div>",
                        unsafe_allow_html=True,
                    )
                if actual.get("fecha_salida_zf"):
                    st.markdown(
                        f"<div style='text-align:center; color:#9B1C1C; font-weight:600; "
                        f"padding:10px; background:#FCE9E9; border-radius:6px; margin-top:10px;'>"
                        f"<b>Salida confirmada por vigilante:</b> {fecha_legible(actual['fecha_salida_zf'])}</div>",
                        unsafe_allow_html=True,
                    )

            tiquet_vivo()

            st.markdown("<br>", unsafe_allow_html=True)
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("Actualizar Estado Manualmente", use_container_width=True):
                    st.rerun()
            with col_b:
                if st.button("Registrar otro vehículo", use_container_width=True):
                    st.session_state.mi_turno_id = None
                    try:
                        st.query_params.clear()
                        st.query_params["modo"] = "conductor"
                    except Exception:
                        pass
                    st.rerun()

        else:
            st.info("📍 Ubicación validada. Puedes continuar con tu registro.")
            st.markdown("Completa tus datos para asignarte un turno.")
            st.caption("Los campos con (*) son obligatorios. La bodega es opcional.")

            with st.form("form_conductor", clear_on_submit=False):
                placa = st.text_input("Placa del Vehículo *", placeholder="Ej: ABC123").upper().strip()
                nombre = st.text_input("Nombre Completo del Conductor *", placeholder="Ej: Juan Pérez").title().strip()
                cedula = st.text_input("Número de Cédula *", placeholder="Ej: 1234567890").strip()
                telefono = st.text_input("Teléfono de Contacto *", placeholder="Ej: 3001234567").strip()
                
                bodega = st.selectbox(
                    "Bodega de Destino (Opcional)",
                    ["Seleccione una bodega..."] + CONFIG.get("bodegas", []),
                )
                
                observaciones_conductor = st.text_area(
                    "Observaciones (Opcional)", 
                    placeholder="Si tienes alguna nota sobre tu vehículo o carga, escríbela aquí..."
                )
                
                submitted = st.form_submit_button("Obtener Turno", use_container_width=True)

                if submitted:
                    errores = []
                    if not placa: errores.append("La placa del vehículo es obligatoria.")
                    if not nombre: errores.append("El nombre del conductor es obligatorio.")
                    if not cedula: errores.append("El número de cédula es obligatorio.")
                    if not telefono: errores.append("El teléfono de contacto es obligatorio.")
                    
                    if cedula and not cedula.isdigit(): errores.append("La cédula debe contener solo números.")
                    if telefono and not telefono.isdigit(): errores.append("El teléfono debe contener solo números.")
                    if telefono and len(telefono) < 7: errores.append("El teléfono debe tener al menos 7 dígitos.")
                    if placa and len(placa) < 5: errores.append("La placa debe tener al menos 5 caracteres.")

                    if errores:
                        st.error("⚠️ Corrige los siguientes errores:\n\n" + "\n".join(f"• {e}" for e in errores))
                    else:
                        actuales = cargar_datos()
                        placa_clean = placa.replace(" ", "").upper()
                        
                        activo = next(
                            (t for t in actuales
                             if t.get("placa", "").replace(" ", "").upper() == placa_clean 
                             and "FINALIZADO" not in (t.get("estado") or "").upper()),
                            None,
                        )
                        if activo:
                            st.warning(f"La placa {placa} ya tiene un turno activo: {activo['id_turno']} del {fecha_legible(activo['fecha_hora'])}")
                            st.session_state.mi_turno_id = activo["uid"]
                            try:
                                st.query_params["t"] = activo["uid"]
                            except Exception:
                                pass
                            st.rerun()
                        else:
                            st.session_state.historial_turnos.append(copy.deepcopy(actuales))
                            
                            ultimo_corte = obtener_ultimo_corte()
                            corte_str = ultimo_corte.strftime("%Y-%m-%d %H:%M:%S")
                            
                            turnos_jornada = [t for t in actuales if t.get("fecha_hora", "") >= corte_str]
                            
                            ids_existentes = [int(t["id_turno"].split("-")[1]) for t in turnos_jornada if t.get("id_turno", "").startswith("T-") and t["id_turno"].split("-")[1].isdigit()]
                            nuevo_num = max(ids_existentes) + 1 if ids_existentes else 1
                            turno_str = f"T-{nuevo_num:03d}"

                            nuevo = {
                                "uid": str(uuid.uuid4()),
                                "id_turno": turno_str,
                                "placa": placa_clean,
                                "nombre": nombre,
                                "cedula": cedula,
                                "telefono": telefono,
                                "bodega": bodega if bodega != "Seleccione una bodega..." else "",
                                "tipo_movimiento": "Ingresar",
                                "estado": "En Espera",
                                "fecha_hora": ahora_str(),
                                "mensaje_admin": "En espera de autorización",
                                "mostrar_en_pantalla": False,
                                "fecha_autorizacion": None,
                                "fecha_ingreso_zf": None,
                                "fecha_salida_zf": None,
                                "checklist_ingreso": {},
                                "checklist_salida": {},
                                "observaciones_vigilante": "",
                                "observaciones_conductor": observaciones_conductor,
                            }
                            actuales.append(nuevo)
                            guardar_datos(actuales)
                            
                            st.session_state.mi_turno_id = nuevo["uid"]
                            try:
                                st.query_params["t"] = nuevo["uid"]
                            except Exception:
                                pass
                            st.rerun()

        st.markdown("---")
        with st.expander("Visualizar estado de mi placa"):
            p_buscar = st.text_input("Ingresa tu placa", key="rec_placa").upper().strip()
            if st.button("Buscar", key="btn_rec"):
                if not p_buscar:
                    st.error("Ingresa una placa.")
                else:
                    tb = cargar_datos()
                    p_buscar_clean = p_buscar.replace(" ", "").upper()
                    encontrado = next(
                        (t for t in reversed(tb)
                         if t.get("placa", "").replace(" ", "").upper() == p_buscar_clean 
                         and "FINALIZADO" not in (t.get("estado") or "").upper()),
                        None,
                    )
                    if encontrado:
                        st.session_state.mi_turno_id = encontrado["uid"]
                        try:
                            st.query_params["t"] = encontrado["uid"]
                        except Exception:
                            pass
                        st.rerun()
                    else:
                        st.error("No se encontró turno activo para esa placa.")

    st.stop()

# ==========================================================================
# MODO: VIGILANTE
# ==========================================================================
if modo_qp == "vigilante":
    render_header_institucional()

    if not st.session_state.auth_vigilante:
        st.markdown("<h3 style='text-align:center;'>Acceso Vigilante</h3>", unsafe_allow_html=True)
        left, center, right = st.columns([1, 1.2, 1])
        with center:
            with st.form("login_vig"):
                u = st.text_input("Usuario", placeholder="Ingresa tu usuario")
                p = st.text_input("Contraseña", type="password", placeholder="Ingresa tu contraseña")
                if st.form_submit_button("Iniciar Sesión", use_container_width=True):
                    if u == CONFIG.get("usuario_vigilante") and p == CONFIG.get("contrasena_vigilante"):
                        st.session_state.auth_vigilante = True
                        try:
                            st.query_params["vig_ts"] = str(time.time())
                        except Exception:
                            pass
                        st.rerun()
                    else:
                        st.error("Credenciales incorrectas.")
        st.stop()

    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        st.markdown("### Portal del Vigilante")
        st.caption(f"Sesión activa · {ahora_str()}")
    with col_h2:
        if st.button("Cerrar Sesión", use_container_width=True, key="logout_vig"):
            st.session_state.auth_vigilante = False
            try:
                if "vig_ts" in st.query_params:
                    del st.query_params["vig_ts"]
            except Exception:
                pass
            st.rerun()

    st.markdown("---")

    tab_control, tab_historial_vig = st.tabs(["Control de Acceso", "Historial"])

    with tab_control:
        @st.fragment(run_every=5)
        def panel_vigilante():
            turnos = cargar_datos()

            pendientes_ing = [
                t for t in turnos
                if "INGRESAR" in (t.get("estado") or "").upper()
                and "ZONA" not in (t.get("estado") or "").upper()
                and "FINALIZADO" not in (t.get("estado") or "").upper()
                and not t.get("fecha_ingreso_zf")
            ]
            pendientes_sal = [
                t for t in turnos
                if "SALIR" in (t.get("estado") or "").upper()
                and "FINALIZADO" not in (t.get("estado") or "").upper()
                and not t.get("fecha_salida_zf")
            ]

            m1, m2, m3 = st.columns(3)
            m1.metric("Pendientes de Ingreso", len(pendientes_ing))
            m2.metric("Pendientes de Salida", len(pendientes_sal))
            m3.metric("Total en Sistema", len(turnos))

            st.markdown("---")

            col_i, col_s = st.columns(2)

            with col_i:
                st.markdown('<div class="col-header-ing">CHECK IN · INGRESOS</div>', unsafe_allow_html=True)
                if not pendientes_ing:
                    st.markdown('<div class="empty-state">Sin vehículos pendientes de ingreso.</div>', unsafe_allow_html=True)
                else:
                    for t in pendientes_ing:
                        with st.expander(f"{t['placa']} · {t['id_turno']}", expanded=False):
                            st.markdown(f"**Conductor:** {t.get('nombre','')} — C.C. {t.get('cedula','')}")
                            st.markdown(f"**Teléfono:** {t.get('telefono','—')}")
                            if t.get("bodega"): st.markdown(f"**Bodega:** {t['bodega']}")
                            st.caption(f"Registrado: {fecha_legible(t.get('fecha_hora',''))}")
                            if t.get("observaciones_conductor"):
                                st.info(f"**Observación del Conductor:** {t['observaciones_conductor']}")
                            if t.get("mensaje_admin"):
                                st.markdown(f"**Mensaje Operador:** {t['mensaje_admin']}")

                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("Ingreso Autorizado", key=f"btn_in_{t['uid']}", use_container_width=True):
                                st.session_state.historial_turnos.append(copy.deepcopy(turnos))
                                hora_conf = ahora_str()
                                todos = cargar_datos()
                                for x in todos:
                                    if x["uid"] == t["uid"]:
                                        x["fecha_ingreso_zf"] = hora_conf
                                        x["estado"] = "EN ZONA FRANCA"
                                        x["mostrar_en_pantalla"] = False
                                        break
                                guardar_datos(todos)
                                st.success(f"Ingreso autorizado para {t['placa']} a las {hora_conf}")
                                st.rerun()

            with col_s:
                st.markdown('<div class="col-header-sal">CHECK OUT · SALIDAS</div>', unsafe_allow_html=True)
                if not pendientes_sal:
                    st.markdown('<div class="empty-state">Sin vehículos pendientes de salida.</div>', unsafe_allow_html=True)
                else:
                    for t in pendientes_sal:
                        with st.expander(f"{t['placa']} · {t['id_turno']}", expanded=False):
                            st.markdown(f"**Conductor:** {t.get('nombre','')} — C.C. {t.get('cedula','')}")
                            st.markdown(f"**Teléfono:** {t.get('telefono','—')}")
                            if t.get("bodega"): st.markdown(f"**Bodega:** {t['bodega']}")
                            if t.get("fecha_ingreso_zf"):
                                st.caption(f"Ingresó: {fecha_legible(t['fecha_ingreso_zf'])}")
                            if t.get("observaciones_conductor"):
                                st.info(f"**Observación del Conductor:** {t['observaciones_conductor']}")
                            if t.get("mensaje_admin"):
                                st.markdown(f"**Mensaje Operador:** {t['mensaje_admin']}")

                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("Salida Autorizada", key=f"btn_out_{t['uid']}", use_container_width=True):
                                st.session_state.historial_turnos.append(copy.deepcopy(turnos))
                                hora_conf = ahora_str()
                                todos = cargar_datos()
                                for x in todos:
                                    if x["uid"] == t["uid"]:
                                        x["fecha_salida_zf"] = hora_conf
                                        x["estado"] = "FINALIZADO"
                                        x["mostrar_en_pantalla"] = False
                                        break
                                guardar_datos(todos)
                                st.success(f"Salida autorizada para {t['placa']} a las {hora_conf}")
                                st.rerun()

        panel_vigilante()

    with tab_historial_vig:
        st.markdown("#### Historial General de Turnos")
        turnos_h = cargar_datos()
        if not turnos_h:
            st.info("No hay turnos registrados.")
        else:
            df = pd.DataFrame(turnos_h)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total", len(df))
            m2.metric("En Espera", len(df[df["estado"].astype(str).str.contains("Espera", case=False, na=False)]))
            m3.metric("Autorizados", len(df[df["estado"].astype(str).str.contains("AUTORIZADO", case=False, na=False)]))
            m4.metric("Finalizados", len(df[df["estado"].astype(str).str.contains("FINALIZADO", case=False, na=False)]))

            st.markdown("---")

            df["fecha_hora"] = df["fecha_hora"].apply(fecha_legible)
            df["fecha_autorizacion"] = df["fecha_autorizacion"].apply(fecha_legible)
            df["fecha_ingreso_zf"] = df["fecha_ingreso_zf"].apply(fecha_legible)
            df["fecha_salida_zf"] = df["fecha_salida_zf"].apply(fecha_legible)

            cols = [
                "id_turno", "placa", "nombre", "cedula", "telefono", "fecha_hora",
                "bodega", "tipo_movimiento", "estado", "mensaje_admin",
                "fecha_autorizacion", "fecha_ingreso_zf", "fecha_salida_zf",
                "observaciones_conductor", "observaciones_vigilante"
            ]
            cols_disp = [c for c in cols if c in df.columns]

            df_v = df[cols_disp].rename(columns={
                "id_turno": "Turno",
                "placa": "Placa",
                "nombre": "Conductor",
                "cedula": "Cédula",
                "telefono": "Teléfono",
                "fecha_hora": "Fecha Registro",
                "bodega": "Bodega",
                "tipo_movimiento": "Movimiento",
                "estado": "Estado",
                "mensaje_admin": "Mensaje",
                "fecha_autorizacion": "Autorizado Operador",
                "fecha_ingreso_zf": "Ingreso Vigilante",
                "fecha_salida_zf": "Salida Vigilante",
                "observaciones_conductor": "Obs. Conductor",
                "observaciones_vigilante": "Obs. Vigilante",
            })

            st.dataframe(df_v, use_container_width=True, hide_index=True, height=450)

            excel_data = generar_excel(df_v)
            st.download_button(
                "Descargar Excel",
                data=excel_data,
                file_name=f"turnos_zfc_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.stop()

# ==========================================================================
# MODO: OPERADOR
# ==========================================================================
if modo_qp == "operador":
    render_header_institucional()

    if not st.session_state.auth_operador:
        st.markdown("<h3 style='text-align:center;'>Acceso Operador</h3>", unsafe_allow_html=True)
        left, center, right = st.columns([1, 1.2, 1])
        with center:
            with st.form("login_op"):
                u = st.text_input("Usuario", placeholder="Ingresa tu usuario")
                p = st.text_input("Contraseña", type="password", placeholder="Ingresa tu contraseña")
                if st.form_submit_button("Iniciar Sesión", use_container_width=True):
                    if u == CONFIG.get("usuario_admin") and p == CONFIG.get("contrasena_admin"):
                        st.session_state.auth_operador = True
                        try:
                            st.query_params["op_ts"] = str(time.time())
                        except Exception:
                            pass
                        st.rerun()
                    else:
                        st.error("Credenciales incorrectas.")
        st.stop()

    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        st.markdown("### Panel del Operador")
        st.caption(f"Sesión activa · {ahora_str()}")
    with col_h2:
        if st.button("Cerrar Sesión", use_container_width=True, key="logout_op"):
            st.session_state.auth_operador = False
            try:
                if "op_ts" in st.query_params:
                    del st.query_params["op_ts"]
            except Exception:
                pass
            st.rerun()

    st.markdown("---")

    tab_gestion, tab_pantalla, tab_qr, tab_config, tab_historial = st.tabs([
        "Gestión de Turnos",
        "Pantalla Pública",
        "Código QR",
        "Configuración",
        "Historial",
    ])

    # --- GESTIÓN ---
    with tab_gestion:
        turnos_g = cargar_datos()
        if not turnos_g:
            st.info("No hay turnos registrados.")
        else:
            if st.session_state.historial_turnos:
                if st.button("Deshacer Última Acción", type="primary", key="btn_undo"):
                    ultimo_estado = st.session_state.historial_turnos.pop()
                    guardar_datos(ultimo_estado)
                    st.success("Se ha deshecho la última acción correctamente.")
                    st.rerun()
            
            filtro = st.selectbox(
                "Filtrar por estado:",
                ["Todos", "En Espera", "Autorizados", "En Zona Franca", "Finalizados"],
                key="filtro_op",
            )

            def filtrar(lst):
                if filtro == "Todos": return lst
                if filtro == "En Espera": return [t for t in lst if "ESPERA" in (t.get("estado") or "").upper()]
                if filtro == "Autorizados": return [t for t in lst if "AUTORIZADO" in (t.get("estado") or "").upper()]
                if filtro == "En Zona Franca": return [t for t in lst if "ZONA FRANCA" in (t.get("estado") or "").upper()]
                if filtro == "Finalizados": return [t for t in lst if "FINALIZADO" in (t.get("estado") or "").upper()]
                return lst

            lista = filtrar(turnos_g)

            if lista:
                opciones_dict = {
                    f"{t['id_turno']} ({fecha_legible(t.get('fecha_hora',''))}) — {t['placa']} — {t['estado']}": t.get("uid", "")
                    for t in lista
                }
                
                sel = st.selectbox("Selecciona un turno para gestionar:", list(opciones_dict.keys()), key="sel_op")
                id_sel = opciones_dict[sel]
                turno_sel = buscar_turno(turnos_g, id_sel)

                if turno_sel:
                    with st.form(f"form_g_{id_sel}"):
                        st.markdown("##### Editar Datos del Conductor")
                        
                        nuevo_tiquet = st.text_input("Número de Ficho (Tiquet)", value=turno_sel.get('id_turno', '')).upper().strip()
                        
                        nueva_placa = st.text_input("Placa del Vehículo", value=turno_sel.get('placa', '')).upper().strip()
                        nuevo_nombre = st.text_input("Nombre del Conductor", value=turno_sel.get('nombre', '')).title().strip()
                        nueva_cedula = st.text_input("Cédula", value=turno_sel.get('cedula', '')).strip()
                        nuevo_telefono = st.text_input("Teléfono", value=turno_sel.get('telefono', '')).strip()
                        
                        lista_bodegas = ["Seleccione una bodega..."] + CONFIG.get("bodegas", [])
                        bodega_actual = turno_sel.get("bodega", "")
                        idx_bodega = lista_bodegas.index(bodega_actual) if bodega_actual in lista_bodegas else 0
                        nueva_bodega = st.selectbox("Bodega de Destino", lista_bodegas, index=idx_bodega)
                        
                        st.text_area(
                            "Observaciones del Conductor (Solo lectura)", 
                            value=turno_sel.get('observaciones_conductor', ''), 
                            disabled=True,
                            help="Este campo solo puede ser modificado por el conductor al momento de registrarse."
                        )
                        
                        st.caption(f"**Registrado originalmente:** {fecha_legible(turno_sel.get('fecha_hora',''))}")

                        st.markdown("---")
                        st.markdown("##### Estado y Gestión")
                        
                        nuevo_tipo = st.selectbox(
                            "Tipo de movimiento",
                            ["Ingresar", "Salir"],
                            index=0 if turno_sel.get("tipo_movimiento") == "Ingresar" else 1,
                        )

                        estados_posibles = [
                            "En Espera",
                            "AUTORIZADO PARA INGRESAR",
                            "AUTORIZADO PARA SALIR",
                        ]
                        estado_actual = turno_sel.get("estado", "En Espera")
                        idx_estado = estados_posibles.index(estado_actual) if estado_actual in estados_posibles else 0

                        nuevo_estado = st.selectbox("Estado", estados_posibles, index=idx_estado)

                        mensaje = st.text_input(
                            "Mensaje para el conductor (ej: Muelle 3)",
                            value=turno_sel.get("mensaje_admin", ""),
                        )

                        mostrar = st.checkbox(
                            "Mostrar en pantalla pública",
                            value=bool(turno_sel.get("mostrar_en_pantalla", False)),
                        )

                        st.markdown("---")
                        eliminar = st.checkbox("Eliminar este turno permanentemente")
                        
                        pass_confirm = ""
                        if eliminar:
                            pass_confirm = st.text_input("Ingrese contraseña de Operador para confirmar eliminación", type="password")

                        if st.form_submit_button("Guardar Cambios", use_container_width=True):
                            if eliminar:
                                if pass_confirm != CONFIG.get("contrasena_admin"):
                                    st.error("❌ Contraseña incorrecta. No se puede eliminar el turno.")
                                else:
                                    st.session_state.historial_turnos.append(copy.deepcopy(turnos_g))
                                    turnos_g = [t for t in turnos_g if t.get("uid") != id_sel]
                                    guardar_datos(turnos_g)
                                    st.success(f"Turno {turno_sel['id_turno']} eliminado.")
                                    st.rerun()
                            else:
                                nueva_placa_clean = nueva_placa.replace(" ", "").upper()
                                placa_duplicada = any(
                                    t.get("placa", "").replace(" ", "").upper() == nueva_placa_clean 
                                    and t.get("uid") != id_sel 
                                    and "FINALIZADO" not in (t.get("estado") or "").upper()
                                    for t in turnos_g
                                )
                                
                                ultimo_corte = obtener_ultimo_corte()
                                corte_str = ultimo_corte.strftime("%Y-%m-%d %H:%M:%S")
                                tiquet_duplicado = any(
                                    t.get("id_turno") == nuevo_tiquet 
                                    and t.get("fecha_hora", "") >= corte_str
                                    and t.get("uid") != id_sel 
                                    for t in turnos_g
                                )

                                if placa_duplicada:
                                    st.error(f"❌ La placa {nueva_placa} ya está registrada en otro turno activo.")
                                elif tiquet_duplicado:
                                    st.error(f"❌ El Ficho {nuevo_tiquet} ya está en uso en la jornada actual. Por favor asigne uno diferente.")
                                elif not nuevo_tiquet:
                                    st.error("❌ El número de Ficho no puede estar vacío.")
                                else:
                                    st.session_state.historial_turnos.append(copy.deepcopy(turnos_g))
                                    for t in turnos_g:
                                        if t.get("uid") == id_sel:
                                            antes = (t.get("estado") or "").upper()
                                            
                                            t["id_turno"] = nuevo_tiquet
                                            t["placa"] = nueva_placa_clean
                                            t["nombre"] = nuevo_nombre
                                            t["cedula"] = nueva_cedula
                                            t["telefono"] = nuevo_telefono
                                            t["bodega"] = nueva_bodega if nueva_bodega != "Seleccione una bodega..." else ""
                                            
                                            t["tipo_movimiento"] = nuevo_tipo
                                            t["estado"] = nuevo_estado
                                            t["mensaje_admin"] = mensaje
                                            t["mostrar_en_pantalla"] = mostrar
                                            
                                            if "AUTORIZADO" in nuevo_estado.upper() and "AUTORIZADO" not in antes:
                                                t["fecha_autorizacion"] = ahora_str()
                                            break
                                    guardar_datos(turnos_g)
                                    st.success("Cambios guardados exitosamente.")
                                    st.rerun()
            else:
                st.info("No hay turnos que coincidan con el filtro.")

    # --- PANTALLA ---
    with tab_pantalla:
        st.markdown("#### Vista previa de Pantalla Pública")
        st.caption(
            f"Los vehículos autorizados se muestran por "
            f"{CONFIG.get('tiempo_en_pantalla_min', 30)} minutos y luego se ocultan automáticamente."
        )

        turnos_p = cargar_datos()

        def visible_p(t):
            if not t.get("mostrar_en_pantalla"):
                return False
            if debe_ocultar_por_tiempo(t, CONFIG):
                return False
            est = (t.get("estado") or "").upper()
            return ("INGRESAR" in est or "SALIR" in est) and "ZONA" not in est

        aut = [t for t in turnos_p if visible_p(t)]
        ing = [t for t in aut if "INGRESAR" in (t.get("estado") or "").upper()]
        sal = [t for t in aut if "SALIR" in (t.get("estado") or "").upper()]

        col_i, col_s = st.columns(2)
        with col_i:
            st.markdown('<div class="col-header-ing">AUTORIZADOS PARA INGRESAR</div>', unsafe_allow_html=True)
            if not ing:
                st.markdown('<div class="empty-state">Sin vehículos visibles.</div>', unsafe_allow_html=True)
            else:
                for t in ing:
                    st.markdown(f"""
                    <div class="card-ingreso">
                        <div class="card-ficho">Ficho {t['id_turno']}</div>
                        <div class="card-placa">{t['placa']}</div>
                        <div class="card-aviso">{t.get('mensaje_admin','') or '—'}</div>
                    </div>
                    """, unsafe_allow_html=True)
        with col_s:
            st.markdown('<div class="col-header-sal">AUTORIZADOS PARA SALIR</div>', unsafe_allow_html=True)
            if not sal:
                st.markdown('<div class="empty-state">Sin vehículos visibles.</div>', unsafe_allow_html=True)
            else:
                for t in sal:
                    st.markdown(f"""
                    <div class="card-salida">
                        <div class="card-ficho">Ficho {t['id_turno']}</div>
                        <div class="card-placa-sal">{t['placa']}</div>
                        <div class="card-aviso">{t.get('mensaje_admin','') or '—'}</div>
                    </div>
                    """, unsafe_allow_html=True)

    # --- QR ---
    with tab_qr:
        st.markdown("#### Centro de Códigos QR")
        st.caption(
            "Configura la URL base de tu aplicación y descarga los códigos QR "
            "que corresponden a cada portal."
        )

        st.markdown("##### Paso 1 · Configura la URL base")
        base_url_actual = CONFIG.get("url_base", "").strip().rstrip("/")

        with st.form("form_url_base_qr"):
            base_url_in = st.text_input(
                "URL base de la aplicación",
                value=base_url_actual,
                placeholder="https://tu-app.streamlit.app",
            )
            if st.form_submit_button("Guardar URL base", use_container_width=True):
                nuevo_cfg = cargar_config()
                nuevo_cfg["url_base"] = base_url_in.strip().rstrip("/")
                guardar_config(nuevo_cfg)
                st.success("URL base actualizada.")
                st.rerun()

        base_url = CONFIG.get("url_base", "").strip().rstrip("/")
        st.markdown("---")

        if not base_url:
            st.warning("Configura primero la URL base (paso 1).")
        elif not HAS_QR:
            st.error("La librería **qrcode[pil]** no está instalada. Agrégala a `requirements.txt`.")
        else:
            st.markdown("##### Paso 2 · Descarga e imprime los códigos QR")

            portales = [
                {
                    "clave": "conductor",
                    "titulo": "CONDUCTOR",
                    "subtitulo": "Punto de enturnamiento",
                    "descripcion": "El conductor escanea este QR y accede al formulario de registro.",
                    "color_hex": "#1B4D3E",
                    "archivo_png": "qr_conductor_zfc.png",
                },
                {
                    "clave": "vigilante",
                    "titulo": "VIGILANTE",
                    "subtitulo": "Punto de control / garita",
                    "descripcion": "El vigilante escanea este QR y accede a su portal de check-in/out.",
                    "color_hex": "#0D4A56",
                    "archivo_png": "qr_vigilante_zfc.png",
                },
                {
                    "clave": "pantalla",
                    "titulo": "PANTALLA TV",
                    "subtitulo": "Televisor informativo",
                    "descripcion": "Abre este enlace en las pantallas para mostrar vehículos autorizados.",
                    "color_hex": "#12402A",
                    "archivo_png": "qr_pantalla_zfc.png",
                },
            ]

            cols_qr = st.columns(3)

            for idx, portal in enumerate(portales):
                url_portal = f"{base_url}/?modo={portal['clave']}"
                qr_bytes = generar_qr_png(url_portal)

                with cols_qr[idx]:
                    st.markdown(f"""
                    <div style="
                        background:#FFFFFF;
                        border: 1px solid #E1E8E5;
                        border-top: 4px solid {portal['color_hex']};
                        border-radius: 10px;
                        padding: 18px 16px 8px 16px;
                        text-align: center;
                        margin-bottom: 12px;
                    ">
                        <div style="color:{portal['color_hex']}; font-weight:700;
                                    font-size:14px; letter-spacing:2px;">
                            {portal['titulo']}
                        </div>
                        <div style="color:#6B7280; font-size:11px; margin-top:2px;">
                            {portal['subtitulo']}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    if qr_bytes:
                        st.image(qr_bytes, use_container_width=True)
                        st.markdown("<div style='text-align:center; margin-top:-10px; margin-bottom:10px;'><b>O ingresa a este enlace:</b></div>", unsafe_allow_html=True)
                        st.code(url_portal, language=None)
                        st.caption(portal["descripcion"])
                        st.download_button(
                            f"Descargar QR · {portal['titulo'].title()}",
                            data=qr_bytes,
                            file_name=portal["archivo_png"],
                            mime="image/png",
                            use_container_width=True,
                            key=f"dl_qr_{portal['clave']}",
                        )
                    else:
                        st.error("No se pudo generar el QR.")

    # --- CONFIG ---
    with tab_config:
        st.markdown("#### Configuración del Sistema")

        with st.form("form_cfg"):
            st.markdown("**Tiempo en pantalla**")
            tiempo = st.number_input(
                "Minutos que un vehículo permanece visible tras su autorización",
                min_value=1, max_value=240,
                value=int(CONFIG.get("tiempo_en_pantalla_min", 30)),
            )

            st.markdown("---")
            st.markdown("**Rango de geolocalización**")
            radio_val = st.number_input(
                "Radio máximo permitido en kilómetros desde la Zona Franca",
                min_value=0.5, max_value=200.0,
                value=float(CONFIG.get("radio_geofence_km", 3.0)),
                step=0.5,
            )

            c_geo1, c_geo2 = st.columns(2)
            with c_geo1:
                zfc_lat_in = st.number_input(
                    "Latitud Zona Franca",
                    value=float(CONFIG.get("zfc_lat", 7.90750)),
                    format="%.5f",
                )
            with c_geo2:
                zfc_lon_in = st.number_input(
                    "Longitud Zona Franca",
                    value=float(CONFIG.get("zfc_lon", -72.50472)),
                    format="%.5f",
                )
            st.caption("Coordenadas de la Zona Franca de Cúcuta. Verifica en Google Maps si es necesario.")

            st.markdown("---")
            st.markdown("**URL base de la aplicación**")
            url_base_in = st.text_input(
                "URL base (ej: https://tu-app.streamlit.app)",
                value=CONFIG.get("url_base", ""),
            )

            st.markdown("---")
            st.markdown("**Credenciales Operador**")
            u_admin = st.text_input("Usuario Admin", value=CONFIG.get("usuario_admin", ""))
            p_admin = st.text_input("Contraseña Admin", value=CONFIG.get("contrasena_admin", ""), type="password")

            st.markdown("---")
            st.markdown("**Credenciales Vigilante**")
            u_vig = st.text_input("Usuario Vigilante", value=CONFIG.get("usuario_vigilante", ""))
            p_vig = st.text_input("Contraseña Vigilante", value=CONFIG.get("contrasena_vigilante", ""), type="password")

            st.markdown("---")
            st.markdown("**Bodegas disponibles** (una por línea)")
            bodegas_txt = st.text_area(
                "Lista de bodegas",
                value="\n".join(CONFIG.get("bodegas", [])),
                height=250,
            )

            if st.form_submit_button("Guardar Configuración", use_container_width=True):
                cfg = {
                    "config_version": DEFAULT_CONFIG["config_version"],
                    "tiempo_en_pantalla_min": int(tiempo),
                    "radio_geofence_km": float(radio_val),
                    "zfc_lat": float(zfc_lat_in),
                    "zfc_lon": float(zfc_lon_in),
                    "usuario_admin": u_admin,
                    "contrasena_admin": p_admin,
                    "usuario_vigilante": u_vig,
                    "contrasena_vigilante": p_vig,
                    "url_base": url_base_in.strip(),
                    "bodegas": [l.strip() for l in bodegas_txt.split("\n") if l.strip()],
                }
                guardar_config(cfg)
                st.success("Configuración guardada.")
                st.rerun()

        st.markdown("---")
        st.markdown("#### Limpieza de Jornada")
        st.caption(
            "Opción para reiniciar la jornada laboral actual y dejar todo limpio para nuevos registros."
        )

        turnos_actuales = cargar_datos()
        total_registros = len(turnos_actuales)

        c_met1, c_met2 = st.columns(2)
        c_met1.metric("Total registros en historial", total_registros)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("##### Limpiar turnos manualmente")
        st.caption(
            "Esta acción borrará TODOS los turnos del historial. Si deseas conservar la información, descarga el respaldo antes de hacer clic."
        )
        if st.button(
            "Limpiar todo el historial",
            use_container_width=True,
            key="clean_todos_hoy",
        ):
            st.session_state.historial_turnos.append(copy.deepcopy(turnos_actuales))
            guardar_datos([])
            st.success("Se eliminaron todos los turnos. Sistema limpio.")
            st.rerun()

        st.markdown("---")
        st.markdown("#### Exportar antes de limpiar")
        st.caption(
            "Recomendación: descarga el historial en Excel antes de limpiar, "
            "para conservar evidencia de la jornada."
        )
        if total_registros > 0:
            df_clean = pd.DataFrame(turnos_actuales)
            
            df_clean["fecha_hora"] = df_clean["fecha_hora"].apply(fecha_legible)
            df_clean["fecha_autorizacion"] = df_clean["fecha_autorizacion"].apply(fecha_legible)
            df_clean["fecha_ingreso_zf"] = df_clean["fecha_ingreso_zf"].apply(fecha_legible)
            df_clean["fecha_salida_zf"] = df_clean["fecha_salida_zf"].apply(fecha_legible)
            
            csv_clean = generar_excel(df_clean)
            st.download_button(
                "Descargar respaldo del historial completo (Excel)",
                data=csv_clean,
                file_name=f"respaldo_turnos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_respaldo",
            )

    # --- HISTORIAL ---
    with tab_historial:
        st.markdown("#### Historial General de Turnos")
        turnos_h = cargar_datos()
        if not turnos_h:
            st.info("No hay turnos registrados.")
        else:
            df = pd.DataFrame(turnos_h)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total", len(df))
            m2.metric("En Espera", len(df[df["estado"].astype(str).str.contains("Espera", case=False, na=False)]))
            m3.metric("Autorizados", len(df[df["estado"].astype(str).str.contains("AUTORIZADO", case=False, na=False)]))
            m4.metric("Finalizados", len(df[df["estado"].astype(str).str.contains("FINALIZADO", case=False, na=False)]))

            st.markdown("---")

            df["fecha_hora"] = df["fecha_hora"].apply(fecha_legible)
            df["fecha_autorizacion"] = df["fecha_autorizacion"].apply(fecha_legible)
            df["fecha_ingreso_zf"] = df["fecha_ingreso_zf"].apply(fecha_legible)
            df["fecha_salida_zf"] = df["fecha_salida_zf"].apply(fecha_legible)

            cols = [
                "id_turno", "placa", "nombre", "cedula", "telefono", "fecha_hora",
                "bodega", "tipo_movimiento", "estado", "mensaje_admin",
                "fecha_autorizacion", "fecha_ingreso_zf", "fecha_salida_zf",
                "observaciones_conductor", "observaciones_vigilante"
            ]
            cols_disp = [c for c in cols if c in df.columns]

            df_v = df[cols_disp].rename(columns={
                "id_turno": "Turno",
                "placa": "Placa",
                "nombre": "Conductor",
                "cedula": "Cédula",
                "telefono": "Teléfono",
                "fecha_hora": "Fecha Registro",
                "bodega": "Bodega",
                "tipo_movimiento": "Movimiento",
                "estado": "Estado",
                "mensaje_admin": "Mensaje",
                "fecha_autorizacion": "Autorizado Operador",
                "fecha_ingreso_zf": "Ingreso Vigilante",
                "fecha_salida_zf": "Salida Vigilante",
                "observaciones_conductor": "Obs. Conductor",
                "observaciones_vigilante": "Obs. Vigilante",
            })

            st.dataframe(df_v, use_container_width=True, hide_index=True, height=450)

            excel_data = generar_excel(df_v)
            st.download_button(
                "Descargar Excel",
                data=excel_data,
                file_name=f"turnos_zfc_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.stop()

# ==========================================================================
# HOME / SELECTOR DE PORTALES
# ==========================================================================
render_header_institucional()

st.markdown("<h3 style='text-align:center; margin-top:10px;'>Selecciona tu Portal de Acceso</h3>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown("""
    <div class="role-card">
        <div class="role-titulo">CONDUCTOR</div>
        <div class="role-desc">Regístrate, obtén tu turno y consulta tu estado desde tu celular.</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Ingresar como Conductor", use_container_width=True, key="home_btn_cond"):
        st.query_params["modo"] = "conductor"
        st.rerun()

with c2:
    st.markdown("""
    <div class="role-card">
        <div class="role-titulo">VIGILANTE</div>
        <div class="role-desc">Autoriza el ingreso y salida de vehículos con un solo clic.</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Ingresar como Vigilante", use_container_width=True, key="home_btn_vig"):
        st.query_params["modo"] = "vigilante"
        st.rerun()

with c3:
    st.markdown("""
    <div class="role-card">
        <div class="role-titulo">OPERADOR</div>
        <div class="role-desc">Autoriza turnos, gestiona la pantalla y administra el sistema.</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Ingresar como Operador", use_container_width=True, key="home_btn_op"):
        st.query_params["modo"] = "operador"
        st.rerun()

with c4:
    st.markdown("""
    <div class="role-card">
        <div class="role-titulo">PANTALLA TV</div>
        <div class="role-desc">Vista pública para televisores. Solo lectura, sin interacción.</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Abrir Pantalla Pública", use_container_width=True, key="home_btn_pant"):
        st.query_params["modo"] = "pantalla"
        st.rerun()

st.markdown("---")
st.caption("Zona Franca de Cúcuta · Operada por Zona Franca Santander · Sistema de Automatización Aduanera")
