"""
Reporte Operativo LiveOps - PedidosYa
Streamlit app para generación de reportes operacionales desde Excel de Tableau.
Persistencia SQLite + exportación de BD.
"""

import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import json
from datetime import date, datetime
from pathlib import Path

# ─── CONFIG ───────────────────────────────────────────────────────────────────
DB_PATH = "reportes_operativos.db"
HORARIOS = ["Breakfast", "Lunch", "Afternoon", "Dinner"]

HORARIO_EMOJI = {
    "Breakfast": "🌅",
    "Lunch":     "☀️",
    "Afternoon": "🌤️",
    "Dinner":    "🌙",
}

# Columnas esperadas del Excel (C2:AH38 → fila 2 = encabezados)
# Mapeamos nombres de columna a claves internas normalizadas
COL_MAP = {
    # Ajusta estos nombres según los encabezados reales de tu Excel de Tableau
    "Region":              "region",
    "Zona":                "region",
    "City":                "region",
    "DT":                  "dt",
    "Delivery Time":       "dt",
    "% Late Orders":       "pct_late",
    "Late Orders":         "pct_late",
    "% FR":                "pct_fr",
    "FR":                  "pct_fr",
    "UTR":                 "utr",
    "At Vendor Time":      "at_vendor",
    "Rider Accepting Time":"rider_accepting",
    "Hold Back Time":      "hold_back",
    "Responsable":         "responsable",
    "Responsible":         "responsable",
}

THRESHOLD_DT = 33.0  # umbral de gestión en minutos

# Mapeo zona → responsable(s). Coincidencia por substring (case-insensitive).
RESPONSABLES = {
    # RM
    "rm":              "@Manuel Montes de Oca Rivero | @Andrés Muñoz Mera | @Nicolás Paredes",
    "santiago":        "@Manuel Montes de Oca Rivero | @Andrés Muñoz Mera | @Nicolás Paredes",
    "metropolitana":   "@Manuel Montes de Oca Rivero | @Andrés Muñoz Mera | @Nicolás Paredes",
    # Cluster A - Matías Morales
    "concepci":        "@Matías Morales",
    "la serena":       "@Matías Morales",
    "rancagua":        "@Matías Morales",
    "puerto montt":    "@Matías Morales",
    # Cluster A - Esteban Jaña
    "viña":            "@Esteban Jaña",
    "vina":             "@Esteban Jaña",
    "antofagasta":     "@Esteban Jaña",
    "talca":           "@Esteban Jaña",
    # Cluster B - Matías Díaz
    "arica":           "@Matías Díaz",
    "iquique":         "@Matías Díaz",
    "calama":          "@Matías Díaz",
    "curic":           "@Matías Díaz",
    "temuco":          "@Matías Díaz",
    "punta arenas":    "@Matías Díaz",
    # Cluster B - Constanza Díaz
    "chill":           "@Constanza Díaz",
    "osorno":          "@Constanza Díaz",
    "valdivia":        "@Constanza Díaz",
    "copiap":          "@Constanza Díaz",
    "los ángeles":     "@Constanza Díaz",
    "los angeles":     "@Constanza Díaz",
    "quillota":        "@Constanza Díaz",
    # Cluster C - Jonathan Matus
    "san felipe":      "@Jonathan Matus",
    "los andes":       "@Jonathan Matus",
    "linares":         "@Jonathan Matus",
    "melipilla":       "@Jonathan Matus",
    "ovalle":          "@Jonathan Matus",
    "san fernando":    "@Jonathan Matus",
    "coyhaique":       "@Jonathan Matus",
    "santa cruz":      "@Jonathan Matus",
    "vallenar":        "@Jonathan Matus",
    "villarrica":      "@Jonathan Matus",
    "castro":          "@Jonathan Matus",
    "san antonio":     "@Jonathan Matus",
    "puerto varas":    "@Jonathan Matus",
    "maitencillo":     "@Jonathan Matus",
    "puc":             "@Jonathan Matus",
}

def get_responsable(region: str) -> str:
    """Retorna el responsable de una región buscando por substring."""
    reg_lower = str(region).strip().lower()
    for key, val in RESPONSABLES.items():
        if key in reg_lower:
            return val
    return ""


# ─── DB ───────────────────────────────────────────────────────────────────────

def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS reportes (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha       TEXT NOT NULL,
            horario     TEXT NOT NULL,
            mensaje     TEXT NOT NULL,
            datos_json  TEXT,
            creado_en   TEXT DEFAULT (datetime('now','localtime'))
        )
    """)
    con.commit()
    con.close()


def guardar_reporte(fecha: str, horario: str, mensaje: str, datos: dict):
    con = sqlite3.connect(DB_PATH)
    con.execute(
        "INSERT INTO reportes (fecha, horario, mensaje, datos_json) VALUES (?,?,?,?)",
        (fecha, horario, mensaje, json.dumps(datos, ensure_ascii=False))
    )
    con.commit()
    con.close()


def listar_reportes() -> pd.DataFrame:
    con = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT id, fecha, horario, creado_en FROM reportes ORDER BY id DESC", con)
    con.close()
    return df


def obtener_reporte(reporte_id: int) -> dict:
    con = sqlite3.connect(DB_PATH)
    row = con.execute(
        "SELECT fecha, horario, mensaje, datos_json, creado_en FROM reportes WHERE id=?",
        (reporte_id,)
    ).fetchone()
    con.close()
    if row:
        return {"fecha": row[0], "horario": row[1], "mensaje": row[2],
                "datos": json.loads(row[3] or "{}"), "creado_en": row[4]}
    return {}


def eliminar_reporte(reporte_id: int):
    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM reportes WHERE id=?", (reporte_id,))
    con.commit()
    con.close()


def exportar_db() -> bytes:
    with open(DB_PATH, "rb") as f:
        return f.read()


# ─── EXCEL PARSING ────────────────────────────────────────────────────────────

# Claves internas reservadas por columna fija (no deben ser tomadas por COL_MAP)
_SENTINEL_RESERVED = {
    "__pct_fr_fixed__":          "pct_fr",
    "__dt_fixed__":              "dt",
    "__rider_accepting_fixed__": "rider_accepting",
}

def normalizar_columnas(df: pd.DataFrame, skip_sentinels: list = None) -> pd.DataFrame:
    """
    Renombra columnas del Excel al esquema interno via COL_MAP.
    - Solo la PRIMERA columna que coincide toma el nombre interno.
    - Columnas sin mapeo o duplicadas se nombran _col_N y luego se eliminan.
    - skip_sentinels: lista de nombres de columna ya asignados por posicion fija;
      se preservan intactos y sus claves internas quedan bloqueadas en COL_MAP.
    """
    skip_sentinels = skip_sentinels or []

    # Bloquear claves internas que ya estan cubiertas por sentinels
    claves_asignadas: set = {
        _SENTINEL_RESERVED[s] for s in skip_sentinels if s in _SENTINEL_RESERVED
    }

    nuevos_nombres = []
    contador_sin_clave = 0

    for col in df.columns:
        if col in skip_sentinels:
            nuevos_nombres.append(col)
            continue

        col_str = str(col).strip()
        asignado = False
        for key, val in COL_MAP.items():
            if key.lower() in col_str.lower() and val not in claves_asignadas:
                nuevos_nombres.append(val)
                claves_asignadas.add(val)
                asignado = True
                break
        if not asignado:
            nuevos_nombres.append(f"_col_{contador_sin_clave}")
            contador_sin_clave += 1

    df.columns = nuevos_nombres
    df = df[[c for c in df.columns if not c.startswith("_col_")]]
    return df


def leer_excel(uploaded_file) -> pd.DataFrame:
    """
    Lee C2:AH38 del primer sheet.
    - Fila 2 = encabezados.
    - Columna J (indice absoluto 9, base-1) = Fail Rate (pct_fr), asignada por posicion fija.
    - El resto de columnas se normalizan por nombre via COL_MAP.
    """
    # Leer hoja completa con encabezados para tener acceso a columna J por nombre real
    df_full = pd.read_excel(
        uploaded_file,
        header=1,       # fila 2 del Excel (0-indexed: fila index 1) = encabezados
        usecols="C:AH",
        nrows=36,       # filas 3-38 = 36 filas de datos
    )
    df_full = df_full.dropna(how="all")

    # Posiciones fijas dentro del rango C:AH (C=0, D=1, ..., J=7, K=8, L=9)
    # Posiciones fijas dentro del rango C:AH (C=0, D=1, ... J=7, L=9, ... AG=30)
    COL_IDX_FR            = 7   # Columna J  = Fail Rate
    COL_IDX_DT            = 9   # Columna L  = Delivery Time (DT)
    COL_IDX_RIDER         = 29  # Columna AF = Rider Accepting Time
    cols = list(df_full.columns)

    SENTINEL_FR    = "__pct_fr_fixed__"
    SENTINEL_DT    = "__dt_fixed__"
    SENTINEL_RIDER = "__rider_accepting_fixed__"

    if COL_IDX_FR    < len(cols): cols[COL_IDX_FR]    = SENTINEL_FR
    if COL_IDX_DT    < len(cols): cols[COL_IDX_DT]    = SENTINEL_DT
    if COL_IDX_RIDER < len(cols): cols[COL_IDX_RIDER]  = SENTINEL_RIDER
    df_full.columns = cols

    # Normalizar el resto por nombre, preservando sentinels
    df_full = normalizar_columnas(df_full, skip_sentinels=[SENTINEL_FR, SENTINEL_DT, SENTINEL_RIDER])

    # Renombrar sentinels a claves internas
    rename_map = {}
    if SENTINEL_FR    in df_full.columns: rename_map[SENTINEL_FR]    = "pct_fr"
    if SENTINEL_DT    in df_full.columns: rename_map[SENTINEL_DT]    = "dt"
    if SENTINEL_RIDER in df_full.columns: rename_map[SENTINEL_RIDER] = "rider_accepting"
    if rename_map:
        df_full = df_full.rename(columns=rename_map)

    return df_full


def fmt_num(val, decimals=1, suffix="") -> str:
    try:
        return f"{float(val):.{decimals}f}{suffix}"
    except (TypeError, ValueError):
        return str(val) if val else "N/D"


def fmt_pct(val) -> str:
    try:
        v = float(val)
        if v <= 1.0:
            v *= 100
        return f"{v:.1f}%"
    except (TypeError, ValueError):
        return "N/D"


# ─── DIAGNÓSTICO AUTOMÁTICO ───────────────────────────────────────────────────

def diagnosticar(row: dict, es_nacional: bool = False) -> str:
    """Genera diagnóstico automático basado en los KPIs."""
    dt         = float(row.get("dt", 0) or 0)
    pct_late   = float(row.get("pct_late", 0) or 0)
    at_vendor  = float(row.get("at_vendor", 0) or 0)
    rider_acc  = float(row.get("rider_accepting", 0) or 0)
    hold_back  = float(row.get("hold_back", 0) or 0)
    utr        = float(row.get("utr", 0) or 0)

    if pct_late <= 1.0:
        pct_late *= 100

    frases = []

    # Estado general
    if dt >= THRESHOLD_DT:
        frases.append(f"Supera el umbral de gestión de {THRESHOLD_DT:.0f} min.")
    elif dt >= 30:
        frases.append(f"En alerta preventiva (bajo {THRESHOLD_DT:.0f} min).")
    else:
        frases.append("Opera en rangos saludables.")

    # Cuellos de botella
    cuellos = []
    if hold_back >= 15:
        cuellos.append(f"Hold Back Time elevado ({hold_back:.1f} min) como principal retención pre-despacho")
    if at_vendor >= 10:
        cuellos.append(f"latencia en Tienda ({at_vendor:.1f} min)")
    if rider_acc >= 10:
        cuellos.append(f"tiempo de respuesta de flota alto ({rider_acc:.1f} min)")

    if cuellos:
        frases.append("Se identifican: " + ", ".join(cuellos) + ".")

    # UTR
    if utr < 1.2:
        frases.append(f"UTR bajo ({utr:.2f}) sugiere exceso de flota para el volumen actual.")
    elif utr > 2.0:
        frases.append(f"UTR elevado ({utr:.2f}) indica alta productividad de flota.")

    # Late orders
    if pct_late >= 25:
        frases.append(f"Alta proporción de órdenes tardías ({pct_late:.1f}%), impacto directo en experiencia.")
    elif pct_late >= 15:
        frases.append(f"Órdenes tardías ({pct_late:.1f}%) sobre el umbral objetivo.")

    return " ".join(frases) if frases else "Sin datos suficientes para diagnóstico automático."


# ─── GENERADOR DE MENSAJE ─────────────────────────────────────────────────────

def generar_mensaje(df: pd.DataFrame, fecha: str, horario: str) -> tuple[str, dict]:
    """
    Genera el mensaje de reporte Slack/WhatsApp y retorna (mensaje, datos_dict).
    Asume que el DataFrame tiene columna 'region' y columnas de KPI.
    """
    emoji_h = HORARIO_EMOJI.get(horario, "📊")
    lineas = []
    datos_guardados = {"fecha": fecha, "horario": horario, "regiones": []}

    # ── Encabezado ──
    lineas.append(f"📊 *Reporte Operativo: Status {horario}*")
    lineas.append(f"📅 Fecha: {fecha}")
    lineas.append("")

    # ── Nacional: SIEMPRE la ultima fila (fila 38 del Excel) ──
    idx_nacional = len(df) - 1
    nacional = df.iloc[idx_nacional] if len(df) > 0 else None

    if nacional is not None:
        r = nacional.to_dict()
        lineas.append("🇨🇱 *TOTAL CL (Resumen Nacional)*")
        lineas.append(
            f"Dato: DT {fmt_num(r.get('dt'))} mins | "
            f"% Late Orders {fmt_pct(r.get('pct_late'))} | "
            f"% FR {fmt_pct(r.get('pct_fr'))} | "
            f"UTR {fmt_num(r.get('utr'), 2)}"
        )
        lineas.append(
            f"Validacion: At Vendor Time {fmt_num(r.get('at_vendor'))} mins | "
            f"Rider Accepting Time {fmt_num(r.get('rider_accepting'))} mins | "
            f"Hold Back Time {fmt_num(r.get('hold_back'))} mins"
        )
        lineas.append(f"Diagnostico: {diagnosticar(r, es_nacional=True)}")
        datos_guardados["regiones"].append({"tipo": "nacional", **{k: r.get(k) for k in
            ["region","dt","pct_late","pct_fr","utr","at_vendor","rider_accepting","hold_back"]}})
        lineas.append("")

    # ── Regiones: todas las filas EXCEPTO la ultima (nacional) ──
    df_regiones = df.iloc[:idx_nacional].copy()

    # ── RM / Santiago: detectar por nombre dentro de las regiones ──
    idx_rm = None
    rm = None
    for idx, row in df_regiones.iterrows():
        reg = str(row.get("region", "")).strip().upper()
        if any(k in reg for k in ["RM", "SANTIAGO", "METROPOLITANA"]):
            idx_rm = idx
            rm = row
            break

    if rm is not None:
        r = rm.to_dict()
        responsable = get_responsable(r.get("region", "RM"))
        lineas.append("📍 *RM (Santiago)*")
        if responsable:
            lineas.append(f"{responsable}")
        lineas.append(
            f"Dato: DT {fmt_num(r.get('dt'))} mins | "
            f"% Late Orders {fmt_pct(r.get('pct_late'))} | "
            f"% FR {fmt_pct(r.get('pct_fr'))} | "
            f"UTR {fmt_num(r.get('utr'), 2)}"
        )
        lineas.append(
            f"Validacion: At Vendor Time {fmt_num(r.get('at_vendor'))} mins | "
            f"Rider Accepting Time {fmt_num(r.get('rider_accepting'))} mins | "
            f"Hold Back Time {fmt_num(r.get('hold_back'))} mins"
        )
        lineas.append(f"Diagnostico: {diagnosticar(r)}")
        datos_guardados["regiones"].append({"tipo": "rm", **{k: r.get(k) for k in
            ["region","dt","pct_late","pct_fr","utr","at_vendor","rider_accepting","hold_back"]}})
        lineas.append("")

    # ── Top 3 regiones (excluye RM por indice, no por nombre) ──
    df_otras = df_regiones.drop(index=idx_rm) if idx_rm is not None else df_regiones.copy()

    otras = []
    for _, row in df_otras.iterrows():
        reg = str(row.get("region", "")).strip()
        if not reg or reg.upper() in ("NAN", ""):
            continue
        try:
            dt_val = float(row.get("dt", 0) or 0)
            pct_late_val = float(row.get("pct_late", 0) or 0)
            if pct_late_val <= 1.0:
                pct_late_val *= 100
            score = (dt_val / THRESHOLD_DT) + (pct_late_val / 100)
        except (TypeError, ValueError):
            score = 0
        otras.append((score, row))

    otras.sort(key=lambda x: x[0], reverse=True)
    top3 = otras[:3]

    if top3:
        lineas.append("⚠️ *TOP 3 REGIONES (Zonas con Desviación de Flujo)*")
        for i, (_, row) in enumerate(top3, 1):
            r = row.to_dict()
            reg_nombre = str(r.get("region", f"Región {i}")).strip()
            responsable = get_responsable(reg_nombre)
            resp_str = f"\n{responsable}" if responsable else ""
            lineas.append(f"{i}. 🏭 *{reg_nombre}*{resp_str}")
            lineas.append(
                f"Dato: DT {fmt_num(r.get('dt'))} mins | "
                f"🕒 % Late Orders {fmt_pct(r.get('pct_late'))} | "
                f"% FR {fmt_pct(r.get('pct_fr'))} | "
                f"UTR {fmt_num(r.get('utr'), 2)}"
            )
            lineas.append(
                f"Validación: At Vendor Time {fmt_num(r.get('at_vendor'))} mins | "
                f"Rider Accepting Time {fmt_num(r.get('rider_accepting'))} mins | "
                f"Hold Back Time {fmt_num(r.get('hold_back'))} mins"
            )
            lineas.append(f"Diagnóstico: {diagnosticar(r)}")
            datos_guardados["regiones"].append({"tipo": f"top{i}", **{k: r.get(k) for k in
                ["region","dt","pct_late","pct_fr","utr","at_vendor","rider_accepting","hold_back"]}})
            if i < len(top3):
                lineas.append("")

    mensaje = "\n".join(lineas)
    return mensaje, datos_guardados


# ─── UI STREAMLIT ─────────────────────────────────────────────────────────────

def estilo_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans', sans-serif;
    }
    .stApp {
        background: #0d0f14;
        color: #e8eaf0;
    }
    .block-container {
        padding-top: 2rem;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #111318 !important;
        border-right: 1px solid #1e2130;
    }
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3 {
        color: #f97316;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.85rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
    }

    /* Títulos */
    h1 {
        font-family: 'IBM Plex Mono', monospace !important;
        color: #f97316 !important;
        font-size: 1.6rem !important;
        letter-spacing: -0.02em;
        border-bottom: 2px solid #f97316;
        padding-bottom: 0.5rem;
        margin-bottom: 1.5rem !important;
    }
    h2 {
        font-family: 'IBM Plex Sans', sans-serif !important;
        color: #e8eaf0 !important;
        font-size: 1.1rem !important;
        font-weight: 600;
    }
    h3 {
        color: #94a3b8 !important;
        font-size: 0.9rem !important;
        font-weight: 400;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    /* Botones */
    .stButton > button {
        background: #f97316 !important;
        color: #0d0f14 !important;
        border: none !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-weight: 600 !important;
        letter-spacing: 0.05em;
        border-radius: 4px !important;
        transition: all 0.15s ease;
    }
    .stButton > button:hover {
        background: #fb923c !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(249,115,22,0.35) !important;
    }

    /* Textarea del mensaje */
    .stTextArea textarea {
        background: #111318 !important;
        color: #e8eaf0 !important;
        border: 1px solid #1e2130 !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.82rem !important;
        border-radius: 6px !important;
    }
    .stTextArea textarea:focus {
        border-color: #f97316 !important;
        box-shadow: 0 0 0 2px rgba(249,115,22,0.2) !important;
    }

    /* Métricas */
    [data-testid="stMetric"] {
        background: #111318;
        border: 1px solid #1e2130;
        border-radius: 8px;
        padding: 0.75rem 1rem;
    }
    [data-testid="stMetricLabel"] {
        color: #64748b !important;
        font-size: 0.75rem !important;
        font-family: 'IBM Plex Mono', monospace !important;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }
    [data-testid="stMetricValue"] {
        color: #f97316 !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 1.4rem !important;
    }

    /* Tablas */
    [data-testid="stDataFrame"] {
        border: 1px solid #1e2130;
        border-radius: 8px;
        overflow: hidden;
    }

    /* Alertas */
    .stAlert {
        border-radius: 6px !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: #111318;
        border-bottom: 1px solid #1e2130;
        gap: 0;
    }
    .stTabs [data-baseweb="tab"] {
        color: #64748b !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.8rem !important;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        padding: 0.6rem 1.2rem !important;
    }
    .stTabs [aria-selected="true"] {
        color: #f97316 !important;
        border-bottom: 2px solid #f97316 !important;
    }

    /* Divider */
    hr {
        border-color: #1e2130 !important;
    }

    /* Badge horario */
    .badge-horario {
        display: inline-block;
        background: rgba(249,115,22,0.15);
        color: #f97316;
        border: 1px solid rgba(249,115,22,0.4);
        border-radius: 20px;
        padding: 2px 12px;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.8rem;
        font-weight: 600;
        letter-spacing: 0.06em;
    }
    .kpi-card {
        background: #111318;
        border: 1px solid #1e2130;
        border-left: 3px solid #f97316;
        border-radius: 6px;
        padding: 0.8rem 1rem;
        margin-bottom: 0.5rem;
    }
    .region-header {
        font-family: 'IBM Plex Mono', monospace;
        color: #f97316;
        font-weight: 600;
        font-size: 0.85rem;
        letter-spacing: 0.05em;
        margin-bottom: 0.3rem;
    }
    </style>
    """, unsafe_allow_html=True)


def render_sidebar():
    st.sidebar.markdown("## ⚙️ Configuración")
    st.sidebar.markdown("---")

    fecha_sel = st.sidebar.date_input(
        "📅 Fecha del reporte",
        value=date.today(),
        help="Selecciona la fecha del período reportado"
    )

    st.sidebar.markdown("---")
    horario_sel = st.sidebar.selectbox(
        "🕐 Tipo de horario",
        HORARIOS,
        format_func=lambda h: f"{HORARIO_EMOJI[h]} {h}"
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📂 Archivo Excel")
    uploaded = st.sidebar.file_uploader(
        "Sube el Excel de Tableau",
        type=["xlsx", "xls"],
        help="El rango C2:AH38 será procesado automáticamente"
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 💾 Base de Datos")
    if st.sidebar.button("⬇️ Descargar BD SQLite", use_container_width=True):
        st.session_state["descargar_db"] = True

    return fecha_sel, horario_sel, uploaded


def render_kpi_row(label: str, valor, col):
    """Renderiza una fila de KPI en la columna dada."""
    col.metric(label, valor)


def render_vista_previa(df: pd.DataFrame):
    """Muestra el DataFrame cargado de forma limpia."""
    st.markdown("### 📋 Vista previa de datos")
    cols_mostrar = [c for c in ["region", "dt", "pct_late", "pct_fr", "utr",
                                 "at_vendor", "rider_accepting", "hold_back", "responsable"]
                    if c in df.columns]
    if cols_mostrar:
        st.dataframe(
            df[cols_mostrar].fillna("—"),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.dataframe(df.head(20), use_container_width=True)
        st.warning("⚠️ No se detectaron columnas con nombres estándar. "
                   "Verifica que el Excel tenga encabezados como: Region, DT, % Late Orders, UTR, etc.")


def render_historial():
    df_hist = listar_reportes()
    if df_hist.empty:
        st.info("Aún no hay reportes guardados.")
        return

    st.dataframe(
        df_hist.rename(columns={"id": "ID", "fecha": "Fecha", "horario": "Horario", "creado_en": "Creado"}),
        use_container_width=True,
        hide_index=True
    )

    col1, col2 = st.columns([3, 1])
    ids = df_hist["id"].tolist()
    sel_id = col1.selectbox("Seleccionar reporte", ids, format_func=lambda i: f"#{i} — {df_hist[df_hist.id==i].iloc[0]['fecha']} {df_hist[df_hist.id==i].iloc[0]['horario']}")

    if col2.button("🗑️ Eliminar", key="btn_del"):
        eliminar_reporte(sel_id)
        st.success("Reporte eliminado.")
        st.rerun()

    rep = obtener_reporte(sel_id)
    if rep:
        st.markdown(f"**Creado:** {rep['creado_en']}")
        st.text_area("Mensaje guardado", rep["mensaje"], height=400, key="hist_msg")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Reporte Operativo · LiveOps",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    estilo_css()
    init_db()

    # Sidebar
    fecha_sel, horario_sel, uploaded = render_sidebar()
    fecha_str = fecha_sel.strftime("%Y-%m-%d")
    fecha_display = fecha_sel.strftime("%d/%m/%Y")

    # Descarga BD
    if st.session_state.get("descargar_db"):
        st.session_state["descargar_db"] = False
        db_bytes = exportar_db()
        st.sidebar.download_button(
            "💾 Confirmar descarga",
            data=db_bytes,
            file_name=f"reportes_operativos_{date.today()}.db",
            mime="application/octet-stream",
            key="dl_db"
        )

    # ── Header ──
    st.markdown(f"# 📊 Reporte Operativo LiveOps")
    c1, c2, c3 = st.columns([2, 2, 3])
    c1.markdown(f"**Fecha:** `{fecha_display}`")
    c2.markdown(f"**Horario:** <span class='badge-horario'>{HORARIO_EMOJI[horario_sel]} {horario_sel}</span>", unsafe_allow_html=True)
    c3.markdown(f"**Umbral DT:** `{THRESHOLD_DT:.0f} min`")

    st.markdown("---")

    # ── Tabs ──
    tab_gen, tab_hist = st.tabs(["✏️  Generar Reporte", "🗂️  Historial"])

    # ─── TAB GENERAR ────────────────────────────────────────────────────────
    with tab_gen:
        if uploaded is None:
            st.info("👈 Sube un archivo Excel de Tableau desde el panel lateral para continuar.")
            st.markdown("""
            **Requisitos del archivo:**
            - Rango utilizado: `C2:AH38`
            - Fila 2 = encabezados de columna
            - Columnas esperadas: `Region`, `DT`, `% Late Orders`, `% FR`, `UTR`,
              `At Vendor Time`, `Rider Accepting Time`, `Hold Back Time`
            """)
            return

        # Leer Excel
        try:
            df = leer_excel(uploaded)
        except Exception as e:
            st.error(f"❌ Error al leer el Excel: {e}")
            return

        if df.empty:
            st.error("El archivo no contiene datos en el rango C2:AH38.")
            return

        # Vista previa
        with st.expander("📋 Vista previa de datos del Excel", expanded=False):
            render_vista_previa(df)

        st.markdown("### 🔄 Generación del mensaje")

        col_gen, col_opts = st.columns([1, 1])
        with col_gen:
            generar = st.button("⚡ Generar reporte", use_container_width=True, type="primary")

        # Generar o recuperar de session_state
        if generar or "ultimo_mensaje" in st.session_state:
            if generar:
                with st.spinner("Procesando datos..."):
                    mensaje, datos = generar_mensaje(df, fecha_str, horario_sel)
                st.session_state["ultimo_mensaje"] = mensaje
                st.session_state["ultimo_datos"]   = datos
                st.session_state["ultimo_horario"]  = horario_sel
                st.session_state["ultima_fecha"]    = fecha_str
            else:
                mensaje = st.session_state["ultimo_mensaje"]
                datos   = st.session_state["ultimo_datos"]

            st.success("✅ Reporte generado")

            # Mostrar KPIs nacionales si existen
            nac = next((r for r in datos.get("regiones", []) if r.get("tipo") == "nacional"), None)
            if nac:
                st.markdown("#### 🇨🇱 KPIs Nacionales")
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("DT", fmt_num(nac.get("dt")) + " min")
                k2.metric("🕒 Late Orders", fmt_pct(nac.get("pct_late")))
                k3.metric("✕ FR", fmt_pct(nac.get("pct_fr")))
                k4.metric("UTR", fmt_num(nac.get("utr"), 2))

                k5, k6, k7, _ = st.columns(4)
                k5.metric("🏪 At Vendor", fmt_num(nac.get("at_vendor")) + " min")
                k6.metric("🛵 Rider Accept.", fmt_num(nac.get("rider_accepting")) + " min")
                k7.metric("⌛ Hold Back", fmt_num(nac.get("hold_back")) + " min")

            st.markdown("---")

            # Mensaje editable
            st.markdown("#### 💬 Mensaje para Slack / WhatsApp")
            mensaje_edit = st.text_area(
                "Puedes editar el mensaje antes de guardarlo:",
                value=mensaje,
                height=520,
                key="msg_edit",
                label_visibility="collapsed"
            )

            # Acciones
            c_copy, c_save, c_dl = st.columns([2, 2, 2])

            with c_save:
                if st.button("💾 Guardar en BD", use_container_width=True):
                    guardar_reporte(
                        st.session_state.get("ultima_fecha", fecha_str),
                        st.session_state.get("ultimo_horario", horario_sel),
                        mensaje_edit,
                        st.session_state.get("ultimo_datos", {})
                    )
                    st.success("✅ Reporte guardado en la base de datos.")

            with c_dl:
                st.download_button(
                    "⬇️ Descargar .txt",
                    data=mensaje_edit.encode("utf-8"),
                    file_name=f"reporte_{horario_sel.lower()}_{fecha_str}.txt",
                    mime="text/plain",
                    use_container_width=True
                )

    # ─── TAB HISTORIAL ──────────────────────────────────────────────────────
    with tab_hist:
        st.markdown("### 🗂️ Reportes guardados")

        c_ref, c_dl_db = st.columns([3, 1])
        with c_ref:
            if st.button("🔄 Actualizar lista", use_container_width=True):
                st.rerun()
        with c_dl_db:
            db_bytes = exportar_db()
            st.download_button(
                "⬇️ Exportar BD",
                data=db_bytes,
                file_name=f"reportes_operativos_{date.today()}.db",
                mime="application/octet-stream",
                use_container_width=True,
                key="dl_db_hist"
            )

        st.markdown("---")
        render_historial()


if __name__ == "__main__":
    main()