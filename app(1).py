import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACION GENERAL
# ============================================================
st.set_page_config(
    page_title="Optimización de Generadores",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_THRESHOLD = 20.0
DEFAULT_MIN_HOURS = 5.0
DEFAULT_ACTIVE_THRESHOLD = 1.0  # Ajustado a >1.0% para ignorar el ruido cerca de cero

GEN_LOAD_RE = re.compile(r"(?:percent\s+power\s+used\s+)?GEN\s*([1-4])", re.IGNORECASE)

# ============================================================
# ESTILO — DASHBOARD OSCURO TIPO PRESENTACION
# ============================================================
st.markdown(
    """
    <style>
        .stApp {
            background: #081421;
            color: #ffffff;
        }

        [data-testid="stHeader"] {
            background: #081421;
        }

        [data-testid="stSidebar"] {
            background: #0b1b2d;
            border-right: 1px solid #18324d;
        }

        [data-testid="stSidebar"] * {
            color: #eaf2fa !important;
        }

        /* Estilo para los cuadros de entrada en la barra lateral */
        [data-testid="stSidebar"] input {
            background-color: #112a47 !important;
            color: #ffffff !important;
            border: 1px solid #18324d !important;
        }

        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 1.2rem;
            max-width: 1500px;
        }

        .dashboard-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid #00a6d6;
            padding: 5px 0 12px 0;
            margin-bottom: 18px;
        }

        .dashboard-title {
            font-size: 24px;
            font-weight: 800;
            letter-spacing: .4px;
            color: #f5f8fb;
        }

        .dashboard-title span {
            font-weight: 400;
            margin-left: 8px;
        }

        .date-box {
            text-align: right;
            min-width: 190px;
        }

        .date-label {
            color: #a9b8c8;
            font-size: 13px;
            font-weight: 700;
        }

        .date-value {
            color: #ffffff;
            font-size: 16px;
            font-weight: 700;
            margin-top: 3px;
        }

        .metric-card {
            background: #112a47;
            border-radius: 10px;
            min-height: 85px;
            padding: 12px 14px;
            position: relative;
            overflow: hidden;
            border: 1px solid #18324d;
            margin-bottom: 10px;
        }

        .metric-card .accent {
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 4px;
        }

        .metric-label {
            color: #b0c4de;
            font-size: 11px;
            font-weight: 700;
            line-height: 1.2;
            text-transform: uppercase;
        }

        .metric-value {
            color: #ffffff;
            font-size: 24px;
            font-weight: 800;
            margin-top: 6px;
            white-space: nowrap;
        }

        .metric-sub {
            color: #8fa2b7;
            font-size: 11px;
            margin-top: 2px;
        }

        .criterion {
            background: #112a47;
            border-radius: 12px;
            padding: 18px 20px;
            border: 1px solid #18324d;
            margin-bottom: 15px;
        }

        .criterion h3 {
            color: #ffffff;
            font-size: 15px;
            font-weight: 700;
            margin: 0 0 12px 0;
            border-bottom: 1px solid #18324d;
            padding-bottom: 6px;
        }

        .criterion p {
            color: #eaf2fa;
            font-size: 13px;
            line-height: 1.4;
            margin: 0 0 8px 0;
        }

        .criterion .dash {
            color: #00a6d6;
            font-weight: 800;
            margin-right: 6px;
        }

        .section-title {
            color: #ffffff;
            font-size: 20px;
            font-weight: 700;
            margin: 15px 0 10px 0;
        }

        .brand-footer {
            border-top: 1px solid #1d344c;
            padding-top: 12px;
            margin-top: 25px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            color: #ffffff;
        }

        .brand-name {
            font-size: 21px;
            font-weight: 800;
            font-style: italic;
            letter-spacing: .5px;
        }

        .brand-mark {
            display: inline-flex;
            gap: 4px;
            margin-left: 10px;
            vertical-align: middle;
        }

        .brand-mark i {
            display: inline-block;
            width: 8px;
            height: 24px;
            transform: skew(-18deg);
        }

        .footer-right {
            font-size: 11px;
            letter-spacing: 2px;
            color: #8fa2b7;
        }

        /* Estilo oscuro forzado para DataFrames de Streamlit */
        div[data-testid="stDataFrame"] {
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid #18324d;
            background-color: #112a47;
        }

        /* Estilos de Métricas Nativas de Streamlit */
        [data-testid="stMetricValue"] {
            color: #ffffff !important;
            font-size: 26px !important;
        }
        [data-testid="stMetricLabel"] {
            color: #b0c4de !important;
            font-size: 13px !important;
        }

        .stSelectbox label, .stFileUploader label {
            color: #dce7f1 !important;
            font-weight: 700 !important;
        }

        .stButton button, .stDownloadButton button {
            background: #123456;
            color: white;
            border: 1px solid #24537a;
            border-radius: 8px;
            width: 100%;
        }

        @media (max-width: 900px) {
            .dashboard-header { flex-direction: column; align-items: flex-start; gap: 10px; }
            .dashboard-title { font-size: 19px; }
            .date-box { text-align: left; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# UTILIDADES
# ============================================================
def normalize_name(value: str) -> str:
    return re.sub(r"\s+", " ", str(value).strip())


def read_uploaded_file(uploaded_file) -> Dict[str, pd.DataFrame]:
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw))
        rig_name = name.replace(".csv", "").upper()
        return {rig_name: df}

    if name.endswith((".xlsx", ".xlsm", ".xls")):
        xls = pd.ExcelFile(io.BytesIO(raw))
        result = {}
        for sheet in xls.sheet_names:
            if sheet.lower() in {"parametros", "parameters", "readme"}:
                continue
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet)
            if df is not None and not df.empty:
                result[normalize_name(sheet)] = df
        return result

    raise ValueError("Formato no soportado. Usa CSV o Excel (.xlsx/.xlsm/.xls).")


def find_time_column(df: pd.DataFrame) -> str:
    cols = {normalize_name(c).lower(): c for c in df.columns}
    for candidate in ["time", "timestamp", "datetime", "date time", "fecha", "fecha/hora"]:
        if candidate in cols:
            return cols[candidate]
    for col in df.columns:
        text = str(col).lower()
        if "timestamp" in text or "datetime" in text or "time" in text:
            return col
    raise ValueError("No encuentro la columna de fecha/hora. Se esperaba 'Time' o equivalente.")


def find_generator_load_columns(df: pd.DataFrame) -> List[str]:
    found = {}
    for col in df.columns:
        text = str(col).strip()
        low = text.lower()
        if "gen" not in low and "generator" not in low and "generador" not in low:
            continue
        m = re.search(r"(?:gen(?:erator|erador)?)[\s_\-]*(\d+)", low, re.IGNORECASE)
        if not m:
            continue
        n = int(m.group(1))
        if n < 1 or n > 4:
            continue
        priority = 2 if "percent power used" in low else (1 if "power used" in low or "load" in low or "carga" in low else 0)
        if n not in found or priority > found[n][0]:
            found[n] = (priority, col)
    return [found[n][1] for n in sorted(found)]


def find_generator_power_columns(df: pd.DataFrame) -> List[str]:
    found = {}
    for col in df.columns:
        low = str(col).lower()
        if not any(x in low for x in ["kw", "kilowatt", "power"]):
            continue
        if "percent power used" in low:
            continue
        m = re.search(r"(?:gen(?:erator|erador)?)[\s_\-]*(\d+)", low, re.IGNORECASE)
        if not m:
            continue
        n = int(m.group(1))
        if 1 <= n <= 4:
            found[n] = col
    return [found[n] for n in sorted(found)]


def prepare_rig(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str], str, List[str]]:
    df = df.copy()
    df.columns = [normalize_name(c) for c in df.columns]

    time_col = find_time_column(df)
    gen_cols = find_generator_load_columns(df)
    if not gen_cols:
        raise ValueError("No encuentro columnas de carga como 'Percent power used GEN 1' ... 'GEN 4'.")

    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df = df.dropna(subset=[time_col]).sort_values(time_col).reset_index(drop=True)

    for col in gen_cols:
        df[col] = (
            df[col].astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    power_cols = find_generator_power_columns(df)
    for col in power_cols:
        df[col] = (
            df[col].astype(str)
            .str.replace("kW", "", regex=False, case=False)
            .str.replace(",", ".", regex=False)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, gen_cols, time_col, power_cols


def gen_label(col: str) -> str:
    m = re.search(r"(?:gen(?:erator|erador)?)[\s_\-]*(\d+)", str(col), re.IGNORECASE)
    return f"GEN {m.group(1)}" if m else str(col)


def detect_events(
    df: pd.DataFrame,
    gen_cols: List[str],
    time_col: str,
    threshold: float,
    min_hours: float,
    active_threshold: float,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    work = df.copy()
    values = work[gen_cols]

    work["generadores_activos"] = (values > active_threshold).sum(axis=1)
    work["generadores_baja"] = ((values > active_threshold) & (values <= threshold)).sum(axis=1)
    work["condicion_alerta"] = (
        (work["generadores_activos"] >= 2)
        & (work["generadores_baja"] >= 2)
    )

    diffs = work[time_col].diff().dropna().dt.total_seconds().div(60)
    interval_min = float(diffs.median()) if not diffs.empty else 20.0
    if interval_min <= 0:
        interval_min = 20.0

    alert = work[work["condicion_alerta"]].copy()
    if alert.empty:
        return work, pd.DataFrame()

    alert["gap_min"] = alert[time_col].diff().dt.total_seconds().div(60)
    alert["new_group"] = alert["gap_min"].isna() | (alert["gap_min"] > interval_min * 1.5)
    alert["grupo"] = alert["new_group"].cumsum()

    events = []
    for _, g in alert.groupby("grupo"):
        start = g[time_col].min()
        end = g[time_col].max()
        duration_hours = ((end - start).total_seconds() / 3600.0) + (interval_min / 60.0)
        if duration_hours <= min_hours:
            continue

        involved = []
        for col in gen_cols:
            if (g[col] > active_threshold).any():
                involved.append(gen_label(col))

        active_vals = g[gen_cols].values
        active_vals_filtered = active_vals[active_vals > active_threshold]
        avg_load = active_vals_filtered.mean() if len(active_vals_filtered) > 0 else 0.0
        max_active = int(g["generadores_activos"].max())

        events.append({
            "Inicio": start,
            "Fin": end + pd.Timedelta(minutes=interval_min),
            "Duración (h)": round(duration_hours, 2),
            "GEN involucrados": " + ".join(involved),
            "Carga promedio GEN involucrados (%)": round(float(avg_load), 2),
            "Máx. GEN activos": max_active,
        })

    if not events:
        return work, pd.DataFrame()

    return work, pd.DataFrame(events).sort_values("Inicio").reset_index(drop=True)


def calculate_metrics(df, gen_cols, power_cols, active_threshold):
    load_avgs = {}
    active_loads = []

    for col in gen_cols:
        m = re.search(r"(?:gen(?:erator|erador)?)[\s_\-]*(\d+)", str(col), re.IGNORECASE)
        # SOLUCIÓN: Filtrar solo valores mayores al umbral activo (excluir ceros)
        valid_series = df[df[col] > active_threshold][col]
        if m:
            gen_idx = int(m.group(1))
            load_avgs[gen_idx] = float(valid_series.mean()) if not valid_series.empty else None
        if not valid_series.empty:
            active_loads.extend(valid_series.tolist())

    overall_load = float(pd.Series(active_loads).mean()) if active_loads else 0.0

    overall_power = None
    if power_cols:
        p_vals = []
        for col in power_cols:
            v_p = df[df[col] > 10.0][col]  # Filtrar potencias menores o iguales a 10 kW
            if not v_p.empty:
                p_vals.extend(v_p.tolist())
        if p_vals:
            overall_power = float(pd.Series(p_vals).mean())

    return load_avgs, overall_load, overall_power


def date_range_text(df, time_col):
    start = df[time_col].min()
    end = df[time_col].max()
    if pd.isna(start) or pd.isna(end):
        return "Fecha no disponible"
    return f"{start.day} {start.strftime('%b')} – {end.day} {end.strftime('%b')}"


def metric_card(label, value, accent=None, sub=None):
    accent_html = f'<div class="accent" style="background:{accent};"></div>' if accent else ""
    sub_html = f'<div class="metric-sub">{sub}</div>' if sub else ""
    return f"""
    <div class="metric-card">
        {accent_html}
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {sub_html}
    </div>
    """


def make_load_chart(df, gen_cols, time_col, threshold, events, rig):
    colors = {
        "GEN 1": "#00A8E8",
        "GEN 2": "#FFC000",
        "GEN 3": "#2CA02C",
        "GEN 4": "#FF4B23",
    }

    fig = go.Figure()

    for col in gen_cols:
        label = gen_label(col)
        fig.add_trace(
            go.Scatter(
                x=df[time_col],
                y=df[col],
                mode="lines",
                name=label,
                line=dict(color=colors.get(label, "#FFFFFF"), width=2.0),
                hovertemplate=f"%{{x|%d/%m %H:%M}}<br>{label}: %{{y:.2f}}%<extra></extra>",
            )
        )

    # SOLUCIÓN: Posicionar el texto del umbral en top left para evitar superposiciones
    fig.add_hline(
        y=threshold,
        line_color="#ffffff",
        line_dash="dot",
        line_width=2,
        annotation_text=f"UMBRAL {threshold:.0f}%",
        annotation_position="top left",
        annotation_font_color="#ffffff",
        annotation_font_size=11,
    )

    for _, ev in events.iterrows():
        fig.add_vrect(
            x0=ev["Inicio"],
            x1=ev["Fin"],
            fillcolor="rgba(0, 166, 214, 0.22)",
            line_width=0,
            layer="below",
        )

    fig.update_layout(
        title=dict(
            text=f"Carga de los 4 generadores — Rig {rig}",
            font=dict(color="#ffffff", size=16),
            x=0.0,
            xanchor="left",
        ),
        plot_bgcolor="#112a47",
        paper_bgcolor="#112a47",
        font=dict(color="#ffffff", size=11),
        legend=dict(
            font=dict(size=10, color="#ffffff"),
            bgcolor="rgba(8, 20, 33, 0.6)",
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
        xaxis=dict(
            title=dict(text="Tiempo", font=dict(color="#b0c4de", size=12)),
            tickfont=dict(color="#ffffff", size=10),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.08)",
            color="#ffffff",
            zeroline=False,
        ),
        yaxis=dict(
            title=dict(text="Carga (%)", font=dict(color="#b0c4de", size=12)),
            tickfont=dict(color="#ffffff", size=10),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.08)",
            color="#ffffff",
            range=[0, max(55, df[gen_cols].max().max() + 5)],
            zeroline=False,
        ),
        hovermode="x unified",
        margin=dict(l=45, r=25, t=50, b=40),
        height=420,
    )
    return fig


def make_fleet_chart(summary):
    # SOLUCIÓN: Asegurar que la columna 'Taladro' sea de tipo string para evitar ejes decimales
    s = summary.copy()
    s["Taladro"] = "Rig " + s["Taladro"].astype(str)
    s = s.sort_values("Eventos >5h", ascending=True)

    fig = go.Figure(
        go.Bar(
            x=s["Eventos >5h"],
            y=s["Taladro"],
            orientation="h",
            text=s["Eventos >5h"],
            textposition="outside",
            marker_color="#00A8E8",
            hovertemplate="%{y}<br>Eventos >5h: %{x}<extra></extra>",
        )
    )
    fig.update_layout(
        plot_bgcolor="#112a47",
        paper_bgcolor="#112a47",
        font=dict(color="#ffffff"),
        title=dict(text="Eventos prolongados por taladro", font=dict(size=15, color="#ffffff")),
        xaxis=dict(
            title="Cantidad de eventos >5 h",
            color="#ffffff",
            gridcolor="rgba(255,255,255,.08)",
            dtick=1,
        ),
        yaxis=dict(
            title="Taladro",
            color="#ffffff",
            type="category",  # SOLUCIÓN: Eje puramente categórico
            gridcolor="rgba(255,255,255,.08)",
        ),
        margin=dict(l=50, r=45, t=45, b=35),
        height=350,
    )
    return fig


# ============================================================
# SIDEBAR / CONFIGURACION
# ============================================================
with st.sidebar:
    st.header("⚙️ Configuración")
    threshold = st.number_input(
        "Umbral de baja carga (%)",
        min_value=1.0,
        max_value=100.0,
        value=DEFAULT_THRESHOLD,
        step=1.0,
    )
    min_hours = st.number_input(
        "Duración mínima para alerta (h)",
        min_value=0.5,
        max_value=72.0,
        value=DEFAULT_MIN_HOURS,
        step=0.5,
    )
    active_threshold = st.number_input(
        "Umbral para GEN activo (%)",
        min_value=0.0,
        max_value=20.0,
        value=DEFAULT_ACTIVE_THRESHOLD,
        step=0.5,
    )
    st.markdown("---")
    st.caption("Regla: ≥2 GEN activos + ≥2 GEN en baja carga + duración continua > umbral.")

uploaded = st.file_uploader(
    "📂 Carga el Excel consolidado o CSV",
    type=["xlsx", "xlsm", "xls", "csv"],
    help="Soporta múltiples taladros en pestañas o archivos individuales.",
)

if not uploaded:
    st.info("Por favor, carga el archivo de datos para iniciar el análisis.")
    st.stop()

# ============================================================
# LECTURA Y ANALISIS
# ============================================================
try:
    rig_data = read_uploaded_file(uploaded)
except Exception as e:
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

all_results = []
rig_prepared = {}
errors = {}

for rig, raw_df in rig_data.items():
    try:
        df, gen_cols, time_col, power_cols = prepare_rig(raw_df)
        work, events = detect_events(
            df, gen_cols, time_col, threshold, min_hours, active_threshold
        )
        load_avgs, overall_load, overall_power = calculate_metrics(
            df, gen_cols, power_cols, active_threshold
        )

        rig_prepared[rig] = {
            "df": df,
            "gen_cols": gen_cols,
            "time_col": time_col,
            "power_cols": power_cols,
            "work": work,
            "events": events,
            "load_avgs": load_avgs,
            "overall_load": overall_load,
            "overall_power": overall_power,
        }

        all_results.append({
            "Taladro": str(rig),
            "Registros": len(df),
            "GEN máx. activos": int(work["generadores_activos"].max()) if len(work) else 0,
            "Eventos >5h": len(events),
            "Horas en eventos >5h": round(events["Duración (h)"].sum(), 2) if not events.empty else 0.0,
            "Mayor evento (h)": round(events["Duración (h)"].max(), 2) if not events.empty else 0.0,
        })
    except Exception as e:
        errors[rig] = str(e)

if not rig_prepared:
    st.error("No se pudo procesar ningún taladro con la estructura requerida.")
    if errors:
        for rig, error in errors.items():
            st.write(f"**{rig}:** {error}")
    st.stop()

summary = pd.DataFrame(all_results)

# ============================================================
# SELECTOR Y HEADER
# ============================================================
rig_options = list(rig_prepared.keys())
selected_rig = st.selectbox("🛢️ Seleccione el taladro", rig_options)

result = rig_prepared[selected_rig]
df = result["df"]
gen_cols = result["gen_cols"]
time_col = result["time_col"]
power_cols = result["power_cols"]
work = result["work"]
events = result["events"]
load_avgs = result["load_avgs"]
overall_load = result["overall_load"]
overall_power = result["overall_power"]

period = date_range_text(df, time_col)

# SOLUCIÓN: Título principal corregido con estructura HTML limpia
st.markdown(
    f"""
    <div class="dashboard-header">
        <div class="dashboard-title">RIG {selected_rig} <span>— CARGA INDIVIDUAL DE LOS GENERADORES</span></div>
        <div class="date-box">
            <div class="date-label">Fecha analizada</div>
            <div class="date-value">{period}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# TARJETAS SUPERIORES DE METRICAS
# ============================================================
metric_cols = st.columns(6, gap="small")
accents = ["#00A8E8", "#FFC000", "#2CA02C", "#FF4B23"]

for i, col in enumerate(metric_cols[:4], start=1):
    value = load_avgs.get(i)
    value_text = f"{value:.2f}%" if value is not None else "Sin Operación"
    sub_text = "Promedio en operación" if value is not None else "Inactivo / Sin datos"
    with col:
        st.markdown(
            metric_card(f"Promedio GEN {i}", value_text, accents[i - 1], sub=sub_text),
            unsafe_allow_html=True,
        )

with metric_cols[4]:
    st.markdown(
        metric_card("Promedio General Carga", f"{overall_load:.2f}%", sub="Generadores activos"),
        unsafe_allow_html=True,
    )

with metric_cols[5]:
    power_text = f"{overall_power:,.1f} kW" if overall_power is not None else "N/D"
    power_sub = "Promedio en kW" if overall_power is not None else "Columna kW no detectada"
    st.markdown(
        metric_card("Promedio General Potencia", power_text, sub=power_sub),
        unsafe_allow_html=True,
    )

# ============================================================
# CUERPO PRINCIPAL
# ============================================================
left, right = st.columns([2.1, 0.9], gap="small")

with left:
    fig = make_load_chart(df, gen_cols, time_col, threshold, events, selected_rig)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

with right:
    # SOLUCIÓN: Renderizado HTML asegurado con unsafe_allow_html=True
    st.markdown(
        f"""
        <div class="criterion">
            <h3>CRITERIO PARA DETECTAR EVENTOS</h3>
            <p><span class="dash">-</span><b>Condición de revisión:</b> 2 o más generadores activos con carga ≤ {threshold:.0f}%.</p>
            <p><span class="dash">-</span><b>Alerta prolongada:</b> la condición anterior permanece &gt; {min_hours:g} h continuas.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    total_event_hours = events["Duración (h)"].sum() if not events.empty else 0.0
    max_duration = events["Duración (h)"].max() if not events.empty else 0.0
    max_active = int(work["generadores_activos"].max()) if len(work) else 0

    c1, c2 = st.columns(2, gap="small")
    with c1:
        st.markdown(
            metric_card("EVENTOS >5 H", f"{len(events)}", sub=f"{total_event_hours:.1f} h acumuladas"),
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            metric_card("MÁX. DURACIÓN", f"{max_duration:.1f} h" if max_duration else "0 h", sub="Evento mayor"),
            unsafe_allow_html=True,
        )

    c3, c4 = st.columns(2, gap="small")
    with c3:
        st.markdown(
            metric_card("MÁX. GEN ACTIVOS", f"{max_active}", sub=f"{len(gen_cols)} instalados"),
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            metric_card("HORAS EN EVENTOS", f"{total_event_hours:.1f} h", sub="Tiempo acumulado"),
            unsafe_allow_html=True,
        )

# ============================================================
# TABLA DE EVENTOS DETECTADOS
# ============================================================
st.markdown('<div class="section-title">EVENTOS DETECTADOS</div>', unsafe_allow_html=True)

if events.empty:
    st.success(
        f"No se encontraron eventos ineficientes con los parámetros seleccionados (Carga ≤ {threshold:.0f}%, Duración > {min_hours:g} h)."
    )
else:
    display_events = events.copy()
    display_events["Inicio"] = display_events["Inicio"].dt.strftime("%Y-%m-%d %H:%M")
    display_events["Fin"] = display_events["Fin"].dt.strftime("%Y-%m-%d %H:%M")
    st.dataframe(display_events, use_container_width=True, hide_index=True)

# ============================================================
# VISTA DE FLOTA
# ============================================================
st.markdown("---")
st.markdown('<div class="section-title">VISTA DE FLOTA</div>', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns(4)
f1.metric("Taladros cargados", len(summary))
f2.metric("Total Eventos >5 h", int(summary["Eventos >5h"].sum()))
f3.metric("Horas acumuladas flota", f"{summary['Horas en eventos >5h'].sum():.1f} h")
f4.metric("Máxima duración flota", f"{summary['Mayor evento (h)'].max():.1f} h")

fc1, fc2 = st.columns([1.2, 0.8], gap="small")
with fc1:
    st.plotly_chart(make_fleet_chart(summary), use_container_width=True, config={"displaylogo": False})
with fc2:
    st.write("")
    st.write("**Resumen detallado por taladro**")
    st.dataframe(summary.sort_values("Horas en eventos >5h", ascending=False), use_container_width=True, hide_index=True)

# ============================================================
# EXPANDER DE DIAGNÓSTICO
# ============================================================
with st.expander("🔎 Diagnóstico de columnas del taladro seleccionado"):
    st.write("**Columna de Tiempo:**", time_col)
    for col in gen_cols:
        st.write(f"**{gen_label(col)}:** `{col}`")
    if power_cols:
        st.write("**Columnas de Potencia:**", ", ".join([f"`{c}`" for c in power_cols]))

# ============================================================
# PIE DE PÁGINA
# ============================================================
st.markdown(
    """
    <div class="brand-footer">
        <div>
            <span class="brand-name">NABORS</span>
            <span class="brand-mark">
                <i style="background:#1c4bd8"></i>
                <i style="background:#1aa0f2"></i>
                <i style="background:#28a745"></i>
            </span>
        </div>
        <div class="footer-right">NABORS.COM</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# DESCARGA DE RESULTADOS
# ============================================================
out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="Resumen Flota", index=False)
    events.to_excel(writer, sheet_name="Eventos", index=False)
    work.to_excel(writer, sheet_name="Datos Procesados", index=False)
out.seek(0)

st.download_button(
    "📥 Descargar Análisis Consolidado en Excel",
    data=out,
    file_name=f"Analisis_Generadores_{selected_rig}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
