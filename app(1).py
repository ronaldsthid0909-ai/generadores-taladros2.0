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
DEFAULT_ACTIVE_THRESHOLD = 0.0

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

        .block-container {
            padding-top: 1.0rem;
            padding-bottom: 1.2rem;
            max-width: 1500px;
        }

        .dashboard-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            border-bottom: 2px solid #00a6d6;
            padding: 0 0 12px 0;
            margin-bottom: 18px;
        }

        .dashboard-title {
            font-size: 25px;
            font-weight: 800;
            letter-spacing: .4px;
            color: #f5f8fb;
        }

        .dashboard-title span {
            font-weight: 400;
            margin-left: 8px;
        }

        .date-box {
            text-align: left;
            min-width: 190px;
        }

        .date-label {
            color: #a9b8c8;
            font-size: 13px;
            font-weight: 700;
        }

        .date-value {
            color: #ffffff;
            font-size: 17px;
            font-weight: 700;
            margin-top: 3px;
        }

        .metric-card {
            background: #112a47;
            border-radius: 13px;
            min-height: 88px;
            padding: 12px 16px 10px 18px;
            position: relative;
            overflow: hidden;
            box-shadow: none;
        }

        .metric-card .accent {
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 5px;
        }

        .metric-label {
            color: #9eafc1;
            font-size: 11px;
            font-weight: 700;
            line-height: 1.2;
        }

        .metric-value {
            color: #f7f9fb;
            font-size: 27px;
            font-weight: 800;
            margin-top: 9px;
            white-space: nowrap;
        }

        .metric-sub {
            color: #91a4b8;
            font-size: 11px;
            margin-top: 2px;
        }

        .panel {
            background: #102642;
            border-radius: 14px;
            padding: 18px;
        }

        .panel-title {
            color: #ffffff;
            font-size: 17px;
            font-weight: 800;
            margin-bottom: 8px;
        }

        .criterion {
            background: #102642;
            border-radius: 28px;
            padding: 22px 26px;
            min-height: 172px;
        }

        .criterion h3 {
            color: #ffffff;
            font-size: 16px;
            margin: 0 0 17px 0;
        }

        .criterion p {
            color: #f1f5f8;
            font-size: 13px;
            line-height: 1.35;
            margin: 0 0 13px 0;
        }

        .criterion .dash {
            color: #ffffff;
            font-weight: 800;
            margin-right: 8px;
        }

        .section-title {
            color: #ffffff;
            font-size: 23px;
            font-weight: 500;
            margin: 7px 0 8px 0;
        }

        .small-footer {
            color: #8fa2b7;
            font-size: 10px;
            letter-spacing: 3px;
            text-align: right;
            margin-top: 6px;
        }

        .brand-footer {
            border-top: 1px solid #1d344c;
            padding-top: 8px;
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
            height: 28px;
            transform: skew(-18deg);
        }

        .footer-right {
            font-size: 9px;
            letter-spacing: 4px;
            color: #ffffff;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 5px;
            overflow: hidden;
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
        }

        @media (max-width: 900px) {
            .dashboard-header { align-items: flex-start; gap: 10px; }
            .dashboard-title { font-size: 19px; }
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


def read_uploaded_file(uploaded_file, selected_sheet: str | None = None) -> Dict[str, pd.DataFrame]:
    """Read CSV or Excel and return {rig/sheet_name: dataframe}."""
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw))
        return {"992": df}

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
        if "timestamp" in text or "datetime" in text:
            return col
    raise ValueError("No encuentro la columna de fecha/hora. Se esperaba 'Time' o una columna equivalente.")


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
        # Prefer exact load columns.
        priority = 2 if "percent power used" in low else (1 if "power used" in low or "load" in low or "carga" in low else 0)
        if n not in found or priority > found[n][0]:
            found[n] = (priority, col)
    return [found[n][1] for n in sorted(found)]


def find_generator_power_columns(df: pd.DataFrame) -> List[str]:
    """Best-effort detection of generator kW columns."""
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
        low_involved = []
        for col in gen_cols:
            if (g[col] > active_threshold).any():
                involved.append(gen_label(col))
            if ((g[col] > active_threshold) & (g[col] <= threshold)).any():
                low_involved.append(gen_label(col))

        avg_load = g[[c for c in gen_cols if (g[c] > active_threshold).any()]].mean().mean()
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


def calculate_metrics(df, gen_cols, power_cols):
    load_avgs = {}
    for col in gen_cols:
        m = re.search(r"(?:gen(?:erator|erador)?)[\s_\-]*(\d+)", str(col), re.IGNORECASE)
        if m:
            load_avgs[int(m.group(1))] = float(df[col].mean())

    overall_load = float(df[gen_cols].mean().mean()) if gen_cols else 0.0

    overall_power = None
    if power_cols:
        vals = df[power_cols].apply(pd.to_numeric, errors="coerce")
        overall_power = float(vals.mean(axis=1).mean())

    return load_avgs, overall_load, overall_power


def date_range_text(df, time_col):
    start = df[time_col].min()
    end = df[time_col].max()
    if pd.isna(start) or pd.isna(end):
        return "Fecha no disponible"
    return f"{start.day} {start.strftime('%b')} – {end.day} {end.strftime('%b')}".replace("Sep", "Sep")


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
                line=dict(color=colors.get(label, "#FFFFFF"), width=2.2),
                hovertemplate=f"%{{x|%d/%m %H:%M}}<br>{label}: %{{y:.2f}}%<extra></extra>",
            )
        )

    fig.add_hline(
        y=threshold,
        line_color="white",
        line_dash="dot",
        line_width=2.5,
        annotation_text=f"UMBRAL {threshold:.0f}%",
        annotation_position="top right",
        annotation_font_color="white",
        annotation_font_size=12,
    )

    for _, ev in events.iterrows():
        fig.add_vrect(
            x0=ev["Inicio"],
            x1=ev["Fin"],
            fillcolor="rgba(91,165,217,0.18)",
            line_width=0,
            layer="below",
        )

    fig.update_layout(
        title=dict(
            text=f"Carga de los 4 generadores — Rig {rig}",
            font=dict(color="white", size=17),
            x=0.0,
            xanchor="left",
        ),
        plot_bgcolor="#102642",
        paper_bgcolor="#102642",
        font=dict(color="white", size=11),
        legend=dict(
            font=dict(size=10, color="white"),
            bgcolor="rgba(0,0,0,0.08)",
            orientation="v",
            x=0.985,
            xanchor="right",
            y=0.98,
        ),
        xaxis=dict(
            title=dict(text="Tiempo", font=dict(color="white", size=12)),
            tickfont=dict(color="white", size=10),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.10)",
            color="white",
            zeroline=False,
        ),
        yaxis=dict(
            title=dict(text="Carga (%)", font=dict(color="white", size=12)),
            tickfont=dict(color="white", size=10),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.10)",
            color="white",
            range=[0, max(50, threshold + 10)],
            zeroline=False,
        ),
        hovermode="x unified",
        margin=dict(l=45, r=25, t=45, b=40),
        height=430,
    )
    return fig


def make_fleet_chart(summary):
    s = summary.sort_values("Eventos >5h", ascending=True)
    fig = go.Figure(
        go.Bar(
            x=s["Eventos >5h"],
            y=s["Taladro"],
            orientation="h",
            text=s["Eventos >5h"],
            textposition="outside",
            marker_color="#1E6FA8",
            hovertemplate="Taladro %{y}<br>Eventos >5h: %{x}<extra></extra>",
        )
    )
    fig.update_layout(
        plot_bgcolor="#102642",
        paper_bgcolor="#102642",
        font=dict(color="white"),
        title=dict(text="Eventos prolongados por taladro", font=dict(size=17)),
        xaxis=dict(title="Cantidad de eventos >5 h", color="white", gridcolor="rgba(255,255,255,.10)"),
        yaxis=dict(title="Taladro", color="white", gridcolor="rgba(255,255,255,.10)"),
        margin=dict(l=35, r=45, t=55, b=35),
        height=350,
    )
    return fig

# ============================================================
# ENCABEZADO
# ============================================================
st.markdown(
    """
    <div class="dashboard-header">
        <div class="dashboard-title">RIG <span id="rig-title">—</span> <span>CARGA INDIVIDUAL DE LOS GENERADORES</span></div>
        <div class="date-box">
            <div class="date-label">Fecha analizada</div>
            <div class="date-value" id="date-title">Cargue la información</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# SIDEBAR / CARGA
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
    "📂 Carga el Excel consolidado",
    type=["xlsx", "xlsm", "xls", "csv"],
    help="Para varios taladros: una hoja por taladro. Ej.: 992, M47, M48, X38, X40, X42, X43 y X45.",
)

if not uploaded:
    st.info("Carga el Excel consolidado para mostrar todos los taladros en una sola aplicación.")
    st.stop()

# ============================================================
# LECTURA Y ANALISIS DE TODAS LAS HOJAS
# ============================================================
try:
    rig_data = read_uploaded_file(uploaded)
except Exception as e:
    st.error(f"No pude leer el archivo: {e}")
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
        load_avgs, overall_load, overall_power = calculate_metrics(df, gen_cols, power_cols)

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
    st.error("No encontré ningún taladro con la estructura esperada.")
    if errors:
        for rig, error in errors.items():
            st.write(f"**{rig}:** {error}")
    st.stop()

summary = pd.DataFrame(all_results)

# ============================================================
# SELECTOR PRINCIPAL
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

# ============================================================
# ACTUALIZAR TITULO VISUAL
# ============================================================
period = date_range_text(df, time_col)
st.markdown(
    f"""
    <script>
    const rt = window.parent.document.querySelector('#rig-title');
    if (rt) rt.innerText = '{selected_rig}';
    const dt = window.parent.document.querySelector('#date-title');
    if (dt) dt.innerText = '{period}';
    </script>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# TARJETAS SUPERIORES
# ============================================================
metric_cols = st.columns(6, gap="small")
accents = ["#5AA5DD", "#FFC000", "#2C7E1D", "#FF431B"]

for i, col in enumerate(metric_cols[:4], start=1):
    value = load_avgs.get(i)
    value_text = f"{value:.2f}%" if value is not None else "N/D"
    with col:
        st.markdown(
            metric_card(f"Promedio de Potencia de GEN {i}", value_text, accents[i-1]),
            unsafe_allow_html=True,
        )

with metric_cols[4]:
    st.markdown(
        metric_card("Promedio General de carga", f"{overall_load:.2f}%"),
        unsafe_allow_html=True,
    )

with metric_cols[5]:
    if overall_power is None:
        power_text = "N/D"
        power_sub = "No se detectó columna kW"
    else:
        power_text = f"{overall_power:,.2f} kW"
        power_sub = None
    st.markdown(
        metric_card("Promedio General de Potencia", power_text, sub=power_sub),
        unsafe_allow_html=True,
    )

# ============================================================
# CUERPO PRINCIPAL
# ============================================================
left, right = st.columns([2.05, 1.0], gap="small")

with left:
    fig = make_load_chart(df, gen_cols, time_col, threshold, events, selected_rig)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

with right:
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

    st.write("")
    c1, c2 = st.columns(2, gap="small")
    total_event_hours = events["Duración (h)"].sum() if not events.empty else 0.0
    max_duration = events["Duración (h)"].max() if not events.empty else 0.0
    max_active = int(work["generadores_activos"].max()) if len(work) else 0

    with c1:
        st.markdown(
            metric_card(
                "EVENTOS >5 H",
                f"{len(events)}",
                sub=f"{total_event_hours:.2f} h acumuladas",
            ),
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            metric_card(
                "MÁX. DURACIÓN",
                f"{max_duration:.0f} h" if max_duration else "0 h",
                sub="evento más prolongado",
            ),
            unsafe_allow_html=True,
        )

    st.write("")
    c3, c4 = st.columns(2, gap="small")
    with c3:
        st.markdown(
            metric_card(
                "Max de GEN activos",
                f"{max_active}",
                sub=f"{len(gen_cols)} disponibles",
            ),
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            metric_card(
                "HORAS EN EVENTOS",
                f"{total_event_hours:.2f} h",
                sub="Total horas acumuladas",
            ),
            unsafe_allow_html=True,
        )

# ============================================================
# EVENTOS
# ============================================================
st.markdown('<div class="section-title">EVENTOS DETECTADOS</div>', unsafe_allow_html=True)

if events.empty:
    st.success(
        f"No se encontraron eventos que cumplan: ≥2 generadores activos, ≥2 en baja carga (≤{threshold:.0f}%) y duración >{min_hours:g} h continuas."
    )
else:
    display_events = events.copy()
    display_events["Inicio"] = display_events["Inicio"].dt.strftime("%Y-%m-%d %H:%M:%S")
    display_events["Fin"] = display_events["Fin"].dt.strftime("%Y-%m-%d %H:%M:%S")
    st.dataframe(display_events, use_container_width=True, hide_index=True)

    csv = events.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "📥 Descargar eventos de este taladro",
        data=csv,
        file_name=f"eventos_{selected_rig}.csv",
        mime="text/csv",
    )

# ============================================================
# VISTA DE FLOTA
# ============================================================
st.markdown("---")
st.markdown('<div class="section-title">VISTA DE FLOTA</div>', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns(4)
f1.metric("Taladros cargados", len(summary))
f2.metric("Eventos >5 h", int(summary["Eventos >5h"].sum()))
f3.metric("Horas acumuladas", f"{summary['Horas en eventos >5h'].sum():.1f} h")
f4.metric("Máx. duración", f"{summary['Mayor evento (h)'].max():.1f} h")

fc1, fc2 = st.columns([1.25, 0.75], gap="small")
with fc1:
    st.plotly_chart(make_fleet_chart(summary), use_container_width=True, config={"displaylogo": False})
with fc2:
    st.dataframe(summary.sort_values("Horas en eventos >5h", ascending=False), use_container_width=True, hide_index=True)

# ============================================================
# ERRORES / ESTRUCTURA
# ============================================================
if errors:
    with st.expander("⚠️ Hojas que no pudieron analizarse"):
        for rig, error in errors.items():
            st.write(f"**{rig}:** {error}")

with st.expander("🔎 Columnas detectadas en el taladro seleccionado"):
    st.write("**Fecha/hora:**", time_col)
    for col in gen_cols:
        st.write(f"**{gen_label(col)}:** `{col}`")
    if power_cols:
        st.write("**Columnas de potencia detectadas:**")
        for col in power_cols:
            st.write(f"`{col}`")
    else:
        st.info("No se detectaron columnas de potencia en kW; por eso el promedio general de potencia aparece como N/D.")

# ============================================================
# PIE
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
# DESCARGA DEL ANALISIS COMPLETO
# ============================================================
out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="Resumen Flota", index=False)
    events.to_excel(writer, sheet_name=f"Eventos {str(selected_rig)[:20]}", index=False)
    work.to_excel(writer, sheet_name=f"Datos {str(selected_rig)[:19]}", index=False)
out.seek(0)

st.download_button(
    "📥 Descargar análisis en Excel",
    data=out,
    file_name=f"Analisis_Generadores_{selected_rig}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
