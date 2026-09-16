import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACIÓN GENERAL Y ESTILOS
# ============================================================
st.set_page_config(
    page_title="Optimización de Generadores",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_THRESHOLD = 20.0
DEFAULT_MIN_HOURS = 5.0
ACTIVE_THRESHOLD = 0.0

st.markdown(
    """
    <style>
        .stApp {
            background-color: #081421 !important;
            color: #ffffff !important;
        }

        [data-testid="stHeader"] {
            background-color: #081421 !important;
        }

        [data-testid="stSidebar"] {
            background-color: #0b1b2d !important;
            border-right: 1px solid #18324d !important;
        }

        /* Corrección de color de texto para el menú lateral e inputs */
        [data-testid="stSidebar"] label {
            color: #ffffff !important;
            font-size: 13px !important;
            font-weight: 600 !important;
        }
        
        [data-testid="stSidebar"] input {
            color: #000000 !important;
            background-color: #ffffff !important;
            font-weight: bold !important;
        }

        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 1rem;
            max-width: 1650px;
        }

        .dashboard-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            border-bottom: 2px solid #00a6d6;
            padding-bottom: 8px;
            margin-bottom: 15px;
        }

        .dashboard-title {
            font-size: 20px;
            font-weight: 800;
            color: #ffffff;
            letter-spacing: 0.5px;
        }

        .dashboard-title span {
            font-weight: 400;
            font-size: 16px;
            margin-left: 8px;
        }

        .date-box {
            text-align: right;
        }

        .date-label {
            color: #ffffff;
            font-size: 11px;
            font-weight: 600;
        }

        .date-value {
            color: #ffffff;
            font-size: 16px;
            font-weight: 800;
        }

        /* Tarjetas Métricas */
        .metric-card {
            background: #112a47;
            border-radius: 6px;
            height: 95px;
            padding: 10px 12px;
            position: relative;
            overflow: hidden;
            border: 1px solid #18324d;
            box-sizing: border-box;
        }

        .metric-card .accent {
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 6px;
        }

        .metric-label {
            color: #a0b2c6;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
        }

        .metric-value {
            color: #ffffff;
            font-size: 24px;
            font-weight: 800;
            margin-top: 2px;
            white-space: nowrap;
        }

        .metric-sub {
            color: #7a8f05;
            font-size: 11px;
            margin-top: 2px;
        }

        /* Criterio Box */
        .criterion {
            background: #112a47;
            border-radius: 8px;
            padding: 16px;
            border: 1px solid #18324d;
            margin-bottom: 12px;
            height: 180px;
            box-sizing: border-box;
        }

        .criterion h3 {
            color: #ffffff;
            font-size: 13px;
            font-weight: 800;
            margin: 0 0 12px 0;
            letter-spacing: 0.5px;
        }

        .criterion p {
            color: #ffffff;
            font-size: 12px;
            line-height: 1.4;
            margin: 0 0 8px 0;
        }

        .section-title {
            color: #ffffff;
            font-size: 16px;
            font-weight: 800;
            margin: 20px 0 10px 0;
            letter-spacing: 0.5px;
        }

        /* Footer */
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
            font-size: 18px;
            font-weight: 900;
            font-style: italic;
            letter-spacing: 1px;
        }

        .footer-right {
            font-size: 11px;
            letter-spacing: 2px;
            color: #a0b2c6;
            font-weight: 600;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# FUNCIONES DE PROCESAMIENTO DE DATOS
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

    raise ValueError("Formato no soportado.")


def find_time_column(df: pd.DataFrame) -> str:
    cols = {normalize_name(c).lower(): c for c in df.columns}
    for candidate in ["time", "timestamp", "datetime", "date time", "fecha", "fecha/hora"]:
        if candidate in cols:
            return cols[candidate]
    for col in df.columns:
        text = str(col).lower()
        if "timestamp" in text or "datetime" in text or "time" in text:
            return col
    raise ValueError("No se encontró columna de tiempo.")


def find_generator_load_columns(df: pd.DataFrame) -> List[str]:
    found = {}
    for col in df.columns:
        text = str(col).strip()
        low = text.lower()
        if "gen" not in low and "generator" not in low and "generador" not in low:
            continue
        if "total" in low:
            continue
        m = re.search(r"(?:gen(?:erator|erador)?[\s_\-]*(\d+))", low, re.IGNORECASE)
        if not m:
            continue
        n = int(m.group(1))
        if n < 1 or n > 4:
            continue
        priority = 2 if "percent power used" in low else 1
        if n not in found or priority > found[n][0]:
            found[n] = (priority, col)
    return [found[n][1] for n in sorted(found)]


def find_overall_columns(df: pd.DataFrame) -> Tuple[str, str]:
    total_load_col = None
    total_power_col = None

    for col in df.columns:
        low = str(col).lower().strip()
        if "percent power used total" in low or ("percent" in low and "total" in low):
            total_load_col = col
        elif "gen total power" in low or ("total" in low and "power" in low and "percent" not in low):
            total_power_col = col

    return total_load_col, total_power_col


def prepare_rig(df: pd.DataFrame):
    df = df.copy()
    df.columns = [normalize_name(c) for c in df.columns]

    time_col = find_time_column(df)
    gen_cols = find_generator_load_columns(df)
    total_load_col, total_power_col = find_overall_columns(df)

    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df = df.dropna(subset=[time_col]).sort_values(time_col).reset_index(drop=True)

    for col in gen_cols:
        df[col] = (
            df[col].astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    if total_load_col:
        df[total_load_col] = (
            df[total_load_col].astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df[total_load_col] = pd.to_numeric(df[total_load_col], errors="coerce")

    if total_power_col:
        df[total_power_col] = (
            df[total_power_col].astype(str)
            .str.replace("kW", "", regex=False, case=False)
            .str.replace(",", ".", regex=False)
        )
        df[total_power_col] = pd.to_numeric(df[total_power_col], errors="coerce")

    return df, gen_cols, time_col, total_load_col, total_power_col


def gen_label(col: str) -> str:
    m = re.search(r"(?:gen(?:erator|erador)?[\s_\-]*(\d+))", str(col), re.IGNORECASE)
    return f"GEN {m.group(1)}" if m else str(col)


def detect_events(df, gen_cols, time_col, threshold, min_hours):
    work = df.copy()
    values = work[gen_cols]

    work["generadores_activos"] = (values > ACTIVE_THRESHOLD).sum(axis=1)
    work["generadores_baja"] = ((values > ACTIVE_THRESHOLD) & (values <= threshold)).sum(axis=1)
    work["condicion_alerta"] = (
        (work["generadores_activos"] >= 2)
        & (work["generadores_baja"] >= 2)
    )

    diffs = work[time_col].diff().dropna().dt.total_seconds().div(60)
    interval_min = float(diffs.median()) if not diffs.empty else 20.0

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

        involved = [gen_label(c) for c in gen_cols if (g[c] > ACTIVE_THRESHOLD).any()]
        active_vals = g[gen_cols].values
        active_vals_filtered = active_vals[active_vals > ACTIVE_THRESHOLD]
        avg_load = active_vals_filtered.mean() if len(active_vals_filtered) > 0 else 0.0

        events.append({
            "Inicio": start,
            "Fin": end + pd.Timedelta(minutes=interval_min),
            "Duración (h)": round(duration_hours, 2),
            "GEN involucrados": " + ".join(involved),
            "Carga promedio GEN involucrados (%)": round(float(avg_load), 2),
            "Máx. GEN activos": int(g["generadores_activos"].max()),
        })

    return work, pd.DataFrame(events).sort_values("Inicio").reset_index(drop=True) if events else (work, pd.DataFrame())


def render_card(label, value, accent="#00A8E8", sub=""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="accent" style="background:{accent};"></div>
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# Gráfica con leyenda superior
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
                line=dict(color=colors.get(label, "#FFFFFF"), width=1.5),
                hovertemplate=f"%{{x|%d/%m %H:%M}}<br>{label}: %{{y:.2f}}%<extra></extra>",
            )
        )

    fig.add_hline(
        y=threshold,
        line_color="#ffffff",
        line_dash="dot",
        line_width=1.2,
        annotation_text=f"UMBRAL {threshold:.0f}%",
        annotation_position="top right",
        annotation_font_color="#ffffff",
        annotation_font_size=11,
    )

    for _, ev in events.iterrows():
        fig.add_vrect(
            x0=ev["Inicio"],
            x1=ev["Fin"],
            fillcolor="rgba(0, 166, 214, 0.25)",
            line_width=0,
            layer="below",
        )

    fig.update_layout(
        title=dict(
            text=f"Carga de los 4 generadores — Rig {rig}",
            font=dict(color="#ffffff", size=14, family="Arial Black"),
            x=0.0,
            y=0.96,
        ),
        plot_bgcolor="#112a47",
        paper_bgcolor="#112a47",
        font=dict(color="#ffffff", size=11),
        # Leyenda en la parte superior derecha dentro del gráfico
        legend=dict(
            font=dict(size=10, color="#ffffff"),
            bgcolor="rgba(17, 42, 71, 0.8)",
            orientation="v",
            yanchor="top",
            y=0.98,
            xanchor="right",
            x=0.98,
        ),
        xaxis=dict(
            title="Tiempo",
            tickfont=dict(color="#ffffff", size=11),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.1)",
        ),
        yaxis=dict(
            title="Carga (%)",
            tickfont=dict(color="#ffffff", size=11),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.1)",
            range=[0, max(55, df[gen_cols].max().max() + 5)],
        ),
        hovermode="x unified",
        margin=dict(l=40, r=20, t=50, b=40),
        height=380,
    )
    return fig


# ============================================================
# INTERFAZ Y CONTROL
# ============================================================
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
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
    st.markdown("---")
    st.markdown(
        "<p style='color: #a0b2c6; font-size: 11px;'>"
        "Criterio: ≥2 GEN activos (>0%) con carga ≤ umbral durante un tiempo continuo mayor a la duración mínima."
        "</p>",
        unsafe_allow_html=True,
    )

uploaded = st.file_uploader("📂 Cargar archivo (.xlsx, .csv)", type=["xlsx", "xlsm", "xls", "csv"])

if not uploaded:
    st.info("Cargue un archivo para desplegar la interfaz.")
    st.stop()

rig_data = read_uploaded_file(uploaded)
rig_prepared = {}

for rig, raw_df in rig_data.items():
    df, gen_cols, time_col, total_load_col, total_power_col = prepare_rig(raw_df)
    work, events = detect_events(df, gen_cols, time_col, threshold, min_hours)

    load_avgs = {}
    for col in gen_cols:
        m = re.search(r"(?:gen(?:erator|erador)?[\s_\-]*(\d+))", str(col), re.IGNORECASE)
        valid_series = df[df[col] > ACTIVE_THRESHOLD][col]
        if not valid_series.empty and m:
            load_avgs[int(m.group(1))] = float(valid_series.mean())

    overall_load = float(df[total_load_col].mean()) if total_load_col else 0.0
    overall_power = float(df[total_power_col].mean()) if total_power_col else 0.0

    rig_prepared[rig] = {
        "df": df,
        "gen_cols": gen_cols,
        "time_col": time_col,
        "work": work,
        "events": events,
        "load_avgs": load_avgs,
        "overall_load": overall_load,
        "overall_power": overall_power,
    }

selected_rig = st.selectbox("Seleccionar Taladro", list(rig_prepared.keys()))
res = rig_prepared[selected_rig]

start_t = res["df"][res["time_col"]].min()
end_t = res["df"][res["time_col"]].max()
period_str = f"{start_t.strftime('%d %b')} – {end_t.strftime('%d %b')}" if pd.notna(start_t) else ""

# Header
st.markdown(
    f"""
    <div class="dashboard-header">
        <div class="dashboard-title">RIG {selected_rig} <span>CARGA INDIVIDUAL DE LOS GENERADORES</span></div>
        <div class="date-box">
            <div class="date-label">Fecha analizada</div>
            <div class="date-value">{period_str}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# 6 Tarjetas superiores divididas uniformemente
accents = ["#00A8E8", "#FFC000", "#2CA02C", "#FF4B23"]
cols = st.columns(6)

for i in range(1, 5):
    val = res["load_avgs"].get(i)
    txt = f"{val:.2f}%" if val is not None else "N/A"
    with cols[i - 1]:
        render_card(f"Promedio de Potencia de GEN {i}", txt, accents[i - 1])

with cols[4]:
    render_card("Promedio General de carga", f"{res['overall_load']:.2f}%", "#00A8E8")

with cols[5]:
    p_txt = f"{res['overall_power']:.2f} kW" if res["overall_power"] else "N/D"
    render_card("Promedio General de Potencia", p_txt, "#00A8E8")

st.markdown("<br>", unsafe_allow_html=True)

# ALINEACIÓN PERFECTA: 4 columnas a la izquierda (Abarca la gráfica), 2 a la derecha (Abarca los cuadros)
left_box, right_box = st.columns([4, 2], gap="large")

with left_box:
    fig = make_load_chart(res["df"], res["gen_cols"], res["time_col"], threshold, res["events"], selected_rig)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

with right_box:
    st.markdown(
        f"""
        <div class="criterion">
            <h3>CRITERIO PARA DETECTAR EVENTOS</h3>
            <p>- <b>Condición de revisión:</b> 2 o más generadores activos con carga ≤ {threshold:.0f}%.</p>
            <p>- <b>Alerta prolongada:</b> la condición anterior permanece &gt; {min_hours:g} h continuas.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    ev_df = res["events"]
    tot_h = ev_df["Duración (h)"].sum() if not ev_df.empty else 0.0
    max_h = ev_df["Duración (h)"].max() if not ev_df.empty else 0.0
    max_act = int(res["work"]["generadores_activos"].max()) if len(res["work"]) else 0

    r1_1, r1_2 = st.columns(2)
    with r1_1:
        render_card("EVENTOS >5 H", f"{len(ev_df)}", "#00A8E8", f"{tot_h:.2f} h acumuladas")
    with r1_2:
        render_card("MÁX. DURACIÓN", f"{max_h:.0f} h" if max_h else "0 h", "#00A8E8", "evento más prolongado")

    st.markdown("<div style='margin-top: 8px;'></div>", unsafe_allow_html=True)

    r2_1, r2_2 = st.columns(2)
    with r2_1:
        render_card("Max de GEN activos", f"{max_act}", "#00A8E8", f"{len(res['gen_cols'])} disponibles")
    with r2_2:
        render_card("HORAS EN EVENTOS", f"{tot_h:.2f} h", "#00A8E8", "Total horas acumuladas")

# Tabla de Eventos
st.markdown('<div class="section-title">EVENTOS DETECTADOS</div>', unsafe_allow_html=True)
if res["events"].empty:
    st.info("No se registraron eventos prolongados bajo los parámetros establecidos.")
else:
    disp_events = res["events"].copy()
    disp_events["Inicio"] = disp_events["Inicio"].dt.strftime("%Y-%m-%d %H:%M:%S")
    disp_events["Fin"] = disp_events["Fin"].dt.strftime("%Y-%m-%d %H:%M:%S")
    st.dataframe(disp_events, use_container_width=True, hide_index=True)

# Footer
st.markdown(
    """
    <div class="brand-footer">
        <div>
            <span class="brand-name">NABORS</span>
        </div>
        <div class="footer-right">N A B O R S . C O M &nbsp;&nbsp;|&nbsp;&nbsp; 1</div>
    </div>
    """,
    unsafe_allow_html=True,
)
