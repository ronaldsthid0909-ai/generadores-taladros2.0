import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ============================================================
# CONFIGURACIÓN GENERAL
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

# ============================================================
# ESTILO Y TEMAS OSCUROS
# ============================================================
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

        [data-testid="stSidebar"] p, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] span {
            color: #eaf2fa !important;
        }

        /* Estilos específicos para solucionar la visibilidad de los inputs en el sidebar */
        [data-testid="stSidebar"] div[data-baseweb="input"] {
            background-color: #112a47 !important;
            border: 1px solid #18324d !important;
            border-radius: 6px !important;
        }

        [data-testid="stSidebar"] input {
            color: #ffffff !important;
            background-color: #112a47 !important;
            font-weight: 600 !important;
        }

        [data-testid="stSidebar"] button {
            color: #ffffff !important;
            background-color: #112a47 !important;
            border-color: #18324d !important;
        }

        .stSelectbox label, .stFileUploader label, .stNumberInput label {
            color: #ffffff !important;
            font-weight: 600 !important;
            font-size: 14px !important;
        }

        .block-container {
            padding-top: 1.2rem !important;
            padding-bottom: 1.2rem;
            max-width: 1600px;
        }

        .dashboard-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid #00a6d6;
            padding: 2px 0 10px 0;
            margin-bottom: 15px;
        }

        .dashboard-title {
            font-size: 22px;
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
            font-size: 12px;
            font-weight: 700;
        }

        .date-value {
            color: #ffffff;
            font-size: 15px;
            font-weight: 700;
            margin-top: 2px;
        }

        .metric-card {
            background: #112a47;
            border-radius: 8px;
            height: 90px;
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
            width: 5px;
        }

        .metric-label {
            color: #b0c4de;
            font-size: 10px;
            font-weight: 700;
            line-height: 1.2;
            text-transform: uppercase;
        }

        .metric-value {
            color: #ffffff;
            font-size: 21px;
            font-weight: 800;
            margin-top: 4px;
            white-space: nowrap;
        }

        .metric-sub {
            color: #8fa2b7;
            font-size: 10px;
            margin-top: 2px;
        }

        .criterion {
            background: #112a47;
            border-radius: 8px;
            padding: 14px 16px;
            border: 1px solid #18324d;
            margin-bottom: 12px;
            height: 175px;
            box-sizing: border-box;
        }

        .criterion h3 {
            color: #ffffff;
            font-size: 13px;
            font-weight: 700;
            margin: 0 0 8px 0;
            border-bottom: 1px solid #18324d;
            padding-bottom: 5px;
        }

        .criterion p {
            color: #eaf2fa;
            font-size: 11px;
            line-height: 1.35;
            margin: 0 0 5px 0;
        }

        .criterion .dash {
            color: #00a6d6;
            font-weight: 800;
            margin-right: 5px;
        }

        .section-title {
            color: #ffffff;
            font-size: 16px;
            font-weight: 700;
            margin: 15px 0 8px 0;
        }

        .brand-footer {
            border-top: 1px solid #1d344c;
            padding-top: 10px;
            margin-top: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            color: #ffffff;
        }

        .brand-name {
            font-size: 18px;
            font-weight: 800;
            font-style: italic;
        }

        .brand-mark {
            display: inline-flex;
            gap: 4px;
            margin-left: 10px;
        }

        .brand-mark i {
            display: inline-block;
            width: 7px;
            height: 20px;
            transform: skew(-18deg);
        }

        .footer-right {
            font-size: 10px;
            letter-spacing: 2px;
            color: #8fa2b7;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# FUNCIONES AUXILIARES Y GENERACIÓN DE POWERPOINT
# ============================================================
def create_powerpoint_slide(selected_rig, period_str, res, threshold, min_hours):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Fondo oscuro #081421
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(8, 20, 33)
    
    def add_card(left, top, width, height, bg_rgb, border_rgb=None):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_rgb
        if border_rgb:
            shape.line.color.rgb = border_rgb
            shape.line.width = Pt(1)
        else:
            shape.line.fill.background()
        return shape

    # Encabezado
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(8.0), Inches(0.6))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    r1 = p.add_run(); r1.text = f"RIG {selected_rig} "; r1.font.bold = True; r1.font.size = Pt(20); r1.font.color.rgb = RGBColor(255, 255, 255)
    r2 = p.add_run(); r2.text = "CARGA INDIVIDUAL DE LOS GENERADORES"; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(255, 255, 255)

    txBox2 = slide.shapes.add_textbox(Inches(9.5), Inches(0.15), Inches(3.3), Inches(0.6))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.RIGHT
    r_lbl = p2.add_run(); r_lbl.text = "Fecha analizada\n"; r_lbl.font.size = Pt(11); r_lbl.font.color.rgb = RGBColor(169, 184, 200); r_lbl.font.bold = True
    p2_val = tf2.add_paragraph(); p2_val.alignment = PP_ALIGN.RIGHT
    r_val = p2_val.add_run(); r_val.text = period_str; r_val.font.size = Pt(15); r_val.font.bold = True; r_val.font.color.rgb = RGBColor(255, 255, 255)

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.85), Inches(12.333), Inches(0.02))
    line.fill.solid(); line.fill.fore_color.rgb = RGBColor(0, 166, 214); line.line.fill.background()

    # 6 Tarjetas Superiores
    card_w = Inches(1.92); card_h = Inches(0.95); top_pos = Inches(1.0)
    accents_rgb = [RGBColor(0, 168, 232), RGBColor(255, 192, 0), RGBColor(44, 160, 44), RGBColor(255, 75, 35), RGBColor(0, 168, 232), RGBColor(0, 168, 232)]
    bg_card = RGBColor(17, 42, 71); border_card = RGBColor(24, 50, 77)

    cards_data = []
    for i in range(1, 5):
        val = res["load_avgs"].get(i)
        txt = f"{val:.2f}%" if val is not None else "N/A"
        cards_data.append((f"Promedio de Potencia de GEN {i}", txt))
    cards_data.append(("Promedio General de carga", f"{res['overall_load']:.2f}%"))
    p_txt = f"{res['overall_power']:.2f} kW" if res["overall_power"] else "N/D"
    cards_data.append(("Promedio General de Potencia", p_txt))

    for idx, (lbl, val_str) in enumerate(cards_data):
        left_pos = Inches(0.5) + idx * Inches(2.08)
        add_card(left_pos, top_pos, card_w, card_h, bg_card, border_card)
        add_card(left_pos, top_pos, Inches(0.08), card_h, accents_rgb[idx])
        
        tb = slide.shapes.add_textbox(left_pos + Inches(0.12), top_pos + Inches(0.05), card_w - Inches(0.15), card_h - Inches(0.1))
        tf_c = tb.text_frame; tf_c.word_wrap = True
        p_l = tf_c.paragraphs[0]; r_l = p_l.add_run(); r_l.text = lbl; r_l.font.size = Pt(8); r_l.font.bold = True; r_l.font.color.rgb = RGBColor(176, 196, 222)
        p_v = tf_c.add_paragraph(); r_v = p_v.add_run(); r_v.text = val_str; r_v.font.size = Pt(17); r_v.font.bold = True; r_v.font.color.rgb = RGBColor(255, 255, 255)

    # Imagen de Gráfica Plotly
    fig = make_load_chart(res["df"], res["gen_cols"], res["time_col"], threshold, res["events"], selected_rig)
    try:
        img_bytes = fig.to_image(format="png", width=1200, height=600, scale=2)
        image_stream = io.BytesIO(img_bytes)
        slide.shapes.add_picture(image_stream, Inches(0.5), Inches(2.1), width=Inches(8.1))
    except Exception:
        pass

    # Criterio y métricas derechas
    right_left = Inches(8.75); right_w = Inches(4.08)
    add_card(right_left, Inches(2.1), right_w, Inches(1.85), bg_card, border_card)
    tb_crit = slide.shapes.add_textbox(right_left + Inches(0.15), Inches(2.15), right_w - Inches(0.3), Inches(1.75))
    tf_crit = tb_crit.text_frame; tf_crit.word_wrap = True
    p_cr1 = tf_crit.paragraphs[0]; r_cr1 = p_cr1.add_run(); r_cr1.text = "CRITERIO PARA DETECTAR EVENTOS\n"; r_cr1.font.size = Pt(11); r_cr1.font.bold = True; r_cr1.font.color.rgb = RGBColor(255, 255, 255)
    p_cr2 = tf_crit.add_paragraph(); r_cr2 = p_cr2.add_run(); r_cr2.text = f"- Condición de revisión: 2 o más generadores activos con carga ≤ {threshold:.0f}%.\n"; r_cr2.font.size = Pt(10); r_cr2.font.color.rgb = RGBColor(234, 242, 250)
    p_cr3 = tf_crit.add_paragraph(); r_cr3 = p_cr3.add_run(); r_cr3.text = f"- Alerta prolongada: la condición anterior permanece > {min_hours:g} h continuas."; r_cr3.font.size = Pt(10); r_cr3.font.color.rgb = RGBColor(234, 242, 250)

    ev_df = res["events"]
    tot_h = ev_df["Duración (h)"].sum() if not ev_df.empty else 0.0
    max_h = ev_df["Duración (h)"].max() if not ev_df.empty else 0.0
    max_act = int(res["work"]["generadores_activos"].max()) if len(res["work"]) else 0

    sub_w = Inches(1.98); sub_h = Inches(0.9)
    
    # Tarjeta 1
    add_card(right_left, Inches(4.08), sub_w, sub_h, bg_card, border_card)
    tb_m1 = slide.shapes.add_textbox(right_left + Inches(0.1), Inches(4.1), sub_w - Inches(0.2), sub_h)
    tf_m1 = tb_m1.text_frame
    p1 = tf_m1.paragraphs[0]; r = p1.add_run(); r.text = "EVENTOS >5 H\n"; r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = RGBColor(176, 196, 222)
    p2 = tf_m1.add_paragraph(); r = p2.add_run(); r.text = f"{len(ev_df)}\n"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    p3 = tf_m1.add_paragraph(); r = p3.add_run(); r.text = f"{tot_h:.2f} h acumuladas"; r.font.size = Pt(8); r.font.color.rgb = RGBColor(143, 162, 183)

    # Tarjeta 2
    add_card(right_left + Inches(2.1), Inches(4.08), sub_w, sub_h, bg_card, border_card)
    tb_m2 = slide.shapes.add_textbox(right_left + Inches(2.2), Inches(4.1), sub_w - Inches(0.2), sub_h)
    tf_m2 = tb_m2.text_frame
    p1 = tf_m2.paragraphs[0]; r = p1.add_run(); r.text = "MÁX. DURACIÓN\n"; r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = RGBColor(176, 196, 222)
    p2 = tf_m2.add_paragraph(); r = p2.add_run(); r.text = f"{max_h:.0f} h\n" if max_h else "0 h\n"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    p3 = tf_m2.add_paragraph(); r = p3.add_run(); r.text = "evento más prolongado"; r.font.size = Pt(8); r.font.color.rgb = RGBColor(143, 162, 183)

    # Tarjeta 3
    add_card(right_left, Inches(5.1), sub_w, sub_h, bg_card, border_card)
    tb_m3 = slide.shapes.add_textbox(right_left + Inches(0.1), Inches(5.12), sub_w - Inches(0.2), sub_h)
    tf_m3 = tb_m3.text_frame
    p1 = tf_m3.paragraphs[0]; r = p1.add_run(); r.text = "Max de GEN activos\n"; r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = RGBColor(176, 196, 222)
    p2 = tf_m3.add_paragraph(); r = p2.add_run(); r.text = f"{max_act}\n"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    p3 = tf_m3.add_paragraph(); r = p3.add_run(); r.text = f"{len(res['gen_cols'])} disponibles"; r.font.size = Pt(8); r.font.color.rgb = RGBColor(143, 162, 183)

    # Tarjeta 4
    add_card(right_left + Inches(2.1), Inches(5.1), sub_w, sub_h, bg_card, border_card)
    tb_m4 = slide.shapes.add_textbox(right_left + Inches(2.2), Inches(5.12), sub_w - Inches(0.2), sub_h)
    tf_m4 = tb_m4.text_frame
    p1 = tf_m4.paragraphs[0]; r = p1.add_run(); r.text = "HORAS EN EVENTOS\n"; r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = RGBColor(176, 196, 222)
    p2 = tf_m4.add_paragraph(); r = p2.add_run(); r.text = f"{tot_h:.2f} h\n"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    p3 = tf_m4.add_paragraph(); r = p3.add_run(); r.text = "Total horas acumuladas"; r.font.size = Pt(8); r.font.color.rgb = RGBColor(143, 162, 183)

    # Tabla Eventos Detectados
    tx_ev = slide.shapes.add_textbox(Inches(0.5), Inches(5.85), Inches(12.333), Inches(0.3))
    tf_ev = tx_ev.text_frame; p_ev = tf_ev.paragraphs[0]; r_ev = p_ev.add_run(); r_ev.text = "EVENTOS DETECTADOS"; r_ev.font.size = Pt(12); r_ev.font.bold = True; r_ev.font.color.rgb = RGBColor(255, 255, 255)

    if not res["events"].empty:
        disp_events = res["events"].copy()
        disp_events["Inicio"] = disp_events["Inicio"].dt.strftime("%Y-%m-%d %H:%M:%S")
        disp_events["Fin"] = disp_events["Fin"].dt.strftime("%Y-%m-%d %H:%M:%S")
        
        rows = len(disp_events) + 1
        cols_cnt = len(disp_events.columns)
        
        table_shape = slide.shapes.add_table(rows, cols_cnt, Inches(0.5), Inches(6.15), Inches(12.333), Inches(0.25 * rows))
        table = table_shape.table
        
        for col_idx, col_name in enumerate(disp_events.columns):
            cell = table.cell(0, col_idx)
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(17, 42, 71)
            cell.text = str(col_name)
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(8.5); run.font.bold = True; run.font.color.rgb = RGBColor(255, 255, 255)
        
        for row_idx, row_data in disp_events.reset_index(drop=True).iterrows():
            for col_idx, val in enumerate(row_data):
                cell = table.cell(row_idx + 1, col_idx)
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(8, 20, 33)
                cell.text = str(val)
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8); run.font.color.rgb = RGBColor(234, 242, 250)

    # Footer
    tx_f = slide.shapes.add_textbox(Inches(0.5), Inches(7.05), Inches(3.0), Inches(0.3))
    tf_f = tx_f.text_frame; p_f = tf_f.paragraphs[0]; r_f = p_f.add_run(); r_f.text = "NABORS"; r_f.font.size = Pt(13); r_f.font.bold = True; r_f.font.italic = True; r_f.font.color.rgb = RGBColor(255, 255, 255)

    tx_fr = slide.shapes.add_textbox(Inches(9.0), Inches(7.08), Inches(3.8), Inches(0.3))
    tf_fr = tx_fr.text_frame; p_fr = tf_fr.paragraphs[0]; p_fr.alignment = PP_ALIGN.RIGHT
    r_fr = p_fr.add_run(); r_fr.text = "NABORS.COM     1 "; r_fr.font.size = Pt(9); r_fr.font.color.rgb = RGBColor(143, 162, 183)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io


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
            
            df_raw = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=None)
            if df_raw is not None and not df_raw.empty:
                header_row = 0
                for idx, row in df_raw.head(20).iterrows():
                    row_str = " ".join(row.dropna().astype(str)).lower()
                    if any(k in row_str for k in ["time", "timestamp", "fecha", "gen", "generator", "power"]):
                        header_row = idx
                        break
                
                df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=header_row)
                if df is not None and not df.empty:
                    result[normalize_name(sheet)] = df
        return result

    raise ValueError("Formato no soportado.")


def find_time_column(df: pd.DataFrame) -> str:
    cols = {normalize_name(c).lower(): c for c in df.columns}
    for candidate in ["time", "timestamp", "datetime", "date time", "fecha", "fecha/hora", "date_time"]:
        if candidate in cols:
            return cols[candidate]
    
    for col in df.columns:
        text = str(col).lower()
        if "timestamp" in text or "datetime" in text or "fecha" in text or text == "time":
            return col
            
    for col in df.columns:
        text = str(col).lower()
        if "time" in text:
            return col

    raise ValueError("No se encontró la columna de fecha/hora.")


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


def detect_events(
    df: pd.DataFrame,
    gen_cols: List[str],
    time_col: str,
    threshold: float,
    min_hours: float,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
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
            if (g[col] > ACTIVE_THRESHOLD).any():
                involved.append(gen_label(col))

        active_vals = g[gen_cols].values
        active_vals_filtered = active_vals[active_vals > ACTIVE_THRESHOLD]
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


def calculate_metrics(df, gen_cols, total_load_col, total_power_col):
    load_avgs = {}

    for col in gen_cols:
        m = re.search(r"(?:gen(?:erator|erador)?[\s_\-]*(\d+))", str(col), re.IGNORECASE)
        valid_series = df[df[col] > ACTIVE_THRESHOLD][col]
        if not valid_series.empty and m:
            load_avgs[int(m.group(1))] = float(valid_series.mean())

    overall_load = float(df[total_load_col].mean()) if total_load_col and total_load_col in df else 0.0
    overall_power = float(df[total_power_col].mean()) if total_power_col and total_power_col in df else 0.0

    return load_avgs, overall_load, overall_power


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
                line=dict(color=colors.get(label, "#FFFFFF"), width=1.8),
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
        annotation_font_size=10,
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
            font=dict(color="#ffffff", size=14),
            x=0.0,
        ),
        plot_bgcolor="#112a47",
        paper_bgcolor="#112a47",
        font=dict(color="#ffffff", size=10),
        legend=dict(
            font=dict(size=9, color="#ffffff"),
            bgcolor="rgba(8, 20, 33, 0.6)",
            orientation="v",
            yanchor="top",
            y=0.98,
            xanchor="right",
            x=0.98,
        ),
        xaxis=dict(
            title="Tiempo",
            tickfont=dict(color="#ffffff"),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.08)",
        ),
        yaxis=dict(
            title="Carga (%)",
            tickfont=dict(color="#ffffff"),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.08)",
            range=[0, max(55, df[gen_cols].max().max() + 5)],
        ),
        hovermode="x unified",
        margin=dict(l=35, r=15, t=40, b=35),
        height=365,
    )
    return fig


def make_fleet_chart(summary):
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
        )
    )
    fig.update_layout(
        plot_bgcolor="#112a47",
        paper_bgcolor="#112a47",
        font=dict(color="#ffffff"),
        title=dict(text="Eventos prolongados por taladro", font=dict(size=14, color="#ffffff")),
        xaxis=dict(title="Cantidad de eventos >5 h", color="#ffffff", gridcolor="rgba(255,255,255,.08)", dtick=1),
        yaxis=dict(title="Taladro", color="#ffffff", type="category"),
        margin=dict(l=30, r=30, t=35, b=30),
        height=300,
    )
    return fig


# ============================================================
# SIDEBAR
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
    st.markdown("---")
    st.caption("Criterio: ≥2 GEN activos (>0%) con carga ≤ umbral durante un tiempo continuo mayor a la duración mínima.")

uploaded = st.file_uploader(
    "📂 Cargar archivo (.xlsx, .csv)",
    type=["xlsx", "xlsm", "xls", "csv"],
)

if not uploaded:
    st.info("Cargue un archivo para visualizar el análisis.")
    st.stop()

# ============================================================
# PROCESAMIENTO
# ============================================================
try:
    rig_data = read_uploaded_file(uploaded)
except Exception as e:
    st.error(f"Error al leer el archivo: {e}")
    st.stop()

all_results = []
rig_prepared = {}

for rig, raw_df in rig_data.items():
    try:
        df, gen_cols, time_col, total_load_col, total_power_col = prepare_rig(raw_df)
        work, events = detect_events(df, gen_cols, time_col, threshold, min_hours)
        load_avgs, overall_load, overall_power = calculate_metrics(df, gen_cols, total_load_col, total_power_col)

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

        all_results.append({
            "Taladro": str(rig),
            "Registros": len(df),
            "GEN máx. activos": int(work["generadores_activos"].max()) if len(work) else 0,
            "Eventos >5h": len(events),
            "Horas en eventos >5h": round(events["Duración (h)"].sum(), 2) if not events.empty else 0.0,
            "Mayor evento (h)": round(events["Duración (h)"].max(), 2) if not events.empty else 0.0,
        })
    except Exception as e:
        st.warning(f"No se pudo procesar {rig}: {e}")

if not rig_prepared:
    st.stop()

summary = pd.DataFrame(all_results)

# ============================================================
# SELECCIÓN Y DASHBOARD
# ============================================================
col_select, col_btn = st.columns([3, 1])
with col_select:
    selected_rig = st.selectbox("Seleccionar Taladro", list(rig_prepared.keys()))

res = rig_prepared[selected_rig]

start_t = res["df"][res["time_col"]].min()
end_t = res["df"][res["time_col"]].max()
period_str = f"{start_t.strftime('%d %b')} – {end_t.strftime('%d %b')}" if pd.notna(start_t) else ""

# Botón para descargar el PowerPoint con el diseño exacto
with col_btn:
    st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True)
    pptx_file = create_powerpoint_slide(selected_rig, period_str, res, threshold, min_hours)
    st.download_button(
        label="📊 Descargar PowerPoint",
        data=pptx_file,
        file_name=f"RIG_{selected_rig}_Reporte_Generadores.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

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

# Tarjetas Superiores
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

# Sección central
c_left, c_right = st.columns([2.3, 1.0], gap="medium")

with c_left:
    fig = make_load_chart(res["df"], res["gen_cols"], res["time_col"], threshold, res["events"], selected_rig)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

with c_right:
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

    ev_df = res["events"]
    tot_h = ev_df["Duración (h)"].sum() if not ev_df.empty else 0.0
    max_h = ev_df["Duración (h)"].max() if not ev_df.empty else 0.0
    max_act = int(res["work"]["generadores_activos"].max()) if len(res["work"]) else 0

    r1_1, r1_2 = st.columns(2)
    with r1_1:
        render_card("EVENTOS >5 H", f"{len(ev_df)}", "#00A8E8", f"{tot_h:.2f} h acumuladas")
    with r1_2:
        render_card("MÁX. DURACIÓN", f"{max_h:.0f} h" if max_h else "0 h", "#00A8E8", "evento más prolongado")

    st.markdown("<div style='margin-top: 6px;'></div>", unsafe_allow_html=True)

    r2_1, r2_2 = st.columns(2)
    with r2_1:
        render_card("Max de GEN activos", f"{max_act}", "#00A8E8", f"{len(res['gen_cols'])} disponibles")
    with r2_2:
        render_card("HORAS EN EVENTOS", f"{tot_h:.2f} h", "#00A8E8", "Total horas acumuladas")

# Tablas y Resultados
st.markdown('<div class="section-title">EVENTOS DETECTADOS</div>', unsafe_allow_html=True)
if res["events"].empty:
    st.info("No se registraron eventos prolongados bajo los parámetros establecidos.")
else:
    disp_events = res["events"].copy()
    disp_events["Inicio"] = disp_events["Inicio"].dt.strftime("%Y-%m-%d %H:%M:%S")
    disp_events["Fin"] = disp_events["Fin"].dt.strftime("%Y-%m-%d %H:%M:%S")
    st.dataframe(disp_events, use_container_width=True, hide_index=True)

# Vista de Flota
st.markdown("---")
st.markdown('<div class="section-title">VISTA DE FLOTA</div>', unsafe_allow_html=True)

fc1, fc2 = st.columns([1.3, 0.9], gap="medium")
with fc1:
    st.plotly_chart(make_fleet_chart(summary), use_container_width=True, config={"displaylogo": False})
with fc2:
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    st.dataframe(summary.sort_values("Horas en eventos >5h", ascending=False), use_container_width=True, hide_index=True)

# Footer
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
        <div class="footer-right">NABORS.COM &nbsp;&nbsp;|&nbsp;&nbsp; 1</div>
    </div>
    """,
    unsafe_allow_html=True,
)
