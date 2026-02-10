"""
app.py  ─  DQ Assessment  +  Data Maturity Assessment  (Fully Integrated)
=========================================================================
Single Streamlit entry-point  |  Three pages: Home → DQ → Maturity

Project layout:
    app.py
    modules/          ← Data Quality engine (existing, unchanged)
    DataMaturity/     ← Data Maturity module
    output/
    temp/
    rules/

Run:
    pip install streamlit pandas numpy openpyxl xlsxwriter matplotlib reportlab
    streamlit run app.py
"""

# ── stdlib ────────────────────────────────────────────────────────────────
import traceback, datetime
from io import BytesIO

# ── third-party ───────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Circle, Rectangle
from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak, Image as RLImage, HRFlowable,
)
from reportlab.lib.units import inch

# ── DQ Engine modules ─────────────────────────────────────────────────────
from modules.config           import AppConfig
from modules.file_loader      import FileLoaderService
from modules.rulebook_builder import RulebookBuilderService
from modules.rule_executor    import RuleExecutorEngine
from modules.scoring_engine   import ScoringService
from modules.report_generator import ExcelReportGenerator
from modules.ui_components    import UIComponents
from modules.utils            import (setup_directories, save_uploaded_file,
                                      clean_temp_directory)

# ── DataMaturity modules ──────────────────────────────────────────────────
from DataMaturity.config import (
    UNIQU_PURPLE, UNIQU_MAGENTA, UNIQU_LAVENDER,
    UNIQU_LIGHT_BG, UNIQU_TEXT, UNIQU_GREY,
    RATING_LABELS, RATING_TO_SCORE,
    DEFAULT_MASTER_OBJECTS, MATURITY_DIMS, QUESTION_BANK,
)
from DataMaturity.helpers import (
    dq_score_to_maturity_level,
    init_maturity_state,
    build_question_df,
    sync_response_tables,
    autofill_dq_dimension,
    compute_all_scores,
    validate_responses,
    to_excel_bytes,
)
from DataMaturity.visualizations  import render_slide_png
from DataMaturity.report_generator import build_pdf_bytes

# ══════════════════════════════════════════════════════════════════════════
#  APP SETUP
# ══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="DQ & Maturity Assessment",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
setup_directories()


def _css():
    st.markdown(f"""<style>
      .stApp{{background:{UNIQU_LIGHT_BG};color:{UNIQU_TEXT}}}
      section[data-testid="stSidebar"]{{background:#fff;border-right:1px solid {UNIQU_GREY}}}
      h1,h2,h3,h4{{color:{UNIQU_PURPLE};font-family:"Segoe UI",Arial,sans-serif}}
      .stButton>button{{background:{UNIQU_PURPLE};color:#fff;border:0;
          border-radius:10px;padding:.55rem 1.1rem;font-weight:600}}
      .stButton>button:hover{{background:{UNIQU_MAGENTA};color:#fff}}
      button[kind="primary"]{{background:{UNIQU_PURPLE}!important}}
      div[data-testid="stDataFrame"]{{border:1px solid {UNIQU_GREY};
          border-radius:10px;background:#fff;padding:6px}}
      .step-card{{background:#fff;border-radius:12px;padding:20px;
          text-align:center;height:100%}}
    </style>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════
def _init_state():
    # Navigation
    if "page" not in st.session_state:
        st.session_state["page"] = "home"

    # DQ results (shared with maturity page)
    dq_keys = {
        "dq_score":        None,   # float 0-100
        "dq_dim_scores":   None,   # dict {dim: score}
        "dq_results_df":   None,   # full results DataFrame
        "dq_object_name":  "Customer",
        "dq_excel_path":   None,
    }
    for k, v in dq_keys.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # Maturity state
    init_maturity_state()

    # Extra maturity key used only in this file
    if "mat_client" not in st.session_state:
        st.session_state["mat_client"] = ""


# ══════════════════════════════════════════════════════════════════════════
#  DQ SCORE VISUALS  (gauge + dimension bars)
# ══════════════════════════════════════════════════════════════════════════
def _gauge_png(score: float) -> bytes:
    fig, ax = plt.subplots(figsize=(4.5, 2.8), dpi=130)
    fig.patch.set_facecolor(UNIQU_LIGHT_BG)
    ax.set_xlim(0, 1); ax.set_ylim(0, 0.65); ax.axis("off")

    # Background arc
    ax.add_patch(Wedge((0.5, 0.05), 0.40, 0, 180, width=0.10,
        facecolor="#ddd6f0", edgecolor="white", lw=2))

    # Coloured arc
    ang  = score / 100 * 180
    col  = "#22c55e" if score >= 80 else ("#f59e0b" if score >= 60 else "#ef4444")
    ax.add_patch(Wedge((0.5, 0.05), 0.40, 0, ang, width=0.10,
        facecolor=col, edgecolor="white", lw=2))

    ax.text(0.5, 0.30, f"{score:.1f}%",
        ha="center", va="center", fontsize=22,
        fontweight="bold", color=UNIQU_PURPLE)
    ax.text(0.5, 0.16, "Overall DQ Score",
        ha="center", va="center", fontsize=9.5, color="#555")

    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.1)
    plt.close(fig)
    return buf.getvalue()


def _dim_bar_png(dim_scores: dict) -> bytes:
    if not dim_scores:
        return None
    dims   = list(dim_scores.keys())
    scores = [dim_scores[d] for d in dims]
    cols   = ["#22c55e" if s >= 80 else ("#f59e0b" if s >= 60 else "#ef4444")
              for s in scores]

    fig, ax = plt.subplots(figsize=(7, max(2.5, len(dims) * 0.7)), dpi=120)
    fig.patch.set_facecolor(UNIQU_LIGHT_BG)
    ax.set_facecolor(UNIQU_LIGHT_BG)

    bars = ax.barh(dims, scores, color=cols, height=0.5, edgecolor="white")
    ax.set_xlim(0, 112)
    ax.set_xlabel("DQ Score (%)", color=UNIQU_TEXT, fontsize=9)
    ax.tick_params(colors=UNIQU_TEXT, labelsize=9)
    ax.spines[["top","right","bottom"]].set_visible(False)
    ax.spines["left"].set_color(UNIQU_GREY)
    ax.axvline(80, color=UNIQU_PURPLE, lw=1, ls="--", alpha=0.5, label="Good (80%)")
    ax.axvline(60, color=UNIQU_MAGENTA, lw=1, ls=":", alpha=0.5, label="Fair (60%)")
    ax.legend(fontsize=8, loc="lower right", framealpha=0.7)

    for bar, sc in zip(bars, scores):
        ax.text(bar.get_width() + 1.5, bar.get_y() + bar.get_height() / 2,
            f"{sc:.1f}%", va="center", fontsize=9,
            fontweight="bold", color=UNIQU_TEXT)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════
#  MATURITY VISUALS  (scores bar for report page)
# ══════════════════════════════════════════════════════════════════════════
def _mat_bar_png(dim_vals: dict) -> bytes:
    dims   = list(dim_vals.keys())
    scores = [dim_vals[d] for d in dims]
    cols   = [UNIQU_PURPLE if s >= 4 else (UNIQU_MAGENTA if s >= 3 else "#9a79d4")
              for s in scores]

    fig, ax = plt.subplots(figsize=(9, max(2.5, len(dims) * 0.8)), dpi=120)
    fig.patch.set_facecolor(UNIQU_LIGHT_BG)
    ax.set_facecolor(UNIQU_LIGHT_BG)

    bars = ax.barh(dims, scores, color=cols, height=0.5, edgecolor="white")
    ax.set_xlim(0, 6.0)
    ax.set_xlabel("Maturity Score (1 = Adhoc  →  5 = Optimised)",
        color=UNIQU_TEXT, fontsize=9)
    ax.axvline(3.0, color="#aaa", lw=1, ls="--", alpha=0.6, label="Defined (3)")
    ax.axvline(4.0, color=UNIQU_PURPLE, lw=1, ls="--", alpha=0.6, label="Managed (4)")
    ax.legend(fontsize=8, loc="lower right", framealpha=0.7)
    ax.tick_params(colors=UNIQU_TEXT, labelsize=9)
    ax.spines[["top","right","bottom"]].set_visible(False)
    ax.spines["left"].set_color(UNIQU_GREY)

    for bar, sc in zip(bars, scores):
        ax.text(bar.get_width() + 0.08, bar.get_y() + bar.get_height() / 2,
            f"{sc:.2f}", va="center", fontsize=9.5,
            fontweight="bold", color=UNIQU_TEXT)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════
#  COMBINED EXCEL  (DQ scores + Maturity scores in one workbook)
# ══════════════════════════════════════════════════════════════════════════
def _combined_excel(dq_score, dq_dim_scores, dim_table, overall,
                    detail_tables, low_thr, objects) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:

        # ── DQ Summary ────────────────────────────────────────
        rows = [
            {"Metric": "Overall DQ Score (%)",     "Value": f"{dq_score:.1f}%"},
            {"Metric": "Mapped Maturity Level",     "Value": dq_score_to_maturity_level(dq_score)},
        ]
        if dq_dim_scores:
            for dim, sc in dq_dim_scores.items():
                rows.append({"Metric": f"DQ – {dim}", "Value": f"{sc:.1f}%"})
        pd.DataFrame(rows).to_excel(w, sheet_name="DQ Score Summary", index=False)

        # ── Maturity dimension scores ─────────────────────────
        dim_table.to_excel(w, sheet_name="Maturity - Dimension Scores")

        # ── Overall maturity ──────────────────────────────────
        pd.DataFrame(overall).to_excel(w, sheet_name="Maturity - Overall Scores")

        # ── Detail per dimension ──────────────────────────────
        for dim, df in detail_tables.items():
            d = df.copy()
            d.insert(0, "Dimension", dim)
            d.to_excel(w, sheet_name=f"Detail - {dim[:20]}", index=False)

        # ── Exception sheets ──────────────────────────────────
        for dim, df in detail_tables.items():
            s = df.copy()
            for o in objects:
                s[o] = s[o].map(RATING_TO_SCORE).astype(float)
            for o in objects:
                exc = s[s[o] <= low_thr][
                    ["Question ID", "Section", "Question", "Weight", o]
                ].copy()
                if len(exc):
                    exc.to_excel(
                        w,
                        sheet_name=f"Exc-{o[:10]}-{dim[:8]}"[:31],
                        index=False,
                    )
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════
#  COMBINED PDF  (DQ scores + maturity slide + tables)
# ══════════════════════════════════════════════════════════════════════════
def _rl_table(df: pd.DataFrame, max_rows: int = 35) -> Table:
    d    = df.head(max_rows).copy()
    data = [list(d.columns)] + [
        [str(cell) for cell in row] for row in d.values.tolist()
    ]
    t = Table(data, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1,  0), rl_colors.HexColor(UNIQU_PURPLE)),
        ("TEXTCOLOR",      (0, 0), (-1,  0), rl_colors.white),
        ("FONTNAME",       (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",       (0, 0), (-1,  0), 9),
        ("FONTSIZE",       (0, 1), (-1, -1), 8),
        ("GRID",           (0, 0), (-1, -1), 0.25, rl_colors.lightgrey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [rl_colors.HexColor("#f9f5ff"), rl_colors.white]),
        ("VALIGN",         (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",    (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 5),
        ("TOPPADDING",     (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 3),
    ]))
    return t


def _combined_pdf(client_name, slide_png_bytes,
                  dim_table, overall, detail_tables,
                  dq_score=None, dq_dim_scores=None) -> bytes:
    buff  = BytesIO()
    doc   = SimpleDocTemplate(
        buff, pagesize=landscape(A4),
        leftMargin=.5*inch, rightMargin=.5*inch,
        topMargin=.5*inch,  bottomMargin=.5*inch,
    )
    stl = getSampleStyleSheet()
    cn  = client_name.strip() or "Client"

    title_sty = ParagraphStyle("TS", parent=stl["Title"],
        textColor=rl_colors.HexColor(UNIQU_PURPLE), spaceAfter=4)
    h2_sty = ParagraphStyle("H2", parent=stl["Heading2"],
        textColor=rl_colors.HexColor(UNIQU_PURPLE))
    note_sty = ParagraphStyle("Note", parent=stl["Italic"],
        textColor=rl_colors.HexColor("#666666"), fontSize=8)

    story = []

    # Page 1 – Slide image
    story.append(RLImage(BytesIO(slide_png_bytes),
        width=10.8*inch, height=6.1*inch))
    story.append(PageBreak())

    # Page 2 – DQ + Maturity summary
    story.append(Paragraph(f"Assessment Summary  ·  {cn}", title_sty))
    story.append(HRFlowable(width="100%", thickness=2,
        color=rl_colors.HexColor(UNIQU_PURPLE), spaceAfter=10))

    # DQ block
    if dq_score is not None:
        story.append(Paragraph("Data Quality Engine Results", h2_sty))
        story.append(Spacer(1, 6))
        dq_rows = [
            {"Metric": "Overall DQ Score (%)", "Value": f"{dq_score:.1f}%"},
            {"Metric": "Mapped Maturity Level",
             "Value": dq_score_to_maturity_level(dq_score)},
        ]
        if dq_dim_scores:
            for d, sc in dq_dim_scores.items():
                dq_rows.append({"Metric": f"DQ Dimension – {d}",
                                "Value": f"{sc:.1f}%"})
        story.append(_rl_table(pd.DataFrame(dq_rows)))
        story.append(Spacer(1, 16))

    # Maturity dimension table
    story.append(Paragraph(
        "Data Maturity – Dimension-wise Scores  (Weighted Average 1–5)", h2_sty))
    story.append(Spacer(1, 6))
    story.append(_rl_table(
        dim_table.reset_index()
                 .rename(columns={"index": "Dimension"})
                 .round(2),
        max_rows=50,
    ))
    story.append(Spacer(1, 16))

    # Overall maturity
    story.append(Paragraph(
        "Overall Maturity Score  (Average of Dimensions)", h2_sty))
    story.append(Spacer(1, 6))
    story.append(_rl_table(pd.DataFrame({
        "Master Data Object": list(overall.index),
        "Overall Score":      [round(float(v), 2) for v in overall.values],
    }), max_rows=200))
    story.append(PageBreak())

    # Pages 3+ – Per-dimension detail
    for dim, df in detail_tables.items():
        story.append(Paragraph(f"Detailed Responses  ·  {dim}", h2_sty))
        story.append(Spacer(1, 6))
        obj_cols = [c for c in df.columns
                    if c not in ["Question ID", "Section", "Question", "Weight"]]
        cmp = df[["Question ID", "Section", "Question", "Weight"] + obj_cols].copy()
        for o in obj_cols:
            cmp[o] = cmp[o].map(RATING_TO_SCORE)
        cmp["Question"] = cmp["Question"].astype(str).str.slice(0, 88)
        story.append(_rl_table(cmp, max_rows=35))
        story.append(Spacer(1, 8))
        story.append(Paragraph(
            "Scores: 1 = Adhoc  |  2 = Repeatable  |  3 = Defined  "
            "|  4 = Managed  |  5 = Optimised",
            note_sty,
        ))
        story.append(PageBreak())

    doc.build(story)
    return buff.getvalue()


# ══════════════════════════════════════════════════════════════════════════
#  PAGE 1 – HOME
# ══════════════════════════════════════════════════════════════════════════
def page_home():
    st.title("📊 Data Quality & Maturity Assessment")
    st.markdown("##### Integrated Enterprise Assessment Tool")
    st.markdown("---")

    c1, c2, c3 = st.columns(3)
    cards = [
        (c1, UNIQU_PURPLE, "🔍 Step 1", "DQ Assessment",
         "Upload data & rules → validate → get DQ score per column & dimension"),
        (c2, UNIQU_MAGENTA, "🔗 Step 2", "Auto-Populate",
         "DQ score maps to a maturity level and pre-fills the Data Quality dimension"),
        (c3, UNIQU_PURPLE, "📈 Step 3", "Maturity Report",
         "Complete questionnaire → Download slide PDF + Excel"),
    ]
    for col, clr, title, subtitle, body in cards:
        with col:
            st.markdown(
                f'<div style="background:#fff;border:2px solid {clr};'
                f'border-radius:12px;padding:20px;min-height:160px">'
                f'<h3 style="color:{clr};margin:0 0 4px 0">{title}</h3>'
                f'<b>{subtitle}</b><br><br>{body}</div>',
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # DQ result banner
    if st.session_state.dq_score is not None:
        sc  = st.session_state.dq_score
        lvl = dq_score_to_maturity_level(sc)
        bc  = "#22c55e" if sc >= 80 else ("#f59e0b" if sc >= 60 else "#ef4444")
        st.markdown(
            f'<div style="background:#fff;border-left:6px solid {bc};'
            f'border-radius:8px;padding:14px 20px;margin-bottom:16px">'
            f'✅ <b>DQ Assessment Complete</b> — '
            f'Score: <span style="color:{UNIQU_PURPLE};font-size:1.2em;'
            f'font-weight:700">{sc:.1f}%</span>'
            f' → Maturity Level: '
            f'<span style="color:{UNIQU_MAGENTA};font-weight:700">{lvl}</span>'
            f' &nbsp;|&nbsp; Object: '
            f'<b>{st.session_state.dq_object_name}</b></div>',
            unsafe_allow_html=True,
        )

    b1, b2 = st.columns(2)
    with b1:
        if st.button("🔍 Start DQ Assessment", type="primary",
                     use_container_width=True):
            st.session_state.page = "dq"; st.rerun()
    with b2:
        if st.button("📈 Go to Maturity Assessment",
                     use_container_width=True):
            st.session_state.page = "maturity"; st.rerun()

    st.markdown("---")
    st.markdown("#### DQ Score → Maturity Level Mapping")
    st.dataframe(
        pd.DataFrame([
            {"DQ Score Range": "95% – 100%", "Maturity Level": "⭐ Optimised",
             "Description": "Continuous improvement, proactive governance"},
            {"DQ Score Range": "80% – 94%",  "Maturity Level": "✅ Managed",
             "Description": "Monitored, measured, formalized roles"},
            {"DQ Score Range": "60% – 79%",  "Maturity Level": "🔵 Defined",
             "Description": "Standardized and consistently followed"},
            {"DQ Score Range": "40% – 59%",  "Maturity Level": "🟡 Repeatable",
             "Description": "Some processes defined but inconsistent"},
            {"DQ Score Range": "0%  – 39%",  "Maturity Level": "🔴 Adhoc",
             "Description": "Unstructured, reactive, varies widely"},
        ]),
        use_container_width=True,
        hide_index=True,
    )


# ══════════════════════════════════════════════════════════════════════════
#  PAGE 2 – DQ ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════
def page_dq():
    if st.button("🏠 Home"):
        st.session_state.page = "home"; st.rerun()

    UIComponents.render_header()
    UIComponents.render_sidebar()
    st.markdown("---")
    st.subheader("📤 Upload Files")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**📊 Master Dataset**")
        data_file = st.file_uploader(
            "Upload your data file",
            type=AppConfig.SUPPORTED_DATA_FORMATS,
            key="dq_data_upload",
        )
        object_name = st.session_state.dq_object_name
        if data_file:
            object_name = st.text_input(
                "Master Data Object Name",
                value=st.session_state.dq_object_name,
                help="e.g. Customer, Vendor – carried into Maturity Assessment",
            )

    with col2:
        st.markdown("**📋 Rules Configuration**")
        utype = st.radio(
            "Upload Type",
            ["Rules Dataset (CSV/Excel)", "JSON Rulebook"],
            horizontal=True,
        )
        if utype == "Rules Dataset (CSV/Excel)":
            rules_file = st.file_uploader(
                "Upload rules file",
                type=AppConfig.SUPPORTED_RULES_FORMATS,
                key="dq_rules_csv",
            )
        else:
            rules_file = st.file_uploader(
                "Upload JSON rulebook",
                type=["json"],
                key="dq_rules_json",
            )

    UIComponents.render_file_format_help()
    st.markdown("---")

    if not (data_file and rules_file):
        UIComponents.render_welcome_screen()
        st.markdown("---")
        w1, w2, w3 = st.columns(3)
        with w1: st.success("**Step 1** – DQ Assessment")
        with w2: st.info("**Step 2** – Auto-Populate Maturity")
        with w3: st.warning("**Step 3** – Download PDF & Excel")
        st.markdown("---")
        UIComponents.render_footer()
        return

    if not st.button("🚀 Run Data Quality Check", type="primary",
                     use_container_width=True):
        return

    try:
        clean_temp_directory()
        pb   = st.progress(0)
        stxt = st.empty()

        stxt.text("📂 Saving files…");          pb.progress(5)
        dp = save_uploaded_file(data_file,  AppConfig.TEMP_DIR)
        rp = save_uploaded_file(rules_file, AppConfig.TEMP_DIR)

        stxt.text("📊 Loading dataset…");        pb.progress(15)
        loader = FileLoaderService()
        df     = loader.load_dataframe(dp)
        cols   = df.columns.tolist()
        st.info(f"✅ Loaded **{len(df):,}** records · **{len(cols)}** columns")

        stxt.text("🔧 Building rulebook…");      pb.progress(30)
        rb = RulebookBuilderService()
        if utype == "JSON Rulebook":
            rulebook = rb.load_json_rulebook(rp)
        else:
            rulebook = rb.build_from_rules_dataset(
                loader.load_dataframe(rp), cols)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📊 Records",  f"{len(df):,}")
        m2.metric("📋 Columns",  len(cols))
        m3.metric("⚙️ Rules",    len(rulebook.get("rules", [])))
        m4.metric("🎯 Targets",  len({
            r.get("column") or "+".join(r.get("columns", []))
            for r in rulebook["rules"]
        }))

        stxt.text("✅ Executing rules…");        pb.progress(50)
        exe    = RuleExecutorEngine(df, rulebook)
        res    = exe.execute_all_rules()
        combos = exe.get_combination_duplicates()

        stxt.text("📊 Calculating scores…");     pb.progress(70)
        svc      = ScoringService()
        overall  = svc.calculate_overall_score(res)
        col_sc   = svc.calculate_column_scores(res, cols)
        dim_sc   = svc.calculate_dimension_scores(res)

        stxt.text("💾 Generating Excel…");       pb.progress(85)
        rgen = ExcelReportGenerator(
            results_df=res, rulebook=rulebook, all_columns=cols,
            column_scores=col_sc, overall_score=overall,
            dimension_scores=dim_sc, duplicate_combinations=combos,
        )
        out_path = rgen.generate_report(AppConfig.OUTPUT_DIR)
        rb_path  = rgen.save_rulebook_json(AppConfig.OUTPUT_DIR, rulebook)
        pb.progress(100); stxt.text("✅ Complete!")

        # ── Persist for Maturity ───────────────────────────────
        st.session_state.dq_score       = overall
        st.session_state.dq_dim_scores  = dim_sc
        st.session_state.dq_results_df  = res
        st.session_state.dq_object_name = object_name or "Customer"
        st.session_state.dq_excel_path  = out_path
        # Pre-set maturity object
        st.session_state.mat_objects = (
            [object_name] if object_name else DEFAULT_MASTER_OBJECTS[:]
        )
        # Autofill DQ dimension immediately
        autofill_dq_dimension(overall)

        st.success("✅ Data Quality Check Completed Successfully!")
        st.markdown("---")

        # ── Visual dashboard ───────────────────────────────────
        st.subheader("📊 DQ Results Dashboard")
        g1, g2 = st.columns([1, 2])
        with g1:
            st.image(_gauge_png(overall), use_container_width=True)
        with g2:
            if dim_sc:
                bar = _dim_bar_png(dim_sc)
                if bar:
                    st.image(bar, use_container_width=True)

        st.markdown("---")
        UIComponents.render_results_dashboard(overall, res, col_sc, dim_sc)
        st.markdown("---")

        # Annexure info
        a1, a2, a3 = st.columns(3)
        with a1:
            st.info("**🔍 Uniqueness Issues**\n"
                    "Duplicate & combination duplicate records")
        with a2:
            st.info("**📝 Completeness Issues**\n"
                    "Null / missing value records")
        with a3:
            st.info("**📐 Standardization Issues**\n"
                    "Format & pattern violation records")
        st.markdown("---")

        UIComponents.render_download_section(out_path, rb_path, len(cols) + 3)
        st.markdown("---")
        UIComponents.render_detailed_views(rulebook, res, col_sc, dim_sc)

        # ── Continue to Maturity ───────────────────────────────
        st.markdown("---")
        st.subheader("📈 Continue to Data Maturity Assessment")
        lvl = dq_score_to_maturity_level(overall)
        st.markdown(
            f'<div style="background:{UNIQU_LAVENDER};border-left:5px solid '
            f'{UNIQU_PURPLE};border-radius:8px;padding:14px 20px">'
            f'🔗 DQ Score <b>{overall:.1f}%</b> has been mapped to maturity level '
            f'<b style="color:{UNIQU_PURPLE}">{lvl}</b>. '
            f'The <i>Data Quality</i> dimension is already '
            f'auto-filled. Click below to proceed.</div>',
            unsafe_allow_html=True,
        )
        st.markdown("")
        if st.button("📈 Continue to Maturity Assessment →",
                     type="primary", use_container_width=True):
            st.session_state.page = "maturity"
            st.rerun()

    except Exception as e:
        st.error(f"❌ Error: {e}")
        with st.expander("🔍 Full traceback"):
            st.code(traceback.format_exc())

    st.markdown("---")
    UIComponents.render_footer()


# ══════════════════════════════════════════════════════════════════════════
#  PAGE 3 – MATURITY ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════
def _do_submit():
    objects   = st.session_state.mat_objects
    dims      = st.session_state.mat_dims
    responses = st.session_state.mat_responses
    cn        = st.session_state.mat_client or "Client"
    bm        = float(st.session_state.mat_benchmark)
    tg        = float(st.session_state.mat_target)
    lt        = float(st.session_state.mat_low_thr)

    ok, msg = validate_responses(responses, dims, objects)
    if not ok:
        st.error(msg); return

    with st.spinner("⚙️ Computing scores and building all reports…"):
        dim_table, overall = compute_all_scores(objects, dims, responses)
        detail_tables      = {d: responses[d].copy() for d in dims}
        exec_score         = (float(np.nanmean(overall.values))
                              if len(overall) else 0.0)

        # Per-dimension raw averages
        raw = {
            dim: float(np.nanmean(dim_table.loc[dim].values))
            for dim in dims
        }
        for d in MATURITY_DIMS:
            if d not in raw:
                raw[d] = np.nan

        # Display labels for the slide's right-panel table
        domain_display = {
            "Data Governance":
                raw.get("Data Governance", np.nan),
            "Data Quality":
                raw.get("Data Quality", np.nan),
            "Data Integration and\nInteroperability":
                raw.get("Data Integration & Interoperability", np.nan),
        }

        dq_score   = st.session_state.dq_score
        dq_dim_sc  = st.session_state.dq_dim_scores

        # ── Build all output artefacts ─────────────────────────
        slide_png = render_slide_png(
            client_name=cn,
            domain_scores=domain_display,
            exec_score=exec_score if np.isfinite(exec_score) else 0.0,
            benchmark=bm,
            target=tg,
        )

        mat_excel_bytes = to_excel_bytes(
            dim_table, overall, detail_tables, lt, objects
        )

        combined_excel_bytes = _combined_excel(
            dq_score, dq_dim_sc,
            dim_table, overall, detail_tables, lt, objects,
        )

        pdf_bytes = _combined_pdf(
            client_name=cn,
            slide_png_bytes=slide_png,
            dim_table=dim_table.round(2),
            overall=overall.round(2),
            detail_tables=detail_tables,
            dq_score=dq_score,
            dq_dim_scores=dq_dim_sc,
        )

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    st.session_state.mat_payload = {
        "dim_table":      dim_table,
        "overall":        overall,
        "slide_png":      slide_png,
        "mat_excel":      mat_excel_bytes,
        "combined_excel": combined_excel_bytes,
        "pdf_bytes":      pdf_bytes,
        "exec_score":     exec_score,
        "client_name":    cn,
        "ts":             ts,
    }
    st.session_state.mat_submitted = True
    st.rerun()


def page_maturity():
    dq_score = st.session_state.dq_score
    disabled = st.session_state.mat_submitted

    # ── Sidebar ───────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Configuration")

        st.session_state.mat_client = st.text_input(
            "Client Name",
            value=st.session_state.mat_client,
            placeholder="Organisation name",
            disabled=disabled,
        )

        all_opts = list(dict.fromkeys(
            DEFAULT_MASTER_OBJECTS + st.session_state.mat_objects
        ))
        st.session_state.mat_objects = st.multiselect(
            "Master Data Objects",
            options=all_opts,
            default=st.session_state.mat_objects,
            disabled=disabled,
        )

        st.session_state.mat_dims = st.multiselect(
            "Maturity Dimensions",
            options=MATURITY_DIMS,
            default=st.session_state.mat_dims,
            disabled=disabled,
        )

        st.divider()
        st.header("📊 Scoring")
        st.session_state.mat_low_thr = st.slider(
            "Exception threshold (≤)",
            1.0, 5.0, float(st.session_state.mat_low_thr), 0.5,
            disabled=disabled,
        )

        st.divider()
        st.header("🎯 Benchmark / Target")
        st.session_state.mat_benchmark = st.number_input(
            "Industry Benchmark", 1.0, 5.0,
            float(st.session_state.mat_benchmark), 0.1,
            disabled=disabled,
        )
        st.session_state.mat_target = st.number_input(
            "Target", 1.0, 5.0,
            float(st.session_state.mat_target), 0.1,
            disabled=disabled,
        )

        st.divider()
        if st.button("🔍 Back to DQ Assessment"):
            st.session_state.page = "dq"; st.rerun()
        if st.button("🏠 Home"):
            st.session_state.page = "home"; st.rerun()

    # Guards
    if not st.session_state.mat_objects or not st.session_state.mat_dims:
        st.info("Select at least one Master Object and "
                "Dimension in the sidebar.")
        st.stop()

    sync_response_tables()

    # ════════════════════════════════════════════════════════
    # REPORT VIEW  (after submit)
    # ════════════════════════════════════════════════════════
    if st.session_state.mat_submitted:
        p  = st.session_state.mat_payload
        cn = p["client_name"]
        ts = p["ts"]

        st.title("📈 Data Maturity Assessment Report")

        # DQ banner
        if dq_score is not None:
            lvl = dq_score_to_maturity_level(dq_score)
            st.markdown(
                f'<div style="background:{UNIQU_LAVENDER};border-left:6px solid '
                f'{UNIQU_PURPLE};border-radius:8px;padding:12px 20px;'
                f'margin-bottom:16px">'
                f'🔗 <b>DQ Engine</b> — Score: '
                f'<span style="color:{UNIQU_PURPLE};font-weight:700;'
                f'font-size:1.15em">{dq_score:.1f}%</span>'
                f' → Maturity: '
                f'<span style="color:{UNIQU_MAGENTA};font-weight:700">{lvl}'
                f'</span> &nbsp;(auto-applied to <i>Data Quality</i> dimension)'
                f'</div>',
                unsafe_allow_html=True,
            )

        # ── Slide image ───────────────────────────────────────
        st.markdown("### 📊 Summary Slide")
        st.image(p["slide_png"], use_container_width=True)
        st.divider()

        # ── Dimension + Overall tables ────────────────────────
        t1, t2 = st.columns(2)
        with t1:
            st.markdown("#### Dimension-wise Maturity (1–5)")
            st.dataframe(
                p["dim_table"].style
                  .format("{:.2f}")
                  .background_gradient(cmap="RdYlGn", axis=None,
                                       vmin=1, vmax=5),
                use_container_width=True,
            )
        with t2:
            st.markdown("#### Overall Maturity Score")
            st.dataframe(
                pd.DataFrame(p["overall"]).T.style
                  .format("{:.2f}")
                  .background_gradient(cmap="RdYlGn", axis=None,
                                       vmin=1, vmax=5),
                use_container_width=True,
            )

        st.divider()

        # ── Maturity bar chart ────────────────────────────────
        st.markdown("#### Scores by Dimension")
        dim_vals = {
            dim: float(np.nanmean(p["dim_table"].loc[dim].values))
            for dim in p["dim_table"].index
        }
        bar_img = _mat_bar_png(dim_vals)
        if bar_img:
            st.image(bar_img, use_container_width=True)

        st.divider()

        # ── Downloads ─────────────────────────────────────────
        st.markdown("### 📥 Download Reports")
        safe_cn = cn.replace(" ", "_")

        d1, d2, d3 = st.columns(3)

        with d1:
            st.markdown(
                f'<div style="background:#fff;border:2px solid {UNIQU_PURPLE};'
                f'border-radius:10px;padding:16px;text-align:center">'
                f'<h4 style="color:{UNIQU_PURPLE};margin:0">📄 PDF Report</h4>'
                f'<p style="font-size:.85rem;color:#666;margin:6px 0 12px">Slide visual + '
                f'DQ summary + dimension tables + detailed responses</p></div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                "📄 Download PDF",
                data=p["pdf_bytes"],
                file_name=f"Maturity_Report_{safe_cn}_{ts}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        with d2:
            st.markdown(
                f'<div style="background:#fff;border:2px solid {UNIQU_MAGENTA};'
                f'border-radius:10px;padding:16px;text-align:center">'
                f'<h4 style="color:{UNIQU_MAGENTA};margin:0">📊 Maturity Excel</h4>'
                f'<p style="font-size:.85rem;color:#666;margin:6px 0 12px">Dimension scores + '
                f'overall scores + exception sheets</p></div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                "📊 Download Maturity Excel",
                data=p["mat_excel"],
                file_name=f"Maturity_Assessment_{safe_cn}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with d3:
            st.markdown(
                f'<div style="background:#fff;border:2px solid #7c4dbb;'
                f'border-radius:10px;padding:16px;text-align:center">'
                f'<h4 style="color:#7c4dbb;margin:0">🔗 Combined Excel</h4>'
                f'<p style="font-size:.85rem;color:#666;margin:6px 0 12px">DQ scores + '
                f'maturity scores + exceptions in one workbook</p></div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                "🔗 Download Combined Excel",
                data=p["combined_excel"],
                file_name=f"DQ_Maturity_Combined_{safe_cn}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.markdown("---")
        if st.button("✏️ Edit Responses (Back to Questionnaire)",
                     use_container_width=True):
            st.session_state.mat_submitted = False
            st.session_state.mat_payload   = {}
            st.rerun()
        st.stop()

    # ════════════════════════════════════════════════════════
    # QUESTIONNAIRE VIEW
    # ════════════════════════════════════════════════════════
    st.title("📈 Data Maturity Assessment")

    if dq_score is not None:
        lvl = dq_score_to_maturity_level(dq_score)
        st.success(
            f"✅ DQ Score: **{dq_score:.1f}%** → "
            f"Maturity level **'{lvl}'** has been auto-filled "
            f"in the *Data Quality* tab below."
        )
    else:
        st.info(
            "💡 Tip: Run the **DQ Assessment** first to auto-populate "
            "the *Data Quality* dimension scores."
        )

    st.caption(
        "Rate each question for every master data object, then click **Submit**."
    )
    st.markdown("---")

    tabs = st.tabs(st.session_state.mat_dims)
    for i, dim in enumerate(st.session_state.mat_dims):
        with tabs[i]:
            st.markdown(f"### {dim}")

            if dim == "Data Quality" and dq_score is not None:
                lvl = dq_score_to_maturity_level(dq_score)
                st.markdown(
                    f'<div style="background:{UNIQU_LAVENDER};border-radius:6px;'
                    f'padding:8px 14px;margin-bottom:10px">'
                    f'ℹ️ Auto-populated from DQ Score <b>{dq_score:.1f}%</b> → '
                    f'<b>{lvl}</b>.  You may still adjust individual ratings.</div>',
                    unsafe_allow_html=True,
                )

            df  = st.session_state.mat_responses[dim].copy()
            cfg = {
                "Weight": st.column_config.NumberColumn(
                    "Weight", min_value=0.0, step=0.5)
            }
            for obj in st.session_state.mat_objects:
                cfg[obj] = st.column_config.SelectboxColumn(
                    obj, options=RATING_LABELS, required=True)

            edited = st.data_editor(
                df,
                use_container_width=True,
                hide_index=True,
                column_config=cfg,
                disabled=["Question ID", "Section", "Question"],
                key=f"mat_editor_{dim}",
            )
            st.session_state.mat_responses[dim] = edited

    st.markdown("---")
    left, right = st.columns([1, 3])
    with left:
        if st.button("🚀 Submit & Generate Report", type="primary"):
            _do_submit()
    with right:
        st.info(
            "Clicking **Submit** computes all scores, builds the "
            "slide-style PNG, and enables **PDF + 2× Excel** downloads."
        )


# ══════════════════════════════════════════════════════════════════════════
#  ROUTER
# ══════════════════════════════════════════════════════════════════════════
_css()
_init_state()

{
    "home":     page_home,
    "dq":       page_dq,
    "maturity": page_maturity,
}[st.session_state.page]()