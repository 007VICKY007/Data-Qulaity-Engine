# ── stdlib ────────────────────────────────────────────────────────────────
import traceback, datetime
from io import BytesIO
from pathlib import Path

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


def load_css():
    """Load external CSS file"""
    css_file = Path(__file__).parent / "assets" / "styles.css"
    if css_file.exists():
        with open(css_file, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    else:
        st.warning("CSS file not found. Using default styles.")


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
            f"{sc:.1f}%", va="center", fontsize=9.5,
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
#  COMBINED PDF BUILDER
# ══════════════════════════════════════════════════════════════════════════
def _build_combined_pdf(
    client_name: str,
    dq_score:    float,
    dq_dim:      dict,
    dq_results:  pd.DataFrame,
    mat_slide:   bytes,
    mat_dim:     pd.DataFrame,
    mat_overall: pd.Series,
    mat_detail:  dict,
) -> bytes:
    """
    Single PDF:
      1) DQ Summary Page
      2) Maturity Slide (full-page image)
      3) Maturity summary table
      4) Maturity per-dimension pages
    """
    buff = BytesIO()
    doc  = SimpleDocTemplate(
        buff, pagesize=landscape(A4),
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.5  * inch, bottomMargin=0.5 * inch,
    )
    styles = getSampleStyleSheet()
    story  = []
    cn     = client_name.strip() or "Client"

    # ──────────────────────────────────────────────────────────
    # PAGE 1 : DQ Engine Summary
    # ──────────────────────────────────────────────────────────
    story.append(Paragraph(f"Data Quality Assessment – {cn}", styles["Title"]))
    story.append(Spacer(1, 12))

    # Gauge
    gauge_img = _gauge_png(dq_score)
    story.append(RLImage(BytesIO(gauge_img), width=3.5 * inch, height=2.2 * inch))
    story.append(Spacer(1, 12))

    # Dimension scores
    if dq_dim:
        story.append(Paragraph("DQ Scores by Dimension", styles["Heading2"]))
        story.append(Spacer(1, 6))
        dim_df = pd.DataFrame(list(dq_dim.items()), columns=["Dimension", "DQ Score (%)"])
        dim_data = [list(dim_df.columns)] + dim_df.values.tolist()
        t = Table(dim_data, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",     (0, 0), (-1,  0), rl_colors.HexColor(UNIQU_PURPLE)),
            ("TEXTCOLOR",      (0, 0), (-1,  0), rl_colors.white),
            ("FONTNAME",       (0, 0), (-1,  0), "Helvetica-Bold"),
            ("FONTSIZE",       (0, 0), (-1,  0), 10),
            ("FONTSIZE",       (0, 1), (-1, -1), 9),
            ("GRID",           (0, 0), (-1, -1), 0.25, rl_colors.lightgrey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [rl_colors.whitesmoke, rl_colors.white]),
            ("VALIGN",         (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(t)
        story.append(Spacer(1, 12))

    # Sample records (first 10)
    story.append(Paragraph("DQ Results Sample (First 10 Records)", styles["Heading2"]))
    story.append(Spacer(1, 6))
    sample = dq_results.head(10)
    display_cols = [c for c in sample.columns if not c.startswith("_")]
    sample_disp  = sample[display_cols].copy()

    for col in sample_disp.columns:
        sample_disp[col] = sample_disp[col].astype(str).str.slice(0, 40)

    data_sample = [list(sample_disp.columns)] + sample_disp.values.tolist()
    ts = Table(data_sample, repeatRows=1)
    ts.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1,  0), rl_colors.HexColor(UNIQU_PURPLE)),
        ("TEXTCOLOR",      (0, 0), (-1,  0), rl_colors.white),
        ("FONTNAME",       (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",       (0, 0), (-1,  0), 8),
        ("FONTSIZE",       (0, 1), (-1, -1), 7),
        ("GRID",           (0, 0), (-1, -1), 0.25, rl_colors.lightgrey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [rl_colors.whitesmoke, rl_colors.white]),
        ("VALIGN",         (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(ts)
    story.append(PageBreak())

    # ──────────────────────────────────────────────────────────
    # PAGE 2 : Maturity Slide
    # ──────────────────────────────────────────────────────────
    story.append(RLImage(BytesIO(mat_slide), width=10.8 * inch, height=6.1 * inch))
    story.append(PageBreak())

    # ──────────────────────────────────────────────────────────
    # PAGE 3 : Maturity Numeric Summary
    # ──────────────────────────────────────────────────────────
    story.append(Paragraph(f"Maturity Assessment Summary – {cn}", styles["Title"]))
    story.append(Spacer(1, 10))

    # Optional: DQ link
    if dq_score is not None:
        lvl = dq_score_to_maturity_level(dq_score)
        story.append(Paragraph("🔗 DQ Engine Input", styles["Heading2"]))
        story.append(Spacer(1, 4))
        dq_df = pd.DataFrame({
            "Metric": ["DQ Overall Score (%)", "Mapped Maturity Level"],
            "Value":  [f"{dq_score:.1f}%", lvl],
        })
        dq_data = [list(dq_df.columns)] + dq_df.values.tolist()
        t_dq = Table(dq_data)
        t_dq.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor(UNIQU_PURPLE)),
            ("TEXTCOLOR",  (0, 0), (-1, 0), rl_colors.white),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",   (0, 0), (-1, 0), 9),
            ("GRID",       (0, 0), (-1, -1), 0.25, rl_colors.lightgrey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [rl_colors.whitesmoke, rl_colors.white]),
        ]))
        story.append(t_dq)
        story.append(Spacer(1, 12))

    # Dimension table
    story.append(Paragraph("Dimension-wise Maturity (Weighted Average 1–5)", styles["Heading2"]))
    story.append(Spacer(1, 6))
    dim_df = mat_dim.reset_index().rename(columns={"index": "Dimension"}).round(2)
    dim_data = [list(dim_df.columns)] + dim_df.values.tolist()
    t_dim = Table(dim_data, repeatRows=1)
    t_dim.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor(UNIQU_PURPLE)),
        ("TEXTCOLOR",  (0, 0), (-1, 0), rl_colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, 0), 9),
        ("FONTSIZE",   (0, 1), (-1, -1), 8),
        ("GRID",       (0, 0), (-1, -1), 0.25, rl_colors.lightgrey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [rl_colors.whitesmoke, rl_colors.white]),
        ("VALIGN",     (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t_dim)
    story.append(Spacer(1, 12))

    # Overall table
    story.append(Paragraph("Overall Maturity (Average of Dimensions)", styles["Heading2"]))
    story.append(Spacer(1, 6))
    ov_df = pd.DataFrame({
        "Master Data Object": list(mat_overall.index),
        "Overall Score":      [round(float(v), 2) for v in mat_overall.values],
    })
    ov_data = [list(ov_df.columns)] + ov_df.values.tolist()
    t_ov = Table(ov_data)
    t_ov.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor(UNIQU_PURPLE)),
        ("TEXTCOLOR",  (0, 0), (-1, 0), rl_colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, 0), 9),
        ("GRID",       (0, 0), (-1, -1), 0.25, rl_colors.lightgrey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [rl_colors.whitesmoke, rl_colors.white]),
    ]))
    story.append(t_ov)
    story.append(PageBreak())

    # ──────────────────────────────────────────────────────────
    # PAGES 4+ : Per-dimension detailed responses
    # ──────────────────────────────────────────────────────────
    for dim, df in mat_detail.items():
        story.append(Paragraph(f"Detailed Responses – {dim}", styles["Heading2"]))
        story.append(Spacer(1, 6))

        obj_cols = [c for c in df.columns
                    if c not in ["Question ID", "Section", "Question", "Weight"]]
        compact  = df[["Question ID", "Section", "Question", "Weight"] + obj_cols].copy()

        # Convert rating labels → numeric scores
        for o in obj_cols:
            compact[o] = compact[o].map(RATING_TO_SCORE)

        # Truncate long question text
        compact["Question"] = compact["Question"].astype(str).str.slice(0, 95)

        data_det = [list(compact.columns)] + compact.head(35).values.tolist()
        t_det = Table(data_det, repeatRows=1)
        t_det.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor(UNIQU_PURPLE)),
            ("TEXTCOLOR",  (0, 0), (-1, 0), rl_colors.white),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",   (0, 0), (-1, 0), 9),
            ("FONTSIZE",   (0, 1), (-1, -1), 8),
            ("GRID",       (0, 0), (-1, -1), 0.25, rl_colors.lightgrey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [rl_colors.whitesmoke, rl_colors.white]),
            ("VALIGN",     (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",  (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t_det)
        story.append(Spacer(1, 8))
        story.append(Paragraph(
            "Note: Answers shown as numeric scores (1 = Adhoc … 5 = Optimised).",
            styles["Italic"],
        ))
        story.append(PageBreak())

    doc.build(story)
    return buff.getvalue()


# ══════════════════════════════════════════════════════════════════════════
#  PAGE: HOME
# ══════════════════════════════════════════════════════════════════════════
def page_home():
    st.markdown('<h1 class="purple-text">📊 Enterprise DQ & Maturity Platform</h1>',
                unsafe_allow_html=True)
    st.markdown("**One integrated platform** for Data Quality validation and DAMA Maturity assessment.")
    st.divider()

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(
            '<div class="step-card">'
            '<h2>🔍 DQ Assessment</h2>'
            '<p>Upload data + rulebook → get DQ score, dimension breakdown, '
            'column annexures, and automated maturity mapping.</p></div>',
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Go to DQ Assessment →", use_container_width=True):
            st.session_state["page"] = "dq"
            st.rerun()

    with c2:
        st.markdown(
            '<div class="step-card">'
            '<h2>📈 Maturity Assessment</h2>'
            '<p>Structured DAMA questionnaire → Slide-style visuals + PDF + '
            'Excel. Optionally auto-populate from DQ score.</p></div>',
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Go to Maturity Assessment →", use_container_width=True):
            st.session_state["page"] = "maturity"
            st.rerun()

    st.divider()
    st.markdown("### ✨ Key Features")
    st.markdown("""
    - **DQ Engine**: Rules-driven validation with comprehensive dimension analysis
    - **Maturity Assessment**: DAMA-aligned questionnaire with automatic scoring
    - **Integrated Workflow**: DQ score auto-maps to maturity level
    - **Professional Outputs**: PDF reports + Excel workbooks + slide visuals
    """)

    st.divider()
    st.markdown('<p class="text-center">Select an assessment above to begin.</p>',
                unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  PAGE: DQ ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════
def page_dq():
    with st.sidebar:
        st.markdown("### 🧭 Navigation")
        if st.button("← Back to Home", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
        st.divider()
        st.markdown("**Current Page:** DQ Assessment")

    st.markdown('<h1 class="purple-text">🔍 Data Quality Assessment</h1>',
                unsafe_allow_html=True)
    st.caption("Upload master dataset + rules → Generate comprehensive DQ reports")
    st.divider()

    # File uploads
    c1, c2 = st.columns(2)
    with c1:
        data_file = st.file_uploader(
            "📁 Master Dataset",
            type=["csv", "xlsx", "xls", "xlsm", "json"],
            help="Primary data file for validation"
        )
    with c2:
        rules_file = st.file_uploader(
            "📋 Rules Configuration",
            type=["csv", "xlsx", "xls", "json"],
            help="Rulebook (CSV/Excel) or JSON rulebook"
        )

    if not data_file or not rules_file:
        st.info("👆 Please upload both files to proceed")
        return

    object_name = st.text_input(
        "🏷️ Master Data Object Name",
        value=st.session_state.get("dq_object_name", "Customer"),
        help="Used in maturity auto-fill"
    )
    st.session_state["dq_object_name"] = object_name

    sheet_name = None
    if data_file.name.endswith((".xlsx", ".xls", ".xlsm")):
        loader    = FileLoaderService()
        temp_path = AppConfig.TEMP_DIR / data_file.name
        temp_path.write_bytes(data_file.getbuffer())
        sheets    = loader.get_sheet_names(temp_path)
        if len(sheets) > 1:
            sheet_name = st.selectbox("📄 Select Sheet", sheets)

    run_btn = st.button("🚀 Run DQ Assessment", type="primary", use_container_width=True)
    st.divider()

    if not run_btn:
        return

    try:
        with st.spinner("Processing..."):
            loader   = FileLoaderService()
            rb_build = RulebookBuilderService()

            # Save files
            data_path  = save_uploaded_file(data_file, AppConfig.TEMP_DIR)
            rules_path = save_uploaded_file(rules_file, AppConfig.TEMP_DIR)

            # Load data
            df_master = loader.load_dataframe(data_path, sheet_name=sheet_name)
            st.success(f"✅ Loaded {len(df_master):,} records × {len(df_master.columns)} columns")

            # Build rulebook
            if rules_file.name.lower().endswith(".json"):
                rulebook = rb_build.load_json_rulebook(rules_path)
            else:
                df_rules = loader.load_dataframe(rules_path)
                rulebook = rb_build.build_from_rules_dataset(df_rules, list(df_master.columns))

            st.success(f"✅ Rulebook: {len(rulebook['rules'])} rules")

            # Execute rules
            executor = RuleExecutorEngine(df_master, rulebook)
            results  = executor.execute_all_rules()

            # Scoring
            scorer       = ScoringService()
            overall      = scorer.calculate_overall_score(results)
            col_scores   = scorer.calculate_column_scores(results, list(df_master.columns))
            dim_scores   = scorer.calculate_dimension_scores(results)

            # Excel report
            ts      = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            rg      = ExcelReportGenerator()
            xl_path = rg.generate_report(
                results_df=results,
                overall_score=overall,
                column_scores=col_scores,
                dimension_scores=dim_scores,
                rulebook=rulebook,
                master_data_columns=list(df_master.columns),
                output_filename=f"DQ_Report_{object_name}_{ts}.xlsx"
            )

            # Save to session
            st.session_state["dq_score"]      = overall
            st.session_state["dq_dim_scores"] = dim_scores
            st.session_state["dq_results_df"] = results
            st.session_state["dq_excel_path"] = xl_path

            st.success("✅ DQ Assessment Complete!")

    except Exception as e:
        st.error(f"❌ Error: {e}")
        with st.expander("🔍 Detailed Error"):
            st.code(traceback.format_exc())
        return

    # ═══════════════════════════════════════════════════════════
    # RESULTS DISPLAY
    # ═══════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 📊 DQ Results")

    # Gauge + dimension bars
    g1, g2 = st.columns([1, 2])
    with g1:
        gauge_img = _gauge_png(overall)
        st.image(gauge_img, use_container_width=True)
    with g2:
        if dim_scores:
            dim_img = _dim_bar_png(dim_scores)
            if dim_img:
                st.image(dim_img, use_container_width=True)

    st.divider()

    # Metrics
    clean_count = len(results[results["Count of issues"] == 0])
    issue_count = len(results) - clean_count
    pass_cols   = sum(1 for s in col_scores.values() if s == 100)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Overall DQ Score", f"{overall:.1f}%")
    m2.metric("Clean Records", f"{clean_count:,}")
    m3.metric("Records with Issues", f"{issue_count:,}")
    m4.metric("Columns at 100%", f"{pass_cols}/{len(col_scores)}")

    st.divider()

    # Download
    st.markdown("### 📥 Download Reports")
    st.download_button(
        "📊 Download DQ Excel Report",
        data=open(xl_path, "rb"),
        file_name=xl_path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()

    # Preview
    st.markdown("### 🔍 Results Preview")
    display_cols = [c for c in results.columns if not c.startswith("_")]
    issues_df    = results[results["Count of issues"] > 0]

    if len(issues_df) > 0:
        st.write(f"**Records with Issues: {len(issues_df):,}**")
        st.dataframe(issues_df[display_cols].head(50), use_container_width=True)
    else:
        st.success("🎉 No issues found!")
        st.dataframe(results[display_cols].head(20), use_container_width=True)

    st.divider()

    # Link to maturity
    if st.button("📈 Continue to Maturity Assessment →", use_container_width=True):
        st.session_state["page"] = "maturity"
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════
#  PAGE: MATURITY ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════
def page_maturity():
    with st.sidebar:
        st.markdown("### 🧭 Navigation")
        if st.button("← Back to Home", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
        st.divider()
        st.markdown("**Current Page:** Maturity Assessment")

        st.divider()
        st.markdown("### ⚙️ Configuration")

        cn = st.text_input("🏢 Client Name", value=st.session_state.get("mat_client", ""))
        st.session_state["mat_client"] = cn

        st.markdown("**Master Data Objects**")
        new_obj = st.text_input("Add new object:", key="mat_new_obj")
        if st.button("➕ Add Object", use_container_width=True) and new_obj.strip():
            if new_obj not in st.session_state.mat_objects:
                st.session_state.mat_objects.append(new_obj.strip())
                sync_response_tables()
                st.rerun()

        objs_list = st.session_state.mat_objects[:]
        for o in objs_list:
            if st.button(f"❌ {o}", key=f"mat_rm_{o}"):
                st.session_state.mat_objects.remove(o)
                sync_response_tables()
                st.rerun()

        st.divider()
        st.markdown("**Settings**")
        st.session_state["mat_benchmark"] = st.slider(
            "Benchmark (1-5)", 1.0, 5.0, st.session_state.get("mat_benchmark", 3.0), 0.1
        )
        st.session_state["mat_target"] = st.slider(
            "Target (1-5)", 1.0, 5.0, st.session_state.get("mat_target", 3.0), 0.1
        )
        st.session_state["mat_low_thr"] = st.slider(
            "Exception Threshold", 1.0, 5.0, st.session_state.get("mat_low_thr", 2.0), 0.1
        )

    # Ensure tables are in sync
    sync_response_tables()

    dq_score = st.session_state.get("dq_score")
    if dq_score is not None:
        autofill_dq_dimension(dq_score)

    # ═══════════════════════════════════════════════════════════
    # SUBMIT HANDLER
    # ═══════════════════════════════════════════════════════════
    def _do_submit():
        ok, msg = validate_responses(
            st.session_state.mat_responses,
            st.session_state.mat_dims,
            st.session_state.mat_objects,
        )
        if not ok:
            st.error(f"Validation failed: {msg}")
            return

        dim_table, overall = compute_all_scores(
            st.session_state.mat_objects,
            st.session_state.mat_dims,
            st.session_state.mat_responses,
        )

        # Slide visual
        cn        = st.session_state["mat_client"] or "Client"
        dim_vals  = {dim: float(np.nanmean(dim_table.loc[dim].values))
                     for dim in dim_table.index}
        exec_sc   = float(np.nanmean(overall.values))
        bench     = st.session_state.get("mat_benchmark", 3.0)
        targ      = st.session_state.get("mat_target", 3.0)
        slide_png = render_slide_png(cn, dim_vals, exec_sc, bench, targ)

        # PDF
        pdf_bytes = build_pdf_bytes(
            client_name=cn,
            slide_png=slide_png,
            dim_table=dim_table,
            overall=overall,
            detail_tables=st.session_state.mat_responses,
            dq_score=dq_score,
        )

        # Maturity Excel
        low_thr      = st.session_state.get("mat_low_thr", 2.0)
        mat_excel    = to_excel_bytes(
            dim_table, overall, st.session_state.mat_responses,
            low_thr, st.session_state.mat_objects
        )

        # Combined Excel (DQ + Maturity)
        if dq_score is not None and st.session_state.get("dq_results_df") is not None:
            from openpyxl import Workbook, load_workbook
            from io import BytesIO as BIO

            # Load maturity wb
            wb_mat = load_workbook(BIO(mat_excel))

            # DQ results
            dq_results = st.session_state["dq_results_df"]
            display_c  = [c for c in dq_results.columns if not c.startswith("_")]
            ws_dq      = wb_mat.create_sheet("DQ - Results", 0)
            for r_idx, row in enumerate(
                [display_c] + dq_results[display_c].head(1000).values.tolist(), start=1
            ):
                for c_idx, val in enumerate(row, start=1):
                    ws_dq.cell(r_idx, c_idx, value=str(val) if val is not None else "")

            # DQ dimension scores
            dim_sc = st.session_state.get("dq_dim_scores")
            if dim_sc:
                ws_dim = wb_mat.create_sheet("DQ - Dimension Scores", 1)
                ws_dim.append(["Dimension", "DQ Score (%)"])
                for dim, sc in dim_sc.items():
                    ws_dim.append([dim, sc])

            out = BIO()
            wb_mat.save(out)
            combined_excel = out.getvalue()
        else:
            combined_excel = mat_excel

        st.session_state["mat_submitted"] = True
        st.session_state["mat_payload"]   = {
            "slide_png":       slide_png,
            "dim_table":       dim_table,
            "overall":         overall,
            "pdf_bytes":       pdf_bytes,
            "mat_excel":       mat_excel,
            "combined_excel":  combined_excel,
        }
        st.rerun()

    # ═══════════════════════════════════════════════════════════
    # RESULTS VIEW
    # ═══════════════════════════════════════════════════════════
    if st.session_state.get("mat_submitted") and st.session_state.get("mat_payload"):
        p  = st.session_state["mat_payload"]
        cn = st.session_state.get("mat_client") or "Client"
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        st.markdown('<h1 class="purple-text">✅ Maturity Assessment Complete</h1>',
                    unsafe_allow_html=True)
        st.divider()

        # DQ link banner
        if dq_score is not None:
            lvl = dq_score_to_maturity_level(dq_score)
            st.markdown(
                f'<div class="alert-box">'
                f'🔗 <b>DQ Engine</b> — Score: '
                f'<span class="purple-text large-text">{dq_score:.1f}%</span>'
                f' → Maturity: '
                f'<span class="magenta-text">{lvl}</span> &nbsp;'
                f'(auto-applied to <i>Data Quality</i> dimension)'
                f'</div>',
                unsafe_allow_html=True,
            )

        # Slide image
        st.markdown("### 📊 Summary Slide")
        st.image(p["slide_png"], use_container_width=True)
        st.divider()

        # Dimension + Overall tables
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

        # Maturity bar chart
        st.markdown("#### Scores by Dimension")
        dim_vals = {
            dim: float(np.nanmean(p["dim_table"].loc[dim].values))
            for dim in p["dim_table"].index
        }
        bar_img = _mat_bar_png(dim_vals)
        if bar_img:
            st.image(bar_img, use_container_width=True)

        st.divider()

        # Downloads
        st.markdown("### 📥 Download Reports")
        safe_cn = cn.replace(" ", "_")

        d1, d2, d3 = st.columns(3)

        with d1:
            st.markdown(
                '<div class="download-card">'
                '<h4>📄 PDF Report</h4>'
                '<p>Slide visual + DQ summary + dimension tables + detailed responses</p>'
                '</div>',
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
                '<div class="download-card magenta">'
                '<h4>📊 Maturity Excel</h4>'
                '<p>Dimension scores + overall scores + exception sheets</p>'
                '</div>',
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
                '<div class="download-card accent">'
                '<h4>🔗 Combined Excel</h4>'
                '<p>DQ scores + maturity scores + exceptions in one workbook</p>'
                '</div>',
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
                    '<div class="info-box">'
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
load_css()
_init_state()

{
    "home":     page_home,
    "dq":       page_dq,
    "maturity": page_maturity,
}[st.session_state.page]()
