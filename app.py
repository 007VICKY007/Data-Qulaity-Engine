# ── stdlib ─────────────────────────────────────────────────────────────────
import traceback
import datetime
from io import BytesIO

# ── third-party ────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge

# ── DQ Engine modules ──────────────────────────────────────────────────────
from modules.config           import AppConfig
from modules.file_loader      import FileLoaderService
from modules.rulebook_builder import RulebookBuilderService
from modules.rule_executor    import RuleExecutorEngine
from modules.scoring_engine   import ScoringService
from modules.report_generator import ExcelReportGenerator
from modules.ui_components    import UIComponents
from modules.utils            import (
    setup_directories, save_uploaded_file, clean_temp_directory,
)

# ── DataMaturity modules ───────────────────────────────────────────────────
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

from DataMaturity.visualizations   import render_slide_png
from DataMaturity.report_generator import build_pdf_bytes


# ══════════════════════════════════════════════════════════════════════════
#  APP CONFIG & EXTERNAL CSS
# ══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="DQ & Maturity Assessment",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

setup_directories()


def load_css():
    """Load external stylesheet from assets folder."""
    try:
        with open("assets/styles.css", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning("⚠️ styles.css not found in assets/ folder")


# ══════════════════════════════════════════════════════════════════════════
#  SESSION STATE INITIALIZATION
# ══════════════════════════════════════════════════════════════════════════
def _init_state() -> None:
    # Navigation
    if "page" not in st.session_state:
        st.session_state["page"] = "home"

    # DQ results (populated after DQ run; consumed by Maturity)
    for key, default in {
        "dq_score":       None,
        "dq_dim_scores":  None,
        "dq_results_df":  None,
        "dq_object_name": "Customer",
        "dq_excel_path":  None,
    }.items():
        if key not in st.session_state:
            st.session_state[key] = default

    # Maturity state (uses keys defined in DataMaturity/helpers.py)
    init_maturity_state()


# ══════════════════════════════════════════════════════════════════════════
#  VISUALIZATION FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════
def _gauge_png(score: float) -> bytes:
    """Semicircle gauge for overall DQ score."""
    fig, ax = plt.subplots(figsize=(4.5, 2.8), dpi=130)
    fig.patch.set_facecolor(UNIQU_LIGHT_BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 0.65)
    ax.axis("off")

    ax.add_patch(Wedge((0.5, 0.05), 0.40, 0, 180, width=0.10,
                       facecolor="#ddd6f0", edgecolor="white", lw=2))
    ang = score / 100 * 180
    col = "#22c55e" if score >= 80 else ("#f59e0b" if score >= 60 else "#ef4444")
    ax.add_patch(Wedge((0.5, 0.05), 0.40, 0, ang, width=0.10,
                       facecolor=col, edgecolor="white", lw=2))

    ax.text(0.5, 0.30, f"{score:.1f}%", ha="center", va="center",
            fontsize=22, fontweight="bold", color=UNIQU_PURPLE)
    ax.text(0.5, 0.16, "Overall DQ Score", ha="center", va="center",
            fontsize=9.5, color="#555")

    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.1)
    plt.close(fig)
    return buf.getvalue()


def _dim_bar_png(dim_scores: dict) -> bytes | None:
    """Horizontal bar chart – DQ dimension scores."""
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
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.spines["left"].set_color(UNIQU_GREY)

    ax.axvline(80, color=UNIQU_PURPLE,  lw=1, ls="--", alpha=0.5, label="Good (80%)")
    ax.axvline(60, color=UNIQU_MAGENTA, lw=1, ls=":",  alpha=0.5, label="Fair (60%)")
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


def _mat_bar_png(dim_vals: dict) -> bytes | None:
    """Horizontal bar chart – maturity dimension scores."""
    if not dim_vals:
        return None

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

    ax.axvline(3.0, color="#aaa",       lw=1, ls="--", alpha=0.6, label="Defined (3)")
    ax.axvline(4.0, color=UNIQU_PURPLE, lw=1, ls="--", alpha=0.6, label="Managed (4)")
    ax.legend(fontsize=8, loc="lower right", framealpha=0.7)

    ax.tick_params(colors=UNIQU_TEXT, labelsize=9)
    ax.spines[["top", "right", "bottom"]].set_visible(False)
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
#  COMBINED EXCEL (DQ + Maturity)
# ══════════════════════════════════════════════════════════════════════════
def _combined_excel(dq_score: float, dq_dim_scores: dict | None, mat_excel: bytes) -> bytes:
    from openpyxl import load_workbook

    wb = load_workbook(BytesIO(mat_excel))

    # Sheet 1: DQ Score Summary
    ws_dq = wb.create_sheet("DQ Score Summary", 0)
    ws_dq.append(["Metric", "Value"])
    ws_dq.append(["Overall DQ Score (%)",  f"{dq_score:.1f}%"])
    ws_dq.append(["Mapped Maturity Level", dq_score_to_maturity_level(dq_score)])

    if dq_dim_scores:
        for dim, sc in dq_dim_scores.items():
            ws_dq.append([f"DQ – {dim}", f"{sc:.1f}%"])

    # Sheet 2: DQ Results (first 1000 rows)
    dq_df = st.session_state.get("dq_results_df")
    if dq_df is not None:
        display_cols = [c for c in dq_df.columns if not c.startswith("_")]
        ws_res = wb.create_sheet("DQ Results", 1)
        ws_res.append(display_cols)
        for _, row in dq_df[display_cols].head(1000).iterrows():
            ws_res.append([str(v) if v is not None else "" for v in row.tolist()])

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════
#  PAGE: HOME
# ══════════════════════════════════════════════════════════════════════════
def page_home():
    st.title("📊 Enterprise DQ & Maturity Platform")
    st.markdown("**One integrated platform** for Data Quality validation "
                "and DAMA Maturity assessment.")
    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f'<div class="card">'
            f'<h3>🔍 DQ Assessment</h3>'
            f'<p>Upload data + rulebook → DQ score, dimension breakdown, '
            f'column annexures, automated maturity mapping.</p></div>',
            unsafe_allow_html=True,
        )
        if st.button("Go to DQ Assessment →", use_container_width=True):
            st.session_state["page"] = "dq"
            st.rerun()

    with c2:
        st.markdown(
            f'<div class="card">'
            f'<h3>📈 Maturity Assessment</h3>'
            f'<p>DAMA questionnaire → slide-style visuals + PDF + Excel. '
            f'Optionally auto-populated from DQ score.</p></div>',
            unsafe_allow_html=True,
        )
        if st.button("Go to Maturity Assessment →", use_container_width=True):
            st.session_state["page"] = "maturity"
            st.rerun()

    # Show DQ completion banner if available
    if st.session_state.dq_score is not None:
        sc  = st.session_state.dq_score
        lvl = dq_score_to_maturity_level(sc)
        st.divider()
        st.success(
            f"✅ **DQ Assessment completed**  \n"
            f"• Score: **{sc:.1f}%**  \n"
            f"• Maturity mapping: **{lvl}**  \n"
            f"• Object: **{st.session_state.dq_object_name}**"
        )

    st.divider()
    st.markdown("#### DQ Score → Maturity Level Mapping")
    st.dataframe(
        pd.DataFrame([
            {"DQ Score Range": "95% – 100%", "Maturity Level": "⭐ Optimised",
             "Description": "Continuous improvement; proactive governance"},
            {"DQ Score Range": "80% – 94%",  "Maturity Level": "✅ Managed",
             "Description": "Monitored, measured, formalised roles"},
            {"DQ Score Range": "60% – 79%",  "Maturity Level": "🔵 Defined",
             "Description": "Standardised and consistently followed"},
            {"DQ Score Range": "40% – 59%",  "Maturity Level": "🟡 Repeatable",
             "Description": "Some processes defined but inconsistent"},
            {"DQ Score Range": "0%  – 39%",  "Maturity Level": "🔴 Adhoc",
             "Description": "Unstructured, reactive, varies widely"},
        ]),
        use_container_width=True,
        hide_index=True,
    )


# ══════════════════════════════════════════════════════════════════════════
#  PAGE: DQ ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════
def page_dq():
    with st.sidebar:
        st.markdown("### 🧭 Navigation")
        if st.button("🏠 Home", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
        if st.button("📈 Go to Maturity →", use_container_width=True):
            st.session_state["page"] = "maturity"
            st.rerun()
        st.divider()
        UIComponents.render_sidebar()

    st.title("🔍 Data Quality Assessment")
    st.caption("Upload master dataset + rules → comprehensive DQ report")
    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        data_file = st.file_uploader(
            "📁 Master Dataset",
            type=AppConfig.SUPPORTED_DATA_FORMATS,
            help="CSV / Excel / JSON",
        )
    with c2:
        rules_file = st.file_uploader(
            "📋 Rules Configuration",
            type=AppConfig.SUPPORTED_RULES_FORMATS + ["json"],
            help="CSV/Excel rules sheet or JSON rulebook",
        )

    if not data_file or not rules_file:
        UIComponents.render_welcome_screen()
        return

    obj_name = st.text_input(
        "🏷️ Master Data Object Name",
        value=st.session_state.get("dq_object_name", "Customer"),
        help="e.g. Customer, Vendor – carried into Maturity assessment",
    )

    sheet_name = None
    if data_file.name.lower().endswith((".xlsx", ".xls", ".xlsm")):
        loader = FileLoaderService()
        tmp    = AppConfig.TEMP_DIR / data_file.name
        tmp.write_bytes(data_file.getbuffer())
        sheets = loader.get_sheet_names(tmp)
        if len(sheets) > 1:
            sheet_name = st.selectbox("📄 Select Sheet", sheets)

    UIComponents.render_file_format_help()
    st.divider()

    if not st.button("🚀 Run DQ Assessment", type="primary", use_container_width=True):
        return

    try:
        clean_temp_directory()
        pb   = st.progress(0)
        stat = st.empty()

        stat.text("📂 Saving files…");       pb.progress(5)
        data_path  = save_uploaded_file(data_file,  AppConfig.TEMP_DIR)
        rules_path = save_uploaded_file(rules_file, AppConfig.TEMP_DIR)

        stat.text("📊 Loading dataset…");    pb.progress(15)
        loader = FileLoaderService()
        df     = loader.load_dataframe(data_path, sheet_name=sheet_name)
        cols   = list(df.columns)
        st.info(f"✅ Loaded **{len(df):,}** records · **{len(cols)}** columns")

        stat.text("🔧 Building rulebook…");  pb.progress(30)
        rb_svc = RulebookBuilderService()
        if rules_file.name.lower().endswith(".json"):
            rulebook = rb_svc.load_json_rulebook(rules_path)
        else:
            rulebook = rb_svc.build_from_rules_dataset(
                loader.load_dataframe(rules_path), cols)

        stat.text("✅ Executing rules…");    pb.progress(50)
        executor = RuleExecutorEngine(df, rulebook)
        results  = executor.execute_all_rules()
        combos   = executor.get_combination_duplicates()

        stat.text("📊 Scoring…");            pb.progress(70)
        scorer     = ScoringService()
        overall    = scorer.calculate_overall_score(results)
        col_scores = scorer.calculate_column_scores(results, cols)
        dim_scores = scorer.calculate_dimension_scores(results)

        stat.text("💾 Writing Excel…");      pb.progress(85)
        rgen = ExcelReportGenerator(
            results_df=results,
            rulebook=rulebook,
            all_columns=cols,
            column_scores=col_scores,
            overall_score=overall,
            dimension_scores=dim_scores,
            duplicate_combinations=combos,
        )
        xl_path = rgen.generate_report(AppConfig.OUTPUT_DIR)
        rb_path = rgen.save_rulebook_json(AppConfig.OUTPUT_DIR, rulebook)
        pb.progress(100)
        stat.text("✅ Complete!")

        # Save to session
        st.session_state["dq_score"]       = overall
        st.session_state["dq_dim_scores"]  = dim_scores
        st.session_state["dq_results_df"]  = results
        st.session_state["dq_object_name"] = obj_name or "Customer"
        st.session_state["dq_excel_path"]  = xl_path

        # Sync to maturity
        st.session_state["mat_objects"] = [obj_name] if obj_name else DEFAULT_MASTER_OBJECTS[:]
        autofill_dq_dimension(overall)

        st.success("✅ DQ Assessment Completed Successfully!")
        st.divider()

        # ── Results Dashboard ──────────────────────────────────────────
        st.subheader("📊 DQ Results Dashboard")
        g1, g2 = st.columns([1, 2])
        with g1:
            st.image(_gauge_png(overall), use_container_width=True)
        with g2:
            bar = _dim_bar_png(dim_scores)
            if bar:
                st.image(bar, use_container_width=True)

        st.divider()
        UIComponents.render_results_dashboard(overall, results, col_scores, dim_scores)
        st.divider()
        UIComponents.render_download_section(xl_path, rb_path, len(cols))
        st.divider()
        UIComponents.render_detailed_views(rulebook, results, col_scores, dim_scores)

        st.divider()
        lvl = dq_score_to_maturity_level(overall)
        st.info(f"DQ Score **{overall:.1f}%** → Maturity level **{lvl}**. "
                "Data Quality dimension auto-filled in Maturity assessment.")

        if st.button("📈 Continue to Maturity Assessment →", type="primary", use_container_width=True):
            st.session_state["page"] = "maturity"
            st.rerun()

    except Exception as e:
        st.error(f"❌ Error during DQ processing: {e}")
        with st.expander("🔍 Full traceback"):
            st.code(traceback.format_exc())


# ══════════════════════════════════════════════════════════════════════════
#  PAGE: MATURITY ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════
def _do_submit() -> None:
    objects   = st.session_state.mat_objects
    dims      = st.session_state.mat_dims
    responses = st.session_state.mat_responses
    cn        = st.session_state.mat_client_name or "Client"
    bm        = float(st.session_state.mat_benchmark)
    tg        = float(st.session_state.mat_target)
    lt        = float(st.session_state.mat_low_thr)
    dq_score  = st.session_state.get("dq_score")

    ok, msg = validate_responses(responses, dims, objects)
    if not ok:
        st.error(f"⚠️ Validation failed: {msg}")
        return

    with st.spinner("⚙️ Computing scores and building reports…"):
        dim_table, overall = compute_all_scores(objects, dims, responses)

        domain_display = {
            dim: float(np.nanmean(dim_table.loc[dim].values))
            for dim in dims
        }
        exec_score = float(np.nanmean(overall.values)) if len(overall) else 0.0

        slide_png = render_slide_png(
            client_name=cn,
            domain_scores=domain_display,
            exec_score=exec_score if np.isfinite(exec_score) else 0.0,
            benchmark=bm,
            target=tg,
        )

        pdf_bytes = build_pdf_bytes(
            client_name=cn,
            slide_png=slide_png,
            dim_table=dim_table,
            overall=overall,
            detail_tables=responses,
            dq_score=dq_score,
        )

        mat_excel = to_excel_bytes(
            dim_table=dim_table,
            overall=overall,
            detail_tables=responses,
            low_thr=lt,
            objects=objects,
        )

        if dq_score is not None:
            combined_excel = _combined_excel(
                dq_score,
                st.session_state.get("dq_dim_scores"),
                mat_excel,
            )
        else:
            combined_excel = mat_excel

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    st.session_state["mat_submitted"] = True
    st.session_state["mat_payload"]   = {
        "dim_table":      dim_table,
        "overall":        overall,
        "slide_png":      slide_png,
        "mat_excel":      mat_excel,
        "combined_excel": combined_excel,
        "pdf_bytes":      pdf_bytes,
        "client_name":    cn,
        "ts":             ts,
    }
    st.rerun()


def page_maturity():
    dq_score  = st.session_state.get("dq_score")
    submitted = st.session_state.get("mat_submitted", False)

    with st.sidebar:
        st.markdown("### 🧭 Navigation")
        if st.button("🏠 Home", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
        if st.button("🔍 DQ Assessment", use_container_width=True):
            st.session_state["page"] = "dq"
            st.rerun()
        st.divider()

        st.header("⚙️ Configuration")
        st.session_state["mat_client_name"] = st.text_input(
            "Client Name",
            value=st.session_state.get("mat_client_name", ""),
            placeholder="Organisation name",
            disabled=submitted,
        )

        all_obj_opts = list(dict.fromkeys(
            DEFAULT_MASTER_OBJECTS + st.session_state.mat_objects
        ))
        st.session_state["mat_objects"] = st.multiselect(
            "Master Data Objects",
            options=all_obj_opts,
            default=st.session_state.mat_objects,
            disabled=submitted,
        )

        st.session_state["mat_dims"] = st.multiselect(
            "Maturity Dimensions",
            options=MATURITY_DIMS,
            default=st.session_state.mat_dims,
            disabled=submitted,
        )

        st.divider()
        st.header("📊 Thresholds")
        st.session_state["mat_low_thr"] = st.slider(
            "Exception threshold (≤)", 1.0, 5.0,
            float(st.session_state.get("mat_low_thr", 2.0)), 0.5,
            disabled=submitted,
        )

        st.divider()
        st.header("🎯 Benchmark / Target")
        st.session_state["mat_benchmark"] = st.number_input(
            "Industry Benchmark", 1.0, 5.0,
            float(st.session_state.get("mat_benchmark", 3.0)), 0.1,
            disabled=submitted,
        )
        st.session_state["mat_target"] = st.number_input(
            "Target Score", 1.0, 5.0,
            float(st.session_state.get("mat_target", 3.0)), 0.1,
            disabled=submitted,
        )

    if not st.session_state.mat_objects or not st.session_state.mat_dims:
        st.info("⚙️ Please select at least one **Object** and one **Dimension** in the sidebar.")
        st.stop()

    sync_response_tables()

    if dq_score is not None:
        autofill_dq_dimension(dq_score)

    # ── REPORT VIEW (after submit) ─────────────────────────────────────
    if submitted and st.session_state.get("mat_payload"):
        p  = st.session_state["mat_payload"]
        cn = p["client_name"]
        ts = p["ts"]

        st.title("✅ Data Maturity Assessment Report")

        if dq_score is not None:
            lvl = dq_score_to_maturity_level(dq_score)
            st.info(
                f"**DQ Engine** — Score: **{dq_score:.1f}%**  →  "
                f"Maturity level: **{lvl}**  (applied to Data Quality dimension)"
            )

        st.markdown("### 📊 Summary Slide")
        st.image(p["slide_png"], use_container_width=True)
        st.divider()

        t1, t2 = st.columns(2)
        with t1:
            st.markdown("#### Dimension-wise Maturity (1–5)")
            st.dataframe(
                p["dim_table"].style
                  .format("{:.2f}")
                  .background_gradient(cmap="RdYlGn", axis=None, vmin=1, vmax=5),
                use_container_width=True,
            )
        with t2:
            st.markdown("#### Overall Maturity Score")
            st.dataframe(
                pd.DataFrame(p["overall"]).T.style
                  .format("{:.2f}")
                  .background_gradient(cmap="RdYlGn", axis=None, vmin=1, vmax=5),
                use_container_width=True,
            )

        st.divider()
        st.markdown("#### Scores by Dimension")
        dim_vals = {
            dim: float(np.nanmean(p["dim_table"].loc[dim].values))
            for dim in p["dim_table"].index
        }
        bar_img = _mat_bar_png(dim_vals)
        if bar_img:
            st.image(bar_img, use_container_width=True)

        st.divider()
        st.markdown("### 📥 Download Reports")
        safe_cn = cn.replace(" ", "_")
        d1, d2, d3 = st.columns(3)

        with d1:
            st.download_button(
                "📄 Download PDF",
                data=p["pdf_bytes"],
                file_name=f"Maturity_Report_{safe_cn}_{ts}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        with d2:
            st.download_button(
                "📊 Download Maturity Excel",
                data=p["mat_excel"],
                file_name=f"Maturity_Assessment_{safe_cn}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with d3:
            st.download_button(
                "🔗 Download Combined Excel (DQ + Maturity)",
                data=p["combined_excel"],
                file_name=f"DQ_Maturity_Combined_{safe_cn}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.divider()
        if st.button("✏️ Edit Responses", use_container_width=True):
            st.session_state["mat_submitted"] = False
            st.session_state["mat_payload"]   = {}
            st.rerun()

        st.stop()

    # ── QUESTIONNAIRE VIEW ─────────────────────────────────────────────
    st.title("📈 Data Maturity Assessment")

    if dq_score is not None:
        lvl = dq_score_to_maturity_level(dq_score)
        st.success(
            f"✅ DQ Score **{dq_score:.1f}%** → level **'{lvl}'** "
            f"auto-filled in the *Data Quality* dimension."
        )
    else:
        st.info(
            "💡 Run the **DQ Assessment** first to automatically fill "
            "the *Data Quality* dimension."
        )

    st.caption("Rate each question per master data object, then click **Submit**.")
    st.divider()

    dims = st.session_state.mat_dims
    tabs = st.tabs(dims)

    for i, dim in enumerate(dims):
        with tabs[i]:
            st.markdown(f"### {dim}")

            if dim == "Data Quality" and dq_score is not None:
                lvl = dq_score_to_maturity_level(dq_score)
                st.info(
                    f"Auto-populated from DQ Score **{dq_score:.1f}%** → **{lvl}**. "
                    "You can still adjust individual ratings."
                )

            df = st.session_state.mat_responses[dim].copy()

            cfg = {
                "Weight": st.column_config.NumberColumn(
                    "Weight", min_value=0.0, step=0.5),
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
                key=f"mat_editor_{dim}_{hash(dim)}",   # avoid key collision
            )

            st.session_state.mat_responses[dim] = edited

    st.divider()
    lft, rgt = st.columns([1, 3])
    with lft:
        if st.button("🚀 Submit & Generate Report", type="primary"):
            _do_submit()
    with rgt:
        st.info(
            "**Submit** computes scores, generates visuals, "
            "and enables PDF + Excel downloads."
        )


# ══════════════════════════════════════════════════════════════════════════
#  START APPLICATION
# ══════════════════════════════════════════════════════════════════════════
load_css()
_init_state()

{
    "home":     page_home,
    "dq":       page_dq,
    "maturity": page_maturity,
}[st.session_state.page]()