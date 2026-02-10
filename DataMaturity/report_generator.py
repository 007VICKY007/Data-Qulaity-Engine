"""
DataMaturity/report_generator.py
=================================
ReportLab PDF generation for the Data Maturity Assessment report.

PDF structure
-------------
Page 1  – Matplotlib slide image  (from visualizations.render_slide_png)
Page 2  – Numeric summary tables  (dimension scores + overall scores)
Page 3+ – Per-dimension detailed response tables

Public API
----------
build_pdf_bytes(client_name, slide_png, dim_table, overall,
                detail_tables, dq_score=None) → bytes
"""

import numpy as np
import pandas as pd
from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak, Image as RLImage,
)
from reportlab.lib.units import inch

from DataMaturity.config import UNIQU_PURPLE, RATING_TO_SCORE
from DataMaturity.helpers import (
    compute_weighted_scores,
    dq_score_to_maturity_level,
)


# ──────────────────────────────────────────────────────────────
# Internal: DataFrame → styled ReportLab Table
# ──────────────────────────────────────────────────────────────
def _rl_table(df: pd.DataFrame, max_rows: int = 35) -> Table:
    d    = df.head(max_rows).copy()
    data = [list(d.columns)] + [
        [str(cell) for cell in row] for row in d.values.tolist()
    ]
    t = Table(data, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1,  0), colors.HexColor(UNIQU_PURPLE)),
        ("TEXTCOLOR",      (0, 0), (-1,  0), colors.white),
        ("FONTNAME",       (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",       (0, 0), (-1,  0), 9),
        ("FONTSIZE",       (0, 1), (-1, -1), 8),
        ("GRID",           (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
        ("VALIGN",         (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",    (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 4),
        ("TOPPADDING",     (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 3),
    ]))
    return t


# ──────────────────────────────────────────────────────────────
# Public: Build full PDF
# ──────────────────────────────────────────────────────────────
def build_pdf_bytes(
    client_name:   str,
    slide_png:     bytes,
    dim_table:     pd.DataFrame,
    overall:       pd.Series,
    detail_tables: dict,
    dq_score:      float = None,
) -> bytes:
    """
    Generate the full maturity assessment PDF.

    Parameters
    ----------
    client_name   : Organisation name used in headings / footers
    slide_png     : PNG bytes from visualizations.render_slide_png()
    dim_table     : DataFrame (dimensions × objects, scores 1-5)
    overall       : Series   (overall score per object)
    detail_tables : {dim_name: response_DataFrame}
    dq_score      : Optional DQ engine score – adds a summary box on page 2
    """
    buff = BytesIO()
    doc  = SimpleDocTemplate(
        buff,
        pagesize=landscape(A4),
        leftMargin=0.5 * inch,  rightMargin=0.5 * inch,
        topMargin=0.5  * inch,  bottomMargin=0.5 * inch,
    )
    styles = getSampleStyleSheet()
    story  = []
    cn     = client_name.strip() or "Client"

    # ── Page 1 : Slide image ──────────────────────────────────
    story.append(RLImage(BytesIO(slide_png), width=10.8 * inch, height=6.1 * inch))
    story.append(PageBreak())

    # ── Page 2 : Numeric summary ──────────────────────────────
    story.append(Paragraph(f"Detailed Summary – {cn}", styles["Title"]))
    story.append(Spacer(1, 10))

    # Optional DQ Engine input box
    if dq_score is not None:
        story.append(Paragraph("DQ Engine Input", styles["Heading2"]))
        story.append(Spacer(1, 4))
        dq_df = pd.DataFrame({
            "Metric": ["DQ Overall Score (%)", "Mapped Maturity Level"],
            "Value":  [f"{dq_score:.1f}%", dq_score_to_maturity_level(dq_score)],
        })
        story.append(_rl_table(dq_df))
        story.append(Spacer(1, 12))

    # Dimension table
    story.append(Paragraph("Dimension-wise Maturity (Weighted Average 1–5)", styles["Heading2"]))
    story.append(Spacer(1, 6))
    dim_df = (
        dim_table
        .reset_index()
        .rename(columns={"index": "Dimension"})
        .round(2)
    )
    story.append(_rl_table(dim_df, max_rows=50))
    story.append(Spacer(1, 12))

    # Overall table
    story.append(Paragraph("Overall Maturity (Average of Dimensions)", styles["Heading2"]))
    story.append(Spacer(1, 6))
    ov_df = pd.DataFrame({
        "Master Data Object": list(overall.index),
        "Overall Score":      [round(float(v), 2) for v in overall.values],
    })
    story.append(_rl_table(ov_df, max_rows=200))
    story.append(PageBreak())

    # ── Pages 3+ : Per-dimension detailed responses ───────────
    for dim, df in detail_tables.items():
        story.append(Paragraph(f"Detailed Responses – {dim}", styles["Heading2"]))
        story.append(Spacer(1, 6))

        obj_cols = [c for c in df.columns
                    if c not in ["Question ID", "Section", "Question", "Weight"]]
        compact  = df[["Question ID", "Section", "Question", "Weight"] + obj_cols].copy()

        # Convert rating labels → numeric scores for the table
        for o in obj_cols:
            compact[o] = compact[o].map(RATING_TO_SCORE)

        # Truncate very long question text
        compact["Question"] = compact["Question"].astype(str).str.slice(0, 95)

        story.append(_rl_table(compact, max_rows=35))
        story.append(Spacer(1, 8))
        story.append(Paragraph(
            "Note: Answers shown as numeric scores (1 = Adhoc … 5 = Optimised).",
            styles["Italic"],
        ))
        story.append(PageBreak())

    doc.build(story)
    return buff.getvalue()
