"""
modules/report_generator.py
============================
Generates the standard Excel DQ report — identical format to
DQ_Report_20260210_232325.xlsx (the "old" target format).

Sheet order:
  1. DQ Score            – title banner + metrics table + interpretation
  2. Results             – full results dataframe (all rows, cleaned)
  3. Summary             – column-by-column pass/fail summary
  4. Dimension Analysis  – per-dimension score table
  5..N Annexure 1..N     – one sheet per data column (ALL columns, whether PASSED or FAILED)
  N+1. Uniqueness Issues  – duplicate records with Duplicate_Combination + Duplicate_Count
  N+2. Completeness Issues – null/missing-value records with Incomplete_Columns
  N+3. Standardization Issues – format/pattern violation records with Non_Standard_Columns
"""

import json
import datetime
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple
from collections import defaultdict

from modules.utils import clean_value, get_timestamp


class ExcelReportGenerator:
    """Generate the standard DQ Excel report."""

    def __init__(
        self,
        results_df:             pd.DataFrame,
        rulebook:               Dict,
        all_columns:            List[str],
        column_scores:          Dict[str, float],
        overall_score:          float,
        dimension_scores:       Dict[str, float],
        duplicate_combinations: Dict[str, List[Tuple]] = None,
    ):
        self.results_df             = results_df
        self.rulebook               = rulebook
        self.all_columns            = all_columns
        self.column_scores          = column_scores
        self.overall_score          = overall_score
        self.dimension_scores       = dimension_scores
        self.duplicate_combinations = duplicate_combinations or {}

    # ──────────────────────────────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────────────────────────────
    def generate_report(self, output_dir: Path) -> Path:
        """Build the workbook and return its path."""
        output_path = output_dir / f"DQ_Report_{get_timestamp()}.xlsx"

        # ── Build df_out: drop internal helper columns, clean every cell ──
        internal = [c for c in ("_failed_columns_list", "_failed_rules_details")
                    if c in self.results_df.columns]
        df_out = self.results_df.drop(columns=internal, errors="ignore").copy()

        for col in df_out.columns:
            try:
                df_out[col] = df_out[col].apply(clean_value)
            except Exception:
                df_out[col] = df_out[col].astype(str)

        # ── Build failure indexes (uses original results_df, not df_out) ──
        # failed_tracker:    col  → {row_idx}  (for per-column annexures)
        # dimension_tracker: dim  → col → {row_idx}  (for specialised sheets)
        failed_tracker    = defaultdict(set)
        dimension_tracker = defaultdict(lambda: defaultdict(set))

        for idx, row in self.results_df.iterrows():
            failed_cols = row.get("_failed_columns_list", [])
            if isinstance(failed_cols, list):
                for col in failed_cols:
                    failed_tracker[col].add(idx)

            failed_rules = row.get("_failed_rules_details", [])
            if isinstance(failed_rules, list):
                for rd in failed_rules:
                    if isinstance(rd, dict):
                        dim = rd.get("dimension", "General")
                        col = rd.get("column")
                        if col:
                            dimension_tracker[dim][col].add(idx)

        # ── Write workbook ─────────────────────────────────────────────────
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            wb  = writer.book
            fmt = self._make_formats(wb)

            self._sheet_dq_score(writer, fmt)
            self._sheet_results(writer, df_out, fmt)
            self._sheet_summary(writer, df_out, fmt, failed_tracker)
            self._sheet_dimension(writer, fmt)
            self._sheets_annexures(writer, df_out, fmt, failed_tracker)
            self._sheet_uniqueness(writer, df_out, fmt)
            self._sheet_completeness(writer, df_out, fmt, dimension_tracker)
            self._sheet_standardization(writer, df_out, fmt, dimension_tracker)

        return output_path

    def save_rulebook_json(self, output_dir: Path, rulebook: Dict) -> Path:
        path = output_dir / f"Rulebook_{get_timestamp()}.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(rulebook, f, indent=2)
        return path

    # ──────────────────────────────────────────────────────────────────────
    # XlsxWriter formats
    # ──────────────────────────────────────────────────────────────────────
    @staticmethod
    def _make_formats(wb) -> Dict:
        return {
            "title": wb.add_format({
                "bold": True, "font_size": 16, "font_color": "#1F4E78",
            }),
            "subtitle": wb.add_format({
                "italic": True, "font_size": 11,
            }),
            "header": wb.add_format({
                "bold": True, "bg_color": "#4472C4", "font_color": "white",
                "border": 1, "align": "center", "valign": "vcenter",
            }),
            "score": wb.add_format({
                "bold": True, "font_size": 20, "bg_color": "#E7E6E6",
                "border": 1, "align": "center",
            }),
            "metric": wb.add_format({
                "bold": True, "font_size": 12,
            }),
            "pass": wb.add_format({
                "bg_color": "#C6EFCE", "font_color": "#006100",
                "bold": True, "align": "center",
            }),
            "fail": wb.add_format({
                "bg_color": "#FFC7CE", "font_color": "#9C0006",
                "bold": True, "align": "center",
            }),
            "warning": wb.add_format({
                "bg_color": "#FFEB9C", "font_color": "#9C6500",
                "bold": True, "align": "center",
            }),
        }

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 1 : DQ Score
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_dq_score(self, writer, fmt):
        """
        Row 0  : 'DATA QUALITY ASSESSMENT REPORT'   (title)
        Row 1  : 'Generated: YYYY-MM-DD HH:MM:SS'   (subtitle)
        Row 2  : blank
        Row 3  : header  ('Metric', 'Value')
        Rows 4-11 : metric rows
        Row 12 : blank
        Row 13 : 'Score Interpretation:'
        Row 14 : interpretation text
        """
        total  = len(self.results_df)
        clean  = int((self.results_df["Count of issues"] == 0).sum())
        data_cols = [c for c in self.all_columns
                     if not c.startswith("_") and c not in
                     ("Issues", "Count of issues", "Failed_Rules",
                      "Failed_Columns", "Issue categories")]

        metrics = [
            ("Overall DQ Score (%)",  self.overall_score),
            ("Total Records",         total),
            ("Clean Records",         clean),
            ("Records with Issues",   total - clean),
            ("Total Columns",         len(data_cols)),
            ("Columns at 100%",       sum(1 for s in self.column_scores.values() if s == 100)),
            ("Columns with Issues",   sum(1 for s in self.column_scores.values() if s < 100)),
            ("Total Rules Applied",   len(self.rulebook.get("rules", []))),
        ]

        df = pd.DataFrame(metrics, columns=["Metric", "Value"])
        df.to_excel(writer, sheet_name="DQ Score", index=False, startrow=3)

        ws = writer.sheets["DQ Score"]
        ws.write(0, 0, "DATA QUALITY ASSESSMENT REPORT", fmt["title"])
        ws.write(1, 0,
            f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            fmt["subtitle"])

        ws.write(3, 0, "Metric", fmt["header"])
        ws.write(3, 1, "Value",  fmt["header"])
        ws.set_column(0, 0, 35)
        ws.set_column(1, 1, 35)

        # Score value cell gets the big number format
        ws.write(4, 1, self.overall_score, fmt["score"])

        ws.write(13, 0, "Score Interpretation:", fmt["metric"])
        ws.write(14, 0, self._interpret(self.overall_score))

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 2 : Results
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_results(self, writer, df_out, fmt):
        df_out.to_excel(writer, sheet_name="Results", index=False)
        ws = writer.sheets["Results"]
        for c, col in enumerate(df_out.columns):
            ws.write(0, c, col, fmt["header"])
            ws.set_column(c, c, 25)

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 3 : Summary
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_summary(self, writer, df_out, fmt, failed_tracker):
        total = len(self.results_df)
        skip  = {"Issues", "Count of issues", "Failed_Rules",
                  "Failed_Columns", "Issue categories"}
        rows  = []
        for col in self.all_columns:
            if col.startswith("_") or col in skip:
                continue
            failed = len(failed_tracker.get(col, set()))
            rows.append({
                "Column":        col,
                "Total Records": total,
                "Failed":        failed,
                "Passed":        total - failed,
                "DQ Score (%)":  self.column_scores.get(col, 100.0),
                "Status":        "PASSED" if failed == 0 else "FAILED",
            })

        df = pd.DataFrame(rows)
        df.to_excel(writer, sheet_name="Summary", index=False)
        ws = writer.sheets["Summary"]
        for c, col in enumerate(df.columns):
            ws.write(0, c, col, fmt["header"])
            ws.set_column(c, c, 20)
        for r in range(len(df)):
            status = df.iloc[r]["Status"]
            ws.write(r + 1, 5, status,
                     fmt["pass"] if status == "PASSED" else fmt["fail"])

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 4 : Dimension Analysis
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_dimension(self, writer, fmt):
        if not self.dimension_scores:
            return
        rows = [{"Dimension": d, "DQ Score (%)": s,
                 "Status": "PASSED" if s == 100 else "FAILED"}
                for d, s in self.dimension_scores.items()]
        df = pd.DataFrame(rows)
        df.to_excel(writer, sheet_name="Dimension Analysis", index=False)
        ws = writer.sheets["Dimension Analysis"]
        for c, col in enumerate(df.columns):
            ws.write(0, c, col, fmt["header"])
            ws.set_column(c, c, 25)

    # ──────────────────────────────────────────────────────────────────────
    # Sheets 5..N : Annexure per column  (ALL columns, PASSED or FAILED)
    # ──────────────────────────────────────────────────────────────────────
    def _sheets_annexures(self, writer, df_out, fmt, failed_tracker):
        """
        One annexure sheet per data column.
        - PASSED columns : rows 0-4 header only, empty data area, column headers on row 4
        - FAILED columns : rows 0-3 header, row 4 column headers, rows 5+ failed records
        """
        skip = {"Issues", "Count of issues", "Failed_Rules",
                "Failed_Columns", "Issue categories"}
        n = 1
        for col in self.all_columns:
            if col.startswith("_") or col in skip:
                continue

            failed_idx = sorted(failed_tracker.get(col, set()))
            score      = self.column_scores.get(col, 100.0)
            sheet      = f"Annexure {n}"

            if failed_idx:
                ann_df   = df_out.iloc[failed_idx].copy()
                subtitle = f"Records with issues: {len(failed_idx)} | DQ Score: {score}%"
                passed   = False
            else:
                # Empty df with same columns → just headers shown on row 4
                ann_df   = df_out.iloc[0:0].copy()
                subtitle = "No issues found | DQ Score: 100%"
                passed   = True

            # Write data starting at row 4 (leaves rows 0-3 for header info)
            ann_df.to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]

            ws.write(0, 0, f"Annexure {n}: {col}", fmt["title"])
            ws.write(1, 0, subtitle,                 fmt["subtitle"])
            ws.write(2, 0,
                "Status: ✅ PASSED" if passed else "Status: ❌ FAILED",
                fmt["pass"] if passed else fmt["fail"])
            # Row 3 blank, row 4 has column headers from to_excel
            for c, h in enumerate(df_out.columns):
                ws.write(4, c, h, fmt["header"])
                ws.set_column(c, c, 20)

            n += 1

    # ──────────────────────────────────────────────────────────────────────
    # Uniqueness Issues sheet
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_uniqueness(self, writer, df_out, fmt):
        """
        Header rows 0-3, data from row 4.
        Extra columns added: Duplicate_Combination, Duplicate_Count
        'index' column preserved as first column.
        """
        dup_records = set()
        dup_details = []
        for combo_str, groups in self.duplicate_combinations.items():
            for idx_list in groups:
                dup_records.update(idx_list)
                for idx in idx_list:
                    dup_details.append({
                        "Index":                idx,
                        "Duplicate_Combination": combo_str,
                        "Duplicate_Count":       len(idx_list),
                    })

        sheet = "Uniqueness Issues"
        if dup_records:
            dup_df = df_out.iloc[sorted(dup_records)].copy()

            info_df = (
                pd.DataFrame(dup_details)
                .groupby("Index")
                .agg(
                    Duplicate_Combination=("Duplicate_Combination",
                                           lambda x: " | ".join(sorted(set(x)))),
                    Duplicate_Count=("Duplicate_Count", "max"),
                )
                .reset_index()
            )

            dup_df = dup_df.reset_index()           # adds 'index' column
            dup_df = dup_df.merge(info_df, left_on="index",
                                  right_on="Index", how="left")
            dup_df = dup_df.drop(columns=["Index"])

            dup_df.to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "UNIQUENESS VALIDATION - DUPLICATE RECORDS", fmt["title"])
            ws.write(1, 0, f"Total Duplicate Records: {len(dup_records)}", fmt["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", fmt["fail"])
            for c, h in enumerate(dup_df.columns):
                ws.write(4, c, h, fmt["header"])
                ws.set_column(c, c, 20)
        else:
            pd.DataFrame().to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "UNIQUENESS VALIDATION - DUPLICATE RECORDS", fmt["title"])
            ws.write(1, 0, "No duplicate records found",                 fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED",                          fmt["pass"])

    # ──────────────────────────────────────────────────────────────────────
    # Completeness Issues sheet
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_completeness(self, writer, df_out, fmt, dimension_tracker):
        """
        Header rows 0-3, data from row 4.
        Extra column: Incomplete_Columns  (which not_null columns failed)
        'index' column preserved as first column.
        """
        indices = set()
        for col_set in dimension_tracker.get("Completeness", {}).values():
            indices.update(col_set)

        sheet = "Completeness Issues"
        if indices:
            comp_df = df_out.iloc[sorted(indices)].copy()

            info = []
            for idx in sorted(indices):
                bad_cols = []
                row = self.results_df.iloc[idx]
                for rd in (row.get("_failed_rules_details") or []):
                    if isinstance(rd, dict) and rd.get("dimension") == "Completeness":
                        c = rd.get("column", "")
                        if c:
                            bad_cols.append(c)
                info.append({
                    "Index":             idx,
                    "Incomplete_Columns": ", ".join(sorted(set(bad_cols))),
                })

            info_df = pd.DataFrame(info)
            comp_df = comp_df.reset_index()         # adds 'index' column
            comp_df = comp_df.merge(info_df, left_on="index",
                                    right_on="Index", how="left")
            comp_df = comp_df.drop(columns=["Index"])

            comp_df.to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "COMPLETENESS VALIDATION - MISSING VALUES", fmt["title"])
            ws.write(1, 0, f"Total Records with Missing Values: {len(indices)}", fmt["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", fmt["fail"])
            for c, h in enumerate(comp_df.columns):
                ws.write(4, c, h, fmt["header"])
                ws.set_column(c, c, 20)
        else:
            pd.DataFrame().to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "COMPLETENESS VALIDATION - MISSING VALUES", fmt["title"])
            ws.write(1, 0, "No completeness issues found",              fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED",                         fmt["pass"])

    # ──────────────────────────────────────────────────────────────────────
    # Standardization Issues sheet
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_standardization(self, writer, df_out, fmt, dimension_tracker):
        """
        Header rows 0-3, data from row 4.
        Extra column: Non_Standard_Columns
        'index' column preserved as first column.
        Collects from dimension keys: 'Standardization', 'Validation'
        """
        indices = set()
        for dim_key in ("Standardization", "Validation"):
            for col_set in dimension_tracker.get(dim_key, {}).values():
                indices.update(col_set)

        sheet = "Standardization Issues"
        if indices:
            std_df = df_out.iloc[sorted(indices)].copy()

            info = []
            for idx in sorted(indices):
                bad_cols = []
                row = self.results_df.iloc[idx]
                for rd in (row.get("_failed_rules_details") or []):
                    if isinstance(rd, dict) and rd.get("dimension") in (
                            "Standardization", "Validation"):
                        c = rd.get("column", "")
                        if c:
                            bad_cols.append(c)
                info.append({
                    "Index":               idx,
                    "Non_Standard_Columns": ", ".join(sorted(set(bad_cols))),
                })

            info_df = pd.DataFrame(info)
            std_df  = std_df.reset_index()          # adds 'index' column
            std_df  = std_df.merge(info_df, left_on="index",
                                   right_on="Index", how="left")
            std_df  = std_df.drop(columns=["Index"])

            std_df.to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "STANDARDIZATION VALIDATION - FORMAT ISSUES", fmt["title"])
            ws.write(1, 0,
                f"Total Records with Standardization Issues: {len(indices)}",
                fmt["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", fmt["fail"])
            for c, h in enumerate(std_df.columns):
                ws.write(4, c, h, fmt["header"])
                ws.set_column(c, c, 20)
        else:
            pd.DataFrame().to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "STANDARDIZATION VALIDATION - FORMAT ISSUES", fmt["title"])
            ws.write(1, 0, "No standardization issues found",             fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED",                           fmt["pass"])

    # ──────────────────────────────────────────────────────────────────────
    # Helper
    # ──────────────────────────────────────────────────────────────────────
    @staticmethod
    def _interpret(score: float) -> str:
        if score >= 95:
            return "🎉 Excellent! Outstanding data quality."
        if score >= 80:
            return "👍 Good! Minor improvements needed."
        if score >= 60:
            return "⚠️ Fair! Significant improvements required."
        return "❌ Poor! Critical data quality issues detected."