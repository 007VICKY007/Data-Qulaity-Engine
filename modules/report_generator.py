"""
modules/report_generator.py
============================
Generates the standard Excel DQ report with comprehensive validation results.

Sheet Structure:
  1. DQ Score            – title banner + metrics table + interpretation
  2. Results             – full results dataframe (all rows, cleaned)
  3. Summary             – column-by-column pass/fail summary
  4. Dimension Analysis  – per-dimension score table
  5..N Annexure 1..N     – one sheet per data column
  N+1. Uniqueness Issues  – duplicate records
  N+2. Completeness Issues – null/missing-value records
  N+3. Standardization Issues – format/pattern violation records

Features:
  - XlsxWriter formatting for professional appearance
  - Automatic handling of large datasets
  - Comprehensive error tracking and reporting
  - Fixed Excel worksheet name length issues (max 31 chars)
"""

import json
import datetime
import pandas as pd
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from collections import defaultdict

# Configure logging
logger = logging.getLogger(__name__)


def clean_value(val):
    """Clean and format cell values for Excel output."""
    if pd.isna(val):
        return ""
    if isinstance(val, (list, dict)):
        return str(val)
    if isinstance(val, bool):
        return "Yes" if val else "No"
    if isinstance(val, (int, float)):
        return val
    return str(val).strip()


def get_timestamp() -> str:
    """Generate timestamp for report filename."""
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


class ExcelReportGenerator:
    """Generate comprehensive DQ Excel reports with professional formatting."""

    def __init__(
        self,
        results_df: pd.DataFrame,
        rulebook: Dict,
        all_columns: List[str],
        column_scores: Dict[str, float],
        overall_score: float,
        dimension_scores: Dict[str, float],
        duplicate_combinations: Dict[str, List[Tuple]] = None,
    ):
        """
        Initialize the report generator.
        
        Args:
            results_df: DataFrame with validation results
            rulebook: Dictionary containing validation rules
            all_columns: List of all column names in dataset
            column_scores: Dictionary mapping column names to quality scores
            overall_score: Overall data quality score (0-100)
            dimension_scores: Dictionary mapping dimension names to scores
            duplicate_combinations: Dictionary of duplicate record combinations
        """
        self.results_df = results_df
        self.rulebook = rulebook
        self.all_columns = all_columns
        self.column_scores = column_scores
        self.overall_score = overall_score
        self.dimension_scores = dimension_scores
        self.duplicate_combinations = duplicate_combinations or {}
        logger.info(f"Report generator initialized with {len(results_df)} records")

    # ──────────────────────────────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────────────────────────────
    def generate_report(self, output_dir: Path) -> Path:
        """
        Build the complete Excel workbook and return its path.
        
        Args:
            output_dir: Directory where report should be saved
            
        Returns:
            Path to the generated Excel file
            
        Raises:
            IOError: If unable to write to output directory
            Exception: If workbook generation fails
        """
        try:
            output_path = output_dir / f"DQ_Assessment_Report.xlsx"
            output_dir.mkdir(parents=True, exist_ok=True)

            # Build df_out: drop internal helper columns, clean every cell
            internal = [
                c
                for c in ("_failed_columns_list", "_failed_rules_details")
                if c in self.results_df.columns
            ]
            df_out = self.results_df.drop(columns=internal, errors="ignore").copy()

            for col in df_out.columns:
                try:
                    df_out[col] = df_out[col].apply(clean_value)
                except Exception as e:
                    logger.warning(f"Error cleaning column {col}: {e}")
                    df_out[col] = df_out[col].astype(str)

            # Build failure indexes
            failed_tracker = defaultdict(set)
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

            # Write workbook
            import xlsxwriter
            
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                wb = writer.book
                fmt = self._make_formats(wb)

                self._sheet_dq_score(writer, fmt)
                self._sheet_results(writer, df_out, fmt)
                self._sheet_summary(writer, df_out, fmt, failed_tracker)
                self._sheet_dimension(writer, fmt)
                self._sheets_annexures(writer, df_out, fmt, failed_tracker)
                self._sheet_uniqueness(writer, df_out, fmt)
                self._sheet_completeness(writer, df_out, fmt, dimension_tracker)
                self._sheet_standardization(writer, df_out, fmt, dimension_tracker)

            logger.info(f"Report generated successfully: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Error generating report: {e}")
            raise

    def save_rulebook_json(self, output_dir: Path, rulebook: Dict = None) -> Path:
        """
        Save the rulebook as a JSON file.
        
        Args:
            output_dir: Directory where rulebook should be saved
            rulebook: Optional rulebook dict (uses self.rulebook if not provided)
            
        Returns:
            Path to the saved JSON file
        """
        try:
            if rulebook is None:
                rulebook = self.rulebook
                
            output_dir.mkdir(parents=True, exist_ok=True)
            path = output_dir / f"Rulebook_{get_timestamp()}.json"
            
            with open(path, "w", encoding="utf-8") as f:
                json.dump(rulebook, f, indent=2, default=str)
                
            logger.info(f"Rulebook saved: {path}")
            return path
        except Exception as e:
            logger.error(f"Error saving rulebook: {e}")
            raise

    # ──────────────────────────────────────────────────────────────────────
    # Format Definition
    # ──────────────────────────────────────────────────────────────────────
    @staticmethod
    def _make_formats(wb) -> Dict:
        """Create and return format dictionary for Excel cells."""
        return {
            "title": wb.add_format({
                "bold": True,
                "font_size": 16,
                "font_color": "#1F4E78",
            }),
            "subtitle": wb.add_format({
                "italic": True,
                "font_size": 11,
            }),
            "header": wb.add_format({
                "bold": True,
                "bg_color": "#4472C4",
                "font_color": "white",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }),
            "score": wb.add_format({
                "bold": True,
                "font_size": 20,
                "bg_color": "#E7E6E6",
                "border": 1,
                "align": "center",
            }),
            "metric": wb.add_format({
                "bold": True,
                "font_size": 12,
            }),
            "pass": wb.add_format({
                "bg_color": "#C6EFCE",
                "font_color": "#006100",
                "bold": True,
                "align": "center",
            }),
            "fail": wb.add_format({
                "bg_color": "#FFC7CE",
                "font_color": "#9C0006",
                "bold": True,
                "align": "center",
            }),
            "warning": wb.add_format({
                "bg_color": "#FFEB9C",
                "font_color": "#9C6500",
                "bold": True,
                "align": "center",
            }),
            "data": wb.add_format({
                "border": 1,
                "align": "left",
                "valign": "vcenter",
            }),
            "data_center": wb.add_format({
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }),
        }

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 1: DQ Score
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_dq_score(self, writer, fmt):
        """Create the DQ Score sheet with metrics and interpretation."""
        total = len(self.results_df)
        clean = int((self.results_df.get("Count of issues", 0) == 0).sum())
        data_cols = [
            c
            for c in self.all_columns
            if not c.startswith("_")
            and c not in ("Issues", "Count of issues", "Failed_Rules", "Failed_Columns", "Issue categories")
        ]

        metrics = [
            ("Overall DQ Score (%)", f"{self.overall_score:.2f}"),
            ("Total Records", total),
            ("Clean Records", clean),
            ("Records with Issues", total - clean),
            ("Total Columns", len(data_cols)),
            ("Columns at 100%", sum(1 for s in self.column_scores.values() if s == 100)),
            ("Columns with Issues", sum(1 for s in self.column_scores.values() if s < 100)),
            ("Total Rules Applied", len(self.rulebook.get("rules", []))),
        ]

        ws = writer.book.add_worksheet("DQ Score")
        ws.write(0, 0, "DATA QUALITY ASSESSMENT REPORT", fmt["title"])
        ws.write(1, 0, f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", fmt["subtitle"])
        ws.write(3, 0, "Metric", fmt["header"])
        ws.write(3, 1, "Value", fmt["header"])

        for i, (metric, value) in enumerate(metrics, start=4):
            ws.write(i, 0, metric, fmt["metric"])
            if isinstance(value, str) and value.replace(".", "", 1).isdigit():
                ws.write(i, 1, float(value), fmt["data_center"])
            else:
                ws.write(i, 1, value, fmt["data_center"])

        ws.write(13, 0, "Score Interpretation:", fmt["metric"])
        ws.write(14, 0, self._interpret(self.overall_score), fmt["subtitle"])

        ws.set_column(0, 0, 25)
        ws.set_column(1, 1, 20)

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 2: Results
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_results(self, writer, df_out, fmt):
        """Create the Results sheet with all processed data."""
        sheet = "Results"
        df_out.to_excel(writer, sheet_name=sheet, index=False)
        ws = writer.sheets[sheet]

        for col_num, col_name in enumerate(df_out.columns):
            ws.write(0, col_num, col_name, fmt["header"])
            ws.set_column(col_num, col_num, 18)

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 3: Summary
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_summary(self, writer, df_out, fmt, failed_tracker):
        """Create the Summary sheet with column-by-column pass/fail status."""
        summary_data = []
        for col in self.all_columns:
            if col.startswith("_") or col in ("Issues", "Count of issues", "Failed_Rules", 
                                              "Failed_Columns", "Issue categories"):
                continue

            score = self.column_scores.get(col, 0)
            failed_count = len(failed_tracker.get(col, set()))
            status = "PASS" if score == 100 else "FAIL"

            summary_data.append({
                "Column Name": col,
                "Quality Score (%)": f"{score:.2f}",
                "Failed Records": failed_count,
                "Status": status,
            })

        df_summary = pd.DataFrame(summary_data)

        sheet = "Summary"
        df_summary.to_excel(writer, sheet_name=sheet, index=False, startrow=3)
        ws = writer.sheets[sheet]

        ws.write(0, 0, "COLUMN-WISE VALIDATION SUMMARY", fmt["title"])
        ws.write(1, 0, f"Total Columns Analyzed: {len(summary_data)}", fmt["subtitle"])

        for col_num, col_name in enumerate(df_summary.columns):
            ws.write(3, col_num, col_name, fmt["header"])

        for row_num, row_data in df_summary.iterrows():
            status_format = fmt["pass"] if row_data["Status"] == "PASS" else fmt["fail"]
            ws.write(row_num + 4, 3, row_data["Status"], status_format)

            for col_num in range(3):
                ws.write(row_num + 4, col_num, row_data.iloc[col_num], fmt["data"])

        ws.set_column(0, 0, 25)
        ws.set_column(1, 3, 18)

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 4: Dimension Analysis
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_dimension(self, writer, fmt):
        """Create the Dimension Analysis sheet."""
        sheet = "Dimension Analysis"
        dimension_data = [
            {"Dimension": dim, "Quality Score (%)": f"{score:.2f}"}
            for dim, score in sorted(self.dimension_scores.items())
        ]

        df_dimension = pd.DataFrame(dimension_data)
        df_dimension.to_excel(writer, sheet_name=sheet, index=False, startrow=3)
        ws = writer.sheets[sheet]

        ws.write(0, 0, "QUALITY DIMENSIONS ANALYSIS", fmt["title"])
        ws.write(1, 0, f"Total Dimensions: {len(dimension_data)}", fmt["subtitle"])

        for col_num, col_name in enumerate(df_dimension.columns):
            ws.write(3, col_num, col_name, fmt["header"])

        for row_num, row_data in df_dimension.iterrows():
            for col_num, val in enumerate(row_data):
                ws.write(row_num + 4, col_num, val, fmt["data"])

        ws.set_column(0, 0, 25)
        ws.set_column(1, 1, 20)

    # ──────────────────────────────────────────────────────────────────────
    # Sheet 5..N: Annexures (one per column) - FIXED WORKSHEET NAMES
    # ──────────────────────────────────────────────────────────────────────
    def _sheets_annexures(self, writer, df_out, fmt, failed_tracker):
        """Create annexure sheets for each column with safe names (≤31 chars)."""
        data_cols = [
            c
            for c in self.all_columns
            if not c.startswith("_")
            and c not in ("Issues", "Count of issues", "Failed_Rules", "Failed_Columns", "Issue categories")
        ]

        for idx, col in enumerate(data_cols, 1):
            # Create safe sheet name (max 31 chars for Excel)
            safe_col_name = col[:20] if len(col) > 20 else col
            sheet = f"Annex_{idx}_{safe_col_name}"[:31]
            
            failed_indices = sorted(failed_tracker.get(col, set()))
            score = self.column_scores.get(col, 0)

            ws = writer.book.add_worksheet(sheet)
            ws.write(0, 0, f"VALIDATION DETAILS - {col.upper()}", fmt["title"])
            ws.write(1, 0, f"Quality Score: {score:.2f}%", fmt["subtitle"])
            ws.write(2, 0, "Status: " + ("✅ PASSED" if score == 100 else "❌ FAILED"),
                    fmt["pass"] if score == 100 else fmt["fail"])

            if failed_indices:
                annex_df = df_out.iloc[failed_indices].copy()
            else:
                annex_df = df_out.copy()

            annex_df.to_excel(writer, sheet_name=sheet, index=False, startrow=4)
            for col_num, col_name in enumerate(annex_df.columns):
                ws.write(4, col_num, col_name, fmt["header"])
                ws.set_column(col_num, col_num, 18)

    # ──────────────────────────────────────────────────────────────────────
    # Uniqueness Issues Sheet
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_uniqueness(self, writer, df_out, fmt):
        """Create the Uniqueness Issues sheet."""
        dup_records = set()
        dup_details = []
        
        for combo_str, groups in self.duplicate_combinations.items():
            for idx_list in groups:
                dup_records.update(idx_list)
                for idx in idx_list:
                    dup_details.append({
                        "Index": idx,
                        "Duplicate_Combination": combo_str,
                        "Duplicate_Count": len(idx_list),
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

            dup_df = dup_df.reset_index()
            dup_df = dup_df.merge(info_df, left_on="index", right_on="Index", how="left")
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
            ws.write(1, 0, "No duplicate records found", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])

    # ──────────────────────────────────────────────────────────────────────
    # Completeness Issues Sheet
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_completeness(self, writer, df_out, fmt, dimension_tracker):
        """Create the Completeness Issues sheet."""
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
                    "Index": idx,
                    "Incomplete_Columns": ", ".join(sorted(set(bad_cols))),
                })

            info_df = pd.DataFrame(info)
            comp_df = comp_df.reset_index()
            comp_df = comp_df.merge(info_df, left_on="index", right_on="Index", how="left")
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
            ws.write(1, 0, "No completeness issues found", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])

    # ──────────────────────────────────────────────────────────────────────
    # Standardization Issues Sheet
    # ──────────────────────────────────────────────────────────────────────
    def _sheet_standardization(self, writer, df_out, fmt, dimension_tracker):
        """Create the Standardization Issues sheet."""
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
                    if isinstance(rd, dict) and rd.get("dimension") in ("Standardization", "Validation"):
                        c = rd.get("column", "")
                        if c:
                            bad_cols.append(c)
                info.append({
                    "Index": idx,
                    "Non_Standard_Columns": ", ".join(sorted(set(bad_cols))),
                })

            info_df = pd.DataFrame(info)
            std_df = std_df.reset_index()
            std_df = std_df.merge(info_df, left_on="index", right_on="Index", how="left")
            std_df = std_df.drop(columns=["Index"])

            std_df.to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "STANDARDIZATION VALIDATION - FORMAT ISSUES", fmt["title"])
            ws.write(1, 0, f"Total Records with Standardization Issues: {len(indices)}", fmt["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", fmt["fail"])
            
            for c, h in enumerate(std_df.columns):
                ws.write(4, c, h, fmt["header"])
                ws.set_column(c, c, 20)
        else:
            pd.DataFrame().to_excel(writer, sheet_name=sheet, startrow=4, index=False)
            ws = writer.sheets[sheet]
            ws.write(0, 0, "STANDARDIZATION VALIDATION - FORMAT ISSUES", fmt["title"])
            ws.write(1, 0, "No standardization issues found", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])

    # ──────────────────────────────────────────────────────────────────────
    # Helper Methods
    # ──────────────────────────────────────────────────────────────────────
    @staticmethod
    def _interpret(score: float) -> str:
        """Generate human-readable interpretation of quality score."""
        if score >= 95:
            return "🎉 Excellent! Outstanding data quality."
        if score >= 80:
            return "👍 Good! Minor improvements needed."
        if score >= 60:
            return "⚠️ Fair! Significant improvements required."
        return "❌ Poor! Critical data quality issues detected."