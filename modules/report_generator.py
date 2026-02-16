import json
import datetime
import pandas as pd
import numpy as np
import logging
from pathlib import Path
from typing import Dict, List, Tuple
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
    """Generate clean, clear DQ reports."""

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
        """Initialize the report generator."""
        self.results_df = results_df
        self.rulebook = rulebook
        self.all_columns = all_columns
        self.column_scores = column_scores
        self.overall_score = overall_score
        self.dimension_scores = dimension_scores
        self.duplicate_combinations = duplicate_combinations or {}
        logger.info(f"Report generator initialized with {len(results_df)} records")

    def generate_report(self, output_dir: Path) -> Path:
        """Build the complete Excel workbook and return its path."""
        try:
            output_path = output_dir / f"DQ_Assessment_Report.xlsx"
            output_dir.mkdir(parents=True, exist_ok=True)

            # Build clean dataset
            internal = [
                c for c in ("_failed_columns_list", "_failed_rules_details")
                if c in self.results_df.columns
            ]
            df_out = self.results_df.drop(columns=internal, errors="ignore").copy()

            for col in df_out.columns:
                try:
                    df_out[col] = df_out[col].apply(clean_value)
                except Exception as e:
                    logger.warning(f"Error cleaning column {col}: {e}")
                    df_out[col] = df_out[col].astype(str)

            # Build trackers
            failed_tracker = self._build_failed_tracker()
            dimension_tracker = self._build_dimension_tracker()
            uniqueness_failures = self._build_uniqueness_failures()

            # Write workbook
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                wb = writer.book
                fmt = self._make_formats(wb)

                # Create sheets in order
                self._sheet_dq_score(writer, fmt)
                self._sheet_summary(writer, df_out, fmt, failed_tracker)
                self._sheet_results(writer, df_out, fmt)
                self._sheet_dimension(writer, fmt)
                self._sheet_duplicate_summary(writer, fmt)
                self._sheets_annexures(writer, df_out, fmt, failed_tracker)
                self._sheet_uniqueness(writer, df_out, fmt, uniqueness_failures)
                self._sheet_completeness(writer, df_out, fmt, dimension_tracker)
                self._sheet_standardization(writer, df_out, fmt, dimension_tracker)

            logger.info(f"Report generated successfully: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Error generating report: {e}")
            raise

    def save_rulebook_json(self, output_dir: Path, rulebook: Dict = None) -> Path:
        """Save the rulebook as a JSON file."""
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

    def _build_failed_tracker(self) -> Dict[str, set]:
        """Build failed column tracker from results."""
        failed_tracker = defaultdict(set)
        for idx, row in self.results_df.iterrows():
            failed_cols = row.get("_failed_columns_list", [])
            if isinstance(failed_cols, list):
                for col in failed_cols:
                    failed_tracker[col].add(idx)
        return failed_tracker

    def _build_dimension_tracker(self) -> Dict[str, Dict[str, set]]:
        """Build dimension tracker from failed rules details."""
        dimension_tracker = defaultdict(lambda: defaultdict(set))
        for idx, row in self.results_df.iterrows():
            failed_rules = row.get("_failed_rules_details", [])
            if isinstance(failed_rules, list):
                for rd in failed_rules:
                    if isinstance(rd, dict):
                        dim = rd.get("dimension", "General")
                        col = rd.get("column")
                        if col:
                            dimension_tracker[dim][col].add(idx)
        return dimension_tracker

    def _build_uniqueness_failures(self) -> Dict[str, List[int]]:
        """Extract uniqueness failures from results."""
        uniqueness_failures = defaultdict(list)
        for idx, row in self.results_df.iterrows():
            failed_rules = row.get("_failed_rules_details", [])
            if isinstance(failed_rules, list):
                for rd in failed_rules:
                    if isinstance(rd, dict):
                        rule_type = rd.get("rule_type", "")
                        if rule_type == "uniqueness":
                            col = rd.get("column")
                            if col:
                                uniqueness_failures[col].append(idx)
        return uniqueness_failures

    @staticmethod
    def _make_formats(wb) -> Dict:
        """Create Excel cell formats."""
        return {
            "title": wb.add_format({"bold": True, "font_size": 14, "bg_color": "#4472C4", "font_color": "white", "align": "center", "valign": "vcenter", "border": 1}),
            "subtitle": wb.add_format({"bold": True, "font_size": 11, "bg_color": "#D9E1F2", "align": "left", "border": 1}),
            "header": wb.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "white", "border": 1, "align": "center", "valign": "vcenter"}),
            "pass": wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "bold": True, "align": "center", "border": 1}),
            "fail": wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True, "align": "center", "border": 1}),
            "warning": wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500", "bold": True, "align": "center", "border": 1}),
            "data": wb.add_format({"border": 1, "align": "left", "valign": "vcenter"}),
            "data_center": wb.add_format({"border": 1, "align": "center", "valign": "vcenter"}),
            "metric": wb.add_format({"bold": True, "align": "right", "border": 1}),
            "percentage": wb.add_format({"num_format": "0.00%", "align": "right", "border": 1})
        }

    def _sheet_dq_score(self, writer, fmt):
        """Create DQ Score sheet with overall metrics."""
        total = len(self.results_df)
        clean = int((self.results_df.get("Count of issues", 0) == 0).sum())
        data_cols = [c for c in self.all_columns if not c.startswith("_") and c not in ("Issues", "Count of issues", "Failed_Rules", "Failed_Columns", "Issue categories")]
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
            if isinstance(value, str) and value.replace(".", "", 1).replace("-", "", 1).isdigit():
                ws.write(i, 1, float(value), fmt["data_center"])
            else:
                ws.write(i, 1, value, fmt["data_center"])
        ws.write(13, 0, "Score Interpretation:", fmt["metric"])
        ws.write(14, 0, self._interpret(self.overall_score), fmt["subtitle"])
        ws.set_column(0, 0, 30)
        ws.set_column(1, 1, 20)

    def _sheet_summary(self, writer, df_out, fmt, failed_tracker):
        """Create Summary sheet with column pass/fail status."""
        summary_data = []
        for col in self.all_columns:
            if col.startswith("_") or col in ("Issues", "Count of issues", "Failed_Rules", "Failed_Columns", "Issue categories"):
                continue
            score = self.column_scores.get(col, 0)
            failed_count = len(failed_tracker.get(col, set()))
            status = "PASS" if score == 100 else "FAIL"
            summary_data.append({"Column Name": col, "Quality Score (%)": f"{score:.2f}", "Failed Records": failed_count, "Status": status})
        df_summary = pd.DataFrame(summary_data)
        ws = writer.book.add_worksheet("Summary")
        ws.write(0, 0, "COLUMN-WISE VALIDATION SUMMARY", fmt["title"])
        ws.write(1, 0, f"Total Columns: {len(summary_data)}", fmt["subtitle"])
        for col_num, col_name in enumerate(df_summary.columns):
            ws.write(3, col_num, col_name, fmt["header"])
        for row_num, row_data in df_summary.iterrows():
            status_format = fmt["pass"] if row_data["Status"] == "PASS" else fmt["fail"]
            ws.write(row_num + 4, 0, row_data["Column Name"], fmt["data"])
            ws.write(row_num + 4, 1, row_data["Quality Score (%)"], fmt["data_center"])
            ws.write(row_num + 4, 2, row_data["Failed Records"], fmt["data_center"])
            ws.write(row_num + 4, 3, row_data["Status"], status_format)
        ws.set_column(0, 0, 25)
        ws.set_column(1, 3, 18)

    def _sheet_results(self, writer, df_out, fmt):
        """Create Results sheet with all processed data."""
        ws = writer.book.add_worksheet("Results")
        display_cols = [c for c in df_out.columns if not c.startswith("_")]
        for col_num, col_name in enumerate(display_cols):
            ws.write(0, col_num, col_name, fmt["header"])
            ws.set_column(col_num, col_num, 18)
        for row_num, (_, row) in enumerate(df_out[display_cols].head(5000).iterrows(), 1):
            for col_num, value in enumerate(row):
                ws.write(row_num, col_num, str(value) if value is not None else "", fmt["data"])
        ws.freeze_panes(1, 0)

    def _sheet_dimension(self, writer, fmt):
        """Create Dimension Analysis sheet."""
        if not self.dimension_scores:
            return
        ws = writer.book.add_worksheet("Dimension Analysis")
        ws.set_column(0, 0, 30)
        ws.set_column(1, 1, 20)
        ws.write(0, 0, "DIMENSION-WISE ANALYSIS", fmt["title"])
        ws.write(1, 0, f"Total Dimensions: {len(self.dimension_scores)}", fmt["subtitle"])
        ws.write(3, 0, "Dimension", fmt["header"])
        ws.write(3, 1, "DQ Score (%)", fmt["header"])
        for row_num, (dim, score) in enumerate(sorted(self.dimension_scores.items()), start=4):
            ws.write(row_num, 0, dim, fmt["data"])
            status_fmt = fmt["pass"] if score == 100 else (fmt["warning"] if score >= 80 else fmt["fail"])
            ws.write(row_num, 1, f"{score:.2f}%", status_fmt)

    def _calculate_duplicate_count_per_column(self) -> Dict[str, int]:
        """Calculate total duplicate records per column based on validation rules."""
        duplicate_count = {}
        
        # Check if rulebook has rules
        if not self.rulebook or "rules" not in self.rulebook:
            return duplicate_count
        
        rules = self.rulebook.get("rules", [])
        
        # Iterate through rules to find columns with uniqueness/duplicate validation
        for rule in rules:
            if isinstance(rule, dict):
                # Check if rule is for uniqueness/duplicates
                rule_type = rule.get("rule_type", "").lower()
                column = rule.get("column")
                
                if rule_type in ["uniqueness", "duplicate"] and column:
                    # Count duplicates in this column from results_df
                    if column in self.results_df.columns:
                        dup_count = self.results_df[column].duplicated().sum()
                        if dup_count > 0:
                            duplicate_count[column] = int(dup_count)
        
        return duplicate_count

    def _sheet_duplicate_summary(self, writer, fmt):
        """Create Duplicate Summary sheet showing count per column with validation rules."""
        ws = writer.book.add_worksheet("Duplicate Summary")
        ws.write(0, 0, "DUPLICATE COUNT SUMMARY BY COLUMN", fmt["title"])
        
        # Get duplicate counts for columns with rules
        dup_counts = self._calculate_duplicate_count_per_column()
        
        if not dup_counts:
            ws.write(1, 0, "No duplicate validation rules found or no duplicates detected", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])
            return
        
        ws.write(1, 0, f"Columns with Duplicate Rules: {len(dup_counts)}", fmt["subtitle"])
        
        # Calculate total duplicates
        total_duplicates = sum(dup_counts.values())
        ws.write(2, 0, f"Total Duplicate Records: {total_duplicates}", fmt["subtitle"])
        
        # Write headers
        headers = ["Column Name", "Duplicate Count", "Status"]
        for col_num, header in enumerate(headers):
            ws.write(4, col_num, header, fmt["header"])
        
        ws.set_column(0, 0, 30)
        ws.set_column(1, 1, 20)
        ws.set_column(2, 2, 15)
        
        # Write data
        row_num = 5
        for col_name in sorted(dup_counts.keys()):
            dup_count = dup_counts[col_name]
            status = "FAIL" if dup_count > 0 else "PASS"
            status_fmt = fmt["fail"] if dup_count > 0 else fmt["pass"]
            
            ws.write(row_num, 0, col_name, fmt["data"])
            ws.write(row_num, 1, dup_count, fmt["data_center"])
            ws.write(row_num, 2, status, status_fmt)
            row_num += 1

    def _sheets_annexures(self, writer, df_out, fmt, failed_tracker):
        """Create SEPARATE annexure sheets for columns with ISSUES ONLY."""
        workbook = writer.book
        columns_with_issues = [c for c in self.all_columns if c in failed_tracker and len(failed_tracker[c]) > 0 and not c.startswith("_") and c not in ("Issues", "Count of issues", "Failed_Rules", "Failed_Columns", "Issue categories")]
        logger.info(f"Creating {len(columns_with_issues)} column annexures")
        for idx, col in enumerate(columns_with_issues, 1):
            try:
                safe_col_name = col[:20] if len(col) > 20 else col
                sheet_name = f"Annex_{idx}_{safe_col_name}"[:31]
                failed_indices = sorted(failed_tracker.get(col, set()))
                score = self.column_scores.get(col, 0)
                ws = workbook.add_worksheet(sheet_name)
                ws.write(0, 0, f"COLUMN VALIDATION DETAILS: {col.upper()}", fmt["title"])
                ws.write(1, 0, f"Quality Score: {score:.2f}%", fmt["subtitle"])
                ws.write(2, 0, f"Failed Records: {len(failed_indices)} out of {len(df_out)}", fmt["subtitle"])
                status_text = "✅ PASSED - NO ISSUES" if score == 100 else "❌ FAILED - SEE BELOW"
                ws.write(3, 0, f"Status: {status_text}", fmt["pass"] if score == 100 else fmt["fail"])
                if failed_indices:
                    annex_df = df_out.iloc[failed_indices].copy()
                    annex_df.insert(0, "Row_Index", failed_indices)
                    display_cols = list(annex_df.columns)
                    for col_num, col_name in enumerate(display_cols):
                        ws.write(5, col_num, col_name, fmt["header"])
                        ws.set_column(col_num, col_num, 18)
                    for row_num, (_, row_data) in enumerate(annex_df[display_cols].iterrows(), 6):
                        for col_num, value in enumerate(row_data):
                            ws.write(row_num, col_num, str(value) if value is not None else "", fmt["data"])
                    ws.freeze_panes(6, 0)
                else:
                    ws.write(5, 0, "No failed records for this column.", fmt["subtitle"])
            except Exception as e:
                logger.warning(f"Error creating annexure for '{col}': {str(e)}")
                continue

    def _sheet_uniqueness(self, writer, df_out, fmt, uniqueness_failures):
        """
        Create Uniqueness Issues sheet - ULTRA SIMPLIFIED
        No pandas merges - direct indexing only
        """
        ws = writer.book.add_worksheet("Uniqueness Issues")
        ws.write(0, 0, "UNIQUENESS VALIDATION - DUPLICATE RECORDS", fmt["title"])
        
        # Check if there are any uniqueness failures
        if not uniqueness_failures or len(uniqueness_failures) == 0:
            ws.write(1, 0, "No duplicate records found", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])
            return
        
        # Collect all unique failure records
        all_dup_indices = set()
        col_map = {}  # Map index -> list of failed columns
        
        for col, indices in uniqueness_failures.items():
            for idx in indices:
                all_dup_indices.add(idx)
                if idx not in col_map:
                    col_map[idx] = []
                col_map[idx].append(col)
        
        if len(all_dup_indices) == 0:
            ws.write(1, 0, "No duplicate records found", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])
            return
        
        # Write header
        ws.write(1, 0, f"Total Duplicate Records: {len(all_dup_indices)}", fmt["subtitle"])
        ws.write(2, 0, "Status: ❌ FAILED - DUPLICATES FOUND", fmt["fail"])
        
        # Setup column headers
        display_cols = [c for c in df_out.columns if not c.startswith("_")]
        header_cols = ["Row_Index", "Failed_Column", "Issue_Type"] + display_cols
        
        # Write headers (row 4)
        for col_num, col_name in enumerate(header_cols):
            ws.write(4, col_num, col_name, fmt["header"])
            ws.set_column(col_num, col_num, 18)
        
        # Write duplicate records - ULTRA SIMPLE approach
        sorted_indices = sorted(all_dup_indices)
        
        for write_row_num, orig_idx in enumerate(sorted_indices, 5):
            try:
                # Row index
                ws.write(write_row_num, 0, orig_idx, fmt["data"])
                
                # Failed columns
                failed_cols = col_map.get(orig_idx, [])
                ws.write(write_row_num, 1, ", ".join(failed_cols), fmt["data"])
                
                # Issue type
                ws.write(write_row_num, 2, "Uniqueness Violation", fmt["fail"])
                
                # Record data - direct iloc access
                record = df_out.iloc[orig_idx]
                for col_num, col_name in enumerate(display_cols, 3):
                    value = record[col_name]
                    ws.write(write_row_num, col_num, str(value) if value is not None else "", fmt["data"])
                    
            except Exception as e:
                logger.warning(f"Error writing row {orig_idx}: {e}")
                continue

    def _sheet_completeness(self, writer, df_out, fmt, dimension_tracker):
        """Create Completeness Issues sheet."""
        indices = set()
        for col_set in dimension_tracker.get("Completeness", {}).values():
            indices.update(col_set)
        ws = writer.book.add_worksheet("Completeness Issues")
        ws.write(0, 0, "COMPLETENESS VALIDATION - MISSING VALUES", fmt["title"])
        
        if indices:
            sorted_indices = sorted(indices)
            ws.write(1, 0, f"Total Records with Missing Values: {len(indices)}", fmt["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", fmt["fail"])
            
            display_cols = [c for c in df_out.columns if not c.startswith("_")]
            header_cols = ["Row_Index", "Incomplete_Columns"] + display_cols
            
            for col_num, col_name in enumerate(header_cols):
                ws.write(4, col_num, col_name, fmt["header"])
                ws.set_column(col_num, col_num, 20)
            
            for write_row_num, idx in enumerate(sorted_indices, 5):
                try:
                    # Row index
                    ws.write(write_row_num, 0, idx, fmt["data"])
                    
                    # Get incomplete columns
                    bad_cols = []
                    row = self.results_df.iloc[idx]
                    for rd in (row.get("_failed_rules_details") or []):
                        if isinstance(rd, dict) and rd.get("dimension") == "Completeness":
                            c = rd.get("column", "")
                            if c:
                                bad_cols.append(c)
                    
                    ws.write(write_row_num, 1, ", ".join(sorted(set(bad_cols))), fmt["data"])
                    
                    # Record data
                    record = df_out.iloc[idx]
                    for col_num, col_name in enumerate(display_cols, 2):
                        value = record[col_name]
                        ws.write(write_row_num, col_num, str(value) if value is not None else "", fmt["data"])
                except Exception as e:
                    logger.warning(f"Error writing row {idx}: {e}")
                    continue
        else:
            ws.write(1, 0, "No completeness issues found", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])

    def _sheet_standardization(self, writer, df_out, fmt, dimension_tracker):
        """Create Standardization Issues sheet."""
        indices = set()
        for dim_key in ("Standardization", "Validation"):
            for col_set in dimension_tracker.get(dim_key, {}).values():
                indices.update(col_set)
        ws = writer.book.add_worksheet("Standardization Issues")
        ws.write(0, 0, "STANDARDIZATION VALIDATION - FORMAT ISSUES", fmt["title"])
        
        if indices:
            sorted_indices = sorted(indices)
            ws.write(1, 0, f"Total Records with Standardization Issues: {len(indices)}", fmt["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", fmt["fail"])
            
            display_cols = [c for c in df_out.columns if not c.startswith("_")]
            header_cols = ["Row_Index", "Non_Standard_Columns"] + display_cols
            
            for col_num, col_name in enumerate(header_cols):
                ws.write(4, col_num, col_name, fmt["header"])
                ws.set_column(col_num, col_num, 20)
            
            for write_row_num, idx in enumerate(sorted_indices, 5):
                try:
                    # Row index
                    ws.write(write_row_num, 0, idx, fmt["data"])
                    
                    # Get non-standard columns
                    bad_cols = []
                    row = self.results_df.iloc[idx]
                    for rd in (row.get("_failed_rules_details") or []):
                        if isinstance(rd, dict) and rd.get("dimension") in ("Standardization", "Validation"):
                            c = rd.get("column", "")
                            if c:
                                bad_cols.append(c)
                    
                    ws.write(write_row_num, 1, ", ".join(sorted(set(bad_cols))), fmt["data"])
                    
                    # Record data
                    record = df_out.iloc[idx]
                    for col_num, col_name in enumerate(display_cols, 2):
                        value = record[col_name]
                        ws.write(write_row_num, col_num, str(value) if value is not None else "", fmt["data"])
                except Exception as e:
                    logger.warning(f"Error writing row {idx}: {e}")
                    continue
        else:
            ws.write(1, 0, "No standardization issues found", fmt["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", fmt["pass"])

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