"""
Report Generator Module
Generate comprehensive Excel reports with multiple sheets including specialized annexures
"""

import pandas as pd
import datetime
from pathlib import Path
from typing import Dict, List, Set, Tuple
from collections import defaultdict
import json
from modules.utils import clean_value, get_timestamp


class ExcelReportGenerator:
    """Generate comprehensive Excel reports"""
    
    def __init__(
        self, 
        results_df: pd.DataFrame, 
        rulebook: Dict, 
        all_columns: List[str],
        column_scores: Dict[str, float],
        overall_score: float,
        dimension_scores: Dict[str, float],
        duplicate_combinations: Dict[str, List[Tuple]] = None
    ):
        self.results_df = results_df
        self.rulebook = rulebook
        self.all_columns = all_columns
        self.column_scores = column_scores
        self.overall_score = overall_score
        self.dimension_scores = dimension_scores
        self.duplicate_combinations = duplicate_combinations or {}
    
    def generate_report(self, output_dir: Path) -> Path:
        """Generate complete Excel report"""
        
        timestamp = get_timestamp()
        output_path = output_dir / f"DQ_Report_{timestamp}.xlsx"
        
        # Prepare output dataframe - drop internal columns first
        columns_to_drop = ["_failed_columns_list"]
        if "_failed_rules_details" in self.results_df.columns:
            columns_to_drop.append("_failed_rules_details")
        
        df_out = self.results_df.drop(columns=columns_to_drop, errors="ignore")
        
        # Clean values column by column with error handling
        for col in df_out.columns:
            try:
                df_out[col] = df_out[col].apply(clean_value)
            except Exception as e:
                # If cleaning fails, convert to string as fallback
                print(f"Warning: Error cleaning column {col}: {e}")
                df_out[col] = df_out[col].astype(str)
        
        # Track failed columns and categorize by dimension
        failed_tracker = defaultdict(set)
        dimension_tracker = defaultdict(lambda: defaultdict(set))
        
        for idx, row in self.results_df.iterrows():
            # Get failed columns list
            failed_cols = row.get("_failed_columns_list", [])
            if isinstance(failed_cols, list):
                for col in failed_cols:
                    failed_tracker[col].add(idx)
            
            # Track by dimension
            failed_rules = row.get("_failed_rules_details", [])
            if isinstance(failed_rules, list):
                for rule_detail in failed_rules:
                    if isinstance(rule_detail, dict):
                        dimension = rule_detail.get("dimension", "General")
                        column = rule_detail.get("column")
                        if column:
                            dimension_tracker[dimension][column].add(idx)
        
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            wb = writer.book
            
            # Define formats
            formats = self._create_formats(wb)
            
            # Generate sheets
            self._create_score_sheet(writer, formats)
            self._create_results_sheet(writer, df_out, formats)
            self._create_summary_sheet(writer, formats, failed_tracker)
            if self.dimension_scores:
                self._create_dimension_sheet(writer, formats)
            
            # Column-wise annexures
            self._create_column_annexures(writer, df_out, formats, failed_tracker)
            
            # Specialized dimension annexures
            self._create_uniqueness_annexure(writer, df_out, formats)
            self._create_completeness_annexure(writer, df_out, formats, dimension_tracker)
            self._create_standardization_annexure(writer, df_out, formats, dimension_tracker)
        
        return output_path
    
    def save_rulebook_json(self, output_dir: Path, rulebook: Dict) -> Path:
        """Save rulebook as JSON"""
        timestamp = get_timestamp()
        rulebook_path = output_dir / f"Rulebook_{timestamp}.json"
        
        with open(rulebook_path, "w", encoding="utf-8") as f:
            json.dump(rulebook, f, indent=2)
        
        return rulebook_path
    
    def _create_formats(self, workbook):
        """Create Excel formats"""
        return {
            "header": workbook.add_format({
                "bold": True,
                "bg_color": "#4472C4",
                "font_color": "white",
                "border": 1,
                "align": "center",
                "valign": "vcenter"
            }),
            "title": workbook.add_format({
                "bold": True,
                "font_size": 16,
                "font_color": "#1F4E78"
            }),
            "subtitle": workbook.add_format({
                "italic": True,
                "font_size": 11
            }),
            "score": workbook.add_format({
                "bold": True,
                "font_size": 20,
                "bg_color": "#E7E6E6",
                "border": 1,
                "align": "center"
            }),
            "pass": workbook.add_format({
                "bg_color": "#C6EFCE",
                "font_color": "#006100",
                "bold": True,
                "align": "center"
            }),
            "fail": workbook.add_format({
                "bg_color": "#FFC7CE",
                "font_color": "#9C0006",
                "bold": True,
                "align": "center"
            }),
            "metric": workbook.add_format({
                "bold": True,
                "font_size": 12
            }),
            "warning": workbook.add_format({
                "bg_color": "#FFEB9C",
                "font_color": "#9C6500",
                "bold": True,
                "align": "center"
            })
        }
    
    def _create_score_sheet(self, writer, formats):
        """Create DQ Score sheet"""
        clean_records = len(self.results_df[self.results_df["Count of issues"] == 0])
        
        score_data = {
            "Metric": [
                "Overall DQ Score (%)",
                "Total Records",
                "Clean Records",
                "Records with Issues",
                "Total Columns",
                "Columns at 100%",
                "Columns with Issues",
                "Total Rules Applied"
            ],
            "Value": [
                self.overall_score,
                len(self.results_df),
                clean_records,
                len(self.results_df) - clean_records,
                len([c for c in self.all_columns if not c.startswith("_")]),
                sum(1 for s in self.column_scores.values() if s == 100),
                sum(1 for s in self.column_scores.values() if s < 100),
                len(self.rulebook.get("rules", []))
            ]
        }
        
        score_df = pd.DataFrame(score_data)
        score_df.to_excel(writer, sheet_name="DQ Score", index=False, startrow=3)
        
        ws = writer.sheets["DQ Score"]
        ws.write(0, 0, "DATA QUALITY ASSESSMENT REPORT", formats["title"])
        ws.write(1, 0, f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", formats["subtitle"])
        
        for c, col in enumerate(score_df.columns):
            ws.write(3, c, col, formats["header"])
            ws.set_column(c, c, 35)
        
        ws.write(4, 1, self.overall_score, formats["score"])
        
        # Interpretation
        ws.write(13, 0, "Score Interpretation:", formats["metric"])
        interpretation = self._get_score_interpretation(self.overall_score)
        ws.write(14, 0, interpretation)
    
    def _create_results_sheet(self, writer, df_out, formats):
        """Create Results sheet"""
        df_out.to_excel(writer, sheet_name="Results", index=False)
        
        ws = writer.sheets["Results"]
        for c, col in enumerate(df_out.columns):
            ws.write(0, c, col, formats["header"])
            ws.set_column(c, c, 25)
    
    def _create_summary_sheet(self, writer, formats, failed_tracker):
        """Create Summary sheet"""
        summary_data = []
        total = len(self.results_df)
        
        for col in self.all_columns:
            if col.startswith("_") or col in [
                "Issues", "Count of issues", "Failed_Rules", 
                "Failed_Columns", "Issue categories"
            ]:
                continue
            
            failed_count = len(failed_tracker.get(col, set()))
            clean_count = total - failed_count
            score = self.column_scores.get(col, 100)
            
            summary_data.append({
                "Column": col,
                "Total Records": total,
                "Failed": failed_count,
                "Passed": clean_count,
                "DQ Score (%)": score,
                "Status": "PASSED" if failed_count == 0 else "FAILED"
            })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        
        ws = writer.sheets["Summary"]
        for c, col in enumerate(summary_df.columns):
            ws.write(0, c, col, formats["header"])
            ws.set_column(c, c, 20)
        
        # Color code status
        for row_idx in range(len(summary_df)):
            status = summary_df.iloc[row_idx]["Status"]
            fmt = formats["pass"] if status == "PASSED" else formats["fail"]
            ws.write(row_idx + 1, 5, status, fmt)
    
    def _create_dimension_sheet(self, writer, formats):
        """Create dimension analysis sheet"""
        dim_data = []
        
        for dimension, score in self.dimension_scores.items():
            dim_data.append({
                "Dimension": dimension,
                "DQ Score (%)": score,
                "Status": "PASSED" if score == 100 else "FAILED"
            })
        
        dim_df = pd.DataFrame(dim_data)
        dim_df.to_excel(writer, sheet_name="Dimension Analysis", index=False)
        
        ws = writer.sheets["Dimension Analysis"]
        for c, col in enumerate(dim_df.columns):
            ws.write(0, c, col, formats["header"])
            ws.set_column(c, c, 25)
    
    def _create_column_annexures(self, writer, df_out, formats, failed_tracker):
        """Create annexure for each column"""
        annex_no = 1
        
        for col in self.all_columns:
            if col.startswith("_") or col in [
                "Issues", "Count of issues", "Failed_Rules", 
                "Failed_Columns", "Issue categories"
            ]:
                continue
            
            failed_indices = failed_tracker.get(col, set())
            
            if failed_indices:
                annex_df = df_out.iloc[list(failed_indices)]
                subtitle = f"Records with issues: {len(failed_indices)} | DQ Score: {self.column_scores.get(col, 0)}%"
                status_fmt = formats["fail"]
            else:
                annex_df = df_out.iloc[0:0]
                subtitle = "No issues found | DQ Score: 100%"
                status_fmt = formats["pass"]
            
            sheet_name = f"Annexure {annex_no}"[:31]
            annex_df.to_excel(writer, sheet_name=sheet_name, startrow=4, index=False)
            
            ws = writer.sheets[sheet_name]
            ws.write(0, 0, f"Annexure {annex_no}: {col}", formats["title"])
            ws.write(1, 0, subtitle, formats["subtitle"])
            ws.write(2, 0, f"Status: {'✅ PASSED' if not failed_indices else '❌ FAILED'}", status_fmt)
            
            if not annex_df.empty:
                for c, h in enumerate(annex_df.columns):
                    ws.write(4, c, h, formats["header"])
                    ws.set_column(c, c, 20)
            
            annex_no += 1
    
    def _create_uniqueness_annexure(self, writer, df_out, formats):
        """Create specialized annexure for uniqueness violations (duplicates)"""
        
        # Find all duplicate records across all combination columns
        duplicate_records = set()
        duplicate_details = []
        
        for combo_str, duplicate_indices in self.duplicate_combinations.items():
            for indices in duplicate_indices:
                duplicate_records.update(indices)
                # Store details about which combination was duplicated
                for idx in indices:
                    duplicate_details.append({
                        "Index": idx,
                        "Duplicate_Combination": combo_str,
                        "Duplicate_Count": len(indices)
                    })
        
        if duplicate_records:
            # Create dataframe with duplicate records
            dup_df = df_out.iloc[list(sorted(duplicate_records))].copy()
            
            # Add duplicate information
            dup_info_df = pd.DataFrame(duplicate_details)
            dup_info_df = dup_info_df.groupby("Index").agg({
                "Duplicate_Combination": lambda x: " | ".join(x),
                "Duplicate_Count": "max"
            }).reset_index()
            
            # Merge with original data
            dup_df = dup_df.reset_index()
            dup_df = dup_df.merge(dup_info_df, left_on="index", right_on="Index", how="left")
            dup_df = dup_df.drop(columns=["Index"])
            
            sheet_name = "Uniqueness Issues"
            dup_df.to_excel(writer, sheet_name=sheet_name, startrow=4, index=False)
            
            ws = writer.sheets[sheet_name]
            ws.write(0, 0, "UNIQUENESS VALIDATION - DUPLICATE RECORDS", formats["title"])
            ws.write(1, 0, f"Total Duplicate Records: {len(duplicate_records)}", formats["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", formats["fail"])
            
            for c, h in enumerate(dup_df.columns):
                ws.write(4, c, h, formats["header"])
                ws.set_column(c, c, 20)
        else:
            # No duplicates found
            sheet_name = "Uniqueness Issues"
            empty_df = pd.DataFrame()
            empty_df.to_excel(writer, sheet_name=sheet_name, startrow=4, index=False)
            
            ws = writer.sheets[sheet_name]
            ws.write(0, 0, "UNIQUENESS VALIDATION - DUPLICATE RECORDS", formats["title"])
            ws.write(1, 0, "No duplicate records found", formats["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", formats["pass"])
    
    def _create_completeness_annexure(self, writer, df_out, formats, dimension_tracker):
        """Create specialized annexure for completeness issues (null/empty values)"""
        
        completeness_indices = set()
        if "Completeness" in dimension_tracker:
            for col, indices in dimension_tracker["Completeness"].items():
                completeness_indices.update(indices)
        
        if completeness_indices:
            comp_df = df_out.iloc[list(sorted(completeness_indices))].copy()
            
            # Add information about which columns have completeness issues
            comp_info = []
            for idx in sorted(completeness_indices):
                failed_cols = []
                row = self.results_df.iloc[idx]
                if "_failed_rules_details" in row:
                    for rule_detail in row["_failed_rules_details"]:
                        if rule_detail.get("dimension") == "Completeness":
                            failed_cols.append(rule_detail.get("column"))
                
                comp_info.append({
                    "Index": idx,
                    "Incomplete_Columns": ", ".join(set(failed_cols))
                })
            
            comp_info_df = pd.DataFrame(comp_info)
            comp_df = comp_df.reset_index()
            comp_df = comp_df.merge(comp_info_df, left_on="index", right_on="Index", how="left")
            comp_df = comp_df.drop(columns=["Index"])
            
            sheet_name = "Completeness Issues"
            comp_df.to_excel(writer, sheet_name=sheet_name, startrow=4, index=False)
            
            ws = writer.sheets[sheet_name]
            ws.write(0, 0, "COMPLETENESS VALIDATION - MISSING VALUES", formats["title"])
            ws.write(1, 0, f"Total Records with Missing Values: {len(completeness_indices)}", formats["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", formats["fail"])
            
            for c, h in enumerate(comp_df.columns):
                ws.write(4, c, h, formats["header"])
                ws.set_column(c, c, 20)
        else:
            sheet_name = "Completeness Issues"
            empty_df = pd.DataFrame()
            empty_df.to_excel(writer, sheet_name=sheet_name, startrow=4, index=False)
            
            ws = writer.sheets[sheet_name]
            ws.write(0, 0, "COMPLETENESS VALIDATION - MISSING VALUES", formats["title"])
            ws.write(1, 0, "No completeness issues found", formats["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", formats["pass"])
    
    def _create_standardization_annexure(self, writer, df_out, formats, dimension_tracker):
        """Create specialized annexure for standardization issues"""
        
        standardization_indices = set()
        if "Standardization" in dimension_tracker:
            for col, indices in dimension_tracker["Standardization"].items():
                standardization_indices.update(indices)
        
        if standardization_indices:
            std_df = df_out.iloc[list(sorted(standardization_indices))].copy()
            
            # Add information about which columns have standardization issues
            std_info = []
            for idx in sorted(standardization_indices):
                failed_cols = []
                row = self.results_df.iloc[idx]
                if "_failed_rules_details" in row:
                    for rule_detail in row["_failed_rules_details"]:
                        if rule_detail.get("dimension") == "Standardization":
                            failed_cols.append(rule_detail.get("column"))
                
                std_info.append({
                    "Index": idx,
                    "Non_Standard_Columns": ", ".join(set(failed_cols))
                })
            
            std_info_df = pd.DataFrame(std_info)
            std_df = std_df.reset_index()
            std_df = std_df.merge(std_info_df, left_on="index", right_on="Index", how="left")
            std_df = std_df.drop(columns=["Index"])
            
            sheet_name = "Standardization Issues"
            std_df.to_excel(writer, sheet_name=sheet_name, startrow=4, index=False)
            
            ws = writer.sheets[sheet_name]
            ws.write(0, 0, "STANDARDIZATION VALIDATION - FORMAT ISSUES", formats["title"])
            ws.write(1, 0, f"Total Records with Standardization Issues: {len(standardization_indices)}", formats["subtitle"])
            ws.write(2, 0, "Status: ❌ FAILED", formats["fail"])
            
            for c, h in enumerate(std_df.columns):
                ws.write(4, c, h, formats["header"])
                ws.set_column(c, c, 20)
        else:
            sheet_name = "Standardization Issues"
            empty_df = pd.DataFrame()
            empty_df.to_excel(writer, sheet_name=sheet_name, startrow=4, index=False)
            
            ws = writer.sheets[sheet_name]
            ws.write(0, 0, "STANDARDIZATION VALIDATION - FORMAT ISSUES", formats["title"])
            ws.write(1, 0, "No standardization issues found", formats["subtitle"])
            ws.write(2, 0, "Status: ✅ PASSED", formats["pass"])
    
    @staticmethod
    def _get_score_interpretation(score: float) -> str:
        """Get score interpretation"""
        if score >= 95:
            return "🎉 Excellent! Outstanding data quality."
        elif score >= 80:
            return "👍 Good! Minor improvements needed."
        elif score >= 60:
            return "⚠️ Fair! Significant improvements required."
        else:
            return "❌ Poor! Critical data quality issues detected."