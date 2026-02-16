"""
modules/rule_executor.py - CORRECTED VERSION
==============================================
Rule Executor Engine with proper uniqueness detection and failed column tracking

Key Improvements:
  1. Proper failed column tracking in _failed_columns_list
  2. Correct uniqueness detection for scientific notation values
  3. Detailed failure information in _failed_rules_details
  4. Clear dimension categorization
"""

import pandas as pd
import re
import datetime
from collections import defaultdict
from typing import Dict, Any, List, Tuple


class RuleExecutorEngine:
    """Execute validation rules dynamically with proper tracking"""
    
    def __init__(self, df: pd.DataFrame, rulebook: Dict):
        self.df = df
        self.rulebook = rulebook
        self.duplicate_cache = {}
        self.combination_duplicates = {}
        self._precompute_duplicates()
        self._precompute_combination_duplicates()
    
    def _precompute_duplicates(self):
        """
        Pre-compute duplicate indices for all columns.
        
        This is CRITICAL for proper uniqueness detection.
        Handles scientific notation and all data types.
        """
        for col in self.df.columns:
            value_counts = self.df[col].value_counts()
            duplicates = value_counts[value_counts > 1].index.tolist()
            
            dup_indices = set()
            for dup_val in duplicates:
                # Skip null/empty values
                if not self._is_null_or_empty(dup_val):
                    # This comparison works with scientific notation
                    indices = self.df[self.df[col] == dup_val].index.tolist()
                    dup_indices.update(indices)
            
            self.duplicate_cache[col] = dup_indices
    
    def _precompute_combination_duplicates(self):
        """Pre-compute duplicate indices for column combinations"""
        for rule in self.rulebook.get("rules", []):
            if rule.get("rule_type") == "uniqueness_combination":
                columns = rule.get("columns", [])
                if len(columns) >= 2:
                    # Create combination key
                    combo_key = " + ".join(columns)
                    
                    # Find duplicates for this combination
                    combination_values = self.df[columns].apply(
                        lambda row: tuple(row), axis=1
                    )
                    
                    value_counts = combination_values.value_counts()
                    duplicates = value_counts[value_counts > 1].index.tolist()
                    
                    duplicate_groups = []
                    for dup_val in duplicates:
                        # Check if any value in combination is null
                        if any(self._is_null_or_empty(v) for v in dup_val):
                            continue
                        
                        indices = self.df[
                            combination_values == dup_val
                        ].index.tolist()
                        
                        if len(indices) > 1:
                            duplicate_groups.append(indices)
                    
                    self.combination_duplicates[combo_key] = duplicate_groups
    
    def get_combination_duplicates(self) -> Dict[str, List[Tuple]]:
        """Return combination duplicates for reporting"""
        return self.combination_duplicates
    
    def execute_all_rules(self) -> pd.DataFrame:
        """
        Execute all rules and return results dataframe.
        
        IMPORTANT: This properly tracks:
        - _failed_columns_list: Which columns failed validation
        - _failed_rules_details: Details of each failure
        - Issue categories by dimension
        """
        results = []
        
        for idx, row in self.df.iterrows():
            row_issues = []
            row_failed_rules = []
            row_failed_columns = []  # CRITICAL: Track which columns failed
            row_dimensions = set()
            row_failed_details = []
            
            # Execute each rule
            for rule in self.rulebook.get("rules", []):
                result = self._execute_single_rule(row, rule, idx)
                
                # If rule failed
                if not result["passed"]:
                    row_issues.append(result["message"])
                    row_failed_rules.append(result["rule_type"])
                    
                    # CRITICAL: Add column to failed list
                    if "columns" in result and result["columns"]:
                        row_failed_columns.extend(result["columns"])
                    else:
                        row_failed_columns.append(result["column"])
                    
                    row_dimensions.add(result["dimension"])
                    
                    # Store detailed rule information
                    row_failed_details.append({
                        "column": result.get("column") or " + ".join(result.get("columns", [])),
                        "rule_type": result["rule_type"],
                        "dimension": result["dimension"],
                        "message": result["message"]
                    })
            
            # Build result row with tracking information
            result_row = row.to_dict()
            result_row["Issues"] = " | ".join(row_issues) if row_issues else ""
            result_row["Count of issues"] = len(row_issues)
            result_row["Failed_Rules"] = ", ".join(set(row_failed_rules))
            result_row["Failed_Columns"] = ", ".join(set(row_failed_columns))
            result_row["Issue categories"] = ", ".join(sorted(row_dimensions))
            
            # CRITICAL: Store for report generation
            result_row["_failed_columns_list"] = list(set(row_failed_columns))
            result_row["_failed_rules_details"] = row_failed_details
            
            results.append(result_row)
        
        return pd.DataFrame(results)
    
    def _execute_single_rule(self, row: pd.Series, rule: Dict, row_idx: int) -> Dict:
        """
        Execute a single validation rule.
        
        Returns: Dictionary with passed/failed status and details
        """
        rule_type = rule.get("rule_type")
        message = rule.get("message", "Validation failed")
        dimension = rule.get("dimension", "General")
        
        # Handle combination uniqueness separately
        if rule_type == "uniqueness_combination":
            return self._execute_combination_uniqueness(row, rule, row_idx)
        
        # Standard single-column rules
        column = rule.get("column")
        expression = rule.get("expression")
        
        if column not in row.index:
            return {
                "passed": True,
                "message": "",
                "rule_type": rule_type,
                "column": column,
                "dimension": dimension
            }
        
        value = row[column]
        passed = True
        
        try:
            if rule_type == "not_null":
                # Check if value is null/empty
                passed = not self._is_null_or_empty(value)
            
            elif rule_type == "uniqueness":
                # CRITICAL: Check if this row's index is in the duplicate cache
                passed = row_idx not in self.duplicate_cache.get(column, set())
            
            elif rule_type == "regex":
                if not self._is_null_or_empty(value) and expression:
                    passed = bool(re.match(str(expression), str(value)))
            
            elif rule_type == "allowed_values":
                if not self._is_null_or_empty(value) and expression:
                    allowed = [v.strip() for v in str(expression).split(",")]
                    passed = str(value) in allowed
            
            elif rule_type == "range":
                if not self._is_null_or_empty(value) and expression:
                    try:
                        num_val = float(value)
                        min_val, max_val = map(float, str(expression).split(","))
                        passed = min_val <= num_val <= max_val
                    except:
                        passed = False
            
            elif rule_type == "length":
                if not self._is_null_or_empty(value) and expression:
                    try:
                        if "," in str(expression):
                            min_len, max_len = map(int, str(expression).split(","))
                            passed = min_len <= len(str(value)) <= max_len
                        else:
                            passed = len(str(value)) == int(expression)
                    except:
                        passed = False
            
            elif rule_type == "no_special_chars":
                if not self._is_null_or_empty(value):
                    pattern = expression if expression else r'[^A-Za-z0-9\s]'
                    passed = not bool(re.search(str(pattern), str(value)))
            
            elif rule_type == "email_format":
                if not self._is_null_or_empty(value):
                    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                    passed = bool(re.match(email_pattern, str(value)))
            
            elif rule_type == "numeric_only":
                if not self._is_null_or_empty(value):
                    try:
                        float(value)
                        passed = True
                    except:
                        passed = False
            
            elif rule_type == "alpha_only":
                if not self._is_null_or_empty(value):
                    passed = str(value).replace(" ", "").isalpha()
            
            elif rule_type == "date_format":
                if not self._is_null_or_empty(value):
                    try:
                        fmt = expression if expression else "%Y-%m-%d"
                        datetime.datetime.strptime(str(value), fmt)
                        passed = True
                    except:
                        passed = False
            
            elif rule_type == "contains":
                if not self._is_null_or_empty(value) and expression:
                    passed = str(expression) in str(value)
            
            elif rule_type == "not_contains":
                if not self._is_null_or_empty(value) and expression:
                    passed = str(expression) not in str(value)
            
            elif rule_type == "custom_expression":
                if expression:
                    passed = self._evaluate_safe_expression(value, expression)
            
            else:
                # Unknown rule type - pass by default
                passed = True
        
        except Exception as e:
            passed = False
            message = f"{message} (Error: {str(e)})"
        
        return {
            "passed": passed,
            "message": message if not passed else "",
            "rule_type": rule_type,
            "column": column,
            "dimension": dimension
        }
    
    def _execute_combination_uniqueness(self, row: pd.Series, rule: Dict, row_idx: int) -> Dict:
        """Execute uniqueness check for column combinations"""
        columns = rule.get("columns", [])
        message = rule.get("message", f"Duplicate combination found")
        dimension = rule.get("dimension", "Uniqueness")
        
        if not columns or len(columns) < 2:
            return {
                "passed": True,
                "message": "",
                "rule_type": "uniqueness_combination",
                "columns": columns,
                "dimension": dimension
            }
        
        # Check if this row is in any duplicate group for this combination
        combo_key = " + ".join(columns)
        duplicate_groups = self.combination_duplicates.get(combo_key, [])
        
        passed = True
        for dup_group in duplicate_groups:
            if row_idx in dup_group:
                passed = False
                break
        
        return {
            "passed": passed,
            "message": message if not passed else "",
            "rule_type": "uniqueness_combination",
            "columns": columns,
            "column": combo_key,
            "dimension": dimension
        }
    
    @staticmethod
    def _is_null_or_empty(value) -> bool:
        """Check if value is null or empty"""
        return (
            value is None 
            or pd.isna(value) 
            or str(value).strip() == "" 
            or str(value).lower() == "nan"
        )
    
    @staticmethod
    def _evaluate_safe_expression(value, expression: str) -> bool:
        """Safely evaluate custom expression"""
        try:
            safe_dict = {
                "value": value,
                "len": len,
                "str": str,
                "int": int,
                "float": float,
                "abs": abs,
                "min": min,
                "max": max,
            }
            
            if any(keyword in expression for keyword in ["import", "exec", "eval", "__", "open", "file"]):
                return False
            
            result = eval(expression, {"__builtins__": {}}, safe_dict)
            return bool(result)
        except:
            return False