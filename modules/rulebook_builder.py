"""
Rulebook Builder Service
Converts rules datasets to JSON rulebook or loads existing JSON
Supports combination uniqueness rules (e.g., "column1 + column2")
"""

import pandas as pd
import json
from pathlib import Path
from typing import Dict, List, Any
from modules.config import RULE_ALIAS_MAP


class RulebookBuilderService:
    """Service for building and managing rulebooks"""
    
    def __init__(self):
        self.rule_alias_map = RULE_ALIAS_MAP
    
    def load_json_rulebook(self, file_path: Path) -> Dict:
        """Load existing JSON rulebook"""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                rulebook = json.load(f)
            
            if "rules" not in rulebook or not isinstance(rulebook["rules"], list):
                raise ValueError("Rulebook must contain 'rules' array")
            
            return rulebook
            
        except Exception as e:
            raise Exception(f"Error loading JSON rulebook: {str(e)}")
    
    def build_from_rules_dataset(
        self, 
        rules_df: pd.DataFrame, 
        base_columns: List[str]
    ) -> Dict:
        """
        Build rulebook from rules dataset
        
        Expected columns:
        - column_name: Target column (supports "col1 + col2" for combinations)
        - rule or rule_type: Type of validation
        - dimension or rule_category: DQ dimension
        - message: Error message
        - expression: Optional rule expression
        - severity: Optional severity level
        """
        rules = []
        
        # Determine column names (support both formats)
        col_name_field = self._detect_column_field(rules_df)
        rule_type_field = self._detect_rule_field(rules_df)
        dimension_field = self._detect_dimension_field(rules_df)
        
        for _, row in rules_df.iterrows():
            column = row.get(col_name_field)
            
            if not column or pd.isna(column):
                continue
            
            # Check if this is a combination rule (contains "+")
            if "+" in str(column):
                # This is a combination uniqueness rule
                rule = self._build_combination_rule(row, column, rule_type_field, dimension_field, base_columns)
                if rule:
                    rules.append(rule)
            else:
                # Standard single-column rule
                col = str(column).strip()
                if col in base_columns:
                    rule = self._build_single_rule(row, col, rule_type_field, dimension_field)
                    if rule:
                        rules.append(rule)
        
        import datetime
        return {
            "rules": rules,
            "metadata": {
                "created": datetime.datetime.now().isoformat(),
                "total_rules": len(rules),
                "source": "rules_dataset"
            }
        }
    
    def _detect_column_field(self, df: pd.DataFrame) -> str:
        """Detect which field contains column names"""
        if "column_name" in df.columns:
            return "column_name"
        elif "column" in df.columns:
            return "column"
        else:
            raise ValueError("Rules dataset must have 'column_name' or 'column' field")
    
    def _detect_rule_field(self, df: pd.DataFrame) -> str:
        """Detect which field contains rule type"""
        if "rule_type" in df.columns:
            return "rule_type"
        elif "rule" in df.columns:
            return "rule"
        else:
            raise ValueError("Rules dataset must have 'rule_type' or 'rule' field")
    
    def _detect_dimension_field(self, df: pd.DataFrame) -> str:
        """Detect which field contains dimension"""
        if "dimension" in df.columns:
            return "dimension"
        elif "rule_category" in df.columns:
            return "rule_category"
        else:
            return None  # Optional field
    
    def _build_combination_rule(
        self, 
        row: pd.Series, 
        column_combination: str, 
        rule_field: str, 
        dimension_field: str,
        base_columns: List[str]
    ) -> Dict:
        """Build a combination uniqueness rule"""
        
        # Parse columns from combination string
        columns = [c.strip() for c in str(column_combination).split("+")]
        
        # Validate that all columns exist in the dataset
        valid_columns = [c for c in columns if c in base_columns]
        if len(valid_columns) < 2:
            return None  # Need at least 2 valid columns for combination
        
        # Get rule type
        rule_value = row.get(rule_field)
        if not rule_value or pd.isna(rule_value):
            # Default to uniqueness for combinations
            rule_type = "uniqueness"
        else:
            rule_type = self._normalize_rule_type(str(rule_value))
        
        # Get dimension
        dimension = row.get(dimension_field, "Uniqueness") if dimension_field else "Uniqueness"
        if pd.isna(dimension):
            dimension = "Uniqueness"
        
        # Get message
        message = row.get("message")
        if not message or pd.isna(message):
            message = f"Combination {' + '.join(valid_columns)} should be unique"
        
        # Get severity
        severity = row.get("severity", "HIGH")
        if pd.isna(severity):
            severity = "HIGH"
        
        return {
            "rule_type": "uniqueness_combination",
            "columns": valid_columns,
            "dimension": str(dimension),
            "message": str(message),
            "severity": str(severity)
        }
    
    def _build_single_rule(
        self, 
        row: pd.Series, 
        column: str, 
        rule_field: str, 
        dimension_field: str
    ) -> Dict:
        """Build a single rule from row data"""
        
        rule_value = row.get(rule_field)
        if not rule_value or pd.isna(rule_value):
            return None
        
        # Normalize rule type
        rule_type = self._normalize_rule_type(str(rule_value))
        
        # Get dimension
        dimension = row.get(dimension_field, "General") if dimension_field else "General"
        if pd.isna(dimension):
            dimension = "General"
        
        # Get message
        message = row.get("message")
        if not message or pd.isna(message):
            message = f"{column}: {rule_type} validation failed"
        
        # Get expression
        expression = row.get("expression")
        if pd.isna(expression) or str(expression).lower() == "none":
            expression = None
        
        # Get severity
        severity = row.get("severity", "MEDIUM")
        if pd.isna(severity):
            severity = "MEDIUM"
        
        return {
            "column": column,
            "dimension": str(dimension),
            "rule_type": rule_type,
            "expression": expression,
            "message": str(message),
            "severity": str(severity)
        }
    
    def _normalize_rule_type(self, rule_text: str) -> str:
        """Normalize rule type using alias map"""
        rule_lower = rule_text.lower().strip()
        
        # Direct match
        if rule_lower in self.rule_alias_map:
            return self.rule_alias_map[rule_lower]
        
        # Pattern matching
        if "not null" in rule_lower or "not blank" in rule_lower:
            return "not_null"
        elif "unique" in rule_lower or "duplicate" in rule_lower:
            return "uniqueness"
        elif "email" in rule_lower:
            return "email_format"
        elif "regex" in rule_lower or "pattern" in rule_lower:
            return "regex"
        elif "numeric" in rule_lower:
            return "numeric_only"
        elif "alpha" in rule_lower:
            return "alpha_only"
        elif "special char" in rule_lower:
            return "no_special_chars"
        elif "date" in rule_lower:
            return "date_format"
        else:
            return rule_lower  # Use as-is
