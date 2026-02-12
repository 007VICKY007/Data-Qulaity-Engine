"""
Data Maturity Assessment Engine
Evaluates data maturity levels and dimensions
"""

import pandas as pd
from typing import Dict, Any


class DataMaturityEngine:
    """Data Maturity Assessment Engine"""
    
    MATURITY_LEVELS = {
        'Initial': (0, 20),
        'Repeatable': (20, 40),
        'Defined': (40, 60),
        'Managed': (60, 80),
        'Optimized': (80, 100)
    }
    
    def __init__(self, df: pd.DataFrame):
        self.df = df
    
    def run(self) -> Dict[str, Any]:
        """Run maturity assessment"""
        
        completeness = self._assess_completeness()
        structure = self._assess_structure()
        consistency = self._assess_consistency()
        standardization = self._assess_standardization()
        metadata = self._assess_metadata()
        
        dimension_scores = {
            'Completeness': completeness,
            'Structure': structure,
            'Consistency': consistency,
            'Standardization': standardization,
            'Metadata': metadata
        }
        
        overall_maturity = sum(dimension_scores.values()) / len(dimension_scores)
        maturity_level = self._get_maturity_level(overall_maturity)
        
        return {
            'overall_maturity': overall_maturity,
            'maturity_level': maturity_level,
            'completeness': completeness,
            'structure_score': structure,
            'consistency': consistency,
            'standardization': standardization,
            'metadata': metadata,
            'dimension_scores': dimension_scores
        }
    
    def _assess_completeness(self) -> float:
        """Assess data completeness"""
        total_cells = self.df.size
        if total_cells == 0:
            return 0.0
        
        filled_cells = self.df.notna().sum().sum()
        
        for col in self.df.columns:
            for val in self.df[col]:
                if isinstance(val, str) and val.strip() == "":
                    filled_cells -= 1
        
        return round((filled_cells / total_cells) * 100, 2)
    
    def _assess_structure(self) -> float:
        """Assess data structure quality"""
        score = 50
        
        bad_names = 0
        for col in self.df.columns:
            if not col or col.startswith('_') or col.startswith('Unnamed'):
                bad_names += 1
        
        name_score = ((len(self.df.columns) - bad_names) / len(self.df.columns)) * 100 if self.df.columns.size > 0 else 0
        
        type_consistency = 50
        consistency_count = 0
        
        for col in self.df.columns:
            non_null_values = self.df[col].dropna()
            if len(non_null_values) > 0:
                types = set(type(v).__name__ for v in non_null_values if v is not None and str(v).strip() != "")
                if len(types) <= 1:
                    consistency_count += 1
        
        if self.df.columns.size > 0:
            type_consistency = (consistency_count / len(self.df.columns)) * 100
        
        overall = (name_score + type_consistency) / 2
        return round(overall, 2)
    
    def _assess_consistency(self) -> float:
        """Assess data consistency"""
        if len(self.df) == 0:
            return 0.0
        
        total_cells = self.df.size
        duplicate_cells = 0
        
        for col in self.df.columns:
            non_null = self.df[col].dropna()
            if len(non_null) > 0:
                duplicates = non_null.duplicated().sum()
                duplicate_cells += duplicates
        
        consistency = ((total_cells - duplicate_cells) / total_cells) * 100
        return round(min(consistency, 100), 2)
    
    def _assess_standardization(self) -> float:
        """Assess data standardization"""
        score = 0
        checked_cols = 0
        
        for col in self.df.columns:
            non_null = self.df[col].dropna()
            if len(non_null) == 0:
                continue
            
            checked_cols += 1
            values = non_null.astype(str)
            
            trimmed = values.str.strip()
            untrimmed = values[values != trimmed].count()
            
            lowered = values.str.lower()
            original = values[values.str.lower() != values].count()
            
            col_score = ((len(values) - untrimmed - original) / len(values)) * 100
            score += col_score
        
        if checked_cols == 0:
            return 50.0
        
        return round(score / checked_cols, 2)
    
    def _assess_metadata(self) -> float:
        """Assess metadata quality"""
        score = 0
        
        meaningful_names = 0
        for col in self.df.columns:
            col_str = str(col).lower()
            if len(col) > 3 and col_str not in ['col', 'data', 'value', 'id', 'name']:
                meaningful_names += 1
        
        name_score = (meaningful_names / len(self.df.columns)) * 100 if self.df.columns.size > 0 else 0
        
        documentation_score = 60
        
        metadata_score = (name_score + documentation_score) / 2
        return round(metadata_score, 2)
    
    def _get_maturity_level(self, score: float) -> str:
        """Get maturity level name from score"""
        for level, (min_val, max_val) in self.MATURITY_LEVELS.items():
            if min_val <= score < max_val:
                return level
        
        return 'Optimized' if score >= 80 else 'Initial'
