"""
Scoring Engine Module
Calculate DQ scores at multiple levels
"""

import pandas as pd
from collections import defaultdict
from typing import Dict, List


class ScoringService:
    """Service for calculating DQ scores"""
    
    @staticmethod
    def calculate_overall_score(results_df: pd.DataFrame) -> float:
        """Calculate overall DQ score as percentage"""
        total = len(results_df)
        if total == 0:
            return 0.0
        
        clean = len(results_df[results_df["Count of issues"] == 0])
        return round((clean / total) * 100, 2)
    
    @staticmethod
    def calculate_column_scores(
        results_df: pd.DataFrame, 
        all_columns: List[str]
    ) -> Dict[str, float]:
        """Calculate DQ score per column"""
        total = len(results_df)
        if total == 0:
            return {}
        
        column_scores = {}
        
        # Track failed columns
        failed_tracker = defaultdict(set)
        for idx, row in results_df.iterrows():
            for col in row.get("_failed_columns_list", []):
                failed_tracker[col].add(idx)
        
        # Calculate score for each column
        for col in all_columns:
            # Skip internal columns
            if col.startswith("_") or col in [
                "Issues", "Count of issues", "Failed_Rules", 
                "Failed_Columns", "Issue categories"
            ]:
                continue
            
            failed_count = len(failed_tracker.get(col, set()))
            clean_count = total - failed_count
            score = (clean_count / total) * 100
            column_scores[col] = round(score, 2)
        
        return column_scores
    
    @staticmethod
    def calculate_dimension_scores(results_df: pd.DataFrame) -> Dict[str, float]:
        """Calculate scores by DQ dimension"""
        total = len(results_df)
        if total == 0:
            return {}
        
        dimensions = set()
        for dims in results_df["Issue categories"].dropna():
            dimensions.update(d.strip() for d in str(dims).split(",") if d.strip())
        
        dimension_scores = {}
        for dimension in dimensions:
            failed = len(results_df[
                results_df["Issue categories"].str.contains(dimension, na=False)
            ])
            clean = total - failed
            score = (clean / total) * 100
            dimension_scores[dimension] = round(score, 2)
        
        return dimension_scores
