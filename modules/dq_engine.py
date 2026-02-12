"""
Data Quality Engine
Executes quality rules and generates scores
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Any
from .file_loader import FileLoaderService
from .rulebook_builder import RulebookBuilderService


class DataQualityEngine:
    """Data Quality Assessment Engine"""
    
    def __init__(self, df: pd.DataFrame, rules_path: Path):
        self.df = df
        self.rules_path = Path(rules_path)
        self.rulebook = self._load_rulebook()
        self.results_df = None
    
    def _load_rulebook(self) -> Dict:
        """Load rulebook from file"""
        builder = RulebookBuilderService()
        
        if self.rules_path.suffix.lower() == '.json':
            return builder.load_json_rulebook(self.rules_path)
        else:
            loader = FileLoaderService()
            rules_df = loader.load_dataframe(self.rules_path)
            return builder.build_from_rules_dataset(rules_df, self.df.columns.tolist())
    
    def run(self) -> Dict[str, Any]:
        """Run DQ assessment and return results"""
        
        results = []
        total_issues = 0
        
        for idx, row in self.df.iterrows():
            row_issues = []
            row_failed_rules = set()
            row_dimensions = set()
            
            for rule in self.rulebook.get('rules', []):
                is_passed = self._execute_rule(row, rule)
                
                if not is_passed:
                    row_issues.append(rule.get('message', 'Failed'))
                    row_failed_rules.add(rule.get('rule_type', 'unknown'))
                    row_dimensions.add(rule.get('dimension', 'General'))
            
            total_issues += len(row_issues)
            
            result = row.to_dict()
            result['_issues'] = len(row_issues)
            result['_issue_list'] = ' | '.join(row_issues) if row_issues else 'OK'
            result['_failed_rules'] = ', '.join(row_failed_rules) if row_failed_rules else ''
            result['_dimensions'] = ', '.join(sorted(row_dimensions)) if row_dimensions else ''
            
            results.append(result)
        
        self.results_df = pd.DataFrame(results)
        
        overall_score = self._calculate_overall_score()
        dimension_scores = self._calculate_dimension_scores()
        column_scores = self._calculate_column_scores()
        
        clean_records = len(self.results_df[self.results_df['_issues'] == 0])
        
        return {
            'overall_score': overall_score,
            'total_records': len(self.df),
            'clean_records': clean_records,
            'total_issues': total_issues,
            'dimension_scores': dimension_scores,
            'column_scores': column_scores,
            'results_df': self.results_df
        }
    
    def _execute_rule(self, row: pd.Series, rule: Dict) -> bool:
        """Execute single rule against row"""
        rule_type = rule.get('rule_type')
        column = rule.get('column')
        expression = rule.get('expression')
        
        if not column or column not in row.index:
            return True
        
        value = row[column]
        
        if self._is_null_or_empty(value):
            if rule_type == 'not_null':
                return False
            return True
        
        if rule_type == 'not_null':
            return True
        
        elif rule_type == 'uniqueness':
            count = len(self.df[self.df[column] == value])
            return count == 1
        
        elif rule_type == 'regex' and expression:
            import re
            try:
                return bool(re.match(str(expression), str(value)))
            except:
                return False
        
        elif rule_type == 'allowed_values' and expression:
            allowed = [v.strip() for v in str(expression).split(',')]
            return str(value) in allowed
        
        elif rule_type == 'email_format':
            import re
            email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            return bool(re.match(email_pattern, str(value)))
        
        elif rule_type == 'numeric_only':
            try:
                float(value)
                return True
            except:
                return False
        
        elif rule_type == 'length' and expression:
            try:
                if ',' in str(expression):
                    min_len, max_len = map(int, str(expression).split(','))
                    return min_len <= len(str(value)) <= max_len
                else:
                    return len(str(value)) == int(expression)
            except:
                return False
        
        elif rule_type == 'contains' and expression:
            return str(expression) in str(value)
        
        elif rule_type == 'no_special_chars':
            import re
            pattern = expression if expression else r'[^A-Za-z0-9\s]'
            return not bool(re.search(str(pattern), str(value)))
        
        return True
    
    def _is_null_or_empty(self, value) -> bool:
        """Check if value is null/empty"""
        return (
            value is None 
            or pd.isna(value) 
            or str(value).strip() == "" 
            or str(value).lower() == "nan"
        )
    
    def _calculate_overall_score(self) -> float:
        """Calculate overall DQ score"""
        total = len(self.results_df)
        if total == 0:
            return 0.0
        
        clean = len(self.results_df[self.results_df['_issues'] == 0])
        return round((clean / total) * 100, 2)
    
    def _calculate_dimension_scores(self) -> Dict[str, float]:
        """Calculate scores by dimension"""
        scores = {}
        total = len(self.results_df)
        
        if total == 0:
            return scores
        
        dimensions = set()
        for dims in self.results_df['_dimensions'].dropna():
            if dims:
                dimensions.update(d.strip() for d in str(dims).split(',') if d.strip())
        
        for dim in dimensions:
            failed = len(self.results_df[
                self.results_df['_dimensions'].str.contains(dim, na=False)
            ])
            clean = total - failed
            score = (clean / total) * 100
            scores[dim] = round(score, 2)
        
        return scores
    
    def _calculate_column_scores(self) -> Dict[str, float]:
        """Calculate scores by column"""
        scores = {}
        total = len(self.results_df)
        
        if total == 0:
            return scores
        
        for col in self.df.columns:
            if col.startswith('_'):
                continue
            
            clean = total
            score = (clean / total) * 100
            scores[col] = round(score, 2)
        
        return scores
