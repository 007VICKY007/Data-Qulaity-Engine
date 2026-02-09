"""
File Loader Service Module
Universal file loading with support for multiple formats
"""

import pandas as pd
import json
from pathlib import Path
from typing import Optional, List


class FileLoaderService:
    """Service class for loading various file formats"""
    
    def __init__(self):
        self.supported_formats = [".csv", ".tsv", ".xlsx", ".xls", ".xlsm", ".json"]
    
    def load_dataframe(
        self, 
        file_path: Path, 
        sheet_name: Optional[str] = None,
        columns: Optional[List[str]] = None
    ) -> pd.DataFrame:
        """
        Load dataframe from various file formats
        
        Args:
            file_path: Path to the file
            sheet_name: Optional sheet name for Excel files
            columns: Optional list of columns to select
            
        Returns:
            pandas DataFrame
        """
        file_path = Path(file_path)
        ext = file_path.suffix.lower()
        
        if ext not in self.supported_formats:
            raise ValueError(f"Unsupported file format: {ext}")
        
        try:
            # Load based on file type
            if ext in [".csv", ".tsv"]:
                df = self._load_csv(file_path, ext)
            
            elif ext in [".xlsx", ".xls", ".xlsm"]:
                df = self._load_excel(file_path, sheet_name)
            
            elif ext == ".json":
                df = self._load_json(file_path)
            
            else:
                raise ValueError(f"Unsupported file type: {ext}")
            
            # Normalize dataframe
            df = self._normalize_dataframe(df)
            
            # Select columns if specified
            if columns:
                missing_cols = [c for c in columns if c not in df.columns]
                if missing_cols:
                    raise ValueError(f"Columns not found in data: {missing_cols}")
                df = df[columns]
            
            return df
            
        except Exception as e:
            raise Exception(f"Error loading file {file_path}: {str(e)}")
    
    def _load_csv(self, file_path: Path, ext: str) -> pd.DataFrame:
        """Load CSV or TSV file"""
        separator = "\t" if ext == ".tsv" else ","
        return pd.read_csv(file_path, sep=separator, dtype=str, low_memory=False)
    
    def _load_excel(self, file_path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """Load Excel file"""
        xls = pd.ExcelFile(file_path)
        sheet = sheet_name if sheet_name and sheet_name in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(xls, sheet_name=sheet, dtype=str)
    
    def _load_json(self, file_path: Path) -> pd.DataFrame:
        """Load JSON file"""
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        if isinstance(data, list):
            return pd.DataFrame(data)
        elif isinstance(data, dict):
            if "data" in data:
                return pd.DataFrame(data["data"])
            else:
                return pd.DataFrame([data])
        else:
            raise ValueError("JSON must be a list or dict")
    
    def _normalize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Normalize dataframe - clean column names and handle nulls"""
        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Replace NaN with None for consistency
        df = df.where(pd.notnull(df), None)
        
        return df
    
    def get_sheet_names(self, file_path: Path) -> List[str]:
        """Get sheet names from Excel file"""
        if Path(file_path).suffix.lower() not in [".xlsx", ".xls", ".xlsm"]:
            return []
        
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    
    def validate_file(self, file_path: Path) -> bool:
        """Validate if file can be loaded"""
        try:
            ext = Path(file_path).suffix.lower()
            if ext not in self.supported_formats:
                return False
            
            # Try to load first row
            if ext in [".csv", ".tsv"]:
                pd.read_csv(file_path, nrows=1)
            elif ext in [".xlsx", ".xls", ".xlsm"]:
                pd.read_excel(file_path, nrows=1)
            elif ext == ".json":
                with open(file_path, "r") as f:
                    json.load(f)
            
            return True
            
        except Exception:
            return False


class LegacyFileLoader:
    """
    Legacy loader for backward compatibility with existing preprocess.py
    """
    
    @staticmethod
    def load_data(
        path: str,
        sheet: Optional[str] = None,
        columns: Optional[List[str]] = None
    ) -> pd.DataFrame:
        """
        Load CSV or Excel data (backward compatible)
        """
        loader = FileLoaderService()
        return loader.load_dataframe(Path(path), sheet, columns)
    
    @staticmethod
    def load_csv_data(path: str, columns: Optional[List[str]] = None) -> pd.DataFrame:
        """Load CSV file"""
        return LegacyFileLoader.load_data(path=path, columns=columns)
    
    @staticmethod
    def load_excel_data(
        path: str,
        sheet: str,
        columns: Optional[List[str]] = None
    ) -> pd.DataFrame:
        """Load Excel file"""
        return LegacyFileLoader.load_data(path=path, sheet=sheet, columns=columns)
