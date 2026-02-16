"""
Utility Functions Module
Helper functions for file operations, data cleaning, etc.
"""

import os
import shutil
import time
import gc
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np
from modules.config import AppConfig


def setup_directories():
    """Create necessary directories"""
    AppConfig.TEMP_DIR.mkdir(exist_ok=True)
    AppConfig.OUTPUT_DIR.mkdir(exist_ok=True)
    if hasattr(AppConfig, 'RULES_DIR'):
        AppConfig.RULES_DIR.mkdir(exist_ok=True)


def clean_temp_directory():
    """
    Clean temporary directory with robust error handling.
    Handles file locks and permissions issues on Windows.
    """
    if not AppConfig.TEMP_DIR.exists():
        AppConfig.TEMP_DIR.mkdir(exist_ok=True)
        return

    try:
        # Force garbage collection to release file handles
        gc.collect()
        
        # Try to remove directory
        shutil.rmtree(AppConfig.TEMP_DIR)
        AppConfig.TEMP_DIR.mkdir(exist_ok=True)
        
    except PermissionError:
        # If permission denied, try to remove files individually
        try:
            for file_path in AppConfig.TEMP_DIR.iterdir():
                try:
                    if file_path.is_file():
                        # Close any open handles
                        try:
                            file_path.unlink()
                        except PermissionError:
                            # Wait a bit and try again
                            time.sleep(0.5)
                            file_path.unlink()
                    elif file_path.is_dir():
                        shutil.rmtree(file_path)
                except (PermissionError, OSError) as e:
                    # If we still can't delete, just skip and continue
                    pass
            
            # Try to recreate directory
            AppConfig.TEMP_DIR.mkdir(exist_ok=True)
            
        except Exception as e:
            # If all else fails, just ensure directory exists
            AppConfig.TEMP_DIR.mkdir(exist_ok=True)
    
    except Exception as e:
        # Fallback: just ensure directory exists
        AppConfig.TEMP_DIR.mkdir(exist_ok=True)


def clean_temp_directory_safe(max_retries: int = 3):
    """
    Enhanced temp directory cleaning with retry logic.
    
    Args:
        max_retries: Number of times to retry if cleaning fails
    """
    for attempt in range(max_retries):
        try:
            # Force garbage collection
            gc.collect()
            
            if AppConfig.TEMP_DIR.exists():
                shutil.rmtree(AppConfig.TEMP_DIR)
            
            AppConfig.TEMP_DIR.mkdir(exist_ok=True)
            return True
            
        except (PermissionError, OSError) as e:
            if attempt < max_retries - 1:
                # Wait before retrying
                time.sleep(0.5)
                continue
            else:
                # Last attempt failed, just ensure directory exists
                AppConfig.TEMP_DIR.mkdir(exist_ok=True)
                return False
    
    return False


def save_uploaded_file(uploaded_file, directory: Path) -> Path:
    """Save Streamlit uploaded file to directory"""
    file_path = directory / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path


def get_timestamp() -> str:
    """Get current timestamp for filenames"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def format_file_size(size_bytes: int) -> str:
    """Format file size in human readable format"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} TB"


def clean_value(value):
    """Clean value for Excel output - handles all edge cases"""
    # Handle None first
    if value is None:
        return ""
    
    # Handle lists and arrays
    if isinstance(value, (list, tuple, np.ndarray)):
        if len(value) == 0:
            return ""
        # Convert list to comma-separated string
        return ", ".join(str(v) for v in value)
    
    # Handle pandas NA values with try-except for safety
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        # If pd.isna fails, continue with other checks
        pass
    
    # Handle numpy NaN for float values
    if isinstance(value, float):
        try:
            if np.isnan(value):
                return ""
        except (TypeError, ValueError):
            pass
    
    # Handle empty strings
    str_value = str(value).strip()
    if str_value == "" or str_value.lower() == "nan":
        return ""
    
    # Convert to string
    return str(value)


def is_null_or_empty(value) -> bool:
    """Check if value is null or empty"""
    # Handle None
    if value is None:
        return True
    
    # Handle lists and arrays
    if isinstance(value, (list, tuple, np.ndarray)):
        return len(value) == 0
    
    # Handle pandas NA
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    
    # Handle string representation
    str_val = str(value).strip()
    return str_val == "" or str_val.lower() == "nan"