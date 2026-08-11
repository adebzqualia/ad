"""POPS Check — contrôle structurel de classeurs Excel."""

from .compare import analyze_directories, compare_workbooks
from .config import AppConfig, load_config

__all__ = ["AppConfig", "analyze_directories", "compare_workbooks", "load_config"]
__version__ = "1.1.0"
