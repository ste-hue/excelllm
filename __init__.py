"""
ExcelLLM - Convert Excel files to LLM-friendly formats
"""

from .excelllm import ExcelParser, ExcelFormatter, main

__version__ = "0.1.0"
__all__ = ["ExcelParser", "ExcelFormatter", "main"]
