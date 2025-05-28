"""
Date utility functions for consistent date formatting across the application.
"""
from datetime import datetime
from typing import Optional


def format_date_for_display(date_str: str) -> str:
    """
    Convert date string to DD/MM/YYYY format for display.
    
    Args:
        date_str: Date string in various formats (YYYY-MM-DD, DD/MM/YYYY, etc.)
        
    Returns:
        str: Date in DD/MM/YYYY format, or original string if conversion fails
    """
    if not date_str:
        return ""
    
    # If already in DD/MM/YYYY format, return as-is
    if "/" in date_str and len(date_str.split("/")) == 3:
        parts = date_str.split("/")
        if len(parts[0]) <= 2 and len(parts[1]) <= 2 and len(parts[2]) == 4:
            return date_str
    
    # Try to parse from YYYY-MM-DD format
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%d/%m/%Y")
    except ValueError:
        pass
    
    # Try to parse from other common formats
    formats_to_try = [
        "%d/%m/%Y",
        "%d-%m-%Y", 
        "%Y/%m/%d",
        "%m/%d/%Y"
    ]
    
    for fmt in formats_to_try:
        try:
            date_obj = datetime.strptime(date_str, fmt)
            return date_obj.strftime("%d/%m/%Y")
        except ValueError:
            continue
    
    # If all parsing fails, return original string
    return date_str


def format_date_for_storage(date_str: str) -> str:
    """
    Convert date string to DD/MM/YYYY format for consistent storage.
    
    Args:
        date_str: Date string in various formats
        
    Returns:
        str: Date in DD/MM/YYYY format for storage
    """
    return format_date_for_display(date_str)


def get_current_date() -> str:
    """
    Get current date in DD/MM/YYYY format.
    
    Returns:
        str: Current date in DD/MM/YYYY format
    """
    return datetime.now().strftime("%d/%m/%Y")


def convert_date_object_to_display(date_obj) -> str:
    """
    Convert a date object to DD/MM/YYYY format string.
    
    Args:
        date_obj: datetime.date object
        
    Returns:
        str: Date in DD/MM/YYYY format
    """
    if date_obj is None:
        return ""
    
    return date_obj.strftime("%d/%m/%Y") 