"""
Supabase integration package.
"""
from .auth import AuthenticationManager
from .client import get_supabase
from .models import User, Project, Template, Company, DeliveryLocation, AuditLog
from .operations import DatabaseOperations

__all__ = [
    'AuthenticationManager',
    'get_supabase',
    'User',
    'Project',
    'Template',
    'Company',
    'DeliveryLocation',
    'AuditLog',
    'DatabaseOperations'
] 