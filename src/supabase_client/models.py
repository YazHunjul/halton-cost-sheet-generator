"""
Data models for Supabase tables.
"""
from typing import Optional, Dict, Any
from datetime import datetime
from dataclasses import dataclass

@dataclass
class BaseModel:
    """Base model with common fields."""
    id: str
    created_at: datetime
    updated_at: Optional[datetime]

@dataclass
class User(BaseModel):
    """User model matching Supabase auth.users"""
    email: str
    role: str  # admin, user
    is_active: bool
    full_name: Optional[str]
    company: Optional[str]

@dataclass
class Project(BaseModel):
    """Project model for cost sheet projects"""
    name: str
    user_id: str
    data: Dict[str, Any]  # JSON data of the project
    version: str
    is_template: bool = False
    template_name: Optional[str] = None
    company_id: Optional[str] = None
    is_active: bool = True

@dataclass
class Template(BaseModel):
    """Template model for Excel/Word templates"""
    name: str
    file_path: str
    version: str
    type: str  # excel, word
    is_active: bool = True
    uploaded_by: str  # user_id
    metadata: Dict[str, Any] = None

@dataclass
class Company(BaseModel):
    """Company model"""
    name: str
    address: Optional[str]
    contact_person: Optional[str]
    phone: Optional[str]
    email: Optional[str]
    is_active: bool

@dataclass
class DeliveryLocation(BaseModel):
    """Delivery location model"""
    name: str
    distance: float
    is_active: bool = True

@dataclass
class AuditLog(BaseModel):
    """Audit log for tracking important operations"""
    user_id: str
    action: str
    entity_type: str  # project, template, user, etc.
    entity_id: str
    details: Dict[str, Any]  # JSON details of the action 