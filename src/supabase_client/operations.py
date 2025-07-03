"""
Database operations for Supabase tables.
"""
from typing import Dict, List, Optional, Any
from datetime import datetime
import logging
import os
from .client import get_supabase

logger = logging.getLogger(__name__)

class DatabaseOperations:
    """Database operations for all models."""
    
    # User Operations
    @staticmethod
    def get_users() -> List[Dict[str, Any]]:
        """Get all users."""
        try:
            supabase = get_supabase()
            result = supabase.auth.admin.list_users()
            return [
                {
                    "id": user.id,
                    "email": user.email,
                    "role": user.user_metadata.get("role", "user"),
                    "full_name": user.user_metadata.get("full_name"),
                    "company": user.user_metadata.get("company"),
                    "is_active": user.user_metadata.get("is_active", False),
                    "created_at": user.created_at
                }
                for user in result
            ]
        except Exception as e:
            logger.error(f"Error fetching users: {str(e)}")
            raise

    @staticmethod
    def update_user(user_id: str, metadata: Dict[str, Any]) -> Dict[str, Any]:
        """Update user metadata."""
        try:
            supabase = get_supabase()
            result = supabase.auth.admin.update_user_by_id(
                user_id,
                {"user_metadata": metadata}
            )
            return result.user
        except Exception as e:
            logger.error(f"Error updating user: {str(e)}")
            raise

    # Project Operations
    @staticmethod
    def create_project(
        name: str,
        user_id: str,
        data: Dict[str, Any],
        version: str,
        is_template: bool = False,
        template_name: Optional[str] = None,
        company_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a new project."""
        try:
            supabase = get_supabase()
            
            project_data = {
                "name": name,
                "user_id": user_id,
                "data": data,
                "version": version,
                "is_template": is_template,
                "template_name": template_name,
                "company_id": company_id
            }
            
            result = supabase.table("projects").insert(project_data).execute()
            return result.data[0] if result.data else None
            
        except Exception as e:
            logger.error(f"Error creating project: {str(e)}")
            raise

    @staticmethod
    def get_user_projects(user_id: str) -> List[Dict[str, Any]]:
        """Get all projects for a user."""
        try:
            supabase = get_supabase()
            result = supabase.table("projects").select("*").eq("user_id", user_id).execute()
            return result.data
        except Exception as e:
            logger.error(f"Error fetching user projects: {str(e)}")
            raise

    @staticmethod
    def get_project(project_id: str) -> Optional[Dict[str, Any]]:
        """Get a specific project."""
        try:
            supabase = get_supabase()
            result = supabase.table("projects").select("*").eq("id", project_id).single().execute()
            return result.data
        except Exception as e:
            logger.error(f"Error fetching project: {str(e)}")
            raise

    @staticmethod
    def update_project(project_id: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """Update a project."""
        try:
            supabase = get_supabase()
            result = supabase.table("projects").update(data).eq("id", project_id).execute()
            return result.data[0] if result.data else None
        except Exception as e:
            logger.error(f"Error updating project: {str(e)}")
            raise

    @staticmethod
    def delete_project(project_id: str) -> bool:
        """Delete a project."""
        try:
            supabase = get_supabase()
            result = supabase.table("projects").delete().eq("id", project_id).execute()
            return bool(result.data)
        except Exception as e:
            logger.error(f"Error deleting project: {str(e)}")
            raise

    # Template Operations
    @staticmethod
    def create_template(
        name: str,
        version: str,
        type: str,
        file: Any
    ) -> Dict[str, Any]:
        """Create a new template."""
        try:
            supabase = get_supabase()
            
            # Upload file to storage
            file_path = f"templates/{type}/{name}-{version}{os.path.splitext(file.name)[1]}"
            result = supabase.storage.from_("templates").upload(file_path, file)
            
            # Create template record
            template_data = {
                "name": name,
                "version": version,
                "type": type,
                "file_path": file_path,
                "is_active": True,
                "metadata": {}
            }
            
            result = supabase.table("templates").insert(template_data).execute()
            return result.data[0] if result.data else None
            
        except Exception as e:
            logger.error(f"Error creating template: {str(e)}")
            raise

    @staticmethod
    def get_templates() -> List[Dict[str, Any]]:
        """Get all active templates."""
        try:
            supabase = get_supabase()
            result = supabase.table("templates").select("*").eq("is_active", True).execute()
            return result.data
        except Exception as e:
            logger.error(f"Error fetching templates: {str(e)}")
            raise

    @staticmethod
    def get_template_file(template_id: str) -> bytes:
        """Get template file content."""
        try:
            supabase = get_supabase()
            
            # Get template record
            template = supabase.table("templates").select("*").eq("id", template_id).single().execute()
            if not template.data:
                raise ValueError("Template not found")
            
            # Download file from storage
            file_path = template.data["file_path"]
            result = supabase.storage.from_("templates").download(file_path)
            return result
            
        except Exception as e:
            logger.error(f"Error fetching template file: {str(e)}")
            raise

    # Company Operations
    @staticmethod
    def create_company(
        name: str,
        address: Optional[str] = None,
        contact_person: Optional[str] = None,
        phone: Optional[str] = None,
        email: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a new company."""
        try:
            supabase = get_supabase()
            
            company_data = {
                "name": name,
                "address": address,
                "contact_person": contact_person,
                "phone": phone,
                "email": email
            }
            
            result = supabase.table("companies").insert(company_data).execute()
            return result.data[0] if result.data else None
            
        except Exception as e:
            logger.error(f"Error creating company: {str(e)}")
            raise

    @staticmethod
    def get_companies() -> List[Dict[str, Any]]:
        """Get all companies."""
        try:
            supabase = get_supabase()
            result = supabase.table("companies").select("*").execute()
            return result.data
        except Exception as e:
            logger.error(f"Error fetching companies: {str(e)}")
            raise

    @staticmethod
    def get_delivery_locations() -> List[Dict[str, Any]]:
        """Get all active delivery locations."""
        try:
            supabase = get_supabase()
            result = supabase.table("delivery_locations").select("*").eq("is_active", True).execute()
            return result.data
        except Exception as e:
            logger.error(f"Error fetching delivery locations: {str(e)}")
            raise

    @staticmethod
    def get_active_templates() -> List[Dict[str, Any]]:
        """Get all active templates."""
        supabase = get_supabase()
        result = supabase.table("templates").select("*").eq("is_active", True).execute()
        return result.data if result.data else []

    @staticmethod
    def get_active_companies() -> List[Dict[str, Any]]:
        """Get all active companies."""
        supabase = get_supabase()
        result = supabase.table("companies").select("*").eq("is_active", True).execute()
        return result.data if result.data else []

    @staticmethod
    def create_delivery_location(name: str, distance: float) -> Dict[str, Any]:
        """Create a new delivery location."""
        supabase = get_supabase()
        
        location_data = {
            "name": name,
            "distance": distance,
            "is_active": True
        }
        
        result = supabase.table("delivery_locations").insert(location_data).execute()
        return result.data[0] if result.data else None

    @staticmethod
    def log_audit(
        user_id: str,
        action: str,
        entity_type: str,
        entity_id: str,
        details: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Create an audit log entry."""
        supabase = get_supabase()
        
        log_data = {
            "user_id": user_id,
            "action": action,
            "entity_type": entity_type,
            "entity_id": entity_id,
            "details": details
        }
        
        result = supabase.table("audit_logs").insert(log_data).execute()
        return result.data[0] if result.data else None 