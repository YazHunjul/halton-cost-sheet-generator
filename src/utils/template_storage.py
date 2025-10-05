"""
Template storage utilities for managing Word document templates in Supabase Storage.
"""
import os
import io
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple
from config.supabase_config import get_supabase_client

# Storage bucket name
BUCKET_NAME = "templates"

# Template definitions
TEMPLATE_FILES = {
    "canopy_quotation": "Halton Quote Feb 2024.docx",
    "recoair_quotation": "Halton RECO Quotation Jan 2025 (2).docx",
    "ahu_quotation": "Halton AHU quote JAN2020.docx"
}


def ensure_bucket_exists():
    """Ensure the templates bucket exists in Supabase Storage."""
    try:
        client = get_supabase_client(use_service_role=True)

        # Try to get bucket info
        try:
            client.storage.get_bucket(BUCKET_NAME)
            return True, "Bucket exists"
        except:
            # Bucket doesn't exist, create it
            try:
                client.storage.create_bucket(
                    BUCKET_NAME,
                    options={"public": False}  # Private bucket
                )
                return True, "Bucket created"
            except Exception as e:
                return False, f"Failed to create bucket: {str(e)}"
    except Exception as e:
        return False, f"Error checking bucket: {str(e)}"


def upload_template_to_storage(template_key: str, file_bytes: bytes, filename: str) -> Tuple[bool, str]:
    """
    Upload a template file to Supabase Storage.

    Args:
        template_key: Key identifying the template type
        file_bytes: File content as bytes
        filename: Original filename

    Returns:
        Tuple of (success, message)
    """
    try:
        # Ensure bucket exists
        success, message = ensure_bucket_exists()
        if not success:
            return False, message

        client = get_supabase_client(use_service_role=True)

        # Create backup of existing template first
        backup_success, backup_message = backup_template_in_storage(template_key)

        # Upload new template (overwrites if exists)
        storage_path = f"{template_key}/{filename}"

        try:
            # Remove existing file if it exists
            try:
                client.storage.from_(BUCKET_NAME).remove([storage_path])
            except:
                pass  # File might not exist, that's okay

            # Upload new file
            client.storage.from_(BUCKET_NAME).upload(
                storage_path,
                file_bytes,
                file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
            )

            return True, f"Template uploaded successfully{' (backup created)' if backup_success else ''}"
        except Exception as e:
            return False, f"Failed to upload template: {str(e)}"

    except Exception as e:
        return False, f"Error uploading template: {str(e)}"


def download_template_from_storage(template_key: str) -> Tuple[bool, Optional[bytes], str]:
    """
    Download a template file from Supabase Storage.

    Args:
        template_key: Key identifying the template type

    Returns:
        Tuple of (success, file_bytes, message)
    """
    try:
        filename = TEMPLATE_FILES.get(template_key)
        if not filename:
            return False, None, f"Unknown template key: {template_key}"

        client = get_supabase_client(use_service_role=True)
        storage_path = f"{template_key}/{filename}"

        try:
            # Download file
            file_bytes = client.storage.from_(BUCKET_NAME).download(storage_path)
            return True, file_bytes, "Template downloaded successfully"
        except Exception as e:
            return False, None, f"Template not found in storage: {str(e)}"

    except Exception as e:
        return False, None, f"Error downloading template: {str(e)}"


def backup_template_in_storage(template_key: str) -> Tuple[bool, str]:
    """
    Create a backup of a template in Supabase Storage.

    Args:
        template_key: Key identifying the template type

    Returns:
        Tuple of (success, message)
    """
    try:
        filename = TEMPLATE_FILES.get(template_key)
        if not filename:
            return False, f"Unknown template key: {template_key}"

        # Download current template
        success, file_bytes, message = download_template_from_storage(template_key)
        if not success:
            return False, f"No template to backup: {message}"

        # Create backup filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_stem = Path(filename).stem
        file_ext = Path(filename).suffix
        backup_filename = f"{file_stem}_backup_{timestamp}{file_ext}"

        client = get_supabase_client(use_service_role=True)
        backup_path = f"backups/{template_key}/{backup_filename}"

        # Upload backup
        client.storage.from_(BUCKET_NAME).upload(
            backup_path,
            file_bytes,
            file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
        )

        return True, f"Backup created: {backup_filename}"

    except Exception as e:
        return False, f"Error creating backup: {str(e)}"


def list_template_backups(template_key: str) -> Tuple[bool, list, str]:
    """
    List all backups for a specific template.

    Args:
        template_key: Key identifying the template type

    Returns:
        Tuple of (success, list of backup files, message)
    """
    try:
        client = get_supabase_client(use_service_role=True)
        backup_path = f"backups/{template_key}"

        try:
            files = client.storage.from_(BUCKET_NAME).list(backup_path)
            # Sort by created_at descending (newest first)
            files.sort(key=lambda x: x.get('created_at', ''), reverse=True)
            return True, files, "Backups retrieved successfully"
        except Exception as e:
            return False, [], f"No backups found: {str(e)}"

    except Exception as e:
        return False, [], f"Error listing backups: {str(e)}"


def get_template_metadata(template_key: str) -> Tuple[bool, Optional[dict], str]:
    """
    Get metadata about a template file.

    Args:
        template_key: Key identifying the template type

    Returns:
        Tuple of (success, metadata dict, message)
    """
    try:
        filename = TEMPLATE_FILES.get(template_key)
        if not filename:
            return False, None, f"Unknown template key: {template_key}"

        client = get_supabase_client(use_service_role=True)
        storage_path = f"{template_key}/{filename}"

        try:
            # List files to get metadata
            files = client.storage.from_(BUCKET_NAME).list(template_key)

            # Find our file
            for file in files:
                if file['name'] == filename:
                    return True, {
                        'name': file['name'],
                        'size': file.get('metadata', {}).get('size', 0),
                        'updated_at': file.get('updated_at'),
                        'created_at': file.get('created_at')
                    }, "Metadata retrieved"

            return False, None, "Template not found in storage"

        except Exception as e:
            return False, None, f"Template not found: {str(e)}"

    except Exception as e:
        return False, None, f"Error getting metadata: {str(e)}"


def sync_local_templates_to_storage():
    """
    Sync local template files to Supabase Storage (for initial setup).
    This should be run once to upload existing templates.
    """
    results = []
    local_template_dir = Path("templates/word")

    if not local_template_dir.exists():
        return False, "Local template directory not found"

    for template_key, filename in TEMPLATE_FILES.items():
        local_path = local_template_dir / filename

        if local_path.exists():
            with open(local_path, 'rb') as f:
                file_bytes = f.read()

            success, message = upload_template_to_storage(template_key, file_bytes, filename)
            results.append(f"{template_key}: {message}")
        else:
            results.append(f"{template_key}: Local file not found at {local_path}")

    return True, "\n".join(results)


def download_template_to_local(template_key: str, local_dir: str = "templates/word") -> Tuple[bool, str]:
    """
    Download a template from storage to local filesystem.
    Used when generating Word documents.

    Args:
        template_key: Key identifying the template type
        local_dir: Local directory to save template

    Returns:
        Tuple of (success, local_file_path)
    """
    try:
        filename = TEMPLATE_FILES.get(template_key)
        if not filename:
            return False, f"Unknown template key: {template_key}"

        # Download from storage
        success, file_bytes, message = download_template_from_storage(template_key)
        if not success:
            return False, message

        # Ensure local directory exists
        local_path = Path(local_dir)
        local_path.mkdir(parents=True, exist_ok=True)

        # Save to local file
        file_path = local_path / filename
        with open(file_path, 'wb') as f:
            f.write(file_bytes)

        return True, str(file_path)

    except Exception as e:
        return False, f"Error downloading to local: {str(e)}"
