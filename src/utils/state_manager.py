"""
State management utilities for saving and loading form data via URL parameters.
"""
import streamlit as st
import json
import base64
from urllib.parse import quote, unquote
from typing import Dict, Any
import zlib
from datetime import datetime

def compress_state(state_dict: Dict[str, Any]) -> str:
    """
    Compress and encode state dictionary for URL storage.
    
    Args:
        state_dict: Dictionary containing form state
        
    Returns:
        Compressed and encoded string safe for URL
    """
    # Convert to JSON
    json_str = json.dumps(state_dict, separators=(',', ':'))
    
    # Compress using zlib
    compressed = zlib.compress(json_str.encode('utf-8'))
    
    # Encode to base64 and make URL-safe
    encoded = base64.urlsafe_b64encode(compressed).decode('utf-8')
    
    return encoded

def decompress_state(encoded_str: str) -> Dict[str, Any]:
    """
    Decompress and decode state from URL parameter.
    
    Args:
        encoded_str: Compressed and encoded state string
        
    Returns:
        Dictionary containing form state
    """
    try:
        # Decode from base64
        compressed = base64.urlsafe_b64decode(encoded_str.encode('utf-8'))
        
        # Decompress
        json_str = zlib.decompress(compressed).decode('utf-8')
        
        # Parse JSON
        state_dict = json.loads(json_str)
        
        return state_dict
    except Exception as e:
        st.error(f"Failed to load saved state: {str(e)}")
        return {}

def extract_form_state() -> Dict[str, Any]:
    """
    Extract relevant form data from session state.
    
    Returns:
        Dictionary containing form state
    """
    state_dict = {}
    
    # Only save the essential data structures that contain the actual form data
    # This avoids all widget state issues
    essential_keys = {
        'project_info': st.session_state.get('project_info', {}),
        'levels': st.session_state.get('levels', []),
        'template_path': st.session_state.get('template_path', 'templates/excel/Cost Sheet R19.2 Jun 2025.xlsx'),
        'sp_loaded_file': st.session_state.get('sp_loaded_file', None),  # Track loaded file
    }
    
    # Add essential keys to state dict
    for key, value in essential_keys.items():
        if value is not None:
            state_dict[key] = value
    
    # Also save any custom form state values that are simple data types
    # These are typically the state variables we use for forms
    form_state_patterns = ['_state']  # Keys ending with _state are typically our custom state holders
    
    for key in st.session_state:
        if any(pattern in key for pattern in form_state_patterns):
            value = st.session_state.get(key)
            # Only save simple data types
            if isinstance(value, (str, int, float, bool, list, dict)):
                state_dict[key] = value
    
    return state_dict

def restore_form_state(state_dict: Dict[str, Any]):
    """
    Restore form state from dictionary to session state.
    
    Args:
        state_dict: Dictionary containing form state
    """
    # Safe keys that we know can be restored
    safe_keys = ['project_info', 'levels', 'template_path', 'sp_loaded_file']
    
    for key, value in state_dict.items():
        try:
            # Only restore known safe keys or state variables
            if key in safe_keys or key.endswith('_state'):
                st.session_state[key] = value
        except Exception as e:
            # Skip keys that can't be set
            print(f"Skipping key {key}: {str(e)}")
            continue

def generate_save_link() -> str:
    """
    Generate a shareable link with current form state.
    
    Returns:
        Full URL with encoded state
    """
    # Get current form state
    state_dict = extract_form_state()
    
    if not state_dict:
        return None
    
    # Compress and encode
    encoded_state = compress_state(state_dict)
    
    # Get current URL base
    try:
        params = st.query_params
    except AttributeError:
        # Fallback for older Streamlit versions
        params = st.experimental_get_query_params()
    
    # Convert to dict if needed
    if hasattr(params, '_to_dict'):
        params = dict(params)
    elif not isinstance(params, dict):
        params = {}
    
    # Add state parameter
    params['state'] = encoded_state
    
    # Construct URL
    base_url = 'https://haltonsales.streamlit.app'  # Production URL
    
    # Build query string
    query_parts = []
    for key, value in params.items():
        if isinstance(value, list):
            for v in value:
                query_parts.append(f"{key}={quote(str(v))}")
        else:
            query_parts.append(f"{key}={quote(str(value))}")
    
    query_string = '&'.join(query_parts)
    
    return f"{base_url}?{query_string}"

def load_from_url():
    """
    Load form state from URL parameters if present.
    """
    try:
        params = st.query_params
    except AttributeError:
        # Fallback for older Streamlit versions
        params = st.experimental_get_query_params()
    
    # Convert to dict if needed
    if hasattr(params, '_to_dict'):
        params_dict = dict(params)
    else:
        params_dict = params
    
    if 'state' in params_dict:
        encoded_state = params_dict['state'][0] if isinstance(params_dict['state'], list) else params_dict['state']
        
        # Decompress and restore state
        state_dict = decompress_state(encoded_state)
        
        if state_dict:
            restore_form_state(state_dict)
            
            # Clear the state parameter from URL to avoid reloading
            try:
                st.query_params.clear()
            except AttributeError:
                # Fallback for older Streamlit versions
                st.experimental_set_query_params()
            
            return True
    
    return False

def add_save_progress_button():
    """
    Add save progress functionality to the form.
    """
    # Create a container for the save progress feature
    with st.container():
        st.markdown("---")
        st.markdown("### 💾 Save Your Progress")
        st.markdown("Save your current form data and get a shareable link to restore it later.")
        
        if st.button("Generate Save Link", type="primary", help="Generate a link to save your current progress"):
            link = generate_save_link()
            
            if link:
                # Store the link in session state so it persists
                st.session_state['saved_link'] = link
                
                # Create an expander to show the link
                with st.expander("✅ Progress Saved Successfully!", expanded=True):
                    st.markdown("**Your save link:**")
                    
                    # Use a text input for easy copying
                    st.text_input(
                        "Copy this link:", 
                        value=link, 
                        key="save_link_display",
                        help="Click in the box and use Ctrl+A (Cmd+A on Mac) to select all, then Ctrl+C (Cmd+C) to copy"
                    )
                    
                    st.markdown("**Quick Actions:**")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        # Get project information for descriptive content
                        project_info = st.session_state.get('project_info', {})
                        project_name = project_info.get('project_name', 'Unnamed Project')
                        project_number = project_info.get('project_number', 'N/A')
                        customer_name = project_info.get('customer_name', 'N/A')
                        
                        # Create a download button for the link as a text file
                        link_content = f"""Halton Cost Sheet Generator - Saved Progress

Project Information:
-------------------
Project Name: {project_name}
Project Number: {project_number}
Customer: {customer_name}

Save Details:
-------------
Date Saved: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Shareable Link:
---------------
{link}

Instructions:
-------------
1. Copy the link above or use this file to save your progress
2. Open the link in your browser to restore your progress
3. All your form data will be automatically loaded
4. Share this link with colleagues to collaborate on the same project

Note: This link contains all your form data. Anyone with this link can access and modify the project details."""
                        
                        st.download_button(
                            label="📥 Download as Text File",
                            data=link_content,
                            file_name=f"halton_{project_number}_{project_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                            mime="text/plain",
                            help="Download the link as a text file for safekeeping"
                        )
                    
                    with col2:
                        # Open in new tab button
                        st.markdown(
                            f"""
                            <a href="{link}" target="_blank" style="text-decoration: none;">
                                <button style="
                                    background-color: #4CAF50;
                                    color: white;
                                    padding: 10px 20px;
                                    border: none;
                                    border-radius: 4px;
                                    cursor: pointer;
                                    font-size: 14px;
                                    width: 100%;
                                ">
                                    🔗 Test Link
                                </button>
                            </a>
                            """,
                            unsafe_allow_html=True
                        )
                    
                    with col3:
                        # Email link button (creates mailto link)
                        email_subject = f"Halton Cost Sheet - {project_name} ({project_number})"
                        email_body = f"""Here is my saved progress link for the Halton Cost Sheet Generator:

Project Details:
- Project Name: {project_name}
- Project Number: {project_number}
- Customer: {customer_name}
- Date Saved: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Shareable Link:
{link}

To restore this project, simply open the link above in your browser. All form data will be automatically loaded.

This link can be shared with colleagues who need to view or continue working on this cost sheet."""
                        
                        mailto_link = f"mailto:?subject={quote(email_subject)}&body={quote(email_body)}"
                        
                        st.markdown(
                            f"""
                            <a href="{mailto_link}" style="text-decoration: none;">
                                <button style="
                                    background-color: #0066CC;
                                    color: white;
                                    padding: 10px 20px;
                                    border: none;
                                    border-radius: 4px;
                                    cursor: pointer;
                                    font-size: 14px;
                                    width: 100%;
                                ">
                                    📧 Email Link
                                </button>
                            </a>
                            """,
                            unsafe_allow_html=True
                        )
                    
                    st.info("💡 **Tips for copying:**\n- Click inside the text box above\n- Press **Ctrl+A** (or **Cmd+A** on Mac) to select all\n- Press **Ctrl+C** (or **Cmd+C** on Mac) to copy\n- Or use the download button to save as a text file")
            else:
                st.warning("No form data to save yet. Please fill in some fields first.")
        
        # Show the last saved link if it exists
        elif 'saved_link' in st.session_state:
            with st.expander("📌 Last Saved Link", expanded=False):
                st.text_input(
                    "Your last saved link:", 
                    value=st.session_state['saved_link'], 
                    key="last_save_link_display",
                    help="Click in the box and use Ctrl+A (Cmd+A on Mac) to select all, then Ctrl+C (Cmd+C) to copy"
                )
        
        st.markdown("---")