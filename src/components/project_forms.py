"""
Project-specific form components for the HVAC Project Management Tool.
"""
import streamlit as st
from typing import Dict, Any, List
from config.constants import SessionKeys, PROJECT_TYPES, VALID_CANOPY_MODELS

def get_state_key(level_idx: int, area_idx: int = None, canopy_idx: int = None, field: str = None) -> str:
    """Generate a unique key for session state."""
    parts = [f"level_{level_idx}"]
    if area_idx is not None:
        parts.append(f"area_{area_idx}")
    if canopy_idx is not None:
        parts.append(f"canopy_{canopy_idx}")
    if field:
        parts.append(field)
    return "_".join(parts)

def init_state_if_needed(key: str, default_value: Any = None):
    """Initialize session state if not already present."""
    if key not in st.session_state:
        st.session_state[key] = default_value

def canopy_form(level_idx: int, area_idx: int, canopy_idx: int) -> Dict[str, Any]:
    """Renders form fields for a single canopy's details."""
    st.markdown("##### Canopy Details")
    
    # Initialize state for this canopy
    canopy_key = get_state_key(level_idx, area_idx, canopy_idx)
    init_state_if_needed(canopy_key, {})
    
    col1, col2 = st.columns(2)
    
    with col1:
        ref_key = get_state_key(level_idx, area_idx, canopy_idx, "ref")
        init_state_if_needed(ref_key, "")
        ref_number = st.text_input(
            "Reference Number *",
            key=ref_key,
            help="Enter the canopy reference number"
        )
        
        model_key = get_state_key(level_idx, area_idx, canopy_idx, "model")
        init_state_if_needed(model_key, VALID_CANOPY_MODELS[0])
        model = st.selectbox(
            "Model *",
            options=VALID_CANOPY_MODELS,
            key=model_key,
            help="Select the canopy model"
        )
        
        config_key = get_state_key(level_idx, area_idx, canopy_idx, "config")
        init_state_if_needed(config_key, "Wall")
        configuration = st.selectbox(
            "Configuration *",
            options=["Wall", "Island", "Single", "Double"],
            key=config_key,
            help="Select the canopy configuration"
        )
    
    with col2:
        st.markdown("**Wall Cladding**")
        clad_enable_key = get_state_key(level_idx, area_idx, canopy_idx, "clad_enable")
        init_state_if_needed(clad_enable_key, False)
        cladding_enabled = st.checkbox(
            "Wall Cladding",
            key=clad_enable_key,
            help="Check if wall cladding is required"
        )
        
        if cladding_enabled:
            col3, col4 = st.columns(2)
            with col3:
                width_key = get_state_key(level_idx, area_idx, canopy_idx, "width")
                init_state_if_needed(width_key, 0)
                width = st.number_input(
                    "Width (mm)",
                    min_value=0,
                    key=width_key,
                    help="Enter the cladding width in millimeters"
                )
            
            with col4:
                height_key = get_state_key(level_idx, area_idx, canopy_idx, "height")
                init_state_if_needed(height_key, 0)
                height = st.number_input(
                    "Height (mm)",
                    min_value=0,
                    key=height_key,
                    help="Enter the cladding height in millimeters"
                )
            
            # Description (multi-select for position)
            desc_key = get_state_key(level_idx, area_idx, canopy_idx, "description")
            init_state_if_needed(desc_key, [])
            description = st.multiselect(
                "Description",
                options=["rear", "left hand", "right hand"],
                default=st.session_state.get(desc_key, []),
                key=desc_key,
                help="Select cladding positions (can select multiple)"
            )
    
    st.markdown("**Additional Options**")
    col5, col6 = st.columns(2)
    
    with col5:
        fire_sup_key = get_state_key(level_idx, area_idx, canopy_idx, "fire_sup")
        init_state_if_needed(fire_sup_key, False)
        fire_suppression = st.toggle(
            "Fire Suppression System",
            key=fire_sup_key,
            help="Toggle if fire suppression system is needed"
        )
        
        uvc_key = get_state_key(level_idx, area_idx, canopy_idx, "uvc")
        init_state_if_needed(uvc_key, False)
        uvc = st.toggle(
            "UV-C System",
            key=uvc_key,
            help="Toggle if UV-C system is needed"
        )
    
    with col6:
        sdu_key = get_state_key(level_idx, area_idx, canopy_idx, "sdu")
        init_state_if_needed(sdu_key, False)
        sdu = st.toggle(
            "SDU",
            key=sdu_key,
            help="Toggle if SDU is needed"
        )
        
        recoair_key = get_state_key(level_idx, area_idx, canopy_idx, "recoair")
        init_state_if_needed(recoair_key, False)
        recoair = st.toggle(
            "RecoAir",
            key=recoair_key,
            help="Toggle if RecoAir is needed"
        )
    
    st.divider()
    
    return {
        "reference_number": ref_number,
        "model": model,
        "configuration": configuration,
        "wall_cladding": {
            "type": "Custom" if cladding_enabled else "None",
            "width": width if cladding_enabled else None,
            "height": height if cladding_enabled else None,
            "position": description if cladding_enabled else None
        },
        "options": {
            "fire_suppression": fire_suppression,
            "uvc": uvc,
            "sdu": sdu,
            "recoair": recoair
        }
    }

def area_form(level_idx: int, area_idx: int, project_type: str, existing_area: Dict[str, Any] = None) -> Dict[str, Any]:
    """Renders form fields for a single area's details."""
    st.markdown(f"#### Area {area_idx + 1}")
    
    # Initialize state for this area
    area_key = f"area_{level_idx}_{area_idx}"
    if area_key not in st.session_state:
        st.session_state[area_key] = {}
    
    # Load existing data if available
    if existing_area:
        st.session_state[area_key] = existing_area
    
    # Area name input
    name_key = f"area_name_{level_idx}_{area_idx}"
    area_name = st.text_input(
        "Area Name *",
        value=st.session_state[area_key].get("name", ""),
        key=name_key,
        help="Enter the area name (e.g., Kitchen, Storage)"
    )
    
    area_data = {
        "name": area_name,
        "canopies": []
    }
    
    if project_type == "Canopy Project" and area_name:
        st.divider()
        
        # Number of canopies
        canopies_key = f"num_canopies_{level_idx}_{area_idx}"
        default_canopies = len(st.session_state[area_key].get("canopies", [])) if "canopies" in st.session_state[area_key] else 1
        
        num_canopies = st.number_input(
            "Number of Canopies",
            min_value=0,
            value=default_canopies,
            key=canopies_key,
            help="Enter the number of canopies for this area"
        )
        
        # Store canopy data
        for i in range(num_canopies):
            with st.container():
                st.markdown(f"#### Canopy {i + 1}")
                
                # Get existing canopy data if available
                existing_canopy = (st.session_state[area_key].get("canopies", []) or [{}])[i] if i < len(st.session_state[area_key].get("canopies", [])) else {}
                
                # Reference number
                ref_key = f"ref_{level_idx}_{area_idx}_{i}"
                ref_number = st.text_input(
                    "Reference Number *",
                    value=existing_canopy.get("reference_number", ""),
                    key=ref_key,
                    help="Enter the canopy reference number"
                )
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Model
                    model_key = f"model_{level_idx}_{area_idx}_{i}"
                    model = st.selectbox(
                        "Model *",
                        options=VALID_CANOPY_MODELS,
                        index=VALID_CANOPY_MODELS.index(existing_canopy.get("model", VALID_CANOPY_MODELS[0])) if existing_canopy.get("model") in VALID_CANOPY_MODELS else 0,
                        key=model_key,
                        help="Select the canopy model"
                    )
                    
                    # Configuration
                    config_key = f"config_{level_idx}_{area_idx}_{i}"
                    configuration = st.selectbox(
                        "Configuration *",
                        options=["Wall", "Island", "Single", "Double"],
                        index=["Wall", "Island", "Single", "Double"].index(existing_canopy.get("configuration", "Wall")) if existing_canopy.get("configuration") in ["Wall", "Island", "Single", "Double"] else 0,
                        key=config_key,
                        help="Select the canopy configuration"
                    )
                
                with col2:
                    st.markdown("**Wall Cladding**")
                    clad_enable_key = f"clad_enable_{level_idx}_{area_idx}_{i}"
                    existing_cladding = existing_canopy.get("wall_cladding", {})
                    # Determine if cladding is enabled based on existing data
                    existing_enabled = existing_cladding.get("type", "None") != "None"
                    cladding_enabled = st.checkbox(
                        "Wall Cladding",
                        value=existing_enabled,
                        key=clad_enable_key,
                        help="Check if wall cladding is required"
                    )
                    
                    if cladding_enabled:
                        col3, col4 = st.columns(2)
                        with col3:
                            width_key = f"width_{level_idx}_{area_idx}_{i}"
                            width = st.number_input(
                                "Width (mm)",
                                min_value=0,
                                value=existing_cladding.get("width", 0),
                                key=width_key,
                                help="Enter the cladding width in millimeters"
                            )
                        
                        with col4:
                            height_key = f"height_{level_idx}_{area_idx}_{i}"
                            height = st.number_input(
                                "Height (mm)",
                                min_value=0,
                                value=existing_cladding.get("height", 0),
                                key=height_key,
                                help="Enter the cladding height in millimeters"
                            )
                        
                        # Description (multi-select for position)
                        desc_key = f"description_{level_idx}_{area_idx}_{i}"
                        # Handle existing position data - convert to list if it's a string or ensure it's a list
                        existing_position = existing_cladding.get("position", [])
                        if isinstance(existing_position, str):
                            # Handle legacy data that might be a single string
                            existing_position = [existing_position] if existing_position else []
                        elif existing_position is None:
                            existing_position = []
                        
                        description = st.multiselect(
                            "Description",
                            options=["rear", "left hand", "right hand"],
                            default=existing_position,
                            key=desc_key,
                            help="Select cladding positions (can select multiple)"
                        )
                
                st.markdown("**Additional Options**")
                opt_col1, opt_col2 = st.columns(2)
                
                with opt_col1:
                    existing_options = existing_canopy.get("options", {})
                    fire_sup_key = f"fire_sup_{level_idx}_{area_idx}_{i}"
                    fire_suppression = st.toggle(
                        "Fire Suppression System",
                        value=existing_options.get("fire_suppression", False),
                        key=fire_sup_key,
                        help="Toggle if fire suppression system is needed"
                    )
                    
                    uvc_key = f"uvc_{level_idx}_{area_idx}_{i}"
                    uvc = st.toggle(
                        "UV-C System",
                        value=existing_options.get("uvc", False),
                        key=uvc_key,
                        help="Toggle if UV-C system is needed"
                    )
                
                with opt_col2:
                    sdu_key = f"sdu_{level_idx}_{area_idx}_{i}"
                    sdu = st.toggle(
                        "SDU",
                        value=existing_options.get("sdu", False),
                        key=sdu_key,
                        help="Toggle if SDU is needed"
                    )
                    
                    recoair_key = f"recoair_{level_idx}_{area_idx}_{i}"
                    recoair = st.toggle(
                        "RecoAir",
                        value=existing_options.get("recoair", False),
                        key=recoair_key,
                        help="Toggle if RecoAir is needed"
                    )
                
                # Only add canopy if it has a reference number
                if ref_number:
                    canopy_data = {
                        "reference_number": ref_number,
                        "model": model,
                        "configuration": configuration,
                        "wall_cladding": {
                            "type": "Custom" if cladding_enabled else "None",
                            "width": width if cladding_enabled else None,
                            "height": height if cladding_enabled else None,
                            "position": description if cladding_enabled else None
                        },
                        "options": {
                            "fire_suppression": fire_suppression,
                            "uvc": uvc,
                            "sdu": sdu,
                            "recoair": recoair
                        }
                    }
                    area_data["canopies"].append(canopy_data)
                
                st.divider()
        
        # Update session state with the latest area data
        st.session_state[area_key] = area_data
    
    return area_data if area_name else None

def project_structure_form() -> List[Dict[str, Any]]:
    """Renders the project structure form."""
    if SessionKeys.PROJECT_TYPE not in st.session_state:
        st.error("Please select a project type first")
        return None
    
    project_type = st.session_state[SessionKeys.PROJECT_TYPE]
    
    # Initialize state for structure
    if "num_levels" not in st.session_state:
        st.session_state.num_levels = 1
    
    # Initialize project structure if it doesn't exist
    if "project_structure" not in st.session_state:
        st.session_state.project_structure = []
    elif "levels" in st.session_state[SessionKeys.PROJECT_DATA]:
        # Load existing data if available and not already loaded
        if not st.session_state.project_structure:
            st.session_state.project_structure = st.session_state[SessionKeys.PROJECT_DATA]["levels"]
            # Update num_levels to match existing data
            st.session_state.num_levels = len(st.session_state.project_structure)
    
    # Number of levels
    num_levels = st.number_input(
        "Number of Levels",
        min_value=1,
        value=st.session_state.num_levels,
        key="num_levels_input",
        help="Enter the number of levels in the project"
    )
    
    # Update session state if number of levels changed
    if num_levels != st.session_state.num_levels:
        st.session_state.num_levels = num_levels
        # Preserve existing levels data up to the new number of levels
        st.session_state.project_structure = st.session_state.project_structure[:num_levels]
        st.rerun()
    
    levels_data = []
    
    for level_idx in range(num_levels):
        st.markdown("---")
        st.subheader(f"Level {level_idx + 1}")
        
        # Get existing level data if available
        existing_level = None
        if level_idx < len(st.session_state.project_structure):
            existing_level = st.session_state.project_structure[level_idx]
        
        # Level name input
        level_name_key = f"level_name_{level_idx}"
        default_name = existing_level.get("level_name", f"Level {level_idx + 1}") if existing_level else f"Level {level_idx + 1}"
        level_name = st.text_input(
            "Level Name *",
            value=default_name,
            key=level_name_key,
            help="Enter a name for this level (e.g., Ground Floor, First Floor, etc.)"
        )
        
        # Number of areas
        num_areas_key = f"num_areas_{level_idx}"
        default_areas = len(existing_level.get("areas", [])) if existing_level else 1
        if num_areas_key not in st.session_state:
            st.session_state[num_areas_key] = default_areas
        
        num_areas = st.number_input(
            "Number of Areas",
            min_value=1,
            value=st.session_state[num_areas_key],
            key=f"{num_areas_key}_input",
            help="Enter the number of areas for this level"
        )
        
        # Update session state if number of areas changed
        if num_areas != st.session_state[num_areas_key]:
            st.session_state[num_areas_key] = num_areas
            st.rerun()
        
        areas_data = []
        
        for area_idx in range(num_areas):
            with st.container():
                # Get existing area data if available
                existing_area = None
                if existing_level and area_idx < len(existing_level.get("areas", [])):
                    existing_area = existing_level["areas"][area_idx]
                
                area_data = area_form(level_idx, area_idx, project_type, existing_area)
                if area_data and area_data["name"]:
                    areas_data.append(area_data)
        
        if areas_data:
            level_data = {
                "level_number": level_idx + 1,
                "level_name": level_name,
                "areas": areas_data
            }
            levels_data.append(level_data)
            
            # Update session state with the latest data
            if level_idx < len(st.session_state.project_structure):
                st.session_state.project_structure[level_idx] = level_data
            else:
                st.session_state.project_structure.append(level_data)
    
    # Save button handling
    if st.button("Save Project Structure", use_container_width=True, type="primary"):
        if not levels_data:
            st.error("Please fill in at least one area name")
            return None
        # Set the save flag
        st.session_state.save_clicked = True
        return levels_data
    
    # Return the current structure data
    return st.session_state.project_structure if st.session_state.project_structure else None 