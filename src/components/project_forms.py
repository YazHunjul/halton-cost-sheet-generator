"""
Project-specific form components for the Halton Cost Sheet Generator.
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
            options=["Wall", "ISLAND"],
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
    
    fire_sup_key = get_state_key(level_idx, area_idx, canopy_idx, "fire_sup")
    init_state_if_needed(fire_sup_key, False)
    fire_suppression = st.toggle(
        "Fire Suppression System",
        key=fire_sup_key,
        help="Toggle if fire suppression system is needed"
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
            "fire_suppression": fire_suppression
        }
    }

def area_form(level_idx: int, area_idx: int, project_type: str, existing_area: Dict[str, Any] = None) -> Dict[str, Any]:
    """Renders form fields for a single area's details."""
    # Add anchor point for this area
    st.markdown(f'<div id="level-{level_idx + 1}-area-{area_idx + 1}"></div>', unsafe_allow_html=True)
    st.markdown(f"### 📍 Area {area_idx + 1}")
    
    # Initialize state for this area
    area_key = f"area_{level_idx}_{area_idx}"
    if area_key not in st.session_state:
        st.session_state[area_key] = {}
    
    # Load existing data if available
    if existing_area:
        st.session_state[area_key] = existing_area
    
    # Area name input
    name_key = f"area_name_{level_idx}_{area_idx}"
    # Initialize area name in session state if not present
    if name_key not in st.session_state and st.session_state[area_key].get("name"):
        st.session_state[name_key] = st.session_state[area_key].get("name")
    
    area_name = st.text_input(
        "Area Name *",
        key=name_key,
        help="Enter the area name (e.g., Kitchen, Storage)"
    )
    
    # Area-level options (UV-C, SDU, RecoAir)
    if area_name:
        st.markdown("**Area Options**")
        opt_col1, opt_col2, opt_col3 = st.columns(3)
        
        existing_options = st.session_state[area_key].get("options", {})
        
        with opt_col1:
            uvc_key = f"area_uvc_{level_idx}_{area_idx}"
            # Initialize UV-C toggle in session state if not present
            if uvc_key not in st.session_state:
                st.session_state[uvc_key] = existing_options.get("uvc", False)
            
            uvc = st.toggle(
                "UV-C System",
                key=uvc_key,
                help="Toggle if UV-C system is needed for this area"
            )
        
        with opt_col2:
            sdu_key = f"area_sdu_{level_idx}_{area_idx}"
            # Initialize SDU toggle in session state if not present
            if sdu_key not in st.session_state:
                st.session_state[sdu_key] = existing_options.get("sdu", False)
            
            sdu = st.toggle(
                "SDU",
                key=sdu_key,
                help="Toggle if SDU is needed for this area"
            )
        
        with opt_col3:
            recoair_key = f"area_recoair_{level_idx}_{area_idx}"
            # Initialize RecoAir toggle in session state if not present
            if recoair_key not in st.session_state:
                st.session_state[recoair_key] = existing_options.get("recoair", False)
            
            recoair = st.toggle(
                "RecoAir",
                key=recoair_key,
                help="Toggle if RecoAir is needed for this area"
            )
    
    area_data = {
        "name": area_name,
        "options": {
            "uvc": uvc if area_name else False,
            "sdu": sdu if area_name else False,
            "recoair": recoair if area_name else False
        },
        "canopies": []
    }
    
    if project_type == "Canopy Project" and area_name:
        st.divider()
        
        # Number of canopies - store separately in session state to prevent reset
        canopies_key = f"num_canopies_{level_idx}_{area_idx}"
        
        # Initialize the canopy count if not already set
        if canopies_key not in st.session_state:
            # Use existing canopies count if available, otherwise default to 1
            existing_canopies_count = len(st.session_state[area_key].get("canopies", [])) if "canopies" in st.session_state[area_key] else 1
            st.session_state[canopies_key] = max(existing_canopies_count, 1)  # Ensure at least 1
        
        # Initialize the number input in session state
        if f"{canopies_key}_input" not in st.session_state:
            st.session_state[f"{canopies_key}_input"] = st.session_state[canopies_key]
        
        num_canopies = st.number_input(
            "Number of Canopies",
            min_value=0,
            key=f"{canopies_key}_input",
            help="Enter the number of canopies for this area"
        )
        
        # Update session state when the number changes
        if num_canopies != st.session_state[canopies_key]:
            st.session_state[canopies_key] = num_canopies
        
        # Store canopy data
        for i in range(num_canopies):
            with st.container():
                st.markdown(f"#### Canopy {i + 1}")
                
                # Get existing canopy data if available
                existing_canopy = (st.session_state[area_key].get("canopies", []) or [{}])[i] if i < len(st.session_state[area_key].get("canopies", [])) else {}
                
                # Reference number
                ref_key = f"ref_{level_idx}_{area_idx}_{i}"
                # Initialize the reference number in session state if not present
                if ref_key not in st.session_state and existing_canopy.get("reference_number"):
                    st.session_state[ref_key] = existing_canopy.get("reference_number")
                
                ref_number = st.text_input(
                    "Reference Number *",
                    key=ref_key,
                    help="Enter the canopy reference number"
                )
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Model
                    model_key = f"model_{level_idx}_{area_idx}_{i}"
                    # Initialize the model in session state if not present
                    if model_key not in st.session_state and existing_canopy.get("model"):
                        st.session_state[model_key] = existing_canopy.get("model")
                    
                    model = st.selectbox(
                        "Model *",
                        options=VALID_CANOPY_MODELS,
                        key=model_key,
                        help="Select the canopy model"
                    )
                    
                    # Configuration
                    config_key = f"config_{level_idx}_{area_idx}_{i}"
                    # Initialize the configuration in session state if not present
                    if config_key not in st.session_state and existing_canopy.get("configuration"):
                        st.session_state[config_key] = existing_canopy.get("configuration")
                    
                    configuration = st.selectbox(
                        "Configuration *",
                        options=["Wall", "Island", "Single", "Double"],
                        key=config_key,
                        help="Select the canopy configuration"
                    )
                
                with col2:
                    st.markdown("**Wall Cladding**")
                    clad_enable_key = f"clad_enable_{level_idx}_{area_idx}_{i}"
                    existing_cladding = existing_canopy.get("wall_cladding", {})
                    # Initialize the cladding checkbox in session state if not present
                    if clad_enable_key not in st.session_state:
                        existing_enabled = existing_cladding.get("type", "None") != "None"
                        st.session_state[clad_enable_key] = existing_enabled
                    
                    cladding_enabled = st.checkbox(
                        "Wall Cladding",
                        key=clad_enable_key,
                        help="Check if wall cladding is required"
                    )
                    
                    if cladding_enabled:
                        col3, col4 = st.columns(2)
                        with col3:
                            width_key = f"width_{level_idx}_{area_idx}_{i}"
                            # Initialize width in session state if not present
                            if width_key not in st.session_state and existing_cladding.get("width") is not None:
                                st.session_state[width_key] = existing_cladding.get("width", 0)
                            
                            width = st.number_input(
                                "Width (mm)",
                                min_value=0,
                                key=width_key,
                                help="Enter the cladding width in millimeters"
                            )
                        
                        with col4:
                            height_key = f"height_{level_idx}_{area_idx}_{i}"
                            # Initialize height in session state if not present
                            if height_key not in st.session_state and existing_cladding.get("height") is not None:
                                st.session_state[height_key] = existing_cladding.get("height", 0)
                            
                            height = st.number_input(
                                "Height (mm)",
                                min_value=0,
                                key=height_key,
                                help="Enter the cladding height in millimeters"
                            )
                        
                        # Description (multi-select for position)
                        desc_key = f"description_{level_idx}_{area_idx}_{i}"
                        # Initialize description in session state if not present
                        if desc_key not in st.session_state:
                            # Handle existing position data - convert to list if it's a string or ensure it's a list
                            existing_position = existing_cladding.get("position", [])
                            if isinstance(existing_position, str):
                                # Handle legacy data that might be a single string
                                existing_position = [existing_position] if existing_position else []
                            elif existing_position is None:
                                existing_position = []
                            st.session_state[desc_key] = existing_position
                        
                        description = st.multiselect(
                            "Description",
                            options=["rear", "left hand", "right hand"],
                            key=desc_key,
                            help="Select cladding positions (can select multiple)"
                        )
                
                st.markdown("**Additional Options**")
                
                existing_options = existing_canopy.get("options", {})
                fire_sup_key = f"fire_sup_{level_idx}_{area_idx}_{i}"
                # Initialize fire suppression in session state if not present
                if fire_sup_key not in st.session_state:
                    st.session_state[fire_sup_key] = existing_options.get("fire_suppression", False)
                
                fire_suppression = st.toggle(
                    "Fire Suppression System",
                    key=fire_sup_key,
                    help="Toggle if fire suppression system is needed"
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
                            "fire_suppression": fire_suppression
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
    # Initialize num_levels in session state for the input
    if "num_levels_input" not in st.session_state:
        st.session_state["num_levels_input"] = st.session_state.num_levels
    
    num_levels = st.number_input(
        "Number of Levels",
        min_value=1,
        key="num_levels_input",
        help="Enter the number of levels in the project"
    )
    
    # Update session state if number of levels changed
    if num_levels != st.session_state.num_levels:
        st.session_state.num_levels = num_levels
        # Preserve existing levels data up to the new number of levels
        st.session_state.project_structure = st.session_state.project_structure[:num_levels]
        st.rerun()
    
    # Create level and area navigation in sidebar
    with st.sidebar:
        if num_levels >= 1:
            st.markdown("### 🏢 Project Navigation")
            st.markdown("*Click to jump to any section:*")
            
            for level_idx in range(num_levels):
                # Get existing level data for name
                existing_level = None
                if level_idx < len(st.session_state.project_structure):
                    existing_level = st.session_state.project_structure[level_idx]
                
                level_name = existing_level.get("level_name", f"Level {level_idx + 1}") if existing_level else f"Level {level_idx + 1}"
                
                # Create level anchor link
                st.markdown(f"🔗 **[{level_name}](#level-{level_idx + 1})**")
                
                # Add area links if they exist
                if existing_level and existing_level.get("areas"):
                    for area_idx, area in enumerate(existing_level["areas"]):
                        area_name = area.get("name", f"Area {area_idx + 1}")
                        # Show canopy count if available
                        canopy_count = len(area.get("canopies", []))
                        canopy_info = f" ({canopy_count} canopies)" if canopy_count > 0 else ""
                        st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;📍 [{area_name}{canopy_info}](#level-{level_idx + 1}-area-{area_idx + 1})")
                else:
                    # Show placeholder for areas that will be created
                    num_areas_key = f"num_areas_{level_idx}"
                    if num_areas_key in st.session_state:
                        num_areas = st.session_state[num_areas_key]
                        for area_idx in range(num_areas):
                            st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;📍 [Area {area_idx + 1}](#level-{level_idx + 1}-area-{area_idx + 1})")
            
            st.markdown("---")
    
    levels_data = []
    
    for level_idx in range(num_levels):
        # Get existing level data if available
        existing_level = None
        if level_idx < len(st.session_state.project_structure):
            existing_level = st.session_state.project_structure[level_idx]
        
        # Determine the level name for the section title
        default_name = existing_level.get("level_name", f"Level {level_idx + 1}") if existing_level else f"Level {level_idx + 1}"
        
        # Create anchor point and use expander for each level
        st.markdown(f'<div id="level-{level_idx + 1}"></div>', unsafe_allow_html=True)
        
        # Use expander for each level
        with st.expander(f"🏢 {default_name}", expanded=True):
            # Level name input
            level_name_key = f"level_name_{level_idx}"
            # Initialize level name in session state if not present
            if level_name_key not in st.session_state:
                st.session_state[level_name_key] = default_name
            
            level_name = st.text_input(
                "Level Name *",
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
                key=f"{num_areas_key}_input",
                help="Enter the number of areas for this level"
            )
            
            # Update session state if number of areas changed
            if num_areas != st.session_state[num_areas_key]:
                st.session_state[num_areas_key] = num_areas
            
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
                    
                    # Add separator between areas (except for the last one)
                    if area_idx < num_areas - 1:
                        st.markdown("---")
            
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