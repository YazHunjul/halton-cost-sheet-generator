"""
Main Streamlit application for the Halton Cost Sheet Generator.
"""
import streamlit as st
from components.forms import general_project_form
from components.project_forms import project_structure_form
from config.constants import SessionKeys, PROJECT_TYPES
from utils.excel import read_excel_project_data
from utils.word import generate_quotation_document
from openpyxl import load_workbook
import os

def display_project_summary(project_data: dict):
    """Display a formatted summary of the project data."""
    st.header("📋 Project Summary")
    
    # General Information
    st.subheader("General Information")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Project Name:**", project_data.get("project_name"))
        st.write("**Project Number:**", project_data.get("project_number"))
        st.write("**Date:**", project_data.get("date"))
        st.write("**Customer:**", project_data.get("customer"))
        st.write("**Company:**", project_data.get("company"))
    
    with col2:
        st.write("**Project Location:**", project_data.get("project_location") or project_data.get("location"))
        st.write("**Delivery Location:**", project_data.get("delivery_location"))
        st.write("**Address:**", project_data.get("address"))
        st.write("**Sales Contact:**", project_data.get("sales_contact"))
        st.write("**Estimator:**", project_data.get("estimator"))
    
    # Project Structure
    if "levels" in project_data:
        st.markdown("---")
        st.subheader("Project Structure")
        
        for level in project_data["levels"]:
            with st.expander(f"Level {level['level_number']}", expanded=True):
                for area in level["areas"]:
                    st.markdown(f"### 📍 Area: {area['name']}")
                    
                    # Area-level options
                    if "options" in area:
                        st.markdown("**Area Options:**")
                        opt_col1, opt_col2, opt_col3 = st.columns(3)
                        with opt_col1:
                            st.write("✓ UV-C System" if area["options"]["uvc"] else "✗ UV-C System")
                        with opt_col2:
                            st.write("✓ SDU" if area["options"]["sdu"] else "✗ SDU")
                        with opt_col3:
                            st.write("✓ RecoAir" if area["options"]["recoair"] else "✗ RecoAir")
                        st.markdown("---")
                    
                    if area["canopies"]:
                        for i, canopy in enumerate(area["canopies"], 1):
                            st.markdown(f"#### 🔹 Canopy {i}")
                            
                            # Basic Info
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write("**Reference Number:**", canopy["reference_number"])
                                st.write("**Model:**", canopy["model"])
                                st.write("**Configuration:**", canopy["configuration"])
                            
                            # Wall Cladding
                            if canopy["wall_cladding"]["type"] != "None":
                                with col2:
                                    st.markdown("**Wall Cladding:**")
                                    st.write("- Type:", canopy["wall_cladding"]["type"])
                                    st.write("- Width:", f"{canopy['wall_cladding']['width']}mm")
                                    st.write("- Height:", f"{canopy['wall_cladding']['height']}mm")
                                    # Handle position as a list
                                    position = canopy["wall_cladding"]["position"]
                                    if isinstance(position, list):
                                        position_str = ", ".join(position) if position else "None"
                                    else:
                                        position_str = position if position else "None"
                                    st.write("- Position:", position_str)
                            
                            # Canopy Options (only fire suppression now)
                            st.markdown("**Canopy Options:**")
                            st.write("✓ Fire Suppression" if canopy["options"]["fire_suppression"] else "✗ Fire Suppression")
                            
                            st.markdown("---")
                    else:
                        st.write("No canopies in this area")
                    
                    st.markdown("---")

def word_generation_page():
    """Page for generating Word documents from Excel files."""
    st.header("📄 Generate Word Quotation")
    st.markdown("Upload an Excel cost sheet to generate a Word quotation document.")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx', 'xls'],
        help="Upload a cost sheet Excel file generated by this application"
    )
    
    if uploaded_file is not None:
        try:
            # Save uploaded file temporarily
            temp_path = f"temp_excel_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Read project data from Excel
            with st.spinner("Reading project data from Excel..."):
                project_data = read_excel_project_data(temp_path)
                
                # Detect and set project type based on the data (required for Excel generation)
                if not project_data.get("project_type"):
                    # Default to "Commercial Kitchen" as it's the most common type
                    project_data["project_type"] = "Commercial Kitchen"
            
            # Display summary of extracted data
            st.success("✅ Successfully extracted project data from Excel!")
            
            # Analyze project to show what type it is
            from utils.word import analyze_project_areas
            has_canopies, has_recoair, is_recoair_only, has_uv = analyze_project_areas(project_data)
            
            # Show project type analysis
            if is_recoair_only:
                st.info("🔄 **Project Type:** RecoAir-only project detected")
            elif has_canopies and has_recoair:
                st.info("🔄 **Project Type:** Mixed project (Canopies + RecoAir) detected")
            elif has_canopies:
                st.info("🔄 **Project Type:** Canopy-only project detected")
            else:
                st.warning("⚠️ **Project Type:** No canopies or RecoAir systems detected")
            
            # Show download button first for quick access
            st.markdown("---")
            st.subheader("📥 Quick Download")
            
            try:
                with st.spinner("Preparing download..."):
                    # Generate Word documents for download
                    download_word_path = generate_quotation_document(project_data, temp_path)
                
                # Provide download button
                if download_word_path.endswith('.zip'):
                    # Multiple documents in zip file
                    with open(download_word_path, "rb") as file:
                        zip_filename = os.path.basename(download_word_path)
                        st.download_button(
                            label="📥 Download All Documents (ZIP)",
                            data=file.read(),
                            file_name=zip_filename,
                            mime="application/zip",
                            type="primary"
                        )
                    st.info("📦 ZIP file contains both Main Quotation and RecoAir Quotation documents.")
                else:
                    # Single document
                    doc_filename = os.path.basename(download_word_path)
                    with open(download_word_path, "rb") as file:
                        # Determine appropriate label based on document type
                        if is_recoair_only:
                            label = "📥 Download RecoAir Quotation"
                        else:
                            label = "📥 Download Quotation"
                        
                        st.download_button(
                            label=label,
                            data=file.read(),
                            file_name=doc_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    # Show appropriate success message
                    if is_recoair_only:
                        st.info("📄 RecoAir quotation document ready for download.")
                    else:
                        st.info("📄 Quotation document ready for download.")
                        
            except Exception as e:
                st.error(f"❌ Error preparing download: {str(e)}")
            
            # Automatically generate and show preview of existing document
            st.markdown("---")
            st.subheader("📄 Document Preview")
            
            try:
                with st.spinner("Generating preview of current document..."):
                    # Generate Word documents to get the current state
                    word_path = generate_quotation_document(project_data, temp_path)
                
                # Show preview for documents
                if not word_path.endswith('.zip'):
                    # Single document preview
                    from utils.word_preview import check_preview_requirements, preview_word_document
                    capabilities = check_preview_requirements()
                    
                    col1, col2 = st.columns([3, 1])
                    with col2:
                        use_advanced = st.checkbox(
                            "Enhanced Preview", 
                            value=capabilities['advanced_preview'],
                            disabled=not capabilities['advanced_preview'],
                            help="Uses pypandoc for better formatting (if available)",
                            key="upload_preview_advanced"
                        )
                        
                        if not capabilities['advanced_preview']:
                            if "not installed" in str(capabilities.get('pandoc_version', '')):
                                st.info("💡 Install pypandoc for enhanced preview")
                            else:
                                st.warning(f"⚠️ {capabilities.get('pandoc_version', 'Pandoc issue')}")
                        elif capabilities['pandoc_version']:
                            st.success(f"✅ Pandoc v{capabilities['pandoc_version']}")
                    
                    with col1:
                        st.write("**Preview Mode:**", "Enhanced" if use_advanced else "Basic")
                        if capabilities['table_preservation']:
                            st.write("✅ Table preservation enabled")
                    
                    # Generate and display preview
                    try:
                        with st.spinner("Rendering preview..."):
                            preview_html = preview_word_document(word_path, use_advanced)
                        
                        # Display preview
                        st.components.v1.html(preview_html, height=650, scrolling=True)
                        
                        # Preview stats
                        try:
                            from docx import Document
                            doc = Document(word_path)
                            table_count = len(doc.tables)
                            paragraph_count = len([p for p in doc.paragraphs if p.text.strip()])
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("📄 Paragraphs", paragraph_count)
                            with col2:
                                st.metric("📊 Tables", table_count)
                            with col3:
                                file_size = os.path.getsize(word_path)
                                st.metric("💾 File Size", f"{file_size/1024:.1f} KB")
                                
                        except Exception as e:
                            st.write(f"Preview stats unavailable: {str(e)}")
                            
                    except Exception as e:
                        st.error(f"❌ Error generating preview: {str(e)}")
                        st.write("Preview failed, but you can still generate the document below.")
                else:
                    # Multiple documents - show previews for both
                    st.info("📦 Multiple documents detected - showing previews for both documents:")
                    
                    # Extract and preview individual documents from the ZIP
                    import zipfile
                    import tempfile
                    from utils.word_preview import check_preview_requirements, preview_word_document
                    
                    capabilities = check_preview_requirements()
                    
                    # Preview options (shared for both documents)
                    col1, col2 = st.columns([3, 1])
                    with col2:
                        use_advanced = st.checkbox(
                            "Enhanced Preview", 
                            value=capabilities['advanced_preview'],
                            disabled=not capabilities['advanced_preview'],
                            help="Uses pypandoc for better formatting (if available)",
                            key="upload_preview_advanced_multi"
                        )
                        
                        if not capabilities['advanced_preview']:
                            if "not installed" in str(capabilities.get('pandoc_version', '')):
                                st.info("💡 Install pypandoc for enhanced preview")
                            else:
                                st.warning(f"⚠️ {capabilities.get('pandoc_version', 'Pandoc issue')}")
                        elif capabilities['pandoc_version']:
                            st.success(f"✅ Pandoc v{capabilities['pandoc_version']}")
                    
                    with col1:
                        st.write("**Preview Mode:**", "Enhanced" if use_advanced else "Basic")
                        if capabilities['table_preservation']:
                            st.write("✅ Table preservation enabled")
                    
                    try:
                        with zipfile.ZipFile(word_path, 'r') as zip_ref:
                            file_list = zip_ref.namelist()
                            
                            for i, filename in enumerate(file_list):
                                if filename.endswith('.docx'):
                                    st.markdown(f"### 📄 Document {i+1}: {filename}")
                                    
                                    # Extract to temporary file
                                    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
                                        tmp_file.write(zip_ref.read(filename))
                                        tmp_path = tmp_file.name
                                    
                                    try:
                                        # Generate and display preview
                                        with st.spinner(f"Rendering preview for {filename}..."):
                                            preview_html = preview_word_document(tmp_path, use_advanced)
                                        
                                        # Display preview
                                        st.components.v1.html(preview_html, height=500, scrolling=True)
                                        
                                        # Preview stats
                                        try:
                                            from docx import Document
                                            doc = Document(tmp_path)
                                            table_count = len(doc.tables)
                                            paragraph_count = len([p for p in doc.paragraphs if p.text.strip()])
                                            
                                            col1, col2, col3 = st.columns(3)
                                            with col1:
                                                st.metric("📄 Paragraphs", paragraph_count)
                                            with col2:
                                                st.metric("📊 Tables", table_count)
                                            with col3:
                                                file_size = os.path.getsize(tmp_path)
                                                st.metric("💾 File Size", f"{file_size/1024:.1f} KB")
                                                
                                        except Exception as e:
                                            st.write(f"Preview stats unavailable: {str(e)}")
                                    
                                    except Exception as e:
                                        st.error(f"❌ Error generating preview for {filename}: {str(e)}")
                                    
                                    finally:
                                        # Clean up temp file
                                        if os.path.exists(tmp_path):
                                            os.unlink(tmp_path)
                                    
                                    if i < len(file_list) - 1:  # Add separator between documents
                                        st.markdown("---")
                    
                    except Exception as e:
                        st.error(f"❌ Error extracting documents from ZIP: {str(e)}")
                        st.write("Preview failed, but you can still generate the documents below.")
                    
            except Exception as e:
                st.error(f"❌ Error generating preview: {str(e)}")
                st.write("Preview failed, but you can still generate the document below.")
            
            st.markdown("---")
            
            with st.expander("📋 Extracted Project Data", expanded=False):
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Project Name:**", project_data.get("project_name"))
                    st.write("**Project Number:**", project_data.get("project_number"))
                    st.write("**Customer:**", project_data.get("customer"))
                    st.write("**Date:**", project_data.get("date"))
                
                with col2:
                    st.write("**Project Location:**", project_data.get("project_location") or project_data.get("location"))
                    st.write("**Delivery Location:**", project_data.get("delivery_location"))
                    st.write("**Estimator:**", project_data.get("estimator"))
                    st.write("**Estimator Initials (from Excel):**", project_data.get("estimator_initials"))
                    
                    # Show combined initials calculation, reference variable, customer first name, and quote title
                    from utils.word import get_sales_contact_info, get_combined_initials, generate_reference_variable, get_customer_first_name, generate_quote_title
                    estimator_name = project_data.get("estimator", "")
                    sales_contact = get_sales_contact_info(estimator_name)
                    combined_initials = get_combined_initials(sales_contact['name'], estimator_name)
                    reference_variable = generate_reference_variable(
                        project_data.get('project_number', ''), 
                        sales_contact['name'], 
                        estimator_name
                    )
                    customer_first_name = get_customer_first_name(project_data.get('customer', ''))
                    quote_title = generate_quote_title(project_data.get('revision', 'A'))
                    st.write("**Combined Initials (Sales/Estimator):**", combined_initials)
                    st.write("**Reference Variable:**", reference_variable)
                    st.write("**Customer First Name:**", customer_first_name)
                    st.write("**Quote Title:**", quote_title)
                    st.write("**Revision:**", project_data.get('revision', 'A'))
                    st.write("**Sales Contact:**", sales_contact['name'])
                    
                    st.write("**Levels Found:**", len(project_data.get("levels", [])))
                
                # Show detailed analysis
                st.markdown("---")
                st.markdown("**📊 Project Analysis:**")
                analysis_col1, analysis_col2, analysis_col3 = st.columns(3)
                with analysis_col1:
                    st.write("**Has Canopies:**", "✅ Yes" if has_canopies else "❌ No")
                with analysis_col2:
                    st.write("**Has RecoAir:**", "✅ Yes" if has_recoair else "❌ No")
                with analysis_col3:
                    st.write("**RecoAir Only:**", "✅ Yes" if is_recoair_only else "❌ No")
                
                # Show areas and their options
                if project_data.get("levels"):
                    st.markdown("**🏢 Areas Found:**")
                    for level in project_data.get("levels", []):
                        for area in level.get("areas", []):
                            area_name = f"{level.get('level_name', '')} - {area.get('name', '')}"
                            canopy_count = len(area.get('canopies', []))
                            options = area.get('options', {})
                            
                            st.write(f"• **{area_name}**: {canopy_count} canopies")
                            if options.get('uvc'):
                                st.write("  - ✅ UV-C System")
                            if options.get('sdu'):
                                st.write("  - ✅ SDU")
                            if options.get('recoair'):
                                st.write("  - ✅ RecoAir System")
            
            # Show what documents will be generated
            st.markdown("---")
            st.markdown("**📄 Documents to Generate:**")
            if is_recoair_only:
                st.info("📋 **RecoAir Quotation** will be generated (single document)")
                st.write("💡 Your Excel file now has dynamic pricing - totals update automatically!")
            elif has_canopies and has_recoair:
                st.info("📦 **ZIP Package** will be generated containing:")
                st.write("• Main Quotation (for canopies)")
                st.write("• RecoAir Quotation (for RecoAir systems)")
                st.write("💡 Your Excel file now has dynamic pricing - totals update automatically!")
            elif has_canopies:
                st.info("📋 **Main Quotation** will be generated (single document)")
                st.write("💡 Your Excel file now has dynamic pricing - totals update automatically!")
            else:
                st.warning("⚠️ No documents can be generated - no systems detected")
            
            # Generate Word document
            if st.button("📄 Generate Word Quotation", type="primary"):
                try:
                    with st.spinner("Generating Word quotation document(s)..."):
                        # Generate Word documents only (Excel has dynamic pricing now)
                        word_path = generate_quotation_document(project_data, temp_path)
                    
                    st.success("✅ Word quotation document(s) generated successfully!")
                    
                    # Determine file type and provide appropriate download button
                    if word_path.endswith('.zip'):
                        # Multiple documents in zip file
                        with open(word_path, "rb") as file:
                            # Extract filename from the generated path
                            zip_filename = os.path.basename(word_path)
                            st.download_button(
                                label="📥 Download Quotation Documents (ZIP)",
                                data=file.read(),
                                file_name=zip_filename,
                                mime="application/zip"
                            )
                        st.info("📦 Multiple quotation documents generated and packaged in ZIP file.")
                    else:
                        # Single document - automatically show preview with download option
                        doc_filename = os.path.basename(word_path)
                        
                        # Determine appropriate success message based on document type
                        if is_recoair_only:
                            st.info("📄 RecoAir quotation document generated successfully.")
                        else:
                            st.info("📄 Quotation document generated successfully.")
                        
                        # Show download button first
                        with open(word_path, "rb") as file:
                            # Determine appropriate label based on document type
                            if is_recoair_only:
                                label = "📥 Download RecoAir Quotation"
                            else:
                                label = "📥 Download Quotation"
                            
                            st.download_button(
                                label=label,
                                data=file.read(),
                                file_name=doc_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        
                        # Automatically show preview below
                        st.markdown("---")
                        from utils.word_preview import preview_with_download
                        
                        # Show preview with a more compact interface
                        st.subheader("📄 Document Preview")
                        
                        # Check capabilities and show preview options
                        from utils.word_preview import check_preview_requirements
                        capabilities = check_preview_requirements()
                        
                        col1, col2 = st.columns([3, 1])
                        with col2:
                            use_advanced = st.checkbox(
                                "Enhanced Preview", 
                                value=capabilities['advanced_preview'],
                                disabled=not capabilities['advanced_preview'],
                                help="Uses pypandoc for better formatting (if available)"
                            )
                            
                            if not capabilities['advanced_preview']:
                                st.info("💡 Install pypandoc for enhanced preview")
                            elif capabilities['pandoc_version']:
                                st.caption(f"Pandoc version: {capabilities['pandoc_version']}")
                        
                        with col1:
                            st.write("**Preview Mode:**", "Enhanced" if use_advanced else "Basic")
                            if capabilities['table_preservation']:
                                st.write("✅ Table preservation enabled")
                        
                        # Generate and display preview
                        try:
                            from utils.word_preview import preview_word_document
                            with st.spinner("Generating preview..."):
                                preview_html = preview_word_document(word_path, use_advanced)
                            
                            # Display preview
                            st.components.v1.html(preview_html, height=650, scrolling=True)
                            
                            # Preview stats
                            try:
                                from docx import Document
                                doc = Document(word_path)
                                table_count = len(doc.tables)
                                paragraph_count = len([p for p in doc.paragraphs if p.text.strip()])
                                
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("📄 Paragraphs", paragraph_count)
                                with col2:
                                    st.metric("📊 Tables", table_count)
                                with col3:
                                    file_size = os.path.getsize(word_path)
                                    st.metric("💾 File Size", f"{file_size/1024:.1f} KB")
                                    
                            except Exception as e:
                                st.write(f"Preview stats unavailable: {str(e)}")
                                
                        except Exception as e:
                            st.error(f"❌ Error generating preview: {str(e)}")
                            st.write("The document was generated successfully but preview failed. You can still download it above.")
                    
                    # Add note about dynamic pricing
                    st.success("✨ **Your Excel file now has dynamic pricing!** The JOB TOTAL sheet will automatically update when you edit any individual sheet prices.")
                
                except Exception as e:
                    st.error(f"❌ Error generating Word document: {str(e)}")
            
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
        except Exception as e:
            error_message = str(e)
            
            # Check if this is a validation error with detailed information
            if "Data validation errors found:" in error_message:
                st.error("❌ **Excel File Validation Errors**")
                st.markdown("The following data validation errors were found in your Excel file:")
                
                # Split the error message to extract the validation details
                parts = error_message.split("Data validation errors found:")
                if len(parts) > 1:
                    validation_details = parts[1].strip()
                    # Display each validation error in an expandable section
                    with st.expander("📋 **Detailed Error Information**", expanded=True):
                        st.markdown(validation_details)
                
                st.markdown("---")
                st.markdown("### 🔧 **How to Fix:**")
                st.markdown("1. Open your Excel file")
                st.markdown("2. Navigate to the specific cells mentioned above")
                st.markdown("3. Ensure all numeric fields contain valid numbers (not letters or text)")
                st.markdown("4. Save the file and try uploading again")
                
                st.info("💡 **Tip:** The most common issue is entering letters in numeric fields like 'Testing and Commissioning' prices.")
                
            else:
                st.error(f"❌ Error reading Excel file: {error_message}")
            
            if os.path.exists(temp_path):
                os.remove(temp_path)

def revision_page():
    """Page for creating new revisions from existing Excel files."""
    st.header("📝 Create New Revision")
    st.markdown("Upload an existing Excel cost sheet to create a new revision with the same data.")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file to revise",
        type=['xlsx', 'xls'],
        help="Upload an existing cost sheet Excel file to create a new revision"
    )
    
    if uploaded_file is not None:
        try:
            # Save uploaded file temporarily
            temp_path = f"temp_revision_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Read project data from Excel
            with st.spinner("Reading project data from Excel..."):
                project_data = read_excel_project_data(temp_path)
                
                # Ensure project type is set (required for Excel generation)
                if not project_data.get("project_type"):
                    # Default to "Commercial Kitchen" as it's the most common type
                    project_data["project_type"] = "Commercial Kitchen"
            
            st.success("✅ Successfully extracted project data from Excel!")
            
            # Show current revision info
            current_revision = project_data.get('revision', 'A')
            st.info(f"📋 **Current Revision:** {current_revision}")
            
            # Display project summary
            with st.expander("📋 Project Information", expanded=False):
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Project Name:**", project_data.get("project_name"))
                    st.write("**Project Number:**", project_data.get("project_number"))
                    st.write("**Customer:**", project_data.get("customer"))
                    st.write("**Date:**", project_data.get("date"))
                
                with col2:
                    st.write("**Project Location:**", project_data.get("project_location") or project_data.get("location"))
                    st.write("**Delivery Location:**", project_data.get("delivery_location"))
                    st.write("**Estimator:**", project_data.get("estimator"))
                    st.write("**Current Revision:**", current_revision)
                    
                    # Show project analysis
                    from utils.word import analyze_project_areas
                    has_canopies, has_recoair, is_recoair_only, has_uv = analyze_project_areas(project_data)
                    st.write("**Has Canopies:**", "✅ Yes" if has_canopies else "❌ No")
                    st.write("**Has RecoAir:**", "✅ Yes" if has_recoair else "❌ No")
                    st.write("**Has UV Canopies:**", "✅ Yes" if has_uv else "❌ No")
                    st.write("**Levels Found:**", len(project_data.get("levels", [])))
            
            # Revision options
            st.markdown("---")
            st.subheader("🔄 Revision Options")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Auto-increment revision
                next_revision = chr(ord(current_revision) + 1) if current_revision and len(current_revision) == 1 and current_revision < 'Z' else 'B'
                st.write(f"**Suggested Next Revision:** {next_revision}")
                
                revision_choice = st.radio(
                    "Choose revision method:",
                    ["Auto-increment (recommended)", "Custom revision letter"],
                    help="Auto-increment will automatically suggest the next revision letter"
                )
            
            with col2:
                if revision_choice == "Custom revision letter":
                    new_revision = st.text_input(
                        "Enter new revision letter:",
                        value=next_revision,
                        max_chars=3,
                        help="Enter a revision letter (e.g., B, C, D, etc.)"
                    ).upper()
                else:
                    new_revision = next_revision
                    st.write(f"**New Revision will be:** {new_revision}")
            
            # Optional: Update date
            update_date = st.checkbox("Update date to today", value=True)
            
            if update_date:
                from datetime import datetime
                new_date = datetime.now().strftime("%d/%m/%Y")
                st.write(f"**New Date:** {new_date}")
            else:
                new_date = project_data.get("date", "")
                st.write(f"**Date will remain:** {new_date}")
            
            # Generate new revision
            if st.button("🔄 Create New Revision", type="primary"):
                try:
                    with st.spinner(f"Generating revision {new_revision}..."):
                        from utils.excel import create_revision_from_existing
                        
                        # Create new revision from existing file (preserves all data)
                        output_path = create_revision_from_existing(
                            temp_path, 
                            new_revision, 
                            new_date if update_date else None
                        )
                        
                        # Read the file for download
                        with open(output_path, "rb") as file:
                            excel_data = file.read()
                    
                    st.success(f"✅ Revision {new_revision} created successfully!")
                    
                    # Create download filename: "Project Number Cost Sheet Date"
                    project_number = project_data.get('project_number', 'unknown')
                    date_str = new_date.replace('/', '') if new_date else ''
                    download_filename = f"{project_number} Cost Sheet {date_str}.xlsx"
                    
                    st.download_button(
                        label=f"📥 Download Revision {new_revision}",
                        data=excel_data,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Show what changed
                    st.info(f"📋 **Changes Made:**")
                    st.write(f"• Revision updated: {current_revision} → {new_revision}")
                    if update_date:
                        st.write(f"• Date updated: {project_data.get('date', 'N/A')} → {new_date}")
                    st.write("• ✅ All existing data preserved (canopies, pricing, formulas)")
                    st.write("• ✅ Dynamic pricing formulas maintained")
                    st.write("• ✅ All manual entries and calculations preserved")
                    
                except Exception as e:
                    st.error(f"❌ Error creating revision: {str(e)}")
            
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
        except Exception as e:
            error_message = str(e)
            
            # Check if this is a validation error with detailed information
            if "Data validation errors found:" in error_message:
                st.error("❌ **Excel File Validation Errors**")
                st.markdown("The following data validation errors were found in your Excel file:")
                
                # Split the error message to extract the validation details
                parts = error_message.split("Data validation errors found:")
                if len(parts) > 1:
                    validation_details = parts[1].strip()
                    # Display each validation error in an expandable section
                    with st.expander("📋 **Detailed Error Information**", expanded=True):
                        st.markdown(validation_details)
                
                st.markdown("---")
                st.markdown("### 🔧 **How to Fix:**")
                st.markdown("1. Open your Excel file")
                st.markdown("2. Navigate to the specific cells mentioned above")
                st.markdown("3. Ensure all numeric fields contain valid numbers (not letters or text)")
                st.markdown("4. Save the file and try uploading again")
                
                st.info("💡 **Tip:** The most common issue is entering letters in numeric fields like 'Testing and Commissioning' prices.")
                
            else:
                st.error(f"❌ Error reading Excel file: {error_message}")
            
            if os.path.exists(temp_path):
                os.remove(temp_path)

def initialize_session_state():
    """Initialize all session state variables if they don't exist."""
    if SessionKeys.PROJECT_DATA not in st.session_state:
        st.session_state[SessionKeys.PROJECT_DATA] = {}
    
    if SessionKeys.CURRENT_STEP not in st.session_state:
        st.session_state[SessionKeys.CURRENT_STEP] = 1
        
    if SessionKeys.PROJECT_TYPE not in st.session_state:
        st.session_state[SessionKeys.PROJECT_TYPE] = None

def navigation_buttons():
    """Render navigation buttons based on current step."""
    cols = st.columns([1, 1, 1])
    
    with cols[0]:
        if st.session_state[SessionKeys.CURRENT_STEP] > 1:
            if st.button("⬅️ Go Back", use_container_width=True):
                st.session_state[SessionKeys.CURRENT_STEP] -= 1
                # Don't rerun here, let the main flow handle it

def main():
    # Set page config
    st.set_page_config(
        page_title="Halton Cost Sheet Generator",
        page_icon="🏢",
        layout="wide"
    )
    
    # Initialize session state
    initialize_session_state()
    
    # Sidebar for navigation
    with st.sidebar:
        st.title("Navigation")
        page = st.radio("Choose a page:", ["Create New Project", "Generate Word Document", "Create New Revision"])
        
        st.markdown("---")
        st.subheader("Debug Info")
        st.write("Current Step:", st.session_state[SessionKeys.CURRENT_STEP])
        st.write("Project Type:", st.session_state[SessionKeys.PROJECT_TYPE])
        if st.button("Start Over"):
            # Clear all session state except the current page
            for key in st.session_state.keys():
                if key != "pages_initialized":  # Streamlit internal state
                    del st.session_state[key]
            initialize_session_state()
            st.rerun()
    
    # Header
    st.title("Halton Cost Sheet Generator")
    st.markdown("---")
    
    if page == "Generate Word Document":
        word_generation_page()
        return
    elif page == "Create New Revision":
        revision_page()
        return
    
    # Project Type Selection
    if st.session_state[SessionKeys.PROJECT_TYPE] is None:
        st.info("Please select a project type to begin")
        project_type = st.selectbox(
            "Select Project Type *",
            options=PROJECT_TYPES,
            index=None,
            help="Choose the type of project you're creating"
        )
        
        if project_type:
            st.session_state[SessionKeys.PROJECT_TYPE] = project_type
            # Store project type in project data
            st.session_state[SessionKeys.PROJECT_DATA]["project_type"] = project_type
            st.rerun()
    
    # Show current project type
    st.info(f"Project Type: {st.session_state[SessionKeys.PROJECT_TYPE]}")
    
    # Navigation buttons
    navigation_buttons()
    
    # Step 1: General Project Information
    if st.session_state[SessionKeys.CURRENT_STEP] == 1:
        st.header("Step 1: General Project Information")
        project_data = general_project_form()
        
        if project_data:
            # Ensure project type is preserved
            project_data["project_type"] = st.session_state[SessionKeys.PROJECT_TYPE]
            st.session_state[SessionKeys.PROJECT_DATA].update(project_data)
            st.success("Project information saved!")
            st.session_state[SessionKeys.CURRENT_STEP] = 2
            st.rerun()
    
    # Step 2: Project Structure
    elif st.session_state[SessionKeys.CURRENT_STEP] == 2:
        # Show Step 1 data in expander
        with st.expander("Step 1: Project Information", expanded=False):
            st.json(st.session_state[SessionKeys.PROJECT_DATA])
        
        st.header("Step 2: Project Structure")
        
        # Get any existing structure data
        existing_structure = st.session_state[SessionKeys.PROJECT_DATA].get("levels", None)
        if existing_structure:
            st.success("Project structure data exists. You can modify it below.")
        
        levels_data = project_structure_form()
        
        # Only update and proceed if Save button was clicked
        if levels_data and st.session_state.get("save_clicked", False):
            st.session_state[SessionKeys.PROJECT_DATA]["levels"] = levels_data
            st.success("Project structure saved successfully!")
            st.session_state[SessionKeys.CURRENT_STEP] = 3
            # Clear the save flag
            st.session_state.save_clicked = False
            st.rerun()
    
    # Step 3: Review
    elif st.session_state[SessionKeys.CURRENT_STEP] == 3:
        display_project_summary(st.session_state[SessionKeys.PROJECT_DATA])
        
        st.markdown("---")
        st.subheader("📊 Generate Cost Sheet")
        
        if st.button("Generate Excel Cost Sheet", type="primary"):
            try:
                from utils.excel import save_to_excel
                
                # Ensure project type is set
                if "project_type" not in st.session_state[SessionKeys.PROJECT_DATA]:
                    st.session_state[SessionKeys.PROJECT_DATA]["project_type"] = st.session_state[SessionKeys.PROJECT_TYPE]
                
                with st.spinner("Generating Excel cost sheet..."):
                    output_path = save_to_excel(st.session_state[SessionKeys.PROJECT_DATA])
                    
                    # Read the file for download
                    with open(output_path, "rb") as file:
                        excel_data = file.read()
                    
                    st.success("Cost sheet generated successfully!")
                    st.download_button(
                        label="📥 Download Cost Sheet",
                        data=excel_data,
                        file_name=os.path.basename(output_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error generating cost sheet: {str(e)}")
                st.error("Please ensure all required data is filled in correctly.")

if __name__ == "__main__":
    main() 