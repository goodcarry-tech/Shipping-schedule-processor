import streamlit as st
import sys
import os
from io import BytesIO
import pandas as pd
from datetime import datetime

# Import processing functions
from schedule_processor import parse_cosco_pdf, parse_one_pdf, parse_sitc_excel, create_excel_file

# Page configuration
st.set_page_config(
    page_title="Shipping Schedule Organizer",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4788;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.markdown('<div class="main-header">🚢 Shipping Schedule Organizer</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Organize multiple carrier schedules and export to Excel</div>', unsafe_allow_html=True)

# Helper function to filter by month
def filter_by_month(df, start_month, end_month):
    """Filter dataframe by month range"""
    if df.empty:
        return df
    
    # Month name to number mapping
    month_map = {
        "All": 0,
        "January": 1, "February": 2, "March": 3, "April": 4,
        "May": 5, "June": 6, "July": 7, "August": 8,
        "September": 9, "October": 10, "November": 11, "December": 12
    }
    
    start_num = month_map.get(start_month, 0)
    end_num = month_map.get(end_month, 0)
    
    if start_num == 0 and end_num == 0:
        return df
    
    # Extract month from ETD (format: MM-DD)
    df_copy = df.copy()
    df_copy['month'] = df_copy['ETD'].str.split('-').str[0].astype(int)
    
    if start_num > 0 and end_num > 0:
        # Filter between start and end month
        if start_num <= end_num:
            df_copy = df_copy[(df_copy['month'] >= start_num) & (df_copy['month'] <= end_num)]
        else:
            # Handle cross-year case (e.g., Dec to Feb)
            df_copy = df_copy[(df_copy['month'] >= start_num) | (df_copy['month'] <= end_num)]
    elif start_num > 0:
        # Only start month specified
        df_copy = df_copy[df_copy['month'] >= start_num]
    elif end_num > 0:
        # Only end month specified
        df_copy = df_copy[df_copy['month'] <= end_num]
    
    return df_copy.drop('month', axis=1)

# Sidebar
with st.sidebar:
    st.header("📋 Instructions")
    st.markdown("""
    ### How to use:
    1. **Upload schedules** - PDF/Excel formats
    2. **Select carrier** - Choose shipping line
    3. **Preview data** - Check results
    4. **Export Excel** - Download file
    
    ### Supported Carriers:
    - ✅ COSCO (PDF)
    - ✅ ONE (PDF)
    - ✅ SITC (Excel)
    - 🔜 More coming...
    
    ### Supported Formats:
    - 📄 PDF
    - 📊 Excel (.xlsx, .xls)
    """)
    
    st.divider()
    
    st.header("⚙️ Settings")
    
    st.markdown("**Display Options**")
    date_format = st.selectbox(
        "Date Format",
        ["MM-DD", "YYYY-MM-DD", "DD/MM"]
    )
    
    st.markdown("**Export Options**")
    remove_duplicates = st.checkbox(
        "Remove duplicates",
        value=True
    )
    
    include_timestamp = st.checkbox(
        "Add timestamp to filename",
        value=True
    )

# Main tabs
tab1, tab2, tab3 = st.tabs(["📤 Upload & Process", "📊 Data Preview", "📥 Export"])

# Tab 1: Upload
with tab1:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### Step 1: Upload Files")
        uploaded_files = st.file_uploader(
            "Multiple files supported",
            type=["pdf", "xlsx", "xls"],
            accept_multiple_files=True
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
            for file in uploaded_files:
                st.write(f"📄 {file.name} ({file.size / 1024:.1f} KB)")
    
    with col2:
        st.markdown("### Step 2: Select Carrier")
        
        carrier_mapping = {}
        if uploaded_files:
            for file in uploaded_files:
                carrier = st.selectbox(
                    f"File: {file.name[:30]}...",
                    ["Auto-detect", "COSCO", "ONE", "SITC"],
                    key=f"carrier_{file.name}"
                )
                carrier_mapping[file.name] = carrier
    
    st.divider()
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if uploaded_files:
            if st.button("🚀 Start Processing", type="primary", use_container_width=True):
                with st.spinner("Processing schedules..."):
                    try:
                        all_schedules = []
                        
                        # Process each file
                        for file in uploaded_files:
                            carrier = carrier_mapping.get(file.name, "Auto-detect")
                            file_content = BytesIO(file.read())
                            
                            schedules = []
                            
                            # Determine carrier and file type
                            if carrier == "COSCO" or "COSCO" in file.name.upper():
                                schedules = parse_cosco_pdf(file_content)
                            elif carrier == "ONE" or "ONE" in file.name.upper():
                                schedules = parse_one_pdf(file_content)
                            elif carrier == "SITC" or "SITC" in file.name.upper():
                                schedules = parse_sitc_excel(file_content)
                            else:
                                # Auto-detect
                                if file.name.lower().endswith('.pdf'):
                                    try:
                                        schedules = parse_cosco_pdf(file_content)
                                    except:
                                        file_content.seek(0)
                                        schedules = parse_one_pdf(file_content)
                                elif file.name.lower().endswith(('.xlsx', '.xls')):
                                    schedules = parse_sitc_excel(file_content)
                            
                            all_schedules.extend(schedules)
                        
                        # Create DataFrame
                        df = pd.DataFrame(all_schedules)
                        
                        if not df.empty:
                            # Remove duplicates if enabled
                            if remove_duplicates:
                                df = df.drop_duplicates()
                            
                            # Sort by ETD
                            df = df.sort_values('ETD').reset_index(drop=True)
                            
                            # Store in session state
                            st.session_state['processed'] = True
                            st.session_state['df'] = df
                            st.session_state['files'] = uploaded_files
                            st.session_state['carrier_mapping'] = carrier_mapping
                            
                            st.success(f"✅ Processing complete! Found {len(df)} schedules. Switch to 'Data Preview' tab")
                            st.balloons()
                        else:
                            st.error("❌ No schedules found in uploaded files. Please check file format.")
                    
                    except Exception as e:
                        st.error(f"❌ Error processing files: {str(e)}")
                        st.error("Please check if files are in correct format.")

# Tab 2: Preview
with tab2:
    st.markdown("### 📊 Schedule Preview")
    
    if 'processed' in st.session_state and st.session_state['processed']:
        df = st.session_state['df']
        
        st.info("💡 Confirm data accuracy, then go to 'Export' tab to download")
        
        # Statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Schedules", len(df))
        with col2:
            st.metric("Carriers", df['CARRIER'].nunique())
        with col3:
            if not df.empty:
                st.metric("Date Range", f"{df['ETD'].min()}~{df['ETD'].max()}")
        with col4:
            ts_count = len(df[df['T/S Port'] != ''])
            st.metric("T/S Ports", ts_count)
        
        st.divider()
        
        # Filters
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            carriers = ["All"] + sorted(df['CARRIER'].unique().tolist())
            filter_carrier = st.multiselect(
                "Filter Carrier",
                carriers,
                default=["All"]
            )
        with col2:
            services = ["All"] + sorted(df['Service'].unique().tolist())
            filter_service = st.multiselect(
                "Filter Service",
                services,
                default=["All"]
            )
        with col3:
            start_month = st.selectbox(
                "Start Month",
                ["All", "January", "February", "March", "April", "May", "June", 
                 "July", "August", "September", "October", "November", "December"],
                index=0,
                help="Show schedules from this month onwards"
            )
        with col4:
            end_month = st.selectbox(
                "End Month",
                ["All", "January", "February", "March", "April", "May", "June", 
                 "July", "August", "September", "October", "November", "December"],
                index=0,
                help="Show schedules up to this month"
            )
        
        # Apply filters
        df_filtered = df.copy()
        
        if "All" not in filter_carrier:
            df_filtered = df_filtered[df_filtered['CARRIER'].isin(filter_carrier)]
        
        if "All" not in filter_service:
            df_filtered = df_filtered[df_filtered['Service'].isin(filter_service)]
        
        # Apply month filter
        df_filtered = filter_by_month(df_filtered, start_month, end_month)
        
        # Show filter status
        if start_month != "All" or end_month != "All":
            filter_msg = "📅 Active filters: "
            if start_month != "All":
                filter_msg += f"From {start_month}"
            if end_month != "All":
                if start_month != "All":
                    filter_msg += f" to {end_month}"
                else:
                    filter_msg += f"Up to {end_month}"
            st.info(filter_msg)
        
        st.markdown("#### Schedule Details")
        
        # Display full dataframe
        st.dataframe(
            df_filtered,
            use_container_width=True,
            height=600,
            hide_index=True
        )
        
        st.markdown(f"**Showing {len(df_filtered)} of {len(df)} schedules**")
        
        # Download CSV
        csv = df_filtered.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📄 Download CSV Preview",
            data=csv,
            file_name="preview.csv",
            mime="text/csv"
        )
        
    else:
        st.info("👈 Upload and process files in 'Upload & Process' tab first")

# Tab 3: Export
with tab3:
    st.markdown("### 📥 Export Schedule")
    
    if 'processed' in st.session_state and st.session_state['processed']:
        df = st.session_state['df']
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.markdown("#### Export Options")
            
            export_format = st.radio(
                "Format",
                ["Excel (.xlsx)", "CSV (.csv)"]
            )
            
            file_name = st.text_input(
                "Filename",
                value="shipping-schedule"
            )
            
            if export_format == "Excel (.xlsx)":
                include_summary = st.checkbox(
                    "Include summary sheet",
                    value=True
                )
            else:
                include_summary = False
            
            st.divider()
            
            st.markdown("#### Date Range Filter")
            
            export_start_month = st.selectbox(
                "Export from month",
                ["All", "January", "February", "March", "April", "May", "June", 
                 "July", "August", "September", "October", "November", "December"],
                index=0,
                key="export_start",
                help="Only export schedules from this month"
            )
            
            export_end_month = st.selectbox(
                "Export to month",
                ["All", "January", "February", "March", "April", "May", "June", 
                 "July", "August", "September", "October", "November", "December"],
                index=0,
                key="export_end",
                help="Only export schedules up to this month"
            )
            
        with col2:
            st.markdown("#### Export Preview")
            
            # Apply month filter for preview
            df_export = filter_by_month(df, export_start_month, export_end_month)
            
            preview_msg = f"""
            **Ready to export:**
            - 📊 Records: {len(df_export)}
            - 🚢 Carriers: {', '.join([f"{carrier} ({count})" for carrier, count in df_export['CARRIER'].value_counts().items()])}
            - 📅 Date range: {df_export['ETD'].min()} ~ {df_export['ETD'].max()}
            - ✅ Sorted by ETD
            """
            
            if export_start_month != "All" or export_end_month != "All":
                filter_info = "\n\n**📅 Date filter active:**\n"
                if export_start_month != "All":
                    filter_info += f"- From: {export_start_month}\n"
                if export_end_month != "All":
                    filter_info += f"- To: {export_end_month}\n"
                preview_msg += filter_info
            
            st.info(preview_msg)
            
            if export_format == "Excel (.xlsx)" and include_summary:
                st.success("✨ Will include summary sheet")
        
        st.divider()
        
        # Export button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("📥 Export Now", type="primary", use_container_width=True):
                with st.spinner("Generating file..."):
                    try:
                        # Apply month filter
                        df_export = filter_by_month(df, export_start_month, export_end_month)
                        
                        if df_export.empty:
                            st.error("❌ No data to export with current filters")
                        else:
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            
                            if export_format == "Excel (.xlsx)":
                                # Generate Excel
                                excel_file = create_excel_file(df_export, include_summary=include_summary)
                                
                                if include_timestamp:
                                    filename = f"{file_name}_{timestamp}.xlsx"
                                else:
                                    filename = f"{file_name}.xlsx"
                                
                                st.success("✅ Excel file generated successfully!")
                                st.download_button(
                                    label=f"💾 Download {filename}",
                                    data=excel_file,
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary"
                                )
                            
                            else:  # CSV
                                csv_data = df_export.to_csv(index=False).encode('utf-8-sig')
                                
                                if include_timestamp:
                                    filename = f"{file_name}_{timestamp}.csv"
                                else:
                                    filename = f"{file_name}.csv"
                                
                                st.success("✅ CSV file generated successfully!")
                                st.download_button(
                                    label=f"💾 Download {filename}",
                                    data=csv_data,
                                    file_name=filename,
                                    mime="text/csv",
                                    type="primary"
                                )
                    
                    except Exception as e:
                        st.error(f"❌ Error generating file: {str(e)}")
    else:
        st.warning("⚠️ Process data first")
        st.markdown("""
        ### 💡 Before exporting:
        1. Upload schedule files
        2. Select carriers
        3. Process data
        4. Confirm preview
        """)

# Footer
st.divider()
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("""
    <div style='text-align: center; color: #666; padding: 1rem;'>
        <p>🚢 Shipping Schedule Organizer v2.0</p>
        <p>Supports COSCO, ONE, SITC | Full month filtering</p>
    </div>
    """, unsafe_allow_html=True)
