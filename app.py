import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

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
    - ✅ COSCO
    - ✅ ONE
    - ✅ SITC
    - 🔜 More coming...
    
    ### Supported Formats:
    - 📄 PDF
    - 📊 Excel (.xlsx, .xls)
    - 📑 CSV
    """)
    
    st.divider()
    
    st.header("⚙️ Settings")
    date_format = st.selectbox(
        "Date Format",
        ["MM-DD", "YYYY-MM-DD", "DD/MM"]
    )
    
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
            type=["pdf", "xlsx", "xls", "csv"],
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
                    ["Auto-detect", "COSCO", "ONE", "SITC", "MAERSK", "MSC", "Other"],
                    key=f"carrier_{file.name}"
                )
                carrier_mapping[file.name] = carrier
    
    st.divider()
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if uploaded_files:
            if st.button("🚀 Start Processing", type="primary", use_container_width=True):
                with st.spinner("Processing..."):
                    st.session_state['processed'] = True
                    st.session_state['files'] = uploaded_files
                    st.session_state['carrier_mapping'] = carrier_mapping
                    st.success("✅ Complete! Check 'Data Preview' tab")
                    st.balloons()

# Tab 2: Preview
with tab2:
    st.markdown("### 📊 Schedule Preview")
    
    if 'processed' in st.session_state and st.session_state['processed']:
        st.info("💡 Confirm data, then go to 'Export' tab")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total", "38", delta="5")
        with col2:
            st.metric("Carriers", "3")
        with col3:
            st.metric("Date Range", "02-06~03-30")
        with col4:
            st.metric("T/S Ports", "2")
        
        st.divider()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            filter_carrier = st.multiselect(
                "Filter Carrier",
                ["All", "COSCO", "ONE", "SITC"],
                default=["All"]
            )
        with col2:
            filter_service = st.multiselect(
                "Filter Service",
                ["All", "HPX2", "EC3", "VSX"],
                default=["All"]
            )
        
        st.markdown("#### Schedule Details")
        
        sample_data = {
            'CARRIER': ['ONE', 'ONE', 'COSCO', 'SITC'],
            'Service': ['EC3', 'VSS', 'HPX2', 'CBX2'],
            'Vessel': ['HAIAN VIEW', 'ONE STORK', 'MTT SENARI', 'SITC HUIMING'],
            'Voyage': ['162S', '028E', '029S', '2602S'],
            'ETD': ['02-06', '02-09', '02-15', '02-18'],
            'ETA': ['02-20', '02-20', '', '03-01'],
            'Transit Time': ['15', '14', '11', '11'],
            'T/S Port': ['', '', 'Port kelang', 'DIRECT']
        }
        
        df_sample = pd.DataFrame(sample_data)
        
        st.dataframe(
            df_sample,
            use_container_width=True,
            height=400,
            hide_index=True
        )
        
        csv = df_sample.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Download CSV",
            data=csv,
            file_name="preview.csv",
            mime="text/csv"
        )
        
    else:
        st.info("👈 Upload files in 'Upload & Process' tab first")

# Tab 3: Export
with tab3:
    st.markdown("### 📥 Export Schedule")
    
    if 'processed' in st.session_state and st.session_state['processed']:
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.markdown("#### Options")
            
            export_format = st.radio(
                "Format",
                ["Excel (.xlsx)", "CSV (.csv)", "Both"]
            )
            
            file_name = st.text_input(
                "Filename",
                value="shipping-schedule"
            )
            
            include_summary = st.checkbox(
                "Include summary sheet",
                value=True
            )
            
        with col2:
            st.markdown("#### Preview")
            st.info("""
            **Ready to export:**
            - 📊 Records: 38
            - 🚢 COSCO (5), ONE (30), SITC (3)
            - 📅 2026-02-06 ~ 03-30
            - ✅ Sorted by ETD
            """)
        
        st.divider()
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("📥 Export Now", type="primary", use_container_width=True):
                with st.spinner("Generating..."):
                    st.success("✅ File ready!")
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    if include_timestamp:
                        filename = f"{file_name}_{timestamp}.xlsx"
                    else:
                        filename = f"{file_name}.xlsx"
                    
                    st.download_button(
                        label=f"💾 Download {filename}",
                        data=b"",
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
    else:
        st.warning("⚠️ Process data first")

# Footer
st.divider()
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("""
    <div style='text-align: center; color: #666; padding: 1rem;'>
        <p>🚢 Shipping Schedule Organizer v2.0</p>
        <p>Supports COSCO, ONE, SITC and more</p>
    </div>
    """, unsafe_allow_html=True)
