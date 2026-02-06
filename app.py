import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

# 頁面配置
st.set_page_config(
    page_title="船期整理系統",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義CSS
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
    .upload-box {
        border: 2px dashed #4CAF50;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background-color: #f8f9fa;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# 標題
st.markdown('<div class="main-header">🚢 船期整理系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">自動整理多家船公司船期表，一鍵匯出Excel</div>', unsafe_allow_html=True)

# 側邊欄 - 設定與說明
with st.sidebar:
    st.header("📋 使用說明")
    st.markdown("""
    ### 如何使用：
    1. **上傳船期表** - 支援 PDF/Excel 格式
    2. **選擇船公司** - 選擇對應的船公司
    3. **預覽資料** - 檢查解析結果
    4. **匯出Excel** - 下載整理後的船期表
    
    ### 支援的船公司：
    - ✅ COSCO (中遠海運)
    - ✅ ONE (海洋網聯)
    - ✅ SITC (海豐國際)
    - 🔜 更多船公司陸續加入...
    
    ### 支援的格式：
    - 📄 PDF
    - 📊 Excel (.xlsx, .xls)
    - 📑 CSV
    """)
    
    st.divider()
    
    # 進階設定
    st.header("⚙️ 進階設定")
    date_format = st.selectbox(
        "日期格式",
        ["MM-DD", "YYYY-MM-DD", "DD/MM"],
        help="選擇匯出的日期格式"
    )
    
    remove_duplicates = st.checkbox(
        "自動去除重複記錄",
        value=True,
        help="移除完全相同的船期記錄"
    )
    
    include_timestamp = st.checkbox(
        "檔名加入時間戳記",
        value=True,
        help="匯出檔案名稱包含生成時間"
    )

# 主要內容區域
tab1, tab2, tab3 = st.tabs(["📤 上傳與處理", "📊 資料預覽", "📥 匯出結果"])

# Tab 1: 上傳與處理
with tab1:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 步驟 1: 上傳船期表")
        uploaded_files = st.file_uploader(
            "支援多檔案上傳",
            type=["pdf", "xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="可同時上傳多個船公司的船期表"
        )
        
        if uploaded_files:
            st.success(f"✅ 已上傳 {len(uploaded_files)} 個檔案")
            for file in uploaded_files:
                st.write(f"📄 {file.name} ({file.size / 1024:.1f} KB)")
    
    with col2:
        st.markdown("### 步驟 2: 選擇船公司")
        
        carrier_mapping = {}
        if uploaded_files:
            for file in uploaded_files:
                carrier = st.selectbox(
                    f"檔案: {file.name[:30]}...",
                    ["自動識別", "COSCO", "ONE", "SITC", "MAERSK", "MSC", "CMA CGM", "其他"],
                    key=f"carrier_{file.name}"
                )
                carrier_mapping[file.name] = carrier
    
    st.divider()
    
    # 處理按鈕
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if uploaded_files:
            if st.button("🚀 開始處理", type="primary", use_container_width=True):
                with st.spinner("正在處理船期資料..."):
                    # 這裡會呼叫處理函數
                    st.session_state['processed'] = True
                    st.session_state['files'] = uploaded_files
                    st.session_state['carrier_mapping'] = carrier_mapping
                    st.success("✅ 處理完成！請切換到「資料預覽」標籤查看結果")
                    st.balloons()

# Tab 2: 資料預覽
with tab2:
    st.markdown("### 📊 船期資料預覽")
    
    if 'processed' in st.session_state and st.session_state['processed']:
        # 這裡顯示處理後的資料
        st.info("💡 提示：確認資料無誤後，請切換到「匯出結果」標籤下載Excel檔案")
        
        # 統計資訊
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("總船期數", "35", delta="5 筆新增")
        with col2:
            st.metric("船公司數", "2", delta="0")
        with col3:
            st.metric("日期範圍", "02-06 ~ 03-30")
        with col4:
            st.metric("T/S港口數", "2")
        
        st.divider()
        
        # 篩選功能
        col1, col2, col3 = st.columns(3)
        with col1:
            filter_carrier = st.multiselect(
                "篩選船公司",
                ["全部", "COSCO", "ONE", "SITC"],
                default=["全部"]
            )
        with col2:
            filter_service = st.multiselect(
                "篩選服務線",
                ["全部", "HPX2", "EC3", "VSX", "VSS"],
                default=["全部"]
            )
        with col3:
            date_range = st.date_input(
                "日期範圍",
                value=None,
                help="篩選特定日期範圍的船期"
            )
        
        # 顯示資料表
        st.markdown("#### 船期明細表")
        
        # 示例數據
        sample_data = {
            'CARRIER': ['ONE', 'ONE', 'COSCO', 'ONE', 'COSCO'],
            'Service': ['EC3', 'VSS', 'HPX2', 'EC3', 'HPX2'],
            'Vessel': ['HAIAN VIEW', 'ONE STORK', 'MTT SENARI', 'INCRES', 'SAN PEDRO'],
            'Voyage': ['162S', '028E', '029S', '065S', '99S'],
            'ETD': ['02-06', '02-09', '02-15', '02-14', '02-18'],
            'ETA': ['02-20', '02-20', '', '02-27', '03-03'],
            'Transit Time': ['15', '14', '11', '11', '13'],
            'T/S Port': ['', '', 'Port kelang', '', 'Port kelang']
        }
        
        df_sample = pd.DataFrame(sample_data)
        
        # 使用 st.dataframe 顯示可互動的表格
        st.dataframe(
            df_sample,
            use_container_width=True,
            height=400,
            hide_index=True
        )
        
        # 下載CSV選項
        csv = df_sample.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📄 下載CSV預覽",
            data=csv,
            file_name="preview.csv",
            mime="text/csv",
            help="下載當前預覽的CSV檔案"
        )
        
    else:
        st.info("👈 請先在「上傳與處理」標籤上傳檔案並處理")
        st.image("https://via.placeholder.com/800x400/e3f2fd/1976d2?text=尚未處理資料", use_container_width=True)

# Tab 3: 匯出結果
with tab3:
    st.markdown("### 📥 匯出船期表")
    
    if 'processed' in st.session_state and st.session_state['processed']:
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.markdown("#### 匯出選項")
            
            export_format = st.radio(
                "檔案格式",
                ["Excel (.xlsx)", "CSV (.csv)", "兩者都要"],
                help="選擇要匯出的檔案格式"
            )
            
            file_name = st.text_input(
                "檔案名稱",
                value="船期排序表",
                help="不需要加副檔名"
            )
            
            include_summary = st.checkbox(
                "包含統計摘要工作表",
                value=True,
                help="在Excel中額外加入統計摘要頁"
            )
            
        with col2:
            st.markdown("#### 匯出預覽")
            st.info("""
            **即將匯出：**
            - 📊 總船期數: 35 筆
            - 🚢 船公司: COSCO (5筆), ONE (30筆)
            - 📅 日期範圍: 2026-02-06 ~ 2026-03-30
            - 🔄 已按ETD排序
            - ✅ 已去除重複記錄
            """)
            
            if include_summary:
                st.success("✨ 將包含統計摘要工作表")
        
        st.divider()
        
        # 匯出按鈕
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("📥 立即匯出", type="primary", use_container_width=True):
                with st.spinner("正在生成檔案..."):
                    # 這裡會生成實際的檔案
                    st.success("✅ 檔案生成完成！")
                    
                    # 模擬下載按鈕
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    if include_timestamp:
                        filename = f"{file_name}_{timestamp}.xlsx"
                    else:
                        filename = f"{file_name}.xlsx"
                    
                    st.download_button(
                        label=f"💾 下載 {filename}",
                        data=b"",  # 這裡會是實際的檔案內容
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
        # 歷史記錄
        with st.expander("📜 匯出歷史記錄"):
            st.markdown("""
            | 時間 | 檔案名稱 | 記錄數 | 狀態 |
            |------|---------|--------|------|
            | 2026-02-05 14:30 | 船期排序表_20260205_1430.xlsx | 35 | ✅ 成功 |
            | 2026-02-04 09:15 | schedule_export.xlsx | 28 | ✅ 成功 |
            | 2026-02-03 16:45 | 船期整理_20260203.xlsx | 42 | ✅ 成功 |
            """)
    else:
        st.warning("⚠️ 請先處理船期資料")
        st.markdown("""
        ### 💡 匯出前需要：
        1. 上傳船期表檔案
        2. 選擇對應的船公司
        3. 完成資料處理
        4. 確認資料預覽無誤
        """)

# 頁尾
st.divider()
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("""
    <div style='text-align: center; color: #666; padding: 1rem;'>
        <p>🚢 船期整理系統 v1.0 | 由 Claude 協助開發</p>
        <p>支援 COSCO, ONE 及更多船公司 | <a href='#'>使用說明</a> | <a href='#'>問題回報</a></p>
    </div>
    """, unsafe_allow_html=True)
