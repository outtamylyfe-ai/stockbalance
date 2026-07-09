import streamlit as st
import pandas as pd
import numpy as np

# Page Configuration
st.set_page_config(
    page_title="Dynamic Branch Sales & Inventory Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dynamic Branch Sales & Inventory Analytics")
st.markdown("Upload your structural reporting Excel spreadsheet to automatically parse, process, and populate your metrics.")
st.markdown("---")

# ==========================================
# 🛠️ AUTOMATED DATA PROCESSING FUNCTION
# ==========================================
@st.cache_data(ttl=3600)  # Cache for performance, expires in 1 hour
def process_uploaded_excel(uploaded_file):
    """
    Parses an uploaded Excel business spreadsheet dynamically, 
    extracting key reporting views across different branch sheets.
    """
    xls = pd.ExcelFile(uploaded_file)
    processed_data = {}
    
    # 1. Process CCK-NICHE View
    if 'CCK-NICHE' in xls.sheet_names:
        df_niche = pd.read_excel(xls, sheet_name='CCK-NICHE', header=1)
        df_niche.columns = df_niche.columns.str.strip()
        
        # Safely convert calculation dependencies into numeric datatypes
        df_niche['BALANCE_num'] = pd.to_numeric(df_niche['BALANCE'], errors='coerce').fillna(0)
        df_niche['AVG_PO_num'] = pd.to_numeric(df_niche['AVG PO PRICE'], errors='coerce').fillna(0)
        
        # 📊 Dynamic Custom Math: Value of Balance = Balance Units * Average PO Price
        df_niche['Calculated_Value'] = df_niche['BALANCE_num'] * df_niche['AVG_PO_num']
        
        # For block metrics, we aggregate individual lot rows up to the correct block totals 
        # to ensure any custom math logic applies smoothly down the hierarchy.
        df_niche['Block_Group'] = df_niche['BLOCK'].ffill()
        df_clean_rows = df_niche[
            df_niche['LOT TYPE'].notna() & 
            (~df_niche['BLOCK'].str.contains('TOTAL|TOTAL:', na=False, case=False))
        ].copy()
        
        # Block Summary Segment
        block_summary_data = []
        for raw_blk, display_name in [('A', 'Blk A'), ('B', 'Blk B'), ('C', 'Blk C')]:
            blk_data = df_clean_rows[df_clean_rows['Block_Group'] == raw_blk]
            if not blk_data.empty:
                tot = pd.to_numeric(blk_data['TOTAL'], errors='coerce').sum()
                sold = pd.to_numeric(blk_data['TOTAL SOLD'], errors='coerce').sum()
                bal = blk_data['BALANCE_num'].sum()
                val = blk_data['Calculated_Value'].sum()
                
                block_summary_data.append({
                    'Block Unit Name': display_name,
                    'TOTAL': tot, 'SOLD': sold, 'BALANCE': bal, 'VALUE': val
                })
                
        df_blocks_calc = pd.DataFrame(block_summary_data)
        
        summary_blocks = pd.DataFrame({
            'Block Unit Name': df_blocks_calc['Block Unit Name'],
            'Sold (%)': (df_blocks_calc['SOLD'] / df_blocks_calc['TOTAL']).map('{:.2%}'.format),
            'Balance (%)': (df_blocks_calc['BALANCE'] / df_blocks_calc['TOTAL']).map('{:.2%}'.format),
            'Balance Units': df_blocks_calc['BALANCE'].astype(int),
            'Value of balance': df_blocks_calc['VALUE'].map('$  {:,.2f}'.format)
        })
        
        # Appending Global Totals
        total_row = pd.DataFrame([{
            'Block Unit Name': 'Total',
            'Sold (%)': '{:.2%}'.format(df_blocks_calc['SOLD'].sum() / df_blocks_calc['TOTAL'].sum()),
            'Balance (%)': '{:.2%}'.format(df_blocks_calc['BALANCE'].sum() / df_blocks_calc['TOTAL'].sum()),
            'Balance Units': int(df_blocks_calc['BALANCE'].sum()),
            'Value of balance': '$  {:,.2f}'.format(df_blocks_calc['VALUE'].sum())
        }])
        processed_data['cck_niche_blocks'] = pd.concat([summary_blocks, total_row], ignore_index=True)
        
        # Matrix Form Breakdown
        def categorize_lot_type(row):
            lot = str(row['LOT TYPE']).strip().upper()
            if 'SINGLE' in lot or 'SG' in lot: return 'Single'
            elif 'DOUBLE' in lot or 'DB' in lot: return 'Double'
            elif 'FAMILY' in lot: return 'Family'
            elif 'BUDDHA' in lot: return 'Buddha'
            elif 'TOWER' in lot: return 'Tower'
            else: return 'Special'
            
        df_clean_rows['Category'] = df_clean_rows.apply(categorize_lot_type, axis=1)
        
        matrix_aggs = df_clean_rows.groupby('Category').agg(
            Total=('TOTAL', lambda x: pd.to_numeric(x, errors='coerce').sum()), 
            Balance=('BALANCE_num', 'sum'),
            Sold=('TOTAL SOLD', lambda x: pd.to_numeric(x, errors='coerce').sum()), 
            Value_of_B=('Calculated_Value', 'sum')
        ).reindex(['Single', 'Double', 'Family', 'Buddha', 'Tower', 'Special']).fillna(0)
        
        matrix_final = matrix_aggs.T
        matrix_final.loc['Balance(%)'] = (matrix_final.loc['Balance'] / matrix_final.loc['Total']).map('{:.2%}'.format)
        matrix_final.loc['Sold(%)'] = (matrix_final.loc['Sold'] / matrix_final.loc['Total']).map('{:.2%}'.format)
        
        for col in matrix_final.columns:
            matrix_final.at['Total', col] = f"{int(matrix_final.at['Total', col]):,}"
            matrix_final.at['Balance', col] = f"{int(matrix_final.at['Balance', col]):,}"
            matrix_final.at['Sold', col] = f"{int(matrix_final.at['Sold', col]):,}"
            matrix_final.at['Value of b', col] = f"$ {matrix_final.at['Value_of_B', col]:,.2f}"
            
        processed_data['cck_niche_matrix'] = matrix_final.drop(index='Value_of_B')

    # 2. Process LST Inventory View
    if 'LST-TABLE & NICHE' in xls.sheet_names:
        df_lst = pd.read_excel(xls, sheet_name='LST-TABLE & NICHE', header=1)
        df_lst.columns = df_lst.columns.str.strip()
        
        processed_data['lst_tablet'] = df_lst[df_lst['PRODUCT'] == 'TABLET'].groupby('SUITE NO.').agg(
            Total=('TOTAL', 'sum'), Balance=('BALANCE', 'sum'), Sold=('TOTAL SOLD', 'sum'),
            Value=('BALANCE $ \'000 (PO)', lambda x: x.sum() * 1000)
        )
        processed_data['lst_niche'] = df_lst[df_lst['PRODUCT'] == 'NICHE'].groupby('LOT TYPE').agg(
            Total=('TOTAL', 'sum'), Balance=('BALANCE', 'sum'), Sold=('TOTAL SOLD', 'sum'),
            Value=('BALANCE $ \'000 (PO)', lambda x: x.sum() * 1000)
        )

    # 3. Process TLT Inventory View
    if 'TLT-TABLE & NICHE' in xls.sheet_names:
        df_tlt = pd.read_excel(xls, sheet_name='TLT-TABLE & NICHE', header=1)
        df_tlt.columns = df_tlt.columns.str.strip()
        
        processed_data['tlt_tablet'] = df_tlt[df_tlt['PRODUCT'] == 'TABLET'].groupby('SUITE NO.').agg(
            Total=('TOTAL', 'sum'), Balance=('BALANCE', 'sum'), Sold=('TOTAL SOLD', 'sum'),
            Value=('BALANCE $ \'000 (PO)', lambda x: x.sum() * 1000)
        )
        processed_data['tlt_niche'] = df_tlt[df_tlt['PRODUCT'] == 'NICHE'].groupby('LOT TYPE').agg(
            Total=('TOTAL', 'sum'), Balance=('BALANCE', 'sum'), Sold=('TOTAL SOLD', 'sum'),
            Value=('BALANCE $ \'000 (PO)', lambda x: x.sum() * 1000)
        )
        
    return processed_data

# ==========================================
# 📥 USER INTERFACE: FILE UPLOADER
# ==========================================
uploaded_file = st.sidebar.file_uploader(
    "Step 1: Upload Reporting Document", 
    type=["xlsx", "xls"],
    help="Upload your master tracking spreadsheet to autopopulate the views."
)

if uploaded_file is not None:
    # Trigger dynamic extraction pipeline
    with st.spinner("Analyzing spreadsheet structure and calculating summaries..."):
        data_package = process_uploaded_excel(uploaded_file)
    st.sidebar.success("🎉 Data parsed successfully!")
    
    # Navigation Radio
    page = st.sidebar.radio("Step 2: Choose View:", [
        "CCK-NICHE Overview", 
        "LST Inventory Analytics", 
        "TLT Inventory Analytics"
    ])
    
    # Render selected views using processed packages
    if page == "CCK-NICHE Overview" and 'cck_niche_blocks' in data_package:
        st.header("CCK-NICHE Comprehensive Overview")
        col1, col2 = st.columns([1, 1])
        with col1:
            st.subheader("🏢 Niche Block Performance")
            st.dataframe(data_package['cck_niche_blocks'], use_container_width=True, hide_index=True)
        with col2:
            st.subheader("📐 Unit Matrix Breakdown")
            st.dataframe(data_package['cck_niche_matrix'], use_container_width=True)
            
    elif page == "LST Inventory Analytics" and 'lst_tablet' in data_package:
        st.header("LST Branch Inventory Matrix")
        st.subheader("📱 Tablet Matrix Breakdowns")
        st.dataframe(data_package['lst_tablet'], use_container_width=True)
        st.subheader("🏺 Niche Lot Form Factor Breakdowns")
        st.dataframe(data_package['lst_niche'], use_container_width=True)
        
    elif page == "TLT Inventory Analytics" and 'tlt_tablet' in data_package:
        st.header("TLT Branch Inventory Performance")
        st.subheader("📱 Tablet Overview Segment")
        st.dataframe(data_package['tlt_tablet'], use_container_width=True)
        st.subheader("🏺 Niche Form Factor Metrics")
        st.dataframe(data_package['tlt_niche'], use_container_width=True)
else:
    # Fallback view when no sheet is provided yet
    st.info("💡 Please upload an Excel document via the sidebar menu to instantly auto-populate the reporting dashboard dashboards.")
