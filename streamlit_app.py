import streamlit as st
import pandas as pd
import numpy as np

# Page Configuration
st.set_page_config(
    page_title="Branch Sales & Inventory Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Branch Sales & Inventory Analytics")
st.markdown("---")

# ==========================================
# 🛠️ EXACT CELL/ROW DATA PROCESSING PIPELINE
# ==========================================
@st.cache_data(ttl=3600)
def process_uploaded_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    processed_data = {}
    
    # Generic format utility for matrices
    def format_summary_matrix(matrix_df):
        total_row = pd.DataFrame([{
            'Category': 'Total',
            'Total Units': matrix_df['Total Units'].sum(),
            'Sold Units': matrix_df['Sold Units'].sum(),
            'Balance Units': matrix_df['Balance Units'].sum(),
            'Value of Balance': matrix_df['Value of Balance'].sum()
        }]).set_index('Category')
        
        out_df = pd.concat([matrix_df, total_row])
        out_df['Sold (%)'] = (out_df['Sold Units'] / out_df['Total Units']).map('{:.2%}'.format)
        out_df['Balance (%)'] = (out_df['Balance Units'] / out_df['Total Units']).map('{:.2%}'.format)
        
        # Enforce string parsing object casting for commas
        for col in ['Total Units', 'Sold Units', 'Balance Units', 'Value of Balance']:
            out_df[col] = out_df[col].astype(object)
            
        for idx in out_df.index:
            out_df.at[idx, 'Total Units'] = f"{int(float(out_df.at[idx, 'Total Units'])):,}"
            out_df.at[idx, 'Sold Units'] = f"{int(float(out_df.at[idx, 'Sold Units'])):,}"
            out_df.at[idx, 'Balance Units'] = f"{int(float(out_df.at[idx, 'Balance Units'])):,}"
            out_df.at[idx, 'Value of Balance'] = f"$ {float(out_df.at[idx, 'Value of Balance']):,.2f}"
            
        return out_df[['Total Units', 'Sold Units', 'Sold (%)', 'Balance Units', 'Balance (%)', 'Value of Balance']]

    def extract_row_metrics(df, row_indices):
        """Safely extracts columns and computes specific value metrics for selected rows."""
        totals, solds, balances, values = 0, 0, 0, 0
        for idx in row_indices:
            if idx in df.index:
                t = pd.to_numeric(df.at[idx, 'TOTAL'], errors='coerce') or 0
                s = pd.to_numeric(df.at[idx, 'TOTAL SOLD'], errors='coerce') or 0
                b = pd.to_numeric(df.at[idx, 'BALANCE'], errors='coerce') or 0
                avg_p = pd.to_numeric(df.at[idx, 'AVG PO PRICE'], errors='coerce') or 0
                
                totals += t
                solds += s
                balances += b
                values += (b * avg_p)
        return totals, solds, balances, values

    # ==========================================
    # 1. CCK BRANCH PARSING
    # ==========================================
    if 'CCK-TABLET' in xls.sheet_names:
        # Load raw sheet without custom header tracking to lock 1-based index numbers matching Excel layout
        df_cck_tab_raw = pd.read_excel(xls, sheet_name='CCK-TABLET', header=None)
        
        # Assign columns from Row index 1
        cols = df_cck_tab_raw.iloc[1].str.strip().tolist()
        df_cck_tab_raw.columns = cols
        
        # Row slices map directly to Excel row numbers translated to 0-based Python indices
        blk_a_rows = list(range(2, 9))   # Excel rows 3 to 9 (includes all inner data components)
        blk_b_rows = list(range(10, 14)) # Excel rows 11 to 14
        blk_c_rows = list(range(15, 21)) # Excel rows 16 to 21
        
        tab_rows = []
        for blk_lbl, rows in [('Blk A', blk_a_rows), ('Blk B', blk_b_rows), ('Blk C', blk_c_rows)]:
            t, s, b, v = extract_row_metrics(df_cck_tab_raw, rows)
            tab_rows.append({
                'Category': f'Pedestal / Tablet ({blk_lbl})',
                'Total Units': t, 'Sold Units': s, 'Balance Units': b, 'Value of Balance': v
            })
        processed_data['CCK_Pedestal'] = format_summary_matrix(pd.DataFrame(tab_rows).set_index('Category'))

    if 'CCK-NICHE' in xls.sheet_names:
        df_cck_niche = pd.read_excel(xls, sheet_name='CCK-NICHE', header=1)
        df_cck_niche.columns = df_cck_niche.columns.str.strip()
        
        # Drop summary footer/subtotal rows 
        df_niche_clean = df_cck_niche[df_cck_niche['LOT TYPE'].notna() & (~df_cck_niche['BLOCK'].str.contains('TOTAL|TOTAL:', na=False, case=False))].copy()
        
        df_niche_clean['TOTAL_num'] = pd.to_numeric(df_niche_clean['TOTAL'], errors='coerce').fillna(0)
        df_niche_clean['SOLD_num'] = pd.to_numeric(df_niche_clean['TOTAL SOLD'], errors='coerce').fillna(0)
        df_niche_clean['BALANCE_num'] = pd.to_numeric(df_niche_clean['BALANCE'], errors='coerce').fillna(0)
        df_niche_clean['AVG_PO_num'] = pd.to_numeric(df_niche_clean['AVG PO PRICE'], errors='coerce').fillna(0)
        df_niche_clean['Value'] = df_niche_clean['BALANCE_num'] * df_niche_clean['AVG_PO_num']
        
        def classify_niche_type(val):
            val = str(val).strip().upper()
            if 'BUDDHA' in val: return 'Niche - Buddha'
            elif 'SINGLE' in val or 'SG' in val: return 'Niche - Single'
            elif 'DOUBLE' in val or 'DB' in val: return 'Niche - Double'
            else: return 'Niche - Others'
            
        df_niche_clean['Custom_Cat'] = df_niche_clean['LOT TYPE'].apply(classify_niche_type)
        niche_aggs = df_niche_clean.groupby('Custom_Cat').agg(
            Total_Units=('TOTAL_num', 'sum'), Sold_Units=('SOLD_num', 'sum'),
            Balance_Units=('BALANCE_num', 'sum'), Value_of_Balance=('Value', 'sum')
        ).rename(columns={'Total_Units': 'Total Units', 'Sold_Units': 'Sold Units', 'Balance_Units': 'Balance Units', 'Value_of_Balance': 'Value of Balance'})
        
        niche_aggs = niche_aggs.reindex(['Niche - Single', 'Niche - Double', 'Niche - Buddha', 'Niche - Others']).fillna(0)
        processed_data['CCK_Niche'] = format_summary_matrix(niche_aggs)

    # ==========================================
    # 2. LST & TLT BRANCHES PARSING (With Dynasty/Imperial isolation)
    # ==========================================
    for sheet_name, save_key in [('LST-TABLE & NICHE', 'LST'), ('TLT-TABLE & NICHE', 'TLT')]:
        if sheet_name in xls.sheet_names:
            df_branch = pd.read_excel(xls, sheet_name=sheet_name, header=1)
            df_branch.columns = df_branch.columns.str.strip()
            
            # Forward fill product context since rows under primary product header are left blank
            df_branch['PRODUCT_filled'] = df_branch['PRODUCT'].ffill().str.strip().upper()
            
            df_branch['TOTAL_num'] = pd.to_numeric(df_branch['TOTAL'], errors='coerce').fillna(0)
            df_branch['SOLD_num'] = pd.to_numeric(df_branch['TOTAL SOLD'], errors='coerce').fillna(0)
            df_branch['BALANCE_num'] = pd.to_numeric(df_branch['BALANCE'], errors='coerce').fillna(0)
            df_branch['AVG_PO_num'] = pd.to_numeric(df_branch['AVG PO PRICE'], errors='coerce').fillna(0)
            df_branch['Value'] = df_branch['BALANCE_num'] * df_branch['AVG_PO_num']
            
            # Custom classification router tracking Suite names matching user requests
            def classify_lst_tlt_rows(row):
                prod = str(row['PRODUCT_filled']).upper()
                suite = str(row['SUITE NO.']).strip().upper()
                lot = str(row['LOT TYPE']).strip().upper()
                
                if 'TOTAL' in prod:
                    return None
                    
                if 'TABLET' in prod:
                    if 'DYNASTY 2' in suite: return 'Pedestal - Dynasty 2'
                    elif 'IMPERIAL 2' in suite: return 'Pedestal - Imperial 2'
                    else: return 'Pedestal - Others'
                elif 'NICHE' in prod:
                    if 'SINGLE' in lot or 'SG' in lot: return 'Niche - Single'
                    if 'DOUBLE' in lot or 'DB' in lot: return 'Niche - Double'
                return None
                
            df_branch['Custom_Cat'] = df_branch.apply(classify_lst_tlt_rows, axis=1)
            df_branch_clean = df_branch[df_branch['Custom_Cat'].notna()]
            
            branch_aggs = df_branch_clean.groupby('Custom_Cat').agg(
                Total_Units=('TOTAL_num', 'sum'), Sold_Units=('SOLD_num', 'sum'),
                Balance_Units=('BALANCE_num', 'sum'), Value_of_Balance=('Value', 'sum')
            ).rename(columns={'Total_Units': 'Total Units', 'Sold_Units': 'Sold Units', 'Balance_Units': 'Balance Units', 'Value_of_Balance': 'Value of Balance'})
            
            expected_index = ['Pedestal - Dynasty 2', 'Pedestal - Imperial 2', 'Pedestal - Others', 'Niche - Single', 'Niche - Double']
            branch_aggs = branch_aggs.reindex(expected_index).fillna(0)
            processed_data[save_key] = format_summary_matrix(branch_aggs)
            
    return processed_data

# ==========================================
# 📥 USER INTERFACE
# ==========================================
uploaded_file = st.sidebar.file_uploader(
    "Upload Master Spreadsheet", 
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    with st.spinner("Recompiling structural layout matrices..."):
        data_package = process_uploaded_excel(uploaded_file)
    st.sidebar.success("🎉 Custom Matrix Compiled Successfully!")
    
    branch_choice = st.sidebar.radio("Select Branch Dashboard:", ["CCK Branch", "LST Branch", "TLT Branch"])
    
    if branch_choice == "CCK Branch":
        st.header("🏢 CCK Branch Segment Analysis")
        if 'CCK_Pedestal' in data_package:
            st.subheader("📦 Pedestal / Tablet Matrix (by explicit Blocks)")
            st.dataframe(data_package['CCK_Pedestal'], use_container_width=True)
        if 'CCK_Niche' in data_package:
            st.subheader("🏺 Niche Classification Analysis")
            st.dataframe(data_package['CCK_Niche'], use_container_width=True)
            
    elif branch_choice == "LST Branch" and 'LST' in data_package:
        st.header("🏢 LST Branch Matrix Analysis")
        st.dataframe(data_package['LST'], use_container_width=True)
        
    elif branch_choice == "TLT Branch" and 'TLT' in data_package:
        st.header("🏢 TLT Branch Matrix Analysis")
        st.dataframe(data_package['TLT'], use_container_width=True)
else:
    st.info("💡 Upload the tracking file to extract explicit cell block counts.")
