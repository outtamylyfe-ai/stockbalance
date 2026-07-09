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
# 🛠️ DATA PROCESSING PIPELINE
# ==========================================
@st.cache_data(ttl=3600)
def process_uploaded_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    processed_data = {}
    
    # --- HELPER UTILITIES ---
    def clean_dataframe(df):
        df.columns = df.columns.str.strip()
        df['BALANCE_num'] = pd.to_numeric(df['BALANCE'], errors='coerce').fillna(0)
        df['AVG_PO_num'] = pd.to_numeric(df['AVG PO PRICE'], errors='coerce').fillna(0)
        df['TOTAL_num'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
        df['SOLD_num'] = pd.to_numeric(df['TOTAL SOLD'], errors='coerce').fillna(0)
        df['Calculated_Value'] = df['BALANCE_num'] * df['AVG_PO_num']
        return df

    def format_summary_matrix(matrix_df):
        """Converts raw rows of metrics into structural dashboard tables with proper totals/percentages."""
        # Calculate Total Row dynamically
        total_row = pd.DataFrame([{
            'Category': 'Total',
            'Total Units': matrix_df['Total Units'].sum(),
            'Sold Units': matrix_df['Sold Units'].sum(),
            'Balance Units': matrix_df['Balance Units'].sum(),
            'Value of Balance': matrix_df['Value of Balance'].sum()
        }]).set_index('Category')
        
        out_df = pd.concat([matrix_df, total_row])
        
        # Add Percentage Columns safely
        out_df['Sold (%)'] = (out_df['Sold Units'] / out_df['Total Units']).map('{:.2%}'.format)
        out_df['Balance (%)'] = (out_df['Balance Units'] / out_df['Total Units']).map('{:.2%}'.format)
        
        # String Formatter
        for idx in out_df.index:
            out_df.at[idx, 'Total Units'] = f"{int(out_df.at[idx, 'Total Units']):,}"
            out_df.at[idx, 'Sold Units'] = f"{int(out_df.at[idx, 'Sold Units']):,}"
            out_df.at[idx, 'Balance Units'] = f"{int(out_df.at[idx, 'Balance Units']):,}"
            out_df.at[idx, 'Value of Balance'] = f"$ {out_df.at[idx, 'Value of Balance']:,.2f}"
            
        return out_df[['Total Units', 'Sold Units', 'Sold (%)', 'Balance Units', 'Balance (%)', 'Value of Balance']]

    # ==========================================
    # 1. CCK BRANCH PARSING
    # ==========================================
    # A. CCK Pedestal / Tablet (Grouped by Blk A, B, C)
    if 'CCK-TABLET' in xls.sheet_names:
        df_cck_tab = pd.read_excel(xls, sheet_name='CCK-TABLET', header=1)
        df_cck_tab = clean_dataframe(df_cck_tab)
        df_cck_tab['Block_Group'] = df_cck_tab['BLOCK'].ffill().str.strip()
        
        # Clean rows
        df_tab_clean = df_cck_tab[df_cck_tab['SUITE NO.'].notna() & (~df_cck_tab['BLOCK'].str.contains('TOTAL|TOTAL:', na=False, case=False))]
        
        tab_rows = []
        for blk in ['A', 'B', 'C']:
            sub = df_tab_clean[df_tab_clean['Block_Group'] == blk]
            tab_rows.append({
                'Category': f'Pedestal / Tablet (Blk {blk})',
                'Total Units': sub['TOTAL_num'].sum(),
                'Sold Units': sub['SOLD_num'].sum(),
                'Balance Units': sub['BALANCE_num'].sum(),
                'Value of Balance': sub['Calculated_Value'].sum()
            })
        processed_data['CCK_Pedestal'] = format_summary_matrix(pd.DataFrame(tab_rows).set_index('Category'))

    # B. CCK Niche (Grouped by Single, Double, Buddha, Others)
    if 'CCK-NICHE' in xls.sheet_names:
        df_cck_niche = pd.read_excel(xls, sheet_name='CCK-NICHE', header=1)
        df_cck_niche = clean_dataframe(df_cck_niche)
        df_niche_clean = df_cck_niche[df_cck_niche['LOT TYPE'].notna() & (~df_cck_niche['BLOCK'].str.contains('TOTAL|TOTAL:', na=False, case=False))]
        
        def categorize_cck_niche(row):
            lot = str(row['LOT TYPE']).strip().upper()
            if 'BUDDHA' in lot: return 'Niche - Buddha'
            elif 'SINGLE' in lot or 'SG' in lot: return 'Niche - Single'
            elif 'DOUBLE' in lot or 'DB' in lot: return 'Niche - Double'
            else: return 'Niche - Others'
            
        df_niche_clean = df_niche_clean.copy()
        df_niche_clean['Custom_Cat'] = df_niche_clean.apply(categorize_cck_niche, axis=1)
        
        niche_aggs = df_niche_clean.groupby('Custom_Cat').agg(
            Total_Units=('TOTAL_num', 'sum'), Sold_Units=('SOLD_num', 'sum'),
            Balance_Units=('BALANCE_num', 'sum'), Value_of_Balance=('Calculated_Value', 'sum')
        ).rename(columns={'Total_Units': 'Total Units', 'Sold_Units': 'Sold Units', 'Balance_Units': 'Balance Units', 'Value_of_Balance': 'Value of Balance'})
        
        # Ensure ordered layout output
        niche_aggs = niche_aggs.reindex(['Niche - Single', 'Niche - Double', 'Niche - Buddha', 'Niche - Others']).fillna(0)
        processed_data['CCK_Niche'] = format_summary_matrix(niche_aggs)

    # ==========================================
    # 2. LST & TLT BRANCHES PARSING
    # ==========================================
    for sheet_name, save_key in [('LST-TABLE & NICHE', 'LST'), ('TLT-TABLE & NICHE', 'TLT')]:
        if sheet_name in xls.sheet_names:
            df_branch = pd.read_excel(xls, sheet_name=sheet_name, header=1)
            df_branch = clean_dataframe(df_branch)
            
            def categorize_lst_tlt(row):
                prod = str(row['PRODUCT']).strip().upper()
                lot = str(row['LOT TYPE']).strip().upper()
                if 'TABLET' in prod:
                    return 'Pedestal'
                elif 'NICHE' in prod:
                    if 'SINGLE' in lot or 'SG' in lot: return 'Niche - Single'
                    if 'DOUBLE' in lot or 'DB' in lot: return 'Niche - Double'
                return None
                
            df_branch['Custom_Cat'] = df_branch.apply(categorize_lst_tlt, axis=1)
            df_branch_clean = df_branch[df_branch['Custom_Cat'].notna()]
            
            branch_aggs = df_branch_clean.groupby('Custom_Cat').agg(
                Total_Units=('TOTAL_num', 'sum'), Sold_Units=('SOLD_num', 'sum'),
                Balance_Units=('BALANCE_num', 'sum'), Value_of_Balance=('Calculated_Value', 'sum')
            ).rename(columns={'Total_Units': 'Total Units', 'Sold_Units': 'Sold Units', 'Balance_Units': 'Balance Units', 'Value_of_Balance': 'Value of Balance'})
            
            branch_aggs = branch_aggs.reindex(['Pedestal', 'Niche - Single', 'Niche - Double']).fillna(0)
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
    with st.spinner("Compiling structural views..."):
        data_package = process_uploaded_excel(uploaded_file)
    st.sidebar.success("🎉 Matrix Form Factor Parsed!")
    
    branch_choice = st.sidebar.radio("Select Branch Dashboard:", ["CCK Branch", "LST Branch", "TLT Branch"])
    
    if branch_choice == "CCK Branch":
        st.header("🏢 CCK Branch Segment Analysis")
        if 'CCK_Pedestal' in data_package:
            st.subheader("📦 Pedestal / Tablet Matrix")
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
    st.info("💡 Please upload an Excel document to populate the branch matrices.")
