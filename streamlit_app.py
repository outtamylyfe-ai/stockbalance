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
st.markdown("Upload your master tracking sheet to autopopulate metrics separated into **Niche** and **Pedestal / Tablet** columns.")
st.markdown("---")

# ==========================================
# 🛠️ AUTOMATED DATA PROCESSING FUNCTION
# ==========================================
@st.cache_data(ttl=3600)
def process_uploaded_excel(uploaded_file):
    """
    Parses the spreadsheet and aggregates data dynamically into two distinct products:
    1. Niche
    2. Pedestal / Tablet
    """
    xls = pd.ExcelFile(uploaded_file)
    processed_data = {}
    
    # --- HELPER CLEANER FUNCTION ---
    def clean_and_calculate(df):
        df['BALANCE_num'] = pd.to_numeric(df['BALANCE'], errors='coerce').fillna(0)
        df['AVG_PO_num'] = pd.to_numeric(df['AVG PO PRICE'], errors='coerce').fillna(0)
        df['TOTAL_num'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
        df['SOLD_num'] = pd.to_numeric(df['TOTAL SOLD'], errors='coerce').fillna(0)
        df['Calculated_Value'] = df['BALANCE_num'] * df['AVG_PO_num']
        return df

    # ==========================================
    # 1. PROCESS CCK BRANCH
    # ==========================================
    cck_summary = {}
    
    # A. CCK Pedestal / Tablet
    if 'CCK-TABLET' in xls.sheet_names:
        df_cck_tab = pd.read_excel(xls, sheet_name='CCK-TABLET', header=1)
        df_cck_tab.columns = df_cck_tab.columns.str.strip()
        df_cck_tab = clean_and_calculate(df_cck_tab)
        
        # Filter out subtotal/total rows
        df_cck_tab_clean = df_cck_tab[
            df_cck_tab['SUITE NO.'].notna() & 
            (~df_cck_tab['BLOCK'].str.contains('TOTAL|TOTAL:', na=False, case=False))
        ]
        cck_summary['Pedestal / Tablet'] = {
            'Total': df_cck_tab_clean['TOTAL_num'].sum(),
            'Sold': df_cck_tab_clean['SOLD_num'].sum(),
            'Balance': df_cck_tab_clean['BALANCE_num'].sum(),
            'Value': df_cck_tab_clean['Calculated_Value'].sum()
        }
        
    # B. CCK Niche
    if 'CCK-NICHE' in xls.sheet_names:
        df_cck_niche = pd.read_excel(xls, sheet_name='CCK-NICHE', header=1)
        df_cck_niche.columns = df_cck_niche.columns.str.strip()
        df_cck_niche = clean_and_calculate(df_cck_niche)
        
        df_cck_niche_clean = df_cck_niche[
            df_cck_niche['LOT TYPE'].notna() & 
            (~df_cck_niche['BLOCK'].str.contains('TOTAL|TOTAL:', na=False, case=False))
        ]
        cck_summary['Niche'] = {
            'Total': df_cck_niche_clean['TOTAL_num'].sum(),
            'Sold': df_cck_niche_clean['SOLD_num'].sum(),
            'Balance': df_cck_niche_clean['BALANCE_num'].sum(),
            'Value': df_cck_niche_clean['Calculated_Value'].sum()
        }
        
    if cck_summary:
        df_cck = pd.DataFrame(cck_summary)
        # Append Rates
        df_cck.loc['Sold (%)'] = (df_cck.loc['Sold'] / df_cck.loc['Total']).map('{:.2%}'.format)
        df_cck.loc['Balance (%)'] = (df_cck.loc['Balance'] / df_cck.loc['Total']).map('{:.2%}'.format)
        
        # Pretty formatting numbers
        for col in df_cck.columns:
            df_cck.at['Total', col] = f"{int(df_cck.at['Total', col]):,}"
            df_cck.at['Sold', col] = f"{int(df_cck.at['Sold', col]):,}"
            df_cck.at['Balance', col] = f"{int(df_cck.at['Balance', col]):,}"
            df_cck.at['Value', col] = f"$ {df_cck.at['Value', col]:,.2f}"
        processed_data['CCK'] = df_cck

    # ==========================================
    # 2. PROCESS LST BRANCH
    # ==========================================
    if 'LST-TABLE & NICHE' in xls.sheet_names:
        df_lst = pd.read_excel(xls, sheet_name='LST-TABLE & NICHE', header=1)
        df_lst.columns = df_lst.columns.str.strip()
        df_lst = clean_and_calculate(df_lst)
        
        lst_summary = {}
        for p_label, p_key in [('Pedestal / Tablet', 'TABLET'), ('Niche', 'NICHE')]:
            sub_df = df_lst[df_lst['PRODUCT'] == p_key]
            lst_summary[p_label] = {
                'Total': sub_df['TOTAL_num'].sum(),
                'Sold': sub_df['SOLD_num'].sum(),
                'Balance': sub_df['BALANCE_num'].sum(),
                'Value': sub_df['Calculated_Value'].sum()
            }
        
        df_lst_out = pd.DataFrame(lst_summary)
        df_lst_out.loc['Sold (%)'] = (df_lst_out.loc['Sold'] / df_lst_out.loc['Total']).map('{:.2%}'.format)
        df_lst_out.loc['Balance (%)'] = (df_lst_out.loc['Balance'] / df_lst_out.loc['Total']).map('{:.2%}'.format)
        
        for col in df_lst_out.columns:
            df_lst_out.at['Total', col] = f"{int(df_lst_out.at['Total', col]):,}"
            df_lst_out.at['Sold', col] = f"{int(df_lst_out.at['Sold', col]):,}"
            df_lst_out.at['Balance', col] = f"{int(df_lst_out.at['Balance', col]):,}"
            df_lst_out.at['Value', col] = f"$ {df_lst_out.at['Value', col]:,.2f}"
        processed_data['LST'] = df_lst_out

    # ==========================================
    # 3. PROCESS TLT BRANCH
    # ==========================================
    if 'TLT-TABLE & NICHE' in xls.sheet_names:
        df_tlt = pd.read_excel(xls, sheet_name='TLT-TABLE & NICHE', header=1)
        df_tlt.columns = df_tlt.columns.str.strip()
        df_tlt = clean_and_calculate(df_tlt)
        
        tlt_summary = {}
        for p_label, p_key in [('Pedestal / Tablet', 'TABLET'), ('Niche', 'NICHE')]:
            sub_df = df_tlt[df_tlt['PRODUCT'] == p_key]
            tlt_summary[p_label] = {
                'Total': sub_df['TOTAL_num'].sum(),
                'Sold': sub_df['SOLD_num'].sum(),
                'Balance': sub_df['BALANCE_num'].sum(),
                'Value': sub_df['Calculated_Value'].sum()
            }
            
        df_tlt_out = pd.DataFrame(tlt_summary)
        df_tlt_out.loc['Sold (%)'] = (df_tlt_out.loc['Sold'] / df_tlt_out.loc['Total']).map('{:.2%}'.format)
        df_tlt_out.loc['Balance (%)'] = (df_tlt_out.loc['Balance'] / df_tlt_out.loc['Total']).map('{:.2%}'.format)
        
        for col in df_tlt_out.columns:
            df_tlt_out.at['Total', col] = f"{int(df_tlt_out.at['Total', col]):,}"
            df_tlt_out.at['Sold', col] = f"{int(df_tlt_out.at['Sold', col]):,}"
            df_tlt_out.at['Balance', col] = f"{int(df_tlt_out.at['Balance', col]):,}"
            df_tlt_out.at['Value', col] = f"$ {df_tlt_out.at['Value', col]:,.2f}"
        processed_data['TLT'] = df_tlt_out
        
    return processed_data

# ==========================================
# 📥 USER INTERFACE: FILE UPLOADER
# ==========================================
uploaded_file = st.sidebar.file_uploader(
    "Step 1: Upload Master Spreadsheet", 
    type=["xlsx", "xls"],
    help="Upload your master spreadsheet to separate products automatically."
)

if uploaded_file is not None:
    with st.spinner("Processing products..."):
        data_package = process_uploaded_excel(uploaded_file)
    st.sidebar.success("🎉 File Separated Successfully!")
    
    # Branch Navigation Selection
    branch_choice = st.sidebar.radio("Step 2: Select Branch View:", ["CCK Branch", "LST Branch", "TLT Branch"])
    
    if branch_choice == "CCK Branch" and 'CCK' in data_package:
        st.header("🏢 CCK Branch Product Breakdown")
        st.dataframe(data_package['CCK'], use_container_width=True)
        
    elif branch_choice == "LST Branch" and 'LST' in data_package:
        st.header("🏢 LST Branch Product Breakdown")
        st.dataframe(data_package['LST'], use_container_width=True)
        
    elif branch_choice == "TLT Branch" and 'TLT' in data_package:
        st.header("🏢 TLT Branch Product Breakdown")
        st.dataframe(data_package['TLT'], use_container_width=True)
else:
    st.info("💡 Please upload an Excel document via the sidebar menu to view product metrics dynamically.")
