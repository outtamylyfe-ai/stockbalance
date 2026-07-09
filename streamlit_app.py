import streamlit as tf
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# 1. Page Configuration
st.set_page_config(
    page_title="Inventory Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling for a Polished UI
st.markdown("""
    <style>
    .main-title { font-size: 38px; font-weight: 700; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 16px; color: #4B5563; margin-bottom: 25px; }
    .section-header { font-size: 22px; font-weight: 600; color: #1F2937; margin-top: 15px; margin-bottom: 15px; border-bottom: 2px solid #E5E7EB; padding-bottom: 5px; }
    .metric-card { background-color: #F3F4F6; padding: 15px; border-radius: 8px; border-left: 5px solid #2563EB; }
    </style>
""", unsafe_allow_html=True)

# 2. Deployment Instructions (Collapsible)
with st.expander("🚀 Deployment Guide (GitHub & Streamlit Community Cloud)", expanded=False):
    st.markdown("""
    1. **Create a GitHub Repository**: Upload this script as `app.py` and create a `requirements.txt` file including `streamlit`, `pandas`, `openpyxl`, `plotly`, and `reportlab`.
    2. **Deploy via Streamlit**: Go to [share.streamlit.io](https://share.streamlit.io), connect your GitHub repository, choose the branch, select `app.py` as the main entry point, and click **Deploy**.
    """)

st.markdown('<div class="main-title">Inventory Analytics Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your branch inventory workbook to perform data verification, calculations, and export comprehensive reports.</div>', unsafe_allow_html=True)

# 3. File Uploader Widget
uploaded_file = st.file_uploader("Choose your inventory Excel file (.xlsx)", type=["xlsx"])

# Helper function for safe number conversions
def safe_numeric(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

# 4. Main Application Logic
if uploaded_file is not None:
    try:
        # Load sheets with headers starting on Row 2 (index 1)
        xls = pd.ExcelFile(uploaded_file)
        required_sheets = ['CCK-TABLET', 'CCK-NICHE', 'LST-TABLE & NICHE', 'TLT-TABLE & NICHE']
        
        # Verify sheet existence safely
        if not all(s in xls.sheet_names for s in required_sheets):
            st.error(f"Error: The workbook must contain exactly these sheets: {required_sheets}")
            st.stop()
            
        df_cck_tablet = pd.read_excel(uploaded_file, sheet_name='CCK-TABLET', header=1)
        df_cck_niche = pd.read_excel(uploaded_file, sheet_name='CCK-NICHE', header=1)
        df_lst = pd.read_excel(uploaded_file, sheet_name='LST-TABLE & NICHE', header=1)
        df_tlt = pd.read_excel(uploaded_file, sheet_name='TLT-TABLE & NICHE', header=1)
        
        # Strip string whitespace from column headers safely
        for df in [df_cck_tablet, df_cck_niche, df_lst, df_tlt]:
            df.columns = df.columns.str.strip()
            
        # Global dictionary to track processed data frames for PDF export
        pdf_data_summary = {}

        # Define tabs
        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 CCK Tablet", 
            "🪦 CCK Niche", 
            "🏛️ LST Tablet & Niche", 
            "🌅 TLT Tablet & Niche"
        ])

        # ==========================================
        # TAB 1: CCK-TABLET
        # ==========================================
        with tab1:
            st.markdown('<div class="section-header">CCK Tablet Inventory Analysis</div>', unsafe_allow_html=True)
            
            # Filters
            blocks_available = df_cck_tablet['BLOCK'].dropna().unique()
            selected_blocks = st.multiselect("Filter by BLOCK:", options=blocks_available, default=[b for b in ['A', 'B', 'C'] if b in blocks_available])
            
            df_t1_filtered = df_cck_tablet[df_cck_tablet['BLOCK'].isin(selected_blocks)].copy()
            
            if not df_t1_filtered.empty:
                # Calculations
                df_t1_filtered['Balance Unit (%)'] = safe_numeric(df_t1_filtered['BALANCE %'])
                df_t1_filtered['Sold (%)'] = 1.0 - df_t1_filtered['Balance Unit (%)']
                df_t1_filtered['Number of Balance Units'] = safe_numeric(df_t1_filtered['BALANCE'])
                df_t1_filtered['Value of Balance'] = df_t1_filtered['Number of Balance Units'] * safe_numeric(df_t1_filtered['AVG PO PRICE'])
                
                # Summary View / Pivot Table
                pivot_t1 = df_t1_filtered.groupby('BLOCK').agg({
                    'TOTAL': 'sum',
                    'TOTAL SOLD': 'sum',
                    'Number of Balance Units': 'sum',
                    'Value of Balance': 'sum'
                }).reset_index()
                
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.subheader("Block Summary Metrics")
                    st.dataframe(pivot_t1.style.format({
                        'TOTAL': '{:,.0f}', 'TOTAL SOLD': '{:,.0f}', 
                        'Number of Balance Units': '{:,.0f}', 'Value of Balance': '${:,.2f}'
                    }), use_container_width=True)
                
                with col2:
                    st.subheader("Inventory Distribution Chart")
                    fig_t1 = go.Figure()
                    fig_t1.add_trace(go.Bar(x=pivot_t1['BLOCK'], y=pivot_t1['TOTAL SOLD'], name='Units Sold', marker_color='#10B981'))
                    fig_t1.add_trace(go.Bar(x=pivot_t1['BLOCK'], y=pivot_t1['Number of Balance Units'], name='Balance Units', marker_color='#3B82F6'))
                    fig_t1.update_layout(barmode='stack', height=300, margin=dict(t=20, b=20, l=20, r=20))
                    st.plotly_chart(fig_t1, use_container_width=True)
                
                st.subheader("Granular Data View")
                st.dataframe(df_t1_filtered[['BLOCK', 'SUITE NO.', 'TOTAL', 'TOTAL SOLD', 'Balance Unit (%)', 'Sold (%)', 'Number of Balance Units', 'Value of Balance']], use_container_width=True)
                pdf_data_summary['CCK-TABLET'] = pivot_t1
            else:
                st.warning("No data found for the selected Block filters.")

        # ==========================================
        # TAB 2: CCK-NICHE
        # ==========================================
        with tab2:
            st.markdown('<div class="section-header">CCK Niche Lot Type Breakdown</div>', unsafe_allow_html=True)
            
            df_t2 = df_cck_niche.copy()
            if 'LOT TYPE' in df_t2.columns:
                df_t2['LOT TYPE CLEAN'] = df_t2['LOT TYPE'].astype(str).str.strip().str.upper()
                
                # Standard Categories Mapping
                def categorize_lot(val):
                    if val in ['SINGLE', 'DOUBLE', 'FAMILY']:
                        return val
                    return 'OTHERS'
                
                df_t2['Category'] = df_t2['LOT TYPE CLEAN'].apply(categorize_lot)
                
                # Calculations
                df_t2['BALANCE'] = safe_numeric(df_t2['BALANCE'])
                df_t2['TOTAL'] = safe_numeric(df_t2['TOTAL'])
                
                cat_summary = df_t2.groupby('Category').agg({
                    'TOTAL': 'sum',
                    'BALANCE': 'sum'
                }).reindex(['SINGLE', 'DOUBLE', 'FAMILY', 'OTHERS'], fill_value=0).reset_index()
                
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.subheader("Core Category Counts")
                    st.dataframe(cat_summary.style.format({'TOTAL': '{:,.0f}', 'BALANCE': '{:,.0f}'}), use_container_width=True)
                
                with col2:
                    st.subheader("Category Distribution")
                    fig_t2 = px.pie(cat_summary, values='TOTAL', names='Category', color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_t2.update_layout(height=280, margin=dict(t=10, b=10, l=10, r=10))
                    st.plotly_chart(fig_t2, use_container_width=True)
                
                # Expandable view for 'OTHERS' group details
                with st.expander("🔍 Click to view 'OTHERS' drill-down sub-types"):
                    df_others = df_t2[df_t2['Category'] == 'OTHERS']
                    if not df_others.empty:
                        others_breakdown = df_others.groupby('LOT TYPE').agg({
                            'TOTAL': 'sum',
                            'BALANCE': 'sum'
                        }).reset_index()
                        st.dataframe(others_breakdown.style.format({'TOTAL': '{:,.0f}', 'BALANCE': '{:,.0f}'}), use_container_width=True)
                    else:
                        st.info("No sub-types classified under 'OTHERS' found.")
                pdf_data_summary['CCK-NICHE'] = cat_summary

        # ==========================================
        # TAB 3: LST-TABLE & NICHE
        # ==========================================
        with tab3:
            st.markdown('<div class="section-header">LST Tablet and Niche Deep Dive</div>', unsafe_allow_html=True)
            df_lst['PRODUCT_CLEAN'] = df_lst['PRODUCT'].astype(str).str.strip().str.upper()
            
            # SECTION A: Tablet
            st.subheader("🔹 Section A: Tablet Analysis")
            df_lst_tablet = df_lst[df_lst['PRODUCT_CLEAN'] == 'TABLET'].copy()
            
            suites_available = df_lst_tablet['SUITE NO.'].dropna().unique() if 'SUITE NO.' in df_lst_tablet.columns else []
            selected_suites = st.multiselect("Filter Suite (Tablet):", options=suites_available, default=[s for s in ["DYNASTY 2", "IMPERIAL 2"] if s in suites_available])
            
            df_lst_tab_filtered = df_lst_tablet[df_lst_tablet['SUITE NO.'].isin(selected_suites)].copy() if 'SUITE NO.' in df_lst_tablet.columns else pd.DataFrame()
            
            if not df_lst_tab_filtered.empty:
                df_lst_tab_filtered['Balance (%)'] = safe_numeric(df_lst_tab_filtered['BALANCE %'])
                df_lst_tab_filtered['Sold (%)'] = 1.0 - df_lst_tab_filtered['Balance (%)']
                df_lst_tab_filtered['Number of Units'] = safe_numeric(df_lst_tab_filtered['BALANCE'])
                df_lst_tab_filtered['Value of Balance'] = df_lst_tab_filtered['Number of Units'] * safe_numeric(df_lst_tab_filtered['AVG PO PRICE'])
                
                st.dataframe(df_lst_tab_filtered[['SUITE NO.', 'LOT TYPE', 'Balance (%)', 'Sold (%)', 'Number of Units', 'Value of Balance']], use_container_width=True)
                pdf_data_summary['LST-TABLET'] = df_lst_tab_filtered[['SUITE NO.', 'Number of Units', 'Value of Balance']]
            else:
                st.info("No Tablet data matches target criteria.")
                
            # SECTION B: Niche
            st.write("---")
            st.subheader("🔹 Section B: Niche Analysis")
            df_lst_niche = df_lst[df_lst['PRODUCT_CLEAN'] == 'NICHE'].copy()
            
            lot_types_avail = df_lst_niche['LOT TYPE'].dropna().unique() if 'LOT TYPE' in df_lst_niche.columns else []
            selected_lots = st.multiselect("Filter Lot Type (Niche):", options=lot_types_avail, default=list(lot_types_avail))
            
            df_lst_niche_filtered = df_lst_niche[df_lst_niche['LOT TYPE'].isin(selected_lots)].copy() if 'LOT TYPE' in df_lst_niche.columns else pd.DataFrame()
            
            if not df_lst_niche_filtered.empty:
                df_lst_niche_filtered['Balance (%)'] = safe_numeric(df_lst_niche_filtered['BALANCE %'])
                df_lst_niche_filtered['Sold (%)'] = 1.0 - df_lst_niche_filtered['Balance (%)']
                df_lst_niche_filtered['Number of Units'] = safe_numeric(df_lst_niche_filtered['BALANCE'])
                df_lst_niche_filtered['Value of Balance'] = df_lst_niche_filtered['Number of Units'] * safe_numeric(df_lst_niche_filtered['AVG PO PRICE'])
                
                st.dataframe(df_lst_niche_filtered[['SUITE NO.', 'LOT TYPE', 'Balance (%)', 'Sold (%)', 'Number of Units', 'Value of Balance']], use_container_width=True)
                pdf_data_summary['LST-NICHE'] = df_lst_niche_filtered[['LOT TYPE', 'Number of Units', 'Value of Balance']]
            else:
                st.info("No Niche rows found matching active lot filters.")

        # ==========================================
        # TAB 4: TLT-TABLE & NICHE
        # ==========================================
        with tab4:
            st.markdown('<div class="section-header">TLT Tablet & Niche Snapshot</div>', unsafe_allow_html=True)
            df_tlt['PRODUCT_CLEAN'] = df_tlt['PRODUCT'].astype(str).str.strip().str.upper()
            df_tlt_filtered = df_tlt[df_tlt['PRODUCT_CLEAN'].isin(['TABLET', 'NICHE'])].copy()
            
            if not df_tlt_filtered.empty:
                df_tlt_filtered['Balance (%)'] = safe_numeric(df_tlt_filtered['BALANCE %'])
                df_tlt_filtered['Sold (%)'] = 1.0 - df_tlt_filtered['Balance (%)']
                df_tlt_filtered['Value of Balance'] = safe_numeric(df_tlt_filtered['BALANCE']) * safe_numeric(df_tlt_filtered['AVG PO PRICE'])
                
                # Aggregations for KPI displays
                tot_bal_units = safe_numeric(df_tlt_filtered['BALANCE']).sum()
                tot_val_bal = df_tlt_filtered['Value of Balance'].sum()
                avg_sold_pct = df_tlt_filtered['Sold (%)'].mean() * 100
                
                kpi1, kpi2, kpi3 = st.columns(3)
                with kpi1:
                    st.metric("Total Balance Units", f"{tot_bal_units:,.0f}")
                with kpi2:
                    st.metric("Total Portfolio Balance Value", f"${tot_val_bal:,.2f}")
                with kpi3:
                    st.metric("Mean Sales Rate (%)", f"{avg_sold_pct:.2f}%")
                    
                st.subheader("Consolidated Overview Matrix")
                st.dataframe(df_tlt_filtered[['PRODUCT', 'SUITE NO.', 'LOT TYPE', 'Balance (%)', 'Sold (%)', 'Value of Balance']], use_container_width=True)
                pdf_data_summary['TLT-SUMMARY'] = df_tlt_filtered[['PRODUCT', 'Value of Balance']]
            else:
                st.warning("No dynamic Tablet or Niche product lines detected.")

        # ==========================================
        # PDF GENERATION & EXPORT LOGIC
        # ==========================================
        st.sidebar.markdown("---")
        st.sidebar.subheader("📥 Export Summary Data")
        
        def build_pdf_report(summary_dict):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=24, leading=28, textColor=colors.HexColor('#1E3A8A'), spaceAfter=15)
            h1_style = ParagraphStyle('Heading1Custom', parent=styles['Heading2'], fontSize=16, leading=20, textColor=colors.HexColor('#1F2937'), spaceBefore=15, spaceAfter=10)
            body_style = ParagraphStyle('BodyCustom', parent=styles['Normal'], fontSize=10, leading=14, textColor=colors.HexColor('#4B5563'))
            
            # Header
            story.append(Paragraph("Consolidated Branch Inventory Executive Report", title_style))
            story.append(Paragraph("Generated automatically via Streamlit Data Deployment Stack.", body_style))
            story.append(Spacer(1, 15))
            
            # Render individual block statistics safely onto pages
            for section_name, dataframe in summary_dict.items():
                story.append(Paragraph(f"📝 Section: {section_name}", h1_style))
                
                # Transform data values safely to lists of strings for reportlab layout matching
                raw_data = [dataframe.columns.tolist()] + dataframe.astype(str).values.tolist()
                
                table_object = Table(raw_data, hAlign='LEFT')
                table_object.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F9FAFB')),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
                ]))
                story.append(KeepTogether([table_object]))
                story.append(Spacer(1, 12))
                
            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()

        if pdf_data_summary:
            pdf_bytes = build_pdf_report(pdf_data_summary)
            st.sidebar.download_button(
                label="Download Executive PDF Report",
                data=pdf_bytes,
                file_name="Inventory_Consolidated_Executive_Report.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error parsing workbook contents safely: {str(e)}")
else:
    st.info("⚠️ Please upload a valid branch inventory spreadsheet (.xlsx) in the file area to execute system calculations.")
