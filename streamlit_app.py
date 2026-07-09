import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Page Configuration
st.set_page_config(
    page_title="Inventory Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Polished CSS
st.markdown("""
    <style>
    .main-title { font-size: 36px; font-weight: 700; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 15px; color: #4B5563; margin-bottom: 25px; }
    .section-header { font-size: 20px; font-weight: 600; color: #1F2937; margin-top: 15px; margin-bottom: 15px; border-bottom: 2px solid #E5E7EB; padding-bottom: 5px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Inventory Analytics Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your branch inventory workbook to view dynamically cleaned metrics, summaries, and automated PDF export profiles.</div>', unsafe_allow_html=True)

# File Uploader
uploaded_file = st.file_uploader("Choose your inventory Excel file (.xlsx)", type=["xlsx"])

# Robust Data Cleaning Helpers to avoid double-counting subtotal rows
def clean_numeric(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

def filter_out_subtotals(df, check_column):
    """Filters out built-in Excel Subtotal, Total, and Footnote rows dynamically."""
    if check_column not in df.columns:
        return df
    
    # Drop rows where checking column is null or contains 'total', 'sub total', or lengthy footnotes
    mask = df[check_column].astype(str).str.upper().str.strip()
    filtered_df = df[
        df[check_column].notna() & 
        (~mask.str.contains('TOTAL')) & 
        (~mask.str.contains('SUB TOTAL')) & 
        (~mask.str.contains('STATUS:')) & 
        (mask.str.len() < 50) # Drop long descriptive sentences
    ].copy()
    return filtered_df

if uploaded_file is not None:
    try:
        # Load sheets (header starts on row 2 / index 1)
        xls = pd.ExcelFile(uploaded_file)
        required_sheets = ['CCK-TABLET', 'CCK-NICHE', 'LST-TABLE & NICHE', 'TLT-TABLE & NICHE']
        
        if not all(s in xls.sheet_names for s in required_sheets):
            st.error(f"Missing Sheets! The uploaded workbook must contain exactly: {required_sheets}")
            st.stop()
            
        # Global dictionary tracking precise outputs for PDF report generation
        pdf_report_payload = {}

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
            df_raw1 = pd.read_excel(uploaded_file, sheet_name='CCK-TABLET', header=1)
            df_raw1.columns = df_raw1.columns.str.strip()
            
            # Clean and isolate core rows
            df_t1 = filter_out_subtotals(df_raw1, 'BLOCK')
            
            blocks_available = sorted(df_t1['BLOCK'].unique())
            selected_blocks = st.multiselect("Filter by BLOCK:", options=blocks_available, default=[b for b in ['A', 'B', 'C'] if b in blocks_available])
            
            df_t1_filtered = df_t1[df_t1['BLOCK'].isin(selected_blocks)].copy()
            
            if not df_t1_filtered.empty:
                # Direct assignments based on requirement criteria
                df_t1_filtered['Balance Unit (%)'] = clean_numeric(df_t1_filtered['BALANCE %'])
                df_t1_filtered['Sold (%)'] = 1.0 - df_t1_filtered['Balance Unit (%)']
                df_t1_filtered['Number of Balance Units'] = clean_numeric(df_t1_filtered['BALANCE'])
                df_t1_filtered['Value of Balance'] = df_t1_filtered['Number of Balance Units'] * clean_numeric(df_t1_filtered['AVG PO PRICE'])
                
                # Dynamic Pivot
                pivot_t1 = df_t1_filtered.groupby('BLOCK').agg({
                    'TOTAL': 'sum',
                    'TOTAL SOLD': 'sum',
                    'Number of Balance Units': 'sum',
                    'Value of Balance': 'sum'
                }).reset_index()
                
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.subheader("Accurate Block Metrics")
                    st.dataframe(pivot_t1.style.format({
                        'TOTAL': '{:,.0f}', 'TOTAL SOLD': '{:,.0f}', 
                        'Number of Balance Units': '{:,.0f}', 'Value of Balance': '${:,.2f}'
                    }), use_container_width=True)
                
                with col2:
                    st.subheader("Visual Inventory Balance Ratio")
                    fig_t1 = go.Figure()
                    fig_t1.add_trace(go.Bar(x=pivot_t1['BLOCK'], y=pivot_t1['TOTAL SOLD'], name='Units Sold', marker_color='#10B981'))
                    fig_t1.add_trace(go.Bar(x=pivot_t1['BLOCK'], y=pivot_t1['Number of Balance Units'], name='Balance Units', marker_color='#3B82F6'))
                    fig_t1.update_layout(barmode='stack', height=300, margin=dict(t=20, b=20, l=20, r=20))
                    st.plotly_chart(fig_t1, use_container_width=True)
                
                pdf_report_payload['CCK-TABLET'] = pivot_t1
            else:
                st.warning("No data found for the active block criteria.")

        # ==========================================
        # TAB 2: CCK-NICHE
        # ==========================================
        with tab2:
            st.markdown('<div class="section-header">CCK Niche Lot Type Breakdown</div>', unsafe_allow_html=True)
            df_raw2 = pd.read_excel(uploaded_file, sheet_name='CCK-NICHE', header=1)
            df_raw2.columns = df_raw2.columns.str.strip()
            
            df_t2 = filter_out_subtotals(df_raw2, 'BLOCK')
            
            if 'LOT TYPE' in df_t2.columns:
                df_t2['LOT TYPE CLEAN'] = df_t2['LOT TYPE'].astype(str).str.strip().str.upper()
                
                # Interchangeable mapping logic + Categorization
                def assign_niche_category(val):
                    if val in ['SINGLE', 'DOUBLE', 'FAMILY']:
                        return val
                    return 'OTHERS'
                
                df_t2['Category'] = df_t2['LOT TYPE CLEAN'].apply(assign_niche_category)
                df_t2['TOTAL'] = clean_numeric(df_t2['TOTAL'])
                df_t2['BALANCE'] = clean_numeric(df_t2['BALANCE'])
                
                cat_summary = df_t2.groupby('Category').agg({
                    'TOTAL': 'sum',
                    'BALANCE': 'sum'
                }).reindex(['SINGLE', 'DOUBLE', 'FAMILY', 'OTHERS'], fill_value=0).reset_index()
                
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.subheader("Niche Category Metrics Summary")
                    st.dataframe(cat_summary.style.format({'TOTAL': '{:,.0f}', 'BALANCE': '{:,.0f}'}), use_container_width=True)
                
                with col2:
                    st.subheader("Category Distribution Pie Chart")
                    fig_t2 = px.pie(cat_summary, values='TOTAL', names='Category', color_discrete_sequence=px.colors.qualitative.Safe)
                    fig_t2.update_layout(height=280, margin=dict(t=10, b=10, l=10, r=10))
                    st.plotly_chart(fig_t2, use_container_width=True)
                
                with st.expander("🔍 Deep-Dive Breakdown of the 'OTHERS' Bucket"):
                    df_others = df_t2[df_t2['Category'] == 'OTHERS']
                    if not df_others.empty:
                        others_breakdown = df_others.groupby('LOT TYPE').agg({
                            'TOTAL': 'sum',
                            'BALANCE': 'sum'
                        }).reset_index()
                        st.dataframe(others_breakdown.style.format({'TOTAL': '{:,.0f}', 'BALANCE': '{:,.0f}'}), use_container_width=True)
                        pdf_report_payload['CCK-NICHE-OTHERS'] = others_breakdown
                
                pdf_report_payload['CCK-NICHE-MAIN'] = cat_summary

        # ==========================================
        # TAB 3: LST-TABLE & NICHE
        # ==========================================
        with tab3:
            st.markdown('<div class="section-header">LST Tablet and Niche Deep Dive</div>', unsafe_allow_html=True)
            df_raw3 = pd.read_excel(uploaded_file, sheet_name='LST-TABLE & NICHE', header=1)
            df_raw3.columns = df_raw3.columns.str.strip()
            
            df_t3 = filter_out_subtotals(df_raw3, 'PRODUCT')
            df_t3['PRODUCT_CLEAN'] = df_t3['PRODUCT'].astype(str).str.strip().str.upper()
            
            # Section A: Tablet
            st.subheader("🔹 Section A: Tablet Analysis (Dynasty 2 & Imperial 2)")
            df_lst_tab = df_t3[df_t3['PRODUCT_CLEAN'] == 'TABLET'].copy()
            if 'SUITE NO.' in df_lst_tab.columns:
                df_lst_tab['SUITE_CLEAN'] = df_lst_tab['SUITE NO.'].astype(str).str.strip().str.upper()
                df_lst_tab_filtered = df_lst_tab[df_lst_tab['SUITE_CLEAN'].isin(['DYNASTY 2', 'IMPERIAL 2'])].copy()
                
                if not df_lst_tab_filtered.empty:
                    df_lst_tab_filtered['Balance (%)'] = clean_numeric(df_lst_tab_filtered['BALANCE %'])
                    df_lst_tab_filtered['Sold (%)'] = 1.0 - df_lst_tab_filtered['Balance (%)']
                    df_lst_tab_filtered['Number of Units'] = clean_numeric(df_lst_tab_filtered['BALANCE'])
                    df_lst_tab_filtered['Value of Balance'] = df_lst_tab_filtered['Number of Units'] * clean_numeric(df_lst_tab_filtered['AVG PO PRICE'])
                    
                    st.dataframe(df_lst_tab_filtered[['SUITE NO.', 'LOT TYPE', 'Balance (%)', 'Sold (%)', 'Number of Units', 'Value of Balance']].style.format({
                        'Balance (%)': '{:.2%}', 'Sold (%)': '{:.2%}', 'Number of Units': '{:,.0f}', 'Value of Balance': '${:,.2f}'
                    }), use_container_width=True)
                    pdf_report_payload['LST-TABLET'] = df_lst_tab_filtered[['SUITE NO.', 'LOT TYPE', 'BALANCE', 'Value of Balance']]
                else:
                    st.info("No matching data rows found for Dynasty 2 or Imperial 2 Tablet lines.")

            # Section B: Niche
            st.write("---")
            st.subheader("🔹 Section B: Niche Analysis")
            df_lst_niche = df_t3[df_t3['PRODUCT_CLEAN'] == 'NICHE'].copy()
            
            if not df_lst_niche.empty:
                lot_types_lst = sorted(df_lst_niche['LOT TYPE'].dropna().unique())
                selected_lots_lst = st.multiselect("Filter Niche Lot Types:", options=lot_types_lst, default=lot_types_lst)
                
                df_lst_niche_filtered = df_lst_niche[df_lst_niche['LOT TYPE'].isin(selected_lots_lst)].copy()
                if not df_lst_niche_filtered.empty:
                    df_lst_niche_filtered['Balance (%)'] = clean_numeric(df_lst_niche_filtered['BALANCE %'])
                    df_lst_niche_filtered['Sold (%)'] = 1.0 - df_lst_niche_filtered['Balance (%)']
                    df_lst_niche_filtered['Number of Units'] = clean_numeric(df_lst_niche_filtered['BALANCE'])
                    df_lst_niche_filtered['Value of Balance'] = df_lst_niche_filtered['Number of Units'] * clean_numeric(df_lst_niche_filtered['AVG PO PRICE'])
                    
                    st.dataframe(df_lst_niche_filtered[['SUITE NO.', 'LOT TYPE', 'Balance (%)', 'Sold (%)', 'Number of Units', 'Value of Balance']].style.format({
                        'Balance (%)': '{:.2%}', 'Sold (%)': '{:.2%}', 'Number of Units': '{:,.0f}', 'Value of Balance': '${:,.2f}'
                    }), use_container_width=True)
                    pdf_report_payload['LST-NICHE'] = df_lst_niche_filtered[['SUITE NO.', 'LOT TYPE', 'BALANCE', 'Value of Balance']]

        # ==========================================
        # TAB 4: TLT-TABLE & NICHE
        # ==========================================
        with tab4:
            st.markdown('<div class="section-header">TLT Tablet & Niche Snapshot</div>', unsafe_allow_html=True)
            df_raw4 = pd.read_excel(uploaded_file, sheet_name='TLT-TABLE & NICHE', header=1)
            df_raw4.columns = df_raw4.columns.str.strip()
            
            df_tlt = filter_out_subtotals(df_raw4, 'PRODUCT')
            df_tlt['PRODUCT_CLEAN'] = df_tlt['PRODUCT'].astype(str).str.strip().str.upper()
            
            # Filter solely for TABLET and NICHE rows
            df_tlt_filtered = df_tlt[df_tlt['PRODUCT_CLEAN'].isin(['TABLET', 'NICHE'])].copy()
            
            if not df_tlt_filtered.empty:
                df_tlt_filtered['Balance (%)'] = clean_numeric(df_tlt_filtered['BALANCE %'])
                df_tlt_filtered['Sold (%)'] = 1.0 - df_tlt_filtered['Balance (%)']
                df_tlt_filtered['Value of Balance'] = clean_numeric(df_tlt_filtered['BALANCE']) * clean_numeric(df_tlt_filtered['AVG PO PRICE'])
                
                # Accurate Metric Cards
                tot_units = clean_numeric(df_tlt_filtered['BALANCE']).sum()
                tot_val = df_tlt_filtered['Value of Balance'].sum()
                
                m1, m2 = st.columns(2)
                with m1:
                    st.metric("Total Balance Units", f"{tot_units:,.0f}")
                with m2:
                    st.metric("Total Value of Balance", f"${tot_val:,.2f}")
                
                st.subheader("Data Summary Matrix")
                st.dataframe(df_tlt_filtered[['PRODUCT', 'SUITE NO.', 'LOT TYPE', 'Balance (%)', 'Sold (%)', 'Value of Balance']].style.format({
                    'Balance (%)': '{:.2%}', 'Sold (%)': '{:.2%}', 'Value of Balance': '${:,.2f}'
                }), use_container_width=True)
                
                pdf_report_payload['TLT-SUMMARY'] = df_tlt_filtered[['PRODUCT', 'SUITE NO.', 'BALANCE', 'Value of Balance']]

        # ==========================================
        # PDF EXPORT GENERATOR
        # ==========================================
        st.sidebar.markdown("---")
        st.sidebar.subheader("📥 Export Data Summary")
        
        def build_pdf_report(summary_dict):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=40, bottomMargin=40)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=22, leading=26, textColor=colors.HexColor('#1E3A8A'), spaceAfter=15)
            h1_style = ParagraphStyle('Heading1Custom', parent=styles['Heading2'], fontSize=14, leading=18, textColor=colors.HexColor('#1F2937'), spaceBefore=12, spaceAfter=8)
            
            story.append(Paragraph("Consolidated Branch Inventory Executive Report", title_style))
            story.append(Spacer(1, 10))
            
            for section_name, dataframe in summary_dict.items():
                story.append(Paragraph(f"📝 Section: {section_name}", h1_style))
                
                # Format long floating values for readability in standard lists
                formatted_df = dataframe.copy()
                for col in formatted_df.columns:
                    if formatted_df[col].dtype == 'float64':
                        formatted_df[col] = formatted_df[col].apply(lambda x: f"{x:,.2f}" if x > 100 else f"{x:.2%}")
                
                raw_data = [formatted_df.columns.tolist()] + formatted_df.astype(str).values.tolist()
                
                table_object = Table(raw_data, hAlign='LEFT')
                table_object.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F9FAFB')),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
                ]))
                story.append(KeepTogether([table_object]))
                story.append(Spacer(1, 10))
                
            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()

        if pdf_report_payload:
            pdf_bytes = build_pdf_report(pdf_report_payload)
            st.sidebar.download_button(
                label="Download Verified Executive PDF",
                data=pdf_bytes,
                file_name="Inventory_Clean_Executive_Report.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error parsing workbook contents safely: {str(e)}")
else:
    st.info("⚠️ Please upload a valid branch inventory spreadsheet (.xlsx) to get started.")
else:
    st.info("⚠️ Please upload a valid branch inventory spreadsheet (.xlsx) in the file area to execute system calculations.")
