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

# Custom Polished CSS Style Treatments
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

# Robust Data Cleaning Helper
def clean_numeric(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

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
            df_t1_raw = pd.read_excel(uploaded_file, sheet_name='CCK-TABLET', header=1)
            df_t1_raw.columns = df_t1_raw.columns.str.strip()
            
            if 'BLOCK' in df_t1_raw.columns:
                # Filter out pre-existing total/summary rows to avoid double-counting
                df_t1 = df_t1_raw[df_t1_raw['BLOCK'].notna() & (~df_t1_raw['BLOCK'].astype(str).str.upper().str.contains('TOTAL'))].copy()
                
                # Perform clean numerical casts and row-by-row balance calculations
                df_t1['TOTAL'] = clean_numeric(df_t1['TOTAL'])
                df_t1['TOTAL SOLD'] = clean_numeric(df_t1['TOTAL SOLD'])
                df_t1['BALANCE'] = clean_numeric(df_t1['BALANCE'])
                df_t1['Value of Balance'] = df_t1['BALANCE'] * clean_numeric(df_t1['AVG PO PRICE'])
                
                # Aggregate safely across categories
                pivot_t1 = df_t1.groupby('BLOCK').agg({
                    'TOTAL': 'sum', 'TOTAL SOLD': 'sum', 'BALANCE': 'sum', 'Value of Balance': 'sum'
                }).reset_index()
                
                # Dynamic percentage weighting
                pivot_t1['Balance (%)'] = pivot_t1['BALANCE'] / pivot_t1['TOTAL']
                pivot_t1['Sold (%)'] = pivot_t1['TOTAL SOLD'] / pivot_t1['TOTAL']
                
                # Append Grand Total Row
                t1_total = pd.DataFrame([{
                    'BLOCK': 'GRAND TOTAL',
                    'TOTAL': pivot_t1['TOTAL'].sum(),
                    'TOTAL SOLD': pivot_t1['TOTAL SOLD'].sum(),
                    'BALANCE': pivot_t1['BALANCE'].sum(),
                    'Value of Balance': pivot_t1['Value of Balance'].sum(),
                    'Balance (%)': pivot_t1['BALANCE'].sum() / pivot_t1['TOTAL'].sum(),
                    'Sold (%)': pivot_t1['TOTAL SOLD'].sum() / pivot_t1['TOTAL'].sum()
                }])
                
                pivot_t1_final = pd.concat([pivot_t1, t1_total], ignore_index=True)[[
                    'BLOCK', 'TOTAL', 'TOTAL SOLD', 'BALANCE', 'Balance (%)', 'Sold (%)', 'Value of Balance'
                ]]
                
                st.dataframe(pivot_t1_final.style.format({
                    'TOTAL': '{:,.0f}', 'TOTAL SOLD': '{:,.0f}', 'BALANCE': '{:,.0f}',
                    'Balance (%)': '{:.2%}', 'Sold (%)': '{:.2%}', 'Value of Balance': '${:,.2f}'
                }), use_container_width=True)
                
                pdf_report_payload['CCK-TABLET'] = pivot_t1_final

        # ==========================================
        # TAB 2: CCK-NICHE
        # ==========================================
        with tab2:
            st.markdown('<div class="section-header">CCK Niche Lot Type Breakdown</div>', unsafe_allow_html=True)
            df_t2_raw = pd.read_excel(uploaded_file, sheet_name='CCK-NICHE', header=1)
            df_t2_raw.columns = df_t2_raw.columns.str.strip()
            
            if 'LOT TYPE' in df_t2_raw.columns:
                check_col = 'BLOCK' if 'BLOCK' in df_t2_raw.columns else 'LOT TYPE'
                df_t2 = df_t2_raw[df_t2_raw[check_col].notna() & (~df_t2_raw[check_col].astype(str).str.upper().str.contains('TOTAL'))].copy()
                
                df_t2['TOTAL'] = clean_numeric(df_t2['TOTAL'])
                df_t2['TOTAL SOLD'] = clean_numeric(df_t2['TOTAL SOLD'])
                df_t2['BALANCE'] = clean_numeric(df_t2['BALANCE'])
                df_t2['Value of Balance'] = df_t2['BALANCE'] * clean_numeric(df_t2['AVG PO PRICE'])
                
                # Standardize strings to clean categories
                df_t2['Category'] = df_t2['LOT TYPE'].astype(str).str.strip().str.upper().apply(
                    lambda x: x if x in ['SINGLE', 'DOUBLE', 'FAMILY'] else 'OTHERS'
                )
                
                pivot_t2 = df_t2.groupby('Category').agg({
                    'TOTAL': 'sum', 'TOTAL SOLD': 'sum', 'BALANCE': 'sum', 'Value of Balance': 'sum'
                }).reindex(['SINGLE', 'DOUBLE', 'FAMILY', 'OTHERS'], fill_value=0).reset_index()
                
                pivot_t2['Balance (%)'] = pivot_t2['BALANCE'] / pivot_t2['TOTAL']
                pivot_t2['Sold (%)'] = pivot_t2['TOTAL SOLD'] / pivot_t2['TOTAL']
                
                # Append Grand Total Row
                t2_total = pd.DataFrame([{
                    'Category': 'GRAND TOTAL',
                    'TOTAL': pivot_t2['TOTAL'].sum(),
                    'TOTAL SOLD': pivot_t2['TOTAL SOLD'].sum(),
                    'BALANCE': pivot_t2['BALANCE'].sum(),
                    'Value of Balance': pivot_t2['Value of Balance'].sum(),
                    'Balance (%)': pivot_t2['BALANCE'].sum() / pivot_t2['TOTAL'].sum(),
                    'Sold (%)': pivot_t2['TOTAL SOLD'].sum() / pivot_t2['TOTAL'].sum()
                }])
                
                pivot_t2_final = pd.concat([pivot_t2, t2_total], ignore_index=True)[[
                    'Category', 'TOTAL', 'TOTAL SOLD', 'BALANCE', 'Balance (%)', 'Sold (%)', 'Value of Balance'
                ]]
                
                st.dataframe(pivot_t2_final.style.format({
                    'TOTAL': '{:,.0f}', 'TOTAL SOLD': '{:,.0f}', 'BALANCE': '{:,.0f}',
                    'Balance (%)': '{:.2%}', 'Sold (%)': '{:.2%}', 'Value of Balance': '${:,.2f}'
                }), use_container_width=True)
                
                pdf_report_payload['CCK-NICHE'] = pivot_t2_final

        # ==========================================
        # TAB 3: LST-TABLE & NICHE (Dynamic Multi-Filtering Function)
        # ==========================================
        with tab3:
            st.markdown('<div class="section-header">LST Tablet and Niche Deep Dive</div>', unsafe_allow_html=True)
            df_t3_raw = pd.read_excel(uploaded_file, sheet_name='LST-TABLE & NICHE', header=1)
            df_t3_raw.columns = df_t3_raw.columns.str.strip()
            
            if 'PRODUCT' in df_t3_raw.columns:
                # Strip out inner totals
                df_t3 = df_t3_raw[df_t3_raw['PRODUCT'].notna() & (~df_t3_raw['PRODUCT'].astype(str).str.upper().str.contains('TOTAL'))].copy()
                
                df_t3['TOTAL'] = clean_numeric(df_t3['TOTAL'])
                df_t3['TOTAL SOLD'] = clean_numeric(df_t3['TOTAL SOLD'])
                df_t3['BALANCE'] = clean_numeric(df_t3['BALANCE'])
                df_t3['Value of Balance'] = df_t3['BALANCE'] * clean_numeric(df_t3['AVG PO PRICE'])
                
                available_products = sorted(df_t3['PRODUCT'].dropna().unique())
                available_lots = sorted(df_t3['LOT TYPE'].dropna().unique()) if 'LOT TYPE' in df_t3.columns else []
                
                # Function UI controls
                f_col1, f_col2 = st.columns(2)
                with f_col1:
                    selected_products = st.multiselect("Filter LST Category (Column A):", options=available_products, default=available_products)
                with f_col2:
                    selected_lots = st.multiselect("Filter LST Lot Type (Column D):", options=available_lots, default=available_lots) if available_lots else []
                
                # Check criteria masking rules
                mask = df_t3['PRODUCT'].isin(selected_products)
                if available_lots and selected_lots:
                    mask = mask & df_t3['LOT TYPE'].isin(selected_lots)
                    
                df_t3_filtered = df_t3[mask].copy()
                
                if not df_t3_filtered.empty:
                    group_cols = ['PRODUCT', 'LOT TYPE'] if 'LOT TYPE' in df_t3.columns else ['PRODUCT']
                    pivot_t3 = df_t3_filtered.groupby(group_cols).agg({
                        'TOTAL': 'sum', 'TOTAL SOLD': 'sum', 'BALANCE': 'sum', 'Value of Balance': 'sum'
                    }).reset_index()
                    
                    pivot_t3['Balance (%)'] = pivot_t3['BALANCE'] / pivot_t3['TOTAL']
                    pivot_t3['Sold (%)'] = pivot_t3['TOTAL SOLD'] / pivot_t3['TOTAL']
                    
                    # Append Grand Total Row
                    t3_total_data = {
                        'PRODUCT': 'GRAND TOTAL',
                        'TOTAL': pivot_t3['TOTAL'].sum(),
                        'TOTAL SOLD': pivot_t3['TOTAL SOLD'].sum(),
                        'BALANCE': pivot_t3['BALANCE'].sum(),
                        'Value of Balance': pivot_t3['Value of Balance'].sum(),
                        'Balance (%)': pivot_t3['BALANCE'].sum() / pivot_t3['TOTAL'].sum(),
                        'Sold (%)': pivot_t3['TOTAL SOLD'].sum() / pivot_t3['TOTAL'].sum()
                    }
                    if 'LOT TYPE' in df_t3.columns:
                        t3_total_data['LOT TYPE'] = '-'
                        
                    t3_total = pd.DataFrame([t3_total_data])
                    pivot_t3_final = pd.concat([pivot_t3, t3_total], ignore_index=True)
                    
                    st.dataframe(pivot_t3_final.style.format({
                        'TOTAL': '{:,.0f}', 'TOTAL SOLD': '{:,.0f}', 'BALANCE': '{:,.0f}',
                        'Balance (%)': '{:.2%}', 'Sold (%)': '{:.2%}', 'Value of Balance': '${:,.2f}'
                    }), use_container_width=True)
                    
                    pdf_report_payload['LST-SUMMARY'] = pivot_t3_final
                else:
                    st.warning("No rows match current filter functions.")

        # ==========================================
        # TAB 4: TLT-TABLE & NICHE
        # ==========================================
        with tab4:
            st.markdown('<div class="section-header">TLT Tablet & Niche Snapshot</div>', unsafe_allow_html=True)
            df_t4_raw = pd.read_excel(uploaded_file, sheet_name='TLT-TABLE & NICHE', header=1)
            df_t4_raw.columns = df_t4_raw.columns.str.strip()
            
            if 'PRODUCT' in df_t4_raw.columns:
                df_t4 = df_t4_raw[df_t4_raw['PRODUCT'].notna() & (~df_t4_raw['PRODUCT'].astype(str).str.upper().str.contains('TOTAL'))].copy()
                
                df_t4['TOTAL'] = clean_numeric(df_t4['TOTAL'])
                df_t4['TOTAL SOLD'] = clean_numeric(df_t4['TOTAL SOLD'])
                df_t4['BALANCE'] = clean_numeric(df_t4['BALANCE'])
                df_t4['Value of Balance'] = df_t4['BALANCE'] * clean_numeric(df_t4['AVG PO PRICE'])
                
                pivot_t4 = df_t4.groupby('PRODUCT').agg({
                    'TOTAL': 'sum', 'TOTAL SOLD': 'sum', 'BALANCE': 'sum', 'Value of Balance': 'sum'
                }).reset_index()
                
                pivot_t4['Balance (%)'] = pivot_t4['BALANCE'] / pivot_t4['TOTAL']
                pivot_t4['Sold (%)'] = pivot_t4['TOTAL SOLD'] / pivot_t4['TOTAL']
                
                # Append Grand Total Row
                t4_total = pd.DataFrame([{
                    'PRODUCT': 'GRAND TOTAL',
                    'TOTAL': pivot_t4['TOTAL'].sum(),
                    'TOTAL SOLD': pivot_t4['TOTAL SOLD'].sum(),
                    'BALANCE': pivot_t4['BALANCE'].sum(),
                    'Value of Balance': pivot_t4['Value of Balance'].sum(),
                    'Balance (%)': pivot_t4['BALANCE'].sum() / pivot_t4['TOTAL'].sum(),
                    'Sold (%)': pivot_t4['TOTAL SOLD'].sum() / pivot_t4['TOTAL'].sum()
                }])
                
                pivot_t4_final = pd.concat([pivot_t4, t4_total], ignore_index=True)
                
                st.dataframe(pivot_t4_final.style.format({
                    'TOTAL': '{:,.0f}', 'TOTAL SOLD': '{:,.0f}', 'BALANCE': '{:,.0f}',
                    'Balance (%)': '{:.2%}', 'Sold (%)': '{:.2%}', 'Value of Balance': '${:,.2f}'
                }), use_container_width=True)
                
                pdf_report_payload['TLT-SUMMARY'] = pivot_t4_final

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
