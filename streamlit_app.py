import streamlit as st
import pandas as pd
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

# Custom Style Treatments
st.markdown("""
    <style>
    .main-title { font-size: 36px; font-weight: 700; color: #1E3A8A; margin-bottom: 5px; }
    .subtitle { font-size: 15px; color: #4B5563; margin-bottom: 25px; }
    .section-header { font-size: 20px; font-weight: 600; color: #1F2937; margin-top: 15px; margin-bottom: 15px; border-bottom: 2px solid #E5E7EB; padding-bottom: 5px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Inventory Analytics Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your branch inventory workbook to view dynamically cleaned metrics, summaries, and automated PDF export profiles.</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose your inventory Excel file (.xlsx)", type=["xlsx"])

def clean_numeric(val):
    return pd.to_numeric(val, errors='coerce')

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        pdf_report_payload = {}

        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 CCK Tablet", 
            "🪦 CCK Niche", 
            "🏛️ LST Tablet & Niche", 
            "🌅 TLT Tablet & Niche"
        ])

        # ==========================================
        # TAB 1: CCK-TABLET (Exact Formula Target)
        # ==========================================
        with tab1:
            st.markdown('<div class="section-header">CCK Tablet Summary Matrix</div>', unsafe_allow_html=True)
            df = pd.read_excel(uploaded_file, sheet_name='CCK-TABLET', header=1)
            df.columns = df.columns.str.strip()
            
            # Extract precise subtotal rows directly from Excel formulas
            subtotals = df[df['BLOCK'].astype(str).str.upper().str.contains('SUB TOTAL')].copy()
            grand_total = df[df['BLOCK'].astype(str).str.upper() == 'TABLET TOTAL:'].copy()
            
            rows_list = []
            for _, r in subtotals.iterrows():
                label = "Blk " + str(r['BLOCK']).split(' ')[1]  # Extracts 'A', 'B', 'C'
                tot = clean_numeric(r['TOTAL'])
                sold = clean_numeric(r['TOTAL SOLD'])
                bal = clean_numeric(r['BALANCE'])
                val_bal = clean_numeric(r['Value of Balance units'])
                rows_list.append({
                    ' ': label, 'Sold': sold/tot if tot > 0 else 0, 'Balance': bal/tot if tot > 0 else 0,
                    'Balance unit': bal, 'Value of balance': val_bal
                })
                
            for _, r in grand_total.iterrows():
                tot = clean_numeric(r['TOTAL'])
                sold = clean_numeric(r['TOTAL SOLD'])
                bal = clean_numeric(r['BALANCE'])
                val_bal = clean_numeric(r['Value of Balance units'])
                rows_list.append({
                    ' ': 'Total', 'Sold': sold/tot if tot > 0 else 0, 'Balance': bal/tot if tot > 0 else 0,
                    'Balance unit': bal, 'Value of balance': val_bal
                })
                
            pivot_t1_final = pd.DataFrame(rows_list)
            st.dataframe(pivot_t1_final.style.format({
                'Sold': '{:.2%}', 'Balance': '{:.2%}', 'Balance unit': '{:,.0f}', 'Value of balance': '$  {:,.2f}'
            }), use_container_width=True)
            pdf_report_payload['CCK-TABLET'] = pivot_t1_final

        # ==========================================
        # TAB 2: CCK-NICHE
        # ==========================================
        with tab2:
            st.markdown('<div class="section-header">CCK Niche Summary Matrix</div>', unsafe_allow_html=True)
            df = pd.read_excel(uploaded_file, sheet_name='CCK-NICHE', header=1)
            df.columns = df.columns.str.strip()
            
            subtotals = df[df['BLOCK'].astype(str).str.upper().str.contains('SUB TOTAL')].copy()
            grand_total = df[df['BLOCK'].astype(str).str.upper() == 'NICHE TOTAL:'].copy()
            
            rows_list = []
            for _, r in subtotals.iterrows():
                label = "Blk " + str(r['BLOCK']).split(' ')[1]
                tot = clean_numeric(r['TOTAL'])
                sold = clean_numeric(r['TOTAL SOLD'])
                bal = clean_numeric(r['BALANCE'])
                val_bal = clean_numeric(r['Value of Balance'])
                rows_list.append({
                    ' ': label, 'Sold': sold/tot if tot > 0 else 0, 'Balance': bal/tot if tot > 0 else 0,
                    'Balance unit': bal, 'Value of balance': val_bal
                })
                
            for _, r in grand_total.iterrows():
                tot = clean_numeric(r['TOTAL'])
                sold = clean_numeric(r['TOTAL SOLD'])
                bal = clean_numeric(r['BALANCE'])
                val_bal = clean_numeric(r['Value of Balance'])
                rows_list.append({
                    ' ': 'Total', 'Sold': sold/tot if tot > 0 else 0, 'Balance': bal/tot if tot > 0 else 0,
                    'Balance unit': bal, 'Value of balance': val_bal
                })
                
            pivot_t2_final = pd.DataFrame(rows_list)
            st.dataframe(pivot_t2_final.style.format({
                'Sold': '{:.2%}', 'Balance': '{:.2%}', 'Balance unit': '{:,.0f}', 'Value of balance': '$  {:,.2f}'
            }), use_container_width=True)
            pdf_report_payload['CCK-NICHE'] = pivot_t2_final

        # ==========================================
        # TAB 3: LST-TABLE & NICHE
        # ==========================================
        with tab3:
            st.markdown('<div class="section-header">LST Tablet and Niche Summary Matrix</div>', unsafe_allow_html=True)
            df = pd.read_excel(uploaded_file, sheet_name='LST-TABLE & NICHE', header=1)
            df.columns = df.columns.str.strip()
            
            targets = ['TABLET TOTAL:', 'NICHE TOTAL:', 'ALL PRODUCT TOTAL:']
            filtered_rows = df[df['PRODUCT'].astype(str).str.upper().isin(targets)].copy()
            
            rows_list = []
            for _, r in filtered_rows.iterrows():
                p_label = str(r['PRODUCT']).replace('TOTAL:', '').strip().title()
                label = "Total" if "All Product" in p_label else p_label
                
                tot = clean_numeric(r['TOTAL'])
                sold = clean_numeric(r['TOTAL SOLD'])
                bal = clean_numeric(r['BALANCE'])
                val_bal = clean_numeric(r["BALANCE $ '000 (PO)"]) * 1000  # Convert to full currency value
                
                rows_list.append({
                    ' ': label, 'Sold': sold/tot if tot > 0 else 0, 'Balance': bal/tot if tot > 0 else 0,
                    'Balance unit': bal, 'Value of balance': val_bal
                })
                
            pivot_t3_final = pd.DataFrame(rows_list)
            st.dataframe(pivot_t3_final.style.format({
                'Sold': '{:.2%}', 'Balance': '{:.2%}', 'Balance unit': '{:,.0f}', 'Value of balance': '$  {:,.2f}'
            }), use_container_width=True)
            pdf_report_payload['LST-SUMMARY'] = pivot_t3_final

        # ==========================================
        # TAB 4: TLT-TABLE & NICHE
        # ==========================================
        with tab4:
            st.markdown('<div class="section-header">TLT Tablet & Niche Summary Matrix</div>', unsafe_allow_html=True)
            df = pd.read_excel(uploaded_file, sheet_name='TLT-TABLE & NICHE', header=1)
            df.columns = df.columns.str.strip()
            
            targets = ['TABLET TOTAL:', 'NICHE TOTAL:', 'ALL PRODUCT TOTAL:']
            filtered_rows = df[df['PRODUCT'].astype(str).str.upper().isin(targets)].copy()
            
            rows_list = []
            for _, r in filtered_rows.iterrows():
                p_label = str(r['PRODUCT']).replace('TOTAL:', '').strip().title()
                label = "Total" if "All Product" in p_label else p_label
                
                tot = clean_numeric(r['TOTAL'])
                sold = clean_numeric(r['TOTAL SOLD'])
                bal = clean_numeric(r['BALANCE'])
                val_bal = clean_numeric(r["BALANCE $ '000 (PO)"]) * 1000  # Convert to full currency value
                
                rows_list.append({
                    ' ': label, 'Sold': sold/tot if tot > 0 else 0, 'Balance': bal/tot if tot > 0 else 0,
                    'Balance unit': bal, 'Value of balance': val_bal
                })
                
            pivot_t4_final = pd.DataFrame(rows_list)
            st.dataframe(pivot_t4_final.style.format({
                'Sold': '{:.2%}', 'Balance': '{:.2%}', 'Balance unit': '{:,.0f}', 'Value of balance': '$  {:,.2f}'
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
                    if col in ['Sold', 'Balance']:
                        formatted_df[col] = formatted_df[col].apply(lambda x: f"{x:.2%}")
                    elif col == 'Balance unit':
                        formatted_df[col] = formatted_df[col].apply(lambda x: f"{x:,.0f}")
                    elif col == 'Value of balance':
                        formatted_df[col] = formatted_df[col].apply(lambda x: f"${x:,.2f}")
                
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
