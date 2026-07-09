# ==========================================
        # TAB 1: CCK-TABLET (Exact Image Match)
        # ==========================================
        with tab1:
            st.markdown('<div class="section-header">CCK Tablet Summary Matrix</div>', unsafe_allow_html=True)
            df_t1_raw = pd.read_excel(uploaded_file, sheet_name='CCK-TABLET', header=1)
            df_t1_raw.columns = df_t1_raw.columns.str.strip()
            
            if 'BLOCK' in df_t1_raw.columns:
                # Filter out any built-in Excel subtotal/total rows
                df_t1 = df_t1_raw[df_t1_raw['BLOCK'].notna() & (~df_t1_raw['BLOCK'].astype(str).str.upper().str.contains('TOTAL'))].copy()
                
                # Numeric conversions
                df_t1['TOTAL'] = clean_numeric(df_t1['TOTAL'])
                df_t1['TOTAL SOLD'] = clean_numeric(df_t1['TOTAL SOLD'])
                df_t1['BALANCE'] = clean_numeric(df_t1['BALANCE'])
                df_t1['Value of Balance'] = df_t1['BALANCE'] * clean_numeric(df_t1['AVG PO PRICE'])
                
                # Grouping by Block (A, B, C)
                pivot_t1 = df_t1.groupby('BLOCK').agg({
                    'TOTAL': 'sum', 
                    'TOTAL SOLD': 'sum', 
                    'BALANCE': 'sum', 
                    'Value of Balance': 'sum'
                }).reset_index()
                
                # Match your exact percentage column logic
                pivot_t1['Sold'] = pivot_t1['TOTAL SOLD'] / pivot_t1['TOTAL']
                pivot_t1['Balance'] = pivot_t1['BALANCE'] / pivot_t1['TOTAL']
                
                # Match your exact row naming: "Blk A", "Blk B", etc.
                pivot_t1['BLOCK'] = pivot_t1['BLOCK'].apply(lambda x: f"Blk {x}" if not str(x).startswith("Blk") else x)
                
                # Calculate the precise "Total" row metrics
                total_sum = pivot_t1['TOTAL'].sum()
                sold_sum = pivot_t1['TOTAL SOLD'].sum()
                bal_sum = pivot_t1['BALANCE'].sum()
                
                total_row = pd.DataFrame([{
                    'BLOCK': 'Total',
                    'Sold': sold_sum / total_sum if total_sum > 0 else 0,
                    'Balance': bal_sum / total_sum if total_sum > 0 else 0,
                    'BALANCE': bal_sum,
                    'Value of Balance': pivot_t1['Value of Balance'].sum()
                }])
                
                # Format to your exact layout: Block, Sold, Balance, Balance unit, Value of balance
                pivot_t1_final = pd.concat([pivot_t1, total_row], ignore_index=True)
                pivot_t1_final = pivot_t1_final.rename(columns={
                    'BLOCK': ' ',  # Keeps it blank or use 'Block' to match your left column context
                    'BALANCE': 'Balance unit',
                    'Value of Balance': 'Value of balance'
                })[[ ' ', 'Sold', 'Balance', 'Balance unit', 'Value of balance']]
                
                # Output exactly formatted matrix table
                st.dataframe(pivot_t1_final.style.format({
                    'Sold': '{:.2%}', 
                    'Balance': '{:.2%}', 
                    'Balance unit': '{:,.0f}', 
                    'Value of balance': '$  {:,.2f}'
                }), use_container_width=True)
                
                pdf_report_payload['CCK-TABLET'] = pivot_t1_final
