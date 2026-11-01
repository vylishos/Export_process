import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime

st.set_page_config(page_title="Export processor", layout="wide")

def apply_sklad_formatting(writer, df_sorted):
    worksheet = writer.sheets['Sklad']
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    thick_bottom_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thick'))
    header_font = Font(bold=True)
    date_font = Font(bold=True, size=11)

    worksheet['A1'].font = date_font

    # Formátování 3 sloupců skladu (A=Reference, B=Varianta, C=Ks)
    for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=3)):
        for cell in row:
            if row_idx == 0:
                cell.font = header_font
                cell.border = thick_bottom_border
            else:
                cell.border = thin_border

    for col in ['A', 'B', 'C']:
        worksheet.column_dimensions[col].width = 18

def apply_prehled_formatting(writer, df_vystup):
    worksheet = writer.sheets['Prehled']
    max_row = worksheet.max_row
    max_col = worksheet.max_column
    thin_side = Side(style='thin')
    thick_side = Side(style='thick')
    top_border_only = Border(top=thin_side)
    header_border = Border(bottom=thick_side)
    header_font = Font(bold=True)
    date_font = Font(bold=True, size=11)
    no_border = Border()
    align_right = Alignment(horizontal='right')

    worksheet['A1'].font = date_font

    for r_idx in range(2, max_row + 1):
        for c_idx in range(1, max_col + 1):
            cell = worksheet.cell(row=r_idx, column=c_idx)
            
            if c_idx == 2:
                cell.alignment = align_right

            if r_idx == 2:
                cell.font = header_font
                cell.border = header_border
            else:
                val_b = worksheet.cell(row=r_idx, column=2).value
                is_new_order = val_b is not None and str(val_b).strip() != "" and str(val_b).strip().lower() != "nan"
                cell.border = top_border_only if is_new_order else no_border

    for i, col_name in enumerate(df_vystup.columns):
        column_data = df_vystup[col_name].astype(str).fillna('')
        clean_lengths = column_data.replace(['nan', 'None'], '').apply(len)
        
        max_content_len = clean_lengths.max() if not clean_lengths.empty else 0
        max_len = max(max_content_len, len(str(col_name))) + 2
        
        col_letter = chr(65 + i) if i < 26 else f"A{chr(65 + (i-26))}"
        worksheet.column_dimensions[col_letter].width = max_len

def process_sklad_dataframe(df_items):
    """Pomocná funkce pro zpracování a seřazení skladových položek (bez sloupce Název)"""
    if df_items.empty:
        return pd.DataFrame(columns=['Reference', 'Varianta', 'Ks'])
    
    df_res = df_items.copy()
    # Původní vstupy byly iloc[:, [25, 28, 26, 30]] -> indexy: 0: Název, 1: Reference, 2: Varianta, 3: Ks
    # Ponecháme pouze [1, 2, 3] -> Reference, Varianta, Ks
    if df_res.shape[1] == 4:
        df_res = df_res.iloc[:, [1, 2, 3]]
        
    df_res.columns = ['Reference', 'Varianta', 'Ks']
    df_res = df_res.dropna(subset=['Reference']).copy()
    df_res['Ks'] = pd.to_numeric(df_res['Ks'], errors='coerce').fillna(0).astype(int)
    df_res['Varianta'] = df_res['Varianta'].astype(str).replace('nan', '')
    
    # Agregace stejných položek (sečtení Ks dle Reference a Varianty)
    df_res = df_res[df_res['Ks'] > 0].groupby(['Reference', 'Varianta'], as_index=False)['Ks'].sum()
    return df_res.sort_values(by=['Varianta'])

st.title("Export processor")
uploaded_file = st.file_uploader("Nahrajte exportní Excel soubor (.xlsx)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.success("Soubor úspěšně nahrán!")

    datum_soubor = datetime.now().strftime("%d.%m.")
    datum_text = datetime.now().strftime("%d.%m.%Y")

    # --- PŘED-ZPRACOVÁNÍ OBJEDNÁVEK A DOPRAVCŮ PRO OBĚ ČÁSTI ---
    is_main_order = df.iloc[:, 0].notna()
    is_item = df.iloc[:, 28].notna() & (df.iloc[:, 28].astype(str).str.strip() != '')
    
    vystup_rows = []
    zasilkovna_rows = []
    gls_rows = []
    ups_rows = []
    ppl_rows = []

    # Pro skladové seznamy dle dopravců
    sklad_items_by_carrier = {
        'UPS': [],
        'GLS': [],
        'PPL': [],
        'Zásilkovna': []
    }
    
    current_order_rows = []
    current_order_items_raw = []
    current_extra_texts = []
    
    for idx, row in df.iterrows():
        if is_main_order[idx]:
            if current_order_rows:
                combined_text = " ".join(current_extra_texts)
                if "UPS" in combined_text:
                    cisty_dopravce = "UPS"
                elif "GLS" in combined_text:
                    cisty_dopravce = "GLS"
                elif any(kwd in combined_text for kwd in ["PPL", "DHL", "Colis Privé", "Posteitaliane"]):
                    cisty_dopravce = "PPL"
                else:
                    cisty_dopravce = "Zásilkovna"
                
                current_order_rows[0]['Dopravce'] = cisty_dopravce
                current_order_rows.append({
                    'Dopravce': None, 'Číslo objednávky': None, 'Jméno': None,
                    'Reference': None, 'Varianta': None, 'Ks': None
                })
                
                vystup_rows.extend(current_order_rows)
                if cisty_dopravce == "UPS":
                    ups_rows.extend(current_order_rows)
                elif cisty_dopravce == "GLS":
                    gls_rows.extend(current_order_rows)
                elif cisty_dopravce == "PPL":
                    ppl_rows.extend(current_order_rows)
                else:
                    zasilkovna_rows.extend(current_order_rows)

                sklad_items_by_carrier[cisty_dopravce].extend(current_order_items_raw)

            current_order_rows = [{
                'Dopravce': None,
                'Číslo objednávky': row.iloc[0],
                'Jméno': row.iloc[2],
                'Reference': None,
                'Varianta': None,
                'Ks': None
            }]
            current_order_items_raw = []
            current_extra_texts = []
            
        elif is_item[idx]:
            current_order_rows.append({
                'Dopravce': None,
                'Číslo objednávky': row.iloc[0],
                'Jméno': row.iloc[2],
                'Reference': row.iloc[28],
                'Varianta': row.iloc[26],
                'Ks': row.iloc[30]
            })
            current_order_items_raw.append(row.iloc[[25, 28, 26, 30]])
        else:
            puvodni_text = str(row.iloc[25]) if pd.notna(row.iloc[25]) else ""
            if puvodni_text:
                current_extra_texts.append(puvodni_text)

    # Poslední objednávka
    if current_order_rows:
        combined_text = " ".join(current_extra_texts)
        if "UPS" in combined_text:
            cisty_dopravce = "UPS"
        elif "GLS" in combined_text:
            cisty_dopravce = "GLS"
        elif any(kwd in combined_text for kwd in ["PPL", "DHL", "Colis Privé", "Posteitaliane"]):
            cisty_dopravce = "PPL"
        else:
            cisty_dopravce = "Zásilkovna"
        
        current_order_rows[0]['Dopravce'] = cisty_dopravce
        current_order_rows.append({
            'Dopravce': None, 'Číslo objednávky': None, 'Jméno': None,
            'Reference': None, 'Varianta': None, 'Ks': None
        })
        
        vystup_rows.extend(current_order_rows)
        if cisty_dopravce == "UPS":
            ups_rows.extend(current_order_rows)
        elif cisty_dopravce == "GLS":
            gls_rows.extend(current_order_rows)
        elif cisty_dopravce == "PPL":
            ppl_rows.extend(current_order_rows)
        else:
            zasilkovna_rows.extend(current_order_rows)

        sklad_items_by_carrier[cisty_dopravce].extend(current_order_items_raw)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Skladový seznam")
        try:
            # --- 1A. KOMPLETNÍ SKLADOVÝ SEZNAM ---
            df1_raw = df.iloc[:, [25, 28, 26, 30]].copy()
            df1 = process_sklad_dataframe(df1_raw)
            
            output1 = io.BytesIO()
            with pd.ExcelWriter(output1, engine='openpyxl') as writer:
                df1.to_excel(writer, index=False, sheet_name='Sklad', startrow=1)
                worksheet = writer.sheets['Sklad']
                worksheet['A1'] = f"Export ze dne: {datum_text}"
                apply_sklad_formatting(writer, df1)
            
            st.download_button(
                label="Stáhnout Kompletní skladový seznam", 
                data=output1.getvalue(), 
                file_name=f"seznam_pro_sklad_{datum_soubor}xlsx"
            )

            st.markdown("---")
            st.markdown("### ⏱️Rozdělení skladu při zahlcení")
            
            order_cumsum = df.iloc[:, 0].notna().cumsum()
            
            # --- 1B. SKLADOVÝ SEZNAM (PRVNÍCH 30 OBJEDNÁVEK) ---
            df_pouze_30_raw = df[order_cumsum <= 30].iloc[:, [25, 28, 26, 30]]
            df1_30 = process_sklad_dataframe(df_pouze_30_raw)
            
            output1_30 = io.BytesIO()
            with pd.ExcelWriter(output1_30, engine='openpyxl') as writer:
                df1_30.to_excel(writer, index=False, sheet_name='Sklad', startrow=1)
                worksheet = writer.sheets['Sklad']
                worksheet['A1'] = f"Export (Prvních 30 obj.) ze dne: {datum_text}"
                apply_sklad_formatting(writer, df1_30)
                
            st.download_button(
                label="Seznam pro prvních 30 objednávek", 
                data=output1_30.getvalue(), 
                file_name=f"seznam_pro_sklad_prvni_30_{datum_soubor}xlsx"
            )
            
            # --- 1C. SKLADOVÝ SEZNAM (OD 31. OBJEDNÁVKY DÁL) ---
            df_zbytek_raw = df[order_cumsum > 30].iloc[:, [25, 28, 26, 30]]
            
            if not df_zbytek_raw.empty:
                df1_zbytek = process_sklad_dataframe(df_zbytek_raw)
                
                output1_zbytek = io.BytesIO()
                with pd.ExcelWriter(output1_zbytek, engine='openpyxl') as writer:
                    df1_zbytek.to_excel(writer, index=False, sheet_name='Sklad', startrow=1)
                    worksheet = writer.sheets['Sklad']
                    worksheet['A1'] = f"Export (Od 31. obj. dál) ze dne: {datum_text}"
                    apply_sklad_formatting(writer, df1_zbytek)
                    
                st.download_button(
                    label="Seznam pro zbytek objednávek", 
                    data=output1_zbytek.getvalue(), 
                    file_name=f"seznam_pro_sklad_od_31_{datum_soubor}xlsx"
                )
            else:
                st.info("V exportu není více než 30 objednávek (druhá část seznamu je prázdná).")

            # --- 1D. SKLADOVÉ SEZNAMY DLE DOPRAVCŮ ---
            st.markdown("---")
            st.markdown("### 🚚 Skladové seznamy podle dopravců")

            for carrier in ['GLS', 'PPL', 'Zásilkovna', 'UPS']:
                items_list = sklad_items_by_carrier[carrier]
                if items_list:
                    df_carrier_raw = pd.DataFrame(items_list)
                    df_carrier_sklad = process_sklad_dataframe(df_carrier_raw)
                    
                    if not df_carrier_sklad.empty:
                        out_carrier = io.BytesIO()
                        with pd.ExcelWriter(out_carrier, engine='openpyxl') as writer:
                            df_carrier_sklad.to_excel(writer, index=False, sheet_name='Sklad', startrow=1)
                            worksheet = writer.sheets['Sklad']
                            worksheet['A1'] = f"Export Sklad {carrier} ze dne: {datum_text}"
                            apply_sklad_formatting(writer, df_carrier_sklad)
                            
                        st.download_button(
                            label=f"Skladový seznam - {carrier}", 
                            data=out_carrier.getvalue(), 
                            file_name=f"seznam_pro_sklad_{carrier}_{datum_soubor}xlsx"
                        )
            
        except Exception as e:
            st.error(f"Chyba při tvorbě skladu: {e}")

    with col2:
        st.subheader("Přehled objednávek")
        try:
            prehled_columns = ['Dopravce', 'Číslo objednávky', 'Jméno', 'Reference', 'Varianta', 'Ks']
            
            df_vystup = pd.DataFrame(vystup_rows, columns=prehled_columns)
            output2 = io.BytesIO()
            with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                df_vystup.to_excel(writer, index=False, sheet_name='Prehled', startrow=1)
                worksheet = writer.sheets['Prehled']
                worksheet['A1'] = f"Export ze dne: {datum_text}"
                apply_prehled_formatting(writer, df_vystup)
            
            st.download_button(
                label="Stáhnout Kompletní přehled", 
                data=output2.getvalue(), 
                file_name=f"Prehled_objednavek_{datum_soubor}xlsx"
            )
            
            st.markdown("---")
            st.markdown("### 📦 Objednávky podle dopravců")
            
            if zasilkovna_rows:
                df_zasilkovna = pd.DataFrame(zasilkovna_rows, columns=prehled_columns)
                output_z = io.BytesIO()
                with pd.ExcelWriter(output_z, engine='openpyxl') as writer:
                    df_zasilkovna.to_excel(writer, index=False, sheet_name='Prehled', startrow=1)
                    worksheet = writer.sheets['Prehled']
                    worksheet['A1'] = f"Export Zásilkovna ze dne: {datum_text}"
                    apply_prehled_formatting(writer, df_zasilkovna)
                st.download_button(
                    label="Stáhnout objednávky - Zásilkovna", 
                    data=output_z.getvalue(), 
                    file_name=f"Prehled_Zasilkovna_{datum_soubor}xlsx"
                )
                
            if gls_rows:
                df_gls = pd.DataFrame(gls_rows, columns=prehled_columns)
                output_g = io.BytesIO()
                with pd.ExcelWriter(output_g, engine='openpyxl') as writer:
                    df_gls.to_excel(writer, index=False, sheet_name='Prehled', startrow=1)
                    worksheet = writer.sheets['Prehled']
                    worksheet['A1'] = f"Export GLS ze dne: {datum_text}"
                    apply_prehled_formatting(writer, df_gls)
                st.download_button(
                    label="Stáhnout objednávky - GLS", 
                    data=output_g.getvalue(), 
                    file_name=f"Prehled_GLS_{datum_soubor}xlsx"
                )
                
            if ups_rows:
                df_ups = pd.DataFrame(ups_rows, columns=prehled_columns)
                output_u = io.BytesIO()
                with pd.ExcelWriter(output_u, engine='openpyxl') as writer:
                    df_ups.to_excel(writer, index=False, sheet_name='Prehled', startrow=1)
                    worksheet = writer.sheets['Prehled']
                    worksheet['A1'] = f"Export UPS ze dne: {datum_text}"
                    apply_prehled_formatting(writer, df_ups)
                st.download_button(
                    label="Stáhnout objednávky - UPS", 
                    data=output_u.getvalue(), 
                    file_name=f"Prehled_UPS_{datum_soubor}xlsx"
                )

            if ppl_rows:
                df_ppl = pd.DataFrame(ppl_rows, columns=prehled_columns)
                output_p = io.BytesIO()
                with pd.ExcelWriter(output_p, engine='openpyxl') as writer:
                    df_ppl.to_excel(writer, index=False, sheet_name='Prehled', startrow=1)
                    worksheet = writer.sheets['Prehled']
                    worksheet['A1'] = f"Export PPL ze dne: {datum_text}"
                    apply_prehled_formatting(writer, df_ppl)
                st.download_button(
                    label="Stáhnout objednávky - PPL", 
                    data=output_p.getvalue(), 
                    file_name=f"Prehled_PPL_{datum_soubor}xlsx"
                )
                
        except Exception as e:
            st.error(f"Chyba při tvorbě přehledů: {e}")