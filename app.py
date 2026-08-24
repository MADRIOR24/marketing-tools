import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="The Marketing Suite", page_icon="📊", layout="wide")

# --- APP NAVIGATION ---
st.sidebar.title("Navigation")
app_mode = st.sidebar.selectbox("Choose a Tool", ["The Survey", "The Segment Creation"])

# --- SHARED CONSTANTS FOR SEGMENT CREATION ---
BIENNIAL_SHOWS = ["OE", "OI", "SIFER", "FFITA", "FFS", "EFIT", "EUROB", "ICEEE", "INTAI", "EQUDE"]
THREE_YEAR_LABELS = [
    "leads_all_channels", "MSH", "sharers", "Last 3 years exhbiting delegates", 
    "last 2 years registered_PEC", "last 3 years registered VIP", "last 3 years registered", 
    "leads", "Last 3 years attended", "Last 3 years attended_VIP", "Bounceback", 
    "Last 3 years attended Buyer", "last 3 years registered Buyer"
]
MINUS_ONE_LABELS = [
    "Visit to Exhibit Attendee", "Visit to Exhibit non Attendee", "Last year registered VIP", 
    "Last year registered", "Last year attended", "Last year attended_VIP", 
    "External Abandon Basket", "Last year attended Buyer", "Last year registered Buyer"
]

# --- TOOL 1: THE SURVEY ---
if app_mode == "The Survey":
    st.title("📊 The Survey")
    st.write("Extract participant unique codes from Qualtrics URLs.")
    
    uploaded_file = st.file_uploader("Upload Survey Excel File", type=["xlsx"])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        possible_cols = ['Participant_URL', 'URL', 'url', 'Participant URL']
        found_url_col = next((c for c in possible_cols if c in df.columns), None)
        
        if found_url_col:
            def extract_code(url):
                u = str(url)
                if 'Q_DL=' not in u: return ""
                dl = u.split('Q_DL=')[-1].split('&')[0]
                parts = dl.split('_')
                return "_".join(parts[2:]) if len(parts) > 2 else dl

            df['Participant_Unique_Code'] = df[found_url_col].apply(extract_code)
            
            # Reorder
            cols = list(df.columns)
            cols.remove('Participant_Unique_Code')
            idx = cols.index(found_url_col)
            cols.insert(idx + 1, 'Participant_Unique_Code')
            df = df[cols]
            
            st.success("Processing complete!")
            st.dataframe(df.head())
            
            output = io.BytesIO()
            df.to_excel(output, index=False)
            st.download_button("Download Processed File", data=output.getvalue(), file_name="Survey_Processed.xlsx")
        else:
            st.error("Could not find a URL column.")

# --- TOOL 2: THE SEGMENT CREATION ---
elif app_mode == "The Segment Creation":
    st.title("📁 The Segment Creation")
    st.write("Automate Alpha Code fill-down and Criteria logic.")
    
    uploaded_file = st.file_uploader("Upload Customer Journey Template", type=["xlsx"])
    
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file)
        sheet = wb.active
        
        start_row = 4
        m_alpha = str(sheet.cell(row=start_row, column=2).value or "").strip()
        m_year_str = str(sheet.cell(row=start_row, column=3).value or "").strip()
        
        if m_alpha and m_year_str:
            jump = 2 if m_alpha.upper() in BIENNIAL_SHOWS else 1
            st.info(f"Detected {'Biennial' if jump==2 else 'Annual'} show logic for {m_alpha}.")
            
            curr_yr = int(m_year_str)
            prev_yr = str(curr_yr - jump)
            y1, y2, y3 = str(curr_yr-(jump*3))[-2:], str(curr_yr-(jump*2))[-2:], str(curr_yr-jump)[-2:]
            three_yr_str = f"{y1}, {y2}, {y3}"

            row = start_row
            while sheet.cell(row=row, column=4).value:
                label = str(sheet.cell(row=row, column=4).value).strip()
                
                sheet.cell(row=row, column=2).value = m_alpha
                sheet.cell(row=row, column=3).value = m_year_str
                
                # Segment Name
                clean_lbl = label.replace(" ", "_")
                sheet.cell(row=row, column=7).value = f"{m_alpha}{m_year_str}_{sheet.cell(row=row, column=6).value}_{clean_lbl}"
                
                # Criteria Logic
                crit = sheet.cell(row=row, column=9).value
                if crit:
                    lines = str(crit).split('\n')
                    new_lines = []
                    is_3 = any(l.lower() == label.lower() for l in THREE_YEAR_LABELS)
                    is_last = any(l.lower() == label.lower() for l in MINUS_ONE_LABELS)
                    
                    for line in lines:
                        u_line = line.strip().upper()
                        if u_line.startswith("SHOW ="): new_lines.append(f"SHOW = {m_alpha}")
                        elif u_line.startswith("YEARS ="):
                            if is_3: new_lines.append(f"YEARS = {three_yr_str}")
                            elif is_last: new_lines.append(f"YEARS = {prev_yr}")
                            else: new_lines.append(f"YEARS = {m_year_str}")
                        else: new_lines.append(line)
                    sheet.cell(row=row, column=9).value = "\n".join(new_lines)
                row += 1
            
            output = io.BytesIO()
            wb.save(output)
            st.success(f"Processed {row-start_row} segments!")
            st.download_button("Download Generated Segments", data=output.getvalue(), file_name=f"Generated_{m_alpha}_{m_year_str}.xlsx")
        else:
            st.error("Ensure Alpha Code and Year are in Row 4.")