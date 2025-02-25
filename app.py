import streamlit as st
import pandas as pd
import openpyxl  # Sørg for at openpyxl er importeret

# Funktion til at processere Excel-fil
def process_excel(file):
    xls = pd.ExcelFile(file, engine="openpyxl")  # Angiv openpyxl som engine
    data_list = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine="openpyxl")  # Angiv engine her også
        
        if df.shape[1] > 16:  # Tjek for minimum kolonner for generelle data
            receiver_countries = df.iloc[:, 8]  # Kolonne 9 (nulindekseret)
            amounts = df.iloc[:, 15]  # Kolonne 16 (nulindekseret)
            
            additional_amounts = df.iloc[:, 16] if sheet_name == "UPS DE" else 0  # Kolonne 17 (nulindekseret) kun for UPS DE
            
            total_amounts = amounts + additional_amounts  # Samlet beløb
            
            filtered_data = pd.DataFrame({"Receiver Country": receiver_countries, "Amount": total_amounts})
            filtered_data_valid = filtered_data[filtered_data["Receiver Country"].astype(str).str.match(r"^[A-Z]{2}$")]
            data_list.append(filtered_data_valid)
    
    combined_df = pd.concat(data_list, ignore_index=True)
    combined_df = combined_df[pd.to_numeric(combined_df["Amount"], errors="coerce").notna()]
    combined_df["Amount"] = pd.to_numeric(combined_df["Amount"], errors="coerce").round(2)

    summary_df = combined_df.groupby("Receiver Country", as_index=False)["Amount"].sum()
    summary_df["Amount"] = summary_df["Amount"].round(2)
    
    # Håndter manglende landekoder
    missing_country_df = combined_df[combined_df["Receiver Country"].isna() | 
                                     (combined_df["Receiver Country"].astype(str).str.strip() == "")]
    missing_country_sum = missing_country_df["Amount"].sum().round(2)
    
    if missing_country_sum > 0:
        missing_country_row = pd.DataFrame([{ "Receiver Country": "", "Amount": missing_country_sum }])
        summary_df = pd.concat([summary_df, missing_country_row], ignore_index=True)
    
    total_sum = summary_df["Amount"].sum().round(2)
    total_row = pd.DataFrame([{"Receiver Country": "Total", "Amount": total_sum}])
    summary_df = pd.concat([summary_df, total_row], ignore_index=True)
    
    return summary_df

# Streamlit UI
st.title("Fragtbeløb Summering App")
st.write("Upload en Excel-fil, og få summering af fragtbeløb per land.")

uploaded_file = st.file_uploader("Upload Excel-fil", type=["xlsx"])

if uploaded_file:
    summary = process_excel(uploaded_file)
    st.write("### Summering af fragtbeløb per land")
    st.dataframe(summary)
    
    # Download-knap til resultater
    csv = summary.to_csv(index=False).encode('utf-8')
    st.download_button("Download CSV", csv, "fragt_summering.csv", "text/csv")
