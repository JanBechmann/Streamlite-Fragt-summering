import streamlit as st
import pandas as pd
import openpyxl  #

# Funktion til at processere Excel-fil
def process_excel(file):
    xls = pd.ExcelFile(file)
    data_list = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        if df.shape[1] > 15:  # Tjek for minimum kolonner
            receiver_countries = df.iloc[:, 8]  # Kolonne 9 (nulindekseret)
            amounts = df.iloc[:, 15]  # Kolonne 16 (nulindekseret)
            
            filtered_data = pd.DataFrame({"Receiver Country": receiver_countries, "Amount": amounts})
            filtered_data_valid = filtered_data[filtered_data["Receiver Country"].astype(str).str.match(r"^[A-Z]{2}$")]
            data_list.append(filtered_data_valid)
    
    combined_df = pd.concat(data_list, ignore_index=True)
    combined_df = combined_df[pd.to_numeric(combined_df["Amount"], errors="coerce").notna()]
    combined_df["Amount"] = pd.to_numeric(combined_df["Amount"])

    summary_df = combined_df.groupby("Receiver Country", as_index=False)["Amount"].sum()
    
    # Håndter manglende landekoder
    missing_country_df = combined_df[combined_df["Receiver Country"].isna() | 
                                     (combined_df["Receiver Country"].astype(str).str.strip() == "")]
    missing_country_sum = missing_country_df["Amount"].sum().round(2)
    
    if missing_country_sum > 0:
        missing_country_row = pd.DataFrame([{"Receiver Country": "", "Amount": missing_country_sum}])
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
