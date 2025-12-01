
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# App title
st.title("KWh Consumption Calculator")

st.write("""
Upload an Excel file with a sheet named **'duomenys_analizei'**.
The app will sum the **Skirtumas** column per **Obj. Nr.**, show results, and let you download a CSV.
Numbers are formatted with comma as decimal separator and space as thousands separator.
""")

# File upload
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Validate sheet name
        sheet_name = "duomenys_analizei"
        xls = pd.ExcelFile(uploaded_file)
        if sheet_name not in xls.sheet_names:
            st.error(f"Sheet '{sheet_name}' not found. Available sheets: {xls.sheet_names}")
        else:
            # Read the sheet
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl")

            # Ensure Skirtumas column is numeric
            df["Skirtumas"] = pd.to_numeric(df["Skirtumas"], errors="coerce")

            # Group by Obj. Nr. and sum Skirtumas
            summary = df.groupby("Obj. Nr.")["Skirtumas"].sum().reset_index()
            summary.rename(columns={"Skirtumas": "kWh_suvartota"}, inplace=True)

            # Sort by consumption descending
            summary = summary.sort_values(by="kWh_suvartota", ascending=False)

            # Format numbers for display (space for thousands, comma for decimals)
            summary["kWh_suvartota_display"] = summary["kWh_suvartota"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
            )

            # Show table with formatted numbers
            st.subheader("Results")
            st.dataframe(summary[["Obj. Nr.", "kWh_suvartota_display"]].rename(
                columns={"kWh_suvartota_display": "kWh_suvartota"}
            ))

            # Prepare CSV (numeric values, semicolon delimiter)
            csv = summary[["Obj. Nr.", "kWh_suvartota"]].to_csv(index=False, sep=";", encoding="utf-8")
            st.download_button(
                label="Download Full CSV",
                data=csv,
                file_name="obj_nr_kWh_consumption_all_rows.csv",
                mime="text/csv"
            )

            # Plot top 10 objects
            st.subheader("Top 10 Objects by kWh Consumption")
            top10 = summary.head(10)
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.bar(top10["Obj. Nr."], top10["kWh_suvartota"], color="steelblue")
            ax.set_title("Top 10 Objects by kWh Consumption")
            ax.set_xlabel("Obj. Nr.")
            ax.set_ylabel("kWh")
            plt.xticks(rotation=45)
            st.pyplot(fig)

    except Exception as e:
        st.error(f"Error: {e}")
