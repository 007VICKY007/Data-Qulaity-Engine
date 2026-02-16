import streamlit as st
import pandas as pd
from pyxlsb import open_workbook
from io import BytesIO

st.set_page_config(page_title="XLSB Converter", layout="centered")

st.title("📊 XLSB → XLS / CSV Converter")

uploaded_file = st.file_uploader(
    "Upload XLSB file",
    type=["xlsb"]
)

output_format = st.selectbox(
    "Select Output Format",
    ["CSV", "XLS"]
)

def read_xlsb(file):
    data = []
    with open_workbook(file) as wb:
        sheet = wb.get_sheet(1)   # First sheet
        for row in sheet.rows():
            data.append([item.v for item in row])
    df = pd.DataFrame(data[1:], columns=data[0])
    return df

if uploaded_file:

    st.info("Reading XLSB file...")

    df = read_xlsb(uploaded_file)

    st.success("File loaded successfully ✅")
    st.dataframe(df.head())

    buffer = BytesIO()

    if output_format == "CSV":
        df.to_csv(buffer, index=False)
        file_name = "converted.csv"
        mime = "text/csv"

    elif output_format == "XLS":
        with pd.ExcelWriter(buffer, engine="xlwt") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        file_name = "converted.xls"
        mime = "application/vnd.ms-excel"

    st.download_button(
        label="⬇️ Download Converted File",
        data=buffer.getvalue(),
        file_name=file_name,
        mime=mime
    )
