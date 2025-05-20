import io
import streamlit as st
from DeductionConvertUNFIWest import convert_pdf_to_excel, save_to_excel

st.title("PDF → Excel Converter")
uploaded = st.file_uploader("Upload your PDF", type="pdf")
if uploaded:
    # Save to a temp file
    with open("temp.pdf", "wb") as f:
        f.write(uploaded.getbuffer())

    # Run conversion
    result = convert_pdf_to_excel("temp.pdf", output_path=None)
    df = result["main_data"]
    # Prepare DataFrame or Bytes for download
    excel_bytes = io.BytesIO()
    save_to_excel(result, excel_bytes)
    st.download_button(
        label="Download Excel",
        data=excel_bytes.getvalue(),
        file_name="converted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )