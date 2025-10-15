import streamlit as st
import pandas as pd
import io
import xlrd
from datetime import datetime

st.set_page_config(page_title="Crop Data Verification", layout="wide")
st.title("🌾 Crop Data Verification App")

st.warning("Always ensure to upload the Original files that are directly Downloaded from the site")

# Added Help Document 
help_url = "https://github.com/apkanisandeep01/Crop-Verification-Merger/blob/d9bfaa23c14735157eec7e3c3cf9c94d35853a20/Help%20Document.pdf"

st.markdown(
    f'<a href="{help_url}" target="_blank"><button style="padding:10px 20px; font-size:16px;">📘 Open Help Document</button></a>',
    unsafe_allow_html=True
)
st.subheader("Upload Crop Booking Excel Files")
# Upload multiple crop booking files
uploaded_files = st.file_uploader("Upload here", type=None, accept_multiple_files=True)
st.divider()
st.subheader("Upload MAO Verification list of Excel File")
# Upload MAO verification file
mao_file = st.file_uploader("Upload here", type=None, accept_multiple_files=False)

if uploaded_files and mao_file:
    dfs = []

    # Read all crop booking files
    for file in uploaded_files:
        try:
            file_df = pd.read_excel(file)
            dfs.append(file_df)
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

    if dfs:
        crop_df = pd.concat(dfs, ignore_index=True)
        st.dataframe(crop_df.head(5), use_container_width=True)
        # Required columns (only keep if they exist)
        required_cols = ['Season', 'Mandal', 'Village', 'PPBNO', 'FarmerName',
                         'FatherName', 'MobileNo', 'BaseSurveyNo', 'SurveyNo',
                         'SurveyExtent', 'CropName', 'CropVarietyName',
                         'CropSown_Acres','CropSown_Guntas', 'SowingWeek']
        crop_df = crop_df[[c for c in required_cols if c in crop_df.columns]]

        # Read MAO file
        try:
            mao_df = pd.read_excel(mao_file, header=2)
            mao_df.dropna(axis=1, inplace=True)
             st.dataframe(mao_df.head(5), use_container_width=True)
        except Exception as e:
            st.error(f"Error reading MAO file: {e}")
            st.stop()
        try:
            # Merge both dataframes
            verification_df = mao_df.merge(
                crop_df,
                left_on=['VIllage', 'Survey Number'],
                right_on=['Village', 'SurveyNo'],
                how='inner'
            )
    
            # Select required columns
            verification_df = verification_df[[
                'Division', 'Mandal_x', 'VIllage',
                'Pattadar Passbook Number', 'Farmer Name', 'Mobile Number','BaseSurveyNo',
                'Survey Number', 'Survey Extent',
                'CropName', 'CropVarietyName', 'CropSown_Acres',
                'CropSown_Guntas', 'SowingWeek'
            ]]
    
            # Rename columns
            verification_df.columns = [
                'Division', 'Mandal', 'Village',
                'PPB No', 'FarmerName', 'ContactNumber','Base Survey No',
                'Sy No', 'Survey Extent',
                'CropName', 'CropVariety', 'CropSown_Acres',
                'CropSown_Guntas', 'SowingWeek'
            ]
        except Exception as e:
            st.error(f"Error reading MAO file: {e}")

        # Show preview in app
        st.subheader("✅ Preview of Merged Data")
        st.dataframe(verification_df.head(20), use_container_width=True)

        # File name with date
        today = datetime.now().strftime("%d_%m_%Y")
        excel_file_name = f"CropData_{today}.xlsx"

        # --- Excel Export ---
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            verification_df.to_excel(writer, index=False, sheet_name="Crop Verification")
            workbook = writer.book
            worksheet = writer.sheets["Crop Verification"]

            # Header formatting
            header_format = workbook.add_format({
                "bold": True, "text_wrap": True, "valign": "top",
                "fg_color": "#D7E4BC", "border": 1
            })

            for col_num, value in enumerate(verification_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

                # Auto-adjust column width based on content
                series = verification_df[value].astype(str)
                max_len = max(series.map(len).max(), len(str(value))) + 2
                worksheet.set_column(col_num, col_num, max_len)
        st.toast("Don't just count the seeds; make every seed count. Smart work is planting with a purpose", icon="❤️")
        st.toast("Your data is Ready!", icon="😍")
        st.toast("Feel free get your excel Downloaded!", icon="🤗")
        st.download_button(
            label="📥 Download Excel File",
            data=excel_buffer.getvalue(),
            file_name=excel_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
st.info("Feel free to use this tool as much as you need. We respect your privacy, so none of the data you enter here is ever stored.")
