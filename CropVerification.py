import streamlit as st
import pandas as pd
import io
import xlrd
from datetime import datetime

st.set_page_config(page_title="Crop Data Verification", page_icon="🌾", layout="centered")
st.title("🌾 Crop Data Mapper")

st.markdown(
    "### Welcome to the Crop Verification tool\nUpload your crop booking and MAO verification files below, then download the merged report.")

# st.warning("Use original Excel files from the system. Please do not edit, rename, or modify them before uploading.")

# Added Help Document 
help_url = "https://github.com/apkanisandeep01/Crop-Verification-Merger/blob/d9bfaa23c14735157eec7e3c3cf9c94d35853a20/Help%20Document.pdf"

st.markdown(
    f'<a href="{help_url}" target="_blank"><button style="padding:10px 14px; font-size:15px; border-radius:10px; background:#2a9d8f; color:white; border:none;">Access Help Document</button></a>',
    unsafe_allow_html=True
)

with st.container():
    cols = st.columns([2, 1])
    with cols[0]:
        st.subheader("Upload files")
        uploaded_files = st.file_uploader(
            "Crop booking file from portal",
            type=["xls", "xlsx"],
            accept_multiple_files=True,
            help="Upload one or more crop booking Excel files exported from the booking system."
        )
        st.caption("You may upload multiple crop booking files at once.")
        st.divider()
        mao_file = st.file_uploader(
            "MAO/ADA/DAO verification file",
            type=["xls", "xlsx"],
            accept_multiple_files=False,
            help="Upload the MAO/ADA/DAO verification file for matching."
        )
        st.caption("Upload the MAO verification spreadsheet downloaded from the verification system.")
    with cols[1]:
        pass
     
st.divider()

if uploaded_files and mao_file:
    dfs = []
    for file in uploaded_files:
        try:
            file_df = pd.read_excel(file)
            dfs.append(file_df)
        except Exception:
            st.error(f"Could not open '{getattr(file, 'name', 'file')}'. Please make sure it is a valid Excel file (.xls or .xlsx).")

    if not dfs:
        st.error("No valid crop booking files were loaded. Please upload at least one valid Excel file.")
        st.stop()

    crop_df = pd.concat(dfs, ignore_index=True)
    crop_df.columns = [str(col).strip() for col in crop_df.columns]
    required_cols = [
        'Season', 'Mandal', 'Village', 'PPBNO', 'FarmerName',
        'FatherName', 'MobileNo', 'BaseSurveyNo', 'SurveyNo',
        'SurveyExtent', 'CropName', 'CropVarietyName',
        'CropSown_Acres', 'CropSown_Guntas', 'SowingWeek'
    ]
    crop_df = crop_df[[c for c in required_cols if c in crop_df.columns]]
    if crop_df.empty:
        st.error("The crop booking files did not include the expected columns.")
        st.info("Expected columns include Season, Mandal, Village, PPBNO, FarmerName, MobileNo, SurveyNo, CropName, CropVarietyName, CropSown_Acres, CropSown_Guntas, SowingWeek.")
        st.stop()

    st.subheader("Crop Booking data preview")
    st.dataframe(crop_df.head(3), use_container_width=True)
    # st.success("Crop booking files loaded successfully.")
    # st.divider()

    try:
        mao_df = pd.read_excel(mao_file, header=2)
        mao_df.dropna(axis=1, inplace=True)
        mao_df.columns = [str(col).strip() for col in mao_df.columns]
    except Exception:
        st.error("Could not open the MAO verification file. Please make sure it is a valid Excel file (.xls or .xlsx).")
        st.stop()

    st.subheader("MAO verification data preview")
    st.dataframe(mao_df.head(3), use_container_width=True)
    # st.success("MAO verification file loaded successfully.")

    def find_column(df, options):
        for option in options:
            if option in df.columns:
                return option
        return None

    mao_village_col = find_column(mao_df, ['VIllage', 'Village'])
    mao_survey_col = find_column(mao_df, ['Survey Number', 'SurveyNo', 'Survey No'])
    crop_survey_col = find_column(crop_df, ['SurveyNo', 'Survey No', 'SurveyNumber'])

    if not mao_village_col or not mao_survey_col:
        st.error("The MAO file does not have the required merge columns.")
        st.info("Expected MAO columns: 'Village' or 'VIllage', and 'Survey Number', 'SurveyNo', or 'Survey No'.")
        st.stop()

    if not crop_survey_col:
        st.error("The crop booking file does not have the required survey column.")
        st.info("Expected crop booking columns: 'SurveyNo', 'Survey No', or 'SurveyNumber'.")
        st.stop()

    try:
        verification_df = mao_df.merge(
            crop_df,
            left_on=[mao_village_col, mao_survey_col],
            right_on=['Village', crop_survey_col],
            how='inner'
        )
    except Exception as e:
        st.error("Failed to merge crop booking and MAO verification data.")
        st.info(str(e))
        st.stop()

    if verification_df.empty:
        st.warning("No matching records were found after merging the uploaded files.")
        st.stop()

    display_cols = [
        'Division', 'Mandal_x', 'VIllage',
        'Pattadar Passbook Number', 'Farmer Name', 'MobileNo', 'BaseSurveyNo',
        'Survey Number', 'Survey Extent','CropName', 'CropVarietyName', 'CropSown_Acres',
        'CropSown_Guntas', 'SowingWeek'
    ]
    available_cols = [col for col in display_cols if col in verification_df.columns]
    verification_df = verification_df[available_cols]

    rename_map = {
        'Mandal_x': 'Mandal','VIllage': 'Village',
        'Pattadar Passbook Number': 'PPB No',
        'Farmer Name': 'FarmerName','MobileNo': 'ContactNumber',
        'BaseSurveyNo': 'Base Survey No','Survey Number': 'Sy No',
        'CropVarietyName': 'CropVariety'
    }
    verification_df = verification_df.rename(columns=rename_map)
    st.divider()
    st.subheader("✅ Preview of Merged & Processed Data")
    st.dataframe(verification_df.head(5), use_container_width=True)

    today = datetime.now().strftime("%d_%m_%Y")
    excel_file_name = f"CropData_{today}.xlsx"

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        verification_df.to_excel(writer, index=False, sheet_name="Crop Verification")
        workbook = writer.book
        worksheet = writer.sheets["Crop Verification"]

        header_format = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "top",
            "fg_color": "#D7E4BC", "border": 1})

        for col_num, value in enumerate(verification_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            series = verification_df[value].astype(str)
     
    st.success("Your verified report is ready to download.")
    st.download_button(
        label="Download Excel File",
        data=excel_buffer.getvalue(),
        file_name=excel_file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
