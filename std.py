import streamlit as st
import pandas as pd
import os
from io import BytesIO
import base64

# This function is not used in the main logic but kept as it was in the original code
def get_base64_image(image_path):
    """Reads an image file and returns its base64 encoded string."""
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        st.error(f"Image not found at {image_path}")
        return None

def standardize_column_data(df1_orig, df2_orig, common_columns):
    """
    Standardizes data types of common columns with a clear priority:
    1. Numeric: If both columns can be treated as numbers.
    2. Datetime: If they can be parsed as dates (time is removed).
    3. String: As a final fallback.
    """
    df1 = df1_orig.copy()
    df2 = df2_orig.copy()

    for col in common_columns:
        # Step 1: Attempt Numeric Conversion
        # This is a good first check for obviously numeric columns.
        try:
            df1_numeric = pd.to_numeric(df1[col])
            df2_numeric = pd.to_numeric(df2[col])
            # If both conversions succeed without error, apply them
            df1[col] = df1_numeric
            df2[col] = df2_numeric
            # Continue to the next column
            continue
        except (ValueError, TypeError):
            # If conversion to numeric fails, proceed to check for dates.
            pass

        # Step 2: Attempt Datetime Conversion
        # Use errors='coerce' to turn un-parsable values into NaT (Not a Time)
        temp_dt1 = pd.to_datetime(df1[col], errors='coerce')
        temp_dt2 = pd.to_datetime(df2[col], errors='coerce')
        
        # Condition: Proceed if the conversion resulted in at least one valid date
        # in EACH dataframe, preventing columns of names/text from being converted.
        if not temp_dt1.isnull().all() and not temp_dt2.isnull().all():
            # Apply the conversion and use .dt.date to STRIP the time component
            df1[col] = temp_dt1.dt.date
            df2[col] = temp_dt2.dt.date
            # Continue to the next column
            continue

        # Step 3: Default to String Conversion
        # This runs only if both numeric and date conversions fail.
        df1[col] = df1[col].astype(str).str.strip()
        df2[col] = df2[col].astype(str).str.strip()
            
    return df1, df2

def run():
    # Custom CSS for styling
    st.markdown("""
        <style>
        .title { font-size: 36px; color: #FF4B4B; text-align: center; font-weight: bold; margin-bottom: 20px; }
        .instructions { background-color: rgb(128 128 128 / 10%); padding: 15px; border-radius: 10px; border-left: 5px solid #4682B4; margin-bottom: 20px; }
        .file-list { background-color: #F5F5F5; color: #333333; padding: 10px; border-radius: 5px; margin-top: 10px; margin-bottom: 10px; }
        .stButton>button { background-color: #4CAF50; color: white; border: none; padding: 10px 20px; border-radius: 5px; font-weight: bold; }
        .stButton>button:hover { background-color: #45A049; }
        .success-box { background-color: #E6FFE6; color: #333333; padding: 15px; border-radius: 10px; border-left: 5px solid #2ECC71; margin-top: 20px; margin-bottom: 20px; }
        .error-box { background-color: #FFE6E6; color: #333333; padding: 15px; border-radius: 10px; border-left: 5px solid #FF4B4B; margin-top: 20px; margin-bottom: 20px; }
        </style>
    """, unsafe_allow_html=True)

    # Title
    st.markdown('<div class="title">Data Standardiser</div>', unsafe_allow_html=True)

    # Instructions (Updated to reflect new logic)
    st.markdown("""
        <div class="instructions">
        <h3 style="color: #4682B4;">How to Use:</h3>
        <ul>
            <li>Upload an Excel file.</li>
            <li>Ensure the file contains sheets named "excel" and "PBI".</li>
            <li>Columns common to both sheets will be standardized with the following priority:
                <ol>
                    <li><b>Numeric:</b> Columns that are purely numeric.</li>
                    <li><b>Date:</b> Columns with dates or datetimes (time is removed).</li>
                    <li><b>Text:</b> All other columns.</li>
                </ol>
            </li>
            <li>Download the new Excel file with standardized data.</li>
        </ul>
        </div>
    """, unsafe_allow_html=True)

    # File Upload
    st.markdown("### 📤 Upload Excel File")
    uploaded_file = st.file_uploader(
        "Upload an Excel file containing sheets named 'excel' and 'PBI'",
        type=["xlsx"]
    )

    if uploaded_file:
        st.markdown(f'<div class="file-list"><strong>Uploaded File:</strong> {uploaded_file.name}</div>', unsafe_allow_html=True)

        with st.spinner("Standardizing your data... Please wait."):
            try:
                xl = pd.ExcelFile(uploaded_file)
                if 'excel' not in xl.sheet_names:
                    raise ValueError("Sheet 'excel' not found in the uploaded file.")
                if 'PBI' not in xl.sheet_names:
                    raise ValueError("Sheet 'PBI' not found in the uploaded file.")
                
                df_excel_orig = xl.parse('excel')
                df_pbi_orig = xl.parse('PBI')

                common_columns = [col for col in df_excel_orig.columns if col in df_pbi_orig.columns]

                if not common_columns:
                    st.warning("No common columns found between 'excel' and 'PBI' sheets.")
                else:
                    st.markdown(f"**Common columns found:** ` {', '.join(common_columns)} `")
                    
                    df_excel_std, df_pbi_std = standardize_column_data(df_excel_orig, df_pbi_orig, common_columns)

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_excel_std.to_excel(writer, sheet_name='excel', index=False)
                        df_pbi_std.to_excel(writer, sheet_name='PBI', index=False)
                    output.seek(0)

                    original_name = os.path.splitext(uploaded_file.name)[0]
                    output_filename = f"{original_name}_standardized.xlsx" 

                    st.markdown(
                        '<div class="success-box">✅ Standardization complete. Download the standardized file below:</div>',
                        unsafe_allow_html=True
                    )

                    st.download_button(
                        label="📥 Download Standardized Excel",
                        data=output,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except ValueError as ve:
                st.markdown(
                    f'<div class="error-box">⚠️ Processing error: {ve}</div>',
                    unsafe_allow_html=True
                )
            except Exception as e:
                st.markdown(
                    f'<div class="error-box">🚨 Unexpected error: {e}</div>',
                    unsafe_allow_html=True
                )

    st.markdown("---")

# Entry point for running the Streamlit app
if __name__ == "__main__":
    run()
