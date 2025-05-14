import streamlit as st
import pandas as pd
import os
from io import BytesIO
import base64  # For base64 image encoding

def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        return None

def run():
    # Removed st.set_page_config()

    # Custom CSS for styling
    st.markdown("""
        <style>
        .title {
            font-size: 36px;
            color: #FF4B4B;
            text-align: center;
            font-weight: bold;
            margin-bottom: 20px;
        }
        .instructions {
            background-color: #F0F8FF;
            color: #333333;
            padding: 15px;
            border-radius: 10px;
            border-left: 5px solid #4682B4;
            margin-bottom: 20px;
        }
        .file-list {
            background-color: #F5F5F5;
            color: #333333;
            padding: 10px;
            border-radius: 5px;
            margin-top: 10px;
            margin-bottom: 10px;
        }
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            font-weight: bold;
        }
        .stButton>button:hover {
            background-color: #45A049;
        }
        .success-box {
            background-color: #E6FFE6;
            color: #333333;
            padding: 15px;
            border-radius: 10px;
            border-left: 5px solid #2ECC71;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        .error-box {
            background-color: #FFE6E6;
            color: #333333;
            padding: 15px;
            border-radius: 10px;
            border-left: 5px solid #FF4B4B;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        </style>
    """, unsafe_allow_html=True)

    # Title
    st.markdown('<div class="title">Standardiser</div>', unsafe_allow_html=True)

    # Instructions
    st.markdown("""
        <div class="instructions">
        <h3 style="color: #4682B4;">How to Use:</h3>
        <ul>
            <li>Upload an Excel file.</li>
            <li>Ensure the file contains sheets named "excel" and "PBI".</li>
            <li>Columns common to both sheets will be standardized (numeric, date, or string).</li>
            <li>Download the new Excel file with standardized data.</li>
        </ul>
        </div>
    """, unsafe_allow_html=True)

    # Upload
    st.markdown("### 📤 Upload Excel File")
    uploaded_file = st.file_uploader(
        "Upload an Excel file containing sheets named 'excel' and 'PBI'",
        type=["xlsx"]
    )

    if uploaded_file:
        st.markdown(f'<div class="file-list"><strong>Uploaded File:</strong> {uploaded_file.name}</div>', unsafe_allow_html=True)

        with st.spinner("Standardizing your data..."):
            try:
                xl = pd.ExcelFile(uploaded_file)
                df_excel = xl.parse('excel')
                df_pbi = xl.parse('PBI')

                common_columns = [col for col in df_excel.columns if col in df_pbi.columns]

                def standardize_column_data(df1, df2, common_columns):
                    for col in common_columns:
                        if pd.api.types.is_numeric_dtype(df1[col]) and pd.api.types.is_numeric_dtype(df2[col]):
                            df1[col] = pd.to_numeric(df1[col], errors='coerce')
                            df2[col] = pd.to_numeric(df2[col], errors='coerce')
                        elif pd.api.types.is_datetime64_any_dtype(df1[col]) or pd.api.types.is_datetime64_any_dtype(df2[col]):
                            df1[col] = pd.to_datetime(df1[col], errors='coerce').dt.date
                            df2[col] = pd.to_datetime(df2[col], errors='coerce').dt.date
                        else:
                            df1[col] = df1[col].astype(str).str.strip()
                            df2[col] = df2[col].astype(str).str.strip()
                    return df1, df2

                df_excel_std, df_pbi_std = standardize_column_data(df_excel.copy(), df_pbi.copy(), common_columns)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel_std.to_excel(writer, sheet_name='excel', index=False)
                    df_pbi_std.to_excel(writer, sheet_name='PBI', index=False)
                output.seek(0)

                original_name = os.path.splitext(uploaded_file.name)[0]
                output_filename = f"{original_name}_std.xlsx"

                st.markdown(
                    f'<div class="success-box">✅ Standardization complete. Download the standardized file below:</div>',
                    unsafe_allow_html=True
                )

                st.download_button(
                    label="📥 Download Standardized Excel",
                    data=output,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except ValueError as e:
                st.markdown(
                    f'<div class="error-box">⚠️ Sheet error: {e}</div>',
                    unsafe_allow_html=True
                )
            except Exception as e:
                st.markdown(
                    f'<div class="error-box">🚨 Unexpected error: {e}</div>',
                    unsafe_allow_html=True
                )

    st.markdown("---")
# For testing standalone
if __name__ == "__main__":
    run()
