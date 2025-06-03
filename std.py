import streamlit as st
import pandas as pd
import os
from io import BytesIO
import base64

# Function to encode image to base64 (kept from original)
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
    1. Percentage strings (e.g., "69.36%") are converted to decimals (0.6936).
    2. Numeric-convertible columns (including integers and decimals) are converted to numbers.
    3. Date-like columns (including various string formats) are converted to dates.
    4. All other columns default to stripped strings.
    """
    df1 = df1_orig.copy()
    df2 = df2_orig.copy()

    for col in common_columns:
        # Ensure the column exists in both dataframes before proceeding
        if col not in df1.columns or col not in df2.columns:
            continue

        # --- 1. Handle Percentage Columns ---
        # Check if the column is object type and contains '%' character
        is_percent_col1 = pd.api.types.is_object_dtype(df1[col]) and df1[col].str.contains('%', na=False).any()
        is_percent_col2 = pd.api.types.is_object_dtype(df2[col]) and df2[col].str.contains('%', na=False).any()

        if is_percent_col1 or is_percent_col2:
            df1[col] = pd.to_numeric(df1[col].astype(str).str.rstrip('%'), errors='coerce') / 100.0
            df2[col] = pd.to_numeric(df2[col].astype(str).str.rstrip('%'), errors='coerce') / 100.0
            continue # Move to the next column

        # --- 2. Handle Numeric Columns (Integers and Decimals) ---
        # Attempt to convert to numeric. This will handle integers, floats, and numeric strings.
        temp_num1 = pd.to_numeric(df1[col], errors='coerce')
        temp_num2 = pd.to_numeric(df2[col], errors='coerce')

        # If conversion is successful for both (i.e., not all values became NaN)
        if not temp_num1.isnull().all() and not temp_num2.isnull().all():
            df1[col] = temp_num1
            df2[col] = temp_num2
            continue # Move to the next column
            
        # --- 3. Handle Datetime Columns ---
        # This block is only reached if the column could not be treated as numeric.
        try:
            temp_dt1 = pd.to_datetime(df1[col], errors='coerce', dayfirst=True)
            temp_dt2 = pd.to_datetime(df2[col], errors='coerce', dayfirst=True)

            is_original_dt1 = pd.api.types.is_datetime64_any_dtype(df1[col].dtype)
            is_original_dt2 = pd.api.types.is_datetime64_any_dtype(df2[col].dtype)
            
            can_be_converted_dt1 = not temp_dt1.isnull().all()
            can_be_converted_dt2 = not temp_dt2.isnull().all()

            if (is_original_dt1 or is_original_dt2) or (can_be_converted_dt1 and can_be_converted_dt2):
                df1[col] = temp_dt1.dt.date
                df2[col] = temp_dt2.dt.date
                continue # Move to the next column
        except Exception:
            # If to_datetime raises an unexpected error, fall through to string conversion
            pass

        # --- 4. Default to String ---
        # If none of the above conditions were met, treat as string.
        df1[col] = df1[col].astype(str).str.strip()
        df2[col] = df2[col].astype(str).str.strip()

    return df1, df2


def run():
    # Custom CSS for styling (current version)
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
            background-color: rgb(128 128 128 / 10%); 
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
            <li>Common columns will be standardized with the following priority:
                <ul>
                    <li>Percentage strings (e.g., "55.5%") converted to decimals (0.555).</li>
                    <li>Integer and decimal columns preserved as numbers.</li>
                    <li>Date-like columns converted to dates (time part removed).</li>
                    <li>Other columns treated as strings.</li>
                </ul>
            </li>
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

        with st.spinner("Standardizing your data... Please wait."):
            try:
                # Read both sheets from the uploaded Excel file
                xl = pd.ExcelFile(uploaded_file)
                if 'excel' not in xl.sheet_names:
                    raise ValueError("Sheet 'excel' not found in the uploaded file.")
                if 'PBI' not in xl.sheet_names:
                    raise ValueError("Sheet 'PBI' not found in the uploaded file.")
                
                df_excel_orig = xl.parse('excel')
                df_pbi_orig = xl.parse('PBI')

                # Identify common columns
                common_columns = [col for col in df_excel_orig.columns if col in df_pbi_orig.columns]

                if not common_columns:
                    st.warning("No common columns found between 'excel' and 'PBI' sheets.")
                else:
                    st.markdown(f"**Common columns found:** ` {', '.join(common_columns)} `")
                    
                    # Standardize data in common columns using copies
                    df_excel_std, df_pbi_std = standardize_column_data(df_excel_orig, df_pbi_orig, common_columns)

                    # Prepare the output Excel file
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Save standardized dataframes with original sheet names
                        df_excel_std.to_excel(writer, sheet_name='excel', index=False)
                        df_pbi_std.to_excel(writer, sheet_name='PBI', index=False)
                    output.seek(0) # Reset buffer's position to the beginning

                    # Define the output filename
                    original_name = os.path.splitext(uploaded_file.name)[0]
                    output_filename = f"{original_name}_standardized.xlsx" 

                    st.markdown(
                        f'<div class="success-box">✅ Standardization complete. Download the standardized file below:</div>',
                        unsafe_allow_html=True
                    )

                    # Provide download button
                    st.download_button(
                        label="📥 Download Standardized Excel",
                        data=output,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except ValueError as ve: # Specific error for missing sheets or other value issues
                st.markdown(
                    f'<div class="error-box">⚠️ Processing error: {ve}</div>',
                    unsafe_allow_html=True
                )
            except Exception as e: # Catch-all for other unexpected errors
                st.markdown(
                    f'<div class="error-box">🚨 Unexpected error: {e}</div>',
                    unsafe_allow_html=True
                )

    st.markdown("---")

# Entry point for running the Streamlit app
if __name__ == "__main__":
    run()
