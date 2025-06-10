import streamlit as st
import pandas as pd
import os
from io import BytesIO
import base64
import re # Import the regular expression module

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
    - Pre-processes each cell:
        - Strings that are valid percentages (e.g., "75.5%") are converted to decimals.
        - Text containing '%' but is not a valid percentage (e.g. "5% discount") is ignored.
    - Then, attempts column-wide conversion with updated logic:
        1. A column is converted to numeric ONLY IF all non-null values are numeric.
           If even one text value exists, the column is treated as a string.
        2. Date-like columns are converted to dates.
        3. All other columns default to stripped strings.
    """
    df1 = df1_orig.copy()
    df2 = df2_orig.copy()

    # A precise helper function using a regular expression to find percentages.
    def convert_value(val):
        """
        Converts a value only if it's a string that strictly matches a percentage format.
        """
        # Immediately return non-string values to avoid errors.
        if not isinstance(val, str):
            return val

        # Regex to match a string that IS a percentage and nothing else.
        # ^ and $ anchor the match to the start and end of the string.
        percentage_pattern = re.compile(r"^\s*(-?\d+\.?\d*)\s*%\s*$")

        # We use .match() on the stripped string. If it's a match...
        if percentage_pattern.match(val.strip()):
            # ...we can safely perform the conversion.
            return pd.to_numeric(val.strip().rstrip('%'), errors='coerce') / 100.0
        
        # If the string does not match the percentage pattern, return it unchanged.
        return val

    for col in common_columns:
        if col not in df1.columns or col not in df2.columns:
            continue

        # --- Pre-processing Step: Apply cell-wise conversion for percentages ---
        # This step is applied first to handle explicit percentage strings.
        df1[col] = df1[col].apply(convert_value)
        df2[col] = df2[col].apply(convert_value)

        # --- Main Conversion Logic ---

        # 1. Handle Numeric Columns (Stricter Check)
        # We will only convert to numeric if ALL non-null values are numeric.
        try:
            # Attempt to convert to numeric, coercing errors to NaT/NaN
            temp_num1 = pd.to_numeric(df1[col], errors='coerce')
            temp_num2 = pd.to_numeric(df2[col], errors='coerce')

            # The key check: Did 'to_numeric' create new nulls?
            # If the number of nulls increased, it means some values were non-numeric text.
            is_fully_numeric1 = df1[col].isnull().sum() == temp_num1.isnull().sum()
            is_fully_numeric2 = df2[col].isnull().sum() == temp_num2.isnull().sum()

            # Only if BOTH columns are purely numeric, we perform the conversion.
            if is_fully_numeric1 and is_fully_numeric2:
                df1[col] = temp_num1
                df2[col] = temp_num2
                continue # Move to the next column
        except Exception:
            # If any error occurs during numeric check, we'll fall back to string.
            pass


        # 2. Handle Datetime Columns (Logic remains unchanged as requested)
        try:
            # Coerce to datetime, setting invalid parsing as NaT
            temp_dt1 = pd.to_datetime(df1[col], errors='coerce', dayfirst=True)
            temp_dt2 = pd.to_datetime(df2[col], errors='coerce', dayfirst=True)

            is_original_dt1 = pd.api.types.is_datetime64_any_dtype(df1[col].dtype)
            is_original_dt2 = pd.api.types.is_datetime64_any_dtype(df2[col].dtype)
            
            # Check if columns can be fully converted without losing all data
            can_be_converted_dt1 = not temp_dt1.isnull().all()
            can_be_converted_dt2 = not temp_dt2.isnull().all()

            # If either was already a datetime, or both can be successfully converted
            if (is_original_dt1 or is_original_dt2) or (can_be_converted_dt1 and can_be_converted_dt2):
                df1[col] = temp_dt1.dt.date
                df2[col] = temp_dt2.dt.date
                continue # Move to the next column
        except Exception:
            # If date conversion fails, pass and fall through to string conversion
            pass

        # 3. Default to String
        # This is the fallback for columns that are not purely numeric or date-like.
        df1[col] = df1[col].astype(str).str.strip().replace('nan', '')
        df2[col] = df2[col].astype(str).str.strip().replace('nan', '')

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
                    <li>Percentage strings (e.g., "55.5%") converted to decimals.</li>
                    <li>Columns with ONLY numbers will be treated as numeric.</li>
                    <li>Date-like columns converted to dates (time part removed).</li>
                    <li><b>Columns with any text (mixed with numbers) will be preserved as text.</b></li>
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
