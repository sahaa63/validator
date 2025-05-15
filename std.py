import streamlit as st
import pandas as pd
import os
from io import BytesIO
import base64  # For base64 image encoding

# Function to encode image to base64 (currently unused in the main logic but kept from original)
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
    Standardizes the data types of common columns between two DataFrames.
    Prioritizes numeric, then datetime, then string.
    """
    df1 = df1_orig.copy()
    df2 = df2_orig.copy()

    for col in common_columns:
        # Attempt to convert to numeric first
        # Create temporary series to test numeric conversion without altering original until decision
        temp_s1_numeric = pd.to_numeric(df1[col], errors='coerce')
        temp_s2_numeric = pd.to_numeric(df2[col], errors='coerce')

        # Condition for numeric:
        # If BOTH columns can be meaningfully converted to numeric (i.e., not all values become NaN).
        # This handles cases where columns are 'object' dtype but contain numbers, or are already numeric.
        if not temp_s1_numeric.isnull().all() and not temp_s2_numeric.isnull().all():
            df1[col] = temp_s1_numeric
            df2[col] = temp_s2_numeric
            # st.write(f"Column '{col}' standardized as NUMERIC.") # Optional: for debugging

        # Else, if not treated as numeric, attempt datetime conversion.
        # Condition for datetime:
        # 1. EITHER original DataFrame column's dtype is already datetime OR
        # 2. BOTH columns can be meaningfully converted to datetime (i.e., not all values become NaT).
        elif (pd.api.types.is_datetime64_any_dtype(df1[col].dtype) or \
              pd.api.types.is_datetime64_any_dtype(df2[col].dtype)) or \
             (not pd.to_datetime(df1[col], errors='coerce').isnull().all() and \
              not pd.to_datetime(df2[col], errors='coerce').isnull().all()):
            df1[col] = pd.to_datetime(df1[col], errors='coerce').dt.date
            df2[col] = pd.to_datetime(df2[col], errors='coerce').dt.date
            # st.write(f"Column '{col}' standardized as DATE.") # Optional: for debugging
            
        # Else, default to string.
        else:
            df1[col] = df1[col].astype(str).str.strip()
            df2[col] = df2[col].astype(str).str.strip()
            # st.write(f"Column '{col}' standardized as STRING.") # Optional: for debugging
            
    return df1, df2

def run():
    # Custom CSS for styling (unchanged)
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
            <li>Columns common to both sheets will be standardized. The tool will attempt to convert to numeric, then date, then string type.</li>
            <li>Download the new Excel file with standardized data. The sheet names will be preserved as "excel" and "PBI".</li>
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
                
                df_excel = xl.parse('excel')
                df_pbi = xl.parse('PBI')

                # Identify common columns
                common_columns = [col for col in df_excel.columns if col in df_pbi.columns]

                if not common_columns:
                    st.warning("No common columns found between 'excel' and 'PBI' sheets.")
                else:
                    st.markdown(f"**Common columns found:** ` {', '.join(common_columns)} `")
                    
                    # Standardize data in common columns
                    df_excel_std, df_pbi_std = standardize_column_data(df_excel, df_pbi, common_columns)

                    # Prepare the output Excel file
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Save standardized dataframes with original sheet names
                        df_excel_std.to_excel(writer, sheet_name='excel', index=False)
                        df_pbi_std.to_excel(writer, sheet_name='PBI', index=False)
                    output.seek(0) # Reset buffer's position to the beginning

                    # Define the output filename
                    original_name = os.path.splitext(uploaded_file.name)[0]
                    output_filename = f"{original_name}_standardized.xlsx" # Output filename still indicates it's standardized

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
                st.error(f"An unexpected error occurred: {e}")
                st.markdown(
                    f'<div class="error-box">🚨 Unexpected error: {e}</div>',
                    unsafe_allow_html=True
                )

    st.markdown("---")

# Entry point for running the Streamlit app
if __name__ == "__main__":
    run()
