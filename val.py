import streamlit as st
import pandas as pd
import io
import numpy as np
import os
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
import base64  # For base64 image encoding

# Define the checklist data as a DataFrame
checklist_data = {
    "S.No": range(1, 8),
    "Checklist": [
        "All the columns of excel replicated in PBI (No extra columns)",
        "All the filters of excel replicated in PBI",
        "Filters working as expected (single/multi select as usual)",
        "Column names matching with excel",
        "Currency symbols to be replicated",
        "Pre-applied filters while generating validation report?",
        "Sorting is replicated"
    ],
}
checklist_df = pd.DataFrame(checklist_data)



def generate_validation_report(excel_df, pbi_df):
    dims = [col for col in excel_df.columns if col in pbi_df.columns and 
            (excel_df[col].dtype == 'object' or '_id' in col.lower() or '_key' in col.lower() or
             '_ID' in col or '_KEY' in col)]

    excel_df[dims] = excel_df[dims].fillna('NAN')
    pbi_df[dims] = pbi_df[dims].fillna('NAN')

    excel_measures = [col for col in excel_df.columns if col not in dims and np.issubdtype(excel_df[col].dtype, np.number)]
    pbi_measures = [col for col in pbi_df.columns if col not in dims and np.issubdtype(pbi_df[col].dtype, np.number)]
    
    all_measures = list(set(excel_measures) & set(pbi_measures))

    excel_agg = excel_df.groupby(dims)[all_measures].sum().reset_index()
    pbi_agg = pbi_df.groupby(dims)[all_measures].sum().reset_index()

    excel_agg['unique_key'] = excel_agg[dims].astype(str).agg('-'.join, axis=1).str.upper()
    pbi_agg['unique_key'] = pbi_agg[dims].astype(str).agg('-'.join, axis=1).str.upper()

    excel_agg = excel_agg[['unique_key'] + [col for col in excel_agg.columns if col != 'unique_key']]
    pbi_agg = pbi_agg[['unique_key'] + [col for col in pbi_agg.columns if col != 'unique_key']]

    validation_report = pd.DataFrame({'unique_key': list(set(excel_agg['unique_key']) | set(pbi_agg['unique_key']))})

    for dim in dims:
        validation_report[dim] = validation_report['unique_key'].map(dict(zip(excel_agg['unique_key'], excel_agg[dim])))
        validation_report[dim].fillna(validation_report['unique_key'].map(dict(zip(pbi_agg['unique_key'], pbi_agg[dim]))), inplace=True)

    validation_report['presence'] = validation_report['unique_key'].apply(
        lambda key: 'Present in Both' if key in excel_agg['unique_key'].values and key in pbi_agg['unique_key'].values
        else ('Present in excel' if key in excel_agg['unique_key'].values
              else 'Present in PBI')
    )

    for measure in all_measures:
        validation_report[f'{measure}_excel'] = validation_report['unique_key'].map(dict(zip(excel_agg['unique_key'], excel_agg[measure])))
        validation_report[f'{measure}_PBI'] = validation_report['unique_key'].map(dict(zip(pbi_agg['unique_key'], pbi_agg[measure])))
        
        validation_report[f'{measure}_Diff'] = np.where(
            (validation_report[f'{measure}_PBI'].fillna(0) == 0) | (validation_report[f'{measure}_excel'].fillna(0) == 0),
            np.where(
                (validation_report[f'{measure}_PBI'].fillna(0) == 0) & (validation_report[f'{measure}_excel'].fillna(0) == 0),
                0,
                1
            ),
            abs(round((validation_report[f'{measure}_PBI'].fillna(0) - validation_report[f'{measure}_excel'].fillna(0)) / 
                      validation_report[f'{measure}_excel'].fillna(0), 4))
        )

    column_order = ['unique_key'] + dims + ['presence'] + \
                    [col for measure in all_measures for col in 
                     [f'{measure}_excel', f'{measure}_PBI', f'{measure}_Diff']]
    validation_report = validation_report[column_order]

    return validation_report, excel_agg, pbi_agg

def column_checklist(excel_df, pbi_df):
    excel_columns = excel_df.columns.tolist()
    pbi_columns = pbi_df.columns.tolist()

    checklist_df = pd.DataFrame({
        'excel Columns': excel_columns + [''] * (max(len(pbi_columns), len(excel_columns)) - len(excel_columns)),
        'PowerBI Columns': pbi_columns + [''] * (max(len(pbi_columns), len(excel_columns)) - len(pbi_columns))
    })

    checklist_df['Match'] = checklist_df.apply(lambda row: row['excel Columns'] == row['PowerBI Columns'], axis=1)
    
    return checklist_df

def generate_diff_checker(validation_report):
    diff_columns = [col for col in validation_report.columns if col.endswith('_Diff')]

    diff_checker = pd.DataFrame({
        'Diff Column Name': diff_columns,
        'Percentage Difference': [f"{validation_report[col].mean() * 100:.2f}%" for col in diff_columns]
    })

    presence_summary = {
        'Diff Column Name': 'All rows present in both',
        'Percentage Difference': 'Yes' if all(validation_report['presence'] == 'Present in Both') else 'No'
    }
    diff_checker = pd.concat([diff_checker, pd.DataFrame([presence_summary])], ignore_index=True)

    return diff_checker

def apply_conditional_formatting(ws, validation_report, low_thresh, mid_thresh):
    dark_green_fill = PatternFill(start_color='19D119', end_color='19D119', fill_type='solid')
    dark_red_fill = PatternFill(start_color='E82D1C', end_color='E82D1C', fill_type='solid')

    diff_cols = [col for col in validation_report.columns if col.endswith('_Diff')]
    presence_col_idx = validation_report.columns.get_loc('presence') + 1

    for col_idx, col_name in enumerate(validation_report.columns, 1):
        col_letter = get_column_letter(col_idx)

        if col_name.endswith('_Diff'):
            header_cell = ws[f'{col_letter}1']
            header_cell.number_format = '0.00%'

            for row_idx, value in enumerate(validation_report[col_name], 2):
                cell = ws[f'{col_letter}{row_idx}']
                if pd.notna(value):
                    cell.value = value
                    cell.number_format = '0.00%'

                    if value <= low_thresh:
                        cell.fill = dark_green_fill
                    elif value <= mid_thresh:
                        ratio = (value - low_thresh) / (mid_thresh - low_thresh)
                        r = int(255 + (139 - 255) * ratio)
                        g = int(255 - (255 - 0) * ratio)
                        b = 0
                        color = f'{r:02X}{g:02X}{b:02X}'
                        cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                    else:
                        cell.fill = dark_red_fill

        elif col_idx == presence_col_idx:
            for row_idx, value in enumerate(validation_report[col_name], 2):
                cell = ws[f'{col_letter}{row_idx}']
                if value == 'Present in Both':
                    cell.fill = dark_green_fill
                elif value in ['Present in excel', 'Present in PBI']:
                    cell.fill = dark_red_fill

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

    st.markdown('<div class="title">Validation Report Generator</div>', unsafe_allow_html=True)

    # Sidebar thresholds
    st.sidebar.header("⚙️ Diff Color Thresholds")
    low_threshold = st.sidebar.number_input("Green Threshold (≤)", min_value=0.0, max_value=1.0, value=0.1, step=0.01)
    mid_threshold = st.sidebar.number_input("Amber Threshold (≤)", min_value=0.0, max_value=1.0, value=0.5, step=0.01)

    st.markdown("""
    <div class="instructions">
    <h3 style="color: #4682B4;">How to Use:</h3>
    <ul>
        <li>Upload an Excel file with two sheets: "excel" and "PBI".</li>
        <li>Ensure column names are similar in both sheets for accurate comparison.</li>
        <li>For ID/Key/Code columns, include "_ID" or "_KEY" in the names (case insensitive).</li>
        <li>Preview your validation report and download the formatted Excel file!</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Drop Your Excel File Here!",
        type=["xls", "xlsx"],
        help="Upload an Excel file with 'excel' and 'PBI' sheets."
    )

    if uploaded_file is not None:
        st.markdown(f'<div class="file-list"><strong>Uploaded File:</strong> {uploaded_file.name}</div>', unsafe_allow_html=True)
        
        with st.spinner("Generating your validation report... Hang tight!"):
            try:
                xls = pd.ExcelFile(uploaded_file)
                excel_df = pd.read_excel(xls, 'excel')
                pbi_df = pd.read_excel(xls, 'PBI')

                excel_df = excel_df.apply(lambda x: x.str.upper().str.strip() if x.dtype == "object" else x)
                pbi_df = pbi_df.apply(lambda x: x.str.upper().str.strip() if x.dtype == "object" else x)

                validation_report, excel_agg, pbi_agg = generate_validation_report(excel_df, pbi_df)
                column_checklist_df = column_checklist(excel_df, pbi_df)
                diff_checker_df = generate_diff_checker(validation_report)

                st.subheader("Validation Report Preview")
                display_report = validation_report.copy()
                for col in display_report.columns:
                    if col.endswith('_Diff'):
                        display_report[col] = display_report[col].apply(lambda x: f"{x*100:.2f}%" if pd.notna(x) else x)
                st.dataframe(display_report)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    original_filename = os.path.splitext(uploaded_file.name)[0]
                    sheet_name = f"{original_filename}_validation_report"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    validation_report.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    apply_conditional_formatting(ws, validation_report, low_threshold, mid_threshold)

                output.seek(0)
                new_file_name = f"{original_filename}_validation_report.xlsx"
                st.markdown(
                    f'<div class="success-box">Success! Your validation report is ready: <strong>{new_file_name}</strong></div>',
                    unsafe_allow_html=True
                )
                st.download_button(
                    label="Download Your Validation Report!",
                    data=output,
                    file_name=new_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.markdown(
                    f'<div class="error-box">Oops! An error occurred: {str(e)}</div>',
                    unsafe_allow_html=True
                )
    st.markdown("---")
# For testing standalone
if __name__ == "__main__":
    run();

