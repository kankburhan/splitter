import streamlit as st
import pandas as pd
import io
import zipfile

def funding_transform(df):
    df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    
    # Set proper headers
    df.columns = df.iloc[1]
    df = df[2:].reset_index(drop=True)
    
    # Rename relevant columns
    df.rename(columns={
        'EXECUTION SEQUENCE': 'Execution_Sequence',
        'EXPECTED RESULT': 'Expected_Result'
    }, inplace=True)
    
    # Select necessary columns
    columns_to_keep = [
        'NO.', 
        'TEST SCRIPT NUMBER',
        'TEST SCRIPT DESCRIPTION/SCENARIO',
        'TEST OBJECT NAME',
        'Scenario Type',
        'GENERAL INFORMATION / SUMMARY OF THE TEST SCRIPT',
        'PRE-REQUISITES',
        'Execution_Sequence',
        "MODULE\nPut in Defect",
        'Expected_Result',
        'Product / Akad',
        'Assigned Tester',
        'Executed By',
        'PASS OR FAIL',
        'ACTUAL RESULT'
    ]
    df_selected = df[columns_to_keep]
    
    # Function to align lists
    def align_lists(row):
        exec_seq = row['Execution_Sequence']
        exp_res = row['Expected_Result']

        exec_seq = exec_seq.split('\n') if isinstance(exec_seq, str) else []
        exp_res = exp_res.split('\n') if isinstance(exp_res, str) else []

        max_len = max(len(exec_seq), len(exp_res))
        exec_seq.extend([''] * (max_len - len(exec_seq)))
        exp_res.extend([''] * (max_len - len(exp_res)))

        return list(zip(exec_seq, exp_res))
    
    df_selected['Aligned'] = df_selected.apply(align_lists, axis=1)
    df_exploded = df_selected.explode('Aligned', ignore_index=True)
    df_exploded[['Execution_Sequence', 'Expected_Result']] = pd.DataFrame(
        df_exploded['Aligned'].tolist(), 
        index=df_exploded.index
    )
    df_exploded.drop(columns=['Aligned'], inplace=True)
    
    # Create new DataFrame with desired structure
    desired_columns = {
        'NO.': df_exploded['NO.'],
        'Test Script ID': df_exploded['TEST SCRIPT NUMBER'],
        'Skenario': df_exploded['TEST SCRIPT DESCRIPTION/SCENARIO'],
        'Test Object Name': df_exploded['TEST OBJECT NAME'],
        'Scenario Type': df_exploded['Scenario Type'],
        'Summary': df_exploded['GENERAL INFORMATION / SUMMARY OF THE TEST SCRIPT'],
        'Prerequisite': df_exploded['PRE-REQUISITES'],
        'Test Step': df_exploded['Execution_Sequence'],
        'Module': df_exploded["MODULE\nPut in Defect"],
        'Expected Result': df_exploded['Expected_Result'],
        'Product Code_Product Name': df_exploded['Product / Akad'],
        'Assignee': df_exploded['Assigned Tester'],
        'Executed By': df_exploded['Executed By'],
        'Status': df_exploded['PASS OR FAIL'],
        'Actual Result': df_exploded['ACTUAL RESULT']
    }
    
    df_new = pd.DataFrame(desired_columns)
    
    # Define columns to clear for duplicates
    cols_to_clear = [
        'Skenario',
        'Scenario Type', 
        'Test Script ID', 
        'Summary', 
        'Prerequisite', 
        'Test Object Name', 
        'Module',
        'Product Code_Product Name',
        'Assignee',
        'Executed By',
        'Status',
        'Actual Result'
    ]
    
    # Clear duplicate values
    df_new.loc[df_new.duplicated(subset=['Test Script ID']), cols_to_clear] = ''
    
    return df_new

def financing_transform(df):
    df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    
    # Set proper headers
    df.columns = df.iloc[1]
    df = df[2:].reset_index(drop=True)
    
    # Rename relevant columns
    df.rename(columns={
        'EXECUTION SEQUENCE': 'Execution_Sequence',
        'EXPECTED RESULT': 'Expected_Result'
    }, inplace=True)
    
    # Select necessary columns
    columns_to_keep = [
        'NO.', 
        'TEST SCRIPT NUMBER',
        'TEST SCRIPT DESCRIPTION/SCENARIO',
        'TEST OBJECT NAME',
        'Scenario Type',
        'GENERAL INFORMATION / SUMMARY OF THE TEST SCRIPT',
        'PRE-REQUISITES',
        'Execution_Sequence',
        "MODULE\nPut in Defect",
        'Expected_Result',
        'Product / Akad',
        'Feature',
        'Akad',
        'Assigned Tester',
        'Executed By',
        'PASS OR FAIL',
        'ACTUAL RESULT'
    ]
    df_selected = df[columns_to_keep]
    
    # Function to align lists
    def align_lists(row):
        exec_seq = row['Execution_Sequence']
        exp_res = row['Expected_Result']

        exec_seq = exec_seq.split('\n') if isinstance(exec_seq, str) else []
        exp_res = exp_res.split('\n') if isinstance(exp_res, str) else []

        max_len = max(len(exec_seq), len(exp_res))
        exec_seq.extend([''] * (max_len - len(exec_seq)))
        exp_res.extend([''] * (max_len - len(exp_res)))

        return list(zip(exec_seq, exp_res))
    
    df_selected['Aligned'] = df_selected.apply(align_lists, axis=1)
    df_exploded = df_selected.explode('Aligned', ignore_index=True)
    df_exploded[['Execution_Sequence', 'Expected_Result']] = pd.DataFrame(
        df_exploded['Aligned'].tolist(), 
        index=df_exploded.index
    )
    df_exploded.drop(columns=['Aligned'], inplace=True)
    
    # Create new DataFrame with desired structure
    desired_columns = {
        'NO.': df_exploded['NO.'],
        'Test Script ID': df_exploded['TEST SCRIPT NUMBER'],
        'Skenario': df_exploded['TEST SCRIPT DESCRIPTION/SCENARIO'],
        'Test Object Name': df_exploded['TEST OBJECT NAME'],
        'Scenario Type': df_exploded['Scenario Type'],
        'Summary': df_exploded['GENERAL INFORMATION / SUMMARY OF THE TEST SCRIPT'],
        'Prerequisite': df_exploded['PRE-REQUISITES'],
        'Test Step': df_exploded['Execution_Sequence'],
        'Module': df_exploded["MODULE\nPut in Defect"],
        'Expected Result': df_exploded['Expected_Result'],
        'Product Code_Product Name': df_exploded['Product / Akad'],
        'Feature': df_exploded['Feature'],
        'Akad': df_exploded['Akad'],
        'Assignee': df_exploded['Assigned Tester'],
        'Executed By': df_exploded['Executed By'],
        'Status': df_exploded['PASS OR FAIL'],
        'Actual Result': df_exploded['ACTUAL RESULT']
    }
    
    df_new = pd.DataFrame(desired_columns)
    
    # Define columns to clear for duplicates
    cols_to_clear = [
        'Skenario',
        'Scenario Type', 
        'Test Script ID', 
        'Summary', 
        'Prerequisite', 
        'Test Object Name', 
        'Module',
        'Product Code_Product Name',
        'Feature',
        'Akad',
        'Assignee',
        'Executed By',
        'Status',
        'Actual Result'
    ]
    
    # Clear duplicate values
    df_new.loc[df_new.duplicated(subset=['Test Script ID']), cols_to_clear] = ''
    
    return df_new


def process_sheet(input_file, sheet_name):
    # Load the Excel file
    xls = pd.ExcelFile(input_file)
    df = xls.parse(sheet_name=sheet_name)
    
    # Drop empty rows and columns
    filename_lower = input_file.name.lower()
    if '[funding]' in filename_lower or '[t24]' in filename_lower:
        df = funding_transform(df)
    elif '[financing]' in filename_lower:
        df = financing_transform(df)
    else:
        st.error("File name must contain [FUNDING], [FINANCING], or [T24]")
        return None
    
    return df

# Streamlit UI Configuration
st.set_page_config(page_title="Excel Processor", layout="wide")
st.title("📊 Excel Processing Tool")

# File Upload Section
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"], 
                               help="Please upload an Excel file in XLSX format")

if uploaded_file:
    # Sheet Selection
    try:
        output_format = st.radio(
            "Output format", 
            ["Single File (multiple sheets)", "Multiple Files (per sheet per file)", "Single File (one combined sheet)"]
        )
        download_format = st.radio(
            "Select file format:",
            ["XLSX", "CaSV"],
            key="file_format"
        )
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        selected_sheets = st.multiselect("Select worksheets", sheet_names)
        output_format = output_format
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        st.stop()

    # Processing Section
    if st.button("🚀 Start Processing"):
        if not selected_sheets:
            st.error("Please select at least one worksheet.")
            st.stop()
        
        progress_bar = st.progress(0)
        status_message = st.empty()
        
        try:
            status_message.info("Initializing processing...")
            progress_bar.progress(10)
            
            processed_data = []
            total_sheets = len(selected_sheets)
            for i, sheet_name in enumerate(selected_sheets):
                status_message.info(f"Processing sheet {i+1}/{total_sheets}: {sheet_name}...")
                progress = 10 + int((i / total_sheets) * 70)
                progress_bar.progress(progress)
                
                df = process_sheet(uploaded_file, sheet_name)
                if df is not None:
                    processed_data.append((sheet_name, df))
            
            status_message.info("Finalizing output...")
            progress_bar.progress(90)
            
            # Generate output based on format
            if output_format == "Single File (multiple sheets)":
                if download_format == "XLSX":
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        for sheet_name, df in processed_data:
                            safe_sheet_name = sheet_name[:31]  # Excel sheet name limit
                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    output.seek(0)
                    st.success("✅ Processing completed successfully!")
                    st.download_button(
                        label="⬇️ Download Excel File",
                        data=output,
                        file_name="processed_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:  # CSV for multiple sheets
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for sheet_name, df in processed_data:
                            csv_buffer = io.StringIO()
                            df.to_csv(csv_buffer, index=False)
                            zip_file.writestr(f"{sheet_name}.csv", csv_buffer.getvalue())
                    zip_buffer.seek(0)
                    st.success("✅ Processing completed successfully!")
                    st.download_button(
                        label="⬇️ Download ZIP of CSV Files",
                        data=zip_buffer,
                        file_name="processed_sheets.zip",
                        mime="application/zip",
                    )

            elif output_format == "Single File (one combined sheet)":
                combined_df = pd.concat([df for _, df in processed_data], ignore_index=True)
                if download_format == "XLSX":
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        combined_df.to_excel(writer, sheet_name='Combined_Sheet', index=False)
                    output.seek(0)
                    st.download_button(
                        label="⬇️ Download Combined Excel File",
                        data=output,
                        file_name="combined_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    csv_output = combined_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="⬇️ Download Combined CSV File",
                        data=csv_output,
                        file_name="combined_output.csv",
                        mime="text/csv",
                    )

            else:  # Multiple Files (per sheet per file)
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for sheet_name, df in processed_data:
                        if download_format == "XLSX":
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                df.to_excel(writer, index=False, sheet_name='Sheet1')
                            zip_file.writestr(f"{sheet_name}.xlsx", excel_buffer.getvalue())
                        else:
                            csv_buffer = io.StringIO()
                            df.to_csv(csv_buffer, index=False)
                            zip_file.writestr(f"{sheet_name}.csv", csv_buffer.getvalue())
                zip_buffer.seek(0)
                st.download_button(
                    label="⬇️ Download ZIP File",
                    data=zip_buffer,
                    file_name="processed_files.zip",
                    mime="application/zip",
                )

            progress_bar.progress(100)
            st.balloons()
            
        except Exception as e:
            progress_bar.empty()
            status_message.error(f"❌ Processing failed: {str(e)}")
            st.exception(e)

st.markdown("---")
st.markdown("### Instructions")
st.markdown("""
1. Upload your Excel file using the uploader above
2. Select the worksheets you want to process
3. Choose output format (single file with multiple sheets or multiple files)
4. Click the 'Start Processing' button
5. Download your processed file(s) when ready
""")
footer_html = """<div style='text-align: center;'>
  <p>Developed with 🔥 by Puti Andini</p>
</div>"""
st.markdown(footer_html, unsafe_allow_html=True)