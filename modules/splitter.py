import streamlit as st
import pandas as pd
import io
import zipfile
import re

def funding_transform(df, epic_link='', feature='', squad='', priority='High'):
    df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    
    # Dynamically find the row containing the desired column headers
    required_columns = {
        'NO.', 
        'TEST SCRIPT NUMBER',
        'TEST SCRIPT DESCRIPTION/SCENARIO',
        'TEST OBJECT NAME',
        'Scenario Type',
        'GENERAL INFORMATION / SUMMARY OF THE TEST SCRIPT',
        'PRE-REQUISITES',
        'Execution_Sequence',
        'Expected_Result',
        'Product / Akad',
        'Feature',
        'Assigned Tester'
    }
    header_row_index = None
    for i, row in df.iterrows():
        row_set = set(row.dropna().astype(str))  # Drop NaN and convert to strings
        if len(required_columns.intersection(row_set)) >= len(required_columns) * 0.6:  # Match at least 60% of required columns
            header_row_index = i
            break
    
    if header_row_index is None:
        raise ValueError("Could not find the header row with the required columns.")
    
    # Set proper headers and drop unnecessary rows
    df.columns = df.iloc[header_row_index]
    df = df.iloc[header_row_index + 1:].reset_index(drop=True)
    
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
        'Expected_Result',
        'Product / Akad',
        'Feature',
        'Assigned Tester'
    ]
    df = df[[col for col in columns_to_keep if col in df.columns]]  # Keep only available columns
    
    # Align lists using vectorized operations
    def align_lists(row):
        exec_seq = re.split(r'\n(?=\d+\.\s*)', row['Execution_Sequence']) if isinstance(row['Execution_Sequence'], str) else []
        exp_res = re.split(r'\n(?=\d+\.\s*)', row['Expected_Result']) if isinstance(row['Expected_Result'], str) else []
        max_len = max(len(exec_seq), len(exp_res))
        exec_seq.extend([''] * (max_len - len(exec_seq)))
        exp_res.extend([''] * (max_len - len(exp_res)))
        return pd.DataFrame({'Execution_Sequence': exec_seq, 'Expected_Result': exp_res})
    
    exploded_dfs = []
    for _, row in df.iterrows():
        aligned_df = align_lists(row)
        aligned_df = aligned_df.assign(**{col: row[col] for col in df.columns if col not in ['Execution_Sequence', 'Expected_Result']})
        exploded_dfs.append(aligned_df)
    
    df_exploded = pd.concat(exploded_dfs, ignore_index=True)
    
    # Add additional columns directly
    df_exploded['Test Envi'] = 'SIT'
    df_exploded['Test Phase'] = 'SIT1'
    df_exploded['epic link'] = epic_link
    df_exploded['summary'] = df_exploded['TEST SCRIPT NUMBER'] + '_' + df_exploded['GENERAL INFORMATION / SUMMARY OF THE TEST SCRIPT']
    df_exploded['feature'] = feature
    df_exploded['squad'] = squad
    df_exploded['Priority'] = priority
    
    # Group by "TEST SCRIPT NUMBER" and assign "NO." starting from 0 for each group
    df_exploded['NO.'] = df_exploded.groupby('TEST SCRIPT NUMBER').ngroup()
    df_exploded['TEST SCRIPT NUMBER'] = df_exploded['TEST SCRIPT NUMBER'].ffill()
    
    # Clear duplicate values for specific columns
    cols_to_clear = [
        'TEST SCRIPT NUMBER',
    ]
    df_exploded.loc[df_exploded.duplicated(subset=['TEST SCRIPT NUMBER']), cols_to_clear] = ''

    # Fill specific columns only in the first row of each TEST SCRIPT NUMBER
    cols_to_fill_first_row = [
    ]
    for col in cols_to_fill_first_row:
        df_exploded[col] = df_exploded.groupby('TEST SCRIPT NUMBER')[col].transform(
            lambda x: [x.iloc[0]] + [''] * (len(x) - 1)
        )
    
    # Reorder columns to match the specified ordering
    final_columns = [
        'NO.', 
        'TEST SCRIPT NUMBER',
        'TEST SCRIPT DESCRIPTION/SCENARIO',
        'TEST OBJECT NAME',
        'Scenario Type',
        'GENERAL INFORMATION / SUMMARY OF THE TEST SCRIPT',
        'PRE-REQUISITES',
        'Execution_Sequence',
        'Expected_Result',
        'Product / Akad',
        'Feature',
        'Assigned Tester',
        'Test Envi',
        'Test Phase',
        'epic link',
        'summary',
        'feature',
        'squad',
        'Priority'
    ]
    return df_exploded[final_columns]

def financing_transform(df, epic_link='', feature='', squad='', priority='High'):
    return funding_transform(df, epic_link, feature, squad, priority)

def process_sheet(input_file, sheet_name, epic_link, feature, squad, priority):
    # Load the Excel file
    xls = pd.ExcelFile(input_file)
    df = xls.parse(sheet_name=sheet_name)
    
    # Drop empty rows and columns
    filename_lower = input_file.name.lower()
    if '[funding]' in filename_lower or '[t24]' in filename_lower:
        df = funding_transform(df, epic_link, feature, squad, priority)
    elif '[financing]' in filename_lower:
        df = financing_transform(df, epic_link, feature, squad, priority)
    else:
         df = funding_transform(df, epic_link, feature, squad, priority)
    
    return df

def render_splitter():
    st.title("📊 Excel Processing Tool")

    # Add new form fields
    epic_link = st.text_input("Epic Link", help="Enter the epic link")
    feature = st.text_input("Feature", help="Enter the feature name")
    squad = st.selectbox(
        "Squad",
        [
            "",
            "Customer", 
            "Funding", 
            "Branch Transaction", 
            "Financing Retail", 
            "Financing Corporate", 
            "Financing-After DIsbursement",
            "Micro", 
            "Gadai", 
            "Head Office Transaction", 
            "Accounting"
        ],
        help="Select the squad"
    )
    # File Upload Section
    uploaded_files = st.file_uploader(
        "Upload your Excel file(s)", 
        type=["xlsx"], 
        accept_multiple_files=True,
        help="Please upload one or more Excel files in XLSX format"
    )

    if uploaded_files:
        # Sheet Selection
        try:
            output_format = st.radio(
                "Output format", 
                ["Single File (multiple sheets)", "Multiple Files (per sheet per file)", "Single File (one combined sheet)"]
            )
            download_format = st.radio(
                "Select file format:",
                ["XLSX", "CSV"],
                key="file_format"
            )
            
            all_sheet_names = {}
            for uploaded_file in uploaded_files:
                xls = pd.ExcelFile(uploaded_file)
                all_sheet_names[uploaded_file.name] = xls.sheet_names
            
            selected_sheets = {}
            for file_name, sheet_names in all_sheet_names.items():
                selected_sheets[file_name] = st.multiselect(f"Select worksheets for {file_name}", sheet_names)
            
        except Exception as e:
            st.error(f"Error reading files: {str(e)}")
            st.stop()

        # Processing Section
        if st.button("🚀 Start Processing"):
            if not any(selected_sheets.values()):
                st.error("Please select at least one worksheet from the uploaded files.")
                st.stop()
            
            progress_bar = st.progress(0)
            status_message = st.empty()
            
            try:
                status_message.info("Initializing processing...")
                progress_bar.progress(10)
                
                processed_data = []
                total_sheets = sum(len(sheets) for sheets in selected_sheets.values())
                processed_count = 0
                
                for file_name, sheets in selected_sheets.items():
                    for sheet_name in sheets:
                        status_message.info(f"Processing sheet {processed_count + 1}/{total_sheets}: {sheet_name} from {file_name}...")
                        progress = 10 + int((processed_count / total_sheets) * 70)
                        progress_bar.progress(progress)
                        
                        uploaded_file = next(f for f in uploaded_files if f.name == file_name)
                        df = process_sheet(uploaded_file, sheet_name, epic_link, feature, squad, 'High')
                        if df is not None:
                            processed_data.append((sheet_name, df))  # Keep only the sheet name
                        
                        processed_count += 1
                
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

                elif output_format == "Multiple Files (per sheet per file)":
                    try:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for sheet_name, df in processed_data:
                                if download_format == "XLSX":
                                    excel_buffer = io.BytesIO()
                                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])  # Preserve sheet name
                                    zip_file.writestr(f"{sheet_name}.xlsx", excel_buffer.getvalue())
                                else:
                                    csv_buffer = io.StringIO()
                                    df.to_csv(csv_buffer, index=False)
                                    zip_file.writestr(f"{sheet_name}.csv", csv_buffer.getvalue())
                        zip_buffer.seek(0)
                        st.success("✅ Processing completed successfully!")
                        st.download_button(
                            label="⬇️ Download ZIP File",
                            data=zip_buffer,
                            file_name="processed_files.zip",
                            mime="application/zip",
                        )
                    except Exception as e:
                        st.error(f"❌ Error during ZIP creation: {str(e)}")
                        st.exception(e)

                progress_bar.progress(100)
                st.balloons()
                
            except Exception as e:
                progress_bar.empty()
                status_message.error(f"❌ Processing failed: {str(e)}")
                st.exception(e)

    st.markdown("---")
    st.markdown("### Instructions")
    st.markdown("""
    1. Upload one or more Excel files using the uploader above
    2. Select the worksheets you want to process from each file
    3. Choose output format (single file with multiple sheets, multiple files, or one combined sheet)
    4. Click the 'Start Processing' button
    5. Download your processed file(s) when ready
    """)
    footer_html = """<div style='text-align: center;'>
      <p>Developed with 🔥 by Puti Andini</p>
    </div>"""
    st.markdown(footer_html, unsafe_allow_html=True)