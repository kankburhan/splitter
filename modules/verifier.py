import streamlit as st
import pandas as pd
import time

def verify_excel(file):
    try:
        df = pd.read_excel(file, sheet_name='Sheet1')
    except ValueError:
    # Load the first available sheet if 'Sheet1' not found
        sheet_names = pd.ExcelFile(file).sheet_names
        df = pd.read_excel(file, sheet_name=sheet_names[0])

    try:
        df['TEST SCRIPT NUMBER'] = df['TEST SCRIPT NUMBER'].ffill()
        
        invalid_scripts = []
        
        for script_id, group in df.groupby('TEST SCRIPT NUMBER'):
            test_steps = group['Execution_Sequence'].count()
            expected_results = group['Expected_Result'].count()
            
            if test_steps != expected_results:
                invalid_scripts.append({
                    'File Name': file.name,
                    'Script ID': script_id,
                    'Test Steps': test_steps,
                    'Expected Results': expected_results
                })
        
        return pd.DataFrame(invalid_scripts)
    
    except Exception as e:
        return pd.DataFrame({
            'File Name': [file.name],
            'Error': [f"Processing error: {str(e)}"]
        })

def render_verifier():
    st.title('📊 Multi-File Test Script Verifier')
    st.caption("Developed with Streamlit 🚀")
    footer_html = """<div style='text-align: center;'>
      <p>Developed with 🔥 by Puti Andini</p>
    </div>"""
    st.markdown(footer_html, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Upload Excel files (multiple allowed)",
        type=["xlsx"],
        accept_multiple_files=True
    )

    # Debugging: Print the list of uploaded files
    if uploaded_files:
        st.write(f"Uploaded files: {[file.name for file in uploaded_files]}")

    if st.button('🔍 Verify Files', type='primary'):
        if not uploaded_files:
            st.warning("⚠️ Please upload at least one Excel file")
        else:
            progress_bar = st.progress(0)
            status_container = st.empty()
            all_results = []
            total_files = len(uploaded_files)
            
            with st.spinner('🚀 Processing files...'):
                for i, file in enumerate(uploaded_files):
                    progress = int((i + 1) / total_files * 100)
                    progress_bar.progress(progress)
                    
                    status_container.markdown(f"""
                    🛠 **Processing File {i+1}/{total_files}**  
                    📄 File Name: `{file.name}`  
                    ⏳ Status: Analyzing...
                    """)
                    
                    time.sleep(0.1)
                    file_results = verify_excel(file)
                    
                    if not file_results.empty:
                        all_results.append(file_results)
                    
                    status_container.markdown(f"""
                    ✅ **Completed File {i+1}/{total_files}**  
                    📄 File Name: `{file.name}`  
                    🕒 Status: Analysis completed
                    """)
                    time.sleep(0.2)

            progress_bar.empty()
            status_container.empty()

            if all_results:
                combined_results = pd.concat(all_results)
                
                st.error("❌ Validation Results")
                st.write("### Invalid Scripts Grouped by File")
                
                # Group results by file name
                grouped = combined_results.groupby('File Name')
                for file_name, group in grouped:
                    with st.expander(f"📁 {file_name} ({len(group)} issues)", expanded=False):
                        st.dataframe(group.drop(columns=['File Name']).reset_index(drop=True))
                
                st.write("### Summary Statistics")
                summary = combined_results.groupby('File Name').size().reset_index(name='Total Issues')
                st.dataframe(summary)
                
                st.download_button(
                    label="📥 Download Full Report",
                    data=combined_results.to_csv(index=False).encode('utf-8'),
                    file_name='validation_report.csv',
                    mime='text/csv'
                )
            else:
                st.balloons()
                st.success("🎉 All files passed validation!")

            st.markdown("---")
            st.caption(f"Validation completed at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            