import streamlit as st
import pandas as pd
import os
from processor import run_automation

# --- Streamlit Page Configuration ---
st.set_page_config(
    page_title="Data Enrichment Tool",
    page_icon="🤖",
    layout="wide"
)

# --- UI Layout ---
st.title("📄 Data Enrichment Automation Tool (Selenium Pro Version)")
st.write("This tool uses a real browser in the background to find data, just like a human would. This is the most reliable method.")
st.info("Note: The first run might be slow as it downloads the correct browser driver.")


uploaded_file = st.file_uploader(
    "Choose an Excel file with 'Company' and 'Contacts' sheets",
    type=['xlsx']
)

if uploaded_file is not None:
    st.success(f"File '{uploaded_file.name}' uploaded successfully!")
    
    temp_file_path = os.path.join("output", uploaded_file.name)
    with open(temp_file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("🚀 Run Enrichment Automation", use_container_width=True):
        
        status_placeholder = st.empty()
        results_placeholder = st.empty()

        def update_status(message):
            status_placeholder.info(message)

        try:
            with st.spinner("Processing with automated browser... This may take several minutes."):
                excel_path, json_path, df_companies, df_contacts = run_automation(
                    temp_file_path, 
                    update_status
                )
            
            status_placeholder.success("✅ Automation complete! See results below.")
            
            with results_placeholder.container():
                st.subheader("Enriched Companies")
                st.dataframe(df_companies)
                
                st.subheader("Enriched Contacts")
                st.dataframe(df_contacts)

                st.divider()
                st.subheader("Download Your Files")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    with open(excel_path, "rb") as file:
                        st.download_button(
                            label="📥 Download Enriched Excel File",
                            data=file,
                            file_name="Enriched_Results.xlsx",
                            mime="application/vnd.ms-excel",
                            use_container_width=True
                        )
                
                with col2:
                    with open(json_path, "rb") as file:
                        st.download_button(
                            label="📥 Download Enriched JSON File",
                            data=file,
                            file_name="enriched_data.json",
                            mime="application/json",
                            use_container_width=True
                        )

        except Exception as e:
            st.error(f"An error occurred: {e}")
            st.error("Please check your file format and column names. It must be an Excel file with 'Company' and 'Contacts' sheets.")

        finally:
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
