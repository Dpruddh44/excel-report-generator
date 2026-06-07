import streamlit as st
import os

# Import the core backend script
import generate_report

st.set_page_config(page_title="Ops Report Automation", page_icon="📊", layout="wide")

# Hide default Streamlit elements for a cleaner UI
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Main Title & Subtitle
st.title("📊 Ops Report Automation")
st.markdown("### Streamline your reporting workflow")
st.write("Upload the three required files below to automatically generate the consolidated Ops Report.")

st.divider()

# Uploaders Section
st.subheader("📁 1. Data Sources")
col1, col2, col3 = st.columns(3)

with col1:
    st.info("**📋 EAID Resource List**")
    eaid_file = st.file_uploader("Upload EAID File", type=["xlsx"], help="Upload the EAID Resource List Excel file")

with col2:
    st.info("**📝 DMM Ops Template**")
    template_file = st.file_uploader("Upload Template File", type=["xlsx"], help="Upload the DMM Ops Report Template")

with col3:
    st.info("**💰 MTD Revenue**")
    rbr_file = st.file_uploader("Upload Revenue File", type=["xlsx"], help="Upload the MTD Revenue by Resource")

st.divider()

# Action Section
st.subheader("⚙️ 2. Generation")

# Wide primary button
if st.button("✨ Generate Report", type="primary", use_container_width=True):
    if eaid_file and template_file and rbr_file:
        with st.status("Generating Report...", expanded=True) as status:
            try:
                st.write("💾 Saving uploaded files...")
                # The backend script expects these exact filenames on disk
                with open("EAID AT 7 May for Ops team.xlsx", "wb") as f:
                    f.write(eaid_file.getbuffer())
                
                with open("DMM Ops Report_1-29 April (as on 30 April).xlsx", "wb") as f:
                    f.write(template_file.getbuffer())
                
                with open("MTD Revenue_by_Resource_11-05-2026.xlsx", "wb") as f:
                    f.write(rbr_file.getbuffer())
                
                st.write("⚙️ Processing data through backend logic...")
                # Execute the backend core logic
                generate_report.main()
                
                st.write("✅ Finalizing output...")
                # Check if output file was successfully generated
                if os.path.exists("final_report.xlsx"):
                    status.update(label="Report generated successfully!", state="complete", expanded=False)
                    st.success("🎉 Process completed successfully!")
                    
                    with open("final_report.xlsx", "rb") as f:
                        st.download_button(
                            label="📥 Download Final Report",
                            data=f,
                            file_name="final_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                else:
                    status.update(label="Report generation failed.", state="error", expanded=True)
                    st.error("❌ The final output file was not created.")

            except Exception as e:
                status.update(label="An error occurred.", state="error", expanded=True)
                st.error(f"⚠️ An error occurred during report generation: {e}")
    else:
        st.warning("⚠️ Please upload all three files in the Data Sources section before generating the report.")
