import os
import zipfile
import streamlit as st
from salary_slip import generate_salary_slip

# Title of the web app
st.title('Salary Slip Generator')

# Upload an Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Allow user to upload logo (optional)
    logo_path = "https://github.com/sjpradhan/salaryslip_generator/blob/master/Bankbencher%20logo.png"

    # Temporary directory to store the generated PDFs
    output_dir = "static/"

    # Process the file when uploaded
    if st.button("Generate Salary Slips"):
        with st.spinner("Generating salary slips..."):
            # Save the uploaded file locally
            with open("uploaded_file.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Call the generate_salary_slip function
            pdf_files = generate_salary_slip("uploaded_file.xlsx", logo_path, output_dir)

            # Create a zip file containing all PDFs
            zip_filename = "salary_slips.zip"
            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                for pdf_file in pdf_files:
                    zipf.write(pdf_file, os.path.basename(pdf_file))

            st.success("Salary slips generated successfully!")

            # Provide the zip file for download
            with open(zip_filename, "rb") as f:
                st.download_button("Download Salary Slips", f, file_name=zip_filename)

            # Optionally, clean up the temporary files
            os.remove("uploaded_file.xlsx")
            for pdf_file in pdf_files:
                os.remove(pdf_file)
            os.remove(zip_filename)
