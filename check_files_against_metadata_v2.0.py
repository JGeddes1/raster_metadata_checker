import os
import streamlit as st
import openpyxl

# Set page configuration with favicon
st.set_page_config(
    page_title="Check Files Against Metadata",
   
)


# Your main app content
st.title("Check Files Against Metadata")
st.write(" upload images and Excel files to compare metadata.")


# Function to process Excel files and extract filenames and keywords
def read_excel_file(file):
    filenames = []
    subjectkeywords = set()
    excluded_keywords = {"subject keyword 1", "subject keyword 2", "subject keyword 3"}  # Titles to exclude
    workbook = openpyxl.load_workbook(file)
    worksheet = workbook.active
    for row in worksheet.iter_rows(values_only=True):
        if row[0] and not str(row[0]).lower() in ["filename"]:
            filenames.append(row[0].strip().lower())
        # Check and add keywords only if they are not in the excluded list
        for keyword in row[2:5]:  # Assuming columns C, D, and E contain keywords
            if keyword and keyword.strip().lower() not in excluded_keywords:
                subjectkeywords.add(keyword.strip().lower())
    return filenames, subjectkeywords


# Function to compare uploaded files against Excel data
def find_missing_files(uploaded_files, excel_filenames):
    uploaded_filenames = [os.path.basename(file.name).lower() for file in uploaded_files]
    missing_in_directory = [filename for filename in excel_filenames if filename not in uploaded_filenames]
    missing_in_metadata = [filename for filename in uploaded_filenames if filename not in excel_filenames]
    return missing_in_directory, missing_in_metadata



# File upload section
uploaded_files = st.file_uploader("Upload image files (multiple allowed)", type=["jpg", "jpeg", "png", "gif", "bmp", "tiff"], accept_multiple_files=True)

# Excel file upload section
uploaded_excel_file1 = st.file_uploader("Upload first Excel file", type=["xlsx"])
uploaded_excel_file2 = st.file_uploader("Upload second Excel file (optional)", type=["xlsx"])

if st.button("Check Files"):
    if uploaded_files and uploaded_excel_file1:
        # Read filenames from Excel files
        excel_filenames1, subjectkeywords1 = read_excel_file(uploaded_excel_file1)
        if uploaded_excel_file2:
            excel_filenames2, subjectkeywords2 = read_excel_file(uploaded_excel_file2)
            excel_filenames = list(set(excel_filenames1 + excel_filenames2))
            subjectkeywords = subjectkeywords1.union(subjectkeywords2)
        else:
            excel_filenames = excel_filenames1
            subjectkeywords = subjectkeywords1

        # Compare uploaded files against Excel filenames
        missing_in_directory, missing_in_metadata = find_missing_files(uploaded_files, excel_filenames)

        # Display results
        st.subheader("Results")
        st.write("**Missing from directory:**")
        st.write(missing_in_directory if missing_in_directory else "No files missing.")
        st.write("**Not listed in metadata:**")
        st.write(missing_in_metadata if missing_in_metadata else "No extra files found.")
        st.write("**Subject Keywords:**")
        st.write(list(subjectkeywords))
    else:
        st.error("Please upload at least one Excel file and some image files.")
