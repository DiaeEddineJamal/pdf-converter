import streamlit as st
import tempfile
import os
import subprocess

def convert_to_pdf(file):
    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()

    # Save the uploaded document to a temporary file
    docx_file = os.path.join(temp_dir, file.name)
    with open(docx_file, "wb") as f:
        f.write(file.read())

    # Convert the DOCX file to PDF using unoconv
    pdf_file = docx_file.replace(".docx", ".pdf")
    subprocess.run(["unoconv", "-f", "pdf", docx_file], check=True)

    # Read the PDF file data
    with open(pdf_file, "rb") as f:
        pdf_data = f.read()

    # Cleanup: remove temporary files
    os.remove(docx_file)
    os.remove(pdf_file)
    os.rmdir(temp_dir)

    # Return the PDF file data
    return pdf_data

# Streamlit app
def main():
    st.title("Document to PDF Converter")

    # File uploader
    file = st.file_uploader("Upload a document", type=[".doc", ".docx"])

    if file is not None:
        if st.button("Convert"):
            pdf_data = convert_to_pdf(file)

            # Download link for the converted PDF file
            st.download_button("Download PDF", pdf_data, file.name + ".pdf")

if __name__ == "__main__":
    main()
