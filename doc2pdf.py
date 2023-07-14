import streamlit as st
from docx import Document
from docx2pdf import convert
import tempfile
import os

def convert_to_pdf(file):
    # Convert the document to a DOCX file on disk
    temp_dir = tempfile.mkdtemp()
    docx_file = os.path.join(temp_dir, file.name + ".docx")
    document = Document(file)
    document.save(docx_file)

    # Convert the DOCX file to PDF
    pdf_file = docx_file[:-5] + ".pdf"
    convert(docx_file, pdf_file)

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
