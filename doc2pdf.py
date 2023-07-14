import streamlit as st
from docx import Document
from docx2pdf import convert
import io

def convert_to_pdf(file):
    # Convert the document to a DOCX file in memory
    docx_file = file.name + ".docx"
    document = Document(file)
    docx_data = io.BytesIO()
    document.save(docx_data)
    docx_data.seek(0)

    # Convert the DOCX file to PDF in memory
    pdf_data = convert(docx_data)

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
