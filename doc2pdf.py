import streamlit as st
from docx import Document
from docx2pdf import convert

def convert_to_pdf(file):
    # Convert the document to a DOCX file
    docx_file = f"{file.name}.docx"
    document = Document(file)
    document.save(docx_file)

    # Convert the DOCX file to PDF
    pdf_file = f"{file.name}.pdf"
    convert(docx_file, pdf_file)

    # Return the PDF file path
    return pdf_file

# Streamlit app
def main():
    st.title("Document to PDF Converter")

    # File uploader
    file = st.file_uploader("Upload a document", type=[".doc", ".docx"])

    if file is not None:
        if st.button("Convert"):
            pdf_file = convert_to_pdf(file)

            # Download link for the converted PDF file
            st.download_button("Download PDF", pdf_file)

if __name__ == "__main__":
    main()
