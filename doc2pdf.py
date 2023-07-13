import streamlit as st
import os
import tempfile
import requests
from pptx import Presentation
import pdf2image
from PIL import Image
import pdfkit
import asyncio


def convert_docx_to_pdf(file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(file.read())
        tmp.close()
        pdf_file = tmp.name.replace(".docx", ".pdf")
        convert_using_api(tmp.name, pdf_file, "docx", "pdf")
        return pdf_file


def convert_using_api(input_file, output_file, input_format, output_format):
    api_key = "163163096"  # Replace with your ConvertAPI API key
    url = f"https://v2.convertapi.com/convert/{input_format}/to/{output_format}"
    files = {"file": open(input_file, "rb")}
    payload = {"ApiKey": api_key}
    response = requests.post(url, files=files, data=payload)
    response.raise_for_status()
    with open(output_file, "wb") as output:
        output.write(response.content)


async def convert_pptx_to_pdf(file):
    prs = Presentation(file)
    html_path = tempfile.NamedTemporaryFile(delete=False, suffix=".html").name
    pdf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name

    slides_html = []
    for slide in prs.slides:
        slide_content = extract_slide_content(slide)
        slides_html.append(slide_content)

    html_content = "".join(slides_html)

    with open(html_path, "w", encoding="utf-8") as html_file:
        html_file.write(html_content)

    options = {
        "quiet": "",
        "no-outline": None
    }

    await asyncio.sleep(0)  # Allow event loop to run
    pdfkit.from_file(html_path, pdf_path, options=options)

    return pdf_path


def extract_slide_content(slide):
    shapes = slide.shapes
    content = ""
    for shape in shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    content += run.text
                    content += " "
        content += "<br>"
    return content


def main():
    st.title("Document to PDF Converter :notebook:")

    # Upload the document file
    file = st.file_uploader("Upload a document file")

    with st.spinner("Converting..."):
        if file is not None:
            file_extension = os.path.splitext(file.name)[1].lower()

            if file_extension == ".docx":
                # Convert DOCX to PDF
                if st.button("Convert to PDF"):
                    progress_bar = st.progress(0)
                    pdf_file = convert_docx_to_pdf(file)
                    progress_bar.progress(100)
                    st.success("Conversion successful!")
                    st.download_button("Download PDF", data=open(pdf_file, "rb"), file_name="converted.pdf")

            elif file_extension == ".pptx":
                # Convert PPTX to PDF
                if st.button("Convert to PDF"):
                    progress_bar = st.progress(0)
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    pdf_file = loop.run_until_complete(convert_pptx_to_pdf(file))
                    progress_bar.progress(100)
                    st.success("Conversion successful!")
                    st.download_button("Download PDF", data=open(pdf_file, "rb"), file_name="converted.pdf")

            else:
                st.warning("Invalid file format. Please upload a supported file type (DOCX, PPTX or PDF).")
    st.subheader("How to convert:")
    st.video("./media/pdfx.mp4")


if __name__ == "__main__":
    main()
