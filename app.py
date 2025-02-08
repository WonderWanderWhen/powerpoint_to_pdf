import streamlit as st
from pptx import Presentation
import pdfkit
import tempfile
import os

def pptx_to_pdf(input_file):
    """
    Converts a PowerPoint file (PPT/PPTX) to PDF.
    """
    # Create a temporary file for the PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        pdf_path = tmp_file.name

    try:
        # Save the uploaded file temporarily
        with open("temp.pptx", "wb") as f:
            f.write(input_file.getbuffer())

        # Convert the PPTX file to PDF using pdfkit
        pdfkit.from_file("temp.pptx", pdf_path)
    except Exception as e:
        st.error(f"Conversion failed: {e}")
        return None
    finally:
        # Clean up the temporary PPTX file
        if os.path.exists("temp.pptx"):
            os.remove("temp.pptx")

    return pdf_path

def main():
    """
    Main function to run the Streamlit app.
    """
    st.title("PowerPoint to PDF Converter")
    st.write("Upload a PowerPoint file (PPT/PPTX) to convert it to PDF.")

    # File uploader
    uploaded_file = st.file_uploader("Choose a PowerPoint file", type=["pptx", "ppt"])

    if uploaded_file is not None:
        st.write("File uploaded successfully!")

        # Convert to PDF
        pdf_path = pptx_to_pdf(uploaded_file)

        if pdf_path:
            st.success("Conversion successful!")
            st.write("Download your PDF file below:")

            # Provide download link
            with open(pdf_path, "rb") as f:
                st.download_button(
                    label="Download PDF",
                    data=f,
                    file_name="converted.pdf",
                    mime="application/pdf"
                )

            # Clean up the temporary PDF file
            os.remove(pdf_path)

if __name__ == "__main__":
    main()
