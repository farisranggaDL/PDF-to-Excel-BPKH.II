import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
import pandas as pd
import io
import platform
import shutil

st.set_page_config(page_title="PDF OCR to Excel", layout="centered")
st.title("📄➡️📊 Scanned PDF to Excel (OCR)")
st.write("Upload a **scanned PDF**, extract text using OCR, and download it as an Excel file.")

# -------------------------
# Auto-detect Tesseract
# -------------------------
def setup_tesseract():
    tesseract_path = shutil.which("tesseract")
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
    else:
        st.error(
            "⚠️ Tesseract OCR not found on this system.\n\n"
            "Please install Tesseract:\n"
            "- Ubuntu/Debian: `sudo apt-get install tesseract-ocr`\n"
            "- macOS: `brew install tesseract`\n"
            "- Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki"
        )
        st.stop()

setup_tesseract()

pytesseract.image_to_string(image, lang="eng+ind")

# -------------------------
# Upload PDF
# -------------------------
uploaded_file = st.file_uploader(
    "Upload a scanned PDF",
    type=["pdf"]
)

# -------------------------
# OCR function
# -------------------------
def ocr_pdf_to_text(pdf_bytes):
    images = convert_from_bytes(pdf_bytes)
    extracted_text = []

    for page_num, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image, lang="eng")  # add other languages if needed
        extracted_text.append({
            "Page": page_num,
            "Text": text.strip()
        })

    return extracted_text

# -------------------------
# Process uploaded file
# -------------------------
if uploaded_file is not None:
    with st.spinner("Running OCR... this may take a moment ⏳"):
        pdf_bytes = uploaded_file.read()
        ocr_results = ocr_pdf_to_text(pdf_bytes)

    st.success("OCR complete!")

    # Convert to DataFrame
    df = pd.DataFrame(ocr_results)

    st.subheader("Preview Extracted Text")
    st.dataframe(df, use_container_width=True)

    # Save to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="OCR Output")

    st.download_button(
        label="⬇️ Download Excel File",
        data=output.getvalue(),
        file_name="ocr_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

