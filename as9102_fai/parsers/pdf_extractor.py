import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import logging

from as9102_fai.ocr_utils import configure_pytesseract

logger = logging.getLogger(__name__)

class PdfTextExtractor:
    def __init__(self, tesseract_cmd=None):
        # Configure pytesseract to find the tesseract executable on Windows.
        # If not found, OCR calls will raise a clear error which callers can handle.
        configure_pytesseract(tesseract_cmd)

    def extract_text(self, pdf_path: str) -> str:
        full_text = ""
        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                # Try standard text extraction first
                text = page.get_text()
                
                if not text.strip():
                    # If empty, try OCR
                    pix = page.get_pixmap()
                    img_data = pix.tobytes("png")
                    image = Image.open(io.BytesIO(img_data))
                    text = pytesseract.image_to_string(image)
                
                full_text += text + "\n"
        except Exception as e:
            logger.exception("Error extracting text from PDF")
            return ""
            
        return full_text
