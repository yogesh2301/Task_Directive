import fitz
from PyQt6.QtGui import QImage, QPixmap
from PyQt6.QtCore import Qt


class PDFHandler:
    def __init__(self):
        self.doc = None
        self.page_data = []

    def open_pdf(self, path):
        self.doc = fitz.open(path)
        full_text = "".join(page.get_text() for page in self.doc)
        return full_text

    def extract_pages(self, max_pages=15):
        self.page_data.clear()
        for i, page in enumerate(self.doc):
            if i < max_pages:
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                img = QImage(
                    pix.samples, pix.width, pix.height,
                    pix.stride, QImage.Format.Format_RGB888
                ).copy()
                self.page_data.append((QPixmap.fromImage(img), i))
        return self.page_data

    def get_page(self, index):
        if self.doc:
            return self.doc.load_page(index)
        return None

    def close(self):
        if self.doc:
            self.doc.close()
            self.doc = None