import os
import fitz
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
    QLabel, QScrollArea, QWidget,QGridLayout, QLineEdit, QProgressBar
)
from PyQt6.QtCore import Qt, QUrl, pyqtSignal
from PyQt6.QtGui import QImage, QPixmap

try:
    from PyQt6.QtWebEngineWidgets import QWebEngineView
    from PyQt6.QtWebEngineCore import QWebEngineSettings
    HAS_WEBENGINE = True
except ImportError:
    HAS_WEBENGINE = False


class PDFViewerDialog(QDialog):
    def __init__(self, pdf_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Generated PDF Preview")
        self.resize(1000, 800)
        self.pdf_path = pdf_path
        self.zoom_factor = 1.0
        self.fallback_doc = None
        self.fallback_page_labels = []
        self.setStyleSheet("""
            QDialog { background-color: #f8fafc; color: #191c1e; }
            QPushButton#ZoomBtn {
                background-color: #e2e8f0; color: #334155;
                font-size: 16px; font-weight: bold;
                border-radius: 6px; padding: 6px 14px;
                border: 1px solid #cbd5e1;
            }
            QPushButton#ZoomBtn:hover { background-color: #cbd5e1; color: #00264d; }
            QPushButton#ZoomBtn:pressed { background-color: #94a3b8; color: #ffffff; }
            QPushButton#CloseBtn { 
                background-color: #64748b; color: #ffffff; 
                font-size: 14px; font-weight: bold; 
                border-radius: 6px; padding: 10px 30px; border: none; 
            }
            QPushButton#CloseBtn:hover { background-color: #475569; }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        toolbar_layout = QHBoxLayout()
        toolbar_layout.addStretch()
        zoom_out_btn = QPushButton("-")
        zoom_out_btn.setObjectName("ZoomBtn")
        zoom_out_btn.setToolTip("Zoom out")
        zoom_out_btn.clicked.connect(self.zoom_out)
        zoom_in_btn = QPushButton("+")
        zoom_in_btn.setObjectName("ZoomBtn")
        zoom_in_btn.setToolTip("Zoom in")
        zoom_in_btn.clicked.connect(self.zoom_in)
        toolbar_layout.addWidget(zoom_out_btn)
        toolbar_layout.addWidget(zoom_in_btn)
        layout.addLayout(toolbar_layout)
        
        if HAS_WEBENGINE:
            self.viewer = QWebEngineView()
            self.viewer.settings().setAttribute(
                QWebEngineSettings.WebAttribute.PdfViewerEnabled, True
            )
            self.viewer.setUrl(QUrl.fromLocalFile(os.path.abspath(pdf_path)))
            layout.addWidget(self.viewer)
        else:
            self._add_fallback_viewer(layout, pdf_path)

        close_btn = QPushButton("Close Preview")
        close_btn.setObjectName("CloseBtn")
        close_btn.clicked.connect(self.accept)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def zoom_in(self):
        self.zoom_factor = min(self.zoom_factor + 0.15, 3.0)
        self.apply_zoom()

    def zoom_out(self):
        self.zoom_factor = max(self.zoom_factor - 0.15, 0.4)
        self.apply_zoom()

    def apply_zoom(self):
        if HAS_WEBENGINE and hasattr(self, "viewer"):
            self.viewer.setZoomFactor(self.zoom_factor)
            return
        self.render_fallback_pages()

    def _add_fallback_viewer(self, layout, pdf_path):
        self.fallback_scroll = QScrollArea()
        self.fallback_scroll.setWidgetResizable(True)
        self.fallback_container = QWidget()
        self.fallback_layout = QVBoxLayout(self.fallback_container)
        self.fallback_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.fallback_scroll.setWidget(self.fallback_container)
        layout.addWidget(self.fallback_scroll)
        
        try:
            self.fallback_doc = fitz.open(pdf_path)
            for _page in self.fallback_doc:
                lbl = QLabel()
                self.fallback_page_labels.append(lbl)
                self.fallback_layout.addWidget(lbl)
            self.render_fallback_pages()
        except Exception as e:
            err_lbl = QLabel(f"Failed to load PDF: {e}")
            err_lbl.setStyleSheet("color: #b91c1c; font-weight: bold;")
            self.fallback_layout.addWidget(err_lbl)

    def render_fallback_pages(self):
        if not self.fallback_doc:
            return
        render_zoom = 2.0 * self.zoom_factor
        for page, lbl in zip(self.fallback_doc, self.fallback_page_labels):
            pix = page.get_pixmap(matrix=fitz.Matrix(render_zoom, render_zoom))
            img = QImage(
                pix.samples, pix.width, pix.height,
                pix.stride, QImage.Format.Format_RGB888
            ).copy()
            lbl.setPixmap(QPixmap.fromImage(img))


class FileDetailsDialog(QDialog):
    details_saved = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("File Details Extraction")
        self.resize(450, 350)
        self.setStyleSheet("""
            QDialog { background-color: #f8fafc; color: #191c1e; }
            QLabel { font-size: 13px; font-weight: bold; color: #334155; }
            QLineEdit { 
                padding: 8px; border: 1px solid #cbd5e1; 
                border-radius: 4px; background-color: #ffffff; color: #191c1e;
            }
            QLineEdit:focus { border: 1px solid #0a4da2; }
            QPushButton { 
                font-weight: bold; border-radius: 4px; padding: 10px; border: none;
            }
            QPushButton#ExtractBtn { background-color: #00264d; color: #ffffff; }
            QPushButton#ExtractBtn:hover { background-color: #0a4da2; }
            QPushButton#CloseBtn { background-color: #64748b; color: #ffffff; }
            QPushButton#CloseBtn:hover { background-color: #475569; }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        self.form_layout = QGridLayout()
        self.form_layout.setVerticalSpacing(15)
        
        self.fields = {
            "File no": QLineEdit(),
            "Issue no": QLineEdit(),
            "Date of issue": QLineEdit(),
            "Project name": QLineEdit()
        }
        
        for row, (label, edit) in enumerate(self.fields.items()):
            self.form_layout.addWidget(QLabel(f"{label}:"), row, 0)
            self.form_layout.addWidget(edit, row, 1)
            
        layout.addLayout(self.form_layout)
        
        self.status_label = QLabel(
            "Review the detected details, edit anything that looks wrong, "
            "then save them to the generated file."
        )
        self.status_label.setStyleSheet("color: #64748b; font-weight: normal; font-style: italic;")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        
        btn_layout = QHBoxLayout()
        self.extract_btn = QPushButton("Save to File")
        self.extract_btn.setObjectName("ExtractBtn")
        self.close_btn = QPushButton("Close")
        self.close_btn.setObjectName("CloseBtn")
        
        btn_layout.addWidget(self.extract_btn)
        btn_layout.addWidget(self.close_btn)
        layout.addLayout(btn_layout)
        
        self.extract_btn.clicked.connect(self.save_details)
        self.close_btn.clicked.connect(self.accept)

    def save_details(self):
        details = {
            "file_no": self.fields["File no"].text().strip(),
            "issue_no": self.fields["Issue no"].text().strip(),
            "date_of_issue": self.fields["Date of issue"].text().strip(),
            "project_name": self.fields["Project name"].text().strip(),
        }
        self.details_saved.emit(details)
        self.status_label.setText("Saved to generated file details.")
        self.status_label.setStyleSheet("color: #0a4da2; font-weight: bold;")

    def set_extracting_state(self, is_extracting):
        if is_extracting:
            self.extract_btn.setEnabled(False)
            self.progress_bar.show()
            self.status_label.setText("Reading OCR text locally...")
        else:
            self.extract_btn.setEnabled(True)
            self.progress_bar.hide()

    def update_results(self, results):
        self.set_extracting_state(False)
        if "error" in results:
            self.status_label.setText(f"Error: {results['error']}")
            self.status_label.setStyleSheet("color: #b91c1c;")
            return
            
        if results.get("source") == "selected_ocr":
            self.status_label.setText("Details extracted from selected OCR text.")
        else:
            self.status_label.setText("OCR details extracted. Please review before saving.")
        self.status_label.setStyleSheet("color: #0a4da2; font-weight: bold;")
        
        self.fields["File no"].setText(results.get("file_no", "Not found"))
        self.fields["Issue no"].setText(results.get("issue_no", "Not found"))
        self.fields["Date of issue"].setText(results.get("date_of_issue", "Not found"))
        self.fields["Project name"].setText(results.get("project_name", "Not found"))
