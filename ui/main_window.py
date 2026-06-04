import os
import tempfile
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QLabel, QSplitter,
    QMessageBox, QScrollArea, QProgressBar, QCheckBox,
    QGroupBox, QDialog, QGridLayout, QTabWidget,
    QLineEdit, QTextEdit, QApplication, QComboBox, QSizePolicy
)
from PyQt6.QtWidgets import QStackedWidget
from PyQt6.QtCore import Qt, QUrl, QEvent, QSize, QTimer
from PyQt6.QtGui import QImage, QPixmap, QTextCursor

from ui.styles import get_light_theme
from ui.widgets import CollapsibleBox, PopupMultiSelect
from ui.dialogs import PDFViewerDialog, FileDetailsDialog
from components.pdf_viewer import PdfPageLabel, PDFWebEngineView, HAS_WEBENGINE
from components.text_editor import NotesEditor, SelectableTextPreview
from components.tables import ProjectTablesDialog
from core.pdf_handler import PDFHandler
from core.analysis import (
    AnalysisWorker,
    simple_summarize,
    FileDetailsWorker,
    SignatureValidationWorker,
)
from core.export import PDFExporter, DocxExporter
from core.ocr import extract_file_details_from_text, extract_region_text


class OfflineApp(QMainWindow):
    @property
    def manual_input(self):
        if hasattr(self, 'notes_heading_combo') and hasattr(self, 'notes_editors'):
            return self.notes_editors[self.notes_heading_combo.currentText()]
        return None
    
    def store_selected_region(self, pixmap, rect, page):
        self.current_selected_region = (
            pixmap,
            rect,
            page
        )
        QMessageBox.information(
            self,
            "OCR Selection",
            "Region selected successfully.\n\nNow click 'OCR'."
        )

    def perform_selected_ocr(self):
        try:
            if not hasattr(self, "current_selected_region"):
                QMessageBox.warning(
                    self,
                    "OCR",
                    "Please select a region first."
                )
                return

            pixmap, rect, page = self.current_selected_region
            self.perform_region_ocr(pixmap, rect, page)

        except Exception as e:
            QMessageBox.critical(self, "OCR Error", str(e))

    def perform_region_ocr(self, pixmap, rect, page):
        try:
            
            ocr_text = extract_region_text(
                page,
                pixmap.width(),
                pixmap.height(),
                rect
            )
            self.last_region_ocr_text = ocr_text

            if not ocr_text:
                QMessageBox.warning(self, "OCR", "No text detected.")
                return

            dialog = QDialog(self)
            dialog.setWindowTitle("OCR Extracted Text")
            dialog.resize(900, 700)
            layout = QVBoxLayout(dialog)
            text_edit = QTextEdit()
            text_edit.setPlainText(ocr_text)
            text_edit.setReadOnly(False)
            layout.addWidget(text_edit)
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(dialog.close)
            layout.addWidget(close_btn)
            dialog.exec()
            self.last_region_ocr_text = text_edit.toPlainText()

        except ImportError:
            QMessageBox.critical(
                self,
                "OCR Error",
                "Install pytesseract, opencv-python, PyMuPDF, and numpy"
            )
        except Exception as e:
            QMessageBox.critical(self, "OCR Error", str(e))

    def get_all_manual_notes_html(self):
        notes = {}
        if hasattr(self, 'notes_editors'):
            for heading, editor in self.notes_editors.items():
                text = editor.toPlainText().strip()
                if text:
                    notes[heading] = editor.toHtml()
        return notes

    def get_all_manual_notes_text(self):
        notes = {}
        if hasattr(self, 'notes_editors'):
            for heading, editor in self.notes_editors.items():
                text = editor.toPlainText().strip()
                if text:
                    notes[heading] = text
        return notes

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Document Extraction Assistant")
        self.resize(1280, 800)
        self.setMinimumSize(QSize(900, 600))
        self.first_show = True
        self.splitter_initialized = False
        
        self.file_path = ""
        self.pdf_handler = PDFHandler()
        self.current_zoom = 600
        self.last_region_ocr_text = ""
        self.annexure3_image_path = ""
        self.last_docx_path = ""
        
        # Data Variables
        self.extracted_project_name = ""
        self.intro_text_data = ""
        self.issue_details_data = {
            "file_no": "",
            "issue_no": "",
            "date_of_issue": "",
        }
        self.stakeholders_data = []
        self.cert_task_data = []
        self.annexure1_data = []
        self.test_rigs_data = []
        self.aircraft_checks_data = []
        self.annexure2_data = []
        
        # Initialize UI components
        self.text_preview = SelectableTextPreview()
        self.text_preview.text_extracted.connect(self.add_extracted_text)
        self.text_preview.summary_requested.connect(
            self.generate_summary_from_selection
        )
        
        self.init_data_dialog()
        self.init_tables_dialog()
        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        main_widget.setObjectName("mainWidget")
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        # Top toolbar with upload button
        self._create_top_toolbar()
        layout.addLayout(self.top_toolbar_layout)
        
        # Main splitter (50-50 by default but adjustable)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setHandleWidth(5)
        self.splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #cbd5e1;
            }
            QSplitter::handle:hover {
                background-color: #94a3b8;
            }
        """)

        # Left and Right sides
        self._create_left_panel()
        self._create_right_panel()

        # Set equal stretch factors for 50-50 default
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 1)
        
        layout.addWidget(self.splitter, 1)

    def _create_top_toolbar(self):
        """Create top toolbar with upload button only"""
        self.top_toolbar_layout = QHBoxLayout()
        self.top_toolbar_layout.setContentsMargins(0, 0, 0, 0)
        self.top_toolbar_layout.setSpacing(5)
        
        self.upload_btn = QPushButton("Upload PDF")
        self.upload_btn.setObjectName("PrimaryBtn")
        self.upload_btn.setMinimumHeight(40)
        self.upload_btn.setMaximumWidth(150)
        self.upload_btn.clicked.connect(self.upload_file)
        
        self.top_toolbar_layout.addWidget(self.upload_btn)
        self.top_toolbar_layout.addStretch()

    def _create_left_panel(self):
        left_container = QWidget()
        left_container.setMinimumWidth(360)
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(5)

        # Tabs for PDF/Text viewing
        self.left_tabs = QTabWidget()
        self.left_tabs.setStyleSheet("""
            QTabWidget { background: #ffffff; }
            QTabBar { background: #ffffff; border-bottom: 1px solid #cbd5e1; }
            QTabBar::tab { padding: 8px 15px; color: #334155; }
            QTabBar::tab:selected { background: #00264d; color: #ffffff; border-bottom: 2px solid #00264d; }
        """)
        self._setup_left_tabs()
        left_layout.addWidget(self.left_tabs, 1)

        # Buttons in the tab corner (right-aligned)
        corner_widget = QWidget()
        corner_layout = QHBoxLayout(corner_widget)
        corner_layout.setContentsMargins(0, 0, 5, 0)
        corner_layout.setSpacing(4)

        self.preview_zoom_out_btn = QPushButton("-")
        self.preview_zoom_out_btn.setObjectName("SecondaryBtn")
        self.preview_zoom_out_btn.setToolTip("Zoom out")
        self.preview_zoom_out_btn.setFixedSize(32, 32)
        self.preview_zoom_out_btn.clicked.connect(lambda: self.zoom_out(75))

        self.preview_zoom_in_btn = QPushButton("+")
        self.preview_zoom_in_btn.setObjectName("SecondaryBtn")
        self.preview_zoom_in_btn.setToolTip("Zoom in")
        self.preview_zoom_in_btn.setFixedSize(32, 32)
        self.preview_zoom_in_btn.clicked.connect(lambda: self.zoom_in(75))
        for zoom_btn in [self.preview_zoom_out_btn, self.preview_zoom_in_btn]:
            zoom_btn.setStyleSheet("""
                QPushButton {
                    background-color: #e2e8f0;
                    color: #334155;
                    border: 1px solid #cbd5e1;
                    border-radius: 6px;
                    font-size: 18px;
                    font-weight: bold;
                    padding: 0px;
                }
                QPushButton:hover { background-color: #cbd5e1; color: #00264d; }
                QPushButton:pressed { background-color: #94a3b8; color: #ffffff; }
            """)
        
        self.file_details_btn = QPushButton("File Details")
        self.file_details_btn.setObjectName("SecondaryBtn")
        self.file_details_btn.setMinimumHeight(32)
        self.file_details_btn.setMinimumWidth(96)
        self.file_details_btn.clicked.connect(self.show_file_details)
        
        self.ocr_btn = QPushButton("OCR")
        self.ocr_btn.setObjectName("SecondaryBtn")
        self.ocr_btn.setMinimumHeight(32)
        self.ocr_btn.setMinimumWidth(64)
        self.ocr_btn.clicked.connect(self.perform_selected_ocr)
        
        corner_layout.addWidget(self.preview_zoom_out_btn)
        corner_layout.addWidget(self.preview_zoom_in_btn)
        corner_layout.addWidget(self.file_details_btn)
        corner_layout.addWidget(self.ocr_btn)
        
        # Set the corner widget to the right corner of the tab bar
        self.left_tabs.setCornerWidget(corner_widget, Qt.Corner.TopRightCorner)

        self.splitter.addWidget(left_container)

    def _setup_left_tabs(self):
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("""
            QScrollArea { 
                border: none; 
                background: #f8fafc; 
            }
            QScrollBar:vertical {
                border: none;
                background: #f1f5f9;
                width: 12px;
            }
            QScrollBar::handle:vertical {
                background: #cbd5e1;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94a3b8;
            }
        """)
        
        self.page_container = QWidget()
        self.page_layout = QVBoxLayout(self.page_container)
        self.page_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.page_layout.setSpacing(8)
        self.scroll_area.setWidget(self.page_container)
        self.scroll_area.viewport().installEventFilter(self)

        if HAS_WEBENGINE:
            self.pdf_viewer = PDFWebEngineView(self)
            self.pdf_viewer.text_extracted.connect(self.add_extracted_text)
            self.pdf_viewer.summary_requested.connect(
                self.generate_summary_from_selection
            )
            self.left_tabs.addTab(self.scroll_area, "Visual Extractor")
            self.left_tabs.addTab(self.text_preview, "Text Fallback")
        else:
            self.left_tabs.addTab(self.scroll_area, "Visual Extractor")
            self.left_tabs.addTab(self.text_preview, "Text Viewer")
    
    def _create_right_panel(self):
        right_container = QWidget()
        right_scroll = QScrollArea()
        self.right_scroll = right_scroll
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        right_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        right_scroll.verticalScrollBar().setSingleStep(30)
        right_scroll.setStyleSheet("""
            QScrollArea { 
                border: none; 
                background: #ffffff;
            }
            QScrollBar:vertical {
                border: none;
                background: #f1f5f9;
                width: 12px;
            }
            QScrollBar::handle:vertical {
                background: #cbd5e1;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94a3b8;
            }
        """)

        # Main scrollable content widget
        scroll_content = QWidget()
        scroll_content.setMinimumWidth(420)
        scroll_content.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        right_layout = QVBoxLayout(scroll_content)
        right_layout.setContentsMargins(5, 5, 5, 5)
        right_layout.setSpacing(8)

        # Right panel title
        title_label = QLabel("Document Assistant")
        title_label.setObjectName("PanelTitle")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #00264d;")
        right_layout.addWidget(title_label)

        # Toolbar with table buttons
        icon_layout = QHBoxLayout()
        icon_layout.addStretch()
        
        self.tables_icon_btn = QPushButton("Tables")
        self.tables_icon_btn.setObjectName("SecondaryBtn")
        self.tables_icon_btn.setMinimumHeight(35)
        self.tables_icon_btn.setMinimumWidth(80)
        self.tables_icon_btn.clicked.connect(self.show_tables_popup)
        icon_layout.addWidget(self.tables_icon_btn)
        
        self.data_icon_btn = QPushButton("Data")
        self.data_icon_btn.setObjectName("SecondaryBtn")
        self.data_icon_btn.setMinimumHeight(35)
        self.data_icon_btn.setMinimumWidth(70)
        self.data_icon_btn.clicked.connect(lambda: self.show_data_popup(0))
        icon_layout.addWidget(self.data_icon_btn)
        
        icon_layout.addStretch()
        right_layout.addLayout(icon_layout)

        # Status group
        self._create_status_group()
        right_layout.addWidget(self.status_group)

        # Collapsible params
        self._create_params_section()
        right_layout.addWidget(self.collapsible_params)

        # Add stretch to push action buttons to bottom
        right_layout.addStretch()

        # Action buttons (fixed at bottom)
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(10)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        
        self.save_pdf_btn = QPushButton("View PDF")
        self.save_pdf_btn.setObjectName("SecondaryBtn")
        self.save_pdf_btn.setMinimumHeight(45)
        self.save_pdf_btn.setMinimumWidth(110)
        self.save_pdf_btn.clicked.connect(self.export_pdf)
        actions_layout.addWidget(self.save_pdf_btn)

        self.save_docx_btn = QPushButton("Generate Word")
        self.save_docx_btn.setObjectName("PrimaryBtn")
        self.save_docx_btn.setMinimumHeight(45)
        self.save_docx_btn.setMinimumWidth(150)
        self.save_docx_btn.clicked.connect(self.export_docx)
        actions_layout.addWidget(self.save_docx_btn, 1)

        right_layout.addLayout(actions_layout)

        # Set the scroll area widget
        right_scroll.setWidget(scroll_content)
        self.splitter.addWidget(right_scroll)

    def _create_status_group(self):
        self.status_group = QGroupBox("Checklist")
        self.status_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 10px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #00264d;
            }
        """)
        
        status_main_layout = QVBoxLayout(self.status_group)
        status_main_layout.setContentsMargins(10, 10, 10, 10)
        status_main_layout.setSpacing(8)
        
        status_header = QHBoxLayout()
        self.status_label = QLabel("Waiting for document...")
        self.status_label.setStyleSheet("color: #64748b; font-style: italic; font-size: 12px;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setMaximumHeight(20)
        self.progress_bar.hide()
        status_header.addWidget(self.status_label)
        status_header.addWidget(self.progress_bar)

        self.save_checklist_img_btn = QPushButton("⬇")
        self.save_checklist_img_btn.setToolTip("Save checklist as image")
        self.save_checklist_img_btn.setFixedSize(28, 28)
        self.save_checklist_img_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_checklist_img_btn.clicked.connect(self.save_checklist_as_image)
        self.save_checklist_img_btn.setStyleSheet("""
            QPushButton {
                background-color: #e2e8f0;
                color: #00264d;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                font-size: 15px;
                font-weight: bold;
                padding: 0px;
            }
            QPushButton:hover {
                background-color: #cbd5e1;
            }
            QPushButton:pressed {
                background-color: #94a3b8;
                color: white;
            }
        """)
        status_header.addWidget(self.save_checklist_img_btn)

        status_main_layout.addLayout(status_header)
        
        self.topics = [
            "Introduction",
            "Scope",
            "Product Breakdown",
            "System LRU and Module Details",
            "Signatures",
            "Test Rig Simulators and Ground equipment",
            "Work breakdown structure"
        ]
        self.topic_checkboxes = {}
        
        topics_widget = QWidget()
        topics_layout = QGridLayout(topics_widget)
        topics_layout.setSpacing(6)
        topics_layout.setContentsMargins(0, 0, 0, 0)
        
        # 2 columns for better layout
        row, col = 0, 0
        for topic in self.topics:
            cb = QCheckBox(topic)
            cb.setEnabled(False)
            cb.setStyleSheet("QCheckBox { font-size: 12px; }")
            self.topic_checkboxes[topic] = cb
            topics_layout.addWidget(cb, row, col)
            col += 1
            if col > 1:  # 2 columns
                col = 0
                row += 1

        topics_layout.setColumnStretch(0, 1)
        topics_layout.setColumnStretch(1, 1)
        status_main_layout.addWidget(topics_widget)


    def save_checklist_as_image(self):
        try:
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Checklist Image",
                "checklist.png",
                "PNG Image (*.png)"
            )

            if not save_path:
                return

            if not save_path.lower().endswith(".png"):
                save_path += ".png"

            # Hide the Save Image button so it does not appear in the exported image
            self.save_checklist_img_btn.hide()
            QApplication.processEvents()

            pixmap = self.status_group.grab()
            success = pixmap.save(save_path, "PNG")

            self.save_checklist_img_btn.show()

            if success:
                QMessageBox.information(
                    self,
                    "Image Saved",
                    "Checklist image saved successfully."
                )
            else:
                QMessageBox.warning(
                    self,
                    "Save Failed",
                    "Could not save the checklist image."
                )

        except Exception as e:
            if hasattr(self, "save_checklist_img_btn"):
                self.save_checklist_img_btn.show()

            QMessageBox.critical(
                self,
                "Image Export Error",
                str(e)
            )
    def _create_params_section(self):
        self.collapsible_params = CollapsibleBox("ADDITIONAL PARAMETERS")
        self.collapsible_params.setStyleSheet("""
            QGroupBox {
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 10px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #00264d;
            }
        """)
        
        dropdowns_data = [
            ("Scope Of Task Directive", [
                "Development of air system in accordance with IMTAR-21 Subpart- B",
                "Development of Airborne LRUs in accordance with IMTAR-21 Subpart- C1.",
                "Development of Airborne Software and CEH in accordance with IMTAR-21 Subpart- C6.",
                "Development of Ground support system in accordance with IMTAR-21 Subpart- T.",
                "Continued Airworthiness Coverage in accordance with IMTAR-21 Subpart-L",
                "Bought Out Item clearance in accordance with IMTAR-21 Subpart N and ACR 002/2025",
            ]),
            ("Certification Work Breakdown", [
                "Carry out SSA and classify the criticality of each System and LRU",
                "Requirements Analysis of each System and LRU",
                "Finalization and approval of LRU Specification and build standard",
                "Test Requirement Traceability Matrix (Means of Compliance)",
                "Finalization of Type Approval Basis (TAB) and Airworthiness Certification Plan",
                "Software Certification Plan (PSAC)",
                "Hardware and CEH certification Plan",
                "Preliminary Design Review & Critical Design Review",
                "Test Adequacy Review",
                "LRU Test Rig Specification Approval & Rig Acceptance by DGAQA",
                "Functional Test Plan, Integration Test Plan",
                "Realization of Hardware and inspection by DGAQA",
                "Safety of Flight Testing/Qualification Testing",
                "IV & V of Software and CEH, and Software certification activities",
                "System Integration Testing",
                "Clearance for flight trials",
                "Satisfactory Flight trial feedback from Flight Ops/ Users",
                "Compliance to TAB, ACP, QTP, PSAC, TRTM",
                "Verification of Flight Trial feedback",
                "Production clearance & Service use clearance for the system",
                "RMTC/ MTC of the air system",
                "Continued Airworthiness Activities",
            ]),
            ("Internal Distribution", [
                "Director (System's)",
                "Director (A/C)",
                "Director (Propulsion)",
                "Director (Mat & PI)",
                "Group Director (MS)",
                "RD, RCMA(APS)",
                "RD, RCMA (A/C)",
                "RD, RCMA()",
                "RD, RCMA()",
                "RD, RCMA()",
                "RD, RCMA()",
                "E-Certification",
                "C-Cat lab",
                "Head, SRM Center",
            ]),
        ]
        
        self.combos = {}
        param_widget = QWidget()
        param_layout = QVBoxLayout(param_widget)
        param_layout.setSpacing(10)
        param_layout.setContentsMargins(0, 0, 0, 0)
        
        for label_text, items in dropdowns_data:
            lbl = QLabel(label_text.upper())
            lbl.setStyleSheet("font-weight: bold; color: #334155; margin-top: 2px; font-size: 11px;")
            combo = PopupMultiSelect(items)
            combo.setMinimumHeight(32)
            self.combos[label_text] = combo
            param_layout.addWidget(lbl)
            param_layout.addWidget(combo)
            
        param_layout.addStretch()
        self.collapsible_params.content_layout.addWidget(param_widget)

    def init_data_dialog(self):
        self.data_dialog = QDialog(self)
        self.data_dialog.setWindowTitle("Project Data & Notes")
        self.data_dialog.resize(800, 700)
        self.data_dialog.setMinimumSize(QSize(600, 500))
        
        dialog_layout = QVBoxLayout(self.data_dialog)
        self.data_tabs = QTabWidget()

        # Summary tab
        summary_tab = QWidget()
        summary_layout = QVBoxLayout(summary_tab)
        summary_layout.setSpacing(10)
        
        summary_layout.addWidget(QLabel("Project Name:"))
        self.popup_name_edit = QLineEdit()
        self.popup_name_edit.setMinimumHeight(32)
        self.popup_name_edit.textChanged.connect(
            lambda text: setattr(self, 'extracted_project_name', text)
        )
        summary_layout.addWidget(self.popup_name_edit)
        
        summary_layout.addWidget(QLabel("Introduction Summary:"))
        self.popup_summary_edit = QTextEdit()
        self.popup_summary_edit.setMinimumHeight(150)
        self.popup_summary_edit.textChanged.connect(
            lambda: setattr(self, 'intro_text_data', 
                          self.popup_summary_edit.toPlainText())
        )
        summary_layout.addWidget(self.popup_summary_edit)
        summary_layout.addStretch()
        
        # Notes tab
        notes_tab = QWidget()
        notes_layout = QVBoxLayout(notes_tab)
        notes_layout.setSpacing(8)
        
        heading_layout = QHBoxLayout()
        heading_layout.addWidget(QLabel("Select Heading:"))
        self.notes_heading_combo = QComboBox()
        self.notes_heading_combo.setMinimumHeight(32)
        self.headings = [
            "Introduction",
            "Reference",
            "Basis Of Task Directive",
            "Scope Of Task Directive",
            "Stakeholders",
            "Certification Work Breakdown",
            "Task Allocation",
            "Coordinating Directorate",
            "Single Point of Contact( SPoC)",
            "Certification Task Allocation",
            "Issue of Clearance",
            "SCRB And TARB",
            "Communication",
            "Certification Progress Review",
            "External Organization",
            "Internal Distribution",
            "Annexure-1",
            "Work Assignment List of LRUs",
            "Aircraft Integration Checks and Flight Clearance",
            "Annexure-2",
            "Contact details of dealing officers and RDs",
            "Annexure-3",
            "Product Break Down Structure",
            "Other"
        ]
        self.notes_heading_combo.addItems(self.headings)
        heading_layout.addWidget(self.notes_heading_combo)
        notes_layout.addLayout(heading_layout)
        
        self.notes_stack = QStackedWidget()
        self.notes_editors = {}
        for heading in self.headings:
            editor = NotesEditor()
            self.notes_editors[heading] = editor
            self.notes_stack.addWidget(editor)
            
        self.notes_heading_combo.currentTextChanged.connect(
            lambda text: self.notes_stack.setCurrentWidget(self.notes_editors[text])
        )
        notes_layout.addWidget(self.notes_stack)
        
        self.data_tabs.addTab(summary_tab, "Summary")
        self.data_tabs.addTab(notes_tab, "Manual Notes")
        dialog_layout.addWidget(self.data_tabs)
        
        close_btn = QPushButton("Save & Close")
        close_btn.setMinimumHeight(40)
        close_btn.clicked.connect(self.data_dialog.hide)
        dialog_layout.addWidget(close_btn)

    def init_tables_dialog(self):
        self.tables_dialog = ProjectTablesDialog(self)

    def get_docx_table_data(self):
        return {
            "issue_details": self.issue_details_data,
            "annexure1": self.annexure1_data,
            "test_rigs": self.test_rigs_data,
            "aircraft_checks": self.aircraft_checks_data,
            "annexure2": self.annexure2_data,
        }

    def show_data_popup(self, tab_index=0):
        self.popup_name_edit.setText(self.extracted_project_name)
        self.popup_summary_edit.setText(self.intro_text_data)
        self.data_tabs.setCurrentIndex(tab_index)
        self.data_dialog.show()

    def show_tables_popup(self):
        self.tables_dialog.refresh_all_tables()
        self.tables_dialog.show()

    def show_file_details(self):
        if not self.pdf_handler or (
            not self.pdf_handler.page_data and not self.last_region_ocr_text.strip()
        ):
            QMessageBox.warning(self, "No Document", "Please upload a document first.")
            return
            
        self.file_details_dialog = FileDetailsDialog(self)
        self.file_details_dialog.details_saved.connect(
            self.save_file_details_to_document
        )
        self.file_details_dialog.show()

        if self.last_region_ocr_text.strip():
            details = extract_file_details_from_text(self.last_region_ocr_text)
            details["source"] = "selected_ocr"
            self.file_details_dialog.update_results(details)
            return

        self.file_details_dialog.set_extracting_state(True)
        
        self.file_details_worker = FileDetailsWorker(self.pdf_handler)
        self.file_details_worker.finished.connect(self.file_details_dialog.update_results)
        self.file_details_worker.start()

    def save_file_details_to_document(self, details):
        def cleaned(value):
            value = (value or "").strip()
            if value.lower() in {"not found", "notfound"}:
                return ""
            return value

        self.issue_details_data["file_no"] = cleaned(details.get("file_no"))
        self.issue_details_data["issue_no"] = cleaned(details.get("issue_no"))
        self.issue_details_data["date_of_issue"] = cleaned(
            details.get("date_of_issue")
        )

        project_name = cleaned(details.get("project_name"))
        if project_name:
            self.extracted_project_name = project_name
            if hasattr(self, "popup_name_edit"):
                self.popup_name_edit.setText(project_name)

        if hasattr(self, "tables_dialog"):
            self.tables_dialog.refresh_all_tables()

        self.flash_data_icon()

    def flash_data_icon(self):
        self.data_icon_btn.setObjectName("SuccessBtn")
        self.data_icon_btn.style().unpolish(self.data_icon_btn)
        self.data_icon_btn.style().polish(self.data_icon_btn)

    def paste_to_notes(self):
        clipboard_text = QApplication.clipboard().text()
        if clipboard_text.strip():
            self.manual_input.append(clipboard_text + "\n")
            self.flash_data_icon()

    def summarize_clipboard(self):
        clipboard_text = QApplication.clipboard().text()
        if clipboard_text.strip():
            self.generate_summary_from_selection(clipboard_text)

    def generate_summary_from_selection(self, text):
        self.status_label.setText("Summarizing...")
        summary = simple_summarize(text)
        if summary:
            self.intro_text_data = summary
            self.popup_summary_edit.setText(summary)
            self.flash_data_icon()
            self.topic_checkboxes["Introduction"].setChecked(True)
        self.status_label.setText("Ready")

    def add_extracted_text(self, text):
        QApplication.clipboard().setText(text)
        self.manual_input.append(text + "\n")
        self.flash_data_icon()

    def add_extracted_image(self, pixmap):
        if pixmap.width() > 400:
            pixmap = pixmap.scaledToWidth(
                400, Qt.TransformationMode.SmoothTransformation
            )
        cursor = self.manual_input.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.manual_input.setTextCursor(cursor)
        cursor.insertImage(pixmap.toImage())
        self.manual_input.append("\n")
        self.flash_data_icon()

    def set_annexure3_image(self, pixmap):
        if not pixmap or pixmap.isNull():
            QMessageBox.warning(self, "PBS/Annexure-3", "No image was selected.")
            return

        if self.annexure3_image_path and os.path.exists(self.annexure3_image_path):
            try:
                os.remove(self.annexure3_image_path)
            except OSError:
                pass

        temp_file = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        self.annexure3_image_path = temp_file.name
        temp_file.close()
        if not pixmap.save(self.annexure3_image_path, "PNG"):
            QMessageBox.critical(
                self,
                "PBS/Annexure-3",
                "Could not save the selected image."
            )
            self.annexure3_image_path = ""
            return

        QMessageBox.information(
            self,
            "PBS/Annexure-3",
            "Selected image will be added to Annexure-3 in the generated Word document."
        )

    def showEvent(self, event):
        """Ensure 50-50 split after window is properly laid out"""
        super().showEvent(event)
        if self.first_show and not self.splitter_initialized:
            self.first_show = False
            # Delay the resizing until after the window is fully rendered
            QTimer.singleShot(100, self.set_initial_splitter_sizes)

    def set_initial_splitter_sizes(self):
        """Set the initial 50-50 split after window is laid out"""
        if self.splitter_initialized:
            return
            
        total_width = self.splitter.width()
        if total_width > 0:
            half_width = total_width // 2
            self.splitter.setSizes([half_width, half_width])
            self.splitter_initialized = True
        else:
            # Try again if width is still 0
            QTimer.singleShot(50, self.set_initial_splitter_sizes)

    def eventFilter(self, source, event):
        if source == self.scroll_area.viewport():
            if event.type() == QEvent.Type.Wheel:
                if QApplication.keyboardModifiers() == Qt.KeyboardModifier.ControlModifier:
                    delta = event.angleDelta().y()
                    if delta > 0:
                        self.zoom_in()
                    elif delta < 0:
                        self.zoom_out()
                    return True
        return super().eventFilter(source, event)

    def zoom_in(self, step=50):
        self.current_zoom += step
        self.refresh_pdf_view()

    def zoom_out(self, step=50):
        if self.current_zoom - step > 200:
            self.current_zoom -= step
            self.refresh_pdf_view()

    def refresh_pdf_view(self):
        page_data = self.pdf_handler.page_data

        if not page_data:
            return

        # Create labels if needed
        if self.page_layout.count() != len(page_data):
            while self.page_layout.count():
                child = self.page_layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()

            for pixmap, page_idx in page_data:
                page = self.pdf_handler.get_page(page_idx)
                lbl = PdfPageLabel(page, self.scroll_area)
                lbl.mousePressEvent_original = lbl.mousePressEvent

                def wrapped_mouse_press(event, label=lbl):
                    self.current_active_label = label
                    label.mousePressEvent_original(event)

                lbl.mousePressEvent = wrapped_mouse_press
                lbl.text_extracted.connect(self.add_extracted_text)
                lbl.image_extracted.connect(self.add_extracted_image)
                lbl.annexure3_image_extracted.connect(self.set_annexure3_image)
                lbl.summary_requested.connect(self.generate_summary_from_selection)
                lbl.region_ocr_requested.connect(self.store_selected_region)
                self.page_layout.addWidget(lbl)

        # Update zoom
        for i in range(self.page_layout.count()):
            lbl = self.page_layout.itemAt(i).widget()
            if isinstance(lbl, PdfPageLabel) and i < len(page_data):
                scaled_pixmap = page_data[i][0].scaledToWidth(
                    self.current_zoom,
                    Qt.TransformationMode.SmoothTransformation
                )
                lbl.setPixmap(scaled_pixmap)
                lbl.setFixedSize(scaled_pixmap.size())

    def upload_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Document", "", "Documents (*.pdf *.txt)"
        )
        if not path:
            return
            
        self.file_path = path
        
        # Reset state
        for cb in self.topic_checkboxes.values():
            cb.setChecked(False)
        self.extracted_project_name = ""
        self.intro_text_data = ""
        self.issue_details_data = {
            "file_no": "",
            "issue_no": "",
            "date_of_issue": "",
        }
        self.last_region_ocr_text = ""
        self.annexure3_image_path = ""
        self.last_docx_path = ""
        if hasattr(self, 'notes_editors'):
            for editor in self.notes_editors.values():
                editor.clear()
        self.text_preview.clear()
        
        try:
            full_text = ""
            if path.lower().endswith('.pdf'):
                if HAS_WEBENGINE:
                    self.pdf_viewer.setUrl(
                        QUrl.fromLocalFile(os.path.abspath(path))
                    )
                    self.left_tabs.setCurrentIndex(0)
                else:
                    self.left_tabs.setCurrentIndex(0)
                    
                self.status_label.setText("Extracting...")
                full_text = self.pdf_handler.open_pdf(path)
                self.text_preview.setText(full_text)
                self.pdf_handler.extract_pages(15)
                self.current_zoom = 600
                self.refresh_pdf_view()
            else:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    full_text = f.read()
                    self.text_preview.setText(full_text)

            self.start_analysis_thread(full_text)
            if path.lower().endswith('.pdf'):
                self.start_signature_validation_thread()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not load: {str(e)}")

    def start_analysis_thread(self, text):
        self.status_label.setText("Analyzing...")
        self.progress_bar.show()
        self.worker = AnalysisWorker(text, self.topics)
        self.worker.finished.connect(self.handle_analysis_done)
        self.worker.status.connect(
            lambda msg: self.status_label.setText(msg)
        )
        self.worker.start()

    def start_signature_validation_thread(self):
        if not self.pdf_handler or not self.pdf_handler.doc:
            return

        self.signature_worker = SignatureValidationWorker(self.pdf_handler)
        self.signature_worker.finished.connect(self.handle_signature_validation_done)
        self.signature_worker.error.connect(self.handle_signature_validation_error)
        self.signature_worker.start()

    def handle_signature_validation_done(self, has_signature):
        if "Signatures" in self.topic_checkboxes:
            self.topic_checkboxes["Signatures"].setChecked(bool(has_signature))

    def handle_signature_validation_error(self, message):
        if "Signatures" in self.topic_checkboxes:
            self.topic_checkboxes["Signatures"].setChecked(False)
        self.status_label.setText(f"Signature check unavailable: {message}")

    def handle_analysis_done(self, results):
        self.progress_bar.hide()
        self.status_label.setText("Analysis Complete")
        
        if results["project_name"]:
            self.extracted_project_name = results["project_name"]
            
        if results.get("introduction"):
            self.intro_text_data = results["introduction"]
            self.flash_data_icon()
            
        for topic in results["found_topics"]:
            if topic in self.topic_checkboxes:
                self.topic_checkboxes[topic].setChecked(True)

    def export_pdf(self):
        try:
            scope_combo = self.combos.get("Scope Of Task Directive")
            scope_items = scope_combo.checkedItems() if scope_combo else []

            temp_path = DocxExporter.create_preview_pdf(
                self.extracted_project_name,
                self.intro_text_data,
                self.get_all_manual_notes_text(),
                scope_items,
                self.stakeholders_data,
                self.cert_task_data,
                self.combos,
                self.annexure3_image_path,
                self.get_docx_table_data()
            )

            self.preview_dialog = PDFViewerDialog(temp_path, self)
            self.preview_dialog.exec()

        except Exception as e:
            QMessageBox.warning(
                self,
                "PDF Preview",
                "Could not create the preview. Make sure Microsoft Word "
                "or LibreOffice is installed for DOCX-to-PDF conversion."
                f"\n\nTechnical reason: {str(e)}"
            )

    def _write_docx(self, save_path):
        scope_combo = self.combos.get("Scope Of Task Directive")
        scope_items = scope_combo.checkedItems() if scope_combo else []
        
        DocxExporter.export(
            save_path,
            self.extracted_project_name,
            self.intro_text_data,
            self.get_all_manual_notes_text(),
            scope_items,
            self.stakeholders_data,
            self.cert_task_data,
            self.combos,
            self.annexure3_image_path,
            self.get_docx_table_data()
        )
        self.last_docx_path = save_path

    def export_docx(self):
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Word", "Task_Directive.docx", "Word (*.docx)"
        )
        if not save_path:
            return
            
        try:
            self._write_docx(save_path)
            
            QMessageBox.information(
                self, "Success", "Document created successfully!"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Export failed: {str(e)}")
