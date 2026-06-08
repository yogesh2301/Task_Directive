import fitz
from PyQt6.QtWidgets import QLabel, QRubberBand, QMenu, QApplication
from PyQt6.QtCore import Qt, pyqtSignal, QRect, QPoint, QSize, QTimer
from PyQt6.QtGui import QPixmap

try:
    from PyQt6.QtWebEngineWidgets import QWebEngineView
    from PyQt6.QtWebEngineCore import QWebEngineSettings, QWebEnginePage
    HAS_WEBENGINE = True

    class PDFWebEngineView(QWebEngineView):
        text_extracted = pyqtSignal(str)
        summary_requested = pyqtSignal(str)

        def contextMenuEvent(self, event):
            menu = QMenu(self)
            copy_action = menu.addAction("Copy Selected Text")
            copy_action.triggered.connect(
                lambda: self.triggerPageAction(QWebEnginePage.WebAction.Copy)
            )
            
            menu.addSeparator()
            paste_action = menu.addAction("Paste to Notes")
            paste_action.triggered.connect(self.copy_and_paste)
            
            summarize_action = menu.addAction("Summarize Copied Text")
            summarize_action.triggered.connect(self.copy_and_summarize)
            
            menu.addSeparator()
            reload_action = menu.addAction("Reload PDF")
            reload_action.triggered.connect(self.reload)
            menu.exec(event.globalPos())

        def copy_and_paste(self):
            self.triggerPageAction(QWebEnginePage.WebAction.Copy)
            QTimer.singleShot(
                150, 
                lambda: self.text_extracted.emit(QApplication.clipboard().text())
            )

        def copy_and_summarize(self):
            self.triggerPageAction(QWebEnginePage.WebAction.Copy)
            QTimer.singleShot(
                150, 
                lambda: self.summary_requested.emit(QApplication.clipboard().text())
            )

except ImportError:
    HAS_WEBENGINE = False
    PDFWebEngineView = None


class PdfPageLabel(QLabel):
    text_extracted = pyqtSignal(str)
    image_extracted = pyqtSignal(QPixmap)
    annexure3_image_extracted = pyqtSignal(QPixmap)
    summary_requested = pyqtSignal(str)
    region_ocr_requested = pyqtSignal(QPixmap, QRect, object)  # Change 17: Emit page-aware region OCR selection details.

    def __init__(self, page, scroll_area, parent=None):
        super().__init__(parent)
        self.page = page
        self.scroll_area = scroll_area
        self.rubber_band = QRubberBand(QRubberBand.Shape.Rectangle, self)
        self.origin = QPoint()
        self.pan_start_pos = None

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.origin = event.pos()
            self.rubber_band.setGeometry(QRect(self.origin, QSize()))
            self.rubber_band.show()
        elif event.button() in [Qt.MouseButton.RightButton, Qt.MouseButton.MiddleButton]:
            self.pan_start_pos = event.globalPosition().toPoint()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)

    def mouseMoveEvent(self, event):
        if not self.origin.isNull() and (event.buttons() & Qt.MouseButton.LeftButton):
            self.rubber_band.setGeometry(QRect(self.origin, event.pos()).normalized())
        elif self.pan_start_pos and (event.buttons() & (Qt.MouseButton.RightButton | Qt.MouseButton.MiddleButton)):
            current_pos = event.globalPosition().toPoint()
            delta = current_pos - self.pan_start_pos
            h_bar = self.scroll_area.horizontalScrollBar()
            v_bar = self.scroll_area.verticalScrollBar()
            h_bar.setValue(h_bar.value() - delta.x())
            v_bar.setValue(v_bar.value() - delta.y())
            self.pan_start_pos = current_pos

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            rect = self.rubber_band.geometry()
            if rect.width() > 10 and rect.height() > 10:
                self.show_context_menu(event.pos(), rect)
            else:
                self.rubber_band.hide()
            self.origin = QPoint()
        elif event.button() in [Qt.MouseButton.RightButton, Qt.MouseButton.MiddleButton]:
            self.unsetCursor()
            self.pan_start_pos = None

    def get_text_from_rect(self, rect):
        if self.pixmap() and self.page:
            scale_x = self.page.rect.width / self.pixmap().width()
            scale_y = self.page.rect.height / self.pixmap().height()
            pdf_rect = fitz.Rect(
                rect.left() * scale_x, rect.top() * scale_y,
                rect.right() * scale_x, rect.bottom() * scale_y
            )
            pdf_rect = pdf_rect + (-3, -3, 3, 3)
            return self.page.get_textbox(pdf_rect).strip()
        return ""

    def show_context_menu(self, pos, rect):
        self.rubber_band.hide()
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                color: #191c1e;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 6px;
            }
            QMenu::item {
                padding: 8px 30px 8px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #e7f1ff;
                color: #00264d;
            }
            QMenu::separator {
                height: 1px;
                background: #cbd5e1;
                margin: 4px 10px;
            }
        """)
        
        # Image extraction
        copy_img_action = menu.addAction("Extract as Image")
        
        annexure3_img_action = menu.addAction("Add as image to PBS/Annexure-3")

        # Change 9: Add a dedicated copy action so selected PDF text can be copied directly to the clipboard.
        copy_text_action = menu.addAction("Copy Text")
        extract_text_action = menu.addAction("Extract Text to Notes")
        
        menu.addSeparator()
        
        # Change 9: Keep the existing extract-to-notes action separate from the clipboard copy action.
        ocr_action = menu.addAction("Extract to File Details (OCR)")
        ocr_action.setToolTip("Extract metadata from selection to File Details tab")
        
        menu.addSeparator()
        
        # Summarize
        summarize_action = menu.addAction("Summarize and Set as Intro")
        
        action = menu.exec(self.mapToGlobal(pos))
        
        if action == copy_img_action:
            if self.pixmap():
                self.image_extracted.emit(self.pixmap().copy(rect))

        elif action == annexure3_img_action:
            if self.pixmap():
                self.annexure3_image_extracted.emit(self.pixmap().copy(rect))
        
        elif action == copy_text_action:
            text = self.get_text_from_rect(rect)
            if text:
                QApplication.clipboard().setText(text)
                # Change 9: Copy selected text directly to the clipboard without modifying notes.
        elif action == extract_text_action:
            text = self.get_text_from_rect(rect)
            if text:
                self.text_extracted.emit(text)
                # Change 9: Send selected text to notes as before.
        
        elif action == ocr_action:
            # Change 18: Emit the region OCR signal so selected text can be extracted into File Details.
            if self.pixmap():
                self.region_ocr_requested.emit(self.pixmap(), rect, self.page)
        
        elif action == summarize_action:
            text = self.get_text_from_rect(rect)
            if text:
                self.summary_requested.emit(text)
