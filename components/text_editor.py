from PyQt6.QtWidgets import QTextEdit, QApplication, QMenu
from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtGui import QTextCursor


class NotesEditor(QTextEdit):
    def canInsertFromMimeData(self, source):
        if source.hasImage() or source.hasText():
            return True
        return super().canInsertFromMimeData(source)

    def insertFromMimeData(self, source):
        if source.hasImage():
            clipboard = QApplication.clipboard()
            image = clipboard.image()
            if not image.isNull():
                cursor = self.textCursor()
                if image.width() > 500:
                    image = image.scaledToWidth(
                        500, Qt.TransformationMode.SmoothTransformation
                    )
                cursor.insertImage(image)
                self.insertPlainText("\n")
                return
        if source.hasText():
            self.insertPlainText(source.text())
            return
        super().insertFromMimeData(source)


class SelectableTextPreview(QTextEdit):
    text_extracted = pyqtSignal(str)
    summary_requested = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setStyleSheet("""
            QTextEdit { 
                background: #ffffff; color: #191c1e; 
                font-size: 14px; line-height: 1.6; 
                padding: 15px; border: none; 
            }
        """)

    def contextMenuEvent(self, event):
        menu = self.createStandardContextMenu()
        if self.textCursor().hasSelection():
            menu.addSeparator()
            extract_action = menu.addAction("Extract to Notes")
            summarize_action = menu.addAction("Summarize Selection")
            
            action = menu.exec(event.globalPos())
            
            if action == extract_action:
                self.text_extracted.emit(self.textCursor().selectedText())
            elif action == summarize_action:
                self.summary_requested.emit(self.textCursor().selectedText())
        else:
            menu.exec(event.globalPos())
