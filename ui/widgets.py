from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QFrame, QDialog,
    QListWidget, QListWidgetItem, QLineEdit, QLabel, QHBoxLayout,
    QSizePolicy, QToolButton, QApplication
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QCursor

class CollapsibleBox(QWidget):
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

        self.toggle_btn = QPushButton(f"v  {title}")
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                text-align: left; font-weight: bold; background: #f1f5f9; 
                color: #00264d; padding: 12px; font-size: 14px; 
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                border: 1px solid #cbd5e1; border-bottom: none;
            }
            QPushButton:hover { background: #e2e8f0; }
        """)
        self.toggle_btn.clicked.connect(self.on_press)
        self.layout.addWidget(self.toggle_btn)

        self.content_area = QFrame()
        self.content_area.setStyleSheet("""
            QFrame { 
                background: #ffffff; border: 1px solid #cbd5e1; 
                border-bottom-left-radius: 8px; 
                border-bottom-right-radius: 8px; 
            }
        """)
        self.content_layout = QVBoxLayout(self.content_area)
        self.content_layout.setContentsMargins(15, 15, 15, 15)
        self.content_layout.setSpacing(15)
        self.layout.addWidget(self.content_area)
        
        self.is_expanded = True

    def on_press(self):
        self.is_expanded = not self.is_expanded
        if self.is_expanded:
            self.toggle_btn.setText(self.toggle_btn.text().replace("> ", "v "))
            self.content_area.show()
        else:
            self.toggle_btn.setText(self.toggle_btn.text().replace("v ", "> "))
            self.content_area.hide()


class DropdownPopup(QDialog):
    def __init__(self, parent_btn, options):
        super().__init__(parent_btn, Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint)
        self.parent_btn = parent_btn
        self.setStyleSheet("""
            QDialog { 
                background: #ffffff; border: 1px solid #cbd5e1; 
                border-radius: 6px; 
            }
            QListWidget { 
                background: transparent; border: none; 
                color: #191c1e; outline: none; font-size: 13px; 
            }
            QListWidget::item { 
                padding: 6px 6px 6px 4px; border-radius: 4px; color: #191c1e;
                min-height: 22px;
            }
            QListWidget::item:hover { background: #f1f5f9; }
            QListWidget::item:selected { background: #e7f1ff; color: #00264d; }
            QListWidget::indicator {
                width: 15px;
                height: 15px;
                border: 2px solid #94a3b8;
                border-radius: 3px;
                background: #ffffff;
            }
            QListWidget::indicator:unchecked:hover {
                border: 2px solid #0a4da2;
                background: #e7f1ff;
            }
            QListWidget::indicator:checked {
                background-color: #0a4da2;
                border: 2px solid #0a4da2;
                /* Keep the checked indicator simple for reliable rendering. */
                image: url(none);
            }
            QListWidget::indicator:checked:hover {
                background-color: #00264d;
                border: 2px solid #00264d;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header_label = QLabel("SELECT OPTIONS")
        header_label.setStyleSheet("font-size: 11px; color: #64748b; font-weight: bold; letter-spacing: 0.5px;")
        layout.addWidget(header_label)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search options...")
        self.search_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #cbd5e1; border-radius: 5px;
                padding: 6px 8px; font-size: 13px; color: #191c1e;
                background: #ffffff;
            }
            QLineEdit:focus { border: 1px solid #0a4da2; }
        """)
        self.search_input.textChanged.connect(self.filter_options)
        layout.addWidget(self.search_input)

        action_layout = QHBoxLayout()
        action_layout.setSpacing(6)

        self.select_all_btn = QPushButton("Select All")
        self.clear_btn = QPushButton("Clear")
        self.add_btn = QPushButton("+ Add")
        for btn in [self.select_all_btn, self.clear_btn, self.add_btn]:
            btn.setAutoDefault(False)
            btn.setDefault(False)
            btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            btn.setStyleSheet("""
                QPushButton {
                    background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 5px;
                    color: #00264d; font-size: 12px; font-weight: bold;
                    padding: 6px 8px; text-align: center;
                }
                QPushButton:hover { background: #e7f1ff; border-color: #0a4da2; }
                QPushButton:pressed { background: #dbeafe; }
            """)

        self.select_all_btn.clicked.connect(self.select_visible_options)
        self.clear_btn.clicked.connect(self.clear_visible_options)
        self.add_btn.clicked.connect(self.show_new_input)
        action_layout.addWidget(self.select_all_btn)
        action_layout.addWidget(self.clear_btn)
        action_layout.addWidget(self.add_btn)
        layout.addLayout(action_layout)

        self.new_input = QLineEdit()
        self.new_input.setPlaceholderText("Type and press Enter...")
        self.new_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #0a4da2; border-radius: 5px;
                padding: 5px 8px; font-size: 13px; color: #191c1e;
                background: #ffffff;
            }
        """)
        self.new_input.hide()
        self.new_input.returnPressed.connect(self.save_new_option)
        layout.addWidget(self.new_input)

        self.list_widget = QListWidget()
        self.list_widget.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.list_widget.setWordWrap(True)
        self.list_widget.setUniformItemSizes(False)
        self.list_widget.setMinimumHeight(120)
        self.list_widget.setMaximumHeight(260)
        for opt in options:
            if opt != "Enter Other...":
                self.add_item(opt)
                
        self.list_widget.itemPressed.connect(self.on_item_pressed)
        self.list_widget.itemChanged.connect(self.parent_btn.update_display)
        layout.addWidget(self.list_widget)

        self.done_btn = QPushButton("Done")
        self.done_btn.setAutoDefault(False)
        self.done_btn.setDefault(False)
        self.done_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.done_btn.setStyleSheet("""
            QPushButton {
                background: #00264d; border: 1px solid #00264d; border-radius: 5px;
                color: #ffffff; font-size: 12px; font-weight: bold;
                padding: 6px 8px;
            }
            QPushButton:hover { background: #0a4da2; }
        """)
        self.done_btn.clicked.connect(self.hide)
        layout.addWidget(self.done_btn)

    def hideEvent(self, event):
        self.parent_btn.set_popup_open(False)
        super().hideEvent(event)

    def filter_options(self, text):
        text = text.strip().lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setHidden(text not in item.text().lower())

    def select_visible_options(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                item.setCheckState(Qt.CheckState.Checked)
        self.parent_btn.update_display()

    def clear_visible_options(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                item.setCheckState(Qt.CheckState.Unchecked)
        self.parent_btn.update_display()

    def on_item_pressed(self, item):
        pos = self.list_widget.viewport().mapFromGlobal(QCursor.pos())
        if pos.x() > 25:
            state = item.checkState()
            new_state = Qt.CheckState.Unchecked if state == Qt.CheckState.Checked else Qt.CheckState.Checked
            item.setCheckState(new_state)

    def add_item(self, text, checked=False):
        item = QListWidgetItem(text)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self.list_widget.addItem(item)

    def show_new_input(self):
        self.new_input.show()
        self.new_input.setFocus()

    def save_new_option(self):
        text = self.new_input.text().strip()
        if text:
            existing = [self.list_widget.item(i).text() for i in range(self.list_widget.count())]
            if text not in existing:
                self.add_item(text, checked=True)
                self.filter_options(self.search_input.text())
        self.new_input.clear()
        self.new_input.hide()
        self.parent_btn.update_display()


class PopupMultiSelect(QPushButton):
    def __init__(self, options, parent=None):
        super().__init__("Select options", parent)
        self.setStyleSheet("""
            QPushButton {
                text-align: left; padding: 8px 34px 8px 10px; background: #ffffff;
                border: 1px solid #cbd5e1; border-radius: 6px; 
                color: #191c1e; font-size: 13px;
            }
            QPushButton:hover { border: 1px solid #0a4da2; }
        """)
        self.options = options
        self.popup = None
        self.is_popup_open = False
        self.setMinimumHeight(42)
        self.setMaximumHeight(78)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self.arrow_btn = QToolButton(self)
        self.arrow_btn.setText("v")
        self.arrow_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.arrow_btn.setStyleSheet("""
            QToolButton {
                border: none;
                border-left: 1px solid #cbd5e1;
                color: #334155;
                background: transparent;
                font-size: 11px;
                padding: 0px;
            }
            QToolButton:hover {
                background: #f1f5f9;
                color: #00264d;
            }
        """)
        self.arrow_btn.clicked.connect(self.toggle_popup)

    def sizeHint(self):
        return QSize(240, 42)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        arrow_width = 28
        self.arrow_btn.setGeometry(
            self.width() - arrow_width - 1,
            1,
            arrow_width,
            self.height() - 2
        )

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.show_popup()

    def toggle_popup(self):
        if self.popup and self.popup.isVisible():
            self.popup.hide()
        else:
            self.show_popup()

    def set_popup_open(self, is_open):
        self.is_popup_open = is_open
        self.arrow_btn.setText("^" if is_open else "v")

    def show_popup(self):
        if not self.popup:
            self.popup = DropdownPopup(self, self.options)

        screen = QApplication.screenAt(self.mapToGlobal(self.rect().center()))
        available_rect = screen.availableGeometry() if screen else QApplication.primaryScreen().availableGeometry()
        popup_width = max(self.width(), 360)
        self.popup.setFixedWidth(popup_width)
        self.popup.adjustSize()

        below = self.mapToGlobal(self.rect().bottomLeft())
        above = self.mapToGlobal(self.rect().topLeft())
        popup_height = self.popup.sizeHint().height()
        max_below_height = available_rect.bottom() - below.y() - 8
        max_above_height = above.y() - available_rect.top() - 8
        target_x = min(max(below.x(), available_rect.left()), available_rect.right() - popup_width)

        if max_below_height >= min(popup_height, 260) or max_below_height >= max_above_height:
            target_y = below.y() + 2
            max_height = max(220, max_below_height)
        else:
            max_height = max(220, max_above_height)
            target_y = above.y() - min(popup_height, max_height) - 2

        self.popup.setMaximumHeight(max_height)
        self.popup.move(target_x, max(available_rect.top() + 4, target_y))
        self.popup.show()
        self.set_popup_open(True)
        self.popup.search_input.setFocus()

    def update_display(self):
        checked = self.checkedItems()
        if checked:
            preview = checked[0]
            if len(checked) > 1:
                preview = f"{preview} + {len(checked) - 1} more"
            self.setText(f"{len(checked)} selected: {preview}")
            self.setToolTip("\n".join(checked))
        else:
            self.setText("Select options")
            self.setToolTip("")

    def checkedItems(self):
        if not self.popup: 
            return []
        items = []
        for i in range(self.popup.list_widget.count()):
            item = self.popup.list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                items.append(item.text())
        return items
