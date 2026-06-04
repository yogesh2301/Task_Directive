def get_light_theme():
    return """
    /* ================= GLOBAL ================= */
    QMainWindow, QWidget {
        background-color: #f8fafc;
        color: #191c1e;
        font-family: 'Inter', 'Segoe UI', 'Roboto', sans-serif;
        font-size: 13px;
    }

    QDialog {
        background-color: #f8fafc;
    }

    QMessageBox {
        background-color: #ffffff;
        color: #191c1e;
    }

    /* ================= TABLE ================= */
    QTableWidget {
        background-color: #ffffff;
        alternate-background-color: #f8fafc;
        gridline-color: #e2e8f0;
        selection-background-color: #0a4da2;
        selection-color: #ffffff;
        color: #191c1e;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
    }

    QHeaderView::section {
        background-color: #f1f5f9;
        color: #00264d;
        padding: 6px;
        border: 1px solid #cbd5e1;
        font-weight: bold;
    }

    /* ================= CARDS / GROUPS ================= */
    QFrame[class="Card"] {
        background: #ffffff;
        border-radius: 8px;
        border: 1px solid #cbd5e1;
    }

    QGroupBox {
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        margin-top: 15px;
        font-weight: bold;
        background: #ffffff;
        padding: 10px;
        color: #191c1e;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 0 6px;
        color: #00264d;
    }

    /* ================= PANEL TITLES ================= */
    QLabel#PanelTitle {
        font-size: 16px;
        font-weight: bold;
        color: #00264d;
        padding: 8px 0px;
        border-bottom: 2px solid #00264d;
    }

    /* ================= BUTTONS ================= */
    QPushButton {
        padding: 8px 15px;
        border-radius: 6px;
        font-weight: bold;
        border: 1px solid transparent;
    }

    QPushButton#PrimaryBtn {
        background-color: #00264d;
        color: #ffffff;
        border: 1px solid #00264d;
    }

    QPushButton#PrimaryBtn:hover {
        background-color: #0a4da2;
        border-color: #0a4da2;
    }

    QPushButton#SecondaryBtn {
        background-color: #ffffff;
        color: #00264d;
        border: 1px solid #cbd5e1;
    }

    QPushButton#SecondaryBtn:hover {
        background-color: #f1f5f9;
        border-color: #0a4da2;
    }

    QPushButton#ToolbarBtn {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 22px;
        font-size: 18px;
        color: #334155;
    }

    QPushButton#ToolbarBtn:hover {
        background-color: #f1f5f9;
    }

    QPushButton#SuccessBtn {
        background-color: #0a4da2;
        color: #ffffff;
        border-color: #0a4da2;
    }

    QPushButton#SuccessBtn:hover {
        background-color: #00264d;
    }

    /* ================= INPUT ================= */
    QLineEdit, QTextEdit {
        background: #ffffff;
        color: #191c1e;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        padding: 8px;
    }

    QLineEdit:focus, QTextEdit:focus {
        border: 1px solid #0a4da2;
    }

    QComboBox {
        background: #ffffff;
        color: #191c1e;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        padding: 6px 8px;
    }

    QComboBox:focus {
        border: 1px solid #0a4da2;
    }

    QComboBox QAbstractItemView {
        background: #ffffff;
        color: #191c1e;
        border: 1px solid #cbd5e1;
        selection-background-color: #e7f1ff;
        selection-color: #00264d;
    }

    /* ================= LABEL ================= */
    QLabel {
        color: #334155;
    }

    QLabel#HeaderLabel {
        font-size: 16px;
        font-weight: bold;
        color: #00264d;
    }

    /* ================= CHECKBOX ================= */
    QCheckBox {
        color: #191c1e;
        spacing: 6px;
    }

    QCheckBox::indicator {
        width: 16px;
        height: 16px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
    }

    QCheckBox::indicator:checked {
        background: #0a4da2;
        border: 1px solid #0a4da2;
    }

    /* ================= SCROLL ================= */
    QScrollBar:vertical {
        border: none;
        background: #f1f5f9;
        width: 10px;
    }

    QScrollBar::handle:vertical {
        background: #cbd5e1;
        min-height: 20px;
        border-radius: 5px;
    }

    QScrollBar::handle:vertical:hover {
        background: #94a3b8;
    }

    QScrollBar::add-line:vertical,
    QScrollBar::sub-line:vertical {
        border: none;
        background: none;
    }

    /* ================= SPLITTER ================= */
    QSplitter::handle {
        background: #cbd5e1;
        width: 4px;
    }

    /* ================= TABS ================= */
    QTabWidget::pane {
        border: 1px solid #cbd5e1;
        background: #ffffff;
        border-radius: 6px;
    }

    QTabBar::tab {
        background: #f1f5f9;
        color: #334155;
        padding: 8px 16px;
        margin: 2px;
        border-radius: 4px;
    }

    QTabBar::tab:selected {
        background: #00264d;
        color: #ffffff;
    }

    /* ================= PROGRESS ================= */
    QProgressBar {
        border: 1px solid #cbd5e1;
        border-radius: 5px;
        text-align: center;
        background: #ffffff;
        color: #191c1e;
    }

    QProgressBar::chunk {
        background-color: #0a4da2;
        border-radius: 5px;
    }
    """
