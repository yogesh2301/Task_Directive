from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from PyQt6.QtCore import Qt


# Customize each tab's columns here independently: (Display Header, data key)
GENERIC_TAB_CONFIGS = {
    "annexure1_data": [
        ("Certifiable Item",             "certifiable_item"),
        ("Design Agency",                "design_agency"),
        ("Applicable Subpart",           "applicable_subpart"),
        ("Designated Directorate/RCMA",  "designated_directorate"),
        ("Dealing Officer(s)",           "dealing_officers"),
    ],
    "test_rigs_data": [
        ("Test Rig / Ground Equipment Name", "test_rig_name"),
        ("Utilizations Phase",               "utilization_phase"),
        ("Design Agency",                    "design_agency"),
        ("Designated RCMA",                  "designated_rcma"),
        ("Dealing Officer",                  "dealing_officer"),
    ],
    "aircraft_checks_data": [
        ("System Name",       "system_name"),
        ("Design Agency",     "design_agency"),
        ("Designated RCMA",   "designated_rcma"),
        ("Dealing Officer",   "dealing_officer"),
    ],
    "annexure2_data": [
        ("Name & Designation",   "name_designation"),
        ("Certification Centre", "certification_centre"),
        ("E-mail ID",            "email"),
        ("Phone Number",         "phone"),
    ],
}


class ProjectTablesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_app = parent
        self.generic_inputs = {}
        self.generic_tables = {}

        self.setWindowTitle("Project Tables Builder")
        self.resize(980, 620)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)

        self.tabs = QTabWidget()

        self._create_issue_details_tab()
        self._create_stakeholders_tab()
        self._create_cert_task_tab()
        self._create_generic_table_tab(
            "Annexure-1",
            "annexure1_data",
            "Work Assignment List of LRUs",
        )
        self._create_generic_table_tab(
            "Test Rigs",
            "test_rigs_data",
            "Test Rigs, Simulators and Ground Equipment Required",
        )
        self._create_generic_table_tab(
            "Aircraft Checks",
            "aircraft_checks_data",
            "Aircraft Integration Checks and Flight Clearance",
        )
        self._create_generic_table_tab(
            "Annexure-2",
            "annexure2_data",
            "Contact details of dealing officers and RDs",
        )

        layout.addWidget(self.tabs)

        main_btn_layout = QHBoxLayout()
        main_btn_layout.addStretch()

        self.close_btn = QPushButton("Save & Close")
        self.close_btn.setObjectName("PrimaryBtn")
        self.close_btn.clicked.connect(self.accept)

        main_btn_layout.addWidget(self.close_btn)
        layout.addLayout(main_btn_layout)

        self.refresh_all_tables()

    def _create_issue_details_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)

        form = QFormLayout()
        self.file_no_input = QLineEdit()
        self.file_no_input.setPlaceholderText("File No.")
        self.issue_no_input = QLineEdit()
        self.issue_no_input.setPlaceholderText("Issue No.")
        self.issue_date_input = QLineEdit()
        self.issue_date_input.setPlaceholderText("Date of Issue")

        self.file_no_input.textChanged.connect(
            lambda text: self._set_issue_detail("file_no", text)
        )
        self.issue_no_input.textChanged.connect(
            lambda text: self._set_issue_detail("issue_no", text)
        )
        self.issue_date_input.textChanged.connect(
            lambda text: self._set_issue_detail("date_of_issue", text)
        )

        form.addRow(QLabel("File No."), self.file_no_input)
        form.addRow(QLabel("Issue No."), self.issue_no_input)
        form.addRow(QLabel("Date of Issue"), self.issue_date_input)
        layout.addLayout(form)
        layout.addStretch()

        self.tabs.addTab(tab, "Issue Details")

    def _set_issue_detail(self, key, value):
        self.parent_app.issue_details_data[key] = value.strip()

    def _create_stakeholders_tab(self):
        tab_sh = QWidget()
        layout_sh = QVBoxLayout(tab_sh)
        layout_sh.setContentsMargins(15, 15, 15, 15)

        input_layout_sh = QHBoxLayout()

        self.org_combo = QComboBox()
        self.org_combo.setEditable(True)
        self.org_combo.addItems([
            "CEMILAC", "RCMA", "DGAQA", "DRDO", "User", "Production Agency"
        ])

        self.role_combo = QComboBox()
        self.role_combo.setEditable(True)
        self.role_combo.addItems([
            "Certification Authority", "Design Agency",
            "Production Agency", "Inspection Agency", "End User"
        ])

        self.act_input = QLineEdit()
        self.act_input.setPlaceholderText("e.g. Review & Approval")

        self.add_btn_sh = QPushButton("Add Row")
        self.add_btn_sh.setObjectName("PrimaryBtn")
        self.add_btn_sh.clicked.connect(self.add_stakeholder)

        input_layout_sh.addWidget(QLabel("Organisation:"))
        input_layout_sh.addWidget(self.org_combo, stretch=2)
        input_layout_sh.addWidget(QLabel("Role:"))
        input_layout_sh.addWidget(self.role_combo, stretch=2)
        input_layout_sh.addWidget(QLabel("Activities:"))
        input_layout_sh.addWidget(self.act_input, stretch=3)
        input_layout_sh.addWidget(self.add_btn_sh)

        layout_sh.addLayout(input_layout_sh)

        self.table_sh = QTableWidget(0, 4)
        self.table_sh.setHorizontalHeaderLabels([
            "Sl No.", "Organisation Agency", "Role", "Activities towards Certification"
        ])
        self._configure_table(self.table_sh)
        self.table_sh.cellChanged.connect(self.on_sh_cell_changed)

        layout_sh.addWidget(self.table_sh)
        layout_sh.addLayout(
            self._row_action_layout(self.delete_selected_sh, self.clear_all_sh)
        )

        self.tabs.addTab(tab_sh, "Stakeholders")

    def _create_cert_task_tab(self):
        tab_ct = QWidget()
        layout_ct = QVBoxLayout(tab_ct)
        layout_ct.setContentsMargins(15, 15, 15, 15)

        input_layout_ct = QHBoxLayout()

        self.act_input_ct = QLineEdit()
        self.act_input_ct.setPlaceholderText("e.g. Ground Testing")

        self.centre_combo = QComboBox()
        self.centre_combo.setEditable(True)
        self.centre_combo.addItems([
            "CEMILAC", "RCMA", "DGAQA", "DRDO", "User", "Production Agency"
        ])

        self.head_input = QLineEdit()
        self.head_input.setPlaceholderText("e.g. Director, RCMA")

        self.add_btn_ct = QPushButton("Add Row")
        self.add_btn_ct.setObjectName("PrimaryBtn")
        self.add_btn_ct.clicked.connect(self.add_task)

        input_layout_ct.addWidget(QLabel("Activity:"))
        input_layout_ct.addWidget(self.act_input_ct, stretch=2)
        input_layout_ct.addWidget(QLabel("Work Centre:"))
        input_layout_ct.addWidget(self.centre_combo, stretch=2)
        input_layout_ct.addWidget(QLabel("Resp. Head:"))
        input_layout_ct.addWidget(self.head_input, stretch=2)
        input_layout_ct.addWidget(self.add_btn_ct)

        layout_ct.addLayout(input_layout_ct)

        self.table_ct = QTableWidget(0, 4)
        self.table_ct.setHorizontalHeaderLabels([
            "Sl No.", "Certification Activity",
            "Certification Work Centre", "Responsible Head"
        ])
        self._configure_table(self.table_ct)
        self.table_ct.cellChanged.connect(self.on_ct_cell_changed)

        layout_ct.addWidget(self.table_ct)
        layout_ct.addLayout(
            self._row_action_layout(self.delete_selected_ct, self.clear_all_ct)
        )

        self.tabs.addTab(tab_ct, "Cert Task Allocation")

    def _create_generic_table_tab(self, tab_name, data_attr, title):
        columns = GENERIC_TAB_CONFIGS[data_attr]  # this tab's own column defs

        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)

        layout.addWidget(QLabel(title))

        # Input row — one labelled field per column
        input_layout = QHBoxLayout()
        self.generic_inputs[data_attr] = {}
        for label, key in columns:
            field = QLineEdit()
            field.setPlaceholderText(label)
            self.generic_inputs[data_attr][key] = field
            # input_layout.addWidget(QLabel(f"{label}:"))
            input_layout.addWidget(field, stretch=1)

        add_btn = QPushButton("Add Row")
        add_btn.setObjectName("PrimaryBtn")
        add_btn.clicked.connect(lambda: self.add_generic_row(data_attr))
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)

        # Table headers driven entirely by this tab's column config
        table = QTableWidget(0, len(columns) + 1)  # +1 for Sl No.
        table.setHorizontalHeaderLabels(
            ["Sl No."] + [label for label, _ in columns]
        )
        self._configure_table(table)
        table.cellChanged.connect(
            lambda row, col: self.on_generic_cell_changed(data_attr, row, col)
        )
        self.generic_tables[data_attr] = table

        layout.addWidget(table)
        layout.addLayout(
            self._row_action_layout(
                lambda: self.delete_selected_generic(data_attr),
                lambda: self.clear_all_generic(data_attr),
            )
        )

        self.tabs.addTab(tab, tab_name)

    def _configure_table(self, table):
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )

    def _row_action_layout(self, delete_handler, clear_handler):
        btn_layout = QHBoxLayout()

        del_btn = QPushButton("Delete Selected Row")
        del_btn.setObjectName("SecondaryBtn")
        del_btn.clicked.connect(delete_handler)

        clear_btn = QPushButton("Clear All")
        clear_btn.setObjectName("SecondaryBtn")
        clear_btn.clicked.connect(clear_handler)

        btn_layout.addWidget(del_btn)
        btn_layout.addWidget(clear_btn)
        btn_layout.addStretch()
        return btn_layout

    def refresh_all_tables(self):
        self.file_no_input.setText(
            self.parent_app.issue_details_data.get("file_no", "")
        )
        self.issue_no_input.setText(
            self.parent_app.issue_details_data.get("issue_no", "")
        )
        self.issue_date_input.setText(
            self.parent_app.issue_details_data.get("date_of_issue", "")
        )
        self.refresh_sh_table()
        self.refresh_ct_table()
        for data_attr in self.generic_tables:
            self.refresh_generic_table(data_attr)

    def add_stakeholder(self):
        org = self.org_combo.currentText().strip()
        role = self.role_combo.currentText().strip()
        act = self.act_input.text().strip()

        if org:
            self.parent_app.stakeholders_data.append({
                "org": org, "role": role, "activities": act
            })
            self.role_combo.setCurrentText("")
            self.act_input.clear()
            self.refresh_sh_table()

    def on_sh_cell_changed(self, row, column):
        if 0 <= row < len(self.parent_app.stakeholders_data):
            item = self.table_sh.item(row, column)
            if item:
                value = item.text().strip()
                if column == 1:
                    self.parent_app.stakeholders_data[row]['org'] = value
                elif column == 2:
                    self.parent_app.stakeholders_data[row]['role'] = value
                elif column == 3:
                    self.parent_app.stakeholders_data[row]['activities'] = value

    def delete_selected_sh(self):
        row = self.table_sh.currentRow()
        if row >= 0:
            self.parent_app.stakeholders_data.pop(row)
            self.refresh_sh_table()

    def clear_all_sh(self):
        self.parent_app.stakeholders_data.clear()
        self.refresh_sh_table()

    def refresh_sh_table(self):
        self.table_sh.blockSignals(True)
        self.table_sh.setRowCount(0)
        for i, data in enumerate(self.parent_app.stakeholders_data):
            self.table_sh.insertRow(i)
            self._set_sl_item(self.table_sh, i)
            self.table_sh.setItem(i, 1, QTableWidgetItem(data.get('org', '')))
            self.table_sh.setItem(i, 2, QTableWidgetItem(data.get('role', '')))
            self.table_sh.setItem(i, 3, QTableWidgetItem(data.get('activities', '')))
        self.table_sh.blockSignals(False)

    def add_task(self):
        act = self.act_input_ct.text().strip()
        centre = self.centre_combo.currentText().strip()
        head = self.head_input.text().strip()

        if act or centre or head:
            self.parent_app.cert_task_data.append({
                "activity": act, "centre": centre, "head": head
            })
            self.act_input_ct.clear()
            self.head_input.clear()
            self.refresh_ct_table()

    def on_ct_cell_changed(self, row, column):
        if 0 <= row < len(self.parent_app.cert_task_data):
            item = self.table_ct.item(row, column)
            if item:
                value = item.text().strip()
                if column == 1:
                    self.parent_app.cert_task_data[row]['activity'] = value
                elif column == 2:
                    self.parent_app.cert_task_data[row]['centre'] = value
                elif column == 3:
                    self.parent_app.cert_task_data[row]['head'] = value

    def delete_selected_ct(self):
        row = self.table_ct.currentRow()
        if row >= 0:
            self.parent_app.cert_task_data.pop(row)
            self.refresh_ct_table()

    def clear_all_ct(self):
        self.parent_app.cert_task_data.clear()
        self.refresh_ct_table()

    def refresh_ct_table(self):
        self.table_ct.blockSignals(True)
        self.table_ct.setRowCount(0)
        for i, data in enumerate(self.parent_app.cert_task_data):
            self.table_ct.insertRow(i)
            self._set_sl_item(self.table_ct, i)
            self.table_ct.setItem(i, 1, QTableWidgetItem(data.get('activity', '')))
            self.table_ct.setItem(i, 2, QTableWidgetItem(data.get('centre', '')))
            self.table_ct.setItem(i, 3, QTableWidgetItem(data.get('head', '')))
        self.table_ct.blockSignals(False)

    def add_generic_row(self, data_attr):
        values = {
            key: field.text().strip()
            for key, field in self.generic_inputs[data_attr].items()
        }
        if any(values.values()):
            getattr(self.parent_app, data_attr).append(values)
            for field in self.generic_inputs[data_attr].values():
                field.clear()
            self.refresh_generic_table(data_attr)

    def on_generic_cell_changed(self, data_attr, row, column):
        rows = getattr(self.parent_app, data_attr)
        if not 0 <= row < len(rows) or column == 0:
            return
        item = self.generic_tables[data_attr].item(row, column)
        if item:
            key = GENERIC_TAB_CONFIGS[data_attr][column - 1][1]
            rows[row][key] = item.text().strip()

    def delete_selected_generic(self, data_attr):
        table = self.generic_tables[data_attr]
        row = table.currentRow()
        if row >= 0:
            getattr(self.parent_app, data_attr).pop(row)
            self.refresh_generic_table(data_attr)

    def clear_all_generic(self, data_attr):
        getattr(self.parent_app, data_attr).clear()
        self.refresh_generic_table(data_attr)

    def refresh_generic_table(self, data_attr):
        table = self.generic_tables[data_attr]
        columns = GENERIC_TAB_CONFIGS[data_attr]
        table.blockSignals(True)
        table.setRowCount(0)
        for i, data in enumerate(getattr(self.parent_app, data_attr)):
            table.insertRow(i)
            self._set_sl_item(table, i)
            for col_idx, (_, key) in enumerate(columns, start=1):
                table.setItem(i, col_idx, QTableWidgetItem(data.get(key, "")))
        table.blockSignals(False)

    def _set_sl_item(self, table, row):
        sl = QTableWidgetItem(str(row + 1))
        sl.setFlags(sl.flags() & ~Qt.ItemFlag.ItemIsEditable)
        table.setItem(row, 0, sl)
