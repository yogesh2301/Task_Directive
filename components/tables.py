# -*- coding: utf-8 -*-
# =============================================================================
# components/tables.py
#
# CHANGE HISTORY (all changes documented inline below)
# -----------------------------------------------------------------------------
# CHANGE 1 — Excel file replaced
#   The original Employee_List_50.xlsx was deleted and replaced with a new
#   Employee_List_50(1).xlsx that was copied into the project root as
#   Employee_List_50.xlsx.  The new file has a different column layout that
#   maps directly to the Annexure-2 table columns:
#       col 0 → Sl No
#       col 1 → Name & Designation  (e.g. "Employee 1 - Executive")
#       col 2 → Certification Centre (serves as the department / filter key)
#       col 3 → E-mail ID
#       col 4 → Phone Number
#       col 5 → Office Mail
#
# CHANGE 2 — _EMP_FIELD dict removed
#   The previous version had a module-level _EMP_FIELD dictionary that mapped
#   logical field names (like "name", "department", "email") to lambda accessors
#   on an employee row.  This was used by the generic _make_employee_selector
#   helper.  Since that helper was removed (see CHANGE 4), _EMP_FIELD is no
#   longer needed and has been deleted.
#
# CHANGE 3 — EMPLOYEE_COLUMNS constant removed
#   A module-level EMPLOYEE_COLUMNS list (used only by the old Employee
#   Directory tab) was removed because that tab was also removed (see CHANGE 5).
#
# CHANGE 4 — _make_employee_selector helper removed
#   Previously a shared _make_employee_selector(on_select) method added a
#   "Dept: / Employee:" bar at the TOP of every tab (Stakeholders, Cert Task
#   Allocation, Annexure-1, Test Rigs, Aircraft Checks, Annexure-2).  The user
#   requested this bar be removed from all tabs.  The method has been deleted
#   entirely.  The Annexure-2 tab now uses its own inline dropdowns instead
#   (see CHANGE 7).
#
# CHANGE 5 — Employee Directory tab removed
#   A standalone "Employee Directory" tab was previously added to the dialog so
#   users could browse all employees filtered by department.  The user asked for
#   this tab to be removed.  The _create_employee_directory_tab and
#   _filter_employee_table methods have been deleted, and the call to
#   _create_employee_directory_tab() in __init__ has been removed.
#
# CHANGE 6 — emp_fill_map parameter removed from _create_generic_table_tab
#   The previous _create_generic_table_tab signature accepted an optional
#   emp_fill_map dict that, when provided, inserted the _make_employee_selector
#   bar and wired it to auto-fill specific columns.  Because CHANGE 4 removed
#   the selector, this parameter is no longer meaningful and has been dropped.
#   Annexure-1, Test Rigs and Aircraft Checks now call _create_generic_table_tab
#   without any employee arguments, restoring them to plain text input rows.
#
# CHANGE 7 — _create_annexure2_tab dedicated method added (NEW)
#   Annexure-2 is no longer built by the generic _create_generic_table_tab.
#   A new dedicated _create_annexure2_tab method builds the tab with employee
#   dropdowns integrated directly into the input row:
#     • "Dept:" QComboBox  — filters the Name & Designation combo by department
#     • Name & Designation QComboBox — lists employees for the selected dept;
#       selecting one auto-fills the three remaining fields
#     • Certification Centre QLineEdit — auto-filled from emp[2], still editable
#     • E-mail ID QLineEdit           — auto-filled from emp[3], still editable
#     • Phone Number QLineEdit        — auto-filled from emp[4], still editable
#   The table, cell-edit handling, delete and clear buttons are wired the same
#   way as all other tabs so existing save/export logic is unaffected.
#
# CHANGE 8 — _add_annexure2_row method added (NEW)
#   A dedicated Add Row handler for Annexure-2 reads the name combo and the
#   three QLineEdit fields, appends a dict to self.parent_app.annexure2_data,
#   resets the inputs, and refreshes the table.
# =============================================================================

import os

# CHANGE 1 (continued): openpyxl is used to read the replacement Excel file.
# It is already available in the project's requirements (via pandas dependency).
import openpyxl
from PyQt6.QtCore import Qt
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
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

# -----------------------------------------------------------------------------
# CHANGE 1 (continued): Path to the replacement Excel file.
# os.path.abspath(__file__) resolves to the absolute path of THIS file
# (components/tables.py).  Two dirname() calls walk up to the project root,
# then "Employee_List_50.xlsx" is appended.  This works regardless of the
# current working directory when the app is launched.
#
# *** WHEN NEW EXCEL — STEP 1 of 6 ***
#   Replace "Employee_List_50.xlsx" below with the filename of your new Excel.
#   The file must be placed in the project root directory
#   (same folder as main.py).  Example:
#       "My_New_Staff_List.xlsx"
# -----------------------------------------------------------------------------
_EXCEL_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "Employee_List_50.xlsx",  # <-- WHEN NEW EXCEL: change this filename
)

# -----------------------------------------------------------------------------
# CHANGE 1 (continued): New Excel column layout reference.
# The replacement file columns differ from the original.  They now match the
# Annexure-2 table columns directly, which is why Annexure-2 can use the data
# without any field-name remapping.
#
# *** WHEN NEW EXCEL — STEP 2 of 6 ***
#   Open your new Excel file and note the column ORDER (0-based index).
#   Update the reference table below to match your file, then use those
#   new index numbers everywhere you see an emp[N] reference in the code.
#   The five places are marked with "WHEN NEW EXCEL" in STEPS 3-6.
#
#   Current column layout (Employee_List_50.xlsx):
#   row[0] = Sl No
#   row[1] = Name & Designation   (e.g. "Employee 1 - Executive")
#   row[2] = Certification Centre  (department: HR, Finance, IT …)  ← dept key
#   row[3] = E-mail ID
#   row[4] = Phone Number
#   row[5] = Office Mail           (not displayed in Annexure-2 but kept in row)
#
#   Example — if your new file has columns in a different order, update like:
#   row[0] = Sl No
#   row[1] = Department            ← becomes the dept key  (was row[2])
#   row[2] = Employee Name
#   row[3] = Designation
#   row[4] = Email
#   row[5] = Phone
#   … then change every emp[N] below accordingly.
# -----------------------------------------------------------------------------


def _load_employee_data():
    """
    Read the employee Excel file and return all data rows as a list of
    string lists (one inner list per employee), skipping the header row.

    Returns an empty list silently if the file is missing or unreadable,
    so the app still opens even if the Excel is not present.
    """
    data = []
    try:
        # read_only=True  — faster; does not load formatting
        # data_only=True  — returns cell values, not formulas
        wb = openpyxl.load_workbook(_EXCEL_PATH, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        for row in rows[1:]:  # rows[0] is the header; skip it
            if any(row):  # skip completely empty rows
                # Convert every value to str so the UI never receives None
                data.append([str(v) if v is not None else "" for v in row])
        wb.close()
    except Exception:
        # Broad catch: file not found, corrupt file, permission error, etc.
        pass
    return data


# -----------------------------------------------------------------------------
# GENERIC_TAB_CONFIGS
# Defines the columns for each "generic" tab as a list of
#   (Display Header string, data-dict key string) tuples.
# Used by _create_generic_table_tab to build input fields and table headers
# without duplicating code for each tab.
#
# NOTE (CHANGE 7): "annexure2_data" is still listed here because the
# on_generic_cell_changed and refresh_generic_table helpers reference this dict
# when the user edits a cell in the Annexure-2 table directly.  The INPUT FORM
# for Annexure-2, however, is built by _create_annexure2_tab (not the generic
# method), so it never calls _create_generic_table_tab for "annexure2_data".
# -----------------------------------------------------------------------------
GENERIC_TAB_CONFIGS = {
    "annexure1_data": [
        ("Certifiable Item", "certifiable_item"),
        ("Design Agency", "design_agency"),
        ("Designated Directorate/RCMA", "designated_directorate"),
        ("Dealing Officer HW", "dealing_officer_hw"),
        ("Dealing Officer CH", "dealing_officer_ch"),
        ("Dealing Officer SW", "dealing_officer_sw"),
    ],
    "test_rigs_data": [
        ("Test Rig / Ground Equipment Name", "test_rig_name"),
        ("Utilizations Phase", "utilization_phase"),
        ("Design Agency", "design_agency"),
        ("Designated RCMA", "designated_rcma"),
        ("Dealing Officer", "dealing_officer"),
    ],
    "aircraft_checks_data": [
        ("System Name", "system_name"),
        ("Design Agency", "design_agency"),
        ("Designated RCMA", "designated_rcma"),
        ("Dealing Officer", "dealing_officer"),
    ],
    # CHANGE 7 note: Annexure-2 input form is built by _create_annexure2_tab.
    # This entry is retained only so that cell-edit and table-refresh helpers
    # (on_generic_cell_changed, refresh_generic_table) can look up column keys.
    #
    # *** WHEN NEW EXCEL — STEP 3 of 6 ***
    #   If your new Excel has different column names for the Annexure-2 table,
    #   update the tuples below:
    #       ("Display Header shown in the table", "dict key used in code")
    #   Rules:
    #   - The NUMBER of tuples must match the number of data columns your
    #     Annexure-2 table will have (not counting "Sl No.").
    #   - The dict key (second item) must also match the keys used in
    #     _add_annexure2_row (see STEP 6) and parent_app.annexure2_data.
    #   - Do NOT change the "Sl No." column — it is added automatically.
    "annexure2_data": [
        (
            "Name & Designation",
            "name_designation",
        ),  # <-- WHEN NEW EXCEL: update if column name differs
        (
            "Certification Centre",
            "certification_centre",
        ),  # <-- WHEN NEW EXCEL: update if column name differs
        ("E-mail ID", "email"),  # <-- WHEN NEW EXCEL: update if column name differs
        ("Phone Number", "phone"),  # <-- WHEN NEW EXCEL: update if column name differs
    ],
}


class ProjectTablesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_app = parent

        # Dicts filled by each tab-creation method:
        #   generic_inputs[data_attr][col_key] = QLineEdit widget
        #   generic_tables[data_attr]          = QTableWidget widget
        self.generic_inputs = {}
        self.generic_tables = {}

        # -----------------------------------------------------------------
        # CHANGE 1 / CHANGE 7: Load employee data once at dialog creation.
        # Three structures are built from the flat list for fast access:
        #
        #   _all_employees  — the full list of rows, used when "All Depts" is
        #                     selected in the Annexure-2 dept filter.
        #
        #   _all_depts      — sorted list of unique Certification Centre values
        #                     (e.g. ['Admin','Finance','HR','IT', …]).
        #                     Populated into the Annexure-2 dept combo.
        #
        #   _emp_by_dept    — dict mapping each dept name → list of employee
        #                     rows in that dept.  Used to filter the name combo
        #                     when the user selects a specific department.
        # -----------------------------------------------------------------
        self._all_employees = _load_employee_data()
        # *** WHEN NEW EXCEL — STEP 4 of 6 ***
        #   Change the index [2] in e[2] and emp[2] below to whichever column
        #   in your new Excel contains the DEPARTMENT / FILTER category.
        #   Example: if dept is in the 2nd column (0-based index 1), use e[1].
        #   This controls which values populate the "Dept:" dropdown.
        self._all_depts = sorted(
            {e[2] for e in self._all_employees if e[2]}
        )  # <-- WHEN NEW EXCEL: change [2] to your dept column index
        self._emp_by_dept = {}
        for emp in self._all_employees:
            # setdefault creates the list on first encounter for each dept
            self._emp_by_dept.setdefault(emp[2], []).append(
                emp
            )  # <-- WHEN NEW EXCEL: change [2] to match dept column index above

        self.setWindowTitle("Project Tables Builder")
        self.resize(980, 620)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)

        self.tabs = QTabWidget()

        # Build each tab in display order
        self._create_issue_details_tab()
        self._create_stakeholders_tab()
        self._create_cert_task_tab()

        # RCMA + Dealing Officer dropdowns — dedicated tab methods
        self._create_annexure1_tab()
        self._create_test_rigs_tab()
        self._create_aircraft_checks_tab()

        # CHANGE 7: Annexure-2 now uses its own dedicated method instead of
        # the generic one, so that it can embed employee dropdowns directly
        # inside the input row rather than relying on the removed selector bar.
        self._create_annexure2_tab()

        # CHANGE 5: The call to self._create_employee_directory_tab() that
        # previously appeared here has been removed.  That method and its
        # companion _filter_employee_table have been deleted entirely.

        layout.addWidget(self.tabs)

        main_btn_layout = QHBoxLayout()
        main_btn_layout.addStretch()

        self.close_btn = QPushButton("Save & Close")
        self.close_btn.setObjectName("PrimaryBtn")
        self.close_btn.clicked.connect(self.accept)

        main_btn_layout.addWidget(self.close_btn)
        layout.addLayout(main_btn_layout)

        # Populate every table with any data already stored in the parent app
        self.refresh_all_tables()

    # =========================================================================
    # Issue Details tab — unchanged from original
    # =========================================================================
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

        # Each field syncs its value to the parent app's issue_details_data
        # dict immediately on every keystroke via textChanged signal.
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
        # Strips whitespace before storing so exported documents are clean
        self.parent_app.issue_details_data[key] = value.strip()

    # =========================================================================
    # Stakeholders tab
    # CHANGE 4: The employee selector bar (Dept / Employee dropdowns) that was
    # placed above the input row has been removed.  The tab now shows only the
    # original Organisation, Role and Activities input fields.
    # =========================================================================
    def _create_stakeholders_tab(self):
        tab_sh = QWidget()
        layout_sh = QVBoxLayout(tab_sh)
        layout_sh.setContentsMargins(15, 15, 15, 15)

        # Input row: Organisation combo | Role combo | Activities text | Add btn
        input_layout_sh = QHBoxLayout()

        # Editable combo so the user can type a custom organisation name
        self.org_combo = QComboBox()
        self.org_combo.setEditable(True)
        self.org_combo.addItems(
            ["CEMILAC", "RCMA", "DGAQA", "DRDO", "User", "Production Agency"]
        )

        # Editable combo so the user can type a custom role
        self.role_combo = QComboBox()
        self.role_combo.setEditable(True)
        self.role_combo.addItems(
            [
                "Certification Authority",
                "Design Agency",
                "Production Agency",
                "Inspection Agency",
                "End User",
            ]
        )

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
        self.table_sh.setHorizontalHeaderLabels(
            [
                "Sl No.",
                "Organisation Agency",
                "Role",
                "Activities towards Certification",
            ]
        )
        self._configure_table(self.table_sh)
        # cellChanged fires when the user edits a cell directly in the table;
        # on_sh_cell_changed keeps the in-memory data list in sync.
        self.table_sh.cellChanged.connect(self.on_sh_cell_changed)

        layout_sh.addWidget(self.table_sh)
        layout_sh.addLayout(
            self._row_action_layout(self.delete_selected_sh, self.clear_all_sh)
        )

        self.tabs.addTab(tab_sh, "Stakeholders")

    # =========================================================================
    # Cert Task Allocation tab
    # CHANGE 4: The employee selector bar has been removed from this tab too.
    # The tab now shows only the original Activity, Work Centre and
    # Responsible Head input fields.
    # =========================================================================
    def _create_cert_task_tab(self):
        tab_ct = QWidget()
        layout_ct = QVBoxLayout(tab_ct)
        layout_ct.setContentsMargins(15, 15, 15, 15)

        # Input row: Activity text | Work Centre combo | Resp Head text | Add btn
        input_layout_ct = QHBoxLayout()

        self.act_input_ct = QLineEdit()
        self.act_input_ct.setPlaceholderText("e.g. Ground Testing")

        # Editable so the user can type a custom work centre name
        self.centre_combo = QComboBox()
        self.centre_combo.setEditable(True)
        self.centre_combo.addItems(
            ["CEMILAC", "RCMA", "DGAQA", "DRDO", "User", "Production Agency"]
        )

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
        self.table_ct.setHorizontalHeaderLabels(
            [
                "Sl No.",
                "Certification Activity",
                "Certification Work Centre",
                "Responsible Head",
            ]
        )
        self._configure_table(self.table_ct)
        self.table_ct.cellChanged.connect(self.on_ct_cell_changed)

        layout_ct.addWidget(self.table_ct)
        layout_ct.addLayout(
            self._row_action_layout(self.delete_selected_ct, self.clear_all_ct)
        )

        self.tabs.addTab(tab_ct, "Cert Task Allocation")

    # =========================================================================
    # Generic table tab builder  (Annexure-1 · Test Rigs · Aircraft Checks)
    #
    # CHANGE 6: The emp_fill_map=None optional parameter has been removed.
    # The previous version used it to inject the shared employee selector bar
    # above the input row and wire auto-fill callbacks.  Since the selector bar
    # was removed (CHANGE 4), the parameter served no purpose and was deleted.
    # The method now builds a purely text-based input row for every tab it
    # handles.
    # =========================================================================
    def _create_generic_table_tab(self, tab_name, data_attr, title):
        # Look up this tab's column definitions from the config dict
        columns = GENERIC_TAB_CONFIGS[data_attr]

        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)
        # Section heading label shown above the input row
        layout.addWidget(QLabel(title))

        # Build one QLineEdit per column, stored in generic_inputs[data_attr]
        # so that add_generic_row can read them all generically.
        input_layout = QHBoxLayout()
        self.generic_inputs[data_attr] = {}
        for label, key in columns:
            field = QLineEdit()
            field.setPlaceholderText(label)
            self.generic_inputs[data_attr][key] = field
            input_layout.addWidget(field, stretch=1)

        add_btn = QPushButton("Add Row")
        add_btn.setObjectName("PrimaryBtn")
        # Lambda captures data_attr by value so each tab's button calls the
        # correct add_generic_row(data_attr).
        add_btn.clicked.connect(lambda: self.add_generic_row(data_attr))
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)

        # +1 column for the auto-numbered "Sl No." column
        table = QTableWidget(0, len(columns) + 1)
        table.setHorizontalHeaderLabels(["Sl No."] + [label for label, _ in columns])
        self._configure_table(table)
        table.cellChanged.connect(
            lambda row, col: self.on_generic_cell_changed(data_attr, row, col)
        )
        # Store in dict so refresh_all_tables can iterate over all generic tabs
        self.generic_tables[data_attr] = table

        layout.addWidget(table)
        layout.addLayout(
            self._row_action_layout(
                lambda: self.delete_selected_generic(data_attr),
                lambda: self.clear_all_generic(data_attr),
            )
        )
        self.tabs.addTab(tab, tab_name)

    # =========================================================================
    # CHANGE 7 — NEW: Annexure-2 dedicated tab with employee dropdowns
    #
    # Why a dedicated method instead of using _create_generic_table_tab?
    # The generic method only supports QLineEdit fields in its input row.
    # Annexure-2 needs a QComboBox for "Name & Designation" (employee picker)
    # and a Dept filter combo to narrow the list — widgets that cannot be
    # expressed through the generic column-config system without over-complicating
    # it for all other tabs.
    #
    # Layout of the input row:
    #   [Dept: combo] [Name & Designation combo] [Cert Centre field]
    #   [E-mail field] [Phone field] [Add Row button]
    #
    # Data flow when the user picks an employee:
    #   1. User selects a dept from _a2_dept_combo  →  populate_names() fires
    #      and rebuilds _a2_name_combo with only employees in that dept.
    #   2. User selects a name from _a2_name_combo  →  on_name_selected() fires
    #      and auto-fills _a2_cert_field, _a2_email_field, _a2_phone_field from
    #      the employee row stored as QComboBox item userData.
    #   3. User may edit any field manually before clicking "Add Row".
    #   4. "Add Row" calls _add_annexure2_row() which appends a dict to
    #      self.parent_app.annexure2_data and refreshes the table.
    # =========================================================================
    def _create_annexure2_tab(self):
        # data_attr key must match the key in GENERIC_TAB_CONFIGS and the
        # attribute name on parent_app so existing helpers (refresh_generic_table,
        # on_generic_cell_changed, delete_selected_generic, clear_all_generic)
        # continue to work without modification.
        data_attr = "annexure2_data"

        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.addWidget(QLabel("Contact details of dealing officers and RDs"))

        # -----------------------------------------------------------------
        # Input row widgets
        # -----------------------------------------------------------------
        input_layout = QHBoxLayout()

        # --- Dept filter combo -------------------------------------------
        # Populated from _all_depts (unique Certification Centre values from
        # the Excel).  Selecting a dept narrows the name combo below.
        self._a2_dept_combo = QComboBox()
        self._a2_dept_combo.setMinimumWidth(110)
        self._a2_dept_combo.addItem("All Depts")  # show all when selected
        self._a2_dept_combo.addItems(self._all_depts)

        # --- Name & Designation combo ------------------------------------
        # Each item's display text  = emp[1]  ("Employee 1 - Executive")
        # Each item's userData      = the full employee row list, so that
        # on_name_selected can read all columns without a second lookup.
        # setEditable(True) lets the user type a name manually if needed.
        self._a2_name_combo = QComboBox()
        self._a2_name_combo.setEditable(True)
        self._a2_name_combo.setMinimumWidth(200)

        # --- Auto-filled text fields -------------------------------------
        # These are regular QLineEdits; they are filled automatically when
        # the user picks a name, but remain editable for manual correction.
        self._a2_cert_field = QLineEdit()
        self._a2_cert_field.setPlaceholderText("Certification Centre")

        self._a2_email_field = QLineEdit()
        self._a2_email_field.setPlaceholderText("E-mail ID")

        self._a2_phone_field = QLineEdit()
        self._a2_phone_field.setPlaceholderText("Phone Number")

        # -----------------------------------------------------------------
        # populate_names(dept_text)
        # Called whenever the dept combo value changes.
        # Rebuilds the name combo with employees from the selected dept
        # (or all employees when "All Depts" is chosen), then clears the
        # three auto-filled fields so stale data does not persist.
        # blockSignals(True/False) prevents on_name_selected from firing
        # while the combo is being rebuilt.
        # -----------------------------------------------------------------
        def populate_names(dept_text):
            self._a2_name_combo.blockSignals(True)
            self._a2_name_combo.clear()
            self._a2_name_combo.addItem("— select employee —")  # placeholder at index 0
            pool = (
                self._all_employees  # all 50 employees
                if dept_text == "All Depts"
                else self._emp_by_dept.get(dept_text, [])  # filtered subset
            )
            for emp in pool:
                # *** WHEN NEW EXCEL — STEP 5 of 6 ***
                #   Change emp[1] to whichever column index in your new Excel
                #   holds the employee's display name (shown in the dropdown).
                #   Example: if the name is in column 3 (0-based), use emp[3].
                self._a2_name_combo.addItem(
                    emp[1], userData=emp
                )  # <-- WHEN NEW EXCEL: change [1] to your name column index
            self._a2_name_combo.blockSignals(False)
            # Clear auto-filled fields to avoid showing data from a previous
            # dept selection alongside a new (not-yet-chosen) name.
            self._a2_cert_field.clear()
            self._a2_email_field.clear()
            self._a2_phone_field.clear()

        # -----------------------------------------------------------------
        # on_name_selected(index)
        # Called whenever the user picks an employee from the name combo.
        # Reads the userData (full employee row) attached to the selected
        # item and pushes each value into the corresponding QLineEdit.
        # Index 0 is the "— select employee —" placeholder; userData is None
        # for that item so the if-guard prevents a crash.
        # -----------------------------------------------------------------
        def on_name_selected(index):
            emp = self._a2_name_combo.itemData(index)
            if emp:
                # *** WHEN NEW EXCEL — STEP 6 of 6 ***
                #   These three lines read specific column indexes from the
                #   employee row and push them into the auto-filled fields.
                #   Update each index to match your new Excel column layout
                #   (reference the table you updated in STEP 2).
                #
                #   Current mapping:
                #     emp[2] → Certification Centre field  (dept column)
                #     emp[3] → E-mail ID field
                #     emp[4] → Phone Number field
                #
                #   Also update the placeholder text on the three QLineEdit
                #   fields above (_a2_cert_field, _a2_email_field, _a2_phone_field)
                #   and the table column headers in the QTableWidget below
                #   to match your new column names.
                self._a2_cert_field.setText(
                    emp[2]
                )  # <-- WHEN NEW EXCEL: change [2] to your dept/centre column index
                self._a2_email_field.setText(
                    emp[3]
                )  # <-- WHEN NEW EXCEL: change [3] to your email column index
                self._a2_phone_field.setText(
                    emp[4]
                )  # <-- WHEN NEW EXCEL: change [4] to your phone column index

        # Connect signals AFTER defining the slot functions
        self._a2_dept_combo.currentTextChanged.connect(populate_names)
        self._a2_name_combo.currentIndexChanged.connect(on_name_selected)
        # Prime the name combo with all employees on first load
        populate_names("All Depts")

        add_btn = QPushButton("Add Row")
        add_btn.setObjectName("PrimaryBtn")
        add_btn.clicked.connect(self._add_annexure2_row)

        # Assemble the input row in display order
        input_layout.addWidget(QLabel("Dept:"))
        input_layout.addWidget(self._a2_dept_combo)
        input_layout.addWidget(self._a2_name_combo, stretch=2)
        input_layout.addWidget(self._a2_cert_field, stretch=1)
        input_layout.addWidget(self._a2_email_field, stretch=2)
        input_layout.addWidget(self._a2_phone_field, stretch=1)
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)

        # -----------------------------------------------------------------
        # Table — 5 columns (Sl No. + 4 data columns)
        # cellChanged is wired to on_generic_cell_changed so direct cell
        # edits are persisted to parent_app.annexure2_data automatically,
        # just like every other generic tab.
        # -----------------------------------------------------------------
        table = QTableWidget(0, 5)
        table.setHorizontalHeaderLabels(
            [
                "Sl No.",
                "Name & Designation",
                "Certification Centre",
                "E-mail ID",
                "Phone Number",
            ]
        )
        self._configure_table(table)
        table.cellChanged.connect(
            lambda row, col: self.on_generic_cell_changed(data_attr, row, col)
        )
        # Register in generic_tables so refresh_all_tables picks it up
        self.generic_tables[data_attr] = table

        layout.addWidget(table)
        layout.addLayout(
            self._row_action_layout(
                lambda: self.delete_selected_generic(data_attr),
                lambda: self.clear_all_generic(data_attr),
            )
        )
        self.tabs.addTab(tab, "Annexure-2")

    # =========================================================================
    # CHANGE 8 — NEW: Add Row handler for Annexure-2
    #
    # Reads the current values from all Annexure-2 input widgets, appends a
    # dict to parent_app.annexure2_data (same schema as before), resets the
    # widgets to blank/placeholder state, and refreshes the table.
    #
    # The placeholder item "— select employee —" at index 0 is treated as
    # empty so that clicking Add Row without choosing a name does not insert
    # the placeholder text into the data.
    # =========================================================================
    def _add_annexure2_row(self):
        # currentText() works for both combo selection and manual typing
        name = self._a2_name_combo.currentText().strip()
        # Treat the placeholder as an empty selection
        if name == "— select employee —":
            name = ""
        cert = self._a2_cert_field.text().strip()
        email = self._a2_email_field.text().strip()
        phone = self._a2_phone_field.text().strip()

        # Only add a row if at least one field has content
        if any([name, cert, email, phone]):
            self.parent_app.annexure2_data.append(
                {
                    "name_designation": name,
                    "certification_centre": cert,
                    "email": email,
                    "phone": phone,
                }
            )
            # Reset all input widgets for the next entry
            self._a2_name_combo.setCurrentIndex(0)  # back to placeholder
            self._a2_cert_field.clear()
            self._a2_email_field.clear()
            self._a2_phone_field.clear()
            # Reuse the generic refresh helper — works because the table is
            # registered under "annexure2_data" in self.generic_tables and
            # GENERIC_TAB_CONFIGS["annexure2_data"] has the correct key mapping.
            self.refresh_generic_table("annexure2_data")

    # =========================================================================
    # Annexure-1 tab — dedicated method with RCMA + Dealing Officer dropdowns
    # Columns: Certifiable Item | Design Agency | Applicable Subpart |
    #          Designated Directorate/RCMA (combo) | Dealing Officer(s) (combo)
    # =========================================================================
    def _create_annexure1_tab(self):
        data_attr = "annexure1_data"
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.addWidget(QLabel("Work Assignment List of LRUs"))

        # ── Row 1: Certifiable Item | Design Agency | RCMA combo ──────────
        row1 = QHBoxLayout()

        self._a1_item_field = QLineEdit()
        self._a1_item_field.setPlaceholderText("Certifiable Item")
        self._a1_agency_field = QLineEdit()
        self._a1_agency_field.setPlaceholderText("Design Agency")

        # RCMA dropdown — filters all three Dealing Officer combos below
        self._a1_rcma_combo = QComboBox()
        self._a1_rcma_combo.setEditable(True)
        self._a1_rcma_combo.setMinimumWidth(130)
        self._a1_rcma_combo.addItem("— RCMA —")
        self._a1_rcma_combo.addItems(self._all_depts)

        row1.addWidget(self._a1_item_field, stretch=2)
        row1.addWidget(self._a1_agency_field, stretch=2)
        row1.addWidget(self._a1_rcma_combo, stretch=2)
        layout.addLayout(row1)

        # ── Row 2: HW / CH / SW officer combos + Add button ──────────────
        row2 = QHBoxLayout()

        self._a1_officer_hw_combo = QComboBox()
        self._a1_officer_hw_combo.setEditable(True)
        self._a1_officer_ch_combo = QComboBox()
        self._a1_officer_ch_combo.setEditable(True)
        self._a1_officer_sw_combo = QComboBox()
        self._a1_officer_sw_combo.setEditable(True)

        _OFFICER_COMBOS = [
            (self._a1_officer_hw_combo, "— DO (HW) —"),
            (self._a1_officer_ch_combo, "— DO (CH) —"),
            (self._a1_officer_sw_combo, "— DO (SW) —"),
        ]

        def _a1_populate_officers(dept_text):
            pool = (
                self._all_employees
                if dept_text in ("— RCMA —", "All Depts")
                else self._emp_by_dept.get(dept_text, [])
            )
            for combo, placeholder in _OFFICER_COMBOS:
                combo.blockSignals(True)
                combo.clear()
                combo.addItem(placeholder)
                for emp in pool:
                    combo.addItem(emp[1], userData=emp)
                combo.blockSignals(False)

        self._a1_rcma_combo.currentTextChanged.connect(_a1_populate_officers)
        _a1_populate_officers("— RCMA —")

        add_btn = QPushButton("Add Row")
        add_btn.setObjectName("PrimaryBtn")
        add_btn.clicked.connect(self._add_annexure1_row)

        row2.addWidget(QLabel("HW:"))
        row2.addWidget(self._a1_officer_hw_combo, stretch=2)
        row2.addWidget(QLabel("CH:"))
        row2.addWidget(self._a1_officer_ch_combo, stretch=2)
        row2.addWidget(QLabel("SW:"))
        row2.addWidget(self._a1_officer_sw_combo, stretch=2)
        row2.addWidget(add_btn)
        layout.addLayout(row2)

        # ── Table: 7 columns (Sl No. + 6 data) ───────────────────────────
        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels(
            [
                "Sl No.",
                "Certifiable Item",
                "Design Agency",
                "Designated Directorate/RCMA",
                "Dealing Officer HW",
                "Dealing Officer CH",
                "Dealing Officer SW",
            ]
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
        self.tabs.addTab(tab, "Annexure-1")

    def _add_annexure1_row(self):
        item = self._a1_item_field.text().strip()
        agency = self._a1_agency_field.text().strip()
        rcma = self._a1_rcma_combo.currentText().strip()

        def _officer(combo, placeholder):
            v = combo.currentText().strip()
            return "" if v == placeholder else v

        officer_hw = _officer(self._a1_officer_hw_combo, "— DO (HW) —")
        officer_ch = _officer(self._a1_officer_ch_combo, "— DO (CH) —")
        officer_sw = _officer(self._a1_officer_sw_combo, "— DO (SW) —")

        if rcma == "— RCMA —":
            rcma = ""

        if any([item, agency, rcma, officer_hw, officer_ch, officer_sw]):
            self.parent_app.annexure1_data.append(
                {
                    "certifiable_item": item,
                    "design_agency": agency,
                    "designated_directorate": rcma,
                    "dealing_officer_hw": officer_hw,
                    "dealing_officer_ch": officer_ch,
                    "dealing_officer_sw": officer_sw,
                }
            )
            self._a1_item_field.clear()
            self._a1_agency_field.clear()
            self._a1_rcma_combo.setCurrentIndex(0)
            for combo in [
                self._a1_officer_hw_combo,
                self._a1_officer_ch_combo,
                self._a1_officer_sw_combo,
            ]:
                combo.setCurrentIndex(0)
            self.refresh_generic_table("annexure1_data")

    # =========================================================================
    # Test Rigs tab — dedicated method with RCMA + Dealing Officer dropdowns
    # Columns: Test Rig Name | Utilization Phase | Design Agency |
    #          Designated RCMA (combo) | Dealing Officer (combo)
    # =========================================================================
    def _create_test_rigs_tab(self):
        data_attr = "test_rigs_data"
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.addWidget(QLabel("Test Rigs, Simulators and Ground Equipment Required"))

        input_layout = QHBoxLayout()

        self._tr_name_field = QLineEdit()
        self._tr_name_field.setPlaceholderText("Test Rig / Ground Equipment Name")
        self._tr_phase_field = QLineEdit()
        self._tr_phase_field.setPlaceholderText("Utilizations Phase")
        self._tr_agency_field = QLineEdit()
        self._tr_agency_field.setPlaceholderText("Design Agency")

        self._tr_rcma_combo = QComboBox()
        self._tr_rcma_combo.setEditable(True)
        self._tr_rcma_combo.setMinimumWidth(130)
        self._tr_rcma_combo.addItem("\u2014 RCMA \u2014")
        self._tr_rcma_combo.addItems(self._all_depts)

        self._tr_officer_combo = QComboBox()
        self._tr_officer_combo.setEditable(True)
        self._tr_officer_combo.setMinimumWidth(180)

        def _tr_populate_officers(dept_text):
            self._tr_officer_combo.blockSignals(True)
            self._tr_officer_combo.clear()
            self._tr_officer_combo.addItem("\u2014 Dealing Officer \u2014")
            pool = (
                self._all_employees
                if dept_text in ("\u2014 RCMA \u2014", "All Depts")
                else self._emp_by_dept.get(dept_text, [])
            )
            for emp in pool:
                self._tr_officer_combo.addItem(emp[1], userData=emp)
            self._tr_officer_combo.blockSignals(False)

        self._tr_rcma_combo.currentTextChanged.connect(_tr_populate_officers)
        _tr_populate_officers("\u2014 RCMA \u2014")

        add_btn = QPushButton("Add Row")
        add_btn.setObjectName("PrimaryBtn")
        add_btn.clicked.connect(self._add_test_rigs_row)

        input_layout.addWidget(self._tr_name_field, stretch=2)
        input_layout.addWidget(self._tr_phase_field, stretch=2)
        input_layout.addWidget(self._tr_agency_field, stretch=2)
        input_layout.addWidget(self._tr_rcma_combo, stretch=2)
        input_layout.addWidget(self._tr_officer_combo, stretch=2)
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)

        table = QTableWidget(0, 6)
        table.setHorizontalHeaderLabels(
            [
                "Sl No.",
                "Test Rig / Ground Equipment Name",
                "Utilizations Phase",
                "Design Agency",
                "Designated RCMA",
                "Dealing Officer",
            ]
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
        self.tabs.addTab(tab, "Test Rigs")

    def _add_test_rigs_row(self):
        name = self._tr_name_field.text().strip()
        phase = self._tr_phase_field.text().strip()
        agency = self._tr_agency_field.text().strip()
        rcma = self._tr_rcma_combo.currentText().strip()
        officer = self._tr_officer_combo.currentText().strip()
        if rcma == "\u2014 RCMA \u2014":
            rcma = ""
        if officer == "\u2014 Dealing Officer \u2014":
            officer = ""
        if any([name, phase, agency, rcma, officer]):
            self.parent_app.test_rigs_data.append(
                {
                    "test_rig_name": name,
                    "utilization_phase": phase,
                    "design_agency": agency,
                    "designated_rcma": rcma,
                    "dealing_officer": officer,
                }
            )
            self._tr_name_field.clear()
            self._tr_phase_field.clear()
            self._tr_agency_field.clear()
            self._tr_rcma_combo.setCurrentIndex(0)
            self._tr_officer_combo.setCurrentIndex(0)
            self.refresh_generic_table("test_rigs_data")

    # =========================================================================
    # Aircraft Checks tab — dedicated method with RCMA + Dealing Officer dropdowns
    # Columns: System Name | Design Agency |
    #          Designated RCMA (combo) | Dealing Officer (combo)
    # =========================================================================
    def _create_aircraft_checks_tab(self):
        data_attr = "aircraft_checks_data"
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.addWidget(QLabel("Aircraft Integration Checks and Flight Clearance"))

        input_layout = QHBoxLayout()

        self._ac_system_field = QLineEdit()
        self._ac_system_field.setPlaceholderText("System Name")
        self._ac_agency_field = QLineEdit()
        self._ac_agency_field.setPlaceholderText("Design Agency")

        self._ac_rcma_combo = QComboBox()
        self._ac_rcma_combo.setEditable(True)
        self._ac_rcma_combo.setMinimumWidth(130)
        self._ac_rcma_combo.addItem("\u2014 RCMA \u2014")
        self._ac_rcma_combo.addItems(self._all_depts)

        self._ac_officer_combo = QComboBox()
        self._ac_officer_combo.setEditable(True)
        self._ac_officer_combo.setMinimumWidth(180)

        def _ac_populate_officers(dept_text):
            self._ac_officer_combo.blockSignals(True)
            self._ac_officer_combo.clear()
            self._ac_officer_combo.addItem("\u2014 Dealing Officer \u2014")
            pool = (
                self._all_employees
                if dept_text in ("\u2014 RCMA \u2014", "All Depts")
                else self._emp_by_dept.get(dept_text, [])
            )
            for emp in pool:
                self._ac_officer_combo.addItem(emp[1], userData=emp)
            self._ac_officer_combo.blockSignals(False)

        self._ac_rcma_combo.currentTextChanged.connect(_ac_populate_officers)
        _ac_populate_officers("\u2014 RCMA \u2014")

        add_btn = QPushButton("Add Row")
        add_btn.setObjectName("PrimaryBtn")
        add_btn.clicked.connect(self._add_aircraft_checks_row)

        input_layout.addWidget(self._ac_system_field, stretch=2)
        input_layout.addWidget(self._ac_agency_field, stretch=2)
        input_layout.addWidget(self._ac_rcma_combo, stretch=2)
        input_layout.addWidget(self._ac_officer_combo, stretch=2)
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)

        table = QTableWidget(0, 5)
        table.setHorizontalHeaderLabels(
            [
                "Sl No.",
                "System Name",
                "Design Agency",
                "Designated RCMA",
                "Dealing Officer",
            ]
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
        self.tabs.addTab(tab, "Aircraft Checks")

    def _add_aircraft_checks_row(self):
        system = self._ac_system_field.text().strip()
        agency = self._ac_agency_field.text().strip()
        rcma = self._ac_rcma_combo.currentText().strip()
        officer = self._ac_officer_combo.currentText().strip()
        if rcma == "\u2014 RCMA \u2014":
            rcma = ""
        if officer == "\u2014 Dealing Officer \u2014":
            officer = ""
        if any([system, agency, rcma, officer]):
            self.parent_app.aircraft_checks_data.append(
                {
                    "system_name": system,
                    "design_agency": agency,
                    "designated_rcma": rcma,
                    "dealing_officer": officer,
                }
            )
            self._ac_system_field.clear()
            self._ac_agency_field.clear()
            self._ac_rcma_combo.setCurrentIndex(0)
            self._ac_officer_combo.setCurrentIndex(0)
            self.refresh_generic_table("aircraft_checks_data")

    # =========================================================================
    # Shared table configuration helpers — unchanged from original
    # =========================================================================

    def _configure_table(self, table):
        # Stretch all columns evenly to fill the available width
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # Alternating row colours improve readability
        table.setAlternatingRowColors(True)
        # Click anywhere in a row to select the whole row (not just one cell)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)

    def _row_action_layout(self, delete_handler, clear_handler):
        """Return an QHBoxLayout with Delete Selected Row and Clear All buttons."""
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
        """
        Called every time the Tables dialog is opened (show_tables_popup).
        Re-populates every widget from the current in-memory data so that
        any changes made elsewhere in the app are reflected immediately.
        """
        # Restore Issue Details fields
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
        # Iterates over all registered generic tables including Annexure-2
        # (registered in self.generic_tables by _create_annexure2_tab)
        for data_attr in self.generic_tables:
            self.refresh_generic_table(data_attr)

    # =========================================================================
    # Stakeholders helpers — unchanged from original
    # =========================================================================

    def add_stakeholder(self):
        org = self.org_combo.currentText().strip()
        role = self.role_combo.currentText().strip()
        act = self.act_input.text().strip()

        if org:  # require at least the organisation to be filled
            self.parent_app.stakeholders_data.append(
                {"org": org, "role": role, "activities": act}
            )
            self.role_combo.setCurrentText("")
            self.act_input.clear()
            self.refresh_sh_table()

    def on_sh_cell_changed(self, row, column):
        """Keep in-memory stakeholders_data in sync when a cell is edited."""
        if 0 <= row < len(self.parent_app.stakeholders_data):
            item = self.table_sh.item(row, column)
            if item:
                value = item.text().strip()
                if column == 1:
                    self.parent_app.stakeholders_data[row]["org"] = value
                elif column == 2:
                    self.parent_app.stakeholders_data[row]["role"] = value
                elif column == 3:
                    self.parent_app.stakeholders_data[row]["activities"] = value

    def delete_selected_sh(self):
        row = self.table_sh.currentRow()
        if row >= 0:
            self.parent_app.stakeholders_data.pop(row)
            self.refresh_sh_table()

    def clear_all_sh(self):
        self.parent_app.stakeholders_data.clear()
        self.refresh_sh_table()

    def refresh_sh_table(self):
        # blockSignals prevents on_sh_cell_changed from firing while the table
        # is being rebuilt programmatically, which would cause incorrect updates.
        self.table_sh.blockSignals(True)
        self.table_sh.setRowCount(0)
        for i, data in enumerate(self.parent_app.stakeholders_data):
            self.table_sh.insertRow(i)
            self._set_sl_item(self.table_sh, i)
            self.table_sh.setItem(i, 1, QTableWidgetItem(data.get("org", "")))
            self.table_sh.setItem(i, 2, QTableWidgetItem(data.get("role", "")))
            self.table_sh.setItem(i, 3, QTableWidgetItem(data.get("activities", "")))
        self.table_sh.blockSignals(False)

    # =========================================================================
    # Cert Task helpers — unchanged from original
    # =========================================================================

    def add_task(self):
        act = self.act_input_ct.text().strip()
        centre = self.centre_combo.currentText().strip()
        head = self.head_input.text().strip()

        if act or centre or head:  # require at least one field
            self.parent_app.cert_task_data.append(
                {"activity": act, "centre": centre, "head": head}
            )
            self.act_input_ct.clear()
            self.head_input.clear()
            self.refresh_ct_table()

    def on_ct_cell_changed(self, row, column):
        """Keep cert_task_data in sync when a cell is edited directly."""
        if 0 <= row < len(self.parent_app.cert_task_data):
            item = self.table_ct.item(row, column)
            if item:
                value = item.text().strip()
                if column == 1:
                    self.parent_app.cert_task_data[row]["activity"] = value
                elif column == 2:
                    self.parent_app.cert_task_data[row]["centre"] = value
                elif column == 3:
                    self.parent_app.cert_task_data[row]["head"] = value

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
            self.table_ct.setItem(i, 1, QTableWidgetItem(data.get("activity", "")))
            self.table_ct.setItem(i, 2, QTableWidgetItem(data.get("centre", "")))
            self.table_ct.setItem(i, 3, QTableWidgetItem(data.get("head", "")))
        self.table_ct.blockSignals(False)

    # =========================================================================
    # Generic table helpers
    # Used by Annexure-1, Test Rigs, Aircraft Checks for Add / edit / delete.
    # Also re-used by Annexure-2 for cell-edit sync and table refresh, because
    # _create_annexure2_tab registers its table under "annexure2_data" in
    # self.generic_tables and GENERIC_TAB_CONFIGS still contains that key.
    # =========================================================================

    def add_generic_row(self, data_attr):
        """Collect values from all QLineEdit inputs for data_attr and append."""
        values = {
            key: field.text().strip()
            for key, field in self.generic_inputs[data_attr].items()
        }
        if any(values.values()):  # skip if every field is empty
            getattr(self.parent_app, data_attr).append(values)
            for field in self.generic_inputs[data_attr].values():
                field.clear()
            self.refresh_generic_table(data_attr)

    def on_generic_cell_changed(self, data_attr, row, column):
        """
        Keep the in-memory list for data_attr in sync when the user edits a
        cell directly in the table.  Column 0 is always "Sl No." (read-only)
        so it is skipped.  The correct dict key is looked up from
        GENERIC_TAB_CONFIGS using the column index.
        """
        rows = getattr(self.parent_app, data_attr)
        if not 0 <= row < len(rows) or column == 0:
            return
        item = self.generic_tables[data_attr].item(row, column)
        if item:
            # column - 1 because GENERIC_TAB_CONFIGS lists do not include Sl No.
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
        """
        Rebuild the QTableWidget for data_attr from scratch using the current
        in-memory list.  blockSignals prevents on_generic_cell_changed from
        firing and corrupting data while rows are being inserted.
        """
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
        """
        Insert a read-only serial-number cell (1-based) into column 0.
        ItemIsEditable flag is cleared so the user cannot modify it.
        """
        sl = QTableWidgetItem(str(row + 1))
        sl.setFlags(sl.flags() & ~Qt.ItemFlag.ItemIsEditable)
        table.setItem(row, 0, sl)
