import os
import shutil
import subprocess
import sys
import tempfile

# Path to the DRDO logo image stored in the project root.
# Resolved relative to this file (core/export.py) so it works from any
# working directory.
_LOGO_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "drdo_logo.png",
)

import pythoncom
import win32com.client
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PyQt6.QtGui import QTextDocument
from PyQt6.QtPrintSupport import QPrinter


class PDFExporter:
    @staticmethod
    def export(intro_text, manual_notes, scope_text, stakeholders, cert_tasks):
        temp_file = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        temp_path = temp_file.name
        temp_file.close()

        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName(temp_path)

        ai_html = intro_text.replace("\n", "<br>")

        sth_html = PDFExporter._build_stakeholders_table(stakeholders)
        ct_html = PDFExporter._build_cert_task_table(cert_tasks)

        notes_html_blocks = []
        # Change 1: Support manual notes as either a dict or a plain string.
        # If a string is passed, convert newline characters to HTML breaks.
        if isinstance(manual_notes, dict):
            for heading, content in manual_notes.items():
                if content:
                    block = str(content)
                    if not block.strip().startswith("<"):
                        block = block.replace("\n", "<br>")
                    notes_html_blocks.append(f"<h3>{heading}</h3>{block}")
        else:
            notes_html = str(manual_notes)
            notes_html_blocks.append(notes_html.replace("\n", "<br>"))
        notes_section = "".join(notes_html_blocks)

        final_html = f"""
            <h1>Document Analysis Report</h1><hr>
            <h2>1. Introduction Summary</h2><p>{ai_html}</p><br>
            <h2>2. Manual Notes</h2>{notes_section}<br>
            <h2>Scope</h2><p>{scope_text}</p><br>
            <h2>Stakeholders</h2>{sth_html}<br>
            <h2>Certification Task Allocation</h2>{ct_html}
        """

        doc = QTextDocument()
        doc.setHtml(final_html)
        doc.print(printer)

        return temp_path

    @staticmethod
    def _build_stakeholders_table(stakeholders):
        html = """<table border='1' cellspacing='0' cellpadding='5' width='100%'>
                  <tr><th>Sl No.</th><th>Organisation</th><th>Role</th><th>Activities</th></tr>"""
        if stakeholders:
            for i, sh in enumerate(stakeholders):
                html += f"""<tr><td>{i + 1}</td><td>{sh.get("org", "")}</td>
                           <td>{sh.get("role", "")}</td><td>{sh.get("activities", "")}</td></tr>"""
        else:
            html += "<tr><td>&nbsp;</td><td></td><td></td><td></td></tr>"
        html += "</table>"
        return html

    @staticmethod
    def _build_cert_task_table(cert_tasks):
        html = """<table border='1' cellspacing='0' cellpadding='5' width='100%'>
                  <tr><th>Sl No.</th><th>Activity</th><th>Work Centre</th><th>Head</th></tr>"""
        if cert_tasks:
            for i, ct in enumerate(cert_tasks):
                html += f"""<tr><td>{i + 1}</td><td>{ct.get("activity", "")}</td>
                           <td>{ct.get("centre", "")}</td><td>{ct.get("head", "")}</td></tr>"""
        else:
            html += "<tr><td>&nbsp;</td><td></td><td></td><td></td></tr>"
        html += "</table>"
        return html


class DocxExporter:
    @staticmethod
    def export(save_path, **kwargs):
        """
        Export document to DOCX format.

        Args:
            save_path: Path to save the DOCX file
            **kwargs: Any attributes to include in the document:
                - project_name: Name of the project
                - intro_text: Introduction text
                - manual_notes: Manual notes (dict or string)
                - scope_items: Scope items (list)
                - stakeholders: Stakeholders data (list)
                - cert_tasks: Certification tasks (list)
                - combos: Combo selections (dict)
                - annexure3_image_path: Path to annexure 3 image
                - table_data: Table data (dict) OR individual kwargs:
                    - issue_details: Issue details (dict)
                    - annexure1_data: Annexure 1 data (list)
                    - annexure2_data: Annexure 2 data (list)
                    - test_rigs_data: Test rigs data (list)
                    - aircraft_checks_data: Aircraft checks data (list)
                - Any other custom attributes
        """
        # Change 2: Modified to accept **kwargs instead of fixed parameters for flexibility
        # Extract known parameters with defaults
        project_name = kwargs.get("project_name", "")
        intro_text = kwargs.get("intro_text", "")
        manual_notes = kwargs.get("manual_notes", {})
        if isinstance(manual_notes, str):
            # CHANGED: Allow string manual notes as a single fallback field
            manual_notes = {"Other": manual_notes}
        scope_items = kwargs.get("scope_items", [])
        stakeholders = kwargs.get("stakeholders", [])
        cert_tasks = kwargs.get("cert_tasks", [])
        combos = kwargs.get("combos", {})
        annexure3_image_path = kwargs.get("annexure3_image_path", None)

        # Change 3: Handle table_data - can be passed directly or built from individual kwargs
        # This allows flexibility in how data is passed to the exporter
        table_data = kwargs.get("table_data", {})
        if not table_data:
            # Change 4: Build table_data from individual data kwargs for better API
            # Maps individual parameters to the internal table_data structure
            table_data = {
                "issue_details": kwargs.get("issue_details", {}),
                "annexure1": kwargs.get("annexure1_data", []),
                "annexure2": kwargs.get("annexure2_data", []),
                "test_rigs": kwargs.get("test_rigs_data", []),
                "aircraft_checks": kwargs.get("aircraft_checks_data", []),
            }

        # Change 5: Define known parameters for exclusion from custom_attributes
        # Allows any additional custom kwargs to be passed through
        known_params = {
            "project_name",
            "intro_text",
            "manual_notes",
            "scope_items",
            "stakeholders",
            "cert_tasks",
            "combos",
            "annexure3_image_path",
            "table_data",
            "issue_details",
            "annexure1_data",
            "annexure2_data",
            "test_rigs_data",
            "aircraft_checks_data",
        }

        # Change 6: Store any extra custom attributes for extensibility
        # Custom attributes can be added without modifying the function signature
        custom_attributes = {k: v for k, v in kwargs.items() if k not in known_params}

        table_data = table_data or {}
        doc = Document()
        style = doc.styles["Normal"]
        style.font.name = "Calibri"
        style.font.size = Pt(11)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.08

        # Title page
        DocxExporter._add_title_page(doc, project_name, table_data.get("issue_details"))

        # Content pages
        section2 = doc.add_section(WD_SECTION.NEW_PAGE)
        section2.top_margin = Inches(0.75)
        section2.bottom_margin = Inches(0.75)

        DocxExporter._add_content_sections(
            doc,
            intro_text,
            manual_notes,
            scope_items,
            stakeholders,
            cert_tasks,
            combos,
            annexure3_image_path,
            table_data,
            custom_attributes,
            project_name=project_name,
        )
        DocxExporter._apply_document_spacing(doc)

        doc.save(save_path)

    @staticmethod
    def create_temp_docx(**kwargs):
        """Build the same DOCX content as the final export, but into a temp file.

        Use this for Preview so the user does not need to choose/download
        the final Word document first. The caller owns the returned path.
        """
        temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        temp_path = temp_file.name
        temp_file.close()

        # CHANGED: Forward **kwargs instead of individual parameters
        # Maintains flexibility in how data is passed to export()
        DocxExporter.export(temp_path, **kwargs)
        return temp_path

    @staticmethod
    def create_preview_pdf(**kwargs):
        """Create a PDF preview from a temporary DOCX and return the PDF path.

        This intentionally does not write the user's final .docx file. It lets
        the Preview button show the generated document before final download.
        """
        # CHANGED: Forward **kwargs to create_temp_docx for consistency
        # This keeps the preview path in sync with the current export API.
        temp_docx_path = DocxExporter.create_temp_docx(**kwargs)
        temp_pdf_file = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        temp_pdf_path = temp_pdf_file.name
        temp_pdf_file.close()

        try:
            return DocxExporter.convert_to_pdf(temp_docx_path, temp_pdf_path)
        finally:
            try:
                os.remove(temp_docx_path)
            except OSError:
                pass

    @staticmethod
    def convert_to_pdf(docx_path, pdf_path=None):
        docx_path = os.path.abspath(docx_path)
        if not os.path.exists(docx_path):
            raise FileNotFoundError(docx_path)

        if pdf_path is None:
            temp_file = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            pdf_path = temp_file.name
            temp_file.close()
        pdf_path = os.path.abspath(pdf_path)

        errors = []
        if sys.platform.startswith("win"):
            try:
                DocxExporter._convert_with_word(docx_path, pdf_path)
                if os.path.exists(pdf_path):
                    return pdf_path
            except Exception as e:
                errors.append(f"Microsoft Word: {e}")

        try:
            DocxExporter._convert_with_libreoffice(docx_path, pdf_path)
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception as e:
            errors.append(f"LibreOffice: {e}")

        raise RuntimeError(
            "Could not convert the generated Word document to PDF. "
            + " | ".join(errors)
        )

    @staticmethod
    def _convert_with_word(docx_path, pdf_path):
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)
        finally:
            if doc is not None:
                doc.Close(False)
            if word is not None:
                word.Quit()
            pythoncom.CoUninitialize()

    @staticmethod
    def _convert_with_libreoffice(docx_path, pdf_path):
        candidates = [
            shutil.which("soffice"),
            shutil.which("libreoffice"),
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        office = next(
            (path for path in candidates if path and os.path.exists(path)), None
        )
        if not office:
            raise RuntimeError("LibreOffice executable was not found")

        output_dir = os.path.dirname(pdf_path)
        completed = subprocess.run(
            [
                office,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                output_dir,
                docx_path,
            ],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError((completed.stderr or completed.stdout).strip())

        converted_path = os.path.join(
            output_dir,
            os.path.splitext(os.path.basename(docx_path))[0] + ".pdf",
        )
        if not os.path.exists(converted_path):
            raise RuntimeError("PDF output was not created")
        if os.path.abspath(converted_path) != os.path.abspath(pdf_path):
            os.replace(converted_path, pdf_path)

    @staticmethod
    def _add_title_page(doc, project_name, issue_details=None):
        issue_details = issue_details or {}
        section1 = doc.sections[0]
        section1.top_margin = Inches(1.5)
        section1.bottom_margin = Inches(1.5)

        # Title page: show the auto-generated PBS file number.
        # The Reference section (in content pages) shows the extracted document ref number.
        pbs_no = (
            issue_details.get("pbs_no")
            or issue_details.get("file_no")
            or "________________________"
        )
        doc.add_paragraph(f"File No. {pbs_no}")
        p = doc.add_paragraph()
        run = p.add_run("\nTASK DIRECTIVE\n")
        run.bold = True
        run.font.size = Pt(20)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph("")  # spacer

        # Issue No. and Date of Issue row — bold labels, extracted values
        table = doc.add_table(rows=1, cols=2)
        issue_no = issue_details.get("issue_no") or "________________"
        date_of_issue = issue_details.get("date_of_issue") or "________________"

        for cell, label, value in [
            (table.cell(0, 0), "Issue No.: ", issue_no),
            (table.cell(0, 1), "Date of Issue: ", date_of_issue),
        ]:
            cell.paragraphs[0].clear()
            run_label = cell.paragraphs[0].add_run(label)
            run_label.bold = True
            run_label.font.size = Pt(11)
            run_value = cell.paragraphs[0].add_run(value)
            run_value.font.size = Pt(11)

        if not project_name or project_name.lower() == "not found":
            project_name = "__________________________________________"
        doc.add_paragraph(f"\nPROJECT NAME: {project_name}")

        # Insert DRDO logo if the file exists; fall back to a blank spacer.
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(_LOGO_PATH):
            try:
                run = p.add_run()
                run.add_picture(_LOGO_PATH, width=Inches(2.2))
                for _ in range(6):
                    doc.add_paragraph("")
            except Exception:
                p.add_run("[DRDO Logo]")
                for _ in range(8):
                    doc.add_paragraph("")
        else:
            p.add_run("LOGO HERE")
            for _ in range(12):
                doc.add_paragraph("")

        p = doc.add_paragraph(
            "CENTRE FOR MILITARY AIRWORTHINESS AND CERTIFICATION  \n (CEMILAC)"
        )
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph("DRDO, MoD")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph("Marathahalli Colony Post, \n Bengaluru - 560037")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    @staticmethod
    def _add_roman_list(doc, items):
        romans = [
            "i",
            "ii",
            "iii",
            "iv",
            "v",
            "vi",
            "vii",
            "viii",
            "ix",
            "x",
            "xi",
            "xii",
            "xiii",
            "xiv",
            "xv",
            "xvi",
            "xvii",
            "xviii",
            "xix",
            "xx",
        ]

        for index, item in enumerate(items):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.35)
            p.paragraph_format.first_line_indent = Inches(-0.35)
            p.paragraph_format.space_after = Pt(3)

            number = romans[index] if index < len(romans) else str(index + 1)
            p.add_run(f"{number}. ")
            p.add_run(item)

    @staticmethod
    def _add_content_sections(
        doc,
        intro_text,
        manual_notes,
        scope_items,
        stakeholders,
        cert_tasks,
        combos,
        annexure3_image_path=None,
        table_data=None,
        custom_attributes=None,
        project_name="",
    ):
        table_data = table_data or {}
        custom_attributes = custom_attributes or {}
        if not isinstance(manual_notes, dict):
            manual_notes = {"Other": str(manual_notes)}

        issue_details = table_data.get("issue_details", {})

        # ── Project name (appears before Introduction section) ───────────────────────
        if project_name and project_name.lower() != "not found":
            p = doc.add_paragraph()
            run = p.add_run(f"Project: {project_name}")
            run.bold = True
            run.font.size = Pt(13)

        # ── 1. Introduction ──────────────────────────────────────────────
        DocxExporter._add_heading(doc, "1. Introduction")
        # Use the Introduction manual note (populated from raw document extraction).
        # Fall back to intro_text only if no manual note was written.
        intro_content = (
            manual_notes.get("Introduction", "").strip()
            or (intro_text.strip() if intro_text else "")
            or " "
        )
        doc.add_paragraph(intro_content)

        # ── 2. Reference ─────────────────────────────────────────────────
        DocxExporter._add_heading(doc, "\n2. Reference")
        # Reference section: use the document-extracted reference number and date.
        # These are populated from the uploaded PDF's native text on upload.
        # Fallback to blank lines if extraction did not find them.
        file_no = issue_details.get("file_no") or "________________________"
        date_issue = issue_details.get("date_of_issue") or "________________________"
        ref_default = (
            f"1. Document Name      : {project_name or '________________________'}\n"
            f"2. Document Ref No    : {file_no}\n"
            f"3. Document Issue Date: {date_issue}"
        )
        doc.add_paragraph(
            DocxExporter._note(manual_notes, "Reference", default=ref_default)
        )

        # ── 3. Basis of Task Directive ───────────────────────────────────────
        DocxExporter._add_heading(doc, "\n3. Basis Of Task Directive")
        doc.add_paragraph(
            DocxExporter._note(
                manual_notes,
                "Basis Of Task Directive",
                default=(
                    " 1. IMAP-23, Indian Military Airworthiness Procedure - 23\n"
                    " 2. IMTAR-21, Version 2.0 Indian Military Technical Airworthiness Requirements\n"
                    " 3. Applicable Airworthiness Directives and CEMILAC Directives"
                ),
            )
        )

        # ── 4. Scope of Task Directive ───────────────────────────────────────
        scope_text = DocxExporter._numbered_text(scope_items, default=" ")
        DocxExporter._add_heading(doc, "\n4. Scope Of Task Directive")
        doc.add_paragraph(
            "The scope of the Task Directive is to assign the certification "
            "responsibilities to CEMILAC and RCMAs to enable smooth communication "
            "and transactions between the design agencies and the certification agency. "
            "The designated RCMAs provide concurrent certification coverage for:"
        )
        doc.add_paragraph(scope_text)
        scope_note = DocxExporter._note(
            manual_notes, "Scope", "Scope Of Task Directive", default=""
        )
        if scope_note:
            doc.add_paragraph(scope_note)

        # ── 5. Stakeholders ────────────────────────────────────────────────
        DocxExporter._add_heading(doc, "\n5. Stakeholders")
        doc.add_paragraph(
            "The following are the major stakeholders from certification perspective "
            "for the development, integration of the system on the various Rigs."
        )
        DocxExporter._add_stakeholders_table(doc, stakeholders)
        DocxExporter._add_manual_note(doc, manual_notes, "Stakeholders")

        # ── 6. Certification Work ──────────────────────────────────────────
        DocxExporter._add_heading(doc, "\n6. Certification Work Breakdown")
        cert_items = DocxExporter._combo_items(combos, "Certification Work Breakdown")
        doc.add_paragraph("The following are the major certification activities:")
        doc.add_paragraph(
            DocxExporter._numbered_text(
                cert_items, default="______________________________________________"
            )
        )
        DocxExporter._add_manual_note(doc, manual_notes, "Certification Work Breakdown")

        # ── 7. Task Allocation by Coordinating Directorate ───────────────────────
        DocxExporter._add_heading(
            doc, "\n7. Task Allocation by Coordinating Directorate"
        )

        DocxExporter._add_heading(doc, "7.1 Coordinating Directorate")
        doc.add_paragraph(
            DocxExporter._note(
                manual_notes,
                "Coordinating Directorate",
                default=(
                    "The overall certification of the project would be managed by:\n"
                    "Director () / RCMA ()\n\n"
                    "The responsibilities of the Coordinating Directorate would be as follows:\n\n"
                    " i.   Plan, review and monitor the progress of the certification.\n\n"
                    " ii.  Participate in System level and software meetings.\n\n"
                    " iii. Propose Constitution of SCRB and TARB as and when required."
                ),
            )
        )

        DocxExporter._add_heading(doc, "\n7.2 Single Point of Contact (SPoC)")
        spoc_note = DocxExporter._manual_note(
            manual_notes,
            "Single Point of Contact( SPoC)",
            "Single Point of Contact",
            "SPoC",
        )
        if spoc_note:
            doc.add_paragraph(spoc_note)
        else:
            doc.add_paragraph(
                "The following officer would be responsible for overall certification "
                "coordination with all the stakeholders for smooth certification of the project:"
            )
            doc.add_paragraph("Sri ..., Sc-...")
            doc.add_paragraph("RCMA ()/ Dte ()")
            doc.add_paragraph("The responsibilities of the SPoC would be as follows:")
            DocxExporter._add_roman_list(
                doc,
                [
                    "Ensure prompt assignment of Project Leader/ Dealing Officer by all RCMAs.",
                    "Ensure correct uplinking and downlinking of each project registered by the Design Agency, in line with the PBS.",
                    "Work in coordination with the D&D agency and ensure timely execution of certification activities.",
                    "Liaise with concerned RCMAs, CEMILAC directorates and other stakeholders.",
                    "Ensure finalisation of TAB and ACP after taking inputs from all stakeholders.",
                    "Keep a repository on incoming and outgoing artefacts/documents and ensure timely dissemination of information to all certification centres/stakeholders.",
                    "Apprise Coordinating Directorate the progression of certification activities.",
                    "Plan and coordinate Certification Review Meetings.",
                    "Member Secretary for SCRB and TARB as and when required.",
                    "Coordination for installation/integration/flight clearance.",
                ],
            )

        DocxExporter._add_heading(doc, "7.3 Certification Task Allocation")
        doc.add_paragraph(
            "The detailed certification task allocation is given below. The detailed list "
            "of allocation of certification responsibility to RCMAs and dealing officers, "
            "in respect of each System/LRU, is given in Annexure-I and contact details "
            "are given in Annexure-II respectively."
        )
        DocxExporter._add_cert_task_table(doc, cert_tasks)
        DocxExporter._add_manual_note(
            doc, manual_notes, "Certification Task Allocation", "Cert Allocation"
        )

        # ── 8. Issue of Clearances ───────────────────────────────────────────
        DocxExporter._add_heading(doc, "\n8. Issue of Clearances")
        doc.add_paragraph(
            DocxExporter._note(
                manual_notes,
                "Issue of Clearance",
                default=(
                    "a. On satisfactory completion of Software IV&V, clearance for Software would "
                    "be issued by RD, RCMA (...)\n"
                    "b. On satisfactory completion of SOFT/QT tests on the LRU, clearance for Hardware "
                    "and CEH would be issued by RDs of respective system RCMAs as per Annexure-I.\n"
                    "c. On completion of Software, Hardware and CEH clearances of LRUs, and compliance "
                    "of Systems to TARB, the installation clearances for the systems would be issued.\n"
                    "d. On satisfactory completion of ground testing/ground adaptation, the system would "
                    "be cleared by platform RD for ground integration and flight trials.\n"
                    "e. On satisfactory flight trial completion, System RCMAs shall issue Provisional "
                    "Clearance and service use for the LRUs.\n"
                    "f. Service use clearance for the Systems will be issued.\n"
                    "g. Based on the recommendation of Director(Aircraft), the platform will be issued "
                    "with RMTC/MTC by CE(A), CEMILAC."
                ),
            )
        )

        # ── 9. Certification Progress Review ─────────────────────────────────
        DocxExporter._add_heading(doc, "\n9. Certification Progress Review")
        DocxExporter._add_ce_signature_paragraph(
            doc,
            DocxExporter._note(
                manual_notes,
                "Certification Progress Review",
                default=(
                    "In order to ensure smooth progress of certification and provide "
                    "mid-course corrections, review meetings chaired by Coordinating "
                    "Director shall be conducted."
                ),
            )
            + "\n\n(CE CEMILAC)\nOS & Chief Executive (Airworthiness)",
        )

        # ── 10. Distribution List ─────────────────────────────────────────────
        DocxExporter._add_heading(doc, "\n10. Distribution List")

        DocxExporter._add_heading(
            doc, "10.1 External Organizations and Internal Departments"
        )
        doc.add_paragraph(
            DocxExporter._note(
                manual_notes,
                "External Organization",
                "External Organizations",
                default=(
                    "1. Main contractor - Request to circulate to all the work centres of the project.\n"
                    "2. DG, DGAQA, New Delhi - Request to assign field establishments for QA Coverage.\n"
                    "3. AirHQ, IAF"
                ),
            )
        )

        DocxExporter._add_heading(doc, "10.2 Internal Distribution")
        internal_items = DocxExporter._combo_items(combos, "Internal Distribution")
        doc.add_paragraph(
            DocxExporter._note(
                manual_notes,
                "Internal Distribution",
                default="\n".join(internal_items) if internal_items else "",
            )
        )

        # ── Annexure-I ─────────────────────────────────────────────────────
        doc.add_page_break()
        DocxExporter._add_heading(doc, "Annexure-I")
        doc.add_paragraph("Work Assignment List of LRUs")
        DocxExporter._add_annexure_table(
            doc,
            table_data.get("annexure1"),
            [
                "Sl No",
                "Certifiable item",
                "Design agency",
                "Designated Directorate/ RCMA",
                "Dealing Officer HW",
                "Dealing Officer CH",
                "Dealing Officer SW",
            ],
            [
                "certifiable_item",
                "design_agency",
                "designated_directorate",
                "dealing_officer_hw",
                "dealing_officer_ch",
                "dealing_officer_sw",
            ],
            blank_rows=3,
        )
        DocxExporter._add_annexure_note(
            doc, manual_notes, "Annexure-1", "Work Assignment List of LRUs"
        )

        # ── Annexure-I: Test Rigs ───────────────────────────────────────────────
        doc.add_paragraph("Test Rigs, Simulators and Ground Equipment Required")
        DocxExporter._add_annexure_table(
            doc,
            table_data.get("test_rigs"),
            [
                "Sl No",
                "Test Rig / Ground equipment Name",
                "Utilizations Phase (R&D/ Production/ Delivery to services)",
                "Design agency",
                "Designated RCMA",
                "Dealing Officer",
            ],
            [
                "test_rig_name",
                "utilization_phase",
                "design_agency",
                "designated_rcma",
                "dealing_officer",
            ],
            blank_rows=3,
        )
        DocxExporter._add_annexure_note(
            doc, manual_notes, "Test Rigs, Simulators and Ground Equipment Required"
        )

        # ── Annexure-I: Aircraft Integration ───────────────────────────────────
        DocxExporter._add_heading(
            doc, "Aircraft Integration Checks and Flight Clearance"
        )
        DocxExporter._add_annexure_table(
            doc,
            table_data.get("aircraft_checks"),
            [
                "Sl no.",
                "System Name",
                "Design agency",
                "Designated RCMA",
                "Dealing Officer",
            ],
            [
                "system_name",
                "design_agency",
                "designated_rcma",
                "dealing_officer",
            ],
            blank_rows=2,
        )
        DocxExporter._add_annexure_note(
            doc, manual_notes, "Aircraft Integration Checks and Flight Clearance"
        )

        # ── Annexure-II ──────────────────────────────────────────────────────
        doc.add_page_break()
        DocxExporter._add_heading(doc, "Annexure-II")
        doc.add_paragraph("Contact details of dealing officers and RDs")
        DocxExporter._add_annexure_table(
            doc,
            table_data.get("annexure2"),
            [
                "Sl no.",
                "Name & Designation",
                "Certification Centre",
                "E-mail id",
                "Phone number",
            ],
            ["name_designation", "certification_centre", "email", "phone"],
            blank_rows=2,
        )
        DocxExporter._add_annexure_note(
            doc,
            manual_notes,
            "Annexure-2",
            "Contact details of dealing officers and RDs",
        )

        # ── Annexure-III ─────────────────────────────────────────────────────
        doc.add_page_break()
        DocxExporter._add_heading(doc, "Annexure-III")
        DocxExporter._add_heading(doc, "Product Break Down Structure")
        DocxExporter._add_annexure3_image(doc, annexure3_image_path)

    @staticmethod
    def _note(manual_notes, *keys, default="____"):
        note = DocxExporter._manual_note(manual_notes, *keys)
        base = default
        if default == "____" and len(keys) == 1:
            key = keys[0]
            if "\n" in key or len(key) > 80:
                base = key
        if note and base:
            return f"{base}\n\n{note}"
        if note:
            return note
        return base

    @staticmethod
    def _manual_note(manual_notes, *keys):
        for key in keys:
            value = manual_notes.get(key)
            if value:
                return value
        return ""

    @staticmethod
    def _add_manual_note(doc, manual_notes, *keys):
        note = DocxExporter._manual_note(manual_notes, *keys)
        if note:
            doc.add_paragraph(note)

    @staticmethod
    def _combo_items(combos, key):
        combo = combos.get(key) if isinstance(combos, dict) else None
        if not combo or not hasattr(combo, "checkedItems"):
            return []
        return combo.checkedItems()

    @staticmethod
    def _numbered_text(items, default=""):
        items = [str(item).strip() for item in (items or []) if str(item).strip()]
        if not items:
            return default
        return "\n".join(f"{i + 1}. {item}" for i, item in enumerate(items))

    @staticmethod
    def _add_annexure_note(doc, manual_notes, *keys):
        note = DocxExporter._note(manual_notes, *keys, default="")
        if note:
            doc.add_paragraph(note)

    @staticmethod
    def _add_signature_block(doc, manual_notes):
        note = DocxExporter._note(manual_notes, "Signatures", default="")
        if note:
            doc.add_paragraph(note)
            return

        for _ in range(3):
            doc.add_paragraph("")
        p = doc.add_paragraph("(________________)")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph("________________")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    @staticmethod
    def _add_ce_signature_paragraph(doc, text):
        signature_marker = "(CE CEMILAC)"
        if signature_marker not in text:
            doc.add_paragraph(text)
            return

        body, _signature = text.split(signature_marker, 1)
        body = body.strip()
        if body:
            doc.add_paragraph(body)

        for _ in range(2):
            doc.add_paragraph("")

        for line in [
            "______________________________",
            "(CE CEMILAC)",
            "OS & Chief Executive (Airworthiness)",
        ]:
            p = doc.add_paragraph(line)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    @staticmethod
    def _add_annexure3_image(doc, image_paths):
        # Change 19: Support multiple images for Annexure-3 by accepting a list of image paths.
        if not image_paths:
            doc.add_paragraph("____")
            return

        # Handle both single path (string) and multiple paths (list) for backward compatibility
        if isinstance(image_paths, str):
            image_paths = [image_paths] if image_paths else []

        if not image_paths:
            doc.add_paragraph("____")
            return

        try:
            for image_path in image_paths:
                if not image_path or not os.path.exists(image_path):
                    continue
                paragraph = doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.add_run()
                image = DocxImage.from_file(image_path)
                max_width = Inches(6.5)
                max_height = Inches(8.6)
                if image.px_width / image.px_height > max_width / max_height:
                    run.add_picture(image_path, width=max_width)
                else:
                    run.add_picture(image_path, height=max_height)
                # Add page break after each image for clarity
                doc.add_page_break()
        except Exception:
            doc.add_paragraph("Selected PBS/Annexure-3 image(s) could not be added.")

    @staticmethod
    def _add_heading(doc, text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(12)

    @staticmethod
    def _add_stakeholders_table(doc, stakeholders):
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        headers = [
            "Sl No.",
            "Organisation Agency",
            "Role",
            "Activities towards Certification",
        ]
        for i, h in enumerate(headers):
            hdr_cells[i].text = h

        if stakeholders:
            for i, sh in enumerate(stakeholders):
                row_cells = table.add_row().cells
                row_cells[0].text = str(i + 1)
                row_cells[1].text = sh.get("org", "")
                row_cells[2].text = sh.get("role", "")
                row_cells[3].text = sh.get("activities", "")
        else:
            for _ in range(3):
                table.add_row()
        DocxExporter._format_table(table)
        DocxExporter._add_table_gap(doc)

    @staticmethod
    def _add_cert_task_table(doc, cert_tasks):
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        headers = [
            "Sl No.",
            "Certification Activity",
            "Certification Work Centre",
            "Responsible Head",
        ]
        for i, h in enumerate(headers):
            hdr_cells[i].text = h

        if cert_tasks:
            for i, ct in enumerate(cert_tasks):
                row_cells = table.add_row().cells
                row_cells[0].text = str(i + 1)
                row_cells[1].text = ct.get("activity", "")
                row_cells[2].text = ct.get("centre", "")
                row_cells[3].text = ct.get("head", "")
        else:
            for _ in range(3):
                table.add_row()
        DocxExporter._format_table(table)
        DocxExporter._add_table_gap(doc)

    @staticmethod
    def _add_annexure_table(doc, rows=None, headers=None, keys=None, blank_rows=3):
        headers = headers or ["Sl No", "ABC", "Abc2", "Abc3", "Abc4", "Abc5"]
        keys = keys or ["abc", "abc2", "abc3", "abc4", "abc5"]
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            hdr_cells[i].text = header

        if rows:
            for i, data in enumerate(rows):
                row_cells = table.add_row().cells
                row_cells[0].text = str(i + 1)
                for col_idx, key in enumerate(keys, start=1):
                    row_cells[col_idx].text = data.get(key, "")
        else:
            for _ in range(blank_rows):
                table.add_row()
        DocxExporter._format_table(table)
        DocxExporter._add_table_gap(doc)

    @staticmethod
    def _add_table_gap(doc):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(6)
        paragraph.paragraph_format.line_spacing = 1.0

    @staticmethod
    def _format_table(table):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(1)
                    paragraph.paragraph_format.space_after = Pt(1)
                    paragraph.paragraph_format.line_spacing = 1.0
                    for run in paragraph.runs:
                        run.font.name = "Calibri"
                        run.font.size = Pt(9 if row_idx == 0 else 10)
                        if row_idx == 0:
                            run.bold = True

    @staticmethod
    def _apply_document_spacing(doc):
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(4)
                paragraph.paragraph_format.line_spacing = 1.0
                continue

            is_heading = DocxExporter._is_generated_heading(paragraph)
            if is_heading:
                paragraph.paragraph_format.space_before = Pt(12)
                paragraph.paragraph_format.space_after = Pt(6)
                paragraph.paragraph_format.keep_with_next = True
            else:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(6)
                paragraph.paragraph_format.line_spacing = 1.08

            for run in paragraph.runs:
                if run.font.size is None:
                    run.font.size = Pt(11)
                if not run.font.name:
                    run.font.name = "Calibri"

        for table in doc.tables:
            DocxExporter._format_table(table)

    @staticmethod
    def _is_generated_heading(paragraph):
        if not paragraph.runs:
            return False
        text = paragraph.text.strip()
        if not text:
            return False
        if any(run.bold for run in paragraph.runs):
            first = text.split()[0]
            has_heading_size = any(
                run.font.size is not None and int(run.font.size.pt) == 12
                for run in paragraph.runs
            )
            return first[0].isdigit() or text.startswith("Annexure") or has_heading_size
        return False
