import re

import cv2
import fitz
import numpy as np
import pytesseract
import requests
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QImage

from core.ocr import extract_file_details_from_text
from core.signature import detect_signature_in_image, document_has_valid_signature

from .project_name import infer_project_name

# Change 13: Add optional transformer-based summarization support for selected text.
# This enables the app to use an advanced BART + grammar correction model when installed,
# while preserving the existing local summarizer fallback when dependencies are missing.
try:
    import torch
    from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

    HAS_TRANSFORMERS = True
except ImportError:
    torch = None
    AutoTokenizer = None
    AutoModelForSeq2SeqLM = None
    HAS_TRANSFORMERS = False

SUMMARIZER_MODEL = "facebook/bart-large-cnn"
GRAMMAR_MODEL = "vennify/t5-base-grammar-correction"
_device = (
    torch.device("cuda" if torch.cuda.is_available() else "cpu")
    if HAS_TRANSFORMERS
    else None
)

_sum_tok = None
_sum_model = None
_gram_tok = None
_gram_model = None


def simple_summarize(text, target_ratio=0.5, min_sentences=10, max_sentences=15):
    if not text or len(text.strip()) < 20:
        return text.strip()

    clean_text = re.sub(r"(?<!\n)\n(?!\n)", " ", text)
    sentences = re.split(r"(?<=[.!?])\s+", clean_text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 10]

    if not sentences:
        return ""

    num_sentences = int(len(sentences) * target_ratio)
    num_sentences = max(min_sentences, min(num_sentences, max_sentences))

    if len(sentences) <= num_sentences:
        return " ".join(sentences)

    stop_words = {
        "the",
        "is",
        "in",
        "and",
        "to",
        "of",
        "a",
        "for",
        "on",
        "with",
        "as",
        "by",
        "this",
        "that",
        "it",
        "are",
        "be",
        "or",
        "an",
        "at",
        "from",
        "which",
        "will",
    }

    words = re.findall(r"\b[a-zA-Z]{2,}\b", clean_text.lower())
    freq = {}
    for w in words:
        if w not in stop_words:
            freq[w] = freq.get(w, 0) + 1

    max_freq = max(freq.values()) if freq else 1
    for w in freq:
        freq[w] = freq[w] / max_freq

    scores = {}
    for i, s in enumerate(sentences):
        score = 0
        s_words = re.findall(r"\b[a-zA-Z]{2,}\b", s.lower())
        for w in s_words:
            if w in freq:
                score += freq[w]
        score = score / max(len(s_words), 1)
        if i < 2:
            score += 0.5
        scores[i] = score

    top_indices = sorted(sorted(scores, key=scores.get, reverse=True)[:num_sentences])
    return " ".join([sentences[i] for i in top_indices])


# Change 14: Load transformer models lazily on first summary request.
# This avoids heavy startup cost and only loads the model when needed for selected-text summarization.
def _load_transformer_models():
    global _sum_tok, _sum_model, _gram_tok, _gram_model
    if not HAS_TRANSFORMERS:
        raise RuntimeError(
            "Transformers summarization model is unavailable. "
            "Install transformers, torch, and sentencepiece."
        )
    if _sum_tok is None or _gram_tok is None:
        _sum_tok = AutoTokenizer.from_pretrained(SUMMARIZER_MODEL)
        _sum_model = AutoModelForSeq2SeqLM.from_pretrained(SUMMARIZER_MODEL).to(_device)
        _gram_tok = AutoTokenizer.from_pretrained(GRAMMAR_MODEL)
        _gram_model = AutoModelForSeq2SeqLM.from_pretrained(GRAMMAR_MODEL).to(_device)


# Change 15: Use transformer models to produce an abstractive summary and then grammar-correct it.
def summarize_with_transformers(
    text,
    max_length=130,
    min_length=30,
    num_beams=4,
    length_penalty=2.0,
):
    if not text or len(text.strip()) < 20:
        return text.strip()
    _load_transformer_models()

    inputs = _sum_tok(
        text,
        return_tensors="pt",
        max_length=1024,
        truncation=True,
    ).to(_device)

    with torch.no_grad():
        summary_ids = _sum_model.generate(
            inputs["input_ids"],
            max_length=max_length,
            min_length=min_length,
            num_beams=num_beams,
            length_penalty=length_penalty,
            early_stopping=True,
        )
    raw_summary = _sum_tok.decode(summary_ids[0], skip_special_tokens=True)

    gram_inputs = _gram_tok(
        f"grammar: {raw_summary}",
        return_tensors="pt",
        truncation=True,
    ).to(_device)

    with torch.no_grad():
        gram_ids = _gram_model.generate(gram_inputs["input_ids"], max_length=200)
    corrected = _gram_tok.decode(gram_ids[0], skip_special_tokens=True).strip()
    if corrected:
        corrected = corrected[0].upper() + corrected[1:]
        if corrected[-1] not in ".!?":
            corrected += "."
    return corrected or raw_summary


# Change 16: Run advanced summarization in a background thread to keep the UI responsive.
class SummarizationWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    status = pyqtSignal(str)

    def __init__(self, text, parent=None):
        super().__init__(parent)
        self.text = text

    def run(self):
        if not self.text.strip():
            self.finished.emit("")
            return

        self.status.emit("Summarizing selection...")
        try:
            if HAS_TRANSFORMERS:
                summary = summarize_with_transformers(self.text)
            else:
                self.status.emit("Model not installed; using local summarizer.")
                summary = simple_summarize(self.text)
            self.finished.emit(summary)
        except Exception as exc:
            self.error.emit(str(exc))


# ---------------------------------------------------------------------------
# TOPIC ALIAS MAP
# Each entry is a list of keyword sets. Matching is HEADING-AWARE:
# keywords must appear on a short standalone line (heading), not in body text.
#
# "heading-only" topics use _heading_found_in_text() — strict line check.
# "anywhere" topics (like Signatures) use _keywords_found_in_window()
#   since signatures appear in tables, not as headings.
# ---------------------------------------------------------------------------
TOPIC_ALIASES = {
    "Introduction": {
        "mode": "heading",
        "sets": [
            ["Introduction"],
        ],
    },
    "Scope": {
        "mode": "heading",
        "sets": [
            ["Scope"],
        ],
    },
    "Product Breakdown": {
        "mode": "heading",
        "sets": [
            ["Product", "Breakdown"],
            ["Product", "Break", "Down"],
            ["PBS"],
            ["Product", "Structure"],
        ],
    },
    "System LRU and Module Details": {
        "mode": "heading",
        "sets": [
            ["System", "LRU"],
            ["LRU", "Module"],
            ["LRU", "Details"],
            ["Work", "Assignment", "LRU"],
            ["LRU", "List"],
        ],
    },
    "Signatures": {
        # Signatures appear in tables, not as a standalone heading
        "mode": "anywhere",
        "sets": [
            ["Signature"],
            ["Prepared", "by"],
            ["Approved", "by"],
            ["Reviewed", "by"],
            ["Signed"],
        ],
    },
    "Test Rig Simulators and Ground equipment": {
        "mode": "heading",
        "sets": [
            ["Test", "Rig"],
            ["Simulator"],
            ["Ground", "Equipment"],
            ["Ground", "Support"],
            ["Test", "Equipment"],
            ["GSE"],
        ],
    },
    "Work breakdown structure": {
        "mode": "heading",
        "sets": [
            ["Work", "Breakdown", "Structure"],
            ["Work", "Break", "Down"],
            ["WBS"],
            ["Work", "Assignment"],
            ["Annexure"],
        ],
    },
}

# Max characters a line can have to be considered a heading
HEADING_MAX_LEN = 80

# Heading must not contain these patterns (sentence indicators)
HEADING_BODY_PATTERNS = [
    r"\b(is|are|was|were|has|have|had|will|would|can|could|should|may|might)\b",
    r"[,;]",  # commas/semicolons suggest body text
    r"\b(if|when|because|although|however|therefore|thus)\b",
    r"•",  # bullet points
]


class AnalysisWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    status = pyqtSignal(str)

    def __init__(self, text, topics, pdf_handler=None):
        super().__init__()
        self.text = text
        self.topics = topics
        self.pdf_handler = pdf_handler
        # Pre-split lines once for efficiency
        self._lines = [l.strip() for l in text.split("\n") if l.strip()]

    def clean_extracted_name(self, name):
        if not name:
            return ""
        name = name.strip(" \"'\t\n\r")
        name = re.sub(r'^[a-zA-Z]\s*["\']\s*', "", name)
        name = re.sub(r"^[a-zA-Z]\s+", "", name)
        name = re.sub(r"^[^a-zA-Z0-9]+", "", name)
        return name.strip(" \"'")

    def run(self):
        try:
            results = {"project_name": "", "introduction": "", "found_topics": []}

            self.status.emit("Parsing Document Structure...")

            # Extract project name using layered heuristics
            explicit_match = re.search(
                r"(?:Project\s*Name|Project\s*Title|Title|Subject)[\s:]*(.+)",
                self.text,
                re.IGNORECASE,
            )
            if explicit_match and not explicit_match.group(1).strip().startswith("___"):
                results["project_name"] = self.clean_extracted_name(
                    explicit_match.group(1).strip()[:100]
                )

            if not results["project_name"]:
                try:
                    # If a PDF handler is available, use full inference (may scan page spans).
                    if getattr(self, "pdf_handler", None) and getattr(
                        self.pdf_handler, "doc", None
                    ):
                        pn = infer_project_name(
                            self.text, pdf_handler=self.pdf_handler, quick=False
                        )
                    else:
                        pn = infer_project_name(self.text, quick=True)

                    if pn:
                        results["project_name"] = self.clean_extracted_name(pn[:100])
                except Exception:
                    pass

            # Extract introduction
            intro_text = self._extract_introduction()
            intro_is_short = len(intro_text.split()) <= 100

            # Validate checklist
            for topic in self.topics:
                if topic.lower() == "introduction":
                    if intro_text or self._topic_found_in_text(topic, self.text):
                        results["found_topics"].append(topic)
                elif topic.lower() == "signatures":
                    continue
                else:
                    if self._topic_found_in_text(topic, self.text):
                        results["found_topics"].append(topic)

            # Return the raw extracted introduction — do NOT summarize or rewrite.
            # The UI will display this verbatim so the user can review and edit it
            # before generating the PDF.
            results["introduction"] = intro_text if intro_text else ""

            self.finished.emit(results)

        except Exception as e:
            self.error.emit(str(e))

    def _is_heading_line(self, line):
        """
        Returns True if the line looks like a section heading:
        - Short enough (under HEADING_MAX_LEN chars)
        - Does not read like a sentence (no verbs, commas, conjunctions)
        - Not a bullet point or list item
        - Optionally starts with a number like "1." or "1.1"
        """
        # Strip leading numbering like "1.", "7.2", "A."
        core = re.sub(r"^[\d]+[\d.]*\s+", "", line).strip()
        core = re.sub(r"^[A-Z]\.\s+", "", core).strip()

        if len(core) > HEADING_MAX_LEN:
            return False

        for pattern in HEADING_BODY_PATTERNS:
            if re.search(pattern, core, re.IGNORECASE):
                return False

        return True

    def _heading_found_in_text(self, keyword_set):
        """
        Returns True if all keywords in the set appear together
        on a single heading line (not in body text).
        """
        for line in self._lines:
            if not self._is_heading_line(line):
                continue
            if all(
                re.search(r"\b" + re.escape(k) + r"\b", line, re.IGNORECASE)
                for k in keyword_set
            ):
                return True
        return False

    def _keywords_found_in_window(self, keywords, text, window=120):
        """
        Returns True if all keywords appear within `window` characters
        of the first keyword's match position (used for non-heading topics).
        """
        if not keywords:
            return False
        for match in re.finditer(re.escape(keywords[0]), text, re.IGNORECASE):
            snippet = text[match.start() : match.start() + window]
            if all(
                re.search(r"\b" + re.escape(k) + r"\b", snippet, re.IGNORECASE)
                for k in keywords[1:]
            ):
                return True
        return False

    def _topic_found_in_text(self, topic, text):
        """
        Main dispatcher. Uses TOPIC_ALIASES to determine match mode:
        - 'heading': keyword sets must match on a standalone heading line only
        - 'anywhere': keyword sets can match anywhere in the text

        Falls back to separator-flexible regex on the raw topic string
        if no alias entry exists.
        """
        alias_entry = TOPIC_ALIASES.get(topic)

        if alias_entry:
            mode = alias_entry.get("mode", "heading")
            for keyword_set in alias_entry["sets"]:
                if mode == "heading":
                    if self._heading_found_in_text(keyword_set):
                        return True
                else:
                    if self._keywords_found_in_window(keyword_set, text, window=120):
                        return True
            return False

        # --- Fallback: no alias entry, use separator-flexible regex ---
        FILLERS = {"and", "or", "the", "of", "to", "in", "for"}
        raw_words = re.split(r"[\s\-_&/,]+", topic)
        keywords = [w for w in raw_words if w and w.lower() not in FILLERS]

        if not keywords:
            return False

        if len(keywords) == 1:
            return bool(
                re.search(r"\b" + re.escape(keywords[0]) + r"\b", text, re.IGNORECASE)
            )

        # Try separator-flexible pattern
        sep = r"[\s\-_&/,]+(?:and\s+|or\s+)?"
        pattern = sep.join(re.escape(k) for k in keywords)
        if re.search(pattern, text, re.IGNORECASE):
            return True

        # Fallback window check
        return self._keywords_found_in_window(keywords, text, window=80)

    def _extract_introduction(self):
        lines = self.text.split("\n")
        intro_lines = []
        capturing = False

        for line in lines:
            clean_line = line.strip()
            if not clean_line:
                continue

            if not capturing:
                is_intro_heading = re.match(
                    r"^(?:\d+\.?\s*)?Introduction\b", clean_line, re.IGNORECASE
                )
                is_toc_entry = re.search(r"(?:\.{3,}|\b\d+$)", clean_line)

                if is_intro_heading and not is_toc_entry and len(clean_line) < 50:
                    capturing = True
            else:
                is_num_heading = re.match(r"^\d+\.\s+[A-Z]", clean_line)
                is_caps_heading = clean_line.isupper() and len(clean_line) > 4
                if (is_num_heading or is_caps_heading) and len(clean_line) < 60:
                    break
                intro_lines.append(clean_line)

        return " ".join(intro_lines).strip()

    def _try_ai_summarization(self, intro_text, is_short):
        try:
            requests.get("http://localhost:11434/", timeout=2)

            if is_short:
                prompt = f"""Extract the Project Name from:
                {self.text[:2000]}

                Format: PROJECT_NAME: [name or "Not Found"]
                """
            else:
                prompt = f"""Extract two things:
                1. Project Name from: {self.text[:2000]}
                2. Summarize (max 15 sentences): {intro_text[:6000]}

                Format:
                PROJECT_NAME: [name or "Not Found"]
                INTRODUCTION_SUMMARY: [summary]
                """

            response = requests.post(
                "http://localhost:11434/api/generate",
                json={"model": "gemma:2b", "prompt": prompt, "stream": False},
                timeout=180,
            )

            if response.status_code == 200:
                ai_text = response.json()["response"].strip()

                intro_match = re.search(
                    r"INTRODUCTION_SUMMARY:\s*(.+)", ai_text, re.IGNORECASE | re.DOTALL
                )
                if intro_match:
                    return intro_match.group(1).strip()

            return simple_summarize(intro_text) if not is_short else intro_text

        except Exception:
            self.status.emit("AI unreachable. Using offline algorithms...")
            return simple_summarize(intro_text) if not is_short else intro_text


class FileDetailsWorker(QThread):
    finished = pyqtSignal(dict)

    def __init__(self, pdf_handler):
        super().__init__()
        self.pdf_handler = pdf_handler

    def run(self):
        try:
            if not self.pdf_handler or not self.pdf_handler.doc:
                self.finished.emit({"error": "No document loaded"})
                return

            full_ocr_text = ""
            signature_found = document_has_valid_signature(self.pdf_handler.doc)

            # OCR FIRST 3 PAGES
            max_pages = min(3, len(self.pdf_handler.doc))

            for page_index in range(max_pages):
                page = self.pdf_handler.get_page(page_index)

                # HIGH DPI RENDERING
                pix = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0))

                qimg = QImage(
                    pix.samples,
                    pix.width,
                    pix.height,
                    pix.stride,
                    QImage.Format.Format_RGB888,
                ).copy()

                width = qimg.width()
                height = qimg.height()

                ptr = qimg.bits()
                ptr.setsize(height * width * 3)

                arr = np.array(ptr).reshape(height, width, 3)

                # OCR PREPROCESSING
                processed = self.preprocess_for_ocr(arr)

                # OCR CONFIG FOR TABLES
                config = r"--oem 3 --psm 6"

                ocr_text = pytesseract.image_to_string(
                    processed, lang="eng", config=config
                )

                # ALSO INCLUDE NATIVE PDF TEXT
                native_text = page.get_text()

                combined = native_text + "\n" + ocr_text

                full_ocr_text += "\n\n" + combined

            details = self.extract_details(full_ocr_text)

            # If OCR-based extraction didn't find a good project name, try the layered inference
            try:
                # Prefer full PDF-based inference when possible so large title spans
                # override short/incorrect OCR extracts (e.g., stray words like 'matter').
                if getattr(self, "pdf_handler", None) and getattr(
                    self.pdf_handler, "doc", None
                ):
                    pn = infer_project_name(
                        full_ocr_text, pdf_handler=self.pdf_handler, quick=False
                    )
                else:
                    pn = infer_project_name(full_ocr_text, quick=True)

                if pn and pn.lower() not in {"not found", ""}:
                    current = (details.get("project_name") or "").strip()
                    # override when current is missing/placeholder or new candidate is substantially longer
                    if current.lower() in {"not found", ""} or len(pn) > max(
                        12, len(current) + 5
                    ):
                        details["project_name"] = pn
            except Exception:
                pass

            details["has_signature"] = signature_found

            self.finished.emit(details)

        except ImportError:
            self.finished.emit(
                {"error": "Please install pytesseract and opencv-python"}
            )

        except Exception as e:
            self.finished.emit({"error": str(e)})

    def preprocess_for_ocr(self, image):
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)

        # ENLARGE IMAGE
        gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

        # DENOISE
        gray = cv2.fastNlMeansDenoising(gray)

        # SHARPEN
        kernel = np.array([[-1, -1, -1], [-1, 9, -1], [-1, -1, -1]])

        sharp = cv2.filter2D(gray, -1, kernel)

        # THRESHOLD
        thresh = cv2.adaptiveThreshold(
            sharp, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 15
        )

        return thresh

    def detect_signature(self, img_array):
        return detect_signature_in_image(img_array)

    def extract_details(self, text):
        return extract_file_details_from_text(text)


class SignatureValidationWorker(QThread):
    finished = pyqtSignal(bool)
    error = pyqtSignal(str)

    def __init__(self, pdf_handler):
        super().__init__()
        self.pdf_handler = pdf_handler

    def run(self):
        try:
            if not self.pdf_handler or not self.pdf_handler.doc:
                self.finished.emit(False)
                return

            self.finished.emit(document_has_valid_signature(self.pdf_handler.doc))
        except ImportError:
            self.error.emit("Please install opencv-python and numpy")
        except Exception as e:
            self.error.emit(str(e))
