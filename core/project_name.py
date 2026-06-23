import os
import re
from typing import Optional

from .ocr import extract_file_details_from_text

# Single common words that are never valid project names
_GARBAGE_WORDS = {
    "matter", "subject", "title", "project", "name", "document", "file",
    "report", "draft", "final", "copy", "note", "notes", "unknown", "untitled",
    "n/a", "na", "none", "null", "not found", "tbd", "tbc", "yes", "no",
    "true", "false", "page", "sheet", "form", "ref", "reference",
}

_GARBAGE_PATTERNS = [
    r'^\d+$',                        # pure number
    r'^[A-Z0-9\-_]{2,15}$',         # looks like a code/ID e.g. "REV-01"
    r'^\s*$',                        # blank
    r'^(fig|table|appendix)[\s\.]',  # figure/table captions
    r'^\W+$',                        # only punctuation
]


def _is_garbage_result(s: str) -> bool:
    """Return True if the string is clearly NOT a valid project name."""
    if not s:
        return True
    clean = s.strip().lower()
    if clean in _GARBAGE_WORDS:
        return True
    if len(clean) < 5:
        return True
    for pat in _GARBAGE_PATTERNS:
        if re.match(pat, clean, re.IGNORECASE):
            return True
    return False


def _clean_candidate(s: str) -> str:
    if not s:
        return ""
    s = s.strip(' \t\n\r\"\'')
    s = re.sub(r'\s{2,}', ' ', s)
    return s


def _looks_like_title(line: str, strict: bool = False) -> bool:
    if not line or len(line) < 5:
        return False
    low = line.lower()

    NOISE_KEYWORDS = [
        "page", "doc", "file", "issue", "rev", "date", "revision",
        "prepared by", "approved by", "checked by", "sheet", "drawing no",
        "dwg no", "reference", "version", "amendment", "contract no",
        "purchase order", "po no", "author", "signature",
        "www.", "http", "@", "tel:", "fax:", "email:",
        "4.5hours", "module", "hours",
    ]
    if any(k in low for k in NOISE_KEYWORDS):
        return False

    if len(line) > 200:
        return False

    alpha_ratio = sum(c.isalpha() for c in line) / max(len(line), 1)
    if alpha_ratio < 0.4:
        return False

    words = re.findall(r"[A-Za-z0-9&()\-]{2,}", line)
    if len(words) < 2:
        return False

    caps = sum(1 for w in words if w[0].isupper())

    if caps >= max(2, len(words) // 2):
        return True
    if line.strip() == line.strip().upper() and len(line) <= 120:
        return True
    if not strict and len(words) >= 4:
        return True

    return False


def _score_title_candidate(line: str, position: int) -> float:
    score = 0.0
    score += max(0, (12 - position) * 2)

    words = re.findall(r"[A-Za-z0-9&()\-]{2,}", line)
    caps = sum(1 for w in words if w[0].isupper())
    score += caps * 1.5

    length = len(line.strip())
    if 15 <= length <= 120:
        score += 5
    elif length < 15:
        score -= 3

    low = line.lower()
    if any(k in low for k in ["project", "contract", "works", "construction",
                               "development", "registration", "firm", "business",
                               "process", "legal", "step"]):
        score += 8

    # Big bold lines with colons mid-sentence are still titles (e.g. "Legal issues: Firm...")
    # but lines that END in colon are headers/labels
    if line.strip().endswith(":"):
        score -= 10

    if re.search(r'\b(plot|lot|block|phase|sector|zone)\b', low):
        score -= 3

    # Penalise short single-word candidates heavily
    if len(words) == 1:
        score -= 20

    # Reward longer descriptive titles
    if len(words) >= 5:
        score += len(words) * 0.5

    return score


_EXPLICIT_LABEL_RE = re.compile(
    r'\b(?:Project\s*(?:Name|Title|Description)|Contract\s*(?:Name|Title)?|'
    r'Title|Subject|Works\s*Title|Scheme\s*Name)\b[:\-]?\s*(.+)$',
    re.IGNORECASE,
)

_INLINE_PROJECT_RE = re.compile(
    r'(?:for\s+(?:the\s+)?|re[:\s]+|subject[:\s]+)([A-Z][A-Za-z0-9 ,&()\-]{10,80})',
    re.IGNORECASE,
)


def _extract_from_text_heuristics(text: str) -> Optional[str]:
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    top = lines[:25]  # slightly wider window

    # Pass 1: explicit label match
    for i, line in enumerate(top):
        m = _EXPLICIT_LABEL_RE.search(line)
        if m:
            cand = _clean_candidate(m.group(1))
            if cand and len(cand) < 10 and i + 1 < len(top):
                next_line = top[i + 1]
                if _looks_like_title(next_line):
                    cand = next_line
            if cand and len(cand) >= 5 and not _is_garbage_result(cand):
                return cand[:150]

    # Pass 2: score all title-like candidates
    candidates = []
    for i, line in enumerate(top):
        if _looks_like_title(line):
            s = _score_title_candidate(line, i)
            candidates.append((s, line))

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        best_score, best = candidates[0]
        # Only accept if score is meaningful and not garbage
        if best_score > 0 and len(best) >= 5 and not _is_garbage_result(best):
            return best[:150]

    # Pass 3: inline pattern
    full_top = " ".join(top)
    m = _INLINE_PROJECT_RE.search(full_top)
    if m:
        cand = _clean_candidate(m.group(1))
        if cand and len(cand) >= 8 and not _is_garbage_result(cand):
            return cand[:150]

    return None


def _extract_from_pdf_spans(pdf_handler: object) -> Optional[str]:
    pages_to_check = [0, 1]
    best_score, best_text = 0.0, ""

    for page_num in pages_to_check:
        try:
            page = pdf_handler.get_page(page_num)
            if not page:
                continue
            d = page.get_text("dict")
        except Exception:
            continue

        for block in d.get("blocks", []) or []:
            for line in block.get("lines", []) or []:
                for span in line.get("spans", []) or []:
                    size = span.get("size", 0)
                    flags = span.get("flags", 0)
                    is_bold = bool(flags & 16)
                    txt = _clean_candidate(span.get("text", ""))

                    if len(txt) < 4:
                        continue
                    if _is_garbage_result(txt):
                        continue
                    if not _looks_like_title(txt):
                        continue

                    length_bonus = 1 + min(len(txt), 100) / 100.0
                    bold_bonus = 1.2 if is_bold else 1.0
                    score = size * length_bonus * bold_bonus

                    if score > best_score:
                        best_score = score
                        best_text = txt

        if best_text and page_num == 0:
            break

    return best_text[:150] if best_text else None


def infer_project_name(
    text: str = "",
    pdf_handler: Optional[object] = None,
    file_path: Optional[str] = None,
    quick: bool = False,
) -> str:
    """Infer a project name using multiple signals, returning empty string when unknown.

    Signals (priority order):
    1. Structured OCR extraction — only if result passes garbage filter
    2. Explicit label match  (e.g. "Project Name: XYZ")
    3. Large/bold font span on first PDF page
    4. Scored heuristic scan of top raw-text lines
    5. Filename / parent directory last resort
    """

    # 1) Structured OCR extraction — validate before trusting
    try:
        details = extract_file_details_from_text(text or "")
        pn = (details or {}).get("project_name")
        if pn:
            pn_clean = _clean_candidate(str(pn))
            if not _is_garbage_result(pn_clean):
                return pn_clean
            # If OCR gave garbage, fall through — don't return early
    except Exception:
        pass

    # 2) Explicit label match (fast, high-confidence)
    try:
        lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
        for line in lines[:20]:
            m = _EXPLICIT_LABEL_RE.search(line)
            if m:
                cand = _clean_candidate(m.group(1))
                if cand and len(cand) >= 5 and not _is_garbage_result(cand):
                    return cand[:150]
    except Exception:
        pass

    # 3) PDF span analysis (font size + bold)
    if not quick:
        try:
            if pdf_handler and getattr(pdf_handler, "doc", None):
                result = _extract_from_pdf_spans(pdf_handler)
                if result:
                    return result
        except Exception:
            pass

    # 4) Scored heuristic scan of raw text
    try:
        result = _extract_from_text_heuristics(text)
        if result:
            return result
    except Exception:
        pass

    # 5) Filename fallback
    try:
        if file_path:
            base = os.path.splitext(os.path.basename(file_path))[0]
            if base and len(base) > 2:
                cleaned = _clean_candidate(base.replace("_", " ").replace("-", " "))
                if cleaned and not _is_garbage_result(cleaned):
                    return cleaned[:150]
            parent = os.path.basename(os.path.dirname(file_path))
            if parent and len(parent) > 2:
                return _clean_candidate(parent)[:150]
    except Exception:
        pass

    return ""