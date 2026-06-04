import re
import pytesseract
import fitz
import numpy as np
import cv2

def pixmap_to_rgb_array(pix):
    samples = np.frombuffer(pix.samples, dtype=np.uint8)
    channels = pix.n
    arr = samples.reshape(pix.height, pix.width, channels)

    if channels == 4:
        arr = arr[:, :, :3]
    elif channels == 1:
        arr = np.repeat(arr, 3, axis=2)

    return np.ascontiguousarray(arr)


def rect_to_pdf_rect(page, view_width, view_height, rect, padding_points=2.0):
    scale_x = page.rect.width / max(view_width, 1)
    scale_y = page.rect.height / max(view_height, 1)

    pdf_rect = fitz.Rect(
        rect.left() * scale_x,
        rect.top() * scale_y,
        rect.right() * scale_x,
        rect.bottom() * scale_y,
    )
    pdf_rect = pdf_rect + (
        -padding_points,
        -padding_points,
        padding_points,
        padding_points,
    )
    return pdf_rect & page.rect


def render_pdf_region(page, pdf_rect, dpi=500, max_pixels=24000000):
    zoom = dpi / 72.0
    estimated_pixels = (pdf_rect.width * zoom) * (pdf_rect.height * zoom)
    if estimated_pixels > max_pixels:
        zoom = (max_pixels / max(pdf_rect.width * pdf_rect.height, 1)) ** 0.5

    pix = page.get_pixmap(
        matrix=fitz.Matrix(zoom, zoom),
        clip=pdf_rect,
        colorspace=fitz.csRGB,
        alpha=False,
    )
    return pixmap_to_rgb_array(pix)


def _deskew(gray):
    

    coords = np.column_stack(np.where(gray < 245))
    if coords.size == 0:
        return gray

    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle

    if abs(angle) < 0.2 or abs(angle) > 8:
        return gray

    height, width = gray.shape
    matrix = cv2.getRotationMatrix2D((width // 2, height // 2), angle, 1.0)
    return cv2.warpAffine(
        gray,
        matrix,
        (width, height),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_REPLICATE,
    )


def build_ocr_variants(image):
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)

    height, width = gray.shape
    min_side = min(height, width)
    if min_side < 900:
        scale = min(4.0, 900 / max(min_side, 1))
        gray = cv2.resize(
            gray,
            None,
            fx=scale,
            fy=scale,
            interpolation=cv2.INTER_CUBIC,
        )

    gray = _deskew(gray)
    gray = cv2.fastNlMeansDenoising(gray, None, 12, 7, 21)

    clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
    contrast = clahe.apply(gray)

    blur = cv2.GaussianBlur(contrast, (0, 0), 1.0)
    sharp = cv2.addWeighted(contrast, 1.6, blur, -0.6, 0)

    _, otsu = cv2.threshold(
        sharp,
        0,
        255,
        cv2.THRESH_BINARY + cv2.THRESH_OTSU,
    )

    adaptive = cv2.adaptiveThreshold(
        sharp,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        35,
        11,
    )

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 1))
    adaptive = cv2.morphologyEx(adaptive, cv2.MORPH_OPEN, kernel)

    variants = [sharp, otsu, adaptive]
    if np.mean(gray) < 120:
        variants.extend([cv2.bitwise_not(v) for v in variants])

    bordered = []
    for variant in variants:
        bordered.append(
            cv2.copyMakeBorder(
                variant,
                30,
                30,
                30,
                30,
                cv2.BORDER_CONSTANT,
                value=255,
            )
        )
    return bordered


def _clean_text(text):
    text = text.replace("\x0c", "")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _normalize_details_text(text):
    replacements = {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2212": "-",
        "\ufb01": "fi",
        "\ufb02": "fl",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    return text


def _compact_details_text(text):
    text = text.replace("|", " ")
    text = text.replace("]", " ")
    text = text.replace("[", " ")
    return re.sub(r"\s+", " ", text).strip()


def _clean_field_value(value):
    value = value.strip(" :;|[](){}<>\"'")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def _first_match(text, patterns):
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = _clean_field_value(match.group(1))
            if value:
                return value
    return ""


def _looks_like_project_label(line):
    key = re.sub(r"[^a-z0-9]+", "", line.lower())
    label_parts = [
        "title",
        "projectsystem",
        "projsys",
        "pbswbs",
        "lrusystem",
        "systempartno",
        "criticallevel",
        "documentclassification",
        "copyno",
        "noofpage",
        "designagencylogo",
        "namedesignation",
        "signature",
        "preparedby",
        "reviewedby",
        "approvedby",
        "issueno",
        "revno",
        "issuedate",
        "docno",
    ]
    return any(part in key for part in label_parts)


def _is_project_candidate(line):
    value = _clean_field_value(line)
    if not value or len(value) < 2:
        return False

    lower = value.lower()
    if _looks_like_project_label(value):
        return False
    if lower.startswith("for ") and "<" in value:
        return False
    if "<" in value and ">" in value:
        return False
    if re.search(r"\b(name|designation|agency|signature|restricted|secret)\b", lower):
        return False
    if re.fullmatch(r"[-_.\s]+", value):
        return False

    return bool(re.search(r"[A-Za-z0-9]", value))


def _extract_project_name_from_lines(text):
    lines = [
        _clean_field_value(line)
        for line in text.splitlines()
        if _clean_field_value(line)
    ]

    for index, line in enumerate(lines):
        title_match = re.search(r"\bTitle\s*[:\-]?\s*(.+)$", line, re.IGNORECASE)
        if title_match:
            title_value = _clean_field_value(title_match.group(1))
            if _is_project_candidate(title_value):
                return title_value

            for candidate in lines[index + 1:index + 6]:
                if _is_project_candidate(candidate):
                    return candidate

        key = re.sub(r"[^a-z0-9]+", "", line.lower())
        if "projectsystem" in key or "projsys" in key:
            for candidate in lines[index + 1:index + 5]:
                if _is_project_candidate(candidate):
                    return candidate

    return ""


def extract_file_details_from_text(text):
    normalized = _normalize_details_text(text or "")
    compact = _compact_details_text(normalized)

    results = {
        "file_no": "Not Found",
        "issue_no": "Not Found",
        "date_of_issue": "Not Found",
        "project_name": "Not Found",
        "has_signature": True,
    }

    file_no = _first_match(
        compact,
        [
            r"\bDoc(?:ument)?\s*No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\/_.-]{2,})",
            r"\bFile\s*No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\/_.-]{2,})",
            r"\bReference\s*No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\/_.-]{2,})",
        ],
    )
    if file_no:
        results["file_no"] = file_no

    issue_no = _first_match(
        compact,
        [
            r"\bIssue\s*No\.?\s*\/?\s*(?:Rev(?:ision)?\s*)?No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\/_.-]{1,})",
            r"\bIssue\s*No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\/_.-]{1,})",
            r"\bRev\s*No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\/_.-]{1,})",
        ],
    )
    if issue_no:
        results["issue_no"] = issue_no

    issue_date = _first_match(
        compact,
        [
            r"\bIssue\s*date\s*[:\-]?\s*([0-9]{1,2}[\/\-.][0-9]{1,2}[\/\-.][0-9]{2,4})",
            r"\bDate\s*of\s*Issue\s*[:\-]?\s*([0-9]{1,2}[\/\-.][0-9]{1,2}[\/\-.][0-9]{2,4})",
            r"\bDate\s*[:\-]?\s*([0-9]{1,2}[\/\-.][0-9]{1,2}[\/\-.][0-9]{2,4})",
            r"\b([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})\b",
        ],
    )
    if issue_date:
        results["date_of_issue"] = issue_date

    project_name = _first_match(
        compact,
        [
            r"\bProject\s*Name\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\/,().& _-]{1,150})",
            r"\bProject\s*Title\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\/,().& _-]{1,150})",
            r"\bSubject\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\/,().& _-]{1,150})",
        ],
    )

    if project_name:
        project_name = re.split(
            r"\b(?:Doc|Document|File|Issue|Rev|Date|Copy|No\s*of\s*page)\b",
            project_name,
            flags=re.IGNORECASE,
        )[0]
        project_name = _clean_field_value(project_name)

    if not project_name or _looks_like_project_label(project_name):
        project_name = _extract_project_name_from_lines(normalized)

    if project_name:
        results["project_name"] = project_name[:150]

    return results


def _confidence_score(data):
    confidences = []
    for value in data.get("conf", []):
        try:
            conf = float(value)
        except (TypeError, ValueError):
            continue
        if conf >= 0:
            confidences.append(conf)

    if not confidences:
        return 0.0

    return sum(confidences) / len(confidences)


def _text_quality_score(text, confidence):
    if not text:
        return 0.0

    readable_chars = sum(ch.isalnum() for ch in text)
    noise_chars = sum(ch in "{}[]|~`^" for ch in text)
    word_count = len(re.findall(r"\w+", text))
    return confidence + min(readable_chars, 500) * 0.05 + word_count - noise_chars * 3


def run_tesseract_ocr(image, lang="eng"):
    
    variants = build_ocr_variants(image)
    configs = [
        "--oem 3 --psm 6 --dpi 500 -c preserve_interword_spaces=1",
        "--oem 3 --psm 4 --dpi 500 -c preserve_interword_spaces=1",
        "--oem 3 --psm 11 --dpi 500 -c preserve_interword_spaces=1",
    ]

    best_text = ""
    best_confidence = 0.0
    best_score = 0.0

    for variant in variants:
        for config in configs:
            data = pytesseract.image_to_data(
                variant,
                lang=lang,
                config=config,
                output_type=pytesseract.Output.DICT,
            )
            text = _clean_text(
                pytesseract.image_to_string(
                    variant,
                    lang=lang,
                    config=config,
                )
            )
            confidence = _confidence_score(data)
            score = _text_quality_score(text, confidence)

            if score > best_score:
                best_text = text
                best_confidence = confidence
                best_score = score

    return best_text, best_confidence


def extract_region_text(page, view_width, view_height, rect, lang="eng"):
    pdf_rect = rect_to_pdf_rect(page, view_width, view_height, rect)
    native_text = _clean_text(page.get_textbox(pdf_rect))
    image = render_pdf_region(page, pdf_rect)
    ocr_text, confidence = run_tesseract_ocr(image, lang=lang)

    if native_text and (confidence < 65 or len(native_text) >= len(ocr_text) * 0.75):
        return native_text

    if ocr_text:
        return ocr_text

    return native_text
