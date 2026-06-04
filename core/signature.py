import fitz
import numpy as np
import cv2


def page_to_rgb_array(page, zoom=2.5):
    pix = page.get_pixmap(
        matrix=fitz.Matrix(zoom, zoom),
        colorspace=fitz.csRGB,
        alpha=False,
    )
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n
    )


def detect_signature_in_image(img_array):

    if img_array is None or img_array.size == 0:
        return False

    if img_array.ndim == 2:
        img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)
    elif img_array.shape[2] == 1:
        img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)
    elif img_array.shape[2] == 4:
        img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)

    gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(img_array, cv2.COLOR_RGB2HSV)
    saturation = hsv[:, :, 1]
    value = hsv[:, :, 2]

    color_ink = np.where(
        (saturation > 40) & (value > 35) & (value < 245),
        255,
        0,
    ).astype("uint8")
    if _mask_has_signature(color_ink, min_height_ratio=0.02):
        return True

    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    dark_ink = cv2.adaptiveThreshold(
        blur,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        21,
        8,
    )
    return _mask_has_signature(dark_ink, min_height_ratio=0.025)


def _mask_has_signature(ink, min_height_ratio):
    # Join nearby pen strokes without merging whole table rows.
    close_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (23, 7))
    joined = cv2.morphologyEx(ink, cv2.MORPH_CLOSE, close_kernel)
    contours, _ = cv2.findContours(
        joined,
        cv2.RETR_EXTERNAL,
        cv2.CHAIN_APPROX_SIMPLE,
    )

    page_height, page_width = ink.shape
    page_area = page_height * page_width

    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w <= 0 or h <= 0:
            continue

        box_area = w * h
        ratio = w / float(h)
        ink_density = cv2.countNonZero(ink[y:y + h, x:x + w]) / float(box_area)
        contour_area = cv2.contourArea(contour)
        fill_ratio = contour_area / float(box_area)

        width_ok = page_width * 0.018 <= w <= page_width * 0.45
        height_ok = page_height * min_height_ratio <= h <= page_height * 0.13
        area_ok = page_area * 0.00004 <= box_area <= page_area * 0.04
        shape_ok = 1.4 <= ratio <= 14.0
        ink_ok = 0.015 <= ink_density <= 0.45
        contour_ok = fill_ratio <= 0.75

        if all((width_ok, height_ok, area_ok, shape_ok, ink_ok, contour_ok)):
            return True

    return False


def document_has_valid_signature(pdf_doc):
    if not pdf_doc:
        return False

    for page_index in range(len(pdf_doc)):
        page = pdf_doc.load_page(page_index)
        image = page_to_rgb_array(page)
        if detect_signature_in_image(image):
            return True
    return False
