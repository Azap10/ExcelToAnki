"""Helpers for recognizing printed Chinese and English text in PDF pages.

This module intentionally stops after recognition. It does not alter input PDFs
or create a text overlay yet.
"""

from __future__ import annotations

import json
import os
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class OcrSpan:
    """One recognized text region, with coordinates in image pixels and PDF points."""

    text: str
    recognition_confidence: float
    detection_confidence: float | None
    language: str
    image_polygon: tuple[tuple[float, float], ...]
    pdf_polygon: tuple[tuple[float, float], ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "text": self.text,
            "recognition_confidence": self.recognition_confidence,
            "detection_confidence": self.detection_confidence,
            "language": self.language,
            "image_polygon": [list(point) for point in self.image_polygon],
            "pdf_polygon": [list(point) for point in self.pdf_polygon],
        }


@dataclass(frozen=True)
class PageOcrResult:
    """OCR data for one page, suitable for JSON caching or later PDF overlay."""

    page_index: int
    dpi: int
    image_width: int
    image_height: int
    page_width: float
    page_height: float
    spans: tuple[OcrSpan, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "page_index": self.page_index,
            "dpi": self.dpi,
            "image_width": self.image_width,
            "image_height": self.image_height,
            "page_width": self.page_width,
            "page_height": self.page_height,
            "spans": [span.to_dict() for span in self.spans],
        }


@dataclass(frozen=True)
class RenderedPage:
    """A rendered PDF page and the geometry required to map OCR results back to it."""

    page_index: int
    dpi: int
    image_bgr: Any
    image_width: int
    image_height: int
    page_width: float
    page_height: float
    page_origin_x: float
    page_origin_y: float
    point_scale: float


def open_pdf(pdf_path: str | Path) -> Any:
    """Open a PDF for page-by-page OCR; the caller is responsible for closing it."""
    path = Path(pdf_path)
    if not path.is_file():
        raise FileNotFoundError(f"PDF not found: {path}")

    pymupdf = _load_pymupdf()
    document = pymupdf.open(path)
    if not document.is_pdf:
        document.close()
        raise ValueError(f"The input is not a PDF: {path}")
    return document


def create_ocr_engine() -> Any:
    """Create a reusable PaddleOCR engine for printed Chinese and English text.

    Construct this once and pass it to ``process_pdf_page`` for every page. The
    current PaddleOCR 3.x API is used first, with a compatibility fallback for
    older PaddleOCR installations.
    """
    # PaddleX now defaults to Hugging Face for model downloads. BOS is a more
    # reliable source for this Windows desktop workflow; callers can override
    # it by setting PADDLE_PDX_MODEL_SOURCE before launching the app.
    os.environ.setdefault("PADDLE_PDX_MODEL_SOURCE", "BOS")
    try:
        from paddleocr import PaddleOCR
    except ImportError as error:
        raise RuntimeError(
            "PaddleOCR is not installed. Install paddlepaddle and paddleocr before using OCR helpers."
        ) from error

    try:
        return PaddleOCR(
            lang="ch",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
        )
    except TypeError:
        return PaddleOCR(lang="ch", use_angle_cls=True)


def render_pdf_page(document: Any, page_index: int, dpi: int = 225) -> RenderedPage:
    """Render one PDF page to a BGR image for PaddleOCR.

    ``dpi=225`` balances small-font readability and responsiveness. Increase to
    300 when a page's text is particularly small or low-confidence.
    """
    if dpi <= 0:
        raise ValueError("DPI must be greater than zero.")
    if page_index < 0 or page_index >= document.page_count:
        raise IndexError(f"Page index {page_index} is outside this PDF's page range.")

    try:
        import numpy as np
    except ImportError as error:
        raise RuntimeError("NumPy is required to render PDF pages for PaddleOCR.") from error

    pymupdf = _load_pymupdf()
    page = document.load_page(page_index)
    point_scale = dpi / 72
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(point_scale, point_scale), alpha=False)
    image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, pixmap.n)

    if pixmap.n == 1:
        image = np.repeat(image, 3, axis=2)
    else:
        image = image[:, :, :3]
    image_bgr = image[:, :, ::-1].copy()

    return RenderedPage(
        page_index=page_index,
        dpi=dpi,
        image_bgr=image_bgr,
        image_width=pixmap.width,
        image_height=pixmap.height,
        page_width=page.rect.width,
        page_height=page.rect.height,
        page_origin_x=page.rect.x0,
        page_origin_y=page.rect.y0,
        point_scale=point_scale,
    )


def process_pdf_page(
    document: Any,
    page_index: int,
    ocr_engine: Any,
    dpi: int = 225,
    minimum_confidence: float = 0.0,
) -> PageOcrResult:
    """Recognize one PDF page and return normalized, JSON-ready OCR data."""
    if not 0 <= minimum_confidence <= 1:
        raise ValueError("minimum_confidence must be between 0 and 1.")

    rendered_page = render_pdf_page(document, page_index, dpi=dpi)
    spans = recognize_page_text(rendered_page, ocr_engine, minimum_confidence)
    return PageOcrResult(
        page_index=page_index,
        dpi=dpi,
        image_width=rendered_page.image_width,
        image_height=rendered_page.image_height,
        page_width=rendered_page.page_width,
        page_height=rendered_page.page_height,
        spans=tuple(spans),
    )


def process_pdf_region(
    document: Any,
    page_index: int,
    ocr_engine: Any,
    pdf_rect: tuple[float, float, float, float],
    dpi: int = 225,
    minimum_confidence: float = 0.0,
) -> PageOcrResult:
    """Recognize a rectangular region of one PDF page.

    ``pdf_rect`` contains left, top, right, and bottom coordinates in PDF
    points. The page is rendered at OCR resolution before cropping, so the
    screen preview's resolution does not limit recognition quality.
    """
    if not 0 <= minimum_confidence <= 1:
        raise ValueError("minimum_confidence must be between 0 and 1.")

    rendered_page = render_pdf_page(document, page_index, dpi=dpi)
    left, top, right, bottom = pdf_rect
    left, right = sorted((left, right))
    top, bottom = sorted((top, bottom))
    page_left = rendered_page.page_origin_x
    page_top = rendered_page.page_origin_y
    page_right = page_left + rendered_page.page_width
    page_bottom = page_top + rendered_page.page_height
    left = max(page_left, left)
    top = max(page_top, top)
    right = min(page_right, right)
    bottom = min(page_bottom, bottom)
    if right <= left or bottom <= top:
        raise ValueError("The selected OCR region is outside the PDF page.")

    x0 = max(0, round((left - page_left) * rendered_page.point_scale))
    y0 = max(0, round((top - page_top) * rendered_page.point_scale))
    x1 = min(rendered_page.image_width, round((right - page_left) * rendered_page.point_scale))
    y1 = min(rendered_page.image_height, round((bottom - page_top) * rendered_page.point_scale))
    if x1 <= x0 or y1 <= y0:
        raise ValueError("The selected OCR region is too small to process.")

    region = RenderedPage(
        page_index=page_index,
        dpi=dpi,
        image_bgr=rendered_page.image_bgr[y0:y1, x0:x1].copy(),
        image_width=x1 - x0,
        image_height=y1 - y0,
        page_width=(x1 - x0) / rendered_page.point_scale,
        page_height=(y1 - y0) / rendered_page.point_scale,
        page_origin_x=page_left + x0 / rendered_page.point_scale,
        page_origin_y=page_top + y0 / rendered_page.point_scale,
        point_scale=rendered_page.point_scale,
    )
    spans = recognize_page_text(region, ocr_engine, minimum_confidence)
    return PageOcrResult(
        page_index=page_index,
        dpi=dpi,
        image_width=region.image_width,
        image_height=region.image_height,
        page_width=region.page_width,
        page_height=region.page_height,
        spans=tuple(spans),
    )


def recognize_page_text(
    rendered_page: RenderedPage,
    ocr_engine: Any,
    minimum_confidence: float = 0.0,
) -> list[OcrSpan]:
    """Run PaddleOCR on a rendered page and normalize its results into spans."""
    if hasattr(ocr_engine, "predict"):
        predictions = ocr_engine.predict(rendered_page.image_bgr)
        return _spans_from_v3_predictions(rendered_page, predictions, minimum_confidence)
    if hasattr(ocr_engine, "ocr"):
        predictions = ocr_engine.ocr(rendered_page.image_bgr, cls=True)
        return _spans_from_legacy_predictions(rendered_page, predictions, minimum_confidence)
    raise TypeError("ocr_engine must provide PaddleOCR's predict() or ocr() method.")


def save_page_ocr_result(result: PageOcrResult, output_path: str | Path) -> Path:
    """Save one page's OCR result as UTF-8 JSON for fast reuse later."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(result.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def classify_text_language(text: str) -> str:
    """Classify a recognized span as Chinese, English, mixed, or unknown."""
    has_chinese = any("\u3400" <= character <= "\u9fff" for character in text)
    has_english = any(("a" <= character.lower() <= "z") for character in text)
    if has_chinese and has_english:
        return "mixed"
    if has_chinese:
        return "chinese"
    if has_english:
        return "english"
    return "unknown"


def _load_pymupdf() -> Any:
    try:
        import pymupdf
    except ImportError as error:
        raise RuntimeError("PyMuPDF is not installed. Install pymupdf before using OCR helpers.") from error
    return pymupdf


def _spans_from_v3_predictions(
    rendered_page: RenderedPage,
    predictions: Any,
    minimum_confidence: float,
) -> list[OcrSpan]:
    spans: list[OcrSpan] = []
    for prediction in predictions:
        payload = _prediction_payload(prediction)
        texts = list(payload.get("rec_texts", []))
        scores = list(payload.get("rec_scores", []))
        polygons = list(payload.get("rec_polys", []))
        if not polygons:
            polygons = list(payload.get("rec_boxes", []))
        detection_scores = list(payload.get("dt_scores", []))

        for index, text in enumerate(texts):
            if index >= len(scores) or index >= len(polygons):
                continue
            confidence = float(scores[index])
            if confidence < minimum_confidence:
                continue
            detection_confidence = (
                float(detection_scores[index]) if len(detection_scores) == len(texts) else None
            )
            spans.append(
                _build_span(
                    str(text),
                    confidence,
                    detection_confidence,
                    polygons[index],
                    rendered_page,
                )
            )
    return spans


def _spans_from_legacy_predictions(
    rendered_page: RenderedPage,
    predictions: Any,
    minimum_confidence: float,
) -> list[OcrSpan]:
    spans: list[OcrSpan] = []
    page_predictions = predictions[0] if predictions else []
    for prediction in page_predictions or []:
        polygon, (text, confidence) = prediction
        confidence = float(confidence)
        if confidence >= minimum_confidence:
            spans.append(_build_span(str(text), confidence, None, polygon, rendered_page))
    return spans


def _prediction_payload(prediction: Any) -> Mapping[str, Any]:
    payload = prediction
    if not isinstance(payload, Mapping) and hasattr(payload, "json"):
        payload = payload.json
        if callable(payload):
            payload = payload()
    if isinstance(payload, str):
        payload = json.loads(payload)
    if not isinstance(payload, Mapping):
        raise TypeError("PaddleOCR returned an unsupported prediction format.")
    result = payload.get("res", payload)
    if not isinstance(result, Mapping):
        raise TypeError("PaddleOCR prediction did not contain a result mapping.")
    return result


def _build_span(
    text: str,
    recognition_confidence: float,
    detection_confidence: float | None,
    polygon: Any,
    rendered_page: RenderedPage,
) -> OcrSpan:
    image_polygon = _normalize_polygon(polygon)
    pdf_polygon = tuple(
        (
            rendered_page.page_origin_x + x / rendered_page.point_scale,
            rendered_page.page_origin_y + y / rendered_page.point_scale,
        )
        for x, y in image_polygon
    )
    return OcrSpan(
        text=text,
        recognition_confidence=recognition_confidence,
        detection_confidence=detection_confidence,
        language=classify_text_language(text),
        image_polygon=image_polygon,
        pdf_polygon=pdf_polygon,
    )


def _normalize_polygon(polygon: Any) -> tuple[tuple[float, float], ...]:
    points = list(polygon)
    if len(points) == 4 and all(not hasattr(point, "__len__") for point in points):
        left, top, right, bottom = (float(value) for value in points)
        return ((left, top), (right, top), (right, bottom), (left, bottom))
    normalized = tuple((float(point[0]), float(point[1])) for point in points)
    if len(normalized) < 4:
        raise ValueError("An OCR polygon must contain at least four points.")
    return normalized
