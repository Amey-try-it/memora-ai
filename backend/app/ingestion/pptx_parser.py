"""Deterministic local PowerPoint extraction. No LLM/model is used for ingestion."""
import os
import uuid
from datetime import datetime
from io import BytesIO

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.chart import Chart

from app.models.source_document import SourceDocument, SourcePage
from app.ingestion.cleaner import TextCleaner

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None
    Image = None

TESSERACT_PATH = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
if pytesseract is not None:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


class PptxParser:
    """Extract slide text, tables, notes, charts, links, and image OCR locally."""

    def _ocr_image(self, image_blob):
        if pytesseract is None or Image is None:
            return ""
        try:
            image = Image.open(BytesIO(image_blob))
            return pytesseract.image_to_string(image, lang="eng").strip()
        except Exception:
            return ""

    @staticmethod
    def _shape_text(shape):
        output = []
        if getattr(shape, "has_text_frame", False):
            for paragraph in shape.text_frame.paragraphs:
                text = (paragraph.text or "").strip()
                if text:
                    output.append(text)
        if getattr(shape, "has_table", False):
            table = shape.table
            for row in table.rows:
                cells = [TextCleaner.clean(cell.text) for cell in row.cells]
                if any(cells):
                    output.append("TABLE: " + " | ".join(cells))
        if getattr(shape, "has_chart", False):
            output.extend(PptxParser._chart_text(shape.chart))
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            alt = getattr(shape, "alt_text", "") or getattr(shape, "name", "")
            if alt:
                output.append(f"Image description: {alt}")
            ocr = PptxParser()._ocr_image(shape.image.blob)
            if ocr:
                output.append("Image OCR: " + ocr)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                output.extend(PptxParser._shape_text(child))
        # Hyperlinks in text runs
        if getattr(shape, "has_text_frame", False):
            links = []
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    address = getattr(run.hyperlink, "address", None)
                    if address:
                        links.append(address)
            if links:
                output.extend("Hyperlink: " + link for link in dict.fromkeys(links))
        return output

    @staticmethod
    def _chart_text(chart):
        lines = ["Chart:"]
        try:
            if chart.has_title:
                lines.append("Title: " + chart.chart_title.text_frame.text)
        except Exception:
            pass
        try:
            for series in chart.series:
                name = getattr(series, "name", "Series") or "Series"
                lines.append(f"Series {name}:")
                values = list(series.values) if series.values is not None else []
                lines.append("Values: " + ", ".join(str(v) for v in values))
                try:
                    categories = [str(c.label) for c in chart.plots[0].categories]
                    if categories:
                        lines.append("Categories: " + ", ".join(categories))
                except Exception:
                    pass
        except Exception:
            lines.append("Chart data could not be read by python-pptx.")
        return lines

    def parse(self, filepath):
        presentation = Presentation(filepath)
        pages = []
        core = presentation.core_properties
        author = (core.author or "Unknown").strip() or "Unknown"
        created = core.created if core.created else datetime.now()
        if created.tzinfo is not None:
            created = created.replace(tzinfo=None)

        for slide_number, slide in enumerate(presentation.slides, start=1):
            parts = []
            for shape in slide.shapes:
                parts.extend(self._shape_text(shape))
            # Speaker notes are part of the source and should be indexed.
            try:
                if slide.has_notes_slide:
                    notes = slide.notes_slide.notes_text_frame.text
                    notes = TextCleaner.clean(notes)
                    if notes:
                        parts.append("Speaker Notes: " + notes)
            except Exception:
                pass
            clean = TextCleaner.clean("\\n".join(parts))
            if clean:
                pages.append(SourcePage(location=f"Slide {slide_number}", text=clean))

        return SourceDocument(
            id=str(uuid.uuid4()),
            source_type="pptx",
            source_name=os.path.basename(filepath),
            author=author,
            created_at=created,
            pages=pages,
        )
