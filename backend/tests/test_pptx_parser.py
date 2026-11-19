import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

from app.ingestion.pptx_parser import PptxParser


class PptxParserTests(unittest.TestCase):
    def test_extracts_slide_text_tables_and_locations(self):
        presentation = Presentation()
        slide1 = presentation.slides.add_slide(presentation.slide_layouts[6])
        box = slide1.shapes.add_textbox(0, 0, 5000000, 500000)
        box.text = "Title: Architecture"
        table_shape = slide1.shapes.add_table(2, 2, 0, 1000000, 4000000, 1000000)
        table_shape.table.cell(0, 0).text = "Component"
        table_shape.table.cell(0, 1).text = "Role"
        table_shape.table.cell(1, 0).text = "API"
        table_shape.table.cell(1, 1).text = "Entry point"

        slide2 = presentation.slides.add_slide(presentation.slide_layouts[6])
        box2 = slide2.shapes.add_textbox(0, 0, 5000000, 500000)
        box2.text = "Slide two content"

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "sample.pptx"
            presentation.save(path)
            document = PptxParser().parse(str(path))

        self.assertEqual(document.source_type, "pptx")
        self.assertEqual(len(document.pages), 2)
        self.assertEqual(document.pages[0].location, "Slide 1")
        self.assertIn("Architecture", document.pages[0].text)
        self.assertIn("API | Entry point", document.pages[0].text)
        self.assertEqual(document.pages[1].location, "Slide 2")
        self.assertIn("Slide two content", document.pages[1].text)


if __name__ == "__main__":
    unittest.main()
