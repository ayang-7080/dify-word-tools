import unittest
import base64
import zipfile
from io import BytesIO

from PIL import Image

from tools.tools import extract_docx_to_json, normalize_extracted_image


class ExtractedImageConversionTests(unittest.TestCase):
    def test_tiff_is_converted_to_png(self):
        image_stream = BytesIO()
        Image.new("RGB", (2, 2), "red").save(image_stream, format="TIFF")

        result = normalize_extracted_image(image_stream.getvalue(), "image1.tiff")

        self.assertEqual(result["mime_type"], "image/png")
        self.assertEqual(result["filename"], "image1.png")
        self.assertTrue(result["converted"])
        self.assertTrue(result["data"].startswith(b"\x89PNG\r\n\x1a\n"))

    def test_emf_is_not_reported_as_png_when_no_converter_is_available(self):
        result = normalize_extracted_image(b"\x01\x00\x00\x00", "image1.emf")

        self.assertEqual(result["mime_type"], "image/x-emf")
        self.assertEqual(result["filename"], "image1.emf")
        self.assertFalse(result["converted"])

    def test_docx_extraction_returns_converted_tiff_image(self):
        image_stream = BytesIO()
        Image.new("RGB", (2, 2), "blue").save(image_stream, format="TIFF")

        docx_stream = BytesIO()
        with zipfile.ZipFile(docx_stream, "w") as docx_zip:
            docx_zip.writestr(
                "word/document.xml",
                """<?xml version="1.0" encoding="UTF-8"?>
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                    <w:body>
                        <w:p>
                            <w:r>
                                <w:drawing>
                                    <wp:inline>
                                        <wp:extent cx="360000" cy="360000"/>
                                        <a:graphic>
                                            <a:graphicData>
                                                <a:blip r:embed="rId1"/>
                                            </a:graphicData>
                                        </a:graphic>
                                    </wp:inline>
                                </w:drawing>
                            </w:r>
                        </w:p>
                    </w:body>
                </w:document>
                """,
            )
            docx_zip.writestr(
                "word/_rels/document.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship
                        Id="rId1"
                        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                        Target="media/image1.tiff"/>
                </Relationships>
                """,
            )
            docx_zip.writestr("word/media/image1.tiff", image_stream.getvalue())

        docx_stream.seek(0)
        result = extract_docx_to_json(docx_stream)
        image = next(item for item in result if item["type"] == "image")

        self.assertEqual(image["mime_type"], "image/png")
        self.assertEqual(image["filename"], "image1.png")
        self.assertEqual(image["original_filename"], "image1.tiff")
        self.assertTrue(image["converted"])
        self.assertTrue(base64.b64decode(image["content"]).startswith(b"\x89PNG\r\n\x1a\n"))


if __name__ == "__main__":
    unittest.main()
