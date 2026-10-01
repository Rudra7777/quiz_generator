"""Tests for pulling a case study's text out of a PDF or DOCX."""

import io

import pytest
from docx import Document

from case_reader import read_case_paragraphs


def _docx(*paragraphs: str) -> bytes:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def _pdf(*lines: str) -> bytes:
    """A minimal one-page PDF with each line drawn on its own row, 16pt apart."""
    text_ops = "".join(f"({line}) Tj 0 -16 Td " for line in lines)
    stream = f"BT /F1 12 Tf 72 720 Td {text_ops}ET".encode()
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        b"/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
    ]
    out, offsets = b"%PDF-1.4\n", []
    for number, body in enumerate(objects, 1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % number + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offsets)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (
        len(objects) + 1,
        xref,
    )
    return out


def test_docx_paragraphs_come_back_in_order_without_blanks():
    data = _docx("First paragraph.", "", "  Second paragraph.  ")

    assert read_case_paragraphs("case.docx", data) == ["First paragraph.", "Second paragraph."]


def test_pdf_lines_of_one_paragraph_are_joined():
    data = _pdf("The retailer runs", "twelve stores.")

    assert read_case_paragraphs("Case.PDF", data) == ["The retailer runs twelve stores."]


def test_pdf_paragraph_gap_starts_a_new_paragraph():
    # A blank line in the drawing is the vertical gap between two paragraphs.
    data = _pdf("The retailer runs", "twelve stores.", "", "Costs are rising.")

    assert read_case_paragraphs("case.pdf", data) == [
        "The retailer runs twelve stores.",
        "Costs are rising.",
    ]


def test_pdf_word_hyphenated_across_lines_is_rejoined():
    data = _pdf("plans a new electric-", "bike line.")

    assert read_case_paragraphs("case.pdf", data) == ["plans a new electric-bike line."]


def test_other_file_types_are_refused():
    with pytest.raises(ValueError, match="PDF or DOCX"):
        read_case_paragraphs("case.txt", b"hello")


def test_a_file_with_no_text_is_refused():
    with pytest.raises(ValueError, match="No text"):
        read_case_paragraphs("case.docx", _docx())
