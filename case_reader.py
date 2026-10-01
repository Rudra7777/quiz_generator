"""
Case Reader

Pulls the text out of a case study handed in as a PDF or DOCX, so it can be printed
above the questions on every set. Text only — images, tables and page layout are
dropped, since a question paper is an Excel sheet and cells hold text.
"""

import io
import re
from pathlib import Path
from typing import List

from docx import Document
from pypdf import PdfReader


CASE_EXTENSIONS = (".pdf", ".docx")


def _pdf_paragraphs(data: bytes) -> List[str]:
    # A PDF stores visual lines, not paragraphs. Layout mode keeps the vertical gap
    # between paragraphs as a blank line (plain mode drops it), so split on those and
    # re-join the lines within each paragraph.
    paragraphs = []
    for page in PdfReader(io.BytesIO(data)).pages:
        text = page.extract_text(extraction_mode="layout") or ""
        for block in re.split(r"\n\s*\n", text):
            joined = ""
            for line in (line.strip() for line in block.splitlines()):
                if not line:
                    continue
                # A word hyphenated across the line break stays one word: 'electric-bike'.
                separator = "" if not joined or joined.endswith("-") else " "
                joined += separator + line
            if joined:
                paragraphs.append(joined)
    return paragraphs


def _docx_paragraphs(data: bytes) -> List[str]:
    return [p.text.strip() for p in Document(io.BytesIO(data)).paragraphs if p.text.strip()]


def read_case_paragraphs(filename: str, data: bytes) -> List[str]:
    """Return the case's paragraphs, in order, from a .pdf or .docx file's bytes."""
    extension = Path(filename).suffix.lower()
    if extension == ".pdf":
        paragraphs = _pdf_paragraphs(data)
    elif extension == ".docx":
        paragraphs = _docx_paragraphs(data)
    else:
        raise ValueError(f"Case must be a PDF or DOCX file, not '{extension or filename}'.")

    if not paragraphs:
        # Usually a scanned PDF: pages are pictures, so there is no text to pull out.
        raise ValueError("No text found in the case file. Is it a scanned PDF?")
    return paragraphs
