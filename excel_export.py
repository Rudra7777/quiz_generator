"""
Excel Export Helpers

UI-agnostic functions shared by the Streamlit (app.py) and Reflex (quiz_web) UIs:
- Building the styled multi-sheet question-papers workbook.
- Serializing a DataFrame to in-memory Excel bytes.
- Loading the embedded Question_Bank sheet back out of a question papers file.
"""

import io
import math
import os
import tempfile
from collections import Counter, defaultdict
from typing import List, Sequence, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

from answer_checker import _inject_cached_values
from excel_handler import load_question_bank, set_label, FullQuestionBank


# Sheet holding every set stacked one below another, for A4 portrait printing.
ALL_SETS_SHEET = "All_Sets"

# Column widths for the stacked print sheet. The visible ones total 92, which is
# what A4 portrait prints at full size — the per-set sheets stay wide (183) since
# they are read on screen, not printed.
ALL_SETS_WIDTHS = {'A': 4, 'B': 8, 'C': 8, 'D': 26, 'E': 13.5, 'F': 13.5, 'G': 13.5, 'H': 13.5}

# ── Set block styles ──────────────────────────────────────────────────────────
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
WRAP_ALIGN = Alignment(wrap_text=True, vertical='top')
CASE_ALIGN = Alignment(wrap_text=True, vertical='top', horizontal='justify')
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
BOLD_FONT = Font(bold=True)
QSET_FONT = Font(bold=True, color="FF0000")
QSET_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
QCD_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
CORRECT_OPTION_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
SUMMARY_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

# Question_Bank sheet layout: title on row 1, the filter-aware summary grid from
# QB_SUMMARY_ROW, and the bank table itself (with its filter buttons) from QB_HEADER_ROW.
QB_SUMMARY_ROW = 3
QB_HEADER_ROW = 11

# Per-set sheets are read on screen and kept wide; their visible columns (B hidden)
# total this many width units. All_Sets prints at the narrower ALL_SETS_WIDTHS.
SET_SHEET_WIDTHS = dict(zip('ABCDEFGH', (6, 8, 10, 55, 28, 28, 28, 28)))

# Excel will not auto-fit the height of a merged cell, so case rows are sized by
# estimate: roughly 1.1 characters of prose per column-width unit, 15pt per line.
# A row tops out at 409pt, so long paragraphs are split across several rows.
CASE_CHARS_PER_UNIT = 1.1
CASE_LINE_HEIGHT = 15
CASE_MAX_LINES_PER_ROW = 25


def _visible_width(widths: dict) -> float:
    """Total width of the columns a printed set shows — Col B is always hidden."""
    return sum(w for ch, w in widths.items() if ch != 'B')


def _case_rows(paragraphs: Sequence[str], width: float) -> List[Tuple[str, float]]:
    """Split case paragraphs into (text, row height) pairs for a block `width` units wide."""
    chars_per_line = max(1, int(width * CASE_CHARS_PER_UNIT))
    max_chars = chars_per_line * CASE_MAX_LINES_PER_ROW

    rows = []
    for paragraph in paragraphs:
        chunk = ""
        for word in paragraph.split():
            if chunk and len(chunk) + 1 + len(word) > max_chars:
                rows.append(chunk)
                chunk = word
            else:
                chunk = f"{chunk} {word}" if chunk else word
        if chunk:
            rows.append(chunk)

    return [
        (text, CASE_LINE_HEIGHT * math.ceil(len(text) / chars_per_line) + 4)
        for text in rows
    ]


def _write_set_block(
    ws,
    top_row: int,
    label: str,
    quiz: list,
    question_bank: FullQuestionBank,
    case_rows: Sequence[Tuple[str, float]] = (),
) -> int:
    """
    Draw one set's paper starting at `top_row`, returning the row after it.

    Used for both a set's own sheet (top_row=1) and its block on the stacked
    print sheet, so the printed paper is the same paper the sheet shows.

    `case_rows`, from `_case_rows`, puts a case study between the QSet line and
    the questions, under a 'Case Study' heading.
    """
    # Set label, then blank Name / Roll No fields to fill in by hand
    for col in (1, 2, 3):
        cell = ws.cell(row=top_row, column=col, value="QSet:" if col == 1 else label)
        cell.font = QSET_FONT
        cell.fill = QSET_FILL
        cell.border = THIN_BORDER
        cell.alignment = CENTER_ALIGN

    ws.merge_cells(start_row=top_row, start_column=4, end_row=top_row, end_column=6)
    ws.merge_cells(start_row=top_row, start_column=7, end_row=top_row, end_column=8)
    ws.cell(row=top_row, column=4, value="Name:").font = BOLD_FONT
    ws.cell(row=top_row, column=7, value="Roll No:").font = BOLD_FONT
    for col in range(4, 9):
        ws.cell(row=top_row, column=col).border = THIN_BORDER
        ws.cell(row=top_row, column=col).alignment = LEFT_ALIGN

    # Case study, one merged A:H row per paragraph chunk.
    row = top_row + 1
    if case_rows:
        heading = ws.cell(row=row, column=1, value="Case Study")
        heading.font = BOLD_FONT
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1
        for text, height in case_rows:
            ws.cell(row=row, column=1, value=text).alignment = CASE_ALIGN
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.row_dimensions[row].height = height
            row += 1

    # Headers. Two QCd columns — the bare Question Number and its printed
    # 'Q- 27' form, which is what students copy into the form.
    header_row = row
    for col, header in enumerate(['Sr', 'QCd', 'QCd', 'Question', 'A', 'B', 'C', 'D'], 1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = BOLD_FONT
        cell.border = THIN_BORDER
        cell.alignment = CENTER_ALIGN
    ws.cell(row=header_row, column=2).fill = QCD_FILL

    # Questions
    for q_idx, question_id in enumerate(quiz):
        q = question_bank.get_by_id(question_id)
        row = header_row + 1 + q_idx
        ws.cell(row=row, column=1, value=q_idx + 1)
        ws.cell(row=row, column=2, value=q.question_no).fill = QCD_FILL
        ws.cell(row=row, column=3, value=question_code(q.question_no))
        ws.cell(row=row, column=4, value=q.question_text)
        options = (q.option_a, q.option_b, q.option_c, q.option_d)
        for col, (letter, text) in enumerate(zip('ABCD', options), 5):
            ws.cell(row=row, column=col, value=f"{letter}] {text}")

        for col in range(1, 9):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.alignment = CENTER_ALIGN if col <= 3 else WRAP_ALIGN

    return header_row + 1 + len(quiz)


def _write_all_sets_sheet(
    wb, shuffled_matrix: list, question_bank: FullQuestionBank, case_paragraphs: Sequence[str] = ()
) -> None:
    """
    Write every set one below another on a single A4-portrait print sheet.

    Faculty print this rather than opening 65 separate sheets. Each set starts on
    a fresh page so the papers can be separated, and row heights are left unset so
    Excel fits each row to its own wrapped text.
    """
    ws = wb.create_sheet(title=ALL_SETS_SHEET, index=0)
    case_rows = _case_rows(case_paragraphs, _visible_width(ALL_SETS_WIDTHS))

    row = 1
    for student_idx, quiz in enumerate(shuffled_matrix):
        row = _write_set_block(
            ws, row, set_label(student_idx + 1), quiz, question_bank, case_rows
        )
        if student_idx < len(shuffled_matrix) - 1:
            ws.row_breaks.append(Break(id=row - 1))

    for ch, width in ALL_SETS_WIDTHS.items():
        ws.column_dimensions[ch].width = width
    # Col B carries the bare bank number only so the two sheets stay identical;
    # faculty print the 'Q- 27' form in Col C, so B is hidden here as well.
    ws.column_dimensions['B'].hidden = True

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.2, footer=0.2)


def _write_bank_summary(ws, question_bank: FullQuestionBank, last_bank_row: int) -> dict:
    """
    Write a difficulty x correct-option count grid above the Question_Bank table.

    Every count goes through SUBTOTAL(103, ...), which skips rows a filter has
    hidden, so filtering the bank below (say, to Hard questions, or to one topic)
    updates the grid to describe just the rows on screen. It sits above the table
    rather than beside it because a filter hides whole rows, and would hide the
    grid with them.

    The A-D counts sit in columns C-F, directly over option_a-option_d, and the
    Total in G, over the answer column. Returns each formula's current value,
    keyed by cell, for writing in as a cached result.
    """
    first, last = QB_HEADER_ROW + 1, last_bank_row
    visible = f"SUBTOTAL(103,OFFSET($H${first},ROW($H${first}:$H${last})-ROW($H${first}),0))"
    levels = ["Hard", "Medium", "Easy"]
    header_row = QB_SUMMARY_ROW + 1
    total_row = header_row + len(levels) + 1
    pct_row = total_row + 1

    title = ws.cell(row=QB_SUMMARY_ROW, column=2, value="Summary — counts only the rows the filter shows")
    title.font = Font(bold=True, italic=True)

    for col, label in enumerate(["Difficulty / Correct option", "A", "B", "C", "D", "Total"], 2):
        cell = ws.cell(row=header_row, column=col, value=label)
        cell.font = Font(bold=True)
        cell.fill = SUMMARY_FILL
        cell.border = THIN_BORDER
        cell.alignment = LEFT_ALIGN if col == 2 else CENTER_ALIGN

    by_level = Counter(
        (q.difficulty.capitalize(), str(q.answer).strip().upper()) for q in question_bank.get_all()
    )
    cached = {}
    for offset, level in enumerate(levels + ["Total", "% of total"]):
        row = header_row + 1 + offset
        ws.cell(row=row, column=2, value=level).font = Font(bold=True)
        for col in range(2, 8):
            ws.cell(row=row, column=col).border = THIN_BORDER
        for col_offset, letter in enumerate("ABCD"):
            col = 3 + col_offset
            ref = f"{get_column_letter(col)}{row}"
            col_letter = get_column_letter(col)
            if level in levels:
                ws[ref] = (
                    f"=SUMPRODUCT({visible},--($H${first}:$H${last}=$B{row}),"
                    f"--($G${first}:$G${last}={col_letter}${header_row}))"
                )
                cached[ref] = by_level[(level, letter)]
            elif level == "Total":
                ws[ref] = f"=SUM({col_letter}{header_row + 1}:{col_letter}{header_row + len(levels)})"
                cached[ref] = sum(by_level[(lv, letter)] for lv in levels)
            else:
                ws[ref] = f"=IF($G${total_row}=0,0,{col_letter}{total_row}/$G${total_row})"
                ws[ref].number_format = "0%"
            ws[ref].alignment = CENTER_ALIGN

        total_ref = f"G{row}"
        if level == "% of total":
            ws[total_ref] = f"=SUM(C{row}:F{row})"
            ws[total_ref].number_format = "0%"
        else:
            ws[total_ref] = f"=SUM(C{row}:F{row})"
            cached[total_ref] = sum(cached[f"{get_column_letter(c)}{row}"] for c in range(3, 7))
        ws[total_ref].font = Font(bold=True)
        ws[total_ref].alignment = CENTER_ALIGN

    grand_total = cached[f"G{total_row}"]
    for col in range(3, 7):
        share = cached[f"{get_column_letter(col)}{total_row}"] / grand_total if grand_total else 0
        cached[f"{get_column_letter(col)}{pct_row}"] = share
    cached[f"G{pct_row}"] = 1 if grand_total else 0
    return cached


def qid_to_number(question_id: str, question_bank: FullQuestionBank) -> int:
    """Convert internal question_id (e.g., H1, M5) to original question_no."""
    q = question_bank.get_by_id(question_id)
    return q.question_no if q else 0


def question_code(question_no: int) -> str:
    """Printed form of a Question Number, e.g. 27 -> 'Q- 27'."""
    return f"Q- {question_no:02d}"


def create_formatted_excel(
    allocation_matrix: list,
    shuffled_matrix: list,
    usage_counts: dict,
    question_bank: FullQuestionBank,
    include_answer_key: bool = True,
    case_paragraphs: Sequence[str] = (),
) -> bytes:
    """
    Create a formatted Excel file with:
    - Question papers (one sheet per student)
    - Answer Key
    - Allocation Table (original order, numeric IDs)
    - Shuffled Table (shuffled order, numeric IDs)
    - Evaluation Table (min/max/delta stats)

    `case_paragraphs`, if given (see case_reader), is printed above the questions
    on every set.
    """
    wb = Workbook()

    # ── Shared styles ─────────────────────────────────────────────────────
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=11, color="FFFFFF")
    green_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    orange_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    thin_border = THIN_BORDER
    wrap_align = WRAP_ALIGN
    center_align = CENTER_ALIGN

    # Remove default sheet
    wb.remove(wb.active)

    # ══════════════════════════════════════════════════════════════════════
    # Question Paper Sheets (one per student)
    # ══════════════════════════════════════════════════════════════════════
    case_rows = _case_rows(case_paragraphs, _visible_width(SET_SHEET_WIDTHS))
    for student_idx, quiz in enumerate(shuffled_matrix):
        label = set_label(student_idx + 1)
        ws = wb.create_sheet(title=label)
        end_row = _write_set_block(ws, 1, label, quiz, question_bank, case_rows)

        # Column widths & row heights
        for ch, width in SET_SHEET_WIDTHS.items():
            ws.column_dimensions[ch].width = width
        # Col B carries the bare bank number for the answer-checker to read back
        # (see response_generator._attach_bank_no); faculty only want to see the
        # printed 'Q- 27' form in Col C, so hide B from view and from printouts.
        ws.column_dimensions['B'].hidden = True
        for r in range(end_row - len(quiz), end_row):
            ws.row_dimensions[r].height = 45

    # ══════════════════════════════════════════════════════════════════════
    # Answer Key Sheet
    # ══════════════════════════════════════════════════════════════════════
    if include_answer_key:
        ws = wb.create_sheet(title="Answer_Key")
        ws['A1'] = "ANSWER KEY (For Teachers Only)"
        ws['A1'].font = Font(bold=True, size=16, color="FF0000")
        num_q = len(shuffled_matrix[0])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_q + 1)

        headers = ['Set'] + [f'Q{i+1}' for i in range(num_q)]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=h)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.border = thin_border

        for student_idx, quiz in enumerate(shuffled_matrix):
            row = student_idx + 4
            ws.cell(row=row, column=1, value=set_label(student_idx + 1)).border = thin_border
            for q_idx, qid in enumerate(quiz):
                q = question_bank.get_by_id(qid)
                ws.cell(row=row, column=q_idx + 2, value=q.answer).border = thin_border

    # ══════════════════════════════════════════════════════════════════════
    # Allocation Table Sheet (original order, numeric question numbers)
    # ══════════════════════════════════════════════════════════════════════
    ws = wb.create_sheet(title="Allocation_Table")
    ws['A1'] = "Allocation Table (Original Order by Difficulty)"
    ws['A1'].font = Font(bold=True, size=14)
    num_students = len(allocation_matrix)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_students + 1)

    # Headers
    cell = ws.cell(row=3, column=1, value="Position")
    cell.font = header_font_white
    cell.fill = header_fill
    cell.border = thin_border
    for s_idx in range(num_students):
        cell = ws.cell(row=3, column=s_idx + 2, value=set_label(s_idx + 1))
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = thin_border

    # Data (using original question_no, not H/M/E IDs)
    num_positions = len(allocation_matrix[0])
    for pos in range(num_positions):
        cell = ws.cell(row=pos + 4, column=1, value=f"Q{pos + 1}")
        cell.border = thin_border
        cell.font = Font(bold=True)
        for s_idx in range(num_students):
            qid = allocation_matrix[s_idx][pos]
            cell = ws.cell(row=pos + 4, column=s_idx + 2, value=qid_to_number(qid, question_bank))
            cell.border = thin_border
            cell.alignment = center_align

    ws.column_dimensions['A'].width = 10

    # ══════════════════════════════════════════════════════════════════════
    # Shuffled Table Sheet (shuffled order, numeric question numbers)
    # ══════════════════════════════════════════════════════════════════════
    ws = wb.create_sheet(title="Shuffled_Table")
    ws['A1'] = "Shuffled Table (Randomized Order per Student)"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_students + 1)

    # Headers
    cell = ws.cell(row=3, column=1, value="Position")
    cell.font = header_font_white
    cell.fill = green_fill
    cell.border = thin_border
    for s_idx in range(num_students):
        cell = ws.cell(row=3, column=s_idx + 2, value=set_label(s_idx + 1))
        cell.font = header_font_white
        cell.fill = green_fill
        cell.border = thin_border

    # Data
    for pos in range(num_positions):
        cell = ws.cell(row=pos + 4, column=1, value=f"Q{pos + 1}")
        cell.border = thin_border
        cell.font = Font(bold=True)
        for s_idx in range(num_students):
            qid = shuffled_matrix[s_idx][pos]
            cell = ws.cell(row=pos + 4, column=s_idx + 2, value=qid_to_number(qid, question_bank))
            cell.border = thin_border
            cell.alignment = center_align

    ws.column_dimensions['A'].width = 10

    # ══════════════════════════════════════════════════════════════════════
    # Evaluation Table Sheet
    # ══════════════════════════════════════════════════════════════════════
    ws = wb.create_sheet(title="Evaluation")
    ws['A1'] = "Evaluation Summary"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:E1')

    # ── Question Usage Table ──
    ws['A3'] = "Question Usage"
    ws['A3'].font = Font(bold=True, size=12)

    for col, h in enumerate(['Question No', 'Internal ID', 'Difficulty', 'Usage Count'], 1):
        cell = ws.cell(row=4, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = thin_border

    row = 5
    for q in question_bank.get_all():
        count = usage_counts.get(q.question_id, 0)
        ws.cell(row=row, column=1, value=q.question_no).border = thin_border
        ws.cell(row=row, column=1).alignment = center_align
        ws.cell(row=row, column=2, value=q.question_id).border = thin_border
        ws.cell(row=row, column=3, value=q.difficulty.capitalize()).border = thin_border
        ws.cell(row=row, column=4, value=count).border = thin_border
        ws.cell(row=row, column=4).alignment = center_align
        row += 1

    # ── Min/Max/Delta by Difficulty ──
    row += 1
    ws.cell(row=row, column=1, value="Min / Max / Delta by Difficulty").font = Font(bold=True, size=12)
    row += 1

    for col, h in enumerate(['Difficulty', 'Min', 'Max', 'Delta', 'Variance'], 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = header_font_white
        cell.fill = orange_fill
        cell.border = thin_border
    row += 1

    by_diff = defaultdict(list)
    for q in question_bank.get_all():
        by_diff[q.difficulty].append(usage_counts.get(q.question_id, 0))

    all_counts = list(usage_counts.values())

    # Track min/max per difficulty for overall calculation
    diff_stats = []

    for diff in ['hard', 'medium', 'easy']:
        counts = by_diff.get(diff, [])
        if counts:
            mn, mx = min(counts), max(counts)
            var = round(float(np.var(counts)), 4)
        else:
            mn = mx = 0
            var = 0.0

        diff_stats.append((mn, mx))

        ws.cell(row=row, column=1, value=diff.capitalize()).border = thin_border
        ws.cell(row=row, column=2, value=mn).border = thin_border
        ws.cell(row=row, column=3, value=mx).border = thin_border
        ws.cell(row=row, column=4, value=mx - mn).border = thin_border
        ws.cell(row=row, column=5, value=var).border = thin_border
        row += 1

    # Overall row: sum of min/max from each difficulty
    overall_min = sum(mn for mn, mx in diff_stats)
    overall_max = sum(mx for mn, mx in diff_stats)
    overall_delta = overall_max - overall_min

    ws.cell(row=row, column=1, value="OVERALL").border = thin_border
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2, value=overall_min).border = thin_border
    ws.cell(row=row, column=3, value=overall_max).border = thin_border
    ws.cell(row=row, column=4, value=overall_delta).border = thin_border
    ws.cell(row=row, column=5, value="-").border = thin_border

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 12

    # ══════════════════════════════════════════════════════════════════════
    # Question Bank Sheet (embedded for Part 2 answer checking)
    # ══════════════════════════════════════════════════════════════════════
    ws = wb.create_sheet(title="Question_Bank")
    ws['A1'] = "Question Bank (Embedded for Answer Checking)"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:H1')

    qb_headers = ['question_no', 'question', 'option_a', 'option_b',
                   'option_c', 'option_d', 'answer', 'difficulty']
    qb_fill = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    for col, h in enumerate(qb_headers, 1):
        cell = ws.cell(row=QB_HEADER_ROW, column=col, value=h)
        cell.font = header_font_white
        cell.fill = qb_fill
        cell.border = thin_border

    for q_idx, q in enumerate(question_bank.get_all()):
        row = QB_HEADER_ROW + 1 + q_idx
        ws.cell(row=row, column=1, value=q.question_no).border = thin_border
        ws.cell(row=row, column=1).alignment = center_align
        ws.cell(row=row, column=2, value=q.question_text).border = thin_border
        ws.cell(row=row, column=2).alignment = wrap_align
        ws.cell(row=row, column=3, value=q.option_a).border = thin_border
        ws.cell(row=row, column=4, value=q.option_b).border = thin_border
        ws.cell(row=row, column=5, value=q.option_c).border = thin_border
        ws.cell(row=row, column=6, value=q.option_d).border = thin_border
        ws.cell(row=row, column=7, value=q.answer).border = thin_border
        ws.cell(row=row, column=7).alignment = center_align
        ws.cell(row=row, column=8, value=q.difficulty.capitalize()).border = thin_border

        # Highlight the correct option so faculty can check the key at a glance.
        answer = str(q.answer).strip().upper()
        if answer in 'ABCD':
            ws.cell(row=row, column=3 + 'ABCD'.index(answer)).fill = CORRECT_OPTION_FILL

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 50
    for ch in 'CDEF':
        ws.column_dimensions[ch].width = 20
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 12

    last_bank_row = QB_HEADER_ROW + len(question_bank.get_all())
    ws.auto_filter.ref = f"A{QB_HEADER_ROW}:H{last_bank_row}"
    bank_summary_cached = _write_bank_summary(ws, question_bank, last_bank_row)

    # ══════════════════════════════════════════════════════════════════════
    # Combined Print Sheet (every set stacked, A4 portrait)
    # ══════════════════════════════════════════════════════════════════════
    _write_all_sets_sheet(wb, shuffled_matrix, question_bank, case_paragraphs)

    # ── Save ──────────────────────────────────────────────────────────────
    output = io.BytesIO()
    wb.save(output)
    # Give the summary formulas a result to show before Excel recalculates.
    _inject_cached_values(output, "Question_Bank", bank_summary_cached)
    output.seek(0)
    return output.getvalue()


def _make_excel_bytes_from_dataframe(df: pd.DataFrame, sheet_name: str) -> bytes:
    """Serialize DataFrame into an in-memory Excel file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()


def _load_question_bank_from_question_papers(question_papers_path: str) -> FullQuestionBank:
    """Load embedded Question_Bank sheet from question_papers.xlsx."""
    required_cols = [
        "question_no",
        "question",
        "option_a",
        "option_b",
        "option_c",
        "option_d",
        "answer",
        "difficulty",
    ]
    normalized_required = set(required_cols)

    def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        out.columns = [str(c).strip().lower().replace(" ", "_") for c in out.columns]
        return out

    temp_path = None
    try:
        try:
            # The header row moves between formats — first row on a plain table,
            # below the title and summary grid on a styled one — so find it first.
            raw = pd.read_excel(question_papers_path, sheet_name="Question_Bank", header=None)
            question_bank_df = None
            for header_row, values in raw.iterrows():
                found = {str(v).strip().lower().replace(" ", "_") for v in values if pd.notna(v)}
                if normalized_required.issubset(found):
                    candidate = pd.read_excel(
                        question_papers_path, sheet_name="Question_Bank", header=header_row
                    )
                    candidate = _normalize_cols(candidate)
                    question_bank_df = candidate[required_cols].copy()
                    question_bank_df = question_bank_df.dropna(how="all")
                    break

            if question_bank_df is None:
                raise ValueError(
                    f"Missing required columns: {required_cols}"
                )
        except ValueError as exc:
            if "Worksheet named 'Question_Bank' not found" in str(exc):
                raise ValueError(
                    "Question_Bank sheet not found in question papers. "
                    "Regenerate papers with the latest app/CLI."
                ) from exc
            raise ValueError(
                "Question_Bank sheet is present but not in expected format. "
                "Regenerate papers with the latest app/CLI."
            ) from exc
        except Exception as exc:
            raise ValueError(
                "Question_Bank sheet not found in question papers. "
                "Regenerate papers with the latest app/CLI."
            ) from exc

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="embedded_qb_") as tmp:
            temp_path = tmp.name

        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            question_bank_df.to_excel(writer, index=False)

        return load_question_bank(temp_path)
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
