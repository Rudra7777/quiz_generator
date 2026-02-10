"""
Quiz Generator - Streamlit UI

A web interface for generating randomized quiz papers from a question bank.
Upload an Excel file, configure settings, and download formatted question papers.
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from typing import Dict, List
from collections import defaultdict

from allocator import QuizStructure, allocate_quizzes, shuffle_all_quizzes
from excel_handler import load_question_bank, FullQuestionBank
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


# Page configuration
st.set_page_config(
    page_title="Quiz Generator",
    page_icon="📝",
    layout="wide"
)


def qid_to_number(question_id: str, question_bank: FullQuestionBank) -> int:
    """Convert internal question_id (e.g., H1, M5) to original question_no."""
    q = question_bank.get_by_id(question_id)
    return q.question_no if q else 0


def create_formatted_excel(
    allocation_matrix: list,
    shuffled_matrix: list,
    usage_counts: dict,
    question_bank: FullQuestionBank,
    include_answer_key: bool = True
) -> bytes:
    """
    Create a formatted Excel file with:
    - Question papers (one sheet per student)
    - Answer Key
    - Allocation Table (original order, numeric IDs)
    - Shuffled Table (shuffled order, numeric IDs)
    - Evaluation Table (min/max/delta stats)
    """
    wb = Workbook()

    # ── Shared styles ─────────────────────────────────────────────────────
    header_font = Font(bold=True, size=14)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=11, color="FFFFFF")
    green_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    orange_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    wrap_align = Alignment(wrap_text=True, vertical='top')
    center_align = Alignment(horizontal='center', vertical='center')

    # Remove default sheet
    wb.remove(wb.active)

    # ══════════════════════════════════════════════════════════════════════
    # Question Paper Sheets (one per student)
    # ══════════════════════════════════════════════════════════════════════
    for student_idx, quiz in enumerate(shuffled_matrix):
        ws = wb.create_sheet(title=f"Set_{student_idx + 1}")

        # Title
        ws.merge_cells('A1:F1')
        ws['A1'] = f"Question Paper - Set {student_idx + 1}"
        ws['A1'].font = header_font
        ws['A1'].alignment = center_align

        # Headers
        for col, header in enumerate(['Q.No', 'Question', 'Option A', 'Option B', 'Option C', 'Option D'], 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = center_align

        # Questions
        for q_idx, question_id in enumerate(quiz):
            q = question_bank.get_by_id(question_id)
            row = q_idx + 4
            ws.cell(row=row, column=1, value=q_idx + 1).border = thin_border
            ws.cell(row=row, column=1).alignment = center_align
            ws.cell(row=row, column=2, value=q.question_text).border = thin_border
            ws.cell(row=row, column=2).alignment = wrap_align
            ws.cell(row=row, column=3, value=q.option_a).border = thin_border
            ws.cell(row=row, column=4, value=q.option_b).border = thin_border
            ws.cell(row=row, column=5, value=q.option_c).border = thin_border
            ws.cell(row=row, column=6, value=q.option_d).border = thin_border

        # Column widths & row heights
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 50
        for ch in 'CDEF':
            ws.column_dimensions[ch].width = 20
        for r in range(4, len(quiz) + 4):
            ws.row_dimensions[r].height = 30

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
            ws.cell(row=row, column=1, value=f"Set_{student_idx + 1}").border = thin_border
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
        cell = ws.cell(row=3, column=s_idx + 2, value=f"S{s_idx + 1}")
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
        cell = ws.cell(row=3, column=s_idx + 2, value=f"S{s_idx + 1}")
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

    # ── Save ──────────────────────────────────────────────────────────────
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def main():
    st.title("📝 Quiz Question Paper Generator")
    st.markdown("Upload a question bank and generate randomized question papers for all students.")

    st.divider()

    # ========================================================================
    # Step 1: Upload Question Bank
    # ========================================================================
    st.header("1️⃣ Upload Question Bank")

    uploaded_file = st.file_uploader(
        "Upload Excel file (.xlsx)",
        type=['xlsx'],
        help="Excel with columns: question_no, question, option_a, option_b, option_c, option_d, answer, difficulty"
    )

    question_bank = None
    counts = {'hard': 0, 'medium': 0, 'easy': 0}

    if uploaded_file:
        try:
            with open("temp_upload.xlsx", "wb") as f:
                f.write(uploaded_file.getvalue())

            question_bank = load_question_bank("temp_upload.xlsx")
            counts = question_bank.count_by_difficulty()
            total = sum(counts.values())

            st.success(f"✅ Loaded {total} questions successfully!")

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Hard", counts.get('hard', 0))
            col2.metric("Medium", counts.get('medium', 0))
            col3.metric("Easy", counts.get('easy', 0))
            col4.metric("Total", total)

        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
            question_bank = None

    st.divider()

    # ========================================================================
    # Step 2: Configuration
    # ========================================================================
    st.header("2️⃣ Configuration")

    col1, col2 = st.columns(2)

    with col1:
        num_students = st.number_input(
            "Number of Students",
            min_value=1, max_value=500, value=50,
            help="How many question papers to generate"
        )

    with col2:
        total_questions = st.number_input(
            "Questions per Quiz",
            min_value=1, max_value=100, value=15,
            help="Total questions in each student's quiz"
        )

    st.divider()

    # ========================================================================
    # Step 3: Difficulty Distribution
    # ========================================================================
    st.header("3️⃣ Difficulty Distribution")

    mode = st.radio(
        "Select mode:",
        ["Absolute (exact counts)", "Percentage (auto-calculate)"],
        horizontal=True
    )

    hard_count = medium_count = easy_count = 0

    if mode == "Absolute (exact counts)":
        col1, col2, col3 = st.columns(3)
        with col1:
            hard_count = st.number_input("Hard Questions", min_value=0, max_value=total_questions,
                                         value=min(4, counts.get('hard', 4)), key="hard_abs")
        with col2:
            medium_count = st.number_input("Medium Questions", min_value=0, max_value=total_questions,
                                           value=min(6, counts.get('medium', 6)), key="medium_abs")
        with col3:
            easy_count = st.number_input("Easy Questions", min_value=0, max_value=total_questions,
                                         value=min(5, counts.get('easy', 5)), key="easy_abs")
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            hard_pct = st.slider("Hard %", 0, 100, 27, key="hard_pct")
        with col2:
            medium_pct = st.slider("Medium %", 0, 100, 40, key="medium_pct")
        with col3:
            easy_pct = st.slider("Easy %", 0, 100, 33, key="easy_pct")

        total_pct = hard_pct + medium_pct + easy_pct
        if total_pct != 100:
            st.warning(f"⚠️ Percentages sum to {total_pct}%, should be 100%")

        hard_count = round(total_questions * hard_pct / 100)
        medium_count = round(total_questions * medium_pct / 100)
        easy_count = total_questions - hard_count - medium_count

        st.info(f"📊 Calculated: {hard_count} Hard + {medium_count} Medium + {easy_count} Easy = {hard_count + medium_count + easy_count} questions")

    # Validation
    total_selected = hard_count + medium_count + easy_count
    if total_selected != total_questions:
        st.error(f"❌ Selected {total_selected} questions, but quiz requires {total_questions}")

    validation_errors = []
    if question_bank:
        if hard_count > counts.get('hard', 0):
            validation_errors.append(f"Need {hard_count} hard questions, only {counts.get('hard', 0)} available")
        if medium_count > counts.get('medium', 0):
            validation_errors.append(f"Need {medium_count} medium questions, only {counts.get('medium', 0)} available")
        if easy_count > counts.get('easy', 0):
            validation_errors.append(f"Need {easy_count} easy questions, only {counts.get('easy', 0)} available")

    for error in validation_errors:
        st.error(f"❌ {error}")

    st.divider()

    # ========================================================================
    # Step 4: Generate
    # ========================================================================
    st.header("4️⃣ Generate Question Papers")

    can_generate = (
        question_bank is not None
        and total_selected == total_questions
        and len(validation_errors) == 0
    )

    if st.button("🚀 Generate Question Papers", disabled=not can_generate, type="primary"):
        with st.spinner("Generating question papers..."):
            try:
                quiz_structure = QuizStructure(
                    hard_count=hard_count,
                    medium_count=medium_count,
                    easy_count=easy_count
                )

                q_ids_by_diff = {
                    'hard': question_bank.get_question_ids_by_difficulty('hard'),
                    'medium': question_bank.get_question_ids_by_difficulty('medium'),
                    'easy': question_bank.get_question_ids_by_difficulty('easy'),
                }

                # Run allocation
                allocation_matrix, usage_counts = allocate_quizzes(
                    q_ids_by_diff,
                    num_students=num_students,
                    quiz_structure=quiz_structure,
                    seed=42
                )

                # Shuffle
                shuffled_matrix = shuffle_all_quizzes(allocation_matrix, base_seed=42)

                # Create formatted Excel with all sheets
                excel_bytes = create_formatted_excel(
                    allocation_matrix=allocation_matrix,
                    shuffled_matrix=shuffled_matrix,
                    usage_counts=usage_counts,
                    question_bank=question_bank
                )

                st.success(f"✅ Generated {num_students} question papers!")

                st.download_button(
                    label="📥 Download Question Papers (Excel)",
                    data=excel_bytes,
                    file_name="question_papers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

                # Show stats
                st.subheader("📊 Generation Statistics")
                col1, col2, col3 = st.columns(3)
                col1.metric("Students", num_students)
                col2.metric("Questions/Quiz", total_questions)
                col3.metric("Total Sheets", num_students + 4)  # papers + answer key + alloc + shuffled + eval

                st.caption("Sheets: Set_1 … Set_N, Answer_Key, Allocation_Table, Shuffled_Table, Evaluation")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

    # Cleanup temp file
    if os.path.exists("temp_upload.xlsx"):
        os.remove("temp_upload.xlsx")


if __name__ == "__main__":
    main()
