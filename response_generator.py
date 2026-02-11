"""
Response Generator Module

Generates a dummy student response sheet (Excel) simulating Google Form output.
Each student answers only the questions from their assigned set.
Simulates realistic behavior: ~70% correct, ~20% wrong, ~10% blank.
"""

import random
import pandas as pd
from typing import List, Dict, Optional, Tuple
from pathlib import Path

from excel_handler import load_question_bank, FullQuestionBank


def _read_set_sheet(question_papers_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Read a Set_N sheet robustly across formats.

    Supports both:
    - plain sheets where header is on first row
    - styled sheets where title row exists and header appears later
    """
    for header_row in (0, 1, 2, 3, 4, 5):
        try:
            df = pd.read_excel(question_papers_path, sheet_name=sheet_name, header=header_row)
        except Exception:
            continue

        normalized = {
            str(col).strip().lower().replace(" ", "_").replace(".", ""): col
            for col in df.columns
        }
        q_col = normalized.get("question")

        if q_col is None:
            continue

        out = df.rename(columns={q_col: "Question"}).copy()
        out = out[out["Question"].notna()]
        if len(out) > 0:
            return out

    raise ValueError(f"Could not parse '{sheet_name}' with a valid Question column.")


def extract_set_questions(question_papers_path: str) -> Dict[str, List[Tuple[int, str]]]:
    """
    Extract each student's assigned questions from the question papers Excel.

    Reads the Answer_Key sheet to determine:
    - Which set each student has
    - What the correct answer is for each positional question

    Then reads each Set_N sheet to get the actual question numbers
    (from Q.No column mapping to original question_no via question text matching).

    Returns:
        Dict mapping set_name -> list of (original_question_no, correct_answer)
    """
    # Read the Answer_Key sheet to get correct answers per set
    answer_key_df = pd.read_excel(question_papers_path, sheet_name='Answer_Key')

    # Read each Set sheet to get the original question numbers
    xl = pd.ExcelFile(question_papers_path)
    set_sheets = [s for s in xl.sheet_names if s.startswith('Set_')]

    set_questions = {}

    for sheet_name in set_sheets:
        # Read the question paper sheet
        paper_df = _read_set_sheet(question_papers_path, sheet_name)

        # Get answer row for this set from answer key
        set_row = answer_key_df[answer_key_df['Set'] == sheet_name]
        if set_row.empty:
            continue

        questions = []
        for q_idx in range(len(paper_df)):
            # Q.No in the paper is sequential (1, 2, 3, ...)
            # We need to find the original question_no
            # The question text from the paper can be matched to the bank
            q_col = f'Q{q_idx + 1}'
            correct_answer = str(set_row.iloc[0][q_col]).strip().upper()
            questions.append((q_idx + 1, correct_answer))  # positional index, answer

        set_questions[sheet_name] = questions

    return set_questions


def map_paper_to_bank_questions(
    question_papers_path: str,
    question_bank: FullQuestionBank
) -> Dict[str, List[int]]:
    """
    Map each set's positional questions to original question_no from the bank.

    Reads each Set sheet, matches question text to the bank, and returns
    the original question numbers.

    Returns:
        Dict mapping set_name -> list of original question_no values
    """
    xl = pd.ExcelFile(question_papers_path)
    set_sheets = [s for s in xl.sheet_names if s.startswith('Set_')]

    # Build lookup: question_text -> question_no
    text_to_no = {}
    for q in question_bank.get_all():
        text_to_no[q.question_text.strip()] = q.question_no

    set_to_question_nos = {}

    for sheet_name in set_sheets:
        paper_df = _read_set_sheet(question_papers_path, sheet_name)
        question_nos = []

        for _, row in paper_df.iterrows():
            q_text = str(row['Question']).strip()
            q_no = text_to_no.get(q_text)
            if q_no is not None:
                question_nos.append(q_no)
            else:
                raise ValueError(
                    f"Could not match question in {sheet_name}: '{q_text[:50]}...'"
                )

        set_to_question_nos[sheet_name] = question_nos

    return set_to_question_nos


def generate_responses(
    question_papers_path: str,
    question_bank: FullQuestionBank,
    num_students: int,
    correct_rate: float = 0.70,
    wrong_rate: float = 0.20,
    blank_rate: float = 0.10,
    seed: Optional[int] = None
) -> pd.DataFrame:
    """
    Generate a dummy response DataFrame simulating Google Form answers.

    Args:
        question_papers_path: Path to generated question_papers.xlsx
        question_bank: FullQuestionBank with all question data
        num_students: Number of student responses to generate
        correct_rate: Probability of answering correctly (~70%)
        wrong_rate: Probability of answering wrong (~20%)
        blank_rate: Probability of leaving blank (~10%)
        seed: Random seed for reproducibility

    Returns:
        DataFrame with columns: Set_No, Q1, Q2, ..., QT
        where T = total questions in the bank
    """
    rng = random.Random(seed) if seed is not None else random.Random()

    # --- Gather info from question papers and bank ---
    total_questions = len(question_bank.get_all())
    set_to_question_nos = map_paper_to_bank_questions(
        question_papers_path, question_bank
    )

    # Get available set names
    set_names = sorted(set_to_question_nos.keys(),
                       key=lambda s: int(s.split('_')[1]))

    if num_students > len(set_names):
        raise ValueError(
            f"Requested {num_students} students but only "
            f"{len(set_names)} sets available in question papers"
        )

    # --- Build answer key: {question_no -> correct_answer} ---
    # Build from question bank directly
    qno_to_answer = {}
    for q in question_bank.get_all():
        qno_to_answer[q.question_no] = q.answer.strip().upper()

    all_options = ['A', 'B', 'C', 'D']

    # --- Generate responses ---
    rows = []
    for student_idx in range(num_students):
        set_name = set_names[student_idx]
        assigned_qnos = set_to_question_nos[set_name]

        # Initialize all question columns as blank (NaN)
        row = {'Set_No': set_name}
        for q_no in range(1, total_questions + 1):
            row[f'Q{q_no}'] = None  # blank by default

        # Fill in answers only for assigned questions
        for q_no in assigned_qnos:
            correct_answer = qno_to_answer[q_no]
            roll = rng.random()

            if roll < correct_rate:
                # Correct answer
                row[f'Q{q_no}'] = correct_answer
            elif roll < correct_rate + wrong_rate:
                # Wrong answer: pick a random wrong option
                wrong_options = [o for o in all_options if o != correct_answer]
                row[f'Q{q_no}'] = rng.choice(wrong_options)
            else:
                # Leave blank (student skipped)
                row[f'Q{q_no}'] = None

        rows.append(row)

    return pd.DataFrame(rows)


def save_response_sheet(df: pd.DataFrame, output_path: str) -> str:
    """
    Save response DataFrame to Excel.

    Args:
        df: Response DataFrame
        output_path: Path for output Excel file

    Returns:
        Path to created file
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, sheet_name='Responses')
    return output_path
