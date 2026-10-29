"""
Response Generator Module

Owns the shape of the student response sheet — the Google Forms export that Part 2
scores — and generates dummy sheets in that exact shape for rehearsing the pipeline.

Each student answers only the questions from their assigned set, so an answer column
is blank unless that question number was on that student's paper. All assigned
questions are treated as compulsory (no blank responses).
"""

import random
import re
from datetime import datetime, timedelta

import pandas as pd
from typing import List, Dict, Optional, Tuple
from pathlib import Path

from excel_handler import load_question_bank, set_label, parse_set_label, SET_LABEL_RE, FullQuestionBank
from same_questions import SetMap, read_set_map


# ── Response sheet column vocabulary ──────────────────────────────────────────
# These are the headers Google Forms produces. answer_checker parses them back.
TIMESTAMP_COL = "Timestamp"
EMAIL_COL = "Email address"
NAME_COL = "Full Name"
ROLL_COL = "Roll Number"
SET_COL = "Question Set"
IDENTITY_COLS = [TIMESTAMP_COL, EMAIL_COL, NAME_COL, ROLL_COL, SET_COL]

# Matches "Q - 01 [Answer]" as exported, plus the shorter forms people hand-edit to.
ANSWER_COL_RE = re.compile(r"^Q\s*-?\s*0*(\d+)\s*(?:\[Answer\])?$", re.IGNORECASE)

# Column we expose on a parsed set sheet holding each row's Question Number.
BANK_NO_COL = "BankNo"


def answer_column(question_no: int) -> str:
    """Google Forms header for a question's answer column."""
    return f"Q - {question_no:02d} [Answer]"


def _translate_letter(value, mapping: Dict[str, str]):
    """Upper-case and strip a typed letter, then map A-D; anything else is left as it is."""
    if pd.isna(value):
        return value
    letter = str(value).strip().upper()
    if not letter:
        return value
    return mapping.get(letter, letter)


def _translate_columns(
    response_df: pd.DataFrame,
    set_map: SetMap,
    source_number,
    target_number,
    letter_map,
    source_numbers: List[int],
    target_numbers: List[int],
    refusal,
) -> pd.DataFrame:
    """
    Re-key every row's answer columns for its set.

    source_number(position, entry) / target_number(position, entry) give the question
    number an answer column is read from and written to; letter_map(entry) gives the
    letter translation. source_numbers are the numbers a column may validly carry and
    target_numbers the columns written, in order. An answer column outside
    source_numbers is dropped when blank in every row; otherwise refusal(col) is the
    error message.
    """
    answer_cols = {}
    for col in response_df.columns:
        match = ANSWER_COL_RE.match(str(col).strip())
        if match:
            answer_cols[int(match.group(1))] = col

    for number, col in answer_cols.items():
        if number not in source_numbers and response_df[col].notna().any():
            raise ValueError(refusal(col))

    kept = [c for c in response_df.columns if c not in answer_cols.values()]
    out = response_df[kept].copy()
    translated = {no: [] for no in target_numbers}

    for idx, row in response_df.iterrows():
        set_no = str(row.get(SET_COL, "")).strip()
        if set_no not in set_map:
            raise ValueError(f"Unknown or missing {SET_COL} at row {idx + 2}: '{set_no}'")
        by_target = {}
        for position, entry in enumerate(set_map[set_no], start=1):
            col = answer_cols.get(source_number(position, entry))
            value = row[col] if col is not None else None
            by_target[target_number(position, entry)] = _translate_letter(value, letter_map(entry))
        for no in target_numbers:
            translated[no].append(by_target[no])

    for no, values in translated.items():
        out[answer_column(no)] = pd.Series(values, index=out.index, dtype=object)
    return out


def _bank_numbers(set_map: SetMap) -> List[int]:
    """The bank question numbers every paper carries, in order."""
    if not set_map:
        return []
    return sorted(entry.question_no for entry in next(iter(set_map.values())))


def to_bank_columns(response_df: pd.DataFrame, set_map: SetMap) -> pd.DataFrame:
    """
    Turn positional answers (Q1..Qn as printed on each paper) into bank-question answers.

    Identity columns are kept; answer columns become one per bank question_no holding
    the bank's original letter.
    """
    bank_numbers = _bank_numbers(set_map)
    n = len(bank_numbers)
    return _translate_columns(
        response_df, set_map,
        source_number=lambda position, entry: position,
        target_number=lambda position, entry: entry.question_no,
        letter_map=lambda entry: dict(zip("ABCD", entry.slots)),
        source_numbers=list(range(1, n + 1)),
        target_numbers=bank_numbers,
        refusal=lambda col: f"'{col}' is beyond the {n} questions on each paper",
    )


def to_positional_columns(response_df: pd.DataFrame, set_map: SetMap) -> pd.DataFrame:
    """Exact inverse of to_bank_columns: Q - 01 .. Q - n hold the letter as printed on that set."""
    bank_numbers = _bank_numbers(set_map)
    return _translate_columns(
        response_df, set_map,
        source_number=lambda position, entry: entry.question_no,
        target_number=lambda position, entry: position,
        letter_map=lambda entry: {orig: "ABCD"[slot] for slot, orig in enumerate(entry.slots)},
        source_numbers=bank_numbers,
        target_numbers=list(range(1, len(bank_numbers) + 1)),
        refusal=lambda col: f"'{col}' is not a question in the bank",
    )


def _attach_bank_no(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add a BankNo column when the set sheet carries Question Numbers.

    Papers written since the QCd change have two 'QCd' columns — the bare number and
    its 'Q- 27' printed form — which pandas reads as 'QCd' and 'QCd.1'. Only one of
    them parses as a number, and that is the one we want. Papers generated before the
    change have neither; callers fall back to matching on question text.
    """
    for col in [c for c in df.columns if str(c).strip().lower().startswith("qcd")]:
        numbers = pd.to_numeric(df[col], errors="coerce")
        if numbers.notna().all():
            out = df.copy()
            out[BANK_NO_COL] = numbers.astype(int)
            return out
    return df


def _read_set_sheet(source, sheet_name: str) -> pd.DataFrame:
    """
    Read a set sheet robustly across formats.

    Supports both:
    - plain sheets where header is on first row
    - styled sheets where title row exists and header appears later

    `source` is a path or an already-open pd.ExcelFile. Prefer the latter: every
    attempt here re-parses whatever it is handed, and a papers workbook holds one
    sheet per student, so re-opening the path costs the whole file each time.
    """
    def _normalize(value) -> str:
        return str(value).strip().lower().replace(" ", "_").replace(".", "")

    # The header row is not at a fixed position — a case study printed above the
    # questions pushes it down — so find it first, then read from it.
    try:
        raw = pd.read_excel(source, sheet_name=sheet_name, header=None)
    except Exception as exc:
        raise ValueError(f"Could not read '{sheet_name}'.") from exc

    for header_row, values in raw.iterrows():
        if "question" not in {_normalize(v) for v in values if pd.notna(v)}:
            continue

        df = pd.read_excel(source, sheet_name=sheet_name, header=header_row)
        q_col = next(col for col in df.columns if _normalize(col) == "question")

        out = df.rename(columns={q_col: "Question"}).copy()
        out = out[out["Question"].notna()]
        if len(out) > 0:
            return _attach_bank_no(out)

    raise ValueError(f"Could not parse '{sheet_name}' with a valid Question column.")


def extract_set_questions(question_papers_path: str) -> Dict[str, List[Tuple[int, str]]]:
    """
    Extract each student's assigned questions from the question papers Excel.

    Reads the Answer_Key sheet to determine:
    - Which set each student has
    - What the correct answer is for each positional question

    Then reads each set sheet to get the actual question numbers
    (from Q.No column mapping to original question_no via question text matching).

    Returns:
        Dict mapping set_name -> list of (original_question_no, correct_answer)
    """
    # Read the Answer_Key sheet to get correct answers per set
    answer_key_df = pd.read_excel(question_papers_path, sheet_name='Answer_Key')

    # Read each Set sheet to get the original question numbers
    xl = pd.ExcelFile(question_papers_path)
    set_sheets = [s for s in xl.sheet_names if SET_LABEL_RE.match(s)]

    set_questions = {}

    for sheet_name in set_sheets:
        # Read the question paper sheet
        paper_df = _read_set_sheet(xl, sheet_name)

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

    Reads each Set sheet, taking the Question Number straight from its QCd column,
    and falling back to matching question text for papers generated before QCd
    was printed.

    Returns:
        Dict mapping set_name -> list of original question_no values
    """
    xl = pd.ExcelFile(question_papers_path)
    set_sheets = [s for s in xl.sheet_names if SET_LABEL_RE.match(s)]

    # Build lookup: question_text -> question_no
    text_to_no = {}
    for q in question_bank.get_all():
        text_to_no[q.question_text.strip()] = q.question_no

    set_to_question_nos = {}

    for sheet_name in set_sheets:
        paper_df = _read_set_sheet(xl, sheet_name)

        if BANK_NO_COL in paper_df.columns:
            set_to_question_nos[sheet_name] = paper_df[BANK_NO_COL].tolist()
            continue

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
        blank_rate: Deprecated; kept for compatibility. Blanks are not generated.
        seed: Random seed for reproducibility

    Returns:
        DataFrame in Google Forms export shape: Timestamp, Email address, Full Name,
        Roll Number, Question Set, then one "Q - NN [Answer]" column per bank question.
    """
    rng = random.Random(seed) if seed is not None else random.Random()

    # --- Gather info from question papers and bank ---
    total_questions = len(question_bank.get_all())
    set_to_question_nos = map_paper_to_bank_questions(
        question_papers_path, question_bank
    )

    # Get available set names
    set_names = sorted(set_to_question_nos.keys(), key=parse_set_label)

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
    # Submissions trickle in over the quiz window, as a real Form's would.
    submitted_at = datetime(2026, 2, 16, 16, 0, 0)

    set_map = read_set_map(question_papers_path)
    column_numbers = (
        _bank_numbers(set_map) if set_map is not None
        else range(1, total_questions + 1)
    )

    rows = []
    for student_idx in range(num_students):
        set_name = set_names[student_idx]
        assigned_qnos = set_to_question_nos[set_name]
        student_no = student_idx + 1

        submitted_at += timedelta(seconds=rng.randint(20, 180))

        row = {
            TIMESTAMP_COL: submitted_at,
            EMAIL_COL: f"student{student_no:02d}@example.com",
            NAME_COL: f"Student {student_no:02d}",
            ROLL_COL: f"R{student_no:03d}",
            SET_COL: set_name,
        }
        # Every bank question gets a column; only assigned ones get filled.
        for q_no in column_numbers:
            row[answer_column(q_no)] = None

        for q_no in assigned_qnos:
            correct_answer = qno_to_answer[q_no]
            roll = rng.random()

            if roll < correct_rate:
                row[answer_column(q_no)] = correct_answer
            else:
                # Assigned questions are compulsory, so the remaining probability
                # (wrong_rate and anything left over) all maps to a wrong option.
                wrong_options = [o for o in all_options if o != correct_answer]
                row[answer_column(q_no)] = rng.choice(wrong_options)

        rows.append(row)

    df = pd.DataFrame(rows)
    if set_map is not None:
        # Every bank question is on every paper, so the columns above are complete.
        return to_positional_columns(df, set_map)
    return df


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
