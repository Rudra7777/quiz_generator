"""
Answer Checker Module

Validates student responses and computes scoring reports.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional
import re

import pandas as pd
from openpyxl.styles import PatternFill

from excel_handler import FullQuestionBank
from response_generator import map_paper_to_bank_questions


VALID_OPTIONS = {"A", "B", "C", "D"}
QUESTION_COL_RE = re.compile(r"^Q(\d+)$")


@dataclass
class ValidationResult:
    """Validation details for one student's response row."""

    set_no: str
    extra_questions: List[int]

    @property
    def extra_count(self) -> int:
        return len(self.extra_questions)


@dataclass
class StudentReport:
    """Scoring details for one student."""

    student_index: int
    validation: ValidationResult
    assigned: int
    attempted: int
    correct: int
    wrong: int
    unanswered: int

    @property
    def set_no(self) -> str:
        return self.validation.set_no


@dataclass
class ScoringReport:
    """Aggregated scoring report for all students."""

    student_reports: List[StudentReport]
    validation_issues: List[StudentReport]
    avg_score: float
    median_score: float
    pass_count: int
    pass_rate: float
    pass_threshold: float = 6.0

    def grade_distribution(self) -> Dict[str, int]:
        """Return distribution by obtained marks (Correct / Max Marks)."""
        buckets: Dict[str, int] = {}
        for report in self.student_reports:
            key = f"{report.correct}/{report.assigned}"
            buckets[key] = buckets.get(key, 0) + 1
        # Sort by marks descending for readability.
        buckets = dict(sorted(buckets.items(), key=lambda kv: int(kv[0].split("/")[0]), reverse=True))
        return buckets


def _normalize_answer(value: object) -> Optional[str]:
    """Normalize answer value to A/B/C/D, or None for blank."""
    if pd.isna(value):
        return None
    normalized = str(value).strip().upper()
    if not normalized:
        return None
    return normalized


def _extract_answered_questions(row: pd.Series) -> Dict[int, str]:
    """Extract non-blank answered question numbers from a response row."""
    answered: Dict[int, str] = {}
    for col, value in row.items():
        match = QUESTION_COL_RE.match(str(col))
        if not match:
            continue
        answer = _normalize_answer(value)
        if answer is None:
            continue
        answered[int(match.group(1))] = answer
    return answered


def load_response_sheet(filepath: str) -> pd.DataFrame:
    """
    Load response sheet from Excel.

    Expected minimum columns:
    - Set_No
    - Q1..Qn (question columns)
    """
    df = pd.read_excel(filepath, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    if "Set_No" not in df.columns:
        normalized = {c.lower().replace(" ", "").replace("_", ""): c for c in df.columns}
        set_aliases = ("setno", "set", "q")
        alias_col = next((normalized[a] for a in set_aliases if a in normalized), None)

        # Fallback: if first column looks like set labels, treat it as Set_No.
        if alias_col is None and len(df.columns) > 0:
            first_col = df.columns[0]
            sample = df[first_col].dropna().astype(str).head(5)
            if not sample.empty and sample.str.startswith("Set_").all():
                alias_col = first_col

        if alias_col is not None:
            df = df.rename(columns={alias_col: "Set_No"})
        else:
            raise ValueError("Missing required column 'Set_No' in responses sheet")

    return df


def check_all_responses(
    response_df: pd.DataFrame,
    question_papers_path: str,
    question_bank: FullQuestionBank,
    pass_threshold: float = 6.0,
) -> ScoringReport:
    """
    Validate and score all student responses.
    """
    set_to_question_nos = map_paper_to_bank_questions(question_papers_path, question_bank)

    qno_to_answer = {
        q.question_no: str(q.answer).strip().upper()
        for q in question_bank.get_all()
    }

    student_reports: List[StudentReport] = []

    for idx, row in response_df.iterrows():
        set_no = str(row.get("Set_No", "")).strip()
        if set_no not in set_to_question_nos:
            raise ValueError(f"Unknown or missing Set_No at row {idx + 2}: '{set_no}'")

        assigned_qnos = set_to_question_nos[set_no]
        assigned_set = set(assigned_qnos)
        answered = _extract_answered_questions(row)

        extra_questions = sorted(q_no for q_no in answered if q_no not in assigned_set)

        correct = 0
        wrong = 0
        unanswered = 0

        for q_no in assigned_qnos:
            answer = answered.get(q_no)
            if answer is None:
                # Compulsory forms: unanswered is treated as wrong.
                continue
            if answer == qno_to_answer[q_no]:
                correct += 1
            else:
                wrong += 1

        wrong += (len(assigned_qnos) - (correct + wrong))
        attempted = len(assigned_qnos)
        unanswered = 0

        validation = ValidationResult(set_no=set_no, extra_questions=extra_questions)
        student_reports.append(
            StudentReport(
                student_index=idx,
                validation=validation,
                assigned=len(assigned_qnos),
                attempted=attempted,
                correct=correct,
                wrong=wrong,
                unanswered=unanswered,
            )
        )

    score_series = pd.Series([r.correct for r in student_reports], dtype=float)
    avg_score = round(float(score_series.mean()), 2) if len(score_series) else 0.0
    median_score = round(float(score_series.median()), 2) if len(score_series) else 0.0
    pass_count = sum(1 for r in student_reports if r.correct >= pass_threshold)
    pass_rate = round((pass_count / len(student_reports)) * 100, 2) if student_reports else 0.0

    validation_issues = [r for r in student_reports if r.validation.extra_count > 0]

    return ScoringReport(
        student_reports=student_reports,
        validation_issues=validation_issues,
        avg_score=avg_score,
        median_score=median_score,
        pass_count=pass_count,
        pass_rate=pass_rate,
        pass_threshold=pass_threshold,
    )


def generate_scoring_report(
    report: ScoringReport,
    output_path: str,
    response_df: Optional[pd.DataFrame] = None,
    question_papers_path: Optional[str] = None,
    question_bank: Optional[FullQuestionBank] = None,
) -> str:
    """
    Write scoring report to Excel.

    Always includes:
    - Scores
    - Summary
    - Validation

    Additionally includes 'Responses_Review' with colored answer cells when
    response_df, question_papers_path, and question_bank are provided.
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    scores_rows = []
    for r in report.student_reports:
        scores_rows.append(
            {
                "Student": r.student_index + 1,
                "Set": r.set_no,
                "Assigned": r.assigned,
                "Attempted": r.attempted,
                "Correct": r.correct,
                "Wrong": r.wrong,
                "Extra Answers": r.validation.extra_count,
            }
        )
    scores_df = pd.DataFrame(scores_rows)

    max_marks = scores_df["Assigned"].mode().iloc[0] if not scores_df.empty else 0

    summary_rows = [
        {"Metric": "Total Students", "Value": len(report.student_reports)},
        {"Metric": "Max Marks (per student)", "Value": max_marks},
        {"Metric": "Average Correct (marks)", "Value": report.avg_score},
        {"Metric": "Median Correct (marks)", "Value": report.median_score},
        {"Metric": f"Pass Count (≥{int(report.pass_threshold)} marks)", "Value": report.pass_count},
        {"Metric": "Pass Rate (%)", "Value": report.pass_rate},
        {"Metric": "---", "Value": "---"},
        {"Metric": "Score Distribution (Correct/Max)", "Value": None},
    ]
    for marks, count in report.grade_distribution().items():
        summary_rows.append({"Metric": marks, "Value": count})
    summary_df = pd.DataFrame(summary_rows)

    if report.validation_issues:
        validation_rows = []
        for r in report.validation_issues:
            extra_str = ", ".join([f"Q{q}" for q in r.validation.extra_questions])
            validation_rows.append(
                {
                    "Student": r.student_index + 1,
                    "Set": r.set_no,
                    "Extra Count": r.validation.extra_count,
                    "Extra Questions": extra_str,
                }
            )
        validation_df = pd.DataFrame(validation_rows)
    else:
        validation_df = pd.DataFrame(
            [
                {
                    "Status": "✅ No validation issues found",
                    "Details": "All students answered only their assigned questions",
                }
            ]
        )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        scores_df.to_excel(writer, sheet_name="Scores", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        validation_df.to_excel(writer, sheet_name="Validation", index=False)

        if (
            response_df is not None
            and question_papers_path is not None
            and question_bank is not None
        ):
            review_df = response_df.copy()
            review_df.to_excel(writer, sheet_name="Responses_Review", index=False)

            ws = writer.book["Responses_Review"]
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

            set_to_question_nos = map_paper_to_bank_questions(question_papers_path, question_bank)
            qno_to_answer = {
                q.question_no: str(q.answer).strip().upper()
                for q in question_bank.get_all()
            }

            col_to_idx = {str(col): idx + 1 for idx, col in enumerate(review_df.columns)}

            for row_idx, row in review_df.iterrows():
                set_no = str(row.get("Set_No", "")).strip()
                assigned_qnos = set(set_to_question_nos.get(set_no, []))
                excel_row = row_idx + 2  # Header is row 1

                for col in review_df.columns:
                    match = QUESTION_COL_RE.match(str(col))
                    if not match:
                        continue

                    q_no = int(match.group(1))
                    answer = _normalize_answer(row[col])
                    if answer is None:
                        continue

                    excel_col = col_to_idx[str(col)]
                    cell = ws.cell(row=excel_row, column=excel_col)

                    if q_no in assigned_qnos and answer == qno_to_answer.get(q_no):
                        cell.fill = green_fill
                    else:
                        cell.fill = red_fill

    return output_path
