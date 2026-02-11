from pathlib import Path

import pandas as pd

from answer_checker import (
    check_all_responses,
    generate_scoring_report,
    load_response_sheet,
)
from excel_handler import load_question_bank
from response_generator import generate_responses, map_paper_to_bank_questions


ROOT = Path(__file__).resolve().parents[1]
QUESTION_BANK = ROOT / "question_bank_72.xlsx"
QUESTION_PAPERS = ROOT / "output" / "question_papers.xlsx"
RESPONSES = ROOT / "output" / "student_responses.xlsx"


def test_regression_existing_response_sheet():
    """Regression: current fixture responses should keep known metrics."""
    question_bank = load_question_bank(str(QUESTION_BANK))
    response_df = load_response_sheet(str(RESPONSES))

    report = check_all_responses(
        response_df=response_df,
        question_papers_path=str(QUESTION_PAPERS),
        question_bank=question_bank,
    )

    assert len(report.student_reports) == 70
    assert report.avg_score == 10.76
    assert report.median_score == 11.0
    assert report.pass_rate == 100.0
    assert report.pass_count == 70
    assert len(report.validation_issues) == 0
    assert report.grade_distribution() == {
        "14/15": 5,
        "13/15": 4,
        "12/15": 14,
        "11/15": 16,
        "10/15": 15,
        "9/15": 10,
        "8/15": 5,
        "7/15": 1,
    }


def test_seeded_generation_is_deterministic():
    """Same seed + same inputs should produce identical response sheets."""
    question_bank = load_question_bank(str(QUESTION_BANK))

    df1 = generate_responses(
        question_papers_path=str(QUESTION_PAPERS),
        question_bank=question_bank,
        num_students=20,
        correct_rate=0.70,
        wrong_rate=0.20,
        blank_rate=0.10,
        seed=2026,
    )
    df2 = generate_responses(
        question_papers_path=str(QUESTION_PAPERS),
        question_bank=question_bank,
        num_students=20,
        correct_rate=0.70,
        wrong_rate=0.20,
        blank_rate=0.10,
        seed=2026,
    )

    pd.testing.assert_frame_equal(df1, df2, check_dtype=False)

    report = check_all_responses(
        response_df=df1,
        question_papers_path=str(QUESTION_PAPERS),
        question_bank=question_bank,
    )
    assert len(report.validation_issues) == 0
    assert all(r.assigned == 15 for r in report.student_reports)
    assert all(r.attempted == r.assigned for r in report.student_reports)
    assert all(r.unanswered == 0 for r in report.student_reports)
    assert all(r.correct + r.wrong == r.assigned for r in report.student_reports)


def test_validation_flags_extra_answer_on_unassigned_question(tmp_path: Path):
    """If student answers outside assigned set, validation must report it."""
    question_bank = load_question_bank(str(QUESTION_BANK))
    response_df = generate_responses(
        question_papers_path=str(QUESTION_PAPERS),
        question_bank=question_bank,
        num_students=5,
        seed=7,
    )

    set_map = map_paper_to_bank_questions(str(QUESTION_PAPERS), question_bank)
    assigned_qnos = set(set_map["Set_1"])
    total_qnos = set(range(1, len(question_bank.get_all()) + 1))
    extra_qno = min(total_qnos - assigned_qnos)
    response_df.loc[0, f"Q{extra_qno}"] = "A"

    report = check_all_responses(
        response_df=response_df,
        question_papers_path=str(QUESTION_PAPERS),
        question_bank=question_bank,
    )

    assert len(report.validation_issues) >= 1
    first = report.validation_issues[0]
    assert first.student_index == 0
    assert first.validation.extra_count >= 1
    assert extra_qno in first.validation.extra_questions

    output_path = tmp_path / "scoring_report_with_issue.xlsx"
    saved = generate_scoring_report(report, str(output_path))
    assert Path(saved).exists()

    validation_df = pd.read_excel(saved, sheet_name="Validation")
    assert "Extra Count" in validation_df.columns
