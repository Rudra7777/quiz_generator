"""Tests for spreading the correct answer evenly over A-D."""

import pandas as pd
import pytest

from option_shuffler import LETTERS, OPTION_COLS, answer_counts, balance_options

from test_answer_pipeline import QUESTION_BANK


def _bank() -> pd.DataFrame:
    return pd.read_excel(QUESTION_BANK)


def _correct_text(df: pd.DataFrame) -> list:
    return [row[f"option_{row['answer'].lower()}"] for _, row in df.iterrows()]


def _skewed(n_hard=12, n_medium=30, n_easy=30) -> pd.DataFrame:
    """A bank whose key is all 'A' — the worst case for guessing."""
    rows = []
    for level, n in (("H", n_hard), ("M", n_medium), ("L", n_easy)):
        for _ in range(n):
            q = len(rows) + 1
            rows.append({
                "question_no": q,
                "question": f"Question {q}?",
                "option_a": f"right {q}",
                "option_b": f"wrong {q}.1",
                "option_c": f"wrong {q}.2",
                "option_d": f"wrong {q}.3",
                "answer": "A",
                "difficulty": level,
            })
    return pd.DataFrame(rows)


def test_each_difficulty_is_split_as_evenly_as_possible():
    balanced = balance_options(_skewed(), seed=1)
    counts = answer_counts(balanced)

    assert counts.loc["Hard"].tolist() == [3, 3, 3, 3]
    for level in ("Medium", "Easy"):  # 30 does not divide by 4: 8/8/7/7 in some order
        assert sorted(counts.loc[level].tolist()) == [7, 7, 8, 8]


def test_leftovers_are_spread_so_the_whole_bank_is_even_too():
    counts = answer_counts(balance_options(_skewed(), seed=1))

    assert counts.sum().tolist() == [18, 18, 18, 18]  # 72 / 4


def test_every_question_keeps_its_correct_answer_and_its_options():
    original = _skewed()
    balanced = balance_options(original, seed=3)

    assert _correct_text(balanced) == _correct_text(original)
    for (_, before), (_, after) in zip(original.iterrows(), balanced.iterrows()):
        assert sorted(before[OPTION_COLS]) == sorted(after[OPTION_COLS])
        assert before["question"] == after["question"]
        assert before["difficulty"] == after["difficulty"]


def test_real_bank_round_trips_through_the_bank_loader(tmp_path):
    from excel_handler import load_question_bank

    balanced = balance_options(_bank(), seed=5)
    path = tmp_path / "balanced.xlsx"
    balanced.to_excel(path, index=False)

    bank = load_question_bank(str(path))
    assert len(bank.get_all()) == len(balanced)
    assert {q.answer for q in bank.get_all()} == set(LETTERS)


def test_same_seed_gives_the_same_bank_and_input_is_not_modified():
    original = _skewed()
    snapshot = original.copy()

    first = balance_options(original, seed=9)
    second = balance_options(original, seed=9)

    pd.testing.assert_frame_equal(first, second)
    pd.testing.assert_frame_equal(original, snapshot)


def test_a_bad_answer_letter_is_reported():
    bank = _skewed(n_hard=4, n_medium=0, n_easy=0)
    bank.loc[2, "answer"] = "E"

    with pytest.raises(ValueError, match="row 4"):
        balance_options(bank)
