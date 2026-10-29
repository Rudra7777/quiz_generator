"""Tests for building every set's question order and option order in Same Questions mode."""

from collections import Counter

import pytest

from excel_handler import FullQuestion, FullQuestionBank
from same_questions import build_same_question_sets, fixed_slots, held_in_place


def _question(no, options=None, answer="A", difficulty="medium") -> FullQuestion:
    a, b, c, d = options or (f"right {no}", f"wrong {no}.1", f"wrong {no}.2", f"wrong {no}.3")
    return FullQuestion(
        question_no=no, question_id=f"Q{no}", question_text=f"Question {no}?",
        option_a=a, option_b=b, option_c=c, option_d=d,
        answer=answer, difficulty=difficulty,
    )


@pytest.fixture
def bank20() -> FullQuestionBank:
    """20 questions, answers mostly A; Q5, Q9 and Q12 have options that refer to others."""
    levels = ["hard", "medium", "easy", "medium"]
    questions = [_question(no, difficulty=levels[no % 4]) for no in range(1, 21)]
    questions[4] = _question(5, ("one", "two", "three", "All of the above"), "D", "hard")
    questions[8] = _question(9, ("one", "two", "None of these", "four"), "A")
    questions[11] = _question(12, ("one", "two", "three", "Both A and B"), "A")
    return FullQuestionBank(questions)


def test_fixed_slots_detects_reference_options(bank20):
    q = {q.question_no: q for q in bank20.get_all()}
    assert fixed_slots(q[5]) == frozenset({3})
    assert fixed_slots(q[9]) == frozenset({2})
    assert fixed_slots(q[12]) == frozenset({0, 1, 2, 3})
    assert fixed_slots(q[1]) == frozenset()
    assert held_in_place(bank20) == [5, 9, 12]


def test_every_set_has_every_question_once(bank20):
    sets = build_same_question_sets(bank20, 65, seed=1)
    assert list(sets) == [f"S-{i:02d}" for i in range(1, 66)]
    for positions in sets.values():
        assert sorted(p.question_no for p in positions) == list(range(1, 21))


def test_no_two_sets_are_identical(bank20):
    sets = build_same_question_sets(bank20, 65, seed=1)
    assert len({tuple(sets[s]) for s in sets}) == 65


def test_slots_are_a_permutation_and_answer_points_at_the_bank_answer(bank20):
    bank = {q.question_no: q for q in bank20.get_all()}
    for positions in build_same_question_sets(bank20, 10, seed=2).values():
        for p in positions:
            assert sorted(p.slots) == list("ABCD")
            assert p.slots["ABCD".index(p.answer)] == bank[p.question_no].answer


def test_held_options_stay_in_their_slot(bank20):
    for positions in build_same_question_sets(bank20, 30, seed=3).values():
        by_no = {p.question_no: p for p in positions}
        assert by_no[5].slots[3] == "D" and by_no[5].answer == "D"
        assert by_no[9].slots[2] == "C"
        assert by_no[12].slots == "ABCD"


def test_each_set_key_is_balanced(bank20):
    # 20 questions -> 5 of each letter; Q5 (D) and Q12 (fixed) are absorbed.
    for positions in build_same_question_sets(bank20, 65, seed=4).values():
        counts = Counter(p.answer for p in positions)
        assert max(counts.values()) - min(counts.values()) <= 1


def test_seed_makes_generation_repeatable(bank20):
    assert build_same_question_sets(bank20, 5, seed=9) == build_same_question_sets(bank20, 5, seed=9)


def test_more_sets_than_possible_orders_is_refused():
    bank3 = FullQuestionBank([_question(no) for no in (1, 2, 3)])
    with pytest.raises(ValueError, match="only 6 different orders"):
        build_same_question_sets(bank3, 7, seed=0)


# ── Papers file ──────────────────────────────────────────────────────────────

import io

from openpyxl import load_workbook

from excel_export import (
    _load_question_bank_from_question_papers, create_formatted_excel,
    create_same_questions_excel,
)
from same_questions import read_set_map, variant_papers


def _papers_file(bank, tmp_path, num_sets=4, seed=1):
    sets = build_same_question_sets(bank, num_sets, seed=seed)
    path = tmp_path / "papers.xlsx"
    path.write_bytes(create_same_questions_excel(sets, bank, seed=seed))
    return sets, path


def test_same_questions_file_has_the_expected_sheets(bank20):
    data = create_same_questions_excel(build_same_question_sets(bank20, 4, seed=1), bank20, seed=1)
    wb = load_workbook(io.BytesIO(data))
    assert wb.sheetnames == ["All_Sets", "S-01", "S-02", "S-03", "S-04",
                             "Answer_Key", "Set_Map", "Question_Bank"]


def test_set_map_round_trips(bank20, tmp_path):
    sets, path = _papers_file(bank20, tmp_path)
    assert read_set_map(str(path)) == sets


def test_old_papers_have_no_set_map(tmp_path):
    bank = FullQuestionBank([_question(no, difficulty="hard") for no in range(1, 5)])
    matrix = [["Q1", "Q2"], ["Q3", "Q4"]]
    path = tmp_path / "old.xlsx"
    path.write_bytes(create_formatted_excel(matrix, matrix, {}, bank))
    assert read_set_map(str(path)) is None


def test_paper_prints_positions_and_shuffled_options_without_bank_numbers(bank20, tmp_path):
    sets, path = _papers_file(bank20, tmp_path)
    ws = load_workbook(path)["S-02"]
    bank = {q.question_no: q for q in bank20.get_all()}
    assert ws["C2"].value is None
    for k in range(1, 21):
        row = k + 2
        p = sets["S-02"][k - 1]
        assert ws.cell(row=row, column=1).value == k
        assert ws.cell(row=row, column=3).value is None
        assert ws.cell(row=row, column=2).value == p.question_no
        assert ws.cell(row=row, column=4).value == bank[p.question_no].question_text
        original = getattr(bank[p.question_no], f"option_{p.slots[0].lower()}")
        assert ws.cell(row=row, column=5).value == f"A] {original}"


def test_answer_key_is_per_set_by_position(bank20, tmp_path):
    sets, path = _papers_file(bank20, tmp_path)
    ws = load_workbook(path)["Answer_Key"]
    row = next(r for r in range(1, 12) if ws.cell(row=r, column=1).value == "S-03")
    for k in range(1, 21):
        assert ws.cell(row=row, column=k + 1).value == sets["S-03"][k - 1].answer


def test_question_bank_sheet_holds_the_bank_as_uploaded(bank20, tmp_path):
    _, path = _papers_file(bank20, tmp_path)
    loaded = _load_question_bank_from_question_papers(str(path))
    for got, want in zip(loaded.get_all(), bank20.get_all()):
        assert (got.question_no, got.question_text, got.option_a, got.option_b,
                got.option_c, got.option_d, got.answer) == (
            want.question_no, want.question_text, want.option_a, want.option_b,
            want.option_c, want.option_d, want.answer)
    assert len(loaded.get_all()) == 20


def test_variant_papers_print_options_in_set_order(bank20):
    sets = build_same_question_sets(bank20, 3, seed=5)
    variants, matrix = variant_papers(bank20, sets)
    bank = {q.question_no: q for q in bank20.get_all()}
    assert [len(ids) for ids in matrix] == [20, 20, 20]
    p = sets["S-02"][0]
    v = variants.get_by_id(matrix[1][0])
    assert v.question_id == f"S-02:{bank[p.question_no].question_id}"
    assert v.answer == p.answer and v.difficulty == bank[p.question_no].difficulty
    assert v.option_a == getattr(bank[p.question_no], f"option_{p.slots[0].lower()}")


# ── Translating answers between positions and bank questions ─────────────────

from types import SimpleNamespace

import pandas as pd

from response_generator import (
    SET_COL, answer_column, generate_responses, to_bank_columns, to_positional_columns,
)


@pytest.fixture
def same_papers(bank20, tmp_path):
    sets, path = _papers_file(bank20, tmp_path, num_sets=6, seed=3)
    return SimpleNamespace(path=path, bank=bank20, sets=sets)


def _one_student(set_name, answers=None):
    row = {SET_COL: set_name}
    row.update({answer_column(i): None for i in range(1, 21)})
    row.update({answer_column(pos): letter for pos, letter in (answers or {}).items()})
    return pd.DataFrame([row])


def test_translation_round_trips(same_papers):
    df = generate_responses(str(same_papers.path), same_papers.bank, 6, seed=5)
    assert [c for c in df.columns if c.startswith("Q - ")] == [answer_column(i) for i in range(1, 21)]
    back = to_positional_columns(to_bank_columns(df, same_papers.sets), same_papers.sets)
    pd.testing.assert_frame_equal(back, df)


def test_typed_letter_maps_to_the_original_letter(same_papers):
    first = same_papers.sets["S-01"][0]
    bank_answer = next(q for q in same_papers.bank.get_all()
                       if q.question_no == first.question_no).answer
    out = to_bank_columns(_one_student("S-01", {1: first.answer}), same_papers.sets)
    assert out.loc[0, answer_column(first.question_no)] == bank_answer


def test_lowercase_and_spaces_are_read(same_papers):
    plain = to_bank_columns(_one_student("S-01", {1: "B"}), same_papers.sets)
    messy = to_bank_columns(_one_student("S-01", {1: " b"}), same_papers.sets)
    no = same_papers.sets["S-01"][0].question_no
    assert messy.loc[0, answer_column(no)] == plain.loc[0, answer_column(no)]


def test_blank_stays_blank(same_papers):
    out = to_bank_columns(_one_student("S-01"), same_papers.sets)
    assert out[[answer_column(i) for i in range(1, 21)]].isna().all().all()
    assert to_bank_columns(_one_student("S-01", {1: "E"}), same_papers.sets) \
        .loc[0, answer_column(same_papers.sets["S-01"][0].question_no)] == "E"


def test_unknown_set_is_refused(same_papers):
    with pytest.raises(ValueError, match="Unknown or missing Question Set"):
        to_bank_columns(_one_student("S-99"), same_papers.sets)


def test_extra_blank_columns_are_ignored_and_filled_ones_refused(same_papers):
    df = _one_student("S-01", {1: "A"})
    df["Q - 21 [Answer]"] = None
    out = to_bank_columns(df, same_papers.sets)
    assert "Q - 21 [Answer]" not in out.columns
    df.loc[0, "Q - 21 [Answer]"] = "A"
    with pytest.raises(ValueError, match=r"Q - 21 \[Answer\]"):
        to_bank_columns(df, same_papers.sets)
