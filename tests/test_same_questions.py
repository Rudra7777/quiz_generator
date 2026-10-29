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
