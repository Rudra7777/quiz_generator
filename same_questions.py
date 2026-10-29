"""
Same Questions mode

Every student gets every question in the bank. What differs between sets is the
order of the questions and the order of each question's A-D options, so copying a
neighbour's letters does not work.

build_same_question_sets returns a Set Map: for each set, the questions in printed
order, with the bank's option letters in printed slot order and the letter that is
correct in that set. Options that only make sense in place ("All of the above",
"Both A and B") are held in their slots so they keep their meaning.
"""

import math
import random
import re
from dataclasses import dataclass
from itertools import permutations
from typing import Dict, FrozenSet, List, Optional, Tuple

import pandas as pd

from excel_handler import FullQuestion, FullQuestionBank, set_label


SET_MAP_SHEET = "Set_Map"

LETTERS = "ABCD"

# Only checked up to here; a bank of more questions never has too few orders.
MAX_FACTORIAL_CHECK = 10

# Option texts (normalised) that refer to the other options and so must not move.
REFERENCE_PHRASES = (
    "all of the above", "all the above", "none of the above", "none of these",
    "all of these", "both of the above", "neither of the above",
)

# A letter named next to one of these ("both a and b", "only c") refers to an option.
_LETTER_WORD = re.compile(r"\b[abcd]\b")
_LINK_WORD = re.compile(r"\b(and|or|only|both|neither)\b|&")


@dataclass(frozen=True)
class SetPosition:
    question_no: int   # bank Question Number at this position
    slots: str         # bank's original letter printed in slots A-D, e.g. "CABD"
    answer: str        # this set's correct letter at this position


# "S-01" -> positions 1..n in order
SetMap = Dict[str, List[SetPosition]]


def _normalise(text) -> str:
    """Lower case, punctuation (except &) removed, spaces collapsed."""
    text = re.sub(r"[^a-z0-9& ]", " ", str(text).lower())
    return " ".join(text.split())


def fixed_slots(q: FullQuestion) -> FrozenSet[int]:
    """
    The slots (0-3) whose option must stay where it is. If any option names a letter
    ("both A and B"), every slot is fixed; otherwise just the options such as
    "All of the above".
    """
    options = [_normalise(o) for o in (q.option_a, q.option_b, q.option_c, q.option_d)]
    if any(_LETTER_WORD.search(o) and _LINK_WORD.search(o) for o in options):
        return frozenset(range(4))
    return frozenset(
        slot for slot, text in enumerate(options)
        if any(re.search(rf"\b{phrase}\b", text) for phrase in REFERENCE_PHRASES)
    )


def held_in_place(bank: FullQuestionBank) -> List[int]:
    """Question numbers of the questions that have any fixed slot."""
    return [q.question_no for q in bank.get_all() if fixed_slots(q)]


def _target_letters(questions: List[FullQuestion], fixed: Dict[int, FrozenSet[int]],
                    rng: random.Random) -> Dict[int, int]:
    """
    Choose the slot (0-3) the correct option goes to for each question number.

    A question whose correct option is in a fixed slot keeps it. The rest are visited
    in random order, most restricted first so they are not left with only the slots the
    others have filled, and each takes the allowed slot used least so far, ties at
    random, so the key comes out with floor(n/4) of every letter.
    """
    targets, counts = {}, [0] * 4
    flexible = []
    for q in questions:
        current = LETTERS.index(q.answer.strip().upper())
        if current in fixed[q.question_no]:
            targets[q.question_no] = current
            counts[current] += 1
        else:
            flexible.append(q)

    rng.shuffle(flexible)
    flexible.sort(key=lambda q: -len(fixed[q.question_no]))
    for q in flexible:
        allowed = [s for s in range(4) if s not in fixed[q.question_no]]
        fewest = min(counts[s] for s in allowed)
        slot = rng.choice([s for s in allowed if counts[s] == fewest])
        targets[q.question_no] = slot
        counts[slot] += 1
    return targets


def build_same_question_sets(bank: FullQuestionBank, num_sets: int,
                             seed: Optional[int] = None) -> SetMap:
    """
    Build the Set Map: num_sets sets, each holding every question exactly once.

    No two sets share a question order. Within a set the correct answers are spread
    evenly over A-D, and each question's options are shuffled except the fixed ones.
    """
    questions = bank.get_all()
    n = len(questions)
    if n <= MAX_FACTORIAL_CHECK and num_sets > math.factorial(n):
        raise ValueError(
            f"A bank of {n} questions has only {math.factorial(n)} different orders; "
            f"asked for {num_sets} sets."
        )

    rng = random.Random(seed)
    fixed = {q.question_no: fixed_slots(q) for q in questions}
    used_orders = set()
    sets: SetMap = {}

    for number in range(1, num_sets + 1):
        order = questions[:]
        rng.shuffle(order)
        while tuple(q.question_no for q in order) in used_orders:
            rng.shuffle(order)
        used_orders.add(tuple(q.question_no for q in order))

        targets = _target_letters(questions, fixed, rng)
        positions = []
        for q in order:
            current = LETTERS.index(q.answer.strip().upper())
            target = targets[q.question_no]
            # order[slot] is the bank slot whose option is printed in that slot.
            choices = [
                p for p in permutations(range(4))
                if p[target] == current and all(p[s] == s for s in fixed[q.question_no])
            ]
            chosen = rng.choice(choices)
            positions.append(SetPosition(
                question_no=q.question_no,
                slots="".join(LETTERS[s] for s in chosen),
                answer=LETTERS[target],
            ))
        sets[set_label(number)] = positions

    return sets


def variant_papers(bank: FullQuestionBank,
                   set_map: SetMap) -> Tuple[FullQuestionBank, List[List[str]]]:
    """
    One FullQuestion per set per position, options in printed order and the answer
    the set's own letter, so the existing paper writer can print any set. Ids are
    f"{label}:{original_id}"; the matrix lists each set's ids in printed order.
    """
    by_no = {q.question_no: q for q in bank.get_all()}
    variants, matrix = [], []
    for label, positions in set_map.items():
        ids = []
        for p in positions:
            q = by_no[p.question_no]
            options = {"A": q.option_a, "B": q.option_b, "C": q.option_c, "D": q.option_d}
            a, b, c, d = (options[letter] for letter in p.slots)
            variant = FullQuestion(
                question_no=q.question_no, question_id=f"{label}:{q.question_id}",
                question_text=q.question_text, option_a=a, option_b=b, option_c=c,
                option_d=d, answer=p.answer, difficulty=q.difficulty,
            )
            variants.append(variant)
            ids.append(variant.question_id)
        matrix.append(ids)
    return FullQuestionBank(variants), matrix


def read_set_map(papers_path: str) -> Optional[SetMap]:
    """Read the Set_Map sheet back from a papers file; None if it has none."""
    try:
        raw = pd.read_excel(papers_path, sheet_name=SET_MAP_SHEET, header=None)
    except ValueError:
        return None   # no such sheet

    header_row = next(
        (i for i, row in raw.iterrows()
         if str(row[0]).strip() == "Set" and str(row[1]).strip() == "Position"),
        None,
    )
    if header_row is None:
        return None

    set_map: SetMap = {}
    for _, row in raw.iloc[header_row + 1:].iterrows():
        if pd.isna(row[0]):
            continue
        slots = "".join(str(row[i]).strip() for i in (3, 4, 5, 6))
        set_map.setdefault(str(row[0]).strip(), []).append(
            SetPosition(int(row[2]), slots, str(row[7]).strip())
        )
    return set_map
