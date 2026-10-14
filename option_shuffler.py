"""
Option Shuffler

Reorders each question's A-D options so that the correct answer lands on every
letter equally often — 25% each — within each difficulty level. A bank whose key
leans on one letter lets students score by guessing that letter.

Works on the bank as a DataFrame in the question bank file format, so the result
can be saved and fed straight back into Generate Papers.
"""

import random
from itertools import permutations
from typing import Optional

import pandas as pd

from excel_handler import normalize_difficulty


LETTERS = "ABCD"
OPTION_COLS = ["option_a", "option_b", "option_c", "option_d"]


def _column_lookup(df: pd.DataFrame) -> dict:
    """Map normalized column names (as load_question_bank reads them) to the real ones."""
    return {str(c).strip().lower().replace(" ", "_"): c for c in df.columns}


def answer_counts(df: pd.DataFrame) -> pd.DataFrame:
    """Correct-answer counts: one row per difficulty (Hard/Medium/Easy), one column per letter."""
    cols = _column_lookup(df)
    difficulty = df[cols["difficulty"]].map(lambda v: normalize_difficulty(v).capitalize())
    answer = df[cols["answer"]].astype(str).str.strip().str.upper()
    counts = pd.crosstab(difficulty, answer).reindex(
        index=["Hard", "Medium", "Easy"], columns=list(LETTERS), fill_value=0
    )
    return counts.loc[counts.sum(axis=1) > 0]


def _avoid_current_letter(targets: list, current: list, rng: random.Random) -> None:
    """
    Swap targets in place until no question keeps its current answer letter.

    A swap with a question whose current and target letters are both different from
    the clash always exists unless one letter is more than 75% of the level's key;
    in that case the clashes that cannot be removed are left alone.
    """
    for i in range(len(targets)):
        if targets[i] != current[i]:
            continue
        swaps = [
            j for j in range(len(targets))
            if targets[j] != current[i] and current[j] != targets[i]
        ]
        if swaps:
            j = rng.choice(swaps)
            targets[i], targets[j] = targets[j], targets[i]


def _new_order(current: int, target: int, rng: random.Random) -> tuple:
    """
    Pick a new order for the four options: new_order[slot] is the old slot whose
    option goes there. The correct option goes from current to target and no other
    option stays where it was.
    """
    orders = [
        order for order in permutations(range(4))
        if order[target] == current
        and all(order[slot] != slot for slot in range(4) if slot != target)
    ]
    return rng.choice(orders)


def balance_options(df: pd.DataFrame, seed: Optional[int] = None) -> pd.DataFrame:
    """
    Return a copy of the bank with options reordered so the correct answers are
    spread evenly over A-D within each difficulty.

    Each difficulty gets floor(n/4) of every letter; its leftover 1-3 questions go to
    whichever letters the bank has fewest of so far, so the bank as a whole comes out
    even too. No question keeps its old answer letter and no distractor stays in its
    old place, unless one letter makes up more than 75% of a level's key — then a
    few answers have to stay put. Question text, numbers and difficulty are untouched.
    """
    cols = _column_lookup(df)
    missing = [c for c in OPTION_COLS + ["answer", "difficulty"] if c not in cols]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    rng = random.Random(seed)
    out = df.copy()
    option_cols = [cols[c] for c in OPTION_COLS]
    answer_col = cols["answer"]
    difficulty = df[cols["difficulty"]].map(normalize_difficulty)

    totals = {letter: 0 for letter in LETTERS}
    for level in ("hard", "medium", "easy"):
        rows = list(df.index[difficulty == level])
        if not rows:
            continue

        base, extra = divmod(len(rows), 4)
        # Leftovers go to the letters the bank is shortest on; ties broken at random.
        shortest = sorted(LETTERS, key=lambda letter: (totals[letter], rng.random()))
        targets = []
        for letter in LETTERS:
            share = base + (1 if letter in shortest[:extra] else 0)
            targets += [letter] * share
            totals[letter] += share
        rng.shuffle(targets)

        current = []
        for row in rows:
            letter = str(df.at[row, answer_col]).strip().upper()
            if letter not in LETTERS:
                raise ValueError(
                    f"Question at row {row + 2} has answer '{letter}', expected A/B/C/D."
                )
            current.append(letter)
        _avoid_current_letter(targets, current, rng)

        for row, letter, target in zip(rows, current, targets):
            options = [df.at[row, c] for c in option_cols]
            order = _new_order(LETTERS.index(letter), LETTERS.index(target), rng)

            for col, old_slot in zip(option_cols, order):
                out.at[row, col] = options[old_slot]
            out.at[row, answer_col] = target

    return out
