# `AnsC` counts only real question columns

The faculty's original spreadsheet computes `AnsC` as
`{=SUMPRODUCT(--(H$2:CN$2=H4:CN4))}`. The `H:CN` range is 85 columns wide but only 27
carry data, so the 65 trailing blank-vs-blank column pairs each count as a match and
every score is inflated by exactly 65 — a student with 4 correct answers reads 69. We
emit the same array formula ranged to the actual question columns, so `AnsC` is the true
number correct.

## Consequences

**Our numbers will not match the sheet the faculty sent.** Theirs read 69/69/67/71;
ours read 4/4/2/6 for the same students. This is not a bug and should not be "fixed" by
widening the range back to `H:CN` — check the offset is 65 before assuming otherwise.

`AnsC` stays a live formula rather than a value written by `answer_checker.py` so that
editing the answer key in row 2 recalculates the whole column in Excel. It should always
agree with the `Correct` column on the `Scores` sheet; if it does not, the key row and
the question bank have diverged.
