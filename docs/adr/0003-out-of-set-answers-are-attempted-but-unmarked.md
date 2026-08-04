# Out-of-set answers are attempted but never marked

The Google Form shows every question in the bank, and the student's paper tells them which
subset to answer. Some answer questions that were not on their paper. Those answers count
towards `Attempted` / `Count` — the student did the work — but earn nothing, right or
wrong, because a student must not be able to gain marks by answering questions that were
never assigned to them. They render in purple, distinct from blue (correct) and red
(wrong), so "attempted but unmarked" is visible at a glance.

## Consequences

`AnsC` cannot be a plain comparison of the student's row against the key row: that
formula has no way to tell an assigned question from an extra one, and would hand out
marks for correct extras. Verified before this change — a student with 15 assigned, 9
genuinely correct and 2 correct extras had `Scores.Correct` of 9 but `AnsC` of 11.

So the Faculty_Report carries a **hidden helper block** to the right of the answers, one
1/0 per question per student meaning "was this on their paper". `AnsC` multiplies by it:

```
=SUMPRODUCT(--(H$2:CA$2=H4:CA4), CC4:EV4)
```

and the conditional formatting reads the same block, which is what makes purple possible
at all. Do not delete those columns because they look like clutter — `AnsC` silently
starts over-scoring if the second argument goes away.

The block is what keeps `AnsC` both live (editing a key letter in row 2 still
recalculates) and correct. The alternative was writing `AnsC` as a static number from
Python, which would have been simpler but would have broken the live key editing chosen
in [0002](./0002-ansc-counts-only-real-question-columns.md).
