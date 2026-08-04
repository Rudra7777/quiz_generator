# `S-01` is the set label everywhere

Students submit answers through a Google Form whose "Question Set" dropdown is populated
by hand from the generated question papers, so the set label is the join key between a
submission and the answer key it is scored against. We had two names for one concept —
`Set_1` internally, `S-01` on the Form — so we renamed the internal one. Sheet titles in
`question_papers.xlsx`, the `Answer_Key.Set` column, and every output sheet all emit
`S-01`, zero-padded to at least two digits.

## Considered Options

Normalising on read (`S-13` → `Set_13` on ingest, back to `S-13` on output) was the
obvious alternative and would have left existing papers working. We rejected it because
the mapping then lives in two places, and a mistake in either is invisible until marks
come out wrong for the wrong student — the worst possible failure mode for this tool.

## Consequences

Question papers generated before this change no longer match responses collected after
it. Regenerate the papers; do not hand-edit the sheet names, because the `Answer_Key`
sheet has to agree with them.

Set numbers are formatted as strings at every boundary rather than carried as integers.
That is deliberate — the label is the identity, and the padding is part of it.
