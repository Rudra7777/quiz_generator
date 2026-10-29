# Same-Questions Mode — Design

Date: 2026-10-29
Status: agreed with faculty, awaiting approval of this document

## Why

A faculty member asked for a way to write a small bank (say 20 questions) instead of
a large one. Every student gets all of the questions. What stops copying is that each
set has the questions in its own order and its own A–D options, so neighbours' Q5 is
a different question and "B" means different things on their papers.

The faculty also said this is not enough for large cohorts. Once one division has
sat the quiz, the questions can spread to the next. So the current big-bank method
stays, and the new method is a second mode chosen for each quiz.

## Goals

- A second generation mode, **Same Questions**, next to the current **Question Bank**
  mode, chosen on the Generate page.
- In Same Questions mode, every set holds every bank question, in a shuffled order
  and with shuffled options.
- Papers and the Google Form use the **paper's own numbering** (Q1…Qn). The bank
  number is not printed on the paper.
- Scoring checks each student against their own set's key.
- The scoring report shows answers both by question (the current Faculty Report,
  unchanged in behaviour) and as typed.

## Non-goals

- No change to Question Bank mode's output or behaviour.
- No change to the Streamlit UI (`app.py`) or the CLI (`main.py`); this goes into the
  Reflex UI, which is the deployed one. *(my call)*
- No difficulty mix in Same Questions mode. All questions are used, so the H/M/E
  counts and percentages don't apply and are hidden in that mode.

## Anti-goals (would count as failure even if it "works")

- A student is marked wrong because their letter was compared with another set's key,
  or with the bank's original letters, without translation.
- The bank number is visible on the printed paper, which lets students line up papers.
- Two sets come out identical.
- The current mode's papers or reports change in any way.

## How it works

### Generate (Part 1)

- The Generate page gets a mode choice: *Question Bank* (default, as now) and
  *Same Questions*.
- In Same Questions mode the inputs are the bank upload, number of sets (students),
  seed and optional case study. The number of questions is the bank size.
- For each set:
  1. **Question order:** a random permutation of all bank questions. No two sets may
     have the same order; if one repeats, draw again.
  2. **Options:** each question's four options are reordered for that set. Each set's
     key is balanced across A–D: floor(n/4) of every letter, with the 1–3 leftovers on
     randomly chosen letters. This reuses the balancing idea in `option_shuffler.py`,
     applied per set. Unlike the Shuffle page, this mode does not require that no option
     stays in its original place. *(my call)*
- The bank as uploaded is the reference. Its `answer` column is the "original letter"
  every translation goes back to.

### Papers file (output of Part 1)

Same file shape as now, with these differences in Same Questions mode:

- **S-01 … S-NN and All_Sets:** questions numbered by position, Q1…Qn, with options in
  that set's order. No bank number on the printed area.
- **Answer_Key:** one row per set and one column per position (Q1…Qn), giving that set's
  correct letter. (Today's sheet already uses this Set × Q-position layout.)
- **Set_Map (new):** one row per set per position: Set, Position, bank Question
  Number, and the bank's original letter shown in each of the A/B/C/D slots. Evaluate
  reads this to translate answers, and faculty can use it to check any single paper by hand.
- **Question_Bank:** the bank as uploaded (original options and key), as today.
- Allocation_Table / Shuffled_Table / Evaluation: leave out whatever is meaningless
  when every set has every question. The builder decides which to keep, but must
  not leave sheets that show misleading numbers.
- The presence of Set_Map is how Evaluate tells which mode a papers file was made in.

### Google Form

One answer column per position, **Q - 01 … Q - n**, plus the existing identity columns
and the Question Set dropdown. The headers look the same as today, but in this mode
the number means "position on my paper", not bank number.

### Evaluate (Part 2)

When the uploaded papers file has a Set_Map:

- Each submission's set is looked up. For each position, the student's letter is
  translated through Set_Map into the bank question number and the bank's original
  letter.
- **Scoring:** a translated answer is correct when it equals the bank key. This is the
  same as comparing the typed letter with that set's Answer_Key. Blanks are wrong, as
  today. Every question is on every paper, so there are no Extra Answers.
- **Faculty_Report:** built exactly as today from the translated answers. Columns are bank
  question numbers, rows 1–2 hold the bank key, live Count/AnsC formulas apply, editing a key
  letter recolours the sheet, and the Answered / Wrong / Wrong% rows rate each question. No
  yellow "outside your set" cells can appear.
- **As_Typed (new sheet):** one row per student with Roll Number, Name, Set, then their
  Q1…Qn letters exactly as submitted, each coloured right/wrong against that set's
  key. This is for disputes ("I wrote B for Q5").
- Scores / Summary / Validation: unchanged.
- A papers file without Set_Map goes down today's path, untouched.

### Dummy responses

The Evaluate page's dummy-response generator must produce Same Questions mode
responses (position columns Q1…Qn) when given such a papers file, so the full
pipeline can be rehearsed and tested.

## Constraints

- Set labels stay `S-01` style everywhere (docs/adr/0001).
- AnsC/Count formula ranges cover only real question columns (docs/adr/0002).
- Answer header parsing stays tolerant (`Q - 01 [Answer]`, `Q-01`, `Q01`, `Q1`).
- Each printed set still fits one A4 page.
- Existing tests keep passing. New tests cover: every set contains every question
  once; no duplicate sets; per-set key balance; translation round-trip (dummy responses
  → score equals the number of correct answers the generator intended); the bank-number
  column is absent from the printed paper.

## Settled after review *(my calls)*

- **Bank size:** no hard limit. A set uses the same layout rules as today (9pt, then
  8pt, and each set's own sheet is held to one page). A large bank may make All_Sets run
  past one page per set, exactly as a large set does in Question Bank mode.
- **Difficulty column:** stays required, so the bank format is the same in both modes.
  It is ignored for generation in this mode, but the Question_Bank summary grid still uses it.
- **As_Typed colouring:** fixed at scoring time, not live. A note at the top of the sheet
  says a key correction in Faculty_Report is not reflected here, so the scoring report
  has to be regenerated after fixing the bank.

## Options that refer to other options *(faculty: keep them in place)*

Some options only make sense in a fixed position. When generating sets:

- **"All of the above" / "None of the above" style** (also "all of these", "none of
  these", "all the above", and so on): that option stays in its original slot, and the
  other three are shuffled around it.
- **Options that name other letters** ("Both A and B", "A and C", "Only B", "A & B",
  …): none of that question's options are moved, because shuffling any of them would
  break the reference. *(my call)*
- Detection is by phrase matching on the option text, case-insensitive and ignoring
  punctuation. No new column is added to the bank.
- Per-set key balancing works around these questions: their answer letter is
  fixed, and the other questions are balanced as closely as the fixed ones allow.
- Generation reports how many questions were held in place, so faculty can check those
  few by eye.

## Left to the builder

- Module boundaries: where the per-set option shuffle and the translation live.
- Exact sheet styling of Set_Map and As_Typed, following existing sheet conventions.
- Which of Allocation_Table / Shuffled_Table / Evaluation to keep in this mode.
- Wording and layout of the mode choice on the Generate page.
- Updating README and CONTEXT.md vocabulary (e.g. defining "Position" as a term).
