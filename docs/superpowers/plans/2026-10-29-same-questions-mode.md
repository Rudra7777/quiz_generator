# Same-Questions Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a second generation mode where every set holds every bank question, with a per-set question order and per-set option order, and score it correctly.

**Architecture:** A new `same_questions.py` builds a **Set Map** (for each set: the bank question at each position, and which original option sits in each printed slot). Paper export turns the Set Map into per-set copies of the questions ("variants") and reuses today's set-sheet, All_Sets and Answer_Key writers, then adds a `Set_Map` sheet. On the scoring side, positional answers are translated back to bank question numbers and original letters before scoring, so the existing scorer and Faculty_Report run unchanged. A new As_Typed sheet shows the answers exactly as typed.

**Tech Stack:** Python 3.12, pandas, openpyxl, Reflex, pytest (`venv/bin/python -m pytest -q`).

**Spec:** `docs/superpowers/specs/2026-10-29-same-questions-mode-design.md`

## Global Constraints

- Question Bank mode output and behaviour must not change: all 72 existing tests stay green, untouched.
- Set labels are `S-01` style everywhere (docs/adr/0001).
- Answer headers stay `Q - 01 [Answer]`, read tolerantly (`Q-01`, `Q01`, `Q1`).
- No bank number visible on the printed area of a Same Questions paper.
- Each set's key is balanced: floor(n/4) of each letter, the 1–3 leftovers on random letters, adjusted only where held-in-place questions force a letter.
- No two sets identical.
- Bank file format is unchanged in both modes (difficulty column still required).
- Reflex UI only; `app.py` and `main.py` are not touched.
- Mode is detected from the presence of a `Set_Map` sheet in the papers file.

## Review Focus

- Typed answers in lower case or with spaces (`" b"`) are read as `B`, same as today. Pinned in Task 3.
- A blank position counts as wrong and stays blank in As_Typed. Pinned in Tasks 3 and 4.
- A response sheet that still carries answer columns beyond n (a form reused from a bigger quiz): blank extra columns are ignored, a filled one raises a clear error naming the column. Pinned in Task 3.
- A small bank asked for more sets than it has orders (3 questions, 7 sets) raises a clear error rather than looping forever. Pinned in Task 1.
- A question where all four options are held and its letter can't be balanced still generates, and the other questions absorb the balance. Pinned in Task 1.

---

### Task 1: Set Map generation (`same_questions.py`)

**Files:**
- Create: `same_questions.py`
- Create: `tests/test_same_questions.py`
- Modify: `CONTEXT.md` (add **Same Questions mode**, **Position** and **Set Map** to the glossary, in the existing style)

**Interfaces:**
- Consumes: `FullQuestion`, `FullQuestionBank`, `set_label` from `excel_handler`.
- Produces:
  ```python
  SET_MAP_SHEET = "Set_Map"

  @dataclass(frozen=True)
  class SetPosition:
      question_no: int   # bank Question Number at this position
      slots: str         # bank's original letter printed in slots A-D, e.g. "CABD"
      answer: str        # this set's correct letter at this position

  SetMap = Dict[str, List[SetPosition]]   # "S-01" -> positions 1..n in order

  def fixed_slots(q: FullQuestion) -> FrozenSet[int]
  def held_in_place(bank: FullQuestionBank) -> List[int]          # question_nos with any fixed slot
  def build_same_question_sets(bank: FullQuestionBank, num_sets: int, seed: Optional[int] = None) -> SetMap
  ```

- [ ] **Step 1: Write failing tests** in `tests/test_same_questions.py`, with a fixture `bank20` that builds a 20-question `FullQuestionBank` in code (numbers 1–20, mixed difficulties, answers mostly `A`). Q5 has option D `"All of the above"` and is the answer. Q9 has option C `"None of these"` and is not the answer. Q12 has option D `"Both A and B"`.

```python
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
    bank3 = ...  # 3 questions, no held options
    with pytest.raises(ValueError, match="only 6 different orders"):
        build_same_question_sets(bank3, 7, seed=0)
```

- [ ] **Step 2: Run them to see them fail**

Run: `venv/bin/python -m pytest tests/test_same_questions.py -v`
Expected: FAIL, `ModuleNotFoundError: same_questions`.

- [ ] **Step 3: Implement `same_questions.py`**

- `fixed_slots`: normalise each option (lower case, punctuation stripped, spaces collapsed). If any option names a letter as a word, return all four slots. That means it matches `\b[abcd]\b` together with "and", "or", "&", "only", "both", or "neither", as in "both a and b", "a and c", "only b", "a & b". Otherwise return the slots whose text matches "all of the above", "all the above", "none of the above", "none of these", "all of these", "both of the above", or "neither of the above". Keep the phrases in a module constant so they can be extended.
- `build_same_question_sets`: one `random.Random(seed)` for everything. For each set:
  1. Shuffle the question order. Redraw while the order is already used. Before the loop, raise `ValueError(f"A bank of {n} questions has only {math.factorial(n)} different orders; asked for {num_sets} sets.")` when `num_sets > math.factorial(n)`. Only check that for n ≤ 10, because larger factorials are never exceeded.
  2. Choose the target letters. A question whose answer slot is fixed keeps its letter. Each other question can go to any letter not taken by a fixed slot. Visit the flexible questions in random order and give each the allowed letter with the lowest count so far, breaking ties at random. This gives floor(n/4) of each letter with random leftovers, and the tolerance-1 test holds.
  3. For each question, pick uniformly from the `itertools.permutations(range(4))` that put the correct option at its target and leave every fixed slot where it is. Write `slots` as the original letters in printed order. Unlike `option_shuffler._new_order`, other options are allowed to stay in place.

- [ ] **Step 4: Run the tests and see them pass**

Run: `venv/bin/python -m pytest tests/test_same_questions.py -v`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add same_questions.py tests/test_same_questions.py CONTEXT.md
git commit -m "Build per-set question and option orders for Same Questions mode"
```

---

### Task 2: Same Questions papers file

**Files:**
- Modify: `excel_export.py`. Lift the Answer_Key block (the `if include_answer_key:` section) into `_write_answer_key_sheet(wb, shuffled_matrix, question_bank)`, and the Question_Bank block into `_write_question_bank_sheet(wb, question_bank) -> dict` (returns the summary cache). `create_formatted_excel` calls both and its output is unchanged. Add `show_code: bool = True` to `_write_set_block` and `_write_all_sets_sheet`. When it is False, the column C header and cells are left empty, and column B (hidden) still holds the bank number under its `QCd` header. That way `response_generator._attach_bank_no` and `map_paper_to_bank_questions` read each variant sheet's bank numbers unchanged, and Task 4 relies on this.
- Modify: `same_questions.py`. Add `variant_papers` and `read_set_map`.
- Test: `tests/test_same_questions.py`

**Interfaces:**
- Consumes: `SetMap`, `SetPosition`, `SET_MAP_SHEET`, `build_same_question_sets` (Task 1).
- Produces:
  ```python
  # same_questions.py
  def variant_papers(bank: FullQuestionBank, set_map: SetMap) -> Tuple[FullQuestionBank, List[List[str]]]
      # per-set FullQuestion copies, id f"{label}:{original_id}", options in printed order,
      # answer = SetPosition.answer, question_no/difficulty unchanged; plus the id matrix
  def read_set_map(papers_path: str) -> Optional[SetMap]   # None when the sheet is absent

  # excel_export.py
  def create_same_questions_excel(
      set_map: SetMap, question_bank: FullQuestionBank,
      case_paragraphs: Sequence[str] = (), seed: Optional[int] = None,
  ) -> bytes
  ```
- Sheets, in order: `All_Sets`, `S-01`…`S-NN`, `Answer_Key`, `Set_Map`, `Question_Bank`. Allocation_Table, Shuffled_Table and Evaluation are left out, because Set_Map replaces Shuffled_Table and usage is identical for every question.
- `Set_Map` layout: title in A1 (`"Set Map (which bank question and option sits where)"`). The seed goes in A2 when one is given. Headers on row 3 are `Set | Position | Question No | A | B | C | D | Answer`. After that there is one row per set per position: the A–D cells hold the bank's original letter shown in that slot, and Answer is the set's letter. `read_set_map` finds the header row by its `Set`/`Position` cells instead of assuming row 3.

- [ ] **Step 1: Write failing tests**

```python
def test_same_questions_file_has_the_expected_sheets(bank20, tmp_path):
    data = create_same_questions_excel(build_same_question_sets(bank20, 4, seed=1), bank20, seed=1)
    wb = load_workbook(io.BytesIO(data))
    assert wb.sheetnames == ["All_Sets", "S-01", "S-02", "S-03", "S-04",
                             "Answer_Key", "Set_Map", "Question_Bank"]

def test_set_map_round_trips(bank20, tmp_path):
    sets = build_same_question_sets(bank20, 4, seed=1)
    path = tmp_path / "papers.xlsx"; path.write_bytes(create_same_questions_excel(sets, bank20))
    assert read_set_map(str(path)) == sets

def test_old_papers_have_no_set_map(tmp_path):
    # a Question Bank mode file from create_formatted_excel
    assert read_set_map(str(old_papers_path)) is None

def test_paper_prints_positions_and_shuffled_options_without_bank_numbers(bank20, tmp_path):
    # For S-02: column A is 1..20; column C is empty on every row (header included);
    # the question in row k is bank question sets["S-02"][k-1].question_no, and its
    # "A] ..." cell is the bank option named by slots[0].

def test_answer_key_is_per_set_by_position(bank20, tmp_path):
    # Answer_Key row for S-03, column Qk == sets["S-03"][k-1].answer

def test_question_bank_sheet_holds_the_bank_as_uploaded(bank20, tmp_path):
    # _load_question_bank_from_question_papers returns the original options and answers

def test_existing_mode_output_is_unchanged():
    # Existing test_excel_export.py tests stay green. No new assertion needed beyond running them.
```

- [ ] **Step 2: Run them to see them fail**

Run: `venv/bin/python -m pytest tests/test_same_questions.py -v`
Expected: the new tests FAIL with ImportError on `create_same_questions_excel`.

- [ ] **Step 3: Do the two extractions in `excel_export.py`. Then run the whole suite to prove Question Bank mode is unchanged**

Run: `venv/bin/python -m pytest -q`
Expected: all 72 existing tests pass (the new ones still fail).

- [ ] **Step 4: Implement `variant_papers`, `read_set_map`, `create_same_questions_excel` and the `show_code` flag**

`create_same_questions_excel` builds the variant bank and matrix. It then calls `_plan_set_layout`, the set-sheet loop, `_write_answer_key_sheet` (with the variant bank), the Set_Map writer, and `_write_question_bank_sheet` (with the **original** bank). Last come `_write_all_sets_sheet`, the save, and `_inject_cached_values`, exactly as `create_formatted_excel` does.

- [ ] **Step 5: Run the whole suite**

Run: `venv/bin/python -m pytest -q`
Expected: all PASS.

- [ ] **Step 6: Commit**

```bash
git add excel_export.py same_questions.py tests/test_same_questions.py
git commit -m "Write Same Questions papers with a Set_Map sheet"
```

---

### Task 3: Translating answers between positions and bank questions

**Files:**
- Modify: `response_generator.py`. This file owns the response-sheet shape, so the two translations live here, and `generate_responses` emits positional columns when the papers carry a Set_Map.
- Test: `tests/test_same_questions.py`

**Interfaces:**
- Consumes: `SetMap`, `read_set_map` (Tasks 1–2), the existing `answer_column`, `ANSWER_COL_RE`, `SET_COL`, `IDENTITY_COLS`.
- Produces:
  ```python
  def to_bank_columns(response_df: pd.DataFrame, set_map: SetMap) -> pd.DataFrame
      # identity columns kept; answer columns become one per bank question_no
      # (answer_column(no)), holding the bank's ORIGINAL letter
  def to_positional_columns(response_df: pd.DataFrame, set_map: SetMap) -> pd.DataFrame
      # exact inverse: Q - 01 … Q - n holding the letter as printed on that set
  ```
- Rules:
  - A letter is upper-cased and stripped first, the same as `answer_checker._normalize_answer`.
  - A–D are translated through `slots`. Anything else, blanks included, passes through unchanged, so a blank stays blank.
  - An unknown set raises the same "Unknown or missing Question Set" `ValueError` the scorer raises today.
  - A position column numbered above n raises `ValueError(f"'{col}' is beyond the {n} questions on each paper")` if any row fills it, and is dropped if it's blank in every row.
- `generate_responses`: when `read_set_map(question_papers_path)` is not None, build the bank-number responses as today, then return `to_positional_columns(df, set_map)`. Today's `range(1, total_questions + 1)` column loop is fine here, because every bank question is on every paper.

- [ ] **Step 1: Write failing tests**

```python
def test_translation_round_trips(same_papers):            # fixture: 20 Qs, 6 sets, papers file on disk
    df = generate_responses(str(same_papers.path), same_papers.bank, 6, seed=5)
    assert [c for c in df.columns if c.startswith("Q")] == [answer_column(i) for i in range(1, 21)]
    back = to_positional_columns(to_bank_columns(df, same_papers.sets), same_papers.sets)
    pd.testing.assert_frame_equal(back, df)

def test_typed_letter_maps_to_the_original_letter(same_papers):
    # Take S-01 position 1: type the letter whose slot holds the bank answer;
    # to_bank_columns gives the bank answer under answer_column(that question_no).

def test_lowercase_and_spaces_are_read(same_papers):
    # " b" at a position translates the same as "B"

def test_blank_stays_blank(same_papers):

def test_extra_blank_columns_are_ignored_and_filled_ones_refused(same_papers):
    # add an all-blank "Q - 21 [Answer]" -> dropped; fill one cell -> ValueError naming "Q - 21 [Answer]"
```

- [ ] **Step 2: Run them to see them fail**

Run: `venv/bin/python -m pytest tests/test_same_questions.py -v`
Expected: the new tests FAIL with ImportError on `to_bank_columns`.

- [ ] **Step 3: Implement both functions and the `generate_responses` branch**

- [ ] **Step 4: Run the whole suite**

Run: `venv/bin/python -m pytest -q`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add response_generator.py tests/test_same_questions.py
git commit -m "Translate Same Questions answers between positions and bank questions"
```

---

### Task 4: Scoring and the As_Typed sheet

**Files:**
- Modify: `answer_checker.py`.
  - In `check_all_responses`, if `read_set_map(question_papers_path)` is not None, run `response_df = to_bank_columns(response_df, set_map)` before dedup. Everything after that is unchanged: Set sheets carry bank numbers in hidden column B, so `map_paper_to_bank_questions` already returns all n bank numbers per set, and the key comes from Question_Bank.
  - In `generate_scoring_report`, when a Set_Map is present, also write `As_Typed` after Faculty_Report.
- Test: `tests/test_same_questions.py`

**Interfaces:**
- Consumes: `read_set_map`, `to_bank_columns`, `to_positional_columns` (Tasks 2–3).
- Produces: `_write_as_typed(writer, scored_df: pd.DataFrame, set_map: SetMap) -> None`
  - Row 1: the note `"Coloured against each set's key when this report was made. A key fix in Faculty_Report is not reflected here; regenerate the report after correcting the bank."`
  - Row 3: the headers `Roll Number | Full Name | Set | Q1 … Qn | Correct`.
  - Row 4 onwards: one row per scored submission, in Faculty_Report's order (by set), with the typed letters from `to_positional_columns(report.scored_df, set_map)`.
  - Each answer cell uses Faculty_Report's right/wrong font styles (blue / red italic). A blank is left blank with no fill.
  - Correct is a plain number equal to that student's `Correct` on Scores.

- [ ] **Step 1: Write failing tests**

```python
def test_scores_match_what_the_generator_intended(same_papers, tmp_path):
    df = generate_responses(str(same_papers.path), same_papers.bank, 6, correct_rate=0.6, seed=7)
    report = check_all_responses(df, str(same_papers.path), same_papers.bank)
    for r, (_, row) in zip(report.student_reports, df.iterrows()):
        key = [p.answer for p in same_papers.sets[row["Question Set"]]]
        typed = [row[answer_column(i)] for i in range(1, 21)]
        assert r.correct == sum(t == k for t, k in zip(typed, key))
        assert r.assigned == 20 and r.validation.extra_count == 0

def test_blank_position_is_wrong(same_papers):

def test_faculty_report_is_by_bank_question_with_the_bank_key(same_papers, tmp_path):
    # rows 1-2 strip = bank numbers 1..20 and bank answers; no yellow (extra-answer) fills

def test_as_typed_sheet(same_papers, tmp_path):
    # sheet exists; row 1 note; headers Q1..Q20; a student's cells equal what they typed;
    # Correct column equals Scores' Correct

def test_old_papers_have_no_as_typed_sheet(tmp_path):
```

- [ ] **Step 2: Run them to see them fail**

Run: `venv/bin/python -m pytest tests/test_same_questions.py -v`
Expected: the new tests FAIL. Scores come out wrong because translation isn't wired in, and As_Typed is missing.

- [ ] **Step 3: Implement**

- [ ] **Step 4: Run the whole suite**

Run: `venv/bin/python -m pytest -q`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add answer_checker.py tests/test_same_questions.py
git commit -m "Score Same Questions responses and add the As_Typed sheet"
```

---

### Task 5: Generate page mode choice

**Files:**
- Modify: `quiz_web/state.py`. In `GenerateState`:
  - Add `paper_mode: str = "Question Bank"` (or `"Same Questions"`) and a `set_paper_mode` setter.
  - `can_generate` becomes `bank_uploaded and (paper_mode == "Same Questions" or (size_ok and avail_ok))`.
  - Add an `rx.var` `bank_size -> int` (the same as `total_avail`).
  - In `generate_papers`, when `paper_mode == "Same Questions"`: call `build_same_question_sets(bank, self.num_students, run_seed)` and `create_same_questions_excel(set_map, bank, self._case_paragraphs, run_seed)`. The case study works in both modes. The status becomes `f"Generated {n} question papers, each with all {bank size} questions."`, followed by `f" {k} question(s) kept their options in place: Q{a}, Q{b}."` when `held_in_place(bank)` is not empty. `ValueError` from generation shows through the existing `error`.
- Modify: `quiz_web/pages/generate.py`. Add a mode choice (segmented control or radio, following the page's existing controls) above the configuration step. In Same Questions mode, hide the total-questions input and the difficulty step, and show `"Every set gets all {bank_size} questions, in its own order and with its own options."`
- Modify: `README.md`. Add a "Same Questions mode" section after "Balancing the Answer Key". It should say what the mode does, that papers show Q1…Qn with no bank numbers, that the Google Form needs answer columns Q - 01 … Q - n plus Question Set, what the Set_Map and As_Typed sheets are, and which options are held in place.

**Interfaces:**
- Consumes: `build_same_question_sets`, `held_in_place` (Task 1), `create_same_questions_excel` (Task 2).

- [ ] **Step 1: Implement the state and page changes**

- [ ] **Step 2: Run the whole suite**

Run: `venv/bin/python -m pytest -q`
Expected: all PASS.

- [ ] **Step 3: Verify in the running app**

Start with `reflex run` through a `.claude/launch.json` entry, then open the Generate page in the browser pane.
- Upload a 20-question bank (make one from `input/question_bank.xlsx`'s first 20 rows), choose Same Questions and 65 sets, then generate. Expected: the status says 65 papers with all 20 questions, and the download has the sheets from Task 2.
- Switch back to Question Bank. Expected: the difficulty controls return and generation works as before.
- On Evaluate, generate dummy responses from that papers file and score them. Expected: the report has Faculty_Report and As_Typed, and the marks are out of 20.

- [ ] **Step 4: Commit**

```bash
git add quiz_web/state.py quiz_web/pages/generate.py README.md
git commit -m "Add Same Questions mode to the Generate page"
```
