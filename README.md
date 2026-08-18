# 📝 Quiz Question Paper Generator

A web-based quiz allocation system that generates randomized question papers with optimal question distribution.

## Features

- 📤 Upload Excel question bank
- 🎯 Configure difficulty distribution (Absolute or Percentage mode)
- ⚖️ Fair question allocation using greedy load-balancing
- 📥 Download formatted Excel with multiple sheets
- 📊 Built-in evaluation metrics

## Live Demo

🚀 [View App on Reflex Cloud](https://quiz-web-silver-moon.reflex.run)

## Local Setup

Requires Python 3.10–3.12 (3.12 recommended — 3.13+ will fail to install).

```bash
pip install -r requirements.txt
```

Setting up on a fresh machine, especially Windows? See `INSTALL.txt` for the
step-by-step version and the errors that usually come up.

Two web UIs are available (both share the same engine and `excel_export.py` helpers):

**Reflex UI (recommended)** — sidebar multi-page app with Generate / Evaluate pages:

```bash
reflex run
```

**Streamlit UI** — original two-tab app:

```bash
streamlit run app.py
```

## Deployment

The Reflex UI is deployed to [Reflex Cloud](https://build.reflex.dev), which hosts both the frontend and backend for you (no Vercel/servers needed).

```bash
reflex login    # one-time, opens a browser to authenticate
reflex deploy --app-name quiz-web --project 67b7e9f1-6dba-4838-9eec-129ecd156331
```

This rebuilds the app and pushes it to the same live URL above. Currently on the Free tier (1 CPU / 1GB RAM, up to 5 deployments).

## Automated Testing

Run the answer-checking integration tests locally:

```bash
pytest -q
```

CI runs the same tests on every push/PR via:
- `/Users/rudrapatole/Desktop/quiz/.github/workflows/tests.yml`

### Randomization Behavior

- Streamlit app: default is fresh random seed on each generation.
- CLI (`main.py`): runs are random by default; pass `--seed <number>` for reproducible output.

## Question Bank Format

Your Excel file should have these columns:
- `question_no`: Unique question number
- `question`: Question text
- `option_a`, `option_b`, `option_c`, `option_d`: Answer options
- `answer`: Correct answer (A/B/C/D)
- `difficulty`: H/M/L or Hard/Medium/Easy

## Output

The generated Excel contains:
- **All_Sets**: Every set stacked one below another, set up to print on A4 portrait
  with each set starting on a fresh page. This is the sheet to print.
- **S-01 to S-NN**: Individual question papers (no answers)
- **Answer_Key**: Correct answers for teachers
- **Allocation_Table**: Original allocation by difficulty
- **Shuffled_Table**: Randomized order per student
- **Evaluation**: Usage statistics and metrics

Set labels are `S-01`, `S-02`, … everywhere — question paper sheets, the answer key, and
the Google Form's "Question Set" dropdown all use the same string, so responses join
straight back to the right answer key. See `docs/adr/0001`.

## Student Response Format

Part 2 scores the Google Forms export directly. It needs these columns:

| Column | Purpose |
| --- | --- |
| `Timestamp` | Picks the winner when a student submits twice |
| `Email address` | Carried through to the reports |
| `Full Name` | Carried through to the reports |
| `Roll Number` | Student identity; duplicate submissions are deduplicated on it |
| `Question Set` | Joins the row to its answer key (`S-01` style) |
| `Q - 01 [Answer]` … | One column per **question bank number**, blank unless that question was on the student's paper |

Answer cells hold a bare `A`/`B`/`C`/`D`. Answer column headers are read tolerantly —
`Q - 01 [Answer]`, `Q-01`, `Q01` and `Q1` all mean question 1.

If a Roll Number submits more than once, only the latest submission is scored; the
dropped ones are listed on the `Validation` sheet.

## Scoring Report

`scoring_report.xlsx` contains:
- **Scores**: per student — Roll Number, Name, Email, Set, Assigned, Attempted, Correct, Wrong
- **Summary**: cohort averages, pass rate, and mark distribution in columns A–B, plus a live
  Excel chart — a mark-by-mark histogram with a fitted normal curve over it. Chart source
  data sits in columns D–F.
- **Validation**: duplicate submissions, and anyone who answered outside their set
- **Faculty_Report**: the layout the faculty asked for — answer key strip on rows 1–2, then
  one row per student with live `Count` and `AnsC` formulas (see `docs/adr/0002`).
  Answers are colour-coded on two axes: the text says right or wrong against the key
  (**blue** / ***red italic***), the background says whose question it was — light green
  across the whole of the student's own set, answered or not, so the ones they skipped
  show as empty green cells, and light yellow where they answered a question that was not
  on their paper (see `docs/adr/0003`). Conditional
  formatting, so fixing a key letter in row 2 recolours the sheet and recomputes the marks
  together. Hidden helper columns to the right of the answers drive both — don't delete them.
- **Responses_Review**: every answer cell colour-coded green/red

## Sample Files

- `input/question_bank.xlsx` - 70 questions (12H, 30M, 28E)
- `input/question_bank_72.xlsx` - 72 questions (12H, 30M, 30E)

## Tech Stack

- Python 3.8+
- Streamlit
- Pandas
- OpenPyXL
- NumPy
