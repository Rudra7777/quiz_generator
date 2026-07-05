# 📝 Quiz Question Paper Generator

A web-based quiz allocation system that generates randomized question papers with optimal question distribution.

## Features

- 📤 Upload Excel question bank
- 🎯 Configure difficulty distribution (Absolute or Percentage mode)
- ⚖️ Fair question allocation using greedy load-balancing
- 📥 Download formatted Excel with multiple sheets
- 📊 Built-in evaluation metrics

## Live Demo

🚀 [View App on Streamlit Cloud](#) *(Add your URL after deployment)*

## Local Setup

```bash
pip install -r requirements.txt
```

Two web UIs are available (both share the same engine and `excel_export.py` helpers):

**Reflex UI (recommended)** — sidebar multi-page app with Generate / Evaluate pages:

```bash
reflex run
```

**Streamlit UI** — original two-tab app:

```bash
streamlit run app.py
```

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
- **Set_1 to Set_N**: Individual question papers (no answers)
- **Answer_Key**: Correct answers for teachers
- **Allocation_Table**: Original allocation by difficulty
- **Shuffled_Table**: Randomized order per student
- **Evaluation**: Usage statistics and metrics

## Sample Files

- `input/question_bank.xlsx` - 70 questions (12H, 30M, 28E)
- `input/question_bank_72.xlsx` - 72 questions (12H, 30M, 30E)

## Tech Stack

- Python 3.8+
- Streamlit
- Pandas
- OpenPyXL
- NumPy
