"""
Quiz Generator - Streamlit UI

A web interface for generating randomized quiz papers from a question bank.
Upload an Excel file, configure settings, and download formatted question papers.
"""

import streamlit as st
import os
import secrets
import tempfile
from typing import Dict, List

from allocator import QuizStructure, allocate_quizzes, shuffle_all_quizzes
from excel_handler import load_question_bank
from response_generator import generate_responses
from answer_checker import (
    load_response_sheet,
    check_all_responses,
    generate_scoring_report,
)
from excel_export import (
    create_formatted_excel,
    _make_excel_bytes_from_dataframe,
    _load_question_bank_from_question_papers,
)


# Page configuration
st.set_page_config(
    page_title="Quiz Generator",
    page_icon="📝",
    layout="wide"
)


def _save_uploaded_temp(uploaded_file, prefix: str) -> str:
    """Persist uploaded Streamlit file to a temporary .xlsx path."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix=prefix) as tmp:
        tmp.write(uploaded_file.getvalue())
        return tmp.name


def _render_generation_tab():
    """Part 1 UI: generate question papers."""
    st.markdown("Upload a question bank and generate randomized question papers for all students.")
    st.divider()

    st.header("1️⃣ Upload Question Bank")
    uploaded_file = st.file_uploader(
        "Upload Excel file (.xlsx)",
        type=["xlsx"],
        help="Excel with columns: question_no, question, option_a, option_b, option_c, option_d, answer, difficulty",
        key="part1_question_bank",
    )

    question_bank = None
    counts = {"hard": 0, "medium": 0, "easy": 0}

    if uploaded_file:
        temp_path = None
        try:
            temp_path = _save_uploaded_temp(uploaded_file, "part1_qb_")
            question_bank = load_question_bank(temp_path)
            counts = question_bank.count_by_difficulty()
            total = sum(counts.values())

            st.success(f"✅ Loaded {total} questions successfully!")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Hard", counts.get("hard", 0))
            col2.metric("Medium", counts.get("medium", 0))
            col3.metric("Easy", counts.get("easy", 0))
            col4.metric("Total", total)

            st.session_state["part1_question_bank_bytes"] = uploaded_file.getvalue()
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
            question_bank = None
        finally:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)

    st.divider()

    st.header("2️⃣ Configuration")
    col1, col2 = st.columns(2)
    with col1:
        num_students = st.number_input(
            "Number of Students",
            min_value=1,
            max_value=500,
            value=50,
            help="How many question papers to generate",
            key="part1_num_students",
        )
    with col2:
        total_questions = st.number_input(
            "Questions per Quiz",
            min_value=1,
            max_value=100,
            value=15,
            help="Total questions in each student's quiz",
            key="part1_total_questions",
        )

    st.divider()
    st.header("3️⃣ Difficulty Distribution")
    mode = st.radio(
        "Select mode:",
        ["Absolute (exact counts)", "Percentage (auto-calculate)"],
        horizontal=True,
        key="part1_mode",
    )

    hard_count = medium_count = easy_count = 0
    if mode == "Absolute (exact counts)":
        col1, col2, col3 = st.columns(3)
        with col1:
            hard_count = st.number_input(
                "Hard Questions",
                min_value=0,
                max_value=total_questions,
                value=min(4, counts.get("hard", 4)),
                key="part1_hard_abs",
            )
        with col2:
            medium_count = st.number_input(
                "Medium Questions",
                min_value=0,
                max_value=total_questions,
                value=min(6, counts.get("medium", 6)),
                key="part1_medium_abs",
            )
        with col3:
            easy_count = st.number_input(
                "Easy Questions",
                min_value=0,
                max_value=total_questions,
                value=min(5, counts.get("easy", 5)),
                key="part1_easy_abs",
            )
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            hard_pct = st.slider("Hard %", 0, 100, 27, key="part1_hard_pct")
        with col2:
            medium_pct = st.slider("Medium %", 0, 100, 40, key="part1_medium_pct")
        with col3:
            easy_pct = st.slider("Easy %", 0, 100, 33, key="part1_easy_pct")

        total_pct = hard_pct + medium_pct + easy_pct
        if total_pct != 100:
            st.warning(f"⚠️ Percentages sum to {total_pct}%, should be 100%")

        hard_count = round(total_questions * hard_pct / 100)
        medium_count = round(total_questions * medium_pct / 100)
        easy_count = total_questions - hard_count - medium_count
        st.info(
            f"📊 Calculated: {hard_count} Hard + {medium_count} Medium + {easy_count} Easy = "
            f"{hard_count + medium_count + easy_count} questions"
        )

    total_selected = hard_count + medium_count + easy_count
    if total_selected != total_questions:
        st.error(f"❌ Selected {total_selected} questions, but quiz requires {total_questions}")

    validation_errors = []
    if question_bank:
        if hard_count > counts.get("hard", 0):
            validation_errors.append(f"Need {hard_count} hard questions, only {counts.get('hard', 0)} available")
        if medium_count > counts.get("medium", 0):
            validation_errors.append(
                f"Need {medium_count} medium questions, only {counts.get('medium', 0)} available"
            )
        if easy_count > counts.get("easy", 0):
            validation_errors.append(f"Need {easy_count} easy questions, only {counts.get('easy', 0)} available")

    for error in validation_errors:
        st.error(f"❌ {error}")

    st.divider()
    st.header("4️⃣ Randomization")
    use_fixed_seed = st.checkbox(
        "Use fixed seed (reproducible output)",
        value=False,
        help="Enable this if you want the exact same allocation for identical inputs.",
        key="part1_use_fixed_seed",
    )
    fixed_seed = None
    if use_fixed_seed:
        fixed_seed = st.number_input(
            "Seed value",
            min_value=0,
            max_value=2_147_483_647,
            value=42,
            step=1,
            key="part1_seed_value",
        )
        st.caption("Same input + same seed -> same allocation.")
    else:
        st.caption("Each generation uses a fresh random seed.")

    st.divider()
    st.header("5️⃣ Generate Question Papers")

    can_generate = question_bank is not None and total_selected == total_questions and len(validation_errors) == 0

    if st.button("🚀 Generate Question Papers", disabled=not can_generate, type="primary", key="part1_generate"):
        with st.spinner("Generating question papers..."):
            try:
                quiz_structure = QuizStructure(
                    hard_count=hard_count,
                    medium_count=medium_count,
                    easy_count=easy_count,
                )
                q_ids_by_diff = {
                    "hard": question_bank.get_question_ids_by_difficulty("hard"),
                    "medium": question_bank.get_question_ids_by_difficulty("medium"),
                    "easy": question_bank.get_question_ids_by_difficulty("easy"),
                }

                run_seed = int(fixed_seed) if use_fixed_seed else secrets.randbelow(2_147_483_647)
                allocation_matrix, usage_counts = allocate_quizzes(
                    q_ids_by_diff,
                    num_students=num_students,
                    quiz_structure=quiz_structure,
                    seed=run_seed,
                )
                shuffled_matrix = shuffle_all_quizzes(allocation_matrix, base_seed=run_seed)

                excel_bytes = create_formatted_excel(
                    allocation_matrix=allocation_matrix,
                    shuffled_matrix=shuffled_matrix,
                    usage_counts=usage_counts,
                    question_bank=question_bank,
                )

                st.session_state["part1_question_papers_bytes"] = excel_bytes
                st.success(f"✅ Generated {num_students} question papers!")
                st.download_button(
                    label="📥 Download Question Papers (Excel)",
                    data=excel_bytes,
                    file_name="question_papers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="part1_download_papers",
                )

                col1, col2, col3 = st.columns(3)
                col1.metric("Students", num_students)
                col2.metric("Questions/Quiz", total_questions)
                col3.metric("Total Sheets", num_students + 5)  # sets + answer key + alloc + shuffled + eval + qbank

                if use_fixed_seed:
                    st.caption(f"Seed used: {run_seed} (fixed)")
                else:
                    st.caption(f"Seed used: {run_seed} (auto-generated for this run)")

                st.caption(
                    "Sheets: S-01 … S-NN, Answer_Key, Allocation_Table, Shuffled_Table, Evaluation, Question_Bank"
                )
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")


def _render_answer_checking_tab():
    """Part 2 UI: generate responses and score submissions."""
    st.markdown("Generate dummy responses and validate/score answer sheets.")
    st.divider()

    st.header("Generate Dummy Responses")
    gen_qp_upload = st.file_uploader("Upload Question Papers (.xlsx)", type=["xlsx"], key="part2_gen_qp_upload")

    col1, col2, col3 = st.columns(3)
    with col1:
        gen_students = st.number_input("Students", min_value=1, max_value=500, value=70, key="part2_gen_students")
    with col2:
        correct_rate = st.slider("Correct %", 0, 100, 70, key="part2_correct_rate")
    with col3:
        wrong_rate = st.slider("Wrong %", 0, 100, 20, key="part2_wrong_rate")

    extra_rate = 100 - correct_rate - wrong_rate
    if extra_rate < 0:
        st.error("❌ Correct% + Wrong% cannot exceed 100.")
    else:
        st.caption(f"Remaining {extra_rate}% is treated as wrong (all assigned questions are compulsory).")

    use_fixed_gen_seed = st.checkbox(
        "Use fixed seed for dummy responses",
        value=False,
        key="part2_use_fixed_gen_seed",
    )
    gen_seed = None
    if use_fixed_gen_seed:
        gen_seed = st.number_input(
            "Generator seed",
            min_value=0,
            max_value=2_147_483_647,
            value=42,
            step=1,
            key="part2_gen_seed",
        )

    if st.button("🧪 Generate Dummy Responses", type="primary", key="part2_generate_responses"):
        if extra_rate < 0:
            st.error("Fix rates before generating responses.")
        else:
            temp_files = []
            try:
                if not gen_qp_upload:
                    st.error("Upload Question Papers.")
                    return
                qp_bytes = gen_qp_upload.getvalue()

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="part2_qp_") as qp_tmp:
                    qp_tmp.write(qp_bytes)
                    qp_path = qp_tmp.name
                temp_files.append(qp_path)

                question_bank = _load_question_bank_from_question_papers(qp_path)
                response_df = generate_responses(
                    question_papers_path=qp_path,
                    question_bank=question_bank,
                    num_students=int(gen_students),
                    correct_rate=float(correct_rate) / 100.0,
                    wrong_rate=float(wrong_rate) / 100.0,
                    blank_rate=0.0,
                    seed=int(gen_seed) if use_fixed_gen_seed else None,
                )

                response_bytes = _make_excel_bytes_from_dataframe(response_df, "Responses")
                st.session_state["part2_generated_responses_bytes"] = response_bytes

                st.success(f"✅ Generated dummy responses for {len(response_df)} students.")
                st.download_button(
                    "📥 Download student_responses.xlsx",
                    data=response_bytes,
                    file_name="student_responses.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="part2_download_responses",
                )
                st.caption(f"Shape: {response_df.shape[0]} rows × {response_df.shape[1]} columns")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
            finally:
                for path in temp_files:
                    if os.path.exists(path):
                        os.remove(path)

    st.divider()
    st.header("Check & Score Responses")

    chk_qp_upload = st.file_uploader("Upload Question Papers (.xlsx)", type=["xlsx"], key="part2_chk_qp_upload")
    chk_resp_upload = st.file_uploader("Upload Student Responses (.xlsx)", type=["xlsx"], key="part2_chk_resp_upload")

    pass_threshold = st.number_input(
        "Pass Marks (Correct Answers)",
        min_value=0.0,
        max_value=200.0,
        value=6.0,
        step=1.0,
        key="part2_pass_threshold",
    )

    if st.button("✅ Check & Score", type="primary", key="part2_check_score"):
        if not chk_qp_upload:
            st.error("Upload Question Papers for checking.")
        elif not chk_resp_upload:
            st.error("Upload Student Responses.")
        else:
            temp_files = []
            report_temp_path = None
            try:
                qp_bytes = chk_qp_upload.getvalue()
                resp_bytes = chk_resp_upload.getvalue()

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="part2_chk_qp_") as qp_tmp:
                    qp_tmp.write(qp_bytes)
                    qp_path = qp_tmp.name
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="part2_chk_resp_") as resp_tmp:
                    resp_tmp.write(resp_bytes)
                    resp_path = resp_tmp.name
                temp_files.extend([qp_path, resp_path])

                question_bank = _load_question_bank_from_question_papers(qp_path)
                response_df = load_response_sheet(resp_path)
                report = check_all_responses(
                    response_df=response_df,
                    question_papers_path=qp_path,
                    question_bank=question_bank,
                    pass_threshold=float(pass_threshold),
                )

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="part2_report_") as report_tmp:
                    report_temp_path = report_tmp.name
                generate_scoring_report(
                    report,
                    report_temp_path,
                    question_papers_path=qp_path,
                    question_bank=question_bank,
                )
                with open(report_temp_path, "rb") as f:
                    report_bytes = f.read()

                st.success("✅ Scoring completed.")
                max_marks = report.student_reports[0].assigned if report.student_reports else 0
                col1, col2, col3 = st.columns(3)
                col1.metric("Average Correct", f"{report.avg_score:.2f}/{max_marks}")
                col2.metric("Median Correct", f"{report.median_score:.2f}/{max_marks}")
                col3.metric("Pass Rate (%)", f"{report.pass_rate:.2f}")

                if report.validation_issues:
                    st.warning(f"⚠️ Validation issues found for {len(report.validation_issues)} students.")
                else:
                    st.success("No validation issues found.")

                st.download_button(
                    "📥 Download scoring_report.xlsx",
                    data=report_bytes,
                    file_name="scoring_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="part2_download_report",
                )
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
            finally:
                for path in temp_files:
                    if os.path.exists(path):
                        os.remove(path)
                if report_temp_path and os.path.exists(report_temp_path):
                    os.remove(report_temp_path)


def main():
    st.title("📝 Quiz Generator")
    part1_tab, part2_tab = st.tabs(["Part 1: Generate Papers", "Part 2: Answer Checking"])

    with part1_tab:
        _render_generation_tab()

    with part2_tab:
        _render_answer_checking_tab()


if __name__ == "__main__":
    main()
