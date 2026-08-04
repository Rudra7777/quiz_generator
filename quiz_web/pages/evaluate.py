"""Evaluate Answers page (Part 2)."""

import reflex as rx

from quiz_web.state import EvaluateState
from quiz_web.components.sidebar import shell
from quiz_web.components.ui import stat_tile, section, field, upload_zone, feedback

TINT = "ruby"  # grading red — distinguishes evaluation from generation


def _generate_responses_step() -> rx.Component:
    return section(
        "1",
        "Generate dummy responses",
        "Simulate a Google-Form submission set from your question papers.",
        rx.vstack(
            upload_zone(
                "gen_qp_upload",
                EvaluateState.handle_gen_papers_upload(rx.upload_files(upload_id="gen_qp_upload")),
                "Question papers",
                uploaded=EvaluateState.gen_papers_uploaded,
                filename=EvaluateState.gen_papers_filename,
                tint=TINT,
            ),
            rx.hstack(
                field(
                    "Students",
                    rx.input(
                        type="number",
                        value=EvaluateState.gen_students,
                        on_change=EvaluateState.set_gen_students,
                        min=1,
                        max=500,
                    ),
                ),
                rx.vstack(
                    field(
                        f"Correct %",
                        rx.slider(
                            value=[EvaluateState.correct_rate],
                            min=0, max=100,
                            on_change=EvaluateState.set_correct_rate,
                            color_scheme="grass",
                        ),
                    ),
                    rx.text("Correct: ", EvaluateState.correct_rate, "%", size="1", color_scheme="gray"),
                    spacing="1",
                    width="100%",
                ),
                rx.vstack(
                    field(
                        f"Wrong %",
                        rx.slider(
                            value=[EvaluateState.wrong_rate],
                            min=0, max=100,
                            on_change=EvaluateState.set_wrong_rate,
                            color_scheme="ruby",
                        ),
                    ),
                    rx.text("Wrong: ", EvaluateState.wrong_rate, "%", size="1", color_scheme="gray"),
                    spacing="1",
                    width="100%",
                ),
                spacing="4",
                width="100%",
                align="start",
            ),
            rx.text(
                "Remaining ", EvaluateState.extra_rate,
                "% is also treated as wrong (assigned questions are compulsory).",
                size="1", color_scheme="gray",
            ),
            rx.hstack(
                rx.switch(checked=EvaluateState.gen_use_seed, on_change=EvaluateState.toggle_gen_seed),
                rx.text("Use fixed seed", size="2"),
                rx.cond(
                    EvaluateState.gen_use_seed,
                    rx.input(
                        type="number",
                        value=EvaluateState.gen_seed,
                        on_change=EvaluateState.set_gen_seed,
                        min=0,
                        width="8em",
                    ),
                ),
                spacing="3",
                align="center",
            ),
            rx.button(
                rx.icon("flask-conical", size=18),
                "Generate dummy responses",
                on_click=EvaluateState.generate_responses_action,
                size="3",
                color_scheme=TINT,
            ),
            feedback(EvaluateState.gen_error, "red", "triangle_alert"),
            feedback(EvaluateState.gen_status, "grass", "circle_check"),
            rx.cond(
                EvaluateState.responses_ready,
                rx.button(
                    rx.icon("download", size=18),
                    "Download student_responses.xlsx",
                    on_click=EvaluateState.download_responses,
                    variant="soft",
                    color_scheme=TINT,
                    size="3",
                ),
            ),
            spacing="3",
            width="100%",
        ),
        tint=TINT,
    )


def _check_score_step() -> rx.Component:
    return section(
        "2",
        "Check & score responses",
        "Validate submissions against the answer key and produce a scoring report.",
        rx.vstack(
            rx.hstack(
                upload_zone(
                    "chk_qp_upload",
                    EvaluateState.handle_chk_papers_upload(rx.upload_files(upload_id="chk_qp_upload")),
                    "Question papers",
                    uploaded=EvaluateState.chk_papers_uploaded,
                    filename=EvaluateState.chk_papers_filename,
                    tint=TINT,
                ),
                upload_zone(
                    "chk_resp_upload",
                    EvaluateState.handle_chk_responses_upload(rx.upload_files(upload_id="chk_resp_upload")),
                    "Student responses",
                    uploaded=EvaluateState.chk_responses_uploaded,
                    filename=EvaluateState.chk_responses_filename,
                    tint=TINT,
                ),
                spacing="3",
                width="100%",
            ),
            field(
                "Pass marks",
                rx.input(
                    type="number",
                    value=EvaluateState.pass_threshold,
                    on_change=EvaluateState.set_pass_threshold,
                    min=0,
                    width="10em",
                ),
                hint="Correct answers needed to pass.",
            ),
            rx.button(
                rx.icon("clipboard-check", size=18),
                "Check & score",
                on_click=EvaluateState.check_score_action,
                size="3",
                color_scheme=TINT,
            ),
            feedback(EvaluateState.score_error, "red", "triangle_alert"),
            feedback(EvaluateState.score_status, "grass", "circle_check"),
            rx.cond(
                EvaluateState.report_ready,
                rx.vstack(
                    rx.hstack(
                        stat_tile("Avg correct", EvaluateState.avg_score, "indigo"),
                        stat_tile("Median correct", EvaluateState.median_score, "indigo"),
                        stat_tile("Pass rate %", EvaluateState.pass_rate, "grass"),
                        stat_tile("Max marks", EvaluateState.max_marks, "gray"),
                        spacing="3",
                        width="100%",
                    ),
                    rx.cond(
                        EvaluateState.validation_count > 0,
                        rx.callout(
                            rx.hstack(
                                rx.text(EvaluateState.validation_count),
                                rx.text("students answered questions outside their set."),
                                spacing="1",
                            ),
                            icon="triangle_alert",
                            color_scheme="amber",
                            size="1",
                        ),
                        rx.callout(
                            "No validation issues found.",
                            icon="circle_check",
                            color_scheme="grass",
                            size="1",
                        ),
                    ),
                    rx.button(
                        rx.icon("download", size=18),
                        "Download scoring_report.xlsx",
                        on_click=EvaluateState.download_report,
                        variant="soft",
                        color_scheme=TINT,
                        size="3",
                    ),
                    spacing="3",
                    width="100%",
                ),
            ),
            spacing="3",
            width="100%",
        ),
        tint=TINT,
    )


def evaluate_page() -> rx.Component:
    return shell(
        "evaluate",
        rx.vstack(
            rx.heading("Evaluate Answers", size="7"),
            rx.text(
                "Simulate responses, then validate and score them against the answer key.",
                size="3", color_scheme="gray",
            ),
            spacing="1",
            align="start",
        ),
        _generate_responses_step(),
        _check_score_step(),
    )
