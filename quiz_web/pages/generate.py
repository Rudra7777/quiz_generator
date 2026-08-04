"""Generate Papers page (Part 1)."""

import reflex as rx

from quiz_web.state import GenerateState
from quiz_web.components.sidebar import shell
from quiz_web.components.ui import stat_tile, section, field, upload_zone, feedback


def _upload_step() -> rx.Component:
    return section(
        "1",
        "Upload question bank",
        "Excel with columns: question_no, question, option_a–d, answer, difficulty.",
        rx.vstack(
            upload_zone(
                "bank_upload",
                GenerateState.handle_bank_upload(rx.upload_files(upload_id="bank_upload")),
                "Question bank",
                uploaded=GenerateState.bank_uploaded,
                filename=GenerateState.bank_filename,
            ),
            feedback(GenerateState.error, "red", "triangle_alert"),
            rx.cond(
                GenerateState.bank_uploaded,
                rx.hstack(
                    stat_tile("Hard", GenerateState.hard_avail, "ruby"),
                    stat_tile("Medium", GenerateState.medium_avail, "amber"),
                    stat_tile("Easy", GenerateState.easy_avail, "grass"),
                    stat_tile("Total", GenerateState.total_avail, "indigo"),
                    spacing="3",
                    width="100%",
                ),
            ),
            spacing="3",
            width="100%",
        ),
    )


def _config_step() -> rx.Component:
    return section(
        "2",
        "Configuration",
        "How many papers, and how many questions each.",
        rx.hstack(
            field(
                "Number of students",
                rx.input(
                    type="number",
                    value=GenerateState.num_students,
                    on_change=GenerateState.set_num_students,
                    min=1,
                    max=500,
                ),
            ),
            field(
                "Questions per quiz",
                rx.input(
                    type="number",
                    value=GenerateState.total_questions,
                    on_change=GenerateState.set_total_questions,
                    min=1,
                    max=100,
                ),
            ),
            spacing="4",
            width="100%",
        ),
    )


def _difficulty_step() -> rx.Component:
    absolute_inputs = rx.hstack(
        field(
            "Hard",
            rx.input(type="number", value=GenerateState.hard_count, on_change=GenerateState.set_hard_count, min=0),
        ),
        field(
            "Medium",
            rx.input(type="number", value=GenerateState.medium_count, on_change=GenerateState.set_medium_count, min=0),
        ),
        field(
            "Easy",
            rx.input(type="number", value=GenerateState.easy_count, on_change=GenerateState.set_easy_count, min=0),
        ),
        spacing="4",
        width="100%",
    )

    def pct_slider(label, value, handler, tint):
        return field(
            f"{label} — {value}%",
            rx.slider(value=[value], min=0, max=100, on_change=handler, color_scheme=tint),
        )

    percentage_inputs = rx.vstack(
        pct_slider("Hard", GenerateState.hard_pct, GenerateState.set_hard_pct, "ruby"),
        pct_slider("Medium", GenerateState.medium_pct, GenerateState.set_medium_pct, "amber"),
        pct_slider("Easy", GenerateState.easy_pct, GenerateState.set_easy_pct, "grass"),
        rx.cond(
            GenerateState.pct_sum != 100,
            rx.text(
                "Percentages sum to ", GenerateState.pct_sum, "% (should be 100%).",
                size="1", color_scheme="amber",
            ),
        ),
        rx.text(
            "Calculated: ", GenerateState.calc_hard, " Hard + ", GenerateState.calc_medium,
            " Medium + ", GenerateState.calc_easy, " Easy",
            size="2", color_scheme="gray",
        ),
        spacing="3",
        width="100%",
    )

    return section(
        "3",
        "Difficulty distribution",
        "Set exact counts, or split by percentage of the quiz size.",
        rx.vstack(
            rx.radio(
                ["Absolute", "Percentage"],
                value=GenerateState.mode,
                on_change=GenerateState.set_mode,
                direction="row",
                spacing="5",
            ),
            rx.cond(GenerateState.mode == "Absolute", absolute_inputs, percentage_inputs),
            rx.cond(
                ~GenerateState.size_ok,
                rx.text(
                    "Selected ", GenerateState.total_selected, " of ",
                    GenerateState.total_questions, " required questions.",
                    size="1", color_scheme="red",
                ),
            ),
            rx.cond(
                GenerateState.bank_uploaded & ~GenerateState.avail_ok,
                rx.text(
                    "A difficulty exceeds what the bank provides.",
                    size="1", color_scheme="red",
                ),
            ),
            spacing="3",
            width="100%",
        ),
    )


def _seed_step() -> rx.Component:
    return section(
        "4",
        "Randomization",
        "Use a fixed seed to reproduce an identical allocation.",
        rx.vstack(
            rx.hstack(
                rx.switch(checked=GenerateState.use_fixed_seed, on_change=GenerateState.toggle_fixed_seed),
                rx.text("Use fixed seed", size="2"),
                spacing="3",
                align="center",
            ),
            rx.cond(
                GenerateState.use_fixed_seed,
                field(
                    "Seed",
                    rx.input(type="number", value=GenerateState.seed, on_change=GenerateState.set_seed, min=0),
                ),
                rx.text("Each run uses a fresh random seed.", size="1", color_scheme="gray"),
            ),
            spacing="3",
            width="100%",
        ),
    )


def _generate_step() -> rx.Component:
    return section(
        "5",
        "Generate question papers",
        "Produces one sheet per student plus answer key, tables, and an embedded bank.",
        rx.vstack(
            rx.button(
                rx.icon("rocket", size=18),
                "Generate question papers",
                on_click=GenerateState.generate_papers,
                disabled=~GenerateState.can_generate,
                size="3",
            ),
            feedback(GenerateState.status, "grass", "circle_check"),
            rx.cond(
                GenerateState.papers_ready,
                rx.vstack(
                    rx.hstack(
                        stat_tile("Students", GenerateState.generated_count, "indigo"),
                        stat_tile("Questions / quiz", GenerateState.total_questions, "indigo"),
                        stat_tile("Seed used", GenerateState.last_seed, "gray"),
                        spacing="3",
                        width="100%",
                    ),
                    rx.button(
                        rx.icon("download", size=18),
                        "Download question_papers.xlsx",
                        on_click=GenerateState.download_papers,
                        variant="soft",
                        size="3",
                    ),
                    spacing="3",
                    width="100%",
                ),
            ),
            spacing="3",
            width="100%",
        ),
    )


def generate_page() -> rx.Component:
    return shell(
        "generate",
        rx.vstack(
            rx.heading("Generate Papers", size="7"),
            rx.text(
                "Upload a bank and produce fair, randomized question papers for every student.",
                size="3", color_scheme="gray",
            ),
            spacing="1",
            align="start",
        ),
        _upload_step(),
        _config_step(),
        _difficulty_step(),
        _seed_step(),
        _generate_step(),
    )
