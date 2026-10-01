"""Shuffle Options page — balance a bank's answer key over A-D."""

import reflex as rx

from quiz_web.state import ShuffleState
from quiz_web.components.sidebar import shell
from quiz_web.components.ui import section, upload_zone, feedback

TINT = "amber"


def _count_table(title: str, rows) -> rx.Component:
    header = ["Difficulty", "A", "B", "C", "D", "Total"]
    return rx.vstack(
        rx.text(title, size="2", weight="bold"),
        rx.table.root(
            rx.table.header(
                rx.table.row(*[rx.table.column_header_cell(h) for h in header]),
            ),
            rx.table.body(
                rx.foreach(
                    rows,
                    lambda row: rx.table.row(
                        rx.foreach(row, lambda cell: rx.table.cell(cell)),
                    ),
                ),
            ),
            variant="surface",
            size="1",
            width="100%",
        ),
        spacing="2",
        width="100%",
    )


def _upload_step() -> rx.Component:
    return section(
        "1",
        "Upload question bank",
        "The same Excel format as Generate Papers.",
        rx.vstack(
            upload_zone(
                "shuffle_bank_upload",
                ShuffleState.handle_bank_upload(rx.upload_files(upload_id="shuffle_bank_upload")),
                "Question bank",
                uploaded=ShuffleState.bank_uploaded,
                filename=ShuffleState.bank_filename,
                tint=TINT,
            ),
            feedback(ShuffleState.error, "red", "triangle_alert"),
            spacing="3",
            width="100%",
        ),
        tint=TINT,
    )


def _result_step() -> rx.Component:
    return section(
        "2",
        "Download the balanced bank",
        "Correct answers per letter, before and after. Use the balanced file on Generate Papers.",
        rx.cond(
            ShuffleState.bank_uploaded,
            rx.vstack(
                feedback(ShuffleState.status, "grass", "circle_check"),
                rx.hstack(
                    _count_table("Before", ShuffleState.before_rows),
                    _count_table("After", ShuffleState.after_rows),
                    spacing="4",
                    width="100%",
                    align="start",
                ),
                rx.button(
                    rx.icon("download", size=18),
                    "Download ", ShuffleState.balanced_filename,
                    on_click=ShuffleState.download_balanced,
                    color_scheme=TINT,
                    size="3",
                ),
                spacing="3",
                width="100%",
            ),
            rx.text("Upload a bank to see its balance.", size="2", color_scheme="gray"),
        ),
        tint=TINT,
    )


def shuffle_page() -> rx.Component:
    return shell(
        "shuffle",
        rx.vstack(
            rx.heading("Shuffle Options", size="7"),
            rx.text(
                "Reorder each question's options so the correct answer is A, B, C and D "
                "equally often within every difficulty level.",
                size="3", color_scheme="gray",
            ),
            spacing="1",
            align="start",
        ),
        _upload_step(),
        _result_step(),
    )
