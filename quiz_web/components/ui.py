"""Reusable presentational helpers shared across pages."""

import reflex as rx


def stat_tile(label: str, value, color_scheme: str = "gray") -> rx.Component:
    """A small scoreboard-style cell: quiet label, prominent mono figure."""
    return rx.card(
        rx.vstack(
            rx.text(
                label,
                size="1",
                weight="medium",
                color_scheme="gray",
                style={"textTransform": "uppercase", "letterSpacing": "0.08em"},
            ),
            rx.text(
                value,
                size="7",
                weight="bold",
                color_scheme=color_scheme,
                style={"fontFamily": "monospace"},
            ),
            spacing="1",
            align="start",
        ),
        size="1",
        flex="1",
    )


def section(number: str, title: str, description: str, *children, tint: str = "indigo") -> rx.Component:
    """A numbered step card. The generate/evaluate flows are real sequences,
    so the numbering encodes order rather than decoration."""
    return rx.card(
        rx.vstack(
            rx.hstack(
                rx.badge(number, variant="soft", color_scheme=tint, size="2", radius="full"),
                rx.vstack(
                    rx.heading(title, size="4"),
                    rx.text(description, size="2", color_scheme="gray"),
                    spacing="0",
                    align="start",
                ),
                spacing="3",
                align="center",
                width="100%",
            ),
            rx.box(*children, width="100%"),
            spacing="4",
            width="100%",
        ),
        size="3",
        width="100%",
    )


def field(label: str, control: rx.Component, hint: str = "") -> rx.Component:
    """A labelled form control."""
    return rx.vstack(
        rx.text(label, size="2", weight="medium"),
        control,
        rx.cond(hint != "", rx.text(hint, size="1", color_scheme="gray")),
        spacing="1",
        align="start",
        width="100%",
    )


def upload_zone(upload_id: str, on_drop, prompt: str, tint: str = "indigo") -> rx.Component:
    """A dashed drag-and-drop area for a single .xlsx file."""
    return rx.upload(
        rx.vstack(
            rx.icon("file-up", size=22),
            rx.text(prompt, size="2", weight="medium"),
            rx.cond(
                rx.selected_files(upload_id),
                rx.text(rx.selected_files(upload_id), size="1", color_scheme=tint),
                rx.text("Drop an .xlsx file or click to browse", size="1", color_scheme="gray"),
            ),
            spacing="2",
            align="center",
        ),
        id=upload_id,
        accept={
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
        },
        max_files=1,
        multiple=False,
        on_drop=on_drop,
        border=f"1.5px dashed var(--{tint}-7)",
        border_radius="var(--radius-4)",
        background_color=f"var(--{tint}-2)",
        padding="1.75em",
        width="100%",
        cursor="pointer",
    )


def feedback(message, color_scheme: str, icon: str) -> rx.Component:
    """Render a callout only when the message string is non-empty."""
    return rx.cond(
        message != "",
        rx.callout(
            message,
            icon=icon,
            color_scheme=color_scheme,
            size="1",
            width="100%",
        ),
    )
