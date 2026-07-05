"""Persistent left navigation."""

import reflex as rx


def _nav_item(label: str, icon: str, route: str, active: bool) -> rx.Component:
    return rx.link(
        rx.hstack(
            rx.icon(icon, size=18),
            rx.text(label, size="3", weight="medium"),
            spacing="3",
            align="center",
            width="100%",
            padding_x="0.75em",
            padding_y="0.6em",
            border_radius="var(--radius-3)",
            background_color=rx.cond(active, "var(--accent-4)", "transparent"),
            color=rx.cond(active, "var(--accent-11)", "var(--gray-11)"),
            _hover={"background_color": "var(--accent-3)"},
        ),
        href=route,
        width="100%",
        underline="none",
    )


def sidebar(active: str) -> rx.Component:
    return rx.vstack(
        rx.hstack(
            rx.icon("square-pen", size=24, color="var(--accent-11)"),
            rx.vstack(
                rx.heading("Quiz Studio", size="4"),
                rx.text("Generate & evaluate", size="1", color_scheme="gray"),
                spacing="0",
                align="start",
            ),
            spacing="3",
            align="center",
            padding_bottom="0.5em",
        ),
        rx.divider(),
        rx.vstack(
            _nav_item("Generate Papers", "layout-grid", "/", active == "generate"),
            _nav_item("Evaluate Answers", "check-check", "/evaluate", active == "evaluate"),
            spacing="1",
            width="100%",
            padding_top="0.5em",
        ),
        rx.spacer(),
        rx.hstack(
            rx.color_mode.button(),
            rx.text("Theme", size="1", color_scheme="gray"),
            spacing="2",
            align="center",
        ),
        spacing="4",
        height="100vh",
        width="16em",
        padding="1.25em",
        border_right="1px solid var(--gray-4)",
        position="sticky",
        top="0",
    )


def shell(active: str, *content: rx.Component) -> rx.Component:
    """Sidebar + scrollable page body."""
    return rx.hstack(
        sidebar(active),
        rx.box(
            rx.vstack(*content, spacing="5", width="100%", max_width="60em"),
            padding_x="2.5em",
            padding_y="2em",
            width="100%",
            display="flex",
            justify_content="center",
        ),
        spacing="0",
        align="start",
        width="100%",
    )
