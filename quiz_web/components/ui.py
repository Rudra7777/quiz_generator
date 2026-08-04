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


def upload_zone(
    upload_id: str,
    on_drop,
    prompt: str,
    uploaded=False,
    filename="",
    tint: str = "indigo",
) -> rx.Component:
    """A dashed drag-and-drop area for a single .xlsx file with rich animations.

    `uploaded`/`filename` should come from backend state (set once the file has
    actually been read and processed), not `rx.selected_files` — that only reflects
    the browser's pending selection and clears back to empty once the upload completes.
    """

    # ── Idle state (no file confirmed yet) ──
    idle_content = rx.vstack(
        rx.icon("file-up", size=28, color=f"var(--{tint}-9)"),
        rx.text(prompt, size="2", weight="bold"),
        rx.text(
            "Drop an .xlsx file or click to browse",
            size="1",
            color_scheme="gray",
        ),
        spacing="2",
        align="center",
    )

    # ── Success state (backend has confirmed the file was received) ──
    success_content = rx.vstack(
        rx.box(
            rx.icon("circle-check", size=30, color="var(--grass-9)"),
            class_name="upload-icon-pop",
        ),
        rx.text(prompt, size="2", weight="bold", color_scheme="grass"),
        rx.text(
            filename,
            size="1",
            weight="medium",
            color_scheme=tint,
            style={"wordBreak": "break-all"},
        ),
        spacing="2",
        align="center",
    )

    return rx.fragment(
        # Inject keyframe animations once (Reflex deduplicates rx.script)
        rx.script("""
            (function() {
                if (document.getElementById('upload-zone-animations')) return;
                const style = document.createElement('style');
                style.id = 'upload-zone-animations';
                style.textContent = `
                    /* ── Base transition for every upload zone ── */
                    .upload-zone {
                        transition: transform 0.25s ease, border-color 0.25s ease,
                                    background-color 0.25s ease, box-shadow 0.25s ease;
                    }
                    .upload-zone:hover {
                        transform: scale(1.015);
                        box-shadow: 0 0 0 3px var(--accent-a4);
                        border-color: var(--accent-9) !important;
                    }

                    /* ── Drag-over highlight ── */
                    .upload-zone.drag-over {
                        transform: scale(1.04);
                        border-color: var(--accent-9) !important;
                        background-color: var(--accent-a4) !important;
                        box-shadow: 0 0 16px 2px var(--accent-a5);
                    }

                    /* ── Pop animation for checkmark icon on success ── */
                    @keyframes icon-pop {
                        0%   { transform: scale(0.3); opacity: 0; }
                        60%  { transform: scale(1.25); opacity: 1; }
                        100% { transform: scale(1); opacity: 1; }
                    }
                    .upload-icon-pop {
                        animation: icon-pop 0.4s ease-out;
                    }

                    /* ── Pulse glow on file drop ── */
                    @keyframes drop-pulse {
                        0%   { box-shadow: 0 0 0 0 var(--accent-a6); }
                        70%  { box-shadow: 0 0 0 10px transparent; }
                        100% { box-shadow: 0 0 0 0 transparent; }
                    }
                    .upload-zone.drop-flash {
                        animation: drop-pulse 0.6s ease-out;
                    }
                `;
                document.head.appendChild(style);

                // Attach dragover / dragleave listeners to every upload zone
                function attachDragListeners() {
                    document.querySelectorAll('.upload-zone').forEach(el => {
                        if (el.dataset.dragBound) return;
                        el.dataset.dragBound = '1';

                        el.addEventListener('dragover', () => el.classList.add('drag-over'));
                        el.addEventListener('dragleave', () => el.classList.remove('drag-over'));
                        el.addEventListener('drop', () => {
                            el.classList.remove('drag-over');
                            el.classList.add('drop-flash');
                            setTimeout(() => el.classList.remove('drop-flash'), 600);
                        });
                    });
                }
                // Run once now and again after Reflex re-renders
                attachDragListeners();
                new MutationObserver(attachDragListeners)
                    .observe(document.body, {childList: true, subtree: true});
            })();
        """),
        rx.upload(
            rx.cond(
                uploaded,
                success_content,
                idle_content,
            ),
            id=upload_id,
            accept={
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
            },
            max_files=1,
            multiple=False,
            on_drop=on_drop,
            border=f"2px dashed var(--{tint}-7)",
            border_radius="var(--radius-4)",
            background_color=f"var(--{tint}-2)",
            padding="2em 1.75em",
            width="100%",
            cursor="pointer",
            class_name="upload-zone",
        ),
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
