"""Quiz Studio — a Reflex UI for generating and evaluating quiz papers."""

import reflex as rx

from quiz_web.pages.generate import generate_page
from quiz_web.pages.evaluate import evaluate_page

app = rx.App()

app.add_page(generate_page, route="/", title="Quiz Studio · Generate")
app.add_page(evaluate_page, route="/evaluate", title="Quiz Studio · Evaluate")
