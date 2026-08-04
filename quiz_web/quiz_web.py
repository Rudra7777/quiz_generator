"""Quiz Studio — a Reflex UI for generating and evaluating quiz papers."""

import reflex as rx

from quiz_web.pages.generate import generate_page
from quiz_web.pages.evaluate import evaluate_page

# A real browser reload should start a clean session (no carried-over uploads),
# but a websocket reconnect (idle timeout, tab refocus) should NOT wipe
# in-progress work — Reflex fires `on_load` for both, so it can't tell them
# apart. The browser session token, however, only gets cleared here on an
# actual page load (this script re-runs only then); a reconnect reuses the
# token already sitting in memory, so it keeps the same backend state. A
# reload clears the token, so a fresh token — and fresh state — is issued.
_clear_session_token = rx.script("""
try { window.sessionStorage.removeItem('token'); } catch (e) {}
""")

app = rx.App(head_components=[_clear_session_token])

app.add_page(generate_page, route="/", title="Quiz Studio · Generate")
app.add_page(evaluate_page, route="/evaluate", title="Quiz Studio · Evaluate")
