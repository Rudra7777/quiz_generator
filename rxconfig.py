import reflex as rx

config = rx.Config(
    app_name="quiz_web",
    plugins=[
        rx.plugins.RadixThemesPlugin(
            theme=rx.theme(
                appearance="inherit",
                accent_color="indigo",
                gray_color="slate",
                radius="large",
                scaling="100%",
            ),
        ),
    ],
)
