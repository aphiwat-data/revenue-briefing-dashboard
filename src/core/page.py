from __future__ import annotations


def setup_page(st_module) -> None:
    st_module.set_page_config(
        page_title="Daily Revenue Briefing Dashboard",
        page_icon=None,
        layout="wide",
        initial_sidebar_state="expanded",
    )
