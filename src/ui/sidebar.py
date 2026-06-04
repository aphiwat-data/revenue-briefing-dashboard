from __future__ import annotations

import html
import re


def hotel_key(hotel_name) -> str:
    safe = re.sub(r"[^A-Za-z0-9_]+", "_", str(hotel_name))
    return f"hotel_checkbox_{safe}"


def selected_property_chips_html(selected_hotels) -> str:
    chips = []
    for hotel in selected_hotels:
        name = str(hotel)
        chips.append(
            f'<span class="property-chip" title="{html.escape(name)}">'
            f"{html.escape(name)}</span>"
        )
    return "".join(chips)
