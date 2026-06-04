from __future__ import annotations

import base64
from pathlib import Path


def load_logo_b64() -> str | None:
    logo_path = Path(__file__).resolve().parents[2] / "assets" / "logo.png"
    if logo_path.exists():
        return base64.b64encode(logo_path.read_bytes()).decode()
    return None


LOGO_B64 = load_logo_b64()
