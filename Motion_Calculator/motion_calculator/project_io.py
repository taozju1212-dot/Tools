from __future__ import annotations

import json
from pathlib import Path

from .models import AppState


def save_project(path: str | Path, state: AppState) -> None:
    Path(path).write_text(json.dumps(state.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")


def load_project(path: str | Path) -> AppState:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    return AppState.from_dict(data)
