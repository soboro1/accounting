"""投稿履歴。直近のシーン重複を避けるためと、運用ログのため。"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .config import ROOT

HISTORY = ROOT / "state" / "history.jsonl"


def read_history(limit: int | None = None) -> list[dict[str, Any]]:
    if not HISTORY.exists():
        return []
    rows = [
        json.loads(line)
        for line in HISTORY.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    return rows[-limit:] if limit else rows


def recent_scene_ids(window: int) -> set[str]:
    return {row.get("scene_id", "") for row in read_history(limit=window)}


def append(record: dict[str, Any]) -> None:
    record = {"posted_at": datetime.now(timezone.utc).isoformat(), **record}
    HISTORY.parent.mkdir(parents=True, exist_ok=True)
    with HISTORY.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, ensure_ascii=False) + "\n")
