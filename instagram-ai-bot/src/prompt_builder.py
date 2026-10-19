"""シーン選択と画像プロンプトの組み立て。"""
from __future__ import annotations

import random
from typing import Any

from . import state

# 直近この件数で使ったシーンは再抽選する
DEDUPE_WINDOW = 8


def pick_scene(scenes: list[dict[str, Any]]) -> dict[str, Any]:
    recent = state.recent_scene_ids(DEDUPE_WINDOW)
    pool = [s for s in scenes if s["id"] not in recent] or scenes
    weights = [float(s.get("weight", 1)) for s in pool]
    return random.choices(pool, weights=weights, k=1)[0]


def build_image_prompt(character: dict[str, Any], scene: dict[str, Any]) -> str:
    """Flux 系は否定プロンプトを持たないので、品質指示も肯定形で書き切る。"""
    parts: list[str] = []

    trigger = (character.get("trigger_phrase") or "").strip()
    if trigger and character.get("lora_url"):
        parts.append(trigger)

    parts.append(character["appearance"].strip())
    parts.append(scene["setting"].strip())
    parts.append(f"wearing {scene['outfit'].strip()}")
    parts.append(scene["pose"].strip())
    parts.append(character["style"].strip())
    parts.append("full color photograph, sharp focus on the face, natural proportions")

    return ", ".join(p.rstrip(",. ") for p in parts if p)
