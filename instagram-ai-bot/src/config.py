"""設定とシークレットの読み込み。"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")


def _require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise SystemExit(
            f"環境変数 {name} が未設定です。.env か GitHub Secrets に設定してください。"
        )
    return value


@dataclass(frozen=True)
class Secrets:
    fal_key: str
    anthropic_api_key: str
    ig_user_id: str
    ig_access_token: str
    ig_api_host: str
    ig_api_version: str

    @classmethod
    def load(cls, *, need_instagram: bool = True) -> "Secrets":
        return cls(
            fal_key=_require("FAL_KEY"),
            anthropic_api_key=_require("ANTHROPIC_API_KEY"),
            ig_user_id=_require("IG_USER_ID") if need_instagram else "",
            ig_access_token=_require("IG_ACCESS_TOKEN") if need_instagram else "",
            ig_api_host=os.environ.get("IG_API_HOST", "graph.facebook.com"),
            ig_api_version=os.environ.get("IG_API_VERSION", "v25.0"),
        )


def load_yaml(relative: str) -> dict[str, Any]:
    with (ROOT / relative).open(encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def load_character() -> dict[str, Any]:
    return load_yaml("config/character.yaml")


def load_scenes() -> list[dict[str, Any]]:
    return load_yaml("config/scenes.yaml")["scenes"]
