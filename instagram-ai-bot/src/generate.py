"""fal.ai での画像生成。

顔の固定方法は character.yaml の埋まり方で自動的に切り替わる:
  lora_url あり            -> fal-ai/flux-lora        （学習済み LoRA / 最も一貫性が高い）
  identity_reference_url   -> fal-ai/flux-pulid       （参照顔1枚 / 学習不要）
  どちらも無し              -> fal-ai/flux-lora 素の状態（毎回別人。bootstrap 用）
"""
from __future__ import annotations

import random
from dataclasses import dataclass
from typing import Any

import fal_client


class NSFWRejected(RuntimeError):
    """安全性チェックで弾かれた。別シードで引き直す。"""


@dataclass
class Generated:
    url: str
    seed: int
    model: str
    prompt: str


def _on_update(update: Any) -> None:
    if isinstance(update, fal_client.InProgress):
        for log in update.logs or []:
            print(f"    [fal] {log.get('message', '')}")


def _run(model: str, arguments: dict[str, Any]) -> dict[str, Any]:
    print(f"  → {model}")
    return fal_client.subscribe(
        model, arguments=arguments, with_logs=True, on_queue_update=_on_update
    )


def generate(
    character: dict[str, Any],
    prompt: str,
    *,
    seed: int | None = None,
    force_base: bool = False,
) -> Generated:
    img = character["image"]
    models = character["models"]
    seed = seed if seed is not None else random.randint(1, 2**31 - 1)

    common: dict[str, Any] = {
        "prompt": prompt,
        "image_size": {"width": img["width"], "height": img["height"]},
        "num_inference_steps": img["num_inference_steps"],
        "guidance_scale": img["guidance_scale"],
        "seed": seed,
        "num_images": 1,
        "enable_safety_checker": True,
    }

    lora_url = (character.get("lora_url") or "").strip()
    ref_url = (character.get("identity_reference_url") or "").strip()

    if force_base or (not lora_url and not ref_url):
        model = models["base"]
        args = {**common, "output_format": img["output_format"]}
    elif lora_url:
        model = models["lora"]
        args = {
            **common,
            "output_format": img["output_format"],
            "loras": [{"path": lora_url, "scale": float(character.get("lora_scale", 1.0))}],
        }
    else:
        # PuLID は output_format を持たないので指定しない（media.py で JPEG 化する）
        model = models["identity"]
        args = {**common, "reference_image_url": ref_url}

    result = _run(model, args)

    images = result.get("images") or []
    if not images:
        raise RuntimeError(f"{model} が画像を返しませんでした: {result}")
    if any(result.get("has_nsfw_concepts") or []):
        raise NSFWRejected(f"セーフティチェッカーに弾かれました (seed={seed})")

    return Generated(
        url=images[0]["url"],
        seed=int(result.get("seed", seed)),
        model=model,
        prompt=prompt,
    )


def generate_with_retry(
    character: dict[str, Any], prompt: str, *, attempts: int = 3
) -> Generated:
    last: Exception | None = None
    for i in range(1, attempts + 1):
        try:
            return generate(character, prompt)
        except NSFWRejected as exc:
            last = exc
            print(f"  ! {exc} — リトライ {i}/{attempts}")
    raise RuntimeError(f"{attempts} 回とも生成に失敗しました") from last
