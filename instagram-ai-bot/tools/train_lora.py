"""dataset/training.zip から LoRA を学習する（fal-ai/flux-lora-portrait-trainer）。

  python -m tools.train_lora

学習には数分〜十数分かかります。終わると LoRA の URL が出るので、
config/character.yaml の lora_url に貼ってください。以降の投稿はそれを使います。
"""
from __future__ import annotations

import argparse
from pathlib import Path

import fal_client

from src.config import ROOT, Secrets, load_character

MODEL = "fal-ai/flux-lora-portrait-trainer"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--zip", default="dataset/training.zip", help="学習素材の zip")
    parser.add_argument("--steps", type=int, default=2500)
    parser.add_argument(
        "--trigger", default="", help="トリガーワード（例: RINCHAR）。空なら未使用"
    )
    args = parser.parse_args()

    Secrets.load(need_instagram=False)
    character = load_character()
    trigger = args.trigger or character.get("trigger_phrase") or ""

    zip_path = Path(args.zip)
    if not zip_path.is_absolute():
        zip_path = ROOT / zip_path
    if not zip_path.exists():
        raise SystemExit(
            f"{zip_path} がありません。先に `python -m tools.bootstrap_identity` を実行してください。"
        )

    print(f"素材をアップロード: {zip_path.name}")
    images_data_url = fal_client.upload_file(str(zip_path))

    arguments = {
        "images_data_url": images_data_url,
        "steps": args.steps,
        "subject_crop": True,
        "multiresolution_training": True,
    }
    if trigger:
        arguments["trigger_phrase"] = trigger

    print(f"学習開始 ({MODEL}, steps={args.steps})。数分かかります…")
    result = fal_client.subscribe(
        MODEL,
        arguments=arguments,
        with_logs=True,
        on_queue_update=lambda u: [
            print(f"  [fal] {log.get('message','')}")
            for log in (getattr(u, "logs", None) or [])
        ],
    )

    lora_url = result["diffusers_lora_file"]["url"]
    print()
    print("=" * 70)
    print("config/character.yaml に貼る値:")
    print(f'  lora_url: "{lora_url}"')
    if trigger:
        print(f'  trigger_phrase: "{trigger}"')
    print()
    print("貼ったら `python -m src.main --dry-run` で仕上がりを確認してください。")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
