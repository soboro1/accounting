"""キャラクターの「顔」を作る最初の一歩。

  1. 素の Flux で正面ポートレートを1枚生成 → これがこの子の“正体”になる
  2. その顔を PuLID の参照にして、角度・表情・光違いのバリエーションを N 枚生成
  3. 全部を zip にまとめる（tools/train_lora.py の入力になる）

  python -m tools.bootstrap_identity --count 18

終わったら表示される identity_reference_url を config/character.yaml に貼れば、
LoRA を学習しなくても Phase 1 として運用を始められます。
"""
from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

from src import generate, media
from src.config import ROOT, Secrets, load_character

DATASET = ROOT / "dataset"

# 学習素材としてバラけさせたい軸。顔は同じまま、条件だけ変える。
VARIATIONS = [
    "close-up portrait, facing the camera directly, neutral expression, soft window light",
    "portrait, three-quarter view, slight smile, overcast daylight outdoors",
    "portrait, profile view from the side, calm expression, warm indoor lamp light",
    "upper body shot, looking away from the camera, laughing, bright afternoon sun",
    "portrait, head tilted slightly down, thoughtful expression, dim evening light",
    "upper body shot, hair tied up, looking over her shoulder, backlit at golden hour",
    "close-up portrait, hair down covering one side of her face, soft studio light",
    "medium shot, standing, arms relaxed, full daylight, plain background",
    "portrait, wearing a cap, looking up, bright outdoor light",
    "close-up portrait, eyes closed, gentle expression, soft diffused light",
    "upper body shot, wearing a coat with a scarf, cold winter daylight",
    "portrait, wet hair after rain, natural expression, grey overcast light",
    "medium shot, sitting at a table, leaning on her elbows, warm cafe light",
    "close-up portrait, wide genuine smile, bright even light",
    "upper body shot, looking directly at the camera, serious expression, night city light",
    "portrait, glasses on, reading, warm desk lamp light",
    "medium shot, walking outdoors, motion caught mid-step, afternoon sun",
    "close-up portrait, slight frown, concentrating, cool blue daylight",
]


def build_face_prompt(character: dict, variation: str) -> str:
    return ", ".join([
        character["appearance"].strip(),
        variation,
        character["style"].strip(),
        "full color photograph, sharp focus on the face, natural proportions, plain simple background",
    ])


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--count", type=int, default=18, help="学習用に生成する枚数")
    parser.add_argument("--seed", type=int, help="素体のシード（気に入った顔を再現したいとき）")
    args = parser.parse_args()

    Secrets.load(need_instagram=False)
    character = load_character()
    DATASET.mkdir(parents=True, exist_ok=True)

    # ---- 1. 素体（この子の正体になる1枚） --------------------------------
    print("[1/3] 素体ポートレートを生成")
    anchor_prompt = build_face_prompt(character, VARIATIONS[0])
    anchor = generate.generate(character, anchor_prompt, seed=args.seed, force_base=True)
    anchor_jpeg = media.to_instagram_jpeg(media.download(anchor.url), 1024, 1024)
    anchor_path = media.save_local(anchor_jpeg, DATASET / "000-anchor.jpg")
    reference_url = media.publish_url(anchor_path)
    print(f"    seed={anchor.seed}")
    print(f"    {anchor_path.relative_to(ROOT)}")

    # 以降の生成でこの顔を参照させる
    character = {**character, "identity_reference_url": reference_url, "lora_url": ""}

    # ---- 2. 同じ顔でバリエーション --------------------------------------
    count = min(args.count, len(VARIATIONS))
    print(f"[2/3] 同じ顔で {count} 枚のバリエーションを生成")
    made = [anchor_path]
    for i, variation in enumerate(VARIATIONS[:count], start=1):
        prompt = build_face_prompt(character, variation)
        try:
            out = generate.generate(character, prompt)
        except Exception as exc:  # 1枚失敗しても止めない
            print(f"    {i:03d} 失敗: {exc}")
            continue
        jpeg = media.to_instagram_jpeg(media.download(out.url), 1024, 1024)
        made.append(media.save_local(jpeg, DATASET / f"{i:03d}.jpg"))
        print(f"    {i:03d}/{count} ok")

    # ---- 3. zip --------------------------------------------------------
    print("[3/3] zip にまとめる")
    zip_path = DATASET / "training.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in made:
            zf.write(path, arcname=path.name)

    print()
    print("=" * 70)
    print(f"素材 {len(made)} 枚: {zip_path.relative_to(ROOT)}")
    print()
    print("まず dataset/ の画像を目で見て、顔が揃っていない・崩れている写真は削除してください。")
    print("（そのあと再度 zip する場合: cd dataset && zip training.zip *.jpg）")
    print()
    print("config/character.yaml に貼る値:")
    print(f'  identity_reference_url: "{reference_url}"')
    print()
    print("この時点で `python -m src.main --dry-run` が動きます。")
    print("さらに品質を上げるなら次に: python -m tools.train_lora")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
