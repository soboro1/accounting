"""1 回分の投稿を実行する。

  python -m src.main            # 生成 → キャプション → Instagram 投稿
  python -m src.main --dry-run  # 投稿せず、画像とキャプションを out/ に出すだけ
  python -m src.main --scene cafe_morning  # シーンを指定
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime

from . import caption as caption_mod
from . import generate, media, prompt_builder, state
from .config import ROOT, Secrets, load_character, load_scenes
from .instagram import Instagram


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Instagram に投稿しない")
    parser.add_argument("--scene", help="シーンIDを指定（省略時はランダム）")
    args = parser.parse_args()

    secrets = Secrets.load(need_instagram=not args.dry_run)
    character = load_character()
    scenes = load_scenes()

    if args.scene:
        matched = [s for s in scenes if s["id"] == args.scene]
        if not matched:
            print(f"シーン '{args.scene}' が config/scenes.yaml にありません。", file=sys.stderr)
            return 1
        scene = matched[0]
    else:
        scene = prompt_builder.pick_scene(scenes)

    print(f"[1/5] シーン: {scene['id']}")

    prompt = prompt_builder.build_image_prompt(character, scene)
    print(f"[2/5] 画像生成")
    print(f"    prompt: {prompt[:160]}…")
    result = generate.generate_with_retry(character, prompt)

    print("[3/5] Instagram 向けに正規化 → アップロード")
    raw = media.download(result.url)
    jpeg = media.to_instagram_jpeg(raw, character["image"]["width"], character["image"]["height"])
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    local = media.save_local(jpeg, ROOT / "out" / f"{stamp}-{scene['id']}.jpg")
    print(f"    {local.relative_to(ROOT)} ({len(jpeg) / 1024:.0f} KB)")
    public_url = media.publish_url(local)

    print("[4/5] キャプション生成")
    text, tags = caption_mod.build(secrets.anthropic_api_key, character, scene)
    full_caption = caption_mod.render(text, tags)
    print("―" * 40)
    print(full_caption)
    print("―" * 40)

    if args.dry_run:
        (local.with_suffix(".txt")).write_text(full_caption, encoding="utf-8")
        print("[5/5] --dry-run のため投稿はスキップしました。")
        return 0

    print("[5/5] Instagram へ投稿")
    ig = Instagram(
        user_id=secrets.ig_user_id,
        access_token=secrets.ig_access_token,
        host=secrets.ig_api_host,
        version=secrets.ig_api_version,
    )
    ig.assert_quota_available()
    media_id = ig.post_image(public_url, full_caption)

    state.append({
        "scene_id": scene["id"],
        "model": result.model,
        "seed": result.seed,
        "prompt": prompt,
        "image_url": public_url,
        "caption": full_caption,
        "ig_media_id": media_id,
        "permalink": ig.permalink(media_id),
    })
    print("完了。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
