"""投稿前のヘルスチェック。認証・アカウント種別・投稿枠・トークン期限を確認する。

  python -m tools.check_setup
"""
from __future__ import annotations

import datetime as dt

import requests

from src.config import Secrets, load_character, load_scenes
from src.instagram import Instagram, InstagramError


def main() -> int:
    ok = True
    secrets = Secrets.load()

    print("― 設定ファイル ―")
    character = load_character()
    scenes = load_scenes()
    print(f"  キャラクター: {character['name']}")
    print(f"  シーン数: {len(scenes)}")
    if character.get("lora_url"):
        print("  顔の固定: LoRA (Phase 2)")
    elif character.get("identity_reference_url"):
        print("  顔の固定: PuLID 参照画像 (Phase 1)")
    else:
        ok = False
        print("  顔の固定: 未設定 — 毎回別人になります。")
        print("            先に `python -m tools.bootstrap_identity` を実行してください。")

    print("\n― Instagram アカウント ―")
    ig = Instagram(
        user_id=secrets.ig_user_id,
        access_token=secrets.ig_access_token,
        host=secrets.ig_api_host,
        version=secrets.ig_api_version,
    )
    try:
        me = ig._call("GET", ig.user_id, fields="username,account_type,media_count,followers_count")
        print(f"  @{me.get('username')} / {me.get('account_type', 'n/a')}")
        print(f"  投稿数 {me.get('media_count')} / フォロワー {me.get('followers_count')}")
    except InstagramError as exc:
        ok = False
        print(f"  取得失敗: {exc}")

    try:
        info = ig.publishing_limit()
        total = (info.get("config") or {}).get("quota_total", 100)
        print(f"  投稿枠: {info.get('quota_usage', 0)}/{total} (直近24時間)")
    except InstagramError as exc:
        print(f"  投稿枠の取得失敗: {exc}")

    print("\n― アクセストークン ―")
    try:
        resp = requests.get(
            f"https://{secrets.ig_api_host}/debug_token",
            params={"input_token": secrets.ig_access_token,
                    "access_token": secrets.ig_access_token},
            timeout=30,
        ).json()
        data = resp.get("data", {})
        expires = data.get("expires_at", 0)
        if expires == 0:
            print("  有効期限: なし（System User トークン。運用向き）")
        else:
            when = dt.datetime.fromtimestamp(expires, dt.timezone.utc)
            days = (when - dt.datetime.now(dt.timezone.utc)).days
            print(f"  有効期限: {when:%Y-%m-%d} (あと {days} 日)")
            if days < 14:
                ok = False
                print("  ! まもなく切れます。README の『トークンを切らさない』を参照。")
        scopes = data.get("scopes", [])
        needed = {"instagram_basic", "instagram_content_publish"}
        missing = needed - set(scopes)
        if missing and scopes:
            ok = False
            print(f"  ! 権限が足りません: {', '.join(sorted(missing))}")
    except Exception as exc:
        print(f"  確認できませんでした（Instagram Login では未対応）: {exc}")

    print()
    print("すべて問題なし。" if ok else "上の ! を解消してください。")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
