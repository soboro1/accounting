"""Claude API でキャプションとハッシュタグを生成する。"""
from __future__ import annotations

import json
import re
from typing import Any

import anthropic

MODEL = "claude-sonnet-5"

# Meta は AI 生成コンテンツの明示を求めている。ここは外さないこと。
REQUIRED_HASHTAGS = ["#AI生成", "#aiart", "#aigenerated"]

SYSTEM = """あなたは Instagram の投稿キャプションを書くコピーライターです。
与えられたキャラクター設定とシーンから、その人自身が書いたような一人称のキャプションを作ります。

守ること:
- 日本語。2〜3行、合計120文字以内。改行で区切る。
- 説明文にしない。写真の解説ではなく、その瞬間の気分や独り言を書く。
- 「〜な一日でした」のような定型のまとめ方はしない。
- 誇張した美辞麗句、「#今日のコーデ」的なテンプレ感を避ける。
- ハッシュタグは日本語と英語を混ぜて 8〜12 個。シーンに実際に関係するものだけ。
- 出力は JSON のみ。前後に説明を書かない。

出力形式:
{"caption": "本文", "hashtags": ["#tag1", "#tag2"]}"""


def _extract_json(text: str) -> dict[str, Any]:
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError(f"JSON を取り出せませんでした: {text[:300]}")
    return json.loads(match.group(0))


def build(
    api_key: str, character: dict[str, Any], scene: dict[str, Any]
) -> tuple[str, list[str]]:
    client = anthropic.Anthropic(api_key=api_key)

    user = f"""# キャラクター
名前: {character['name']}
人物像:
{character['persona'].strip()}

# 今回のシーン
場所・状況: {scene['setting']}
服装: {scene['outfit']}
様子: {scene['pose']}
本人の気分: {scene.get('mood', '')}"""

    resp = client.messages.create(
        model=MODEL,
        max_tokens=800,
        system=SYSTEM,
        messages=[{"role": "user", "content": user}],
    )
    data = _extract_json("".join(b.text for b in resp.content if b.type == "text"))

    caption = str(data["caption"]).strip()
    tags = [str(t).strip() for t in data.get("hashtags", []) if str(t).strip()]
    tags = [t if t.startswith("#") else f"#{t}" for t in tags]

    lower = {t.lower() for t in tags}
    for required in REQUIRED_HASHTAGS:
        if required.lower() not in lower:
            tags.append(required)

    return caption, tags[:30]  # Instagram のハッシュタグ上限は 30


def render(caption: str, hashtags: list[str]) -> str:
    return f"{caption}\n\n{' '.join(hashtags)}"
