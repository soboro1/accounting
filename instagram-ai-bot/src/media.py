"""生成画像を Instagram が受け付ける形に正規化し、公開URLを用意する。

Instagram Graph API の制約:
  - JPEG のみ（PNG/WebP は拒否される）
  - image_url は公開アクセス可能であること（ローカルファイルは投稿できない）
  - 8MB 以下
ここでは fal のファイルストレージをそのままホスティングに使うので、
S3 / R2 などを別途用意する必要はない。
"""
from __future__ import annotations

import io
from pathlib import Path

import fal_client
import requests
from PIL import Image

MAX_BYTES = 8 * 1024 * 1024


def download(url: str) -> bytes:
    resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    return resp.content


def to_instagram_jpeg(raw: bytes, width: int, height: int) -> bytes:
    """RGB JPEG に変換し、指定比率にセンタークロップして 8MB 以下に収める。"""
    img = Image.open(io.BytesIO(raw))
    if img.mode != "RGB":
        img = img.convert("RGB")

    target_ratio = width / height
    w, h = img.size
    if abs(w / h - target_ratio) > 0.01:
        if w / h > target_ratio:              # 横に広い → 左右を削る
            new_w = round(h * target_ratio)
            left = (w - new_w) // 2
            img = img.crop((left, 0, left + new_w, h))
        else:                                  # 縦に長い → 上下を削る
            new_h = round(w / target_ratio)
            top = (h - new_h) // 2
            img = img.crop((0, top, w, top + new_h))

    if img.size != (width, height):
        img = img.resize((width, height), Image.LANCZOS)

    for quality in (92, 85, 78, 70):
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True, subsampling=0)
        data = buf.getvalue()
        if len(data) <= MAX_BYTES:
            return data
    return data


def save_local(data: bytes, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def publish_url(path: Path) -> str:
    """fal のストレージにアップロードして公開URLを得る。"""
    return fal_client.upload_file(str(path))
