"""Instagram Graph API (Content Publishing) クライアント。

投稿は 2 段階:
  1. POST /{ig-user-id}/media          -> コンテナ作成（Meta 側が image_url を取りに来る）
  2. POST /{ig-user-id}/media_publish  -> 公開
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any

import requests

TIMEOUT = 60


class InstagramError(RuntimeError):
    pass


@dataclass
class Instagram:
    user_id: str
    access_token: str
    host: str = "graph.facebook.com"
    version: str = "v25.0"

    @property
    def base(self) -> str:
        return f"https://{self.host}/{self.version}"

    def _call(self, method: str, path: str, **params: Any) -> dict[str, Any]:
        url = f"{self.base}/{path.lstrip('/')}"
        params = {**params, "access_token": self.access_token}
        resp = requests.request(
            method,
            url,
            params=params if method == "GET" else None,
            data=params if method != "GET" else None,
            timeout=TIMEOUT,
        )
        try:
            body = resp.json()
        except ValueError:
            raise InstagramError(f"{resp.status_code}: {resp.text[:500]}") from None
        if "error" in body:
            err = body["error"]
            raise InstagramError(
                f"{err.get('type')} (code {err.get('code')}"
                f"/{err.get('error_subcode')}): {err.get('message')}"
            )
        return body

    # ---------------------------------------------------------------- 投稿

    def create_container(self, image_url: str, caption: str) -> str:
        body = self._call(
            "POST", f"{self.user_id}/media", image_url=image_url, caption=caption
        )
        return body["id"]

    def wait_until_ready(
        self, container_id: str, *, timeout: int = 180, interval: int = 5
    ) -> None:
        """Meta が image_url を取得し終えるまで待つ。"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            body = self._call("GET", container_id, fields="status_code,status")
            code = body.get("status_code")
            if code == "FINISHED":
                return
            if code in {"ERROR", "EXPIRED"}:
                raise InstagramError(
                    f"コンテナが {code} になりました: {body.get('status')}"
                )
            print(f"  … コンテナ準備中 ({code})")
            time.sleep(interval)
        raise InstagramError(f"コンテナ {container_id} が {timeout}s 以内に準備完了しませんでした")

    def publish(self, container_id: str) -> str:
        body = self._call("POST", f"{self.user_id}/media_publish", creation_id=container_id)
        return body["id"]

    def post_image(self, image_url: str, caption: str) -> str:
        container_id = self.create_container(image_url, caption)
        print(f"  コンテナ作成: {container_id}")
        self.wait_until_ready(container_id)
        media_id = self.publish(container_id)
        print(f"  公開完了: {media_id}")
        return media_id

    # ---------------------------------------------------------------- 補助

    def publishing_limit(self) -> dict[str, Any]:
        """24時間あたりの投稿枠の消費状況（上限 100 件）。"""
        body = self._call(
            "GET", f"{self.user_id}/content_publishing_limit",
            fields="config,quota_usage",
        )
        return (body.get("data") or [{}])[0]

    def assert_quota_available(self) -> None:
        try:
            info = self.publishing_limit()
        except InstagramError as exc:
            print(f"  ! 投稿枠を確認できませんでした（続行します）: {exc}")
            return
        used = info.get("quota_usage", 0)
        total = (info.get("config") or {}).get("quota_total", 100)
        print(f"  投稿枠: {used}/{total} (直近24時間)")
        if used >= total:
            raise InstagramError("24時間の投稿上限に達しています。")

    def permalink(self, media_id: str) -> str | None:
        try:
            return self._call("GET", media_id, fields="permalink").get("permalink")
        except InstagramError:
            return None
