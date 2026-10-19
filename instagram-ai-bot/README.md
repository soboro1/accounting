# instagram-ai-bot

同じ顔のAIキャラクターの写真を毎日自動生成して、Instagram に投稿します。

```
config/scenes.yaml ──┐
                     ├─→ プロンプト組み立て ─→ fal.ai で生成（顔固定）
config/character.yaml┘                              │
                                                    ├─→ JPEG 正規化 + 公開URL化
Claude API ─→ キャプション + ハッシュタグ ──────────┤
                                                    ↓
                                    Instagram Graph API（コンテナ作成 → 公開）
                                                    ↓
                                       GitHub Actions が 1日2回 実行
```

- **顔の固定**: 参照画像1枚の PuLID（学習不要）か、学習済み LoRA。後者のほうが一貫性・画質ともに上。
- **画像ホスティング不要**: Instagram は「公開URL上のJPEG」しか受け付けませんが、fal のストレージをそのまま使うので S3 / R2 は要りません。
- **同じシーンの連投を回避**: `state/history.jsonl` に履歴を残し、直近8件と重複しないよう抽選します。

---

## 1. 必要なもの

| | 用途 | 取得先 |
|---|---|---|
| fal.ai APIキー | 画像生成・LoRA学習・ファイル置き場 | https://fal.ai/dashboard/keys |
| Anthropic APIキー | キャプション生成 | https://console.anthropic.com/settings/keys |
| Instagram プロアカウント | 投稿先 | Instagramアプリ → 設定 → アカウントの種類 |
| Facebook ページ | Graph API の前提 | 上のIGアカウントと連携する |
| Meta 開発者アプリ | アクセストークン発行 | https://developers.facebook.com/apps |

---

## 2. Instagram 側のセットアップ

ここが一番つまずきやすいので順番どおりに。

1. **プロアカウントに切り替える**
   Instagram アプリ → 設定 → アカウントの種類とツール → プロアカウントに切り替える（**ビジネス**を選ぶ）。
   個人アカウントのままでは Graph API から投稿できません。

2. **Facebook ページと連携する**
   Facebook ページを1つ作り、Instagram アプリの 設定 → シェアと再投稿 / アカウントセンター からそのページとリンクします。

3. **Meta 開発者アプリを作る**
   https://developers.facebook.com/apps → アプリを作成 → ユースケースで **「Instagram」** を選択 → 製品に **Instagram** を追加。

4. **アクセストークンを発行する**
   [Graph API Explorer](https://developers.facebook.com/tools/explorer/) で作ったアプリを選び、次の権限を付けてトークンを生成します。

   ```
   instagram_basic
   instagram_content_publish
   pages_show_list
   pages_read_engagement
   ```

5. **長期トークンに交換する**（短期トークンは1時間で切れます）

   ```bash
   curl -s "https://graph.facebook.com/v25.0/oauth/access_token?grant_type=fb_exchange_token&client_id=APP_ID&client_secret=APP_SECRET&fb_exchange_token=SHORT_LIVED_TOKEN"
   ```

6. **IG_USER_ID を調べる**

   ```bash
   curl -s "https://graph.facebook.com/v25.0/me/accounts?access_token=LONG_LIVED_TOKEN"
   ```

   返ってきたページの `id` を使って:

   ```bash
   curl -s "https://graph.facebook.com/v25.0/PAGE_ID?fields=instagram_business_account&access_token=LONG_LIVED_TOKEN"
   ```

   `instagram_business_account.id` が `IG_USER_ID` です。

> **アプリ審査は要りません。** 自分がアプリの管理者になっている Instagram アカウントに投稿するだけなら、アプリが開発モードのままで動きます。他人のアカウントに投稿させる機能を公開する場合だけ App Review が必要です。

### トークンを切らさない

長期ユーザートークンの寿命は **60日** です。毎回手で更新するのは現実的ではないので、運用に入るなら
[ビジネスマネージャ](https://business.facebook.com/settings/system-users) で **システムユーザー** を作り、
そこから発行したトークン（**無期限**）を使ってください。これが唯一の放置できる構成です。

現在のトークンの残り日数は `python -m tools.check_setup` で確認できます。

---

## 3. ローカルのセットアップ

```bash
cd instagram-ai-bot
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # 中身を埋める
```

---

## 4. キャラクターを作る

### 4-1. 見た目を決める

`config/character.yaml` の `appearance` と `style` を書き換えます。ここが顔の設計図です。

### 4-2. 顔を確定させる（Phase 1）

```bash
python -m tools.bootstrap_identity --count 18
```

素の Flux で正面ポートレートを1枚作り、その顔を参照して角度・表情・光を変えた18枚を生成、
`dataset/` に保存して zip までまとめます。

終わると `identity_reference_url` が表示されるので `config/character.yaml` に貼ってください。
**この時点でもう運用を始められます。**

気に入らない顔が出たら、そのまま何度でも実行し直してください（毎回別人が出ます）。
気に入った顔の `seed` が表示されるので、`--seed 12345` で再現もできます。

### 4-3. LoRA を学習する（Phase 2・推奨）

まず `dataset/` を目で見て、**顔が揃っていない写真・手や目が崩れた写真を削除**します。
枚数より質のほうが効きます。消したら zip を作り直して:

```bash
cd dataset && zip -r training.zip *.jpg && cd ..
python -m tools.train_lora --trigger RINCHAR
```

数分〜十数分で `lora_url` が出ます。`config/character.yaml` の `lora_url` と `trigger_phrase` に貼れば、
以降の生成は自動的に LoRA 経由に切り替わります。

---

## 5. テスト

```bash
python -m tools.check_setup    # 認証・アカウント種別・投稿枠・トークン期限
python -m src.main --dry-run   # 生成とキャプションだけ。投稿はしない
```

`out/` に画像とキャプションが出ます。納得できるまで `--dry-run` を回してください。
特定のシーンだけ試すなら `--scene cafe_morning`。

本番:

```bash
python -m src.main
```

---

## 6. GitHub Actions で自動化

1. このディレクトリを GitHub の**プライベート**リポジトリに push します。
2. Settings → Secrets and variables → Actions → **Secrets** に登録:

   | 名前 | 値 |
   |---|---|
   | `FAL_KEY` | fal.ai のキー |
   | `ANTHROPIC_API_KEY` | Anthropic のキー |
   | `IG_USER_ID` | Instagram ビジネスアカウントID |
   | `IG_ACCESS_TOKEN` | 長期 or システムユーザートークン |

3. Actions タブ → `Instagram auto post` → **Run workflow** で手動実行できます。
   `dry_run` にチェックを入れれば投稿せず生成だけ試せます。

デフォルトは **1日2回（JST 8:00 / 20:00）**。頻度を変えるなら `.github/workflows/post.yml` の `cron` を編集してください（UTC 指定）。

> GitHub の cron は混雑時に **5〜30分ほど遅れる**ことがあります。分単位の正確さが要る用途には向きません。

履歴 `state/history.jsonl` はワークフローが自動でコミットして戻します。生成画像は Artifacts に14日残ります。

---

## 7. 運用上の注意

**AI であることを明示する。**
Meta は AI 生成コンテンツのラベル表示を求めています。キャプションには `#AI生成 #aiart #aigenerated` が
自動で付きますが、**プロフィール文にも「AI generated / AIで生成した架空の人物です」と明記**してください。
実在の人物と誤認させる運用は、なりすまし扱いでアカウント削除の対象になります。

**投稿上限は24時間で100件。** `assert_quota_available()` が毎回チェックし、上限に達していれば投稿せず止まります。

**画像の要件**（`src/media.py` が自動で満たします）
JPEG のみ / 8MB以下 / アスペクト比 4:5〜1.91:1。デフォルトは 1080×1350（フィードで最も面積が大きい 4:5）。

**シーンは増やすほどいい。** 20件だと10日でひと回りします。`config/scenes.yaml` を50〜100件まで育てるのが理想です。

**セーフティチェッカーは切らない。** `enable_safety_checker: True` で生成し、弾かれたら別シードで最大3回引き直します。
露出の多いプロンプトは fal 側でも Instagram 側でも弾かれるので、シーン定義はSFWの範囲で書いてください。

---

## 8. コスト目安（1日2投稿）

| 項目 | 概算 |
|---|---|
| 画像生成（Flux, 1.46MP × 60枚/月） | $2〜4 / 月 |
| キャプション（Claude Sonnet） | $1未満 / 月 |
| LoRA 学習 | $2前後 / 1回きり |
| GitHub Actions | 無料枠内 |

**合計 月$3〜5 程度。** 単価は変わるので fal.ai の料金ページで確認してください。

---

## ファイル構成

```
config/character.yaml   キャラ設定（顔・作風・モデル・画像サイズ）
config/scenes.yaml      シーンプール
src/prompt_builder.py   シーン抽選とプロンプト組み立て
src/generate.py         fal.ai 呼び出し（LoRA / PuLID の自動切替）
src/media.py            JPEG 正規化・クロップ・公開URL化
src/caption.py          Claude API でキャプション生成
src/instagram.py        Graph API クライアント
src/state.py            投稿履歴
src/main.py             1投稿分のオーケストレーション
tools/bootstrap_identity.py  顔を作って学習素材を用意する
tools/train_lora.py     LoRA 学習
tools/check_setup.py    ヘルスチェック
```
