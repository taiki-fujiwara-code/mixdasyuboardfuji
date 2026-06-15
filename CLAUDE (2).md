# ATLAPORT CI — AI 向けプロジェクトガイド

**このファイルを最初に読んでください。** プロジェクトの読み方・構成・作業時の指針をまとめています。

---

## 1. プロジェクト概要

- **リポジトリ名**: atlaport-ci
- **本体**: Rails アプリケーション（**メインコードは `app/` 以下**。リポジトリルートは CI/Vagrant/Docker 等の親ディレクトリ）
- **用途**: GMB（Google My Business）関連のダッシュボード・予約・レビュー・アンケート・AI 機能などを含む SaaS 型 Web アプリ。マルチテナント構造。

---

## 2. 技術スタック

| 項目 | 内容 |
|------|------|
| Ruby | 3.3.8（`app/.ruby-version`） |
| Rails | 7.2 |
| DB | MySQL 8.x（utf8mb4, 読み書き分離対応） |
| キャッシュ/ジョブ | Redis + Sidekiq（本番 50 並行、約 180 キュー） |
| フロント | Webpacker, jQuery, Turbolinks |
| 認証 | Devise, devise-two-factor, devise_token_auth（API） |
| 権限 | CanCanCan（`ability.rb` が 158K と巨大） |
| バリデーション | Dry::Schema / Dry::Validation（contracts/） |
| 検索 | Ransack |
| PDF | WickedPDF + wkhtmltopdf |
| 外部 API | Google My Business API, AWS Comprehend/S3/SES, OpenAI, LINE, Instagram, Yahoo, Apple Business, Yext |
| Python 連携 | python-pptx, langchain, llama-index, chromadb（Dockerfile 内でインストール） |
| その他 | Kaminari, Chartkick, CarrierWave, acts_as_paranoid, Lockbox |

---

## 3. ドメインモデル（中核の階層構造）

```
Agent（代理店）
  └─ Client（クライアント企業）
      └─ Branch（店舗）
          └─ Gmb（Google ビジネスプロフィール）
              ├─ GmbReview（口コミ）
              ├─ GmbInsight（インサイト）
              ├─ Questionnaire（アンケート）
              │   ├─ QuestionnaireDetail（設問）
              │   │   └─ QuestionnaireDetailSelect（選択肢）
              │   └─ QuestionnaireAnswer（回答）
              │       └─ QuestionnaireAnswerDetail（回答詳細）
              ├─ Coupon, GmbMedia, WorkFlow, AiAdvice ...
              └─ (Yahoo, Tabelog, Ekiten, Apple 等は Branch 経由)
```

**横断的な概念:**
- **Group**: フォルダ的な階層構造（parent/children, folder_level 1-5）。Branch を GroupBranch で、User を UserGroup で紐付け。
- **User**: Agent/Client/Branch いずれにも所属可能。`role_user_account` enum で master/full/edit/view/approval/custom_role を区別。
- **論理削除**: acts_as_paranoid を使用（GmbReview, User, Report 等）。User は `deleted_flg` カラム。
- **多言語**: Translation モデル（polymorphic: `translatable`）でアンケート等を翻訳。

---

## 4. ディレクトリ構造

```
atlaport-ci/
├── CLAUDE.md              # 本ファイル
├── app/                    # ★ Rails アプリ本体
│   ├── app/
│   │   ├── controllers/   # コントローラ（concerns/ 内に共通処理多数）
│   │   ├── models/        # モデル 296 ファイル（concerns/ 20, master/ あり）
│   │   ├── views/         # ERB ビュー（layouts/ にメインレイアウト）
│   │   ├── services/      # Service Object 314 ファイル
│   │   ├── workers/       # Sidekiq Worker 355 ファイル
│   │   ├── queries/       # Query Object（複雑 SQL のカプセル化）
│   │   ├── contracts/     # Dry::Schema / Dry::Validation
│   │   ├── validators/    # カスタム ActiveModel Validator
│   │   ├── errors/        # カスタム例外
│   │   ├── serializers/   # ErrorSerializer
│   │   ├── helpers/       # ApplicationHelper 2800 行超
│   │   ├── mailers/
│   │   ├── uploaders/     # CarrierWave
│   │   ├── channels/      # Action Cable
│   │   ├── assets/
│   │   └── javascript/
│   ├── config/
│   │   ├── routes.rb      # 1200 行超（namespace: api/v1, gmbs, groups, branches, ai 等）
│   │   ├── database.yml   # primary + primary_replica（読み書き分離）
│   │   ├── sidekiq.yml    # 約 180 キュー定義
│   │   ├── settings.yml   # 約 3700 行のアプリ設定
│   │   ├── schedule.rb    # Whenever cron
│   │   ├── locales/       # ja.yml (429KB), en.yml, ja_api.yml
│   │   └── initializers/  # sidekiq, devise, datadog, lograge 等
│   ├── lib/
│   │   ├── batches/       # バッチ処理 97 ファイル
│   │   ├── tasks/         # Rake タスク多数
│   │   ├── atlaport/      # 外部 API ラッパー（Facebook, Instagram, LINE, Twitter, Yahoo）
│   │   ├── google/apis/   # Google API 統合
│   │   ├── apple_business/
│   │   └── python/        # Python 連携（レポート出力等）
│   ├── spec/              # RSpec（FactoryBot, DatabaseCleaner）
│   ├── docker-compose.yml
│   ├── Dockerfile
│   └── Gemfile
├── app/docs/              # 設計・手順（architecture_*.md, screen_list.md 等）
├── .github/workflows/     # CI（rubocop.yml: 変更ファイルのみ Rubocop）
├── ansible/               # 構成管理
└── Vagrantfile
```

---

## 5. 設計パターンと規約

### ファイル命名規則
| レイヤー | サフィックス | 例 |
|----------|-------------|-----|
| Service | `*Service` | `CrawlGoogleReviewService` |
| Worker | `*Worker` | `AiAdviceWorker` |
| Query | `*Query` | `ReviewSummaryQuery` |
| Contract | `*Contract` / `*ParamsSchema` | `BookingParamsSchema` |
| Validator | `*Validator` | `BaseUriValidator` |

### Service Object
- `ApplicationService` を継承。`call` メソッドで処理を実行。

### Contract（バリデーション）
- `ApplicationSchema` / `ApplicationContract` を継承（Dry::Schema / Dry::Validation）。
- トップレベルのスキーマ検証 → ネストされた Contract 検証、の 2 段階。
- エラーメッセージは I18n 経由で日本語。

### Query Object
- 複雑な SQL をカプセル化。`sanitize_sql_array` でサニタイズ、ヒアドキュメント（`<<~SQL`）で記述。
- `Rails.cache.fetch` でキャッシュを多用。

### ApplicationController の共通処理
- `check_ip_request`: IP 制限
- `check_expired_login`: ログイン期限
- `authenticate_current_user`: 認証
- `detect_action_resource` / `current_ability` / `check_user_permission`: CanCanCan 権限
- `render_200(msg)` / `render_400(msg)`: JSON レスポンスヘルパー

### Rubocop
- `.rubocop.yml` あり。`Layout/LineLength` 最大 130 文字。`Style/Documentation` 無効。

---

## 6. 開発環境（Docker）

- **compose**: `app/docker-compose.yml`
- **サービス**: `db`（MySQL 8.0, :3306）, `redis`(:6379), `app`（Rails :8858）, `webpack`(:3035）, `sidekiq`, `mailcatcher`(:1080）
- **起動**: `cd app && docker-compose up`
- **Dockerfile**: `ruby:3.3.8` ベース。Node.js, Yarn, Python 3.11, LibreOffice, wkhtmltopdf, ffmpeg, kakasi, 日本語フォント等をインストール。
- **DB 投入**: `db.sql` をコンテナに `docker cp` → `mysql` で投入（詳細は `app/README.md`）。
- **メール確認**: `http://localhost:1080`（Mailcatcher）

---

## 7. 環境変数（主要カテゴリ）

`.env.examples` を参照。主な変数:

| カテゴリ | 変数例 |
|----------|--------|
| DB | `DB_HOST`, `DB_USERNAME`, `DB_PASSWORD`, `READ_DB_HOST` |
| Redis | `REDIS_HOST`, `REDIS_PORT`, `REDIS_DB_ID` |
| AWS | `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY`, `AWS_BUCKET_NAME` |
| メール | `SES_SERVER`, `SES_ACCESS_KEY`, `SMTP_USERNAME` |
| Google | `GOOGLE_API_KEY`, `GOOGLE_API_KEY_PRIMARY` |
| SNS | `FACEBOOK_APP_ID`, `LINE_CO_CLIENT_ID`, `IG_APP_ID` |
| AI | `OPENAI_TOKEN`, `OPENAI_DEFAULT_MODEL` |
| ECS/Fargate | `ECS_CLUSTER_ARN`, `ECS_SERVICE_NAME` |
| 暗号化 | `ENCRYPTION_KEY`, `ACTIVE_RECORD_ENCRYPTION_PRIMARY_KEY` |
| Sidekiq | `SIDEKIQ_DASHBOARD_USERNAME`, `SIDEKIQ_DASHBOARD_PASSWORD` |

---

## 8. ルーティングの読み方

`app/config/routes.rb`（1200 行超）の主な名前空間:

| 名前空間 / パス | 内容 |
|------------------|------|
| `api/v1` | トークン認証 API（devise_token_auth） |
| `devise_scope :user` | ユーザー認証関連（多数のカスタムルート） |
| `resources :gmbs` | GMB 管理（ネスト: faqs, booking_hps, review_hps, sentiment_analyses 等） |
| `resources :groups` | グループ管理（ネスト: reviews, surveys, booking_medias 等） |
| `resources :branches` | 店舗管理（ネスト: instagrams, line_steps） |
| `namespace :ai` | AI 機能（レビュー返信生成、画像生成、キーワード検出等） |
| `resources :reports` | レポート |
| `mount Sidekiq::Web` | `/admin/sidekiq`（BASIC 認証付き） |

**探し方**: コントローラ名やパスで `grep` して該当箇所を読むのが効率的。

---

## 9. モデル concerns（よく使われる共通モジュール）

| Concern | 役割 |
|---------|------|
| `CommonScope` (13K) | 多くのモデルで include される汎用スコープ |
| `TrackingUpdateInfoConcern` (26K) | 更新履歴の追跡 |
| `AgentClientValidationConcern` | Agent/Client バリデーション |
| `AiAdviceConcern` / `AiAdviceAasmConcern` | AI アドバイス機能・状態管理 |
| `ImportData` | データインポート |
| `Statusable` | ステータス管理 |
| `StripAttributes` | 属性のトリミング |
| `OtpTwoFactorConcern` | 二要素認証 |
| `DbConnectionManagerConcern` | DB 接続管理 |

---

## 10. 重要ファイル早見表

| 知りたいこと | 場所 |
|--------------|------|
| ルーティング | `app/config/routes.rb` |
| DB スキーマ | `app/db/schema.rb` |
| DB 設定（読み書き分離） | `app/config/database.yml` |
| Sidekiq キュー | `app/config/sidekiq.yml` |
| アプリ設定 | `app/config/settings.yml`（3700 行） |
| 権限定義 | `app/app/models/ability.rb`（158K, 巨大） |
| 定数 | `app/app/models/constants.rb`（73K） |
| 認証 | `app/app/controllers/users/`（Devise 系） |
| API 認証 | `app/app/controllers/api/v1/` |
| テスト | `app/spec/`（RSpec, FactoryBot, DatabaseCleaner） |
| Cron 設定 | `app/config/schedule.rb`（Whenever） |
| ロケール（日本語） | `app/config/locales/ja.yml`（429KB） |
| CI | `.github/workflows/rubocop.yml` |
| 環境変数テンプレート | `app/.env.examples` |

---

## 11. AI が作業するときの指針

1. **まずこの CLAUDE.md を読む** → 全体像を把握してから必要な箇所だけ開く。
2. **変更は `app/` 以下に限定** — `ansible/`, `light-bootstrap-*` 等は通常触らない。
3. **階層を意識** — Agent → Client → Branch → Gmb の流れを理解しておくと、モデル関連やスコープの意図がわかる。
4. **巨大ファイルは grep で絞る** — `routes.rb`（1200 行）, `ability.rb`（158K）, `constants.rb`（73K）, `settings.yml`（3700 行）, `ja.yml`（429K）は全読みせず grep。
5. **命名規則に従う** — Service は `*Service`, Worker は `*Worker`, Query は `*Query`, Contract は `*ParamsSchema` / `*Contract`。
6. **バリデーション** — モデル層は ActiveRecord、パラメータ層は Dry::Validation（contracts/）。用途に応じて使い分け。
7. **SQL** — 複雑なクエリは Query Object に切り出し、`sanitize_sql_array` + ヒアドキュメントで書く。
8. **キャッシュ** — `Rails.cache.fetch` を多用。変更時はキャッシュキーの影響を確認。
9. **テスト** — `app/spec/` に RSpec + FactoryBot。変更に応じて spec の存在確認と更新を行う。
10. **I18n** — エラーメッセージは `ja.yml` から引く。ハードコードの日本語は避ける。
11. **論理削除** — `acts_as_paranoid` 使用モデルは `destroy` が論理削除になる。`really_destroy!` で物理削除。

---

## 12. 参照ドキュメント

- ビルド・DB・デプロイ: `app/README.md`
- 画面一覧: `screen_list.md`
- API 仕様: `app/docs/me.yaml`（OpenAPI 3.0）
- Claude Code お渡し手順: `docs/claude-code-handoff.md`

---

## 13. 仕様・画面設計の調査（企画・ディレクション向け）

### 正とするブランチ

- **実装の正**: `runsystem_ruby3`（事実上の develop。仕様確認・画面設計の質問はこのブランチのコードを前提にする）
- `master`: GitHub の default だが、最新機能の正ではない可能性あり
- `runsystem`: 旧系統。新規調査では使わない

ブランチを切り替えてから Claude Code を起動するか、質問の最初に「runsystem_ruby3 前提で」と書くこと。

### 調査の優先順位（コードを書かず読む）

| 知りたいこと | まず見る場所 |
|--------------|--------------|
| 画面一覧・画面名 | `screen_list.md`（282 画面） |
| URL・コントローラ対応 | `app/config/routes.rb`（grep で絞る） |
| API 仕様 | `app/docs/me.yaml` |
| 画面の実装イメージ | `app/app/views/` の該当 ERB |
| ビジネスルール | 該当 `*Service` / モデル / `ability.rb`（grep） |

### 仕様確認モード（振る舞い）

- **デフォルトは読み取り専門**: 仕様説明・差分整理・設計案の提示まで。ファイル変更・マイグレーション・本番設定変更は依頼がない限り行わない。
- **回答形式**: 現状仕様（根拠ファイルパス付き）→ 不明点 → 設計案（任意）の順。
- **巨大ファイル**: §11 のとおり `ability.rb` / `routes.rb` / `ja.yml` は grep 前提。

---

*このファイル（CLAUDE.md）は AI がプロジェクトを効率的に読むために用意されています。内容が古くなったらプロジェクトの実態に合わせて更新してください。*
