# Slide Engine - 運用・実行マニュアル

**Python 3.12 / Pydantic V2 / Fail-Soft Architecture によるエンタープライズ向け PowerPoint 自動生成エンジン**

本ドキュメントは、システムの環境構築、CLI / Docker 実行、企業スライドマスター（`.pptx`）の配置、および本番運用時の監視・トラブルシューティング手順を定義します。

> **新規テンプレートを追加開発する場合（生成AIへの指示用）:**
> 新しいスライドテンプレート（`CORE-XX`）を追加する際は、別紙 **`TEMPLATE_DEV_GUIDE.md`** を生成AIに読み込ませてください。

---

## 1. ディレクトリ構成

```text
slide-engine/
├── .github/workflows/ci.yml     # CIパイプライン定義 (Ruff, Mypy strict, Pytest, Docker)
├── assets/                      # 企業マスター (.pptx) 格納ディレクトリ
│   └── corporate_blue.pptx      # デフォルト参照されるベーステンプレート (任意配置)
├── examples/
│   └── sample_deck.json         # 全テンプレート・複合分割を網羅したサンプルJSON
├── output/                      # 生成された .pptx の出力先ディレクトリ
├── src/slide_engine/
│   ├── __init__.py              # パッケージ公開API
│   ├── __main__.py              # CLIエントリーポイント (STDIN/STDOUT・終了コード制御)
│   ├── core_base.py             # 基盤層 (SlideMaster, BoundingBox, FontEngine, ShapeBuilder)
│   ├── models.py                # データ契約層 (Pydantic V2 スキーマ・境界防御・レジストリ)
│   ├── controller.py            # オーケストレーター (動的ディスパッチ・退避制御・テレメトリ)
│   └── render_core_*.py         # 自己完結プラグイン群 (Payload定義・描画・サンプルを1ファイルに集約)
├── tests/
│   └── test_engine.py           # 包括的ユニットテスト・シグネチャ検査・境界防御テスト
├── Dockerfile                   # マルチステージ・非root実行コンテナ定義
├── docker-compose.yml           # ローカル実行・ボリュームマウント定義
├── pyproject.toml               # 依存関係・Ruff (TID251, SLF001, C901)・Mypy設定
├── Rule.md                      # アーキテクチャ憲法・開発規約
├── README.md                    # 本ファイル（人間向け運用マニュアル）
└── TEMPLATE_DEV_GUIDE.md        # 生成AI向け：新規テンプレート拡張開発リファレンス
```

---

## 2. セットアップと品質検査コマンド

Python 3.12 以上の環境で動作します。

```bash
# 仮想環境の作成とインストール
python3.12 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"

# 1. リンター＆複雑度チェック (TID251: 禁止API, SLF001: プライベート参照, C901: 複雑度10以下)
ruff check src tests

# 2. 厳格な静的型チェック
mypy src tests

# 3. テストスイートの実行
pytest tests/ -v
```

---

## 3. 実行方法（CLI / 標準入出力 / Docker Compose）

```bash
# パターンA: ローカル CLI でファイルを指定して実行
slide-engine --input examples/sample_deck.json --output output/presentation.pptx

# パターンB: 標準入出力 (STDIN / STDOUT) パイプライン実行
cat examples/sample_deck.json | slide-engine > output/streamed.pptx

# パターンC: Docker Compose による実行 (ホストにPython環境がなくても動作)
docker compose run --rm engine \
  --input /app/examples/sample_deck.json \
  --output /app/output/generated.pptx

# パターンD: CI/CD・品質保証用：退避スライド発生時に異常終了する厳格モード
slide-engine --input examples/sample_deck.json --output output/presentation.pptx --fail-on-fallback
```

---

## 4. 企業スライドマスター（.pptx）の配置と継承仕様

1. **配置場所:** PowerPoint で作成したマスターファイルを `assets/{theme_name}.pptx`（デフォルトは `assets/corporate_blue.pptx`）として配置します。
2. **不在時のフォールバック:** 指定された `.pptx` が存在しない場合、エラー終了せず、自動的に 16:9（13.333 × 7.5 インチ）の白紙プレゼンテーションをメモリ上で初期化します。
3. **既存ダミースライドの自動パージ:** マスター `.pptx` 内にサンプル用のスライドが残っていても、ロード時にすべて自動削除され、マスター定義のみが継承されます。
4. **レイアウトの自動名前解決:**
   - `CORE-01`（表紙）は、レイアウト名に `title`, `cover`, `タイトル`, `表紙` を含むものを優先適用します（該当なしの場合はインデックス `0`）。
   - その他のスライドは、`blank`, `title and content`, `白紙`, `コンテンツ` を含むものを優先適用します（該当なしの場合はインデックス `6` または `0`）。
5. **初期プレースホルダー枠の自動消去:**
   スライド生成直後に `clean_slide_canvas()` が実行され、「タイトルを入力」等のデフォルト空プレースホルダー枠はすべて自動消去されます（マスター側の背景グラフィック・企業ロゴ・フッターはそのまま維持されます）。

---

## 5. 標準テンプレートカタログ一覧

| ID | 名称 | 用途 | 左右分割 |
| :--- | :--- | :--- | :---: |
| `CORE-00` | システム退避 (Fallback) | 例外発生時の自動フォールバック専用 | 不可 |
| `CORE-01` | 表紙 (Cover) | プレゼンテーション表紙タイトル・発表者・日付 | 不可 |
| `CORE-02` | アジェンダ (Agenda) | 目次・セクション中扉・現在地ハイライト（1〜8項目） | 不可 |
| `CORE-10` | データテーブル (Table) | 可変構造化表（1〜8列 × 1〜15行）・特定行ハイライト | **対応** |
| `CORE-20` | プロセスフロー (Process) | 時系列・業務ステップ（2〜6段階） | **対応** |
| `CORE-30` | ネイティブチャート (Chart) | OpenXML 縦棒・横棒・折れ線・円グラフ | **対応** |
| `CORE-40` | 2×2マトリクス (Matrix) | 縦軸・横軸ラベル付き4象限ポジショニングマップ | **対応** |
| `CORE-50` | ビジュアル枠 (Visual Frame) | 手動図解・スクショ貼付用ガイド枠＋下部キャプション | **対応** |
| `CORE-60` | KPIカード (KPI Cards) | 重要指標ビッグナンバー並列カード（2〜4枚） | **対応** |
| `CORE-70` | 左右対比 (Comparison) | Before/After・As-Is/To-Be 比較ボックス | **対応** |
| `CORE-80` | ウォーターフォール (Waterfall) | 累積・増減ステップチャート（2〜8段階） | **対応** |

※ 具体的な入力 JSON の記述例については `examples/sample_deck.json` を参照してください。

---

## 6. 運用監視・終了コード・トラブルシューティング

### 6.1 CLI 終了コード（Exit Codes）
| 終了コード | 状態 | 意味・発生原因 |
| :---: | :--- | :--- |
| **`0`** | **正常完走** | 生成・出力完了（デフォルトでは一部スライドが `CORE-00` 退避になった場合もフェイルソフトとして `0` で完走します）。 |
| **`1`** | **入力検証エラー** | JSON 構文エラー、未知の `template_id`、配列上限超過、チャート要素数の不一致（Pydantic で遮断）。 |
| **`2`** | **エンジン致命的エラー** | デッキ全体の生成継続が不可能なシステム例外。 |
| **`3`** | **出力 I/O エラー** | 出力先ディレクトリの権限不足やディスクフルによる書き込み失敗。 |
| **`4`** | **厳格退避エラー** | `--fail-on-fallback` 指定時に 1 枚以上のスライドが `CORE-00` 退避になった場合の異常終了（CI/CD ゲート用）。 |

### 6.2 構造化テレメトリログ（Silent Degradation の監視）
生成完了時、**標準エラー出力（`STDERR`）** に以下の JSON ログが1行出力されます。バッチ運用時は `fallback_slides > 0` を監視してアラートを発報してください。
```json
{"event": "deck_generation_summary", "total_slides": 10, "rendered_slides": 9, "fallback_slides": 1, "duration_ms": 142.5}
```

### 6.3 退避スライド（CORE-00）発生時の原因調査
生成された `.pptx` に黄色い `[SYSTEM FALLBACK]` スライドが出力された場合は、**PowerPoint で該当スライドの「ノート（スピーカーノート）」欄を確認**してください。例外の完全な Python スタックトレースが自動記録されています。