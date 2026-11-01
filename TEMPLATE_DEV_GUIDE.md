# Slide Engine - 新規テンプレート拡張開発リファレンス (For AI Coding Agents)

本ドキュメントは、生成AI（コーディングエージェント）が Slide Engine に新しいスライドテンプレート（`CORE-XX`）を追加実装する際の**絶対遵守仕様書**である。
本ドキュメント内の API 仕様と制約に100%従ってコードを生成せよ。

---

## 1. 拡張アーキテクチャの原則（完全プラグイン化：変更対象は1ファイルのみ）

本エンジンは完全自己完結型プラグインアーキテクチャ（`SPEC: TemplateSpec` と描画関数の動的ディスカバリ）を採用している。
新規テンプレート（例：`CORE-90`）を追加する際、**以下の1ファイルを作成して配置するだけでよい（`models.py`, `controller.py`, `test_engine.py` などの既存コードの編集は一切不要）**。

1. **`src/slide_engine/render_core_XX.py`（新規作成のみ）:**
   - Pydantic ペイロードクラス `CoreXXPayload` の定義。
   - 分割レイアウト時の境界防御メソッド `validate_split_density()` の実装（要素数上限の制限）。
   - 最大容量ストレステスト用サンプルデータ（`_SAMPLE_COREXX` および `_SAMPLE_COREXX_SPLIT`）の定義。
   - `SPEC = TemplateSpec(...)` によるメタデータ・スキーマ・サンプルの公開。
   - 描画関数 `render_core_xx(slide, payload, box)` の実装。

> [!TIP]
> テンプレートの削除も同様に、対象の `src/slide_engine/render_core_XX.py` を1ファイル削除するだけで完了する。既存コードやテストを修正する必要は一切ない。

---

## 2. 静的解析・テスト通過のための絶対制約 (Governance as Code)

生成したコードは `ruff check src tests`、`mypy src tests`、`pytest tests/` によって機械的に検査される。以下の制約に1つでも違反すると CI が失敗する。

1. **禁止APIの直接インポート禁止（Ruff `TID251`）:**
   - ❌ `from pptx.dml.color import RGBColor` は**インポート禁止**。色は必ず `SlideMaster.COLORS["primary"]` 等の辞書から取得せよ。
   - ❌ `from pptx.util import Pt` は**インポート禁止**。フォントサイズや線幅には必ず `FontEngine.pt(12.0)` または `FontEngine.calculate_font_size(...)` を使用せよ。
   - ⭕ `from pptx.util import Inches, Length` および `from pptx.enum.text import PP_ALIGN` のインポートは許可されている。
2. **座標型と二重スケーリングの厳禁（最重要バグ防止）:**
   - ❌ `Inches(box.left + Inches(0.200))` は**重大バグ**。`box.left` はすでに EMU 単位の `Length` 型であるため、加算結果を `Inches()` に渡すと 914,400 倍に巨大化してスライド外へ飛び出す。
   - ⭕ EMU 同士の加算・減算・乗除算の結果は、必ず **`Length(int(...))`** でラップして `BoundingBox` に渡せ。
3. **関数名と3引数シグネチャの完全一致（Pytest 検査対象）:**
   - ファイル名は `src/slide_engine/render_core_{小文字ID}.py`（例：`render_core_90.py`）。
   - エントリーポイント関数名は `render_core_{小文字ID}`（例：`render_core_90`）。
   - 引数名・順序・型アノテーションは **`(slide: Slide, payload: CoreXXPayload, box: BoundingBox) -> None`** に完全一致させよ。引数名を `s` や `b` に変えると自動検査で即座にテストが落ちる。
4. **`BoundingBox` 相対座標計算の徹底（ハードコード禁止）:**
   - スライドのヘッダー、リード文（`lead_text`）、および左右分割（`split_horizontal`）は上位の `controller.py` が処理し、描画可能な残り領域が `box: BoundingBox` として渡される。
   - 図形やテキストを配置する際は、**必ず `box.left`, `box.top`, `box.width`, `box.height` を起点とした相対座標**で計算せよ。
   - ローカルの余白やギャップ値はマジックナンバーとせず、ファイル先頭に `Inches(...)` の定数として定義せよ。
5. **McCabe 循環的複雑度 10以下（Ruff `C901`）:**
   - 1つの関数の循環的複雑度は最大 **10**（`max-complexity = 10`）に制限されている。
   - 不要な深いネストを避け、可読性の高い関数構造を保て。
6. **物理レイアウト不変量（Physical Invariants）の厳守:**
   - **可読性フォント下限:** すべてのテキストは `min_pt=8.0` を下限とし、`FontEngine.calculate_font_size` で動的縮小せよ。
   - **コンテンツシェイプの最小サイズ:** テキスト、表、またはチャートを持つ図形は `width >= Inches(0.35)` かつ `height >= Inches(0.20)` でなければならない。
   - **非コンテンツ補助図形:** コネクタ線や区切り線などの補助図形は、正の寸法（`width > 0, height > 0`）でありスライド内に収まっていれば良い。
   - **テキスト衝突の禁止:** テキストを持つシェイプ同士が物理的に重なってはならない。余白とゾーン（Y座標帯）を明確に分離せよ。
7. **厳格な型アノテーション（Mypy `--strict`）:**
   - ファイル先頭に `from __future__ import annotations` を記述し、すべての関数（プライベート関数含む）に完全な型ヒントを付与せよ。

---

## 3. `core_base.py` 公開APIチートシート

### 3.1 標準インポートボイラープレート
```python
from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length
from pydantic import Field

from slide_engine.core_base import (
    BoundingBox,
    FontEngine,
    ParagraphSpec,
    ShapeBuilder,
    SlideMaster,
    clamp_length,
)
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide
```

### 3.2 `BoundingBox`（幾何領域計算・自動レイアウト）
手動計算の代わりに、安全なインセットや自動グリッド分割の高水準APIを活用せよ。
```python
# 1. 余白インセット（パディングの自動安全クランプ・比率上限付き）
inner_box = card_box.inset(
    pad_x=Inches(0.200),
    pad_y=Inches(0.160),
    max_ratio_x=0.08,  # 幅に対するパディング比率上限（狭小ペイン時の潰れ防止）
    max_ratio_y=0.08,  # 高さに対するパディング比率上限
)

# 2. アダプティブグリッド（要素数と幅に応じた動的列数・行数決定）
# 幅が min_col_width を下回ると自動で列落ち（狭小分割レイアウト時は1列化）
grid_cells = box.adaptive_grid(
    num_items=len(payload.cards),
    min_col_width=Inches(3.200),
    gap_x=Inches(0.240),
    gap_y=Inches(0.200),
    balance_rows=True,  # 行ごとの要素数を均等配分（例: 4要素なら2x2）
)

# 3. 高度・幅の比率スケーリング
top_h = box.scale_height(0.40)
bottom_h = box.scale_height(0.60)

# 4. 参照可能なプロパティ:
# box.left, box.top, box.width, box.height
# box.right (left + width), box.bottom (top + height)
```

### 3.3 `SlideMaster`（カラーパレット・フォント定数）
* **フォント名定数:**
  * `SlideMaster.FONT_FAMILY_PRIMARY`: `"Meiryo UI"`（日本語見出し・本文の標準）
  * `SlideMaster.FONT_FAMILY_SECONDARY`: `"Arial"`（英数字・大型KPI数値用）
* **カラー辞書 (`SlideMaster.COLORS[key]`):**
  * `"primary"`: ディープネイビー（メイン見出し、強調カード背景）
  * `"secondary"`: スチールブルー（サブ見出し、バッジ番号、単位）
  * `"accent"`: バーントオレンジ（ハイライト枠線、ポジティブ強調文字）
  * `"text_dark"`: オフブラック（標準の本文テキスト）
  * `"text_light"`: オフホワイト（白抜き文字）
  * `"text_muted"`: グレー（注釈、キャプション、非強調ラベル）
  * `"bg_light"`: ライトクールグレー（標準カードの塗りつぶし背景）
  * `"bg_white"`: ピュアホワイト（白背景カード）
  * `"border"`: ボーダーグレー（標準カードの枠線）
  * `"fallback_bg"`: 警告用背景色
  * `"fallback_text"`: 警告用テキスト色

### 3.4 `ShapeBuilder` と `ParagraphSpec`（推奨：高水準描画API）
複数段落のフォントサイズ動的縮小・テキストボックス生成は、`ParagraphSpec` の宣言的スタックで一括描画せよ。
```python
# 1. カード図形の追加 (位置引数で slide, box を渡す)
ShapeBuilder.add_card(
    slide,
    card_box,
    bg_color=SlideMaster.COLORS["bg_light"],   # None で背景透明
    border_color=SlideMaster.COLORS["border"], # None で枠線なし
    border_width_pt=1.0,
)

# 2. 宣言的段落スタックの追加 (weight で高さを垂直配分し、自動フォント調整)
ShapeBuilder.add_styled_paragraph_stack(
    slide,
    inner_box,
    [
        ParagraphSpec(
            text="見出しタイトル",
            weight=0.40,             # 垂直方向の高さ割当比率
            max_pt=13.0,             # 最大フォントサイズ
            min_pt=8.0,              # 最小フォントサイズ下限（可読性不変量）
            color_key="primary",     # SlideMaster.COLORS のキー
            bold=True,
            space_after_pt=4.0,      # 段落後余白
        ),
        ParagraphSpec(
            text="詳細説明文テキスト。動的に安全縮小されテキストボックス内に収容。",
            weight=0.55,
            max_pt=10.0,
            min_pt=8.0,
            color_key="text_dark",
        ),
    ],
)

# 3. 単一テキストボックスの即時描画（バッジ、キャプション、単独ラベル等）
ShapeBuilder.add_styled_textbox(
    slide,
    badge_box,
    "[STEP 01]",
    max_pt=10.5,
    min_pt=8.0,
    color_key="secondary",
    bold=True,
    alignment=PP_ALIGN.CENTER,
)
```

### 3.5 低水準API・特殊装飾（KPI等の混在Run用）
単一の段落内でフォント・色・サイズが混在する場合（例: KPIの巨大数値＋小さな単位）のみ、低水準APIを使用せよ。
```python
# テキストボックスを直接追加
tb = ShapeBuilder.add_textbox(slide, text_box, margin_zero=True)
tf = tb.text_frame

# 手動での動的フォント計算
fitted_size = FontEngine.calculate_font_size(
    text="1,250",
    box_width=text_box.width,
    box_height=text_box.height,
    max_pt=36.0,
    min_pt=12.0,
)
p = tf.paragraphs[0]
p.text = "1,250"
p.font.name = SlideMaster.FONT_FAMILY_SECONDARY
p.font.size = fitted_size
p.font.bold = True
p.font.color.rgb = SlideMaster.COLORS["primary"]

# 単位などの別 Run 追加
run_unit = p.add_run()
run_unit.text = " 億円"
run_unit.font.name = SlideMaster.FONT_FAMILY_PRIMARY
run_unit.font.size = FontEngine.pt(14.0)
run_unit.font.color.rgb = SlideMaster.COLORS["secondary"]
```

### 3.6 幾何ユーティリティ
```python
# 座標・寸法の境界クランプ（負値や過大寸法の防止）
safe_w = clamp_length(box.width - gap, min_val=Inches(0.500), max_val=Inches(5.000))
```

---

## 4. プラグイン自己完結仕様（データ契約と描画の一体化）

各 `render_core_XX.py` は、単独で以下のすべてを定義する自己完結モジュールとして実装する。

1. **基底クラス:** 必ず `BasePayload`（`extra="forbid", frozen=True` 設定済み）を継承せよ。
2. **文字列型:** `ShortText`（最大80文字）、`SafeText`（最大150文字）を使用せよ。
3. **コレクション制約:** `list[...]` には必ず `Field(min_length=..., max_length=...)` を指定せよ。
4. **分割密度検証:** 左右分割で幅が狭くなる場合に備え、`validate_split_density(self, pane_name: str, template_id: str) -> None` をオーバーライドして要素数制限を課せ。
5. **サンプルデータの必須定義:**
   - `_SAMPLE_COREXX`: 単独ストレステスト用。**必ずコレクションの上限（`max_length`）いっぱいの要素数**で定義せよ。
   - `_SAMPLE_COREXX_SPLIT`: 分割ストレステスト用。狭小ペイン（幅30%）に収まる安全な要素数で定義せよ。
   - `_INVALID_SAMPLES_COREXX`: 境界防御テスト用。**最低1件以上**、スキーマ制約違反（文字数超過・要素数下限違反・型不一致など）を起こすペイロード辞書のリストを定義せよ。
6. **`SPEC: TemplateSpec` のトップレベル公開:**
   - モジュール内に `SPEC = TemplateSpec(...)` を定義し、メタデータ、スキーマ、サンプルデータ（`invalid_sample_payloads` 含む）を宣言せよ。エンジン起動時に自動検出・登録される。

---

## 5. 完全な実装リファレンス例（`CORE-90` を追加する場合）

`src/slide_engine/render_core_90.py` という**1ファイルを作成して置くだけ**で、エンジンへの登録、ディスパッチ、テストスイートへの自動追加がすべて完了する。

```python
"""Renderer for Grid Cards slide (CORE-90).

Renders 2 to 6 structured cards in an adaptive grid layout.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.util import Inches
from pydantic import Field

from slide_engine.core_base import BoundingBox, ParagraphSpec, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class GridCardItem(BasePayload):
    """Single item for Grid Cards slide."""

    badge: ShortText | None = None
    title: ShortText
    description: SafeText
    is_highlighted: bool = False


class Core90Payload(BasePayload):
    """Payload for Grid Cards slide (CORE-90)."""

    cards: list[GridCardItem] = Field(min_length=2, max_length=6)

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Enforce card count limit for split layout."""
        if len(self.cards) > 4:
            raise ValueError(
                f"{pane_name} grid cards '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.cards)} cards > 4 cards maximum."
            )


# サンプルデータ (単体用は上限の6枚、分割用は分割上限の4枚)
_SAMPLE_CORE90: dict[str, Any] = {
    "cards": [
        {
            "badge": f"ITEM {i}",
            "title": f"戦略施策タイトル {i}",
            "description": "中長期的な競争優位性を確立するための施策詳細説明テキスト。",
            "is_highlighted": i == 1,
        }
        for i in range(1, 7)
    ]
}

_SAMPLE_CORE90_SPLIT: dict[str, Any] = {
    "cards": [
        {
            "badge": f"ITEM {i}",
            "title": f"戦略施策 {i}",
            "description": "狭小分割ペイン用の簡潔な施策説明。",
            "is_highlighted": i == 1,
        }
        for i in range(1, 5)
    ]
}

_INVALID_SAMPLES_CORE90: list[dict[str, Any]] = [
    {"cards": [{"title": "A", "description": "B"}]},  # min_length=2 違反
]

# テンプレート仕様の宣言（エンジンが自動検出）
SPEC = TemplateSpec(
    payload_type=Core90Payload,
    supports_split=True,
    description="Grid Cards",
    sample_payload=_SAMPLE_CORE90,
    split_sample_payload=_SAMPLE_CORE90_SPLIT,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE90,
    renders_header=True,
    layout_type="body",
)

# Local Geometry Constants
GRID_MIN_COL_WIDTH = Inches(3.200)
GAP_X = Inches(0.240)
GAP_Y = Inches(0.200)
CARD_PAD_X = Inches(0.200)
CARD_PAD_Y = Inches(0.160)


def _render_single_card(slide: Slide, item: GridCardItem, card_box: BoundingBox) -> None:
    """Render an individual card inside the given bounding box."""
    bg_col = SlideMaster.COLORS["bg_white"] if item.is_highlighted else SlideMaster.COLORS["bg_light"]
    border_col = SlideMaster.COLORS["accent"] if item.is_highlighted else SlideMaster.COLORS["border"]
    border_pt = 1.5 if item.is_highlighted else 1.0

    ShapeBuilder.add_card(
        slide,
        card_box,
        bg_color=bg_col,
        border_color=border_col,
        border_width_pt=border_pt,
    )

    inner_box = card_box.inset(CARD_PAD_X, CARD_PAD_Y, max_ratio_x=0.08, max_ratio_y=0.08)

    header_prefix = f"[{item.badge}] " if item.badge else ""
    full_title = f"{header_prefix}{item.title}"
    title_col = "accent" if item.is_highlighted else "primary"

    ShapeBuilder.add_styled_paragraph_stack(
        slide,
        inner_box,
        [
            ParagraphSpec(
                text=full_title,
                weight=0.40,
                max_pt=12.5,
                min_pt=8.0,
                color_key=title_col,
                bold=True,
                space_after_pt=4.0,
            ),
            ParagraphSpec(
                text=item.description,
                weight=0.55,
                max_pt=10.0,
                min_pt=8.0,
                color_key="text_dark",
            ),
        ],
    )


def render_core_90(slide: Slide, payload: Core90Payload, box: BoundingBox) -> None:
    """Render grid cards dynamically within the provided BoundingBox.

    Args:
        slide: Target slide.
        payload: Core90Payload instance.
        box: Content BoundingBox.
    """
    grid = box.adaptive_grid(
        num_items=len(payload.cards),
        min_col_width=GRID_MIN_COL_WIDTH,
        gap_x=GAP_X,
        gap_y=GAP_Y,
        balance_rows=True,
    )
    for item, cell_box in zip(payload.cards, grid, strict=True):
        _render_single_card(slide, item, cell_box)
```

---

## 6. コード生成後の検証チェックリスト

生成AIは以下の項目をすべて満たしていることを確認してから完了報告せよ。
- [ ] `pptx.dml.color.RGBColor` や `pptx.util.Pt` をインポートしていないか？（`TID251` 遵守）
- [ ] `Inches(box.left + ...)` のような EMU 二重スケーリングを行っていないか？（`Length(int(...))` 必須）
- [ ] 関数名は `render_core_xx` で、引数名は正確に `(slide: Slide, payload: CoreXXPayload, box: BoundingBox) -> None` となっているか？
- [ ] McCabe 循環的複雑度が `10` 以下になっているか？
- [ ] テキスト描画は `ParagraphSpec` / `add_styled_paragraph_stack` 等の高水準APIを優先し、`min_pt=8.0` を下限に指定しているか？
- [ ] グリッド配置・余白計算は手動ループではなく `box.adaptive_grid()` や `box.inset()` を活用しているか？
- [ ] モジュール内に `SPEC = TemplateSpec(...)` を定義し、正しくエクスポートしているか？
- [ ] サンプルデータ `_SAMPLE_COREXX` にスキーマ上限（`max_length`）要素数を投入しているか？
- [ ] 境界防御テスト用の `_INVALID_SAMPLES_COREXX` を最低1件定義し、`SPEC` に渡したか？
- [ ] 分割対応テンプレートの場合、`validate_split_density()` と `_SAMPLE_COREXX_SPLIT` を定義したか？
