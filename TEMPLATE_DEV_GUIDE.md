# Slide Engine - 新規テンプレート拡張開発リファレンス (For AI Coding Agents)

本ドキュメントは、生成AI（コーディングエージェント）が Slide Engine に新しいスライドテンプレート（`CORE-XX`）を追加実装する際の**絶対遵守仕様書**である。
既存の `core_base.py` や `controller.py` のソースコードを開いて確認する必要はない。本ドキュメント内の API 仕様と制約に100%従ってコードを生成せよ。

---

## 1. 拡張アーキテクチャの原則（変更対象は2ファイルのみ）

本エンジンはプラグイン型の動的ディスパッチ（`glob("render_core_*.py")`）を採用している。
新規テンプレート（例：`CORE-80`）を追加する際、**以下の2ファイル以外を編集してはならない（`controller.py` や `core_base.py` の改修は一切禁止）**。

1. **`src/slide_engine/models.py`（追記）:**
   Pydantic ペイロードクラス `CoreXXPayload` の定義と、`TEMPLATE_REGISTRY` への1行追加。
2. **`src/slide_engine/render_core_XX.py`（新規作成）:**
   描画関数 `render_core_xx(slide, payload, box)` の実装。

---

## 2. 静的解析・テスト通過のための絶対制約 (Governance as Code)

生成したコードは `ruff check src tests`、`mypy src tests`、`pytest tests/` によって機械的に検査される。以下の6つの制約に1つでも違反すると CI が失敗する。

1. **禁止APIの直接インポート禁止（Ruff `TID251`）:**
   - ❌ `from pptx.dml.color import RGBColor` は**インポート禁止**。色は必ず `SlideMaster.COLORS["primary"]` 等の辞書から取得せよ。
   - ❌ `from pptx.util import Pt` は**インポート禁止**。フォントサイズや線幅には必ず `FontEngine.pt(12.0)` を使用せよ。
   - ⭕ `from pptx.util import Inches` および `from pptx.enum.text import PP_ALIGN, MSO_ANCHOR` のインポートは許可されている。
2. **関数名と3引数シグネチャの完全一致（Pytest 検査対象）:**
   - ファイル名は `src/slide_engine/render_core_{小文字ID}.py`（例：`render_core_80.py`）。
   - エントリーポイント関数名は `render_core_{小文字ID}`（例：`render_core_80`）。
   - 引数名・順序・型アノテーションは **`(slide: Slide, payload: CoreXXPayload, box: BoundingBox) -> None`** に完全一致させよ。引数名を `s` や `b` に変えると `test_renderer_function_signatures` の自動検査で即座にテストが落ちる。
3. **`BoundingBox` 相対座標計算の徹底（ハードコード禁止）:**
   - スライドのヘッダー、リード文（`lead_text`）、および左右分割（`split_horizontal`）は上位の `controller.py` が処理し、描画可能な残り領域が `box: BoundingBox` として渡される。
   - したがって、図形やテキストを配置する際は、**必ず `box.left`, `box.top`, `box.width`, `box.height` を起点とした相対座標**で計算せよ。
   - ローカルの余白やギャップ値はマジックナンバーとせず、ファイル先頭に `Inches(...)` の定数として定義せよ。
4. **McCabe 循環的複雑度 6以下（Ruff `C901`）:**
   - 1つの関数の循環的複雑度は最大 **6**（`max-complexity = 6`）に制限されている。
   - ループ内で複数の条件分岐を行う場合は、必ず `_render_single_item(...)` のようなプライベート関数へ分割し、各関数をシンプルに保て。
5. **プライベートメンバへのアクセス禁止（Ruff `SLF001`）:**
   - 外部オブジェクトの `_` で始まる属性・メソッドにアクセスしてはならない。
6. **厳格な型アノテーション（Mypy `--strict`）:**
   - ファイル先頭に `from __future__ import annotations` を記述し、すべての関数（プライベート関数含む）に完全な型ヒントを付与せよ。

---

## 3. `core_base.py` 公開APIチートシート

`render_core_XX.py` から呼び出し可能な `core_base.py` の全インターフェースは以下の通りである。

### 3.1 標準インポートボイラープレート
```python
from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core80Payload

if TYPE_CHECKING:
    from pptx.slide import Slide
```

### 3.2 `BoundingBox`（矩形領域クラス）
```python
# インスタンス生成（すべて Inches 型を渡すこと）
inner_box = BoundingBox(
    left=Inches(box.left + Inches(0.200)),
    top=Inches(box.top + Inches(0.200)),
    width=Inches(box.width - Inches(0.400)),
    height=Inches(box.height - Inches(0.400)),
)
# 参照可能な属性:
# inner_box.left, inner_box.top, inner_box.width, inner_box.height
# inner_box.right (left + width), inner_box.bottom (top + height)

# 左右2分割ヘルパーが必要な場合:
left_box, right_box = SlideMaster.split_horizontal(box, split_ratio=0.5, gap=Inches(0.300))
```

### 3.3 `SlideMaster`（カラーパレット・フォント定数）
* **フォント名定数:**
  * `SlideMaster.FONT_FAMILY_PRIMARY`: `"Meiryo UI"`（日本語見出し・本文の標準）
  * `SlideMaster.FONT_FAMILY_SECONDARY`: `"Arial"`（英数字・大型KPI数値用）
* **カラー辞書 (`SlideMaster.COLORS[key]`):**
  （定義されていないキーを指定すると `KeyError` になるため、以下の11種のみを使用せよ）
  * `"primary"`: ディープネイビー（メイン見出し、強調カード背景）
  * `"secondary"`: スチールブルー（サブ見出し、バッジ番号、単位）
  * `"accent"`: バーントオレンジ（ハイライト枠線、ポジティブ強調文字）
  * `"text_dark"`: オフブラック（標準の本文テキスト）
  * `"text_light"`: オフホワイト（`"primary"` 背景上の白抜き文字）
  * `"text_muted"`: グレー（注釈、キャプション、非強調ラベル）
  * `"bg_light"`: ライトクールグレー（標準カードの塗りつぶし背景）
  * `"bg_white"`: ピュアホワイト（白背景カード）
  * `"border"`: ボーダーグレー（標準カードの枠線）
  * `"fallback_bg"`: ソフトアンバー（ハイライト行・警告背景）
  * `"fallback_text"`: アンバーダーク（ネガティブ指標・警告文字）

### 3.4 `FontEngine`（ポイント変換・動的サイズ計算）
```python
# 1. 固定フォントサイズまたは線幅の指定 (Ptの代替)
p.font.size = FontEngine.pt(13.0)

# 2. ボックスサイズに応じた自動縮小フォントサイズの算出
fitted_size = FontEngine.calculate_font_size(
    text=item.description,
    box_width=content_box.width,
    box_height=content_box.height,
    max_pt=12.0,
    min_pt=8.5,
)
p.font.size = fitted_size
```

### 3.5 `ShapeBuilder`（図形・テキストボックス生成）
```python
# 1. 角丸長方形カードの追加
shape = ShapeBuilder.add_card(
    slide=slide,
    box=card_box,
    bg_color=SlideMaster.COLORS["bg_light"],   # None で背景透明
    border_color=SlideMaster.COLORS["border"], # None で枠線なし
    border_width_pt=1.0,
)

# 2. テキストボックスの追加 (word_wrap=True, 余白ゼロ初期化済み)
tb = ShapeBuilder.add_textbox(slide=slide, box=text_box, margin_zero=True)
tf = tb.text_frame

# 1段落目 (必ず paragraphs[0] を使用せよ。最初から add_paragraph() すると先頭に空行が入る)
p_title = tf.paragraphs[0]
p_title.text = "タイトル"
p_title.alignment = PP_ALIGN.LEFT
p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
p_title.font.size = FontEngine.pt(14.0)
p_title.font.bold = True
p_title.font.color.rgb = SlideMaster.COLORS["primary"]
p_title.space_after = FontEngine.pt(6.0)

# 2段落目以降
p_body = tf.add_paragraph()
p_body.text = "本文テキスト"
p_body.font.name = SlideMaster.FONT_FAMILY_PRIMARY
p_body.font.size = FontEngine.pt(10.5)
p_body.font.color.rgb = SlideMaster.COLORS["text_dark"]
```

---

## 4. `models.py` の拡張仕様（データ契約層）

`src/slide_engine/models.py` にスキーマを追加する際は、以下の型とルールに従うこと。

1. **基底クラス:** 必ず `BasePayload`（`extra="forbid", frozen=True` 設定済み）を継承せよ。
2. **文字列型:** 生の `str` ではなく、必ず既存の制約付き型エイリアスを使用せよ。
   - `ShortText`: 最大120文字（見出し、ラベル、バッジ、短い値）
   - `SafeText`: 最大800文字（説明文、箇条書き項目、キャプション）
3. **コレクション制約:** `list[...]` には必ず `Field(min_length=..., max_length=...)` を指定し、レイアウト崩壊とメモリ枯渇を防げ。
4. **純粋性の維持:** `@model_validator` 内でファイル I/O やロギング（`logging`, `print`）を行ってはならない。

---

## 5. 完全な実装リファレンス例（`CORE-80` を追加する場合）

### 手順1: `src/slide_engine/models.py` への追記箇所

```python
class GridCardItem(BasePayload):
    """Single item for Grid Cards slide."""

    badge: ShortText | None = None
    title: ShortText
    description: SafeText
    is_highlighted: bool = False


class Core80Payload(BasePayload):
    """Payload for Grid Cards slide (CORE-80)."""

    cards: list[GridCardItem] = Field(min_length=2, max_length=6)


# TEMPLATE_REGISTRY 辞書へのエントリ追加:
TEMPLATE_REGISTRY: dict[str, TemplateSpec] = {
    # ... 既存の CORE-00 〜 CORE-70 ...
    "CORE-80": TemplateSpec(
        payload_type=Core80Payload,
        supports_split=True,
        description="Grid Cards",
    ),
}
```

### 手順2: `src/slide_engine/render_core_80.py` の新規作成コード全文

```python
"""Renderer for Grid Cards slide (CORE-80).

Renders 2 to 6 structured cards in a clean 2-column grid layout.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core80Payload, GridCardItem

if TYPE_CHECKING:
    from pptx.slide import Slide

# Local Geometry Constants (No Magic Numbers)
GRID_COLS = 2
GAP_X = Inches(0.300)
GAP_Y = Inches(0.250)
CARD_PAD = Inches(0.220)


def _render_single_card(slide: Slide, item: GridCardItem, card_box: BoundingBox) -> None:
    """Render an individual card inside the given bounding box."""
    bg_col = SlideMaster.COLORS["bg_white"] if item.is_highlighted else SlideMaster.COLORS["bg_light"]
    border_col = SlideMaster.COLORS["accent"] if item.is_highlighted else SlideMaster.COLORS["border"]
    border_pt = 1.5 if item.is_highlighted else 1.0

    ShapeBuilder.add_card(
        slide=slide,
        box=card_box,
        bg_color=bg_col,
        border_color=border_col,
        border_width_pt=border_pt,
    )

    inner_box = BoundingBox(
        left=Inches(card_box.left + CARD_PAD),
        top=Inches(card_box.top + CARD_PAD),
        width=Inches(card_box.width - (CARD_PAD * 2)),
        height=Inches(card_box.height - (CARD_PAD * 2)),
    )
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    p_title = tf.paragraphs[0]
    header_prefix = f"[{item.badge}] " if item.badge else ""
    p_title.text = f"{header_prefix}{item.title}"
    p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_title.font.size = FontEngine.pt(13.0)
    p_title.font.bold = True
    p_title.font.color.rgb = SlideMaster.COLORS["accent"] if item.is_highlighted else SlideMaster.COLORS["primary"]
    p_title.space_after = FontEngine.pt(6.0)

    p_desc = tf.add_paragraph()
    p_desc.text = item.description
    p_desc.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_desc.font.size = FontEngine.calculate_font_size(
        text=item.description,
        box_width=inner_box.width,
        box_height=Inches(inner_box.height - Inches(0.350)),
        max_pt=10.5,
        min_pt=8.5,
    )
    p_desc.font.color.rgb = SlideMaster.COLORS["text_dark"]


def render_core_80(slide: Slide, payload: Core80Payload, box: BoundingBox) -> None:
    """Render grid cards dynamically within the provided BoundingBox.

    Args:
        slide: Target slide.
        payload: Core80Payload instance.
        box: Content BoundingBox.
    """
    num_cards = len(payload.cards)
    num_rows = (num_cards + GRID_COLS - 1) // GRID_COLS

    card_width = Inches((box.width - (GAP_X * (GRID_COLS - 1))) / GRID_COLS)
    card_height = Inches((box.height - (GAP_Y * (num_rows - 1))) / num_rows)

    for idx, item in enumerate(payload.cards):
        row_idx = idx // GRID_COLS
        col_idx = idx % GRID_COLS

        card_left = Inches(box.left + col_idx * (card_width + GAP_X))
        card_top = Inches(box.top + row_idx * (card_height + GAP_Y))

        card_box = BoundingBox(card_left, card_top, card_width, card_height)
        _render_single_card(slide, item, card_box)
```

---

## 6. コード生成後の検証チェックリスト

生成AIは以下の項目をすべて満たしていることを確認してからコードを提示すること。
- [ ] `pptx.dml.color.RGBColor` や `pptx.util.Pt` をインポートしていないか？（`TID251` 遵守）
- [ ] 関数名は `render_core_xx` で、引数名は正確に `(slide: Slide, payload: CoreXXPayload, box: BoundingBox) -> None` となっているか？
- [ ] すべての座標計算が `box.left`, `box.top`, `box.width`, `box.height` を基準とした相対計算になっているか？
- [ ] 1つの関数に `if` や `for` を詰め込みすぎて McCabe 複雑度が `6` を超えていないか？（`C901` 遵守）
- [ ] `models.py` の `TEMPLATE_REGISTRY` への登録コードを含めているか？