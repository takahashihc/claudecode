"""
PowerPoint generator: creates a 7-slide weekly report from structured news JSON.
Usage: python3 pptx_generator.py --json <path_to_news.json> [--date YYYY-MM-DD]
"""
import argparse
import json
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── Color palette (matching the Excel) ───────────────────────────────────────
C_DARK_BLUE   = RGBColor(0x1F, 0x4E, 0x79)
C_MID_BLUE    = RGBColor(0x2E, 0x75, 0xB6)
C_LIGHT_BLUE  = RGBColor(0xBD, 0xD7, 0xEE)
C_ORANGE      = RGBColor(0xC5, 0x5A, 0x11)
C_PALE_ORANGE = RGBColor(0xFC, 0xE4, 0xD6)
C_GREEN       = RGBColor(0x37, 0x56, 0x23)
C_PALE_GREEN  = RGBColor(0xE2, 0xEF, 0xDA)
C_WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
C_RED         = RGBColor(0xC0, 0x00, 0x00)
C_YELLOW_BG   = RGBColor(0xFF, 0xFA, 0xCD)
C_GRAY        = RGBColor(0xF2, 0xF2, 0xF2)
C_BLACK       = RGBColor(0x1A, 0x1A, 0x1A)

URGENCY_COLORS = {
    "高": (RGBColor(0xFF, 0xD9, 0xD9), C_RED),
    "中": (C_YELLOW_BG, RGBColor(0x7D, 0x66, 0x08)),
    "低": (C_PALE_GREEN, C_GREEN),
}

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def _prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def _blank(prs):
    blank_layout = prs.slide_layouts[6]  # completely blank
    return prs.slides.add_slide(blank_layout)


def _rect(slide, left, top, width, height, fill_color=None, line_color=None):
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.5)
    else:
        shape.line.fill.background()
    return shape


def _textbox(slide, left, top, width, height, text, font_size=12,
             bold=False, color=None, bg_color=None, align=PP_ALIGN.LEFT,
             wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = "Meiryo"
    if color:
        run.font.color.rgb = color
    if bg_color:
        txBox.fill.solid()
        txBox.fill.fore_color.rgb = bg_color
    return txBox


def _header_bar(slide, title: str, subtitle: str = ""):
    _rect(slide, 0, 0, SLIDE_W, Inches(1.1), C_DARK_BLUE)
    _textbox(slide, Inches(0.3), Inches(0.1), Inches(11), Inches(0.55),
             title, font_size=20, bold=True, color=C_WHITE)
    if subtitle:
        _textbox(slide, Inches(0.3), Inches(0.65), Inches(11), Inches(0.4),
                 subtitle, font_size=11, color=C_LIGHT_BLUE)


def _logo_placeholder(slide):
    box = _rect(slide, Inches(11.8), Inches(0.05), Inches(1.3), Inches(0.9),
                fill_color=C_MID_BLUE, line_color=C_WHITE)
    _textbox(slide, Inches(11.8), Inches(0.15), Inches(1.3), Inches(0.7),
             "LOGO", font_size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)


def _footer(slide, report_date: str):
    _rect(slide, 0, Inches(7.1), SLIDE_W, Inches(0.4), C_GRAY)
    _textbox(slide, Inches(0.3), Inches(7.12), Inches(8), Inches(0.3),
             f"高橋将史商会　食品包装業界 情勢レポート　{report_date}", font_size=8, color=C_DARK_BLUE)
    _textbox(slide, Inches(10), Inches(7.12), Inches(3), Inches(0.3),
             "本資料は社外秘です", font_size=8, color=C_ORANGE, align=PP_ALIGN.RIGHT)


# ── Slide 1: 表紙 ─────────────────────────────────────────────────────────────
def slide_cover(prs, news: dict, report_date: str):
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_DARK_BLUE)

    # Accent bar
    _rect(slide, 0, Inches(3.5), Inches(0.15), Inches(2.5), C_ORANGE)

    # Title
    _textbox(slide, Inches(0.5), Inches(1.2), Inches(11), Inches(0.9),
             "食品包装業界　情勢レポート", font_size=32, bold=True, color=C_WHITE,
             align=PP_ALIGN.LEFT)

    # Subtitle
    _textbox(slide, Inches(0.5), Inches(2.2), Inches(11), Inches(0.7),
             "ホルムズ海峡封鎖・中東情勢　商品群別影響分析", font_size=18,
             color=C_LIGHT_BLUE, align=PP_ALIGN.LEFT)

    # Date badge
    _rect(slide, Inches(0.5), Inches(3.2), Inches(3.5), Inches(0.55), C_ORANGE)
    _textbox(slide, Inches(0.5), Inches(3.2), Inches(3.5), Inches(0.55),
             f"　{report_date}　発行", font_size=16, bold=True,
             color=C_WHITE, align=PP_ALIGN.LEFT)

    # Period
    period = news.get("period", report_date)
    _textbox(slide, Inches(0.5), Inches(3.9), Inches(7), Inches(0.4),
             f"収集期間：{period}", font_size=12, color=C_LIGHT_BLUE)

    # Company name
    _textbox(slide, Inches(0.5), Inches(6.3), Inches(8), Inches(0.6),
             "高橋将史商会", font_size=22, bold=True, color=C_WHITE)

    # Logo placeholder
    _rect(slide, Inches(10.5), Inches(5.8), Inches(2.5), Inches(1.4), C_MID_BLUE)
    _textbox(slide, Inches(10.5), Inches(6.0), Inches(2.5), Inches(0.8),
             "LOGO", font_size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)


# ── Slide 2: エグゼクティブサマリー ─────────────────────────────────────────────
def slide_summary(prs, news: dict, report_date: str):
    slide = _blank(prs)
    _header_bar(slide, "エグゼクティブサマリー", f"収集期間：{news.get('period', report_date)}")
    _logo_placeholder(slide)
    _footer(slide, report_date)

    # Key figures box
    _rect(slide, Inches(0.3), Inches(1.2), Inches(12.7), Inches(1.2), C_PALE_ORANGE)
    figures = [
        ("原油（ブレント）", news.get("brent_crude", "—")),
        ("ナフサ", news.get("naphtha_price", "—")),
        ("円ドル為替", news.get("usd_jpy", "—")),
        ("ホルムズ状況", news.get("hormuz_status", "—")[:12]),
    ]
    for i, (label, val) in enumerate(figures):
        x = Inches(0.4 + i * 3.2)
        _textbox(slide, x, Inches(1.25), Inches(3.0), Inches(0.35),
                 label, font_size=9, color=C_ORANGE, bold=True)
        _textbox(slide, x, Inches(1.6), Inches(3.0), Inches(0.65),
                 val, font_size=14, bold=True, color=C_DARK_BLUE)

    # Summary points
    _textbox(slide, Inches(0.3), Inches(2.5), Inches(3.5), Inches(0.35),
             "■ 今週の重要ポイント", font_size=13, bold=True, color=C_DARK_BLUE)

    points = news.get("summary_points", [])
    for i, pt in enumerate(points[:6]):
        y = Inches(2.95 + i * 0.62)
        # bullet marker
        _rect(slide, Inches(0.3), y + Inches(0.08), Inches(0.08), Inches(0.3), C_ORANGE)
        _textbox(slide, Inches(0.5), y, Inches(12.3), Inches(0.55),
                 pt, font_size=11, color=C_BLACK)


# ── Slide 3: 中東・ホルムズ海峡 現状 ─────────────────────────────────────────
def slide_hormuz(prs, news: dict, report_date: str):
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, Inches(1.1), C_ORANGE)
    _textbox(slide, Inches(0.3), Inches(0.1), Inches(11), Inches(0.55),
             "中東・ホルムズ海峡　現状", font_size=20, bold=True, color=C_WHITE)
    _textbox(slide, Inches(0.3), Inches(0.65), Inches(11), Inches(0.4),
             f"As of {report_date}", font_size=11, color=C_PALE_ORANGE)
    _logo_placeholder(slide)
    _footer(slide, report_date)

    # Status badge
    status = news.get("hormuz_status", "情報収集中")
    _rect(slide, Inches(0.3), Inches(1.25), Inches(5.5), Inches(0.75), C_RED)
    _textbox(slide, Inches(0.3), Inches(1.25), Inches(5.5), Inches(0.75),
             f"⚠ {status}", font_size=15, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Key numbers
    kv_items = [
        ("原油（ブレント）", news.get("brent_crude", "—"), C_PALE_ORANGE),
        ("ナフサ価格（CIF Japan）", news.get("naphtha_price", "—"), C_PALE_ORANGE),
        ("円ドル為替", news.get("usd_jpy", "—"), C_LIGHT_BLUE),
    ]
    for i, (label, val, bg) in enumerate(kv_items):
        x = Inches(0.3 + i * 4.3)
        _rect(slide, x, Inches(2.15), Inches(4.0), Inches(1.1), bg)
        _textbox(slide, x + Inches(0.1), Inches(2.2), Inches(3.8), Inches(0.4),
                 label, font_size=10, bold=True, color=C_ORANGE)
        _textbox(slide, x + Inches(0.1), Inches(2.6), Inches(3.8), Inches(0.55),
                 val, font_size=18, bold=True, color=C_DARK_BLUE)

    # Impact note
    _textbox(slide, Inches(0.3), Inches(3.45), Inches(12.7), Inches(0.4),
             "■ 食品包装業界への主要影響", font_size=13, bold=True, color=C_DARK_BLUE)

    impact_lines = [
        "・石油系樹脂（PE/PP/PS/NBR）の原料ナフサ供給が最大24%減少。価格は2月末比+35〜45%で推移",
        "・ニトリル手袋の原料（ブタジエン）価格が+70%急騰。マレーシア主要メーカーが生産縮小",
        "・海上運賃（LR2タンカー）が倍増し、アジア向けナフサ輸送コストが大幅上昇",
        "・日本の原油中東依存度94%・ホルムズ経由9割のため、エネルギー集約型製品全般に波及",
    ]
    for i, line in enumerate(impact_lines):
        _textbox(slide, Inches(0.3), Inches(3.95 + i * 0.65), Inches(12.7), Inches(0.6),
                 line, font_size=10, color=C_BLACK)


# ── Slide 4: 商品群別 影響度マトリクス ───────────────────────────────────────
def slide_matrix(prs, news: dict, report_date: str):
    slide = _blank(prs)
    _header_bar(slide, "商品群別　影響度マトリクス", "全7商品群の緊急度・価格影響・供給リスク一覧")
    _logo_placeholder(slide)
    _footer(slide, report_date)

    products = news.get("products", [])

    # Table header
    cols = ["商品群", "緊急度", "石油\n依存度", "原材料価格\n影響", "供給\nリスク", "推奨在庫", "主要アクション"]
    col_widths = [Inches(1.8), Inches(0.8), Inches(0.9), Inches(2.2), Inches(1.5), Inches(1.2), Inches(4.5)]
    x_offsets = []
    x = Inches(0.15)
    for w in col_widths:
        x_offsets.append(x)
        x += w

    # Header row
    for i, (col, cw, cx) in enumerate(zip(cols, col_widths, x_offsets)):
        _rect(slide, cx, Inches(1.2), cw, Inches(0.5), C_DARK_BLUE)
        _textbox(slide, cx, Inches(1.2), cw, Inches(0.5),
                 col, font_size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    oil_dep_map = {
        "段ボール": "低", "発泡スチロール": "高", "食品トレー": "高",
        "ボードン袋": "高", "ストレッチフィルム": "高",
        "ニトリル手袋": "非常に高", "プラスチックコンテナー": "高",
    }
    supply_risk_map = {
        "段ボール": "低", "発泡スチロール": "高", "食品トレー": "高",
        "ボードン袋": "高", "ストレッチフィルム": "高",
        "ニトリル手袋": "非常に高", "プラスチックコンテナー": "高",
    }
    stock_map = {
        "段ボール": "1〜2ヶ月分", "発泡スチロール": "3ヶ月以上",
        "食品トレー": "3ヶ月以上", "ボードン袋": "3ヶ月以上",
        "ストレッチフィルム": "3ヶ月以上", "ニトリル手袋": "4ヶ月以上",
        "プラスチックコンテナー": "3ヶ月以上",
    }

    for row_i, product in enumerate(products[:7]):
        name = product.get("name", "")
        urgency = product.get("urgency", "中")
        urg_bg, urg_fg = URGENCY_COLORS.get(urgency, (C_GRAY, C_BLACK))

        row_bg = C_PALE_ORANGE if urgency == "高" else (C_YELLOW_BG if urgency == "中" else C_PALE_GREEN)

        y = Inches(1.75 + row_i * 0.67)

        row_vals = [
            (name, C_MID_BLUE, C_WHITE, True),
            (urgency, urg_bg, urg_fg, True),
            (oil_dep_map.get(name, "—"), row_bg, C_BLACK, False),
            (product.get("import_cost_impact", "—"), row_bg, C_BLACK, False),
            (supply_risk_map.get(name, "—"), row_bg, C_BLACK, False),
            (stock_map.get(name, "—"), row_bg, C_BLACK, False),
            (product.get("action", "—")[:40], row_bg, C_BLACK, False),
        ]

        for (val, bg, fg, bold_), cw, cx in zip(row_vals, col_widths, x_offsets):
            _rect(slide, cx, y, cw, Inches(0.62), bg)
            _textbox(slide, cx + Inches(0.05), y + Inches(0.05), cw - Inches(0.1), Inches(0.56),
                     val, font_size=9, bold=bold_, color=fg, align=PP_ALIGN.CENTER)


# ── Slide 5: 緊急度「高」商品 詳細分析 ─────────────────────────────────────────
def slide_urgent(prs, news: dict, report_date: str):
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, Inches(1.1), C_ORANGE)
    _textbox(slide, Inches(0.3), Inches(0.1), Inches(11), Inches(0.55),
             "緊急度「高」商品　詳細分析", font_size=20, bold=True, color=C_WHITE)
    _textbox(slide, Inches(0.3), Inches(0.65), Inches(11), Inches(0.4),
             "即時対応が必要な商品群", font_size=11, color=C_PALE_ORANGE)
    _logo_placeholder(slide)
    _footer(slide, report_date)

    urgent_products = [p for p in news.get("products", []) if p.get("urgency") == "高"]
    if not urgent_products:
        _textbox(slide, Inches(0.3), Inches(2.0), Inches(12.7), Inches(0.5),
                 "緊急度「高」の商品群はありません", font_size=14, color=C_GREEN)
        return

    card_w = Inches(12.7 / min(len(urgent_products), 3))
    for i, product in enumerate(urgent_products[:3]):
        col = i % 3
        row = i // 3
        x = Inches(0.3) + col * card_w
        y = Inches(1.2) + row * Inches(2.9)

        # Card background
        _rect(slide, x, y, card_w - Inches(0.1), Inches(2.75), C_PALE_ORANGE)

        # Card header
        _rect(slide, x, y, card_w - Inches(0.1), Inches(0.45), C_RED)
        _textbox(slide, x, y, card_w - Inches(0.1), Inches(0.45),
                 f"⚠ {product.get('name', '')}", font_size=13, bold=True,
                 color=C_WHITE, align=PP_ALIGN.CENTER)

        fields = [
            ("原材料価格", product.get("raw_material_price_trend", "—")),
            ("輸入コスト影響", product.get("import_cost_impact", "—")),
            ("ホルムズ直接影響", product.get("hormuz_direct_impact", "—")),
            ("代替調達先", product.get("alternative_procurement", "—")),
            ("推奨アクション", product.get("action", "—")),
            ("情報源", product.get("source", "—")),
        ]
        for j, (label, val) in enumerate(fields):
            fy = y + Inches(0.5 + j * 0.38)
            _textbox(slide, x + Inches(0.1), fy, card_w * 0.38, Inches(0.36),
                     label, font_size=8, bold=True, color=C_ORANGE)
            _textbox(slide, x + card_w * 0.4, fy, card_w * 0.55, Inches(0.36),
                     str(val)[:35], font_size=8, color=C_BLACK)


# ── Slide 6: 推奨アクション ──────────────────────────────────────────────────
def slide_actions(prs, news: dict, report_date: str):
    slide = _blank(prs)
    _header_bar(slide, "推奨アクション（優先順位付き）", "経営判断・社員対応・取引先対応")
    _logo_placeholder(slide)
    _footer(slide, report_date)

    products = news.get("products", [])
    # Sort by urgency: 高 first, then 中, then 低
    order = {"高": 0, "中": 1, "低": 2}
    sorted_products = sorted(products, key=lambda p: order.get(p.get("urgency", "低"), 2))

    priority_labels = {0: "【最優先】", 1: "【優先】", 2: "【通常】"}
    urgency_count = 0

    y = Inches(1.25)
    for product in sorted_products:
        urgency = product.get("urgency", "低")
        action = product.get("action", "")
        name = product.get("name", "")
        if not action:
            continue

        bg, fg = URGENCY_COLORS.get(urgency, (C_GRAY, C_BLACK))
        pri_label = priority_labels.get(order.get(urgency, 2), "")

        _rect(slide, Inches(0.3), y, Inches(1.2), Inches(0.52), bg)
        _textbox(slide, Inches(0.3), y, Inches(1.2), Inches(0.52),
                 pri_label, font_size=8, bold=True, color=fg, align=PP_ALIGN.CENTER)

        _rect(slide, Inches(1.55), y, Inches(1.6), Inches(0.52), C_MID_BLUE)
        _textbox(slide, Inches(1.55), y, Inches(1.6), Inches(0.52),
                 name, font_size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

        _textbox(slide, Inches(3.25), y, Inches(9.7), Inches(0.52),
                 action, font_size=10, color=C_BLACK)

        y += Inches(0.58)
        urgency_count += 1
        if y > Inches(6.8):
            break

    # Bottom note
    _rect(slide, Inches(0.3), Inches(6.55), Inches(12.7), Inches(0.45), C_LIGHT_BLUE)
    _textbox(slide, Inches(0.3), Inches(6.55), Inches(12.7), Inches(0.45),
             "※ 上記アクションは収集時点の情報に基づきます。最新情報を確認の上、経営判断を行ってください。",
             font_size=9, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)


# ── Slide 7: 次回レポート予定 ─────────────────────────────────────────────────
def slide_next(prs, news: dict, report_date: str):
    from datetime import date, timedelta
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_DARK_BLUE)
    _logo_placeholder(slide)

    _textbox(slide, Inches(1.0), Inches(1.2), Inches(11), Inches(0.8),
             "次回レポート予定", font_size=28, bold=True, color=C_WHITE)

    # Calculate next 3 report dates (Mon/Wed/Fri pattern)
    today = date.fromisoformat(report_date)
    next_dates = []
    d = today + timedelta(days=1)
    weekday_targets = {0, 2, 4}  # Mon, Wed, Fri
    while len(next_dates) < 3:
        if d.weekday() in weekday_targets:
            next_dates.append(d)
        d += timedelta(days=1)

    weekday_jp = ["月", "火", "水", "木", "金", "土", "日"]
    for i, nd in enumerate(next_dates):
        wd = weekday_jp[nd.weekday()]
        _rect(slide, Inches(1.0 + i * 4.1), Inches(2.2), Inches(3.7), Inches(1.2), C_MID_BLUE)
        _textbox(slide, Inches(1.0 + i * 4.1), Inches(2.3), Inches(3.7), Inches(1.0),
                 f"第{i+1}回\n{nd.strftime('%Y年%m月%d日')}（{wd}）",
                 font_size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    _textbox(slide, Inches(1.0), Inches(3.7), Inches(11), Inches(0.5),
             "収集コマンド：Claude Code で  /週次レポート  と入力してください",
             font_size=13, color=C_LIGHT_BLUE, align=PP_ALIGN.CENTER)

    _rect(slide, Inches(1.0), Inches(4.5), Inches(11), Inches(1.5), C_MID_BLUE)
    _textbox(slide, Inches(1.1), Inches(4.55), Inches(10.8), Inches(1.35),
             "本レポートに関するお問い合わせ\n"
             "高橋将史商会　〒XXXXX　\n"
             "TEL: 0XX-XXXX-XXXX　Email: info@takahashi-hc.co.jp",
             font_size=12, color=C_WHITE)

    _textbox(slide, Inches(1.0), Inches(6.3), Inches(11), Inches(0.5),
             "高橋将史商会　食品業界向け包装資材　卸売業",
             font_size=14, bold=True, color=C_LIGHT_BLUE, align=PP_ALIGN.CENTER)


# ── Main entry ────────────────────────────────────────────────────────────────
def generate_pptx(news: dict, report_date: str) -> Path:
    prs = _prs()

    slide_cover(prs, news, report_date)
    slide_summary(prs, news, report_date)
    slide_hormuz(prs, news, report_date)
    slide_matrix(prs, news, report_date)
    slide_urgent(prs, news, report_date)
    slide_actions(prs, news, report_date)
    slide_next(prs, news, report_date)

    report_dir = Path(__file__).parent / "reports" / report_date
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / f"高橋将史商会_週次レポート_{report_date}.pptx"
    prs.save(out_path)
    print(f"✅ PowerPoint 保存: {out_path}")
    return out_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="Path to news JSON file")
    parser.add_argument("--date", default=str(date.today()))
    args = parser.parse_args()

    with open(args.json, encoding="utf-8") as f:
        news = json.load(f)

    generate_pptx(news, args.date)
