"""
Excel updater: writes structured news JSON into the master Excel and saves a dated copy.
Usage: python3 excel_updater.py --json <path_to_news.json> [--date YYYY-MM-DD]
"""
import argparse
import json
import os
import shutil
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

MASTER_XLSX = Path(__file__).parent / "高橋将史商会_商品群別ニュース収集フォーマット.xlsx"

# Map JSON product name → Sheet1 row number (rows 4-10)
PRODUCT_ROW = {
    "段ボール": 4,
    "発泡スチロール": 5,
    "食品トレー": 6,
    "ボードン袋": 7,
    "ストレッチフィルム": 8,
    "ニトリル手袋": 9,
    "プラスチックコンテナー": 10,
}

# Sheet1: JSON key → column index (1-based)
COL_MAP = {
    "raw_material_price_trend":  3,
    "domestic_price_trend":      4,
    "procurement_source":        5,
    "alternative_procurement":   6,
    "hormuz_direct_impact":      7,
    "middle_east_impact":        8,
    "crude_naphtha":             9,
    "sea_freight":               10,
    "domestic_logistics":        11,
    "lead_time":                 12,
    "exchange_rate_impact":      13,
    "import_cost_impact":        14,
    "market_demand":             15,
    "regulations":               16,
    "competitors":               17,
    "customer_impact":           18,
    "inventory_status":          19,
    "action":                    20,
    "urgency":                   21,
    "source":                    22,
    "collection_date":           23,
    "notes":                     24,
}

# Sheet2: JSON key → column index (1-based)
S2_COL_MAP = {
    "urgency":          21,   # 緊急仕入必要性 proxy
    "risk_detail":      11,
    "action":           12,
}

URGENCY_COLORS = {
    "高": ("FFD9D9", "C00000"),
    "中": ("FFFACD", "7D6608"),
    "低": ("E2EFDA", "375623"),
}


def _cell_font_fill(cell, urgency: str):
    bg, fg = URGENCY_COLORS.get(urgency, ("FFFFFF", "1A1A1A"))
    cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    cell.font = Font(name="Meiryo", bold=True, color=fg, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def update_excel(news: dict, report_date: str) -> Path:
    wb = openpyxl.load_workbook(MASTER_XLSX)
    ws1 = wb["ニュース収集フォーマット"]
    ws2 = wb["ホルムズ影響度評価"]

    # Write period to subtitle row (row 2)
    period = news.get("period", report_date)
    ws1["A2"] = f"収集期間：{period}　　担当者名：Claude Code週次レポート　　注目テーマ：中東・ホルムズ海峡情勢"

    for product in news.get("products", []):
        name = product.get("name", "")
        row = PRODUCT_ROW.get(name)
        if row is None:
            continue

        # Populate Sheet1
        for key, col in COL_MAP.items():
            if key == "collection_date":
                ws1.cell(row=row, column=col).value = report_date
            else:
                val = product.get(key, "")
                cell = ws1.cell(row=row, column=col)
                cell.value = val
                if key == "urgency" and val in URGENCY_COLORS:
                    _cell_font_fill(cell, val)

        # Populate Sheet2 (rows 3-9 map to rows 4-10 in sheet1 ordering)
        s2_row = row - 1   # Sheet2 data starts row 3
        ws2.cell(row=s2_row, column=10).value = product.get("urgency", "")
        ws2.cell(row=s2_row, column=11).value = product.get("notes", "") or product.get("hormuz_direct_impact", "")
        ws2.cell(row=s2_row, column=12).value = product.get("action", "")
        urg_cell = ws2.cell(row=s2_row, column=8)
        urg_cell.value = product.get("urgency", "")
        if product.get("urgency") in URGENCY_COLORS:
            _cell_font_fill(urg_cell, product["urgency"])

    wb.save(MASTER_XLSX)

    # Save dated copy
    report_dir = Path(__file__).parent / "reports" / report_date
    report_dir.mkdir(parents=True, exist_ok=True)
    dated_path = report_dir / f"高橋将史商会_ニュース収集_{report_date}.xlsx"
    shutil.copy2(MASTER_XLSX, dated_path)

    print(f"✅ Excel 保存: {dated_path}")
    return dated_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="Path to news JSON file")
    parser.add_argument("--date", default=str(date.today()))
    args = parser.parse_args()

    with open(args.json, encoding="utf-8") as f:
        news = json.load(f)

    update_excel(news, args.date)
