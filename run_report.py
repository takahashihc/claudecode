"""
Main orchestrator: accepts news JSON via --json flag or stdin, then runs
Excel update + PowerPoint generation in sequence.

Usage:
  python3 run_report.py --json news_2026-04-21.json
  python3 run_report.py --json news_2026-04-21.json --date 2026-04-21
"""
import argparse
import json
import sys
from datetime import date
from pathlib import Path

from excel_updater import update_excel
from pptx_generator import generate_pptx


def main():
    parser = argparse.ArgumentParser(description="週次ニュースレポート生成")
    parser.add_argument("--json", required=True, help="ニュースJSONファイルのパス")
    parser.add_argument("--date", default=str(date.today()), help="レポート日付 YYYY-MM-DD")
    args = parser.parse_args()

    json_path = Path(args.json)
    if not json_path.exists():
        print(f"❌ JSONファイルが見つかりません: {json_path}", file=sys.stderr)
        sys.exit(1)

    with open(json_path, encoding="utf-8") as f:
        news = json.load(f)

    report_date = args.date
    print(f"\n📅 レポート日付: {report_date}")
    print(f"📦 商品群数: {len(news.get('products', []))}")
    print("-" * 50)

    xlsx_path = update_excel(news, report_date)
    pptx_path = generate_pptx(news, report_date)

    print("-" * 50)
    print(f"\n✅ 生成完了:")
    print(f"   Excel:      {xlsx_path}")
    print(f"   PowerPoint: {pptx_path}")
    print(f"\n📂 保存先フォルダ: reports/{report_date}/")


if __name__ == "__main__":
    main()
