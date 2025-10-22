# magapoke_weekly.py
import requests
from bs4 import BeautifulSoup
import pandas as pd
import datetime
from openpyxl import load_workbook

URL = "https://pocket.shonenmagazine.com/ranking/30"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; magapoke-scraper/1.1)"
}

def fetch_titles(url=URL):
    r = requests.get(url, headers=HEADERS, timeout=20)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    # タイトル要素を上から順に抽出（=ランキング順）
    h3s = soup.select("h3.c-ranking-item__ttl")
    titles = [h.get_text(strip=True) for h in h3s if h.get_text(strip=True)]
    return titles

def save_to_excel(titles, file_path="weekly_ranking.xlsx"):
    # 日付をシート名に
    sheet_name = datetime.datetime.now().strftime("magapoke_%Y-%m-%d")

    # DataFrame作成
    df = pd.DataFrame([{"rank": i + 1, "title": t} for i, t in enumerate(titles)])

    try:
        # 既存ファイルがある場合は追記
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"追加保存: {file_path} → シート名 '{sheet_name}'")

    except FileNotFoundError:
        # ファイルがない場合は新規作成
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"新規作成: {file_path}")

def main():
    titles = fetch_titles()
    save_to_excel(titles)

if __name__ == "__main__":
    main()
