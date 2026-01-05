#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import requests
from azure.identity import ClientSecretCredential
from datetime import datetime


# =========================================
# SharePoint 設定（環境変数から取得）
# =========================================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SITE_HOST = "ncnj.sharepoint.com"
SITE_PATH = "/sites/staff_sharedfolders"


# =========================================
# 年度判定
# =========================================
def guess_years():
    today = datetime.now()

    # 1〜3月に実行されたら → current = 前年, next = 当年
    if today.month <= 3:
        current_year = today.year - 1
        next_year = today.year
    else:
        current_year = today.year
        next_year = today.year + 1

    return current_year, next_year


# =========================================
# SharePoint URL を構築
# =========================================
def build_sharepoint_url(year, term):
    """
    term: 'spring' or 'fall'
    """
    jp_year = f"{year}(R{year-2018})年度時間割"

    if term == "spring":
        filename = f"【{year}・04～09月 前期】全学年時間割.xlsx"
    else:
        filename = f"【{year}・10～03月 後期】全学年時間割.xlsx"

    # sharepoint の :x:/ 形式を利用
    return f"https://{SITE_HOST}/:x:/s/staff_sharedfolders/{jp_year}/{filename}"


# =========================================
# SharePoint authentication
# =========================================
def get_token():
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    token = credential.get_token("https://graph.microsoft.com/.default")
    return token.token


# =========================================
# ダウンロード
# =========================================
def download_file(url, save_path, token):
    headers = {"Authorization": f"Bearer {token}"}
    print(f"→ Downloading: {url}")

    resp = requests.get(url, headers=headers)

    if resp.status_code != 200:
        print(f"  ✗ Failed: {resp.status_code}")
        return False

    with open(save_path, "wb") as f:
        f.write(resp.content)

    print(f"  ✓ Saved to {save_path}")
    return True


# =========================================
# メイン処理
# =========================================
def main():
    token = get_token()
    current_year, next_year = guess_years()

    print(f"◆ current_year = {current_year}")
    print(f"◆ next_year    = {next_year}")

    # 保存名は excel2json.py と完全一致させる
    FILES = [
        ("spring", current_year, "schedule_spring_CURRENT.xlsx"),
        ("fall",   current_year, "schedule_fall_CURRENT.xlsx"),
        ("spring", next_year,   "schedule_spring_NEXT.xlsx"),
        ("fall",   next_year,   "schedule_fall_NEXT.xlsx"),
    ]

    os.makedirs(".", exist_ok=True)

    for term, year, save_name in FILES:
        url = build_sharepoint_url(year, term)
        download_file(url, save_name, token)


if __name__ == "__main__":
    main()
