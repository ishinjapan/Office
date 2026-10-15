import os
import sys
import time
import platform
import subprocess
from datetime import datetime, timedelta
import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

# --- 最小修正1: タイムゾーンを日本時間(JST)に固定 ---
os.environ['TZ'] = 'Asia/Tokyo'
if hasattr(time, 'tzset'):
    time.tzset()

# --- Chromiumブラウザの自動インストール ---
def ensure_playwright_installed():
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    except Exception as e:
        print(f"Playwright browser check warning: {e}")

ensure_playwright_installed()

# --- 認証情報の読み込み ---
load_dotenv()
AIRREGI_ID = os.getenv("AIRREGI_ID")
AIRREGI_PASS = os.getenv("AIRREGI_PASS")

if not AIRREGI_ID or not AIRREGI_PASS:
    try:
        import streamlit as st
        if hasattr(st, "secrets"):
            AIRREGI_ID = AIRREGI_ID or st.secrets.get("AIRREGI_ID")
            AIRREGI_PASS = AIRREGI_PASS or st.secrets.get("AIRREGI_PASS")
    except Exception:
        pass

OUTPUT_DIR = os.path.abspath("./output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
USER_DATA_DIR = os.path.abspath("./user_data")

TARGET_STORES = [
    "三田おにぎり 本店",
    "おにぎり維心 恵比寿店",
    "おにぎり維心 五反田店",
    "おにぎり維心 中目黒店"
]

def auto_login(page):
    page.goto("https://airregi.jp/CLP/view/top/")
    page.wait_for_timeout(2000)
    
    if "login" in page.url or page.locator("text='AirID'").first.is_visible() or page.get_by_role("button", name="ログイン").is_visible():
        print("・ログイン処理を自動実行中...")
        
        if not AIRREGI_ID or not AIRREGI_PASS:
            raise ValueError("【認証エラー】AIRREGI_ID または AIRREGI_PASS が設定されていません。Streamlit Secretsをご確認ください。")

        if page.get_by_role("button", name="ログイン").is_visible() and "login" not in page.url:
            page.get_by_role("button", name="ログイン").click()
            page.wait_for_timeout(2000)

        id_input = None
        selectors = [
            "input[name='loginId']:visible",
            "input[name='account']:visible",
            "input[name='userId']:visible",
            "input[placeholder*='AirID']:visible",
            "input[placeholder*='メール']:visible",
            "input[type='text']:visible:not([name*='dummy'])",
            "input[type='email']:visible:not([name*='dummy'])"
        ]
        for sel in selectors:
            loc = page.locator(sel).first
            if loc.is_visible():
                id_input = loc
                break

        if not id_input:
            for inp in page.locator("input:visible").all():
                inp_type = inp.get_attribute("type") or "text"
                inp_name = inp.get_attribute("name") or ""
                if inp_type in ["text", "email"] and "dummy" not in inp_name.lower():
                    id_input = inp
                    break

        if id_input:
            id_input.fill(AIRREGI_ID.strip())
        else:
            raise ValueError("AirIDの入力欄（可視要素）が見つかりませんでした。")
        
        pass_input = page.locator("input[type='password']:visible, input[name='password']:visible").first
        if pass_input.is_visible():
            pass_input.fill(AIRREGI_PASS.strip())
        else:
            raise ValueError("パスワードの入力欄が見つかりませんでした。")

        try:
            checkbox = page.locator("input[type='checkbox']:visible").first
            if checkbox.is_visible() and not checkbox.is_checked():
                checkbox.check()
        except Exception:
            pass

        login_btn = page.locator("button[type='submit']:visible, input[type='submit']:visible, button:has-text('ログイン'):visible").first
        login_btn.click()
        
        print("・認証後の画面遷移を待機中...")
        for _ in range(25):
            page.wait_for_timeout(1000)

            if "choose-store" in page.url or page.locator("text='店舗を選択'").first.is_visible() or page.locator("text='利用する店舗を選択'").first.is_visible():
                print("・店舗選択画面（choose-store）を検知しました。対象店舗を選択して進みます...")
                page.wait_for_timeout(1500)
                
                clicked = False
                for s_name in TARGET_STORES:
                    loc = page.locator(f"text='{s_name}'").first
                    if loc.is_visible():
                        loc.click()
                        clicked = True
                        break
                
                if not clicked:
                    candidate = page.locator("a:visible, button:visible").filter(has_text="店").first
                    if candidate.is_visible():
                        candidate.click()
                    else
