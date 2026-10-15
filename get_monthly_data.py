import os
import sys
import time
import platform
import subprocess
import re
from datetime import datetime, timedelta
import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

# --- 1. Chromiumブラウザの自動インストール ---
def ensure_playwright_installed():
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    except Exception as e:
        print(f"Playwright browser check warning: {e}")

ensure_playwright_installed()

# --- 2. 認証情報の読み込み ---
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
    page.goto("https://airregi.jp/CLP/view/top/", wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    
    # ログイン画面またはログインボタンが表示されているか判定
    if "login" in page.url or page.locator("text='AirID'").first.is_visible() or page.get_by_role("button", name="ログイン").is_visible():
        print("・ログイン画面を検知。自動ログイン処理を実行します...")
        
        if not AIRREGI_ID or not AIRREGI_PASS:
            raise ValueError("【認証エラー】AIRREGI_ID または AIRREGI_PASS が設定されていません。Streamlit Secretsをご確認ください。")

        if page.get_by_role("button", name="ログイン").is_visible() and "login" not in page.url:
            page.get_by_role("button", name="ログイン").click()
            page.wait_for_timeout(2000)

        # ID入力欄の特定
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
        
        # パスワード入力欄
        pass_input = page.locator("input[type='password']:visible, input[name='password']:visible").first
        if pass_input.is_visible():
            pass_input.fill(AIRREGI_PASS.strip())
        else:
            raise ValueError("パスワードの入力欄が見つかりませんでした。")

        # 「次回から入力を省略」にチェック
        try:
            checkbox = page.locator("input[type='checkbox']:visible").first
            if checkbox.is_visible() and not checkbox.is_checked():
                checkbox.check()
        except Exception:
            pass

        # ログインボタン押下
        login_btn = page.locator("button[type='submit']:visible, input[type='submit']:visible, button:has-text('ログイン'):visible").first
        login_btn.click()
        
        # 画面遷移待機（choose-store画面が出た場合は選択して進む）
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
                    else:
                        first_store_link = page.locator("ul li a:visible, main a:visible").first
                        if first_store_link.is_visible():
                            first_store_link.click()
                
                page.wait_for_timeout(3500)

            if "airregi.jp" in page.url and "connect.airregi.jp" not in page.url:
                break
        
        if "connect.airregi.jp" in page.url and "choose-store" not in page.url:
            body_text = page.locator("body").inner_text()
            if "認証コード" in body_text or "2段階認証" in body_text:
                raise RuntimeError("【2段階認証エラー】AirIDの2段階認証画面が表示されています。")
            
            error_elems = page.locator(".error, [class*='error'], [role='alert'], .attention").all_inner_texts()
            error_details = " / ".join([e.strip() for e in error_elems if e.strip()])
            if error_details:
                raise RuntimeError(f"【ログインエラー】Airレジ画面のメッセージ: {error_details}")
            
            raise RuntimeError(f"【ログイン未完了】画面遷移が完了しませんでした。URL: {page.url}")

        print(f"・Airレジへのログイン成功を確認（現在のURL: {page.url}）")
        page.wait_for_timeout(3000)

def select_store(page, store_name, is_first_store=False):
    print(f"\n--- 【店舗切り替え】「{store_name}」へ移動 ---")
    
    short_kw = ""
    for k in ["三田", "恵比寿", "五反田", "中目黒"]:
        if k in store_name:
            short_kw = k
            break

    # 最初から店舗選択一覧画面が出ている場合
    if page.locator("text='利用する店舗を選択'").is_visible() or page.locator("text='店舗を選択'").is_visible():
        target = page.locator(f"text='{store_name}'").or_(page.get_by_text(short_kw)).first
        target.wait_for(state="visible", timeout=15000)
        target.click()
        page.wait_for_timeout(4000)
        return

    # 1店舗目で既にその店舗が表示されている場合はスキップ
    if is_first_store:
        try:
            page.wait_for_selector(f"text='{short_kw}'", timeout=4000)
            print(f"・1店舗目「{store_name}」は既に選択済みです。")
            return
        except Exception:
            pass

    # ヘッダーの店舗名トリガーをクリック
    store_trigger = page.locator("header button, header a, div[class*='shop'], div[class*='Header']").filter(has_text="店").first
    if not store_trigger.is_visible():
        store_trigger = page.get_by_text("三田おにぎり").or_(page.get_by_text("おにぎり維心")).first

    store_trigger.wait_for(state="visible", timeout=15000)
    store_trigger.click()
    page.wait_for_timeout(1500)

    # 「店舗切り替え」ボタンをクリック
    switch_btn = page.locator("text='店舗切り替え'").or_(page.locator("text='店舗切替'")).first
    switch_btn.wait_for(state="visible", timeout=10000)
    switch_btn.click()
    page.wait_for_timeout(2500)

    # 店舗一覧から対象店舗をクリック（完全一致、全角スペース、
