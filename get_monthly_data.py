import os
import sys
import time
import platform
import subprocess
import re
from datetime import datetime, timedelta

# --- 最小修正1: タイムゾーンを日本時間(JST)に固定 ---
os.environ['TZ'] = 'Asia/Tokyo'
if hasattr(time, 'tzset'):
    time.tzset()

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

STORE_SELECT_URL = "https://airregi.jp/CLP/view/displayPlfSettingStoreSelect/?transitionUrl=https%3A%2F%2Fairregi.jp%2FCLP%2Fview%2Ftop%2F"

def auto_login(page):
    page.goto("https://airregi.jp/CLP/view/top/", wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    
    if "login" in page.url or page.locator("text='AirID'").first.is_visible() or page.get_by_role("button", name="ログイン").is_visible():
        print("・ログイン画面を検知。自動ログイン処理を実行します...")
        
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

    # --- 最小修正2: 三田本店の店舗切り替えスキップを廃止（確実に切り替える） ---
    # if is_first_store:
    #     ...

    print("・店舗選択画面へ移動中...")
    switch_link = page.locator("a[href*='displayPlfSettingStoreSelect']").first
    if switch_link.count() > 0:
        try:
            switch_link.click(force=True)
            page.wait_for_timeout(2500)
        except Exception:
            page.goto(STORE_SELECT_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)
    else:
        page.goto(STORE_SELECT_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)

    print(f"・店舗一覧から「{store_name}」({short_kw}) を選択中...")
    target = (
        page.locator(f":is(a, button, li, tr, div):has-text('{short_kw}')")
        .filter(has_text="店")
        .first
    )
    if not target.is_visible():
        target = page.locator(f"text='{short_kw}'").first

    try:
        target.wait_for(state="visible", timeout=15000)
        target.click(force=True)
    except Exception:
        clicked = page.evaluate("""(kw) => {
            const all = Array.from(document.querySelectorAll('a, button, li, tr, div, span'));
            for (let el of all) {
                const txt = (el.innerText || '').trim();
                if (txt.includes(kw)) {
                    el.click();
                    return true;
                }
            }
            return false;
        }""", short_kw)
        if not clicked:
            raise RuntimeError(f"店舗選択画面で「{store_name}」({short_kw}) が見つかりませんでした。")

    page.wait_for_timeout(3500)
    print(f"・「{store_name}」への切り替え完了")

def navigate_to_sales_page(page):
    print("・日別売上画面へ移動中...")
    page.wait_for_timeout(2000)

    if page.locator("button:has-text('表示する'), button:has-text('表示')").first.is_visible():
        print("・既に日別売上画面に滞在しています。")
        return

    try:
        sales_menu = page.locator("text='売上・分析'").first
        sales_menu.wait_for(state="visible", timeout=15000)
        if not page.locator("text='日別売上'").first.is_visible():
            sales_menu.click()
            page.wait_for_timeout(1500)
    except Exception as e:
        print(f"・メニュー待機エラー: {e}")

    try:
        daily_btn = page.locator("text='日別売上'").first
        daily_btn.wait_for(state="visible", timeout=15000)
        daily_btn.click()
        page.wait_for_timeout(3000)
    except Exception:
        page.evaluate("""() => {
            const all = Array.from(document.querySelectorAll('a, button, span, li, div'));
            for (let el of all) {
                if ((el.innerText || '').trim() === '日別売上') {
                    el.click();
                    return true;
                }
            }
            return false;
        }""")
        page.wait_for_timeout(3000)

    try:
        page.wait_for_selector("button:has-text('表示する'), button:has-text('表示'), select", timeout=15000)
        print("・日別売上画面の読み込み完了を確認")
    except Exception:
        print(f"・日別売上画面の読み込み待機タイムアウト (URL: {page.url})")

def extract_sales_for_date(page, target_date):
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    year_str, month_str, day_str = f"{dt.year}年", f"{dt.month}月", f"{dt.day}日"
    day_num_str = f"{dt.day}日"

    print(f"・ {target_date} ({year_str} {month_str} {day_num_str}) のデータを指定中...")
    page.wait_for_timeout(1000)

    try:
        page.evaluate("""() => {
            const selects = Array.from(document.querySelectorAll('select'));
            for (let sel of selects) {
                for (let opt of sel.options) {
                    if (opt.text.includes('時間別')) {
                        sel.value = opt.value;
                        sel.dispatchEvent(new Event('change', { bubbles: true }));
                        break;
                    }
                }
            }
        }""")
        page.wait_for_timeout(1000)
    except Exception:
        pass

    page.evaluate("""({ yearStr, monthStr, dayNumStr }) => {
        const selects = Array.from(document.querySelectorAll('select'));
        for (let sel of selects) {
            for (let opt of sel.options) {
                let t = opt.text.trim();
                if (t === yearStr || t === monthStr || t === dayNumStr || t === String(parseInt(dayNumStr))) {
                    sel.value = opt.value;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }
        }
    }""", {"yearStr": year_str, "monthStr": month_str, "dayNumStr": day_num_str})
    
    page.wait_for_timeout(1000)

    btn = page.locator("button:has-text('表示する'), button:has-text('表示'), input[value*='表示']").first
    if btn.is_visible():
        btn.click(force=True)
        print("・「表示する」ボタンを押下。データ読み込みを待機中...")
        # --- 最小修正3: テーブルが空になるのを防ぐため、3.5秒から5秒へ待機時間を延長 ---
        page.wait_for_timeout(5000)

    try:
        page.wait_for_selector("table tbody tr, table tr", timeout=15000)
        page.wait_for_timeout(2000)
    except Exception:
        pass

    rows = page.locator("table tr").all()
    store_rows = []
    for row in rows:
        cells = row.locator("th, td").all()
        row_text = [cell.inner_text().strip().replace("\n", " ") for cell in cells]
        if any(row_text):
            store_rows.append(row_text)

    print(f"・抽出行数: {len(store_rows)} 行 (現在のURL: {page.url})")
    return store_rows

def main():
    if len(sys.argv) > 2:
        start_date_str = sys.argv[1]
        end_date_str = sys.argv[2]
    elif len(sys.argv) > 1:
        start_date_str = sys.argv[1]
        end_date_str = sys.argv[1]
    else:
        d = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        start_date_str, end_date_str = d, d

    start_dt = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date_str, "%Y-%m-%d")
    
    date_list = []
    curr = start_dt
    while curr <= end_dt:
        date_list.append(curr.strftime("%Y-%m-%d"))
        curr += timedelta(days=1)

    print(f"=== 対象期間: {start_date_str} 〜 {end_date_str} ({len(date_list)}日間) ===")

    is_cloud = (platform.system() == "Linux")
    run_headless = True if is_cloud else False
    
    launch_args = ["--start-maximized"]
    if is_cloud:
        launch_args.extend([
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu"
        ])

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=run_headless,
            viewport={"width": 1920, "height": 1080},
            timezone_id="Asia/Tokyo",  # タイムゾーンの固定
            locale="ja-JP",
            args=launch_args
        )
        page = context.new_page()

        try:
            auto_login(page)

            for target_date in date_list:
                print(f"\n--- 【処理日】 {target_date} ---")
                all_stores_data = []

                for index, store_name in enumerate(TARGET_STORES):
                    is_first = (index == 0)
                    select_store(page, store_name, is_first_store=is_first)
                    
                    navigate_to_sales_page(page)

                    store_rows = extract_sales_for_date(page, target_date)
                    for r in store_rows:
                        all_stores_data.append([store_name] + r)

                if not all_stores_data:
                    raise RuntimeError(f"【データ取得失敗】{target_date} の売上テーブルからデータを1件も抽出できませんでした。現在のURL: {page.url}")

                df = pd.DataFrame(all_stores_data)
                csv_path = os.path.join(OUTPUT_DIR, f"4店舗_時間別売上_{target_date}.csv")
                df.to_csv(csv_path, index=False, header=False, encoding="utf-8-sig")
                print(f"[保存完了] {csv_path} (全 {len(all_stores_data)} 行)")

        except Exception as e:
            print(f"[ERROR] 処理中にエラーが発生しました: {e}")
            raise e
        finally:
            context.close()

if __name__ == "__main__":
    main()
