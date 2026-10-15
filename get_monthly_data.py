import os
import sys
import time
import platform
import subprocess
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
        
        # 画面遷移待機（connect.airregi.jp から離脱するまで最大15秒）
        print("・認証後の画面遷移を待機中...")
        for _ in range(15):
            page.wait_for_timeout(1000)
            if "connect.airregi.jp" not in page.url:
                break
        
        # 15秒経過後もログイン画面にとどまっている場合はエラー内容を特定して通知
        if "connect.airregi.jp" in page.url:
            body_text = page.locator("body").inner_text()
            if "認証コード" in body_text or "2段階認証" in body_text:
                raise RuntimeError("【2段階認証エラー】AirIDの2段階認証（メール認証コード要求）画面が表示されています。海外クラウドサーバーからの接続のため追加セキュリティが作動しています。")
            
            error_elems = page.locator(".error, [class*='error'], [role='alert'], .attention").all_inner_texts()
            error_details = " / ".join([e.strip() for e in error_elems if e.strip()])
            if error_details:
                raise RuntimeError(f"【ログインエラー】Airレジ画面のメッセージ: {error_details}")
            
            raise RuntimeError(f"【ログイン未完了】ログインボタン押下後も画面が遷移しませんでした。URL: {page.url}")

        print(f"・ログイン成功を確認（現在のURL: {page.url}）")
        page.wait_for_timeout(3000)

def select_store(page, store_name, is_first_store=False):
    print(f"\n--- 【店舗切り替え】「{store_name}」へ移動 ---")
    
    # お知らせ等のポップアップモーダルがあれば閉じる
    try:
        close_btn = page.locator("button:has-text('閉じる'), button[aria-label='閉じる'], .modal-close").first
        if close_btn.is_visible():
            close_btn.click()
            page.wait_for_timeout(1000)
    except Exception:
        pass

    # 店舗選択画面（一覧カードなど）が表示されている場合
    target_btn = page.locator(f"text='{store_name}'").first
    if target_btn.is_visible():
        target_btn.click()
        page.wait_for_timeout(4000)
        return

    # すでにその店舗の管理画面にいるか確認（1店舗目の場合）
    if is_first_store:
        try:
            page.wait_for_selector(f"text='{store_name}'", timeout=3000)
            return
        except Exception:
            pass

    # ヘッダーの店舗切り替えボタン操作
    store_trigger = page.locator("header button, header a, div[class*='shop'], div[class*='Header']").filter(has_text="店").first
    if not store_trigger.is_visible():
        store_trigger = page.get_by_text("三田おにぎり").or_(page.get_by_text("おにぎり維心")).first
    
    store_trigger.wait_for(state="visible", timeout=15000)
    store_trigger.click()
    page.wait_for_timeout(1500)
    
    switch_btn = page.locator("text='店舗切り替え', text='店舗切替'").first
    if switch_btn.is_visible():
        switch_btn.click()
        page.wait_for_timeout(2000)
        
    page.locator(f"text='{store_name}'").first.click()
    page.wait_for_timeout(3500)

def navigate_to_sales_page(page):
    sales_menu = page.locator("text='売上・分析'").first
    page.wait_for_selector("text='売上・分析'", timeout=15000)
    if not page.locator("text='日別売上'").first.is_visible():
        sales_menu.click()
        page.wait_for_timeout(1500)
    page.wait_for_selector("text='日別売上'", timeout=15000)
    page.locator("text='日別売上'").first.click()
    page.wait_for_timeout(2500)

def extract_sales_for_date(page, target_date):
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    year_str, month_str, day_str = f"{dt.year}年", f"{dt.month}月", f"{dt.day}日"
    day_num_str = f"{dt.day}日"

    print(f"・ {target_date} ({year_str} {month_str} {day_num_str}) のデータを指定中...")

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
        page.wait_for_timeout(3500)

    rows = page.locator("table tr").all()
    store_rows = []
    for row in rows:
        cells = row.locator("th, td").all()
        row_text = [cell.inner_text().strip().replace("\n", " ") for cell in cells]
        if any(row_text):
            store_rows.append(row_text)
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
                    
                    if index == 0 and target_date == date_list[0]:
                        navigate_to_sales_page(page)
                    else:
                        try:
                            page.locator("text='日別売上'").first.click()
                            page.wait_for_timeout(1500)
                        except Exception:
                            navigate_to_sales_page(page)

                    store_rows = extract_sales_for_date(page, target_date)
                    for r in store_rows:
                        all_stores_data.append([store_name] + r)

                if all_stores_data:
                    df = pd.DataFrame(all_stores_data)
                    csv_path = os.path.join(OUTPUT_DIR, f"4店舗_時間別売上_{target_date}.csv")
                    df.to_csv(csv_path, index=False, header=False, encoding="utf-8-sig")
                    print(f"[保存完了] {csv_path}")

        except Exception as e:
            print(f"[ERROR] 処理中にエラーが発生しました: {e}")
            raise e
        finally:
            context.close()

if __name__ == "__main__":
    main()
