import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime, timedelta
import calendar
import glob
import subprocess
import io

# openpyxlのスタイリング設定用モジュール
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ページ設定
st.set_page_config(
    page_title="売上月報管理システム",
    page_icon="📊",
    layout="wide"
)

OUTPUT_DIR = os.path.abspath("./output")
DATA_DIR = os.path.abspath("./data")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

BUDGET_PREV_PATH = os.path.join(DATA_DIR, "budget_prev.csv")
DELIVERY_DATA_PATH = os.path.join(DATA_DIR, "delivery_sales.csv")

# 店舗定義
STORE_CREDENTIALS = {
    "三田おにぎり 本店": {"id": "001", "pass": "0000", "short_label": "三田店"},
    "おにぎり維心 恵比寿店": {"id": "002", "pass": "0000", "short_label": "恵比寿店"},
    "おにぎり維心 五反田店": {"id": "003", "pass": "0000", "short_label": "五反田店"},
    "おにぎり維心 中目黒店": {"id": "004", "pass": "0000", "short_label": "中目黒店"}
}

STORES = list(STORE_CREDENTIALS.keys())
TIME_SLOTS = ["早番", "中番", "遅番"]

def normalize_time_slot(val):
    s = str(val)
    if "早" in s:
        return "早番"
    elif "中" in s:
        return "中番"
    elif "遅" in s:
        return "遅番"
    return s

def init_data_files():
    if not os.path.exists(BUDGET_PREV_PATH):
        records = []
        for day in range(1, 32):
            date_str = f"2026-10-{day:02d}"
            for store in STORES:
                for slot in TIME_SLOTS:
                    if slot == "早番":
                        b_sales, p_sales, p_cust = 88000, 77645, 90
                    elif slot == "中番":
                        b_sales, p_sales, p_cust = 157000, 153848, 194
                    else:
                        b_sales, p_sales, p_cust = 55000, 48521, 78

                    records.append({
                        "date": date_str,
                        "store": store,
                        "time_slot": slot,
                        "budget_sales": b_sales,
                        "prev_sales": p_sales,
                        "prev_cust": p_cust
                    })
        pd.DataFrame(records).to_csv(BUDGET_PREV_PATH, index=False, encoding="utf-8-sig")

    if not os.path.exists(DELIVERY_DATA_PATH):
        pd.DataFrame(columns=["date", "store", "time_slot", "uber", "rocket_now", "demaekan", "other"]).to_csv(DELIVERY_DATA_PATH, index=False, encoding="utf-8-sig")

init_data_files()

def load_budget_prev_data():
    df = pd.read_csv(BUDGET_PREV_PATH)
    df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
    return df

def load_delivery_data():
    df = pd.read_csv(DELIVERY_DATA_PATH)
    df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
    if "rn" in df.columns and "rocket_now" not in df.columns:
        df = df.rename(columns={"rn": "rocket_now"})
    return df

def save_delivery_data(date_str, store, time_slot, uber, rocket_now, demaekan, other):
    df = load_delivery_data()
    df = df[~((df["date"] == date_str) & (df["store"] == store) & (df["time_slot"] == time_slot))]
    
    new_row = pd.DataFrame([{
        "date": date_str,
        "store": store,
        "time_slot": time_slot,
        "uber": uber,
        "rocket_now": rocket_now,
        "demaekan": demaekan,
        "other": other
    }])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_csv(DELIVERY_DATA_PATH, index=False, encoding="utf-8-sig")

def parse_int(val):
    if pd.isna(val):
        return 0
    clean_str = re.sub(r"[^\d-]", "", str(val))
    try:
        return int(clean_str)
    except ValueError:
        return 0

def find_and_read_airregi(date_str):
    d_dash = date_str
    d_nodash = date_str.replace("-", "")
    d_under = date_str.replace("-", "_")

    all_files = glob.glob(os.path.join(OUTPUT_DIR, "*.csv"))
    matched = []
    for f in all_files:
        base = os.path.basename(f)
        if d_dash in base or d_nodash in base or d_under in base:
            matched.append(f)
    
    if not matched:
        return None, "ファイルなし"
    
    matched.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    try:
        raw_df = pd.read_csv(matched[0], header=None)
        return raw_df, os.path.basename(matched[0])
    except Exception as e:
        return None, str(e)

def get_short_store_key(store_name):
    for key in ["三田", "恵比寿", "五反田", "中目黒"]:
        if key in store_name:
            return key
    return store_name

def match_slot_row(store_df, slot_name):
    short_slot = slot_name[0] if slot_name else ""
    matched = store_df[store_df[1].astype(str).str.contains(short_slot, na=False)]
    return matched

# 月間データフレーム生成用の汎用関数
def generate_month_df(selected_store, year, month, target_slot_keyword, bp_df, deliv_df):
    _, max_days = calendar.monthrange(year, month)
    rows = []

    for day in range(1, max_days + 1):
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
        bp_sub = bp_df[(bp_df["date"] == date_str) & (bp_df["store"] == selected_store)]
        if target_slot_keyword:
            bp_sub = bp_sub[bp_sub["time_slot"] == target_slot_keyword]

        p_sales = parse_int(bp_sub["prev_sales"].sum()) if not bp_sub.empty else 0
        b_sales = parse_int(bp_sub["budget_sales"].sum()) if not bp_sub.empty else 0
        p_cust = parse_int(bp_sub["prev_cust"].sum()) if not bp_sub.empty else 0

        air_sales = 0
        air_cust = 0
        air_other_sales = 0
        raw_df, _ = find_and_read_airregi(date_str)

        if raw_df is not None:
            try:
                short_key = get_short_store_key(selected_store)
                store_df = raw_df[raw_df[0].astype(str).str.contains(short_key, na=False)]

                if target_slot_keyword is None:
                    matched = store_df[store_df[1].astype(str).str.strip() == "合計"]
                else:
                    matched = match_slot_row(store_df, target_slot_keyword)

                if not matched.empty:
                    air_sales = parse_int(matched.iloc[0, 2])
                    air_cust = parse_int(matched.iloc[0, 5])
                    if matched.shape[1] > 11:
                        air_other_sales = parse_int(matched.iloc[0, 11])
            except Exception:
                pass

        deliv_sub = deliv_df[(deliv_df["date"] == date_str) & (deliv_df["store"] == selected_store)]
        if target_slot_keyword:
            deliv_sub = deliv_sub[deliv_sub["time_slot"] == target_slot_keyword]
        
        day_delivery_sales = 0
        deliv_other_sales = 0
        if not deliv_sub.empty:
            day_delivery_sales = int(
                pd.to_numeric(deliv_sub["uber"], errors="coerce").fillna(0).sum() +
                pd.to_numeric(deliv_sub["rocket_now"], errors="coerce").fillna(0).sum() +
                pd.to_numeric(deliv_sub["demaekan"], errors="coerce").fillna(0).sum()
            )
            deliv_other_sales = int(
                pd.to_numeric(deliv_sub["other"], errors="coerce").fillna(0).sum()
            )

        total_other_sales = air_other_sales + deliv_other_sales
        actual_sales = air_sales + day_delivery_sales + total_other_sales
        actual_cust = air_cust

        prev_ratio = f"{(actual_sales / p_sales * 100):.1f}%" if p_sales > 0 and actual_sales != 0 else ("-" if actual_sales == 0 else "0.0%")
        budget_ratio = f"{(actual_sales / b_sales * 100):.1f}%" if b_sales > 0 and actual_sales != 0 else ("-" if actual_sales == 0 else "0.0%")
        
        p_price = int(p_sales / p_cust) if p_cust > 0 else 0
        a_price = int(actual_sales / actual_cust) if actual_cust > 0 else 0

        cust_ratio = f"{(actual_cust / p_cust * 100):.1f}%" if p_cust > 0 and actual_cust > 0 else "-"
        price_ratio = f"{(a_price / p_price * 100):.1f}%" if p_price > 0 and a_price != 0 else "-"

        weekday_kanji = ["月", "火", "水", "木", "金", "土", "日"][pd.to_datetime(date_str).weekday()]

        rows.append({
            "日": day,
            "曜日": weekday_kanji,
            "前年売上": p_sales,
            "予算売上": b_sales,
            "実績売上": actual_sales,
            "前年比": prev_ratio,
            "予算比": budget_ratio,
            "店頭売上": air_sales,
            "デリバリー売上": day_delivery_sales,
            "その他売上": total_other_sales,
            "前年客数": p_cust,
            "実績客数": actual_cust,
            "客数前年比": cust_ratio,
            "前年単価": p_price,
            "実績単価": a_price,
            "単価前年比": price_ratio
        })

    return pd.DataFrame(rows)

# サマリー行（合計, AVE, MIN, MAX）の動的計算関数
def calculate_summary_rows(df_slot):
    df_act = df_slot[df_slot["実績売上"] != 0]
    if df_act.empty:
        df_act = df_slot.iloc[:1]
        
    num_cols = ["前年売上", "予算売上", "実績売上", "店頭売上", "デリバリー売上", "その他売上", "前年客数", "実績客数"]
    
    # 1. 合計
    sum_vals = {col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).sum()) for col in num_cols}
    s_p_ratio = f"{(sum_vals['実績売上'] / sum_vals['前年売上'] * 100):.1f}%" if sum_vals['前年売上'] > 0 else "-"
    s_b_ratio = f"{(sum_vals['実績売上'] / sum_vals['予算売上'] * 100):.1f}%" if sum_vals['予算売上'] > 0 else "-"
    s_c_ratio = f"{(sum_vals['実績客数'] / sum_vals['前年客数'] * 100):.1f}%" if sum_vals['前年客数'] > 0 else "-"
    s_p_price = int(sum_vals['前年売上'] / sum_vals['前年客数']) if sum_vals['前年客数'] > 0 else 0
    s_a_price = int(sum_vals['実績売上'] / sum_vals['実績客数']) if sum_vals['実績客数'] > 0 else 0
    s_pr_ratio = f"{(s_a_price / s_p_price * 100):.1f}%" if s_p_price > 0 else "-"
    
    row_sum = {
        "日": "合計", "曜日": "",
        "前年売上": sum_vals["前年売上"], "予算売上": sum_vals["予算売上"], "実績売上": sum_vals["実績売上"],
        "前年比": s_p_ratio, "予算比": s_b_ratio,
        "店頭売上": sum_vals["店頭売上"], "デリバリー売上": sum_vals["デリバリー売上"], "その他売上": sum_vals["その他売上"],
        "前年客数": sum_vals["前年客数"], "実績客数": sum_vals["実績客数"], "客数前年比": s_c_ratio,
        "前年単価": s_p_price, "実績単価": s_a_price, "単価前年比": s_pr_ratio
    }
    
    # 2. AVE (平均)
    ave_vals = {col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).mean()) for col in num_cols}
    a_p_ratio = f"{(ave_vals['実績売上'] / ave_vals['前年売上'] * 100):.1f}%" if ave_vals['前年売上'] > 0 else "-"
    a_b_ratio = f"{(ave_vals['実績売上'] / ave_vals['予算売上'] * 100):.1f}%" if ave_vals['予算売上'] > 0 else "-"
    a_c_ratio = f"{(ave_vals['実績客数'] / ave_vals['前年客数'] * 100):.1f}%" if ave_vals['前年客数'] > 0 else "-"
    a_p_price = int(ave_vals['前年売上'] / ave_vals['前年客数']) if ave_vals['前年客数'] > 0 else 0
    a_a_price = int(ave_vals['実績売上'] / ave_vals['実績客数']) if ave_vals['実績客数'] > 0 else 0
    a_pr_ratio = f"{(a_a_price / a_p_price * 100):.1f}%" if a_p_price > 0 else "-"
    
    row_ave = {
        "日": "AVE", "曜日": "",
        "前年売上": ave_vals["前年売上"], "予算売上": ave_vals["予算売上"], "実績売上": ave_vals["実績売上"],
        "前年比": a_p_ratio, "予算比": a_b_ratio,
        "店頭売上": ave_vals["店頭売上"], "デリバリー売上": ave_vals["デリバリー売上"], "その他売上": ave_vals["その他売上"],
        "前年客数": ave_vals["前年客数"], "実績客数": ave_vals["実績客数"], "客数前年比": a_c_ratio,
        "前年単価": a_p_price, "実績単価": a_a_price, "単価前年比": a_pr_ratio
    }
    
    # 3. MIN (最小)
    min_vals = {col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).min()) for col in num_cols}
    m_p_ratio = f"{(min_vals['実績売上'] / min_vals['前年売上'] * 100):.1f}%" if min_vals['前年売上'] > 0 else "-"
    m_b_ratio = f"{(min_vals['実績売上'] / min_vals['予算売上'] * 100):.1f}%" if min_vals['予算売上'] > 0 else "-"
    m_c_ratio = f"{(min_vals['実績客数'] / min_vals['前年客数'] * 100):.1f}%" if min_vals['前年客数'] > 0 else "-"
    m_p_price = int(pd.to_numeric(df_act["前年単価"], errors="coerce").fillna(0).min())
    m_a_price = int(pd.to_numeric(df_act["実績単価"], errors="coerce").fillna(0).min())
    m_pr_ratio = f"{(m_a_price / m_p_price * 100):.1f}%" if m_p_price > 0 else "-"
    
    row_min = {
        "日": "MIN", "曜日": "",
        "前年売上": min_vals["前年売上"], "予算売上": min_vals["予算売上"], "実績売上": min_vals["実績売上"],
        "前年比": m_p_ratio, "予算比": m_b_ratio,
        "店頭売上": min_vals["店頭売上"], "デリバリー売上": min_vals["デリバリー売上"], "その他売上": min_vals["その他売上"],
        "前年客数": min_vals["前年客数"], "実績客数": min_vals["実績客数"], "客数前年比": m_c_ratio,
        "前年単価": m_p_price, "実績単価": m_a_price, "単価前年比": m_pr_ratio
    }
    
    # 4. MAX (最大)
    max_vals = {col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).max()) for col in num_cols}
    mx_p_ratio = f"{(max_vals['実績売上'] / max_vals['前年売上'] * 100):.1f}%" if max_vals['前年売上'] > 0 else "-"
    mx_b_ratio = f"{(max_vals['実績売上'] / max_vals['予算売上'] * 100):.1f}%" if max_vals['予算売上'] > 0 else "-"
    mx_c_ratio = f"{(max_vals['実績客数'] / max_vals['前年客数'] * 100):.1f}%" if max_vals['前年客数'] > 0 else "-"
    mx_p_price = int(pd.to_numeric(df_act["前年単価"], errors="coerce").fillna(0).max())
    mx_a_price = int(pd.to_numeric(df_act["実績単価"], errors="coerce").fillna(0).max())
    mx_pr_ratio = f"{(mx_a_price / mx_p_price * 100):.1f}%" if mx_p_price > 0 else "-"
    
    row_max = {
        "日": "MAX", "曜日": "",
        "前年売上": max_vals["前年売上"], "予算売上": max_vals["予算売上"], "実績売上": max_vals["実績売上"],
        "前年比": mx_p_ratio, "予算比": mx_b_ratio,
        "店頭売上": max_vals["店頭売上"], "デリバリー売上": max_vals["デリバリー売上"], "その他売上": max_vals["その他売上"],
        "前年客数": max_vals["前年客数"], "実績客数": max_vals["実績客数"], "客数前年比": mx_c_ratio,
        "前年単価": mx_p_price, "実績単価": mx_a_price, "単価前年比": mx_pr_ratio
    }
    
    return [row_sum, row_ave, row_min, row_max]

# 全時間帯のデータをデザイン整形済みのタブ付きExcelとして一括出力する関数
def create_multi_sheet_excel(selected_store, year, month, bp_df, deliv_df):
    output = io.BytesIO()
    slots = [("合計", None), ("早番", "早番"), ("中番", "中番"), ("遅番", "遅番")]
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
    
    even_row_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    odd_row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    data_font = Font(name="游ゴシック", size=10)
    summary_label_font = Font(name="游ゴシック", size=10, bold=True)
    center_alignment = Alignment(horizontal="center", vertical="center")
    
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        workbook = writer.book
        
        for sheet_name, slot_kw in slots:
            df_slot = generate_month_df(selected_store, year, month, slot_kw, bp_df, deliv_df)
            cols = list(df_slot.columns)
            
            ws = workbook.create_sheet(title=sheet_name)
            
            title_text = f"{selected_store} {year}年{month}月度（{sheet_name}）"
            ws.cell(row=1, column=1, value=title_text)
            ws.cell(row=1, column=1).font = Font(name="游ゴシック", size=13, bold=True, color="1F4E78")
            ws.cell(row=1, column=1).alignment = Alignment(vertical="center")
            
            for col_idx, col_name in enumerate(cols, start=1):
                cell = ws.cell(row=3, column=col_idx)
                if col_idx > 2:
                    cell.value = col_name
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_alignment
                cell.border = thin_border
            
            summary_rows = calculate_summary_rows(df_slot)
            for s_idx, s_row in enumerate(summary_rows, start=4):
                fill = even_row_fill if s_idx % 2 == 0 else odd_row_fill
                for col_idx, col_name in enumerate(cols, start=1):
                    cell = ws.cell(row=s_idx, column=col_idx)
                    val = s_row.get(col_name, "")
                    if col_idx == 1:
                        cell.value = val
                        cell.font = summary_label_font
                    elif col_idx == 2:
                        cell.value = ""
                    else:
                        cell.value = val
                        cell.font = data_font
                    
                    cell.fill = fill
                    cell.alignment = center_alignment
                    cell.border = thin_border

            for col_idx, col_name in enumerate(cols, start=1):
                cell = ws.cell(row=9, column=col_idx, value=col_name)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_alignment
                cell.border = thin_border
                
            for r_idx, row_data in enumerate(df_slot.to_dict(orient="records"), start=10):
                fill = even_row_fill if r_idx % 2 == 0 else odd_row_fill
                for col_idx, col_name in enumerate(cols, start=1):
                    cell = ws.cell(row=r_idx, column=col_idx, value=row_data[col_name])
                    cell.fill = fill
                    cell.font = data_font
                    cell.alignment = center_alignment
                    cell.border = thin_border

            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 8

            for col in list(ws.columns)[2:]:
                max_len = max(len(str(cell.value or '')) for cell in list(col)[2:])
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        if "Sheet" in workbook.sheetnames:
            workbook.remove(workbook["Sheet"])
                
    return output.getvalue()

# --- セッションステートの初期化 ---
if "selected_store" not in st.session_state:
    st.session_state["selected_store"] = STORES[0]
if "selected_year_month" not in st.session_state:
    st.session_state["selected_year_month"] = "2026-10"

# --- 共有日付ステートの初期化（デフォルトは当日） ---
if "target_date_state" not in st.session_state:
    st.session_state["target_date_state"] = datetime.now().date()

# --- サイドバー設定 ---
st.sidebar.title("メニュー")

st.sidebar.markdown("---")
st.sidebar.subheader("店舗・期間設定")

st.sidebar.selectbox(
    "閲覧店舗", 
    STORES, 
    key="selected_store"
)
st.sidebar.selectbox(
    "対象年月", 
    ["2026-10"], 
    key="selected_year_month"
)

with st.sidebar.expander("📥 売上データを取得する"):
    target_date = st.date_input("取得日", value=datetime.now().date(), key="fetch_date_input_sidebar")
    if st.button("🚀 取得実行", type="primary"):
        target_str = target_date.strftime("%Y-%m-%d")
        with st.spinner(f"{target_str} 取得中..."):
            try:
                subprocess.run(
                    ["python", "get_monthly_data.py", target_str, target_str],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=True
                )
                st.success(f"🎉 取得完了！")
                st.rerun()
            except subprocess.CalledProcessError as e:
                st.error(f"❌ エラー:\n{e}")

st.sidebar.markdown("---")
app_mode = st.sidebar.radio(
    "機能を選択してください", 
    [
        "📊 本部用（月報一覧表示）",
        "📱 店舗用（デリバリー入力）",
        "📋 エクセルコピペ用データ生成",
        "📄 報告書作成"
    ]
)

selected_store = st.session_state["selected_store"]
selected_year_month = st.session_state["selected_year_month"]
year, month = map(int, selected_year_month.split("-"))

if app_mode == "📱 店舗用（デリバリー入力）":
    st.title("🛵 デリバリー売上 入力フォーム")
    st.markdown(f"現在選択中の店舗: **{selected_store}** | 対象の日付、時間帯を選択してデリバリー売上を入力・更新します。")

    col_target1, col_target2 = st.columns(2)
    with col_target1:
        st.info(f"📍 対象店舗: **{selected_store}** （サイドバーで変更可能）")
    with col_target2:
        selected_slot = st.selectbox("時間帯を選択", TIME_SLOTS)

    # 共有の日付ステートと同期する対象日付入力
    selected_date_val = st.date_input("対象日付", value=st.session_state["target_date_state"], key="deliv_date_input")
    if selected_date_val != st.session_state["target_date_state"]:
        st.session_state["target_date_state"] = selected_date_val
    selected_date = selected_date_val.strftime("%Y-%m-%d")
    
    st.markdown("---")
    
    deliv_df = load_delivery_data()
    existing_row = deliv_df[(deliv_df["date"] == selected_date) & (deliv_df["store"] == selected_store) & (deliv_df["time_slot"] == selected_slot)]
    
    is_already_registered = not existing_row.empty

    default_uber = 0
    default_rn = 0
    default_demae = 0
    default_other = 0

    if is_already_registered:
        if "uber" in existing_row.columns:
            default_uber = parse_int(existing_row["uber"].values[0])
        if "rocket_now" in existing_row.columns:
            default_rn = parse_int(existing_row["rocket_now"].values[0])
        if "demaekan" in existing_row.columns:
            default_demae = parse_int(existing_row["demaekan"].values[0])
        if "other" in existing_row.columns:
            default_other = parse_int(existing_row["other"].values[0])

        st.info(f"ℹ️ この日時（{selected_date} / {selected_slot}）のデータはすでに登録されています。変更して保存すると上書きされます。")

    # --- 入力ボックス内の文字スタイル設定（登録済みなら黒字太字、マイナス数値は赤字太字） ---
    css_input_rules = []
    if is_already_registered:
        c_uber = "#c0392b" if default_uber < 0 else "#000000"
        c_rn = "#c0392b" if default_rn < 0 else "#000000"
        c_demae = "#c0392b" if default_demae < 0 else "#000000"
        c_other = "#c0392b" if default_other < 0 else "#000000"

        css_input_rules.append(f'input[aria-label="Uber Eats 売上 (円)"] {{ color: {c_uber} !important; font-weight: bold !important; font-size: 16px !important; }}')
        css_input_rules.append(f'input[aria-label="Rocket now 売上 (円)"] {{ color: {c_rn} !important; font-weight: bold !important; font-size: 16px !important; }}')
        css_input_rules.append(f'input[aria-label="出前館 売上 (円)"] {{ color: {c_demae} !important; font-weight: bold !important; font-size: 16px !important; }}')
        css_input_rules.append(f'input[aria-label="その他 売上 (円)"] {{ color: {c_other} !important; font-weight: bold !important; font-size: 16px !important; }}')

    if css_input_rules:
        st.markdown("<style>" + "\n".join(css_input_rules) + "</style>", unsafe_allow_html=True)

    slot_key = f"{selected_date}_{selected_store}_{selected_slot}"

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        uber_val = st.number_input("Uber Eats 売上 (円)", value=default_uber, step=500, key=f"input_uber_{slot_key}")
        rn_val = st.number_input("Rocket now 売上 (円)", value=default_rn, step=500, key=f"input_rn_{slot_key}")
    with col_d2:
        demaekan_val = st.number_input("出前館 売上 (円)", value=default_demae, step=500, key=f"input_demae_{slot_key}")
        other_val = st.number_input("その他 売上 (円)", value=default_other, step=500, key=f"input_other_{slot_key}")

    st.markdown("---")
    total_delivery = uber_val + rn_val + demaekan_val + other_val
    
    if total_delivery < 0:
        st.markdown(f"💡 この時間帯の売上合計: <span style='color: #c0392b; font-weight: bold; font-size: 1.15em;'>¥{total_delivery:,}</span>", unsafe_allow_html=True)
    else:
        st.markdown(f"💡 この時間帯の売上合計: **¥{total_delivery:,}**")

    if st.button("💾 デリバリー売上を保存する", type="primary"):
        save_delivery_data(selected_date, selected_store, selected_slot, uber_val, rn_val, demaekan_val, other_val)
        st.success(f"🎉 {selected_store} ({selected_date} / {selected_slot}) のデリバリー売上を保存しました！")
        st.rerun()

elif app_mode == "📋 エクセルコピペ用データ生成":
    st.title("📋 エクセル実績入力用データ生成")
    st.markdown(f"現在選択中の店舗: **{selected_store}** | エクセルシートへ簡単にコピペするためのデータを生成します。")

    # 共有の日付ステートと同期する対象日付入力
    target_date_val = st.date_input("対象日付", value=st.session_state["target_date_state"], key="copy_date_input")
    if target_date_val != st.session_state["target_date_state"]:
        st.session_state["target_date_state"] = target_date_val
    target_date_str = target_date_val.strftime("%Y-%m-%d")

    raw_df, file_status = find_and_read_airregi(target_date_str)
    
    if raw_df is None:
        st.warning(f"⚠️ **{target_date_str}** のAirレジCSVファイルが見つかりません。（Status: {file_status}）")

    def get_slot_metrics(store_name, slot_name, df_data):
        air_sales = 0
        air_cust = 0
        if df_data is not None:
            try:
                short_key = get_short_store_key(store_name)
                store_df = df_data[df_data[0].astype(str).str.contains(short_key, na=False)]
                matched = match_slot_row(store_df, slot_name)
                if not matched.empty:
                    air_sales = parse_int(matched.iloc[0, 2])
                    air_cust = parse_int(matched.iloc[0, 5])
            except Exception:
                pass
        return air_sales, air_cust

    def get_slot_delivery_total(store_name, date_str, slot_name):
        deliv_df = load_delivery_data()
        sub = deliv_df[
            (deliv_df["date"].astype(str).str.strip() == date_str.strip()) & 
            (deliv_df["store"].astype(str).str.strip() == store_name.strip()) & 
            (deliv_df["time_slot"].astype(str).str.strip() == slot_name.strip())
        ]
        if sub.empty:
            return 0
        return int(
            pd.to_numeric(sub["uber"], errors="coerce").fillna(0).sum() +
            pd.to_numeric(sub["rocket_now"], errors="coerce").fillna(0).sum() +
            pd.to_numeric(sub["demaekan"], errors="coerce").fillna(0).sum()
        )

    early_d = get_slot_delivery_total(selected_store, target_date_str, "早番")
    mid_d = get_slot_delivery_total(selected_store, target_date_str, "中番")
    late_d = get_slot_delivery_total(selected_store, target_date_str, "遅番")
    
    early_sales, early_cust = get_slot_metrics(selected_store, "早番", raw_df)
    mid_sales, mid_cust = get_slot_metrics(selected_store, "中番", raw_df)
    late_sales, late_cust = get_slot_metrics(selected_store, "遅番", raw_df)

    # 該当日のその他売上を取得
    bp_df_tmp = load_budget_prev_data()
    deliv_df_tmp = load_delivery_data()
    df_month_all = generate_month_df(selected_store, target_date_val.year, target_date_val.month, None, bp_df_tmp, deliv_df_tmp)
    day_row = df_month_all[df_month_all["日"] == target_date_val.day]
    target_other_sales = int(day_row["その他売上"].values[0]) if not day_row.empty else 0

    parts = [
        str(early_cust), str(early_sales), str(early_d),
        str(mid_cust), str(mid_sales), str(mid_d),
        str(late_cust), str(late_sales), str(late_d),
        str(target_other_sales)
    ]

    tabs_list = [
        "\t" * 4, "\t" * 1, "\t" * 3,
        "\t" * 4, "\t" * 1, "\t" * 3,
        "\t" * 4, "\t" * 1, "\t" * 1
    ]
    
    tsv_text = "".join([part + tabs_list[i] if i < len(tabs_list) else part for i, part in enumerate(parts)])
    
    st.success(f"🎉 {target_date_str} ({selected_store}) のデータを自動生成しました！")
    st.code(tsv_text, language="text")

elif app_mode == "📄 報告書作成":
    st.title("📄 LINE報告書作成")
    st.markdown("指定した日付（または直近実績日まで）の全店舗の数値を集計し、LINE報告用のテキストを一括生成します。")
    st.markdown("---")

    # 共有の日付ステートと同期する集計基準日入力
    report_date = st.date_input("集計基準日（この日までを累計します）", value=st.session_state["target_date_state"], key="report_date_input")
    if report_date != st.session_state["target_date_state"]:
        st.session_state["target_date_state"] = report_date
    
    if st.button("📝 報告用テキストを生成する", type="primary"):
        bp_df = load_budget_prev_data()
        deliv_df = load_delivery_data()
        
        _, max_days = calendar.monthrange(year, month)
        target_end_day = min(report_date.day, max_days)

        report_lines = [
            "お疲れ様です。",
            f"{report_date.month}月{report_date.day}日時点の中間報告です。",
            ""
        ]

        for store_name in STORES:
            short_label = STORE_CREDENTIALS[store_name]["short_label"]
            
            total_p = 0
            total_b = 0
            total_act = 0

            for day in range(1, target_end_day + 1):
                date_str = f"{year:04d}-{month:02d}-{day:02d}"
                
                bp_sub = bp_df[(bp_df["date"] == date_str) & (bp_df["store"] == store_name)]
                p_sales = parse_int(bp_sub["prev_sales"].sum()) if not bp_sub.empty else 0
                b_sales = parse_int(bp_sub["budget_sales"].sum()) if not bp_sub.empty else 0

                air_sales = 0
                air_other_sales = 0
                raw_df, _ = find_and_read_airregi(date_str)
                if raw_df is not None:
                    try:
                        short_key = get_short_store_key(store_name)
                        store_df = raw_df[raw_df[0].astype(str).str.contains(short_key, na=False)]
                        matched = store_df[store_df[1].astype(str).str.strip() == "合計"]
                        if not matched.empty:
                            air_sales = parse_int(matched.iloc[0, 2])
                            if matched.shape[1] > 11:
                                air_other_sales = parse_int(matched.iloc[0, 11])
                    except Exception:
                        pass

                deliv_sub = deliv_df[(deliv_df["date"] == date_str) & (deliv_df["store"] == store_name)]
                day_delivery_sales = 0
                deliv_other_sales = 0
                if not deliv_sub.empty:
                    day_delivery_sales = int(
                        pd.to_numeric(deliv_sub["uber"], errors="coerce").fillna(0).sum() +
                        pd.to_numeric(deliv_sub["rocket_now"], errors="coerce").fillna(0).sum() +
                        pd.to_numeric(deliv_sub["demaekan"], errors="coerce").fillna(0).sum()
                    )
                    deliv_other_sales = int(
                        pd.to_numeric(deliv_sub["other"], errors="coerce").fillna(0).sum()
                    )

                actual_sales = air_sales + day_delivery_sales + air_other_sales + deliv_other_sales

                total_p += p_sales
                total_b += b_sales
                total_act += actual_sales

            prev_ratio = f"{(total_act / total_p * 100):.1f}%" if total_p > 0 else "-"
            budget_ratio = f"{(total_act / total_b * 100):.1f}%" if total_b > 0 else "-"
            diff_budget = total_act - total_b
            
            if diff_budget > 0:
                diff_str = f"+{diff_budget:,}円"
            elif diff_budget < 0:
                diff_str = f"▲{abs(diff_budget):,}円"
            else:
                diff_str = "0円"

            report_lines.append(f"【{short_label}】")
            report_lines.append(f"前年比 : {prev_ratio}")
            report_lines.append(f"予算比 : {budget_ratio}")
            report_lines.append(f"累売上 : {total_act:,}円")
            report_lines.append(f"予乖離 : {diff_str}")
            report_lines.append("")

        final_report_text = "\n".join(report_lines).strip()

        st.success("🎉 報告用テキストを生成しました！以下の枠内をコピーしてLINEに貼り付けてください。")
        st.code(final_report_text, language="text")

else:
    st.title("📊 売上月報一覧")
    st.subheader(f"🏬 {selected_store} {year}年{month}月度")
    st.markdown("---")

    bp_df = load_budget_prev_data()
    deliv_df = load_delivery_data()

    def format_currency(val):
        if pd.isna(val):
            return "-"
        try:
            num = int(val)
            if num < 0:
                return f"-¥{abs(num):,}"
            elif num > 0:
                return f"¥{num:,}"
            else:
                return "¥0"
        except (ValueError, TypeError):
            return str(val)

    # --- 1. 累計サマリー用のデータ生成（常に店舗全体の合計：target_slot_keyword=None） ---
    df_month_total = generate_month_df(selected_store, year, month, None, bp_df, deliv_df)

    df_achieved = df_month_total[df_month_total["実績売上"] != 0]
    if df_achieved.empty:
        max_day = 1
        df_achieved = df_month_total.iloc[:1]
    else:
        max_day = int(df_achieved["日"].max())

    total_p = pd.to_numeric(df_achieved["前年売上"], errors="coerce").fillna(0).sum()
    total_b = pd.to_numeric(df_achieved["予算売上"], errors="coerce").fillna(0).sum()
    total_act = pd.to_numeric(df_achieved["実績売上"], errors="coerce").fillna(0).sum()
    total_air = pd.to_numeric(df_achieved["店頭売上"], errors="coerce").fillna(0).sum()
    total_deliv = pd.to_numeric(df_achieved["デリバリー売上"], errors="coerce").fillna(0).sum()
    total_other = pd.to_numeric(df_achieved["その他売上"], errors="coerce").fillna(0).sum()
    total_pcust = pd.to_numeric(df_achieved["前年客数"], errors="coerce").fillna(0).sum()
    total_acust = pd.to_numeric(df_achieved["実績客数"], errors="coerce").fillna(0).sum()

    tot_pprice = int(total_p / total_pcust) if total_pcust > 0 else 0
    tot_aprice = int(total_act / total_acust) if total_acust > 0 else 0

    p_ratio_str = f"{(total_act / total_p * 100):.1f}%" if total_p > 0 else "-"
    b_ratio_str = f"{(total_act / total_b * 100):.1f}%" if total_b > 0 else "-"
    cust_ratio_str = f"{(total_acust / total_pcust * 100):.1f}%" if total_pcust > 0 else "-"
    price_ratio_str = f"{(tot_aprice / tot_pprice * 100):.1f}%" if tot_pprice > 0 else "-"

    # 累計サマリーデータ構築
    summary_data = [{
        "日": f"{max_day}日時点",
        "曜日": "合計",
        "前年比": p_ratio_str,
        "予算比": b_ratio_str,
        "前年売上": format_currency(total_p),
        "予算売上": format_currency(total_b),
        "実績売上": format_currency(total_act),
        "店頭売上": format_currency(total_air),
        "デリバリー売上": format_currency(total_deliv),
        "その他売上": format_currency(total_other),
        "前年客数": int(total_pcust),
        "実績客数": int(total_acust),
        "客数前年比": cust_ratio_str,
        "前年単価": format_currency(tot_pprice),
        "実績単価": format_currency(tot_aprice),
        "単価前年比": price_ratio_str
    }]

    df_summary = pd.DataFrame(summary_data)

    def render_styled_html_table(df, is_summary=False):
        header_bg = "#f39c12" if is_summary else "#f8f9fa"
        header_color = "#ffffff" if is_summary else "#333333"
        border_color = "#e0e0e0"
        
        html = f"""
        <div style="overflow-x: auto; margin-bottom: 20px;">
            <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px; white-space: nowrap;">
                <thead>
                    <tr style="background-color: {header_bg}; color: {header_color}; text-align: center;">
        """
        for col in df.columns:
            html += f'<th style="border: 1px solid {border_color}; padding: 8px; white-space: nowrap; font-weight: bold;">{col}</th>'
        html += "</tr></thead><tbody>"
        
        for idx, row in df.iterrows():
            bg_color = "#fffbf0" if is_summary else ("#fcfcfc" if idx % 2 == 0 else "#ffffff")
            html += f'<tr style="background-color: {bg_color}; text-align: center;">'
            for col in df.columns:
                val = str(row[col])
                style = f"border: 1px solid {border_color}; padding: 8px; white-space: nowrap;"
                
                # --- 条件に応じた文字色と太字スタイルの判定 ---
                if col in ["前年比", "予算比", "客数前年比", "単価前年比"] or "%" in val:
                    if "%" in val:
                        try:
                            num = float(val.replace("%", "").strip())
                            if num >= 100.0:
                                style += " color: #27ae60; font-weight: bold;"  # 100%以上は緑文字太字
                            else:
                                style += " color: #c0392b; font-weight: bold;"  # 100%未満は赤文字太字
                        except ValueError:
                            style += " font-weight: bold;"
                    else:
                        style += " font-weight: bold;"
                
                elif "-" in val and any(char.isdigit() for char in val):
                    style += " color: #c0392b; font-weight: bold;"  # マイナス値は赤文字太字
                
                elif col in ["前年売上", "予算売上", "実績売上", "店頭売上", "デリバリー売上", "その他売上", "前年単価", "実績単価"] or is_summary:
                    style += " font-weight: bold;"

                html += f'<td style="{style}">{val}</td>'
            html += "</tr>"
        html += "</tbody></table></div>"
        return html

    # 赤ピン削除＆店舗名・年月を省略してスッキリした見出しに変更
    st.markdown(f"### 【{max_day}日時点 累計サマリー】")
    st.markdown(render_styled_html_table(df_summary, is_summary=True), unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🔍 表示する時間帯を選択してください")
    
    # --- 2. 日別実績一覧用の時間帯選択タブ ---
    if "active_slot_tab" not in st.session_state:
        st.session_state["active_slot_tab"] = "合計"

    tab_labels = ["合計", "早番", "中番", "遅番"]
    selected_tab = st.radio("時間帯タブ", tab_labels, horizontal=True, label_visibility="collapsed")
    st.session_state["active_slot_tab"] = selected_tab

    slot_kw_map = {"合計": None, "早番": "早番", "中番": "中番", "遅番": "遅番"}
    current_kw = slot_kw_map[selected_tab]
    
    # 選択した時間帯の日別一覧データを生成
    df_current_slot = generate_month_df(selected_store, year, month, current_kw, bp_df, deliv_df)
    df_current_slot_formatted = df_current_slot.copy()
    for col in ["前年売上", "予算売上", "実績売上", "店頭売上", "デリバリー売上", "その他売上", "前年単価", "実績単価"]:
        df_current_slot_formatted[col] = df_current_slot_formatted[col].apply(format_currency)

    st.markdown("### 📅 日別実績一覧")
    st.markdown(render_styled_html_table(df_current_slot_formatted), unsafe_allow_html=True)

    st.markdown("---")
    excel_data = create_multi_sheet_excel(selected_store, year, month, bp_df, deliv_df)
    st.download_button(
        label="📥 全時間帯タブ付きExcelをダウンロード",
        data=excel_data,
        file_name=f"{selected_store}_{year}年{month}月度_売上月報.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )