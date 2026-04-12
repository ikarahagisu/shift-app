import streamlit as st
import pandas as pd
import datetime
import calendar
import io
import math
import random
import re
from ortools.sat.python import cp_model
import jpholiday

# ==========================================
# 重いCSV読み込みを一瞬で終わらせる魔法（キャッシュ機能）
# ==========================================
@st.cache_data
def parse_staff_csv(file_bytes):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), encoding='shift_jis')
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8')
    return df

@st.cache_data
def parse_fixed_csv(file_bytes):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), encoding='shift_jis')
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8')
    if '区分' in df.columns:
        df = df.rename(columns={'区分': '平日/休日'})
    return df

# ==========================================
# カレンダー一括操作用の裏側ロジック
# ==========================================
def set_all_ng(doc_name, y, m, ndays, val):
    for d in range(1, ndays + 1):
        st.session_state[f"ng_{doc_name}_{y}_{m}_{d}"] = val

# ページ設定
st.set_page_config(page_title="シフト作成アプリ", layout="wide")
st.title("当直・日直 自動シフト作成アプリ")

# === スマホ＆フォーム内で絶対に崩れないカレンダー用CSS ===
st.markdown("""
<style>
/* 7列のブロック（カレンダー）をCSS Gridで絶対に7列維持する */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) {
    display: grid !important;
    grid-template-columns: repeat(7, minmax(0, 1fr)) !important;
    gap: 2px !important;
    width: 100% !important; 
    box-sizing: border-box !important;
}

/* カレンダーの各マス（セル）のデザイン */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) > div[data-testid="column"] {
    width: 100% !important;
    min-width: 0 !important; 
    box-sizing: border-box !important; 
    border: 1px solid #eee;
    border-radius: 4px;
    padding: 6px 2px !important; 
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: flex-start !important; 
    background-color: #ffffff;
    overflow: hidden; 
}

/* Streamlit特有の余計なマージンを消去 */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) .element-container {
    margin: 0 !important;
    padding: 0 !important;
    display: flex;
    justify-content: center;
    width: 100%;
}

/* スマホ用に文字サイズを調整し、絶対に改行させない */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) p,
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) label,
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) div[data-testid="stMarkdownContainer"],
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) span,
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) b {
    font-size: 0.8rem !important;
    text-align: center;
    margin: 0 !important;
    white-space: nowrap !important;
    word-break: keep-all !important; 
    line-height: 1.5 !important; 
}

/* プルダウン（Selectbox）を極限までコンパクトに */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) div[data-baseweb="select"] {
    font-size: 0.75rem !important;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) div[data-baseweb="select"] > div {
    padding-top: 0px !important;
    padding-bottom: 0px !important;
    padding-left: 2px !important;
    padding-right: 2px !important;
    min-height: 1.8rem !important;
}
</style>
""", unsafe_allow_html=True)
# ==============================================================================

# ==========================================
# 1. 上部ダッシュボード：年月と休日の設定
# ==========================================
st.header("📅 作成するシフトの設定")
st.info("💡 **【使い方】** 作成したい年・月を選びます。年末年始やお盆など、平日でも日直が必要な日を「休日扱い」にしたい場合はカレンダーのチェックをオンにしてください。GWなどで特定の枠を「2名以上」に増やしたい場合は、下の表で増員設定を行います。")

today = datetime.date.today()
if today.month == 12:
    default_year = today.year + 1
    default_month = 1
else:
    default_year = today.year
    default_month = today.month + 1

col_y, col_m = st.columns(2)
year = col_y.number_input("年", min_value=2000, value=default_year, step=1)
month = col_m.number_input("月", min_value=1, max_value=12, value=default_month, step=1)

st.divider()

st.subheader(f"📅 カレンダー確認 （特別休日の設定） - {month}月")

cal_matrix = calendar.monthcalendar(year, month)
weekdays_ja = ["月", "火", "水", "木", "金", "土", "日"]
custom_holidays = []

# 曜日のヘッダー行
cols = st.columns(7)
for i, w in enumerate(weekdays_ja):
    color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
    cols[i].markdown(f"<div style='color: {color}; font-weight: bold;'>{w}</div>", unsafe_allow_html=True)

# 日付とチェックボックス
for week in cal_matrix:
    cols = st.columns(7)
    for i, day in enumerate(week):
        if day != 0:
            date_obj = datetime.date(year, month, day)
            is_weekend_or_hol = date_obj.weekday() >= 5 or jpholiday.is_holiday(date_obj)
            
            with cols[i]:
                if is_weekend_or_hol:
                    st.markdown(f"<div style='display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 6px; color: #ff4b4b; padding-top: 7px;'><b style='font-weight: 600;'>{day}日</b><div style='height: 1.25rem; display: flex; align-items: center; justify-content: center;'><span style='font-size: 0.8rem;'>休</span></div></div>", unsafe_allow_html=True)
                else:
                    if st.checkbox(f"**{day}日**", key=f"hol_{year}_{month}_{day}"):
                        custom_holidays.append(day)
        else:
            with cols[i]:
                st.write("")

st.divider()

st.subheader("👥 複数人シフト（増員）の設定")

_, num_days = calendar.monthrange(year, month)
NIGHT_SHIFTS_UI = ['宿直A', '宿直B', '外来宿直']
DAY_SHIFTS_UI = ['日直A', '日直B', '外来日直']

date_options = [f"{d}日" for d in range(1, num_days + 1)]
shift_options = NIGHT_SHIFTS_UI + DAY_SHIFTS_UI

multi_df_template = pd.DataFrame(columns=["日付", "シフト枠", "人数"])
edited_multi_df = st.data_editor(
    multi_df_template,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    height=150,
    column_config={
        "日付": st.column_config.SelectboxColumn("日付を選択", options=date_options, required=True),
        "シフト枠": st.column_config.SelectboxColumn("増員する枠を選択", options=shift_options, required=True),
        "人数": st.column_config.NumberColumn("人数を指定", min_value=2, max_value=10, step=1, required=True)
    }
)

multi_slots_dict = {}
for _, row in edited_multi_df.iterrows():
    d_str = str(row.get("日付", ""))
    s_val = str(row.get("シフト枠", ""))
    c_val = row.get("人数")
    
    if d_str and s_val and pd.notna(c_val):
        try:
            d_val = int(re.sub(r'\D', '', d_str))
            c_val = int(c_val)
            multi_slots_dict[(d_val, s_val)] = c_val
        except:
            pass

st.divider()

# ==========================================
# 2. 枠数とカレンダー表示（計算・集計）
# ==========================================
shift_counts = {s: 0 for s in NIGHT_SHIFTS_UI + DAY_SHIFTS_UI}

for d in range(1, num_days + 1):
    date_obj = datetime.date(year, month, d)
    is_hol = jpholiday.is_holiday(date_obj) or date_obj.weekday() >= 5 or (d in custom_holidays)
    
    for s in NIGHT_SHIFTS_UI:
        shift_counts[s] += multi_slots_dict.get((d, s), 1)
        
    if is_hol:
        for s in DAY_SHIFTS_UI:
            shift_counts[s] += multi_slots_dict.get((d, s), 1)

total_slots = sum(shift_counts.values())

st.subheader(f"📌 {year}年{month}月の必要シフト枠数")

st.metric("🌙 宿直A", f"{shift_counts['宿直A']} 枠")
st.metric("☀️ 日直A", f"{shift_counts['日直A']} 枠")
st.metric("🌙 宿直B", f"{shift_counts['宿直B']} 枠")
st.metric("☀️ 日直B", f"{shift_counts['日直B']} 枠")
st.metric("☀️ 外来日直", f"{shift_counts['外来日直']} 枠")
st.metric("🌙 外来宿直", f"{shift_counts['外来宿直']} 枠")
st.metric("🏥 月間 総シフト数", f"{total_slots} 枠")

st.divider()

# ==========================================
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================

st.header("1. 過去・決定済みシフトの読み込み・入力（任意）")
st.info("💡 **【使い方】** 先月末のシフト表をアップロードすれば、月初の間隔（連投禁止）ルールを正確に考慮できます。また、今月のシフトで「すでに人間が確定させた枠」があれば入力してください。AIが残りの空き枠だけを計算して埋めてくれます。")

fixed_columns = ["日付", "平日/休日", "宿直A", "宿直B", "外来宿直", "日直A", "日直B", "外来日直"]
fixed_template_df = pd.DataFrame(columns=fixed_columns)
fixed_csv_template = fixed_template_df.to_csv(index=False).encode('utf-8-sig')

col_dl_fixed, col_ul_fixed = st.columns(2)
with col_dl_fixed:
    st.write("▼ Excelで一括入力したい場合")
    st.download_button(
        label="📥 ひな形（CSV）をダウンロード",
        data=fixed_csv_template,
        file_name="fixed_shift_template.csv",
        mime="text/csv",
    )
with col_ul_fixed:
    fixed_file = st.file_uploader("過去・決定済みシフト表（CSV）をアップロード", type="csv", key="fixed_csv")

if fixed_file is not None:
    try:
        base_fixed_df = parse_fixed_csv(fixed_file.getvalue())
    except Exception as e:
        st.warning(f"過去シフトファイルの読み込みに失敗しました。詳細: {e}")
        base_fixed_df = pd.DataFrame(columns=fixed_columns)
else:
    base_fixed_df = pd.DataFrame(columns=fixed_columns)
    base_fixed_df.loc[0] = ["" for _ in range(len(fixed_columns))]

if "日付" in base_fixed_df.columns:
    base_fixed_df = base_fixed_df.set_index("日付")

st.markdown("##### 📅 決定済みシフトの入力・編集")
st.write("※CSVを使わずに、下の表へ直接クリックして「4/1」のように日付と先生の名前を手打ちすることもできます。")
edited_fixed_df_raw = st.data_editor(base_fixed_df, num_rows="dynamic", use_container_width=True, height=200)

edited_fixed_df = edited_fixed_df_raw.reset_index()

st.divider()

st.header("2. 医師条件の読み込み・入力（必須）")
st.info("""
💡 **【使い方・入力項目の説明】**
まずは「ひな形（CSV）」をダウンロードしてExcelで基本情報を入力・アップロードするのが便利です。

* **入れない曜日**: `水,木` のように入力すると、その曜日は自動的に「宿直なし（日直はあり）」として計算されます。**ただし、翌日が休日の場合は宿直に入る可能性があります（当直明けが休みになるため）。**日直も含めて1日完全に休みたい場合は、下のカレンダーで「全NG」にしてください。
* **NG日**: 下のカレンダーを使って**「全NG」「日NG（日直NG）」「宿NG（宿直NG）」**を直感的に選択できます。
* **希望日**: `10, 15`（日付のみ）や、`10:宿直A`（枠まで指定）で入力します。
* **希望優先度**: 絶対外せない希望がある場合は `100` 以上の数字を入れると、回数上限などのルールを無視して【確実】にそのシフトに入ります。（通常は `1` です）
* **各種ルールについて**:
    * **最低空ける日数**: 勤務と勤務の間を最低何日空けるかを指定します。
    * **月間最小回数 / 月間最大回数**: その月に割り当てる総シフト数の下限と上限です。
    * **休日最大回数**: 土日祝などの「休日扱い」の日に割り当てる最大回数です。
    * **各枠の上限（宿直A上限、日直B上限など）**: 特定のシフト枠ごとに入る最大回数です。
    * ⚠️ **【重要】**: 通常の「希望日（優先度1など）」は、これらのルールを満たす範囲内でのみ叶えられます。ルールと矛盾する希望は反映されないためご注意ください。
* **備考**: 管理用のメモ欄です。「学会のため休み多め」など自由にご記入ください（AIの計算には影響しません）。
""")

template_data = {
    "先生の名前": ["Dr. A", "Dr. B", "Dr. C", "Dr. D", "Dr. E"],
    "入れない曜日(半角カンマ区切り)": ["水,木", "", "土,日", "", ""],
    "NG日(半角カンマ区切り)": ["", "15:日NG", "10:宿NG", "", ""],
    "希望日(半角カンマ区切り)": ["10:宿直A, 15:日直B", "", "8", "20", ""], 
    "希望優先度(数字が大きいほど優先)": [100, 1, 1, 1, 1], 
    "最低空ける日数": [5, 4, 6, 5, 3],  
    "月間最小回数": [1, 2, 0, 1, 0],
    "月間最大回数": [5, 6, 4, 5, 7],
    "休日最大回数": [2, 2, 2, 2, 2],    
    "宿直A上限": [2, 2, 2, 2, 2],
    "宿直B上限": [2, 2, 2, 2, 2],
    "外来宿直上限": [2, 2, 2, 2, 2],
    "日直A上限": [2, 2, 2, 2, 2],
    "日直B上限": [2, 2, 2, 2, 2],
    "外来日直上限": [2, 2, 2, 2, 2],
    "備考（メモ・説明など自由記入）": ["学会のため休み多め", "15日は午後休", "", "当直明け休み希望", ""]
}
df_template = pd.DataFrame(template_data)
csv_template = df_template.to_csv(index=False).encode('shift_jis')

col_dl, col_ul = st.columns(2)
with col_dl:
    st.write("▼ Excelで一括入力したい場合")
    st.download_button(
        label="📥 ひな形（CSV）をダウンロード",
        data=csv_template,
        file_name="shift_template.csv",
        mime="text/csv",
    )
with col_ul:
    uploaded_file = st.file_uploader("医師条件（途中保存CSVも可）をアップロード", type="csv", key="staff_csv")

if uploaded_file is not None:
    if st.session_state.get('last_uploaded_file_id') != uploaded_file.file_id:
        for key in list(st.session_state.keys()):
            if key.startswith("ng_"):
                del st.session_state[key]
        st.session_state['last_uploaded_file_id'] = uploaded_file.file_id
        
    base_df = parse_staff_csv(uploaded_file.getvalue())
else:
    base_df = df_template.copy()

if "先生の名前" in base_df.columns:
    base_df = base_df.set_index("先生の名前")

if "希望優先度(数字が大きいほど優先)" in base_df.columns:
    base_df["希望優先度(数字が大きいほど優先)"] = pd.to_numeric(base_df["希望優先度(数字が大きいほど優先)"], errors='coerce')

text_cols = ["入れない曜日(半角カンマ区切り)", "NG日(半角カンマ区切り)", "希望日(半角カンマ区切り)", "備考（メモ・説明など自由記入）"]
for c in text_cols:
    if c in base_df.columns:
        base_df[c] = base_df[c].apply(lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "none", "<na>"] else str(x))

st.markdown("##### 👩‍⚕️ 医師条件の入力・編集")
st.write("※以下の表は直接クリックして文字を入力できます。（ヘッダーの列名にマウスを合わせるとヒントが出ます）")

edited_df = st.data_editor(
    base_df, 
    num_rows="dynamic", 
    use_container_width=True, 
    height=300,
    column_config={
        "入れない曜日(半角カンマ区切り)": st.column_config.TextColumn(
            "入れない曜日",
            help="例: 水,木 (半角カンマ区切りで入力。カレンダーに⚠️が表示されます)"
        ),
        "NG日(半角カンマ区切り)": None, 
        "希望日(半角カンマ区切り)": st.column_config.TextColumn(
            "希望日",
            help="例: 10, 15 または 10:宿直A (半角カンマ区切りで入力)"
        ),
        "希望優先度(数字が大きいほど優先)": st.column_config.NumberColumn(
            "希望優先度",
            help="数字が大きいほど優先（100以上で絶対希望）"
        ),
        "備考（メモ・説明など自由記入）": st.column_config.TextColumn(
            "備考",
            help="メモ・説明など自由記入"
        )
    }
)

staff_df = edited_df.reset_index()

st.markdown("##### ⚖️ シフト枠と医師の余裕度チェック")

if "月間最大回数" in staff_df.columns:
    total_max_capacity = pd.to_numeric(staff_df["月間最大回数"], errors='coerce').fillna(0).sum()
    total_max_capacity = int(total_max_capacity)
    
    margin = total_max_capacity - total_slots
    
    c1, c2, c3 = st.columns(3)
    c1.metric("🏥 必要な総シフト枠数", f"{total_slots} 枠")
    c2.metric("👩‍⚕️ 医師の月間最大回数の合計", f"{total_max_capacity} 回分")
    
    if margin >= 0:
        c3.metric("✨ 枠の余裕度（バッファ）", f"+{margin} 回分")
    else:
        c3.metric("🚨 枠の余裕度（バッファ）", f"{margin} 回分", delta_color="inverse")
st.divider()

st.markdown("##### 🚫 先生ごとのNG日設定（カレンダーで詳細選択）")
st.info("💡 **【使い方】** カレンダー内のプルダウンから「OK」「全NG」「日NG（日直のみ不可）」「宿NG（宿直のみ不可）」を選べます。選び終わったら、必ず赤い「NG日を確定する」ボタンを押して保存してください。")

valid_staff = staff_df[staff_df["先生の名前"].astype(str).str.strip() != ""]
if not valid_staff.empty:
    doctor_names = valid_staff["先生の名前"].astype(str).tolist()
    tabs = st.tabs(doctor_names)
    
    for t_idx, doc_name in enumerate(doctor_names):
        original_idx = valid_staff.index[t_idx]
        with tabs[t_idx]:
            
            hard_str = str(valid_staff.loc[original_idx].get("入れない曜日(半角カンマ区切り)", ""))
            hard_days = []
            for i, w in enumerate(["月", "火", "水", "木", "金", "土", "日"]):
                if w in hard_str:
                    hard_days.append(i)

            # NG日のパース（全NG, 日NG, 宿NG）
            current_ng_str = str(valid_staff.loc[original_idx].get("NG日(半角カンマ区切り)", ""))
            current_ng_str = current_ng_str.translate(str.maketrans('０１２３４５６７８９，．：', '0123456789,.:'))
            current_ng_dict = {}
            if current_ng_str and current_ng_str.lower() not in ["nan", "none", ""]:
                for x in current_ng_str.split(','):
                    x = x.strip()
                    if not x: continue
                    if ':' in x:
                        parts = x.split(':')
                        try:
                            val = int(float(parts[0].strip()))
                            if 1 <= val <= num_days:
                                current_ng_dict[val] = parts[1].strip()
                        except:
                            pass
                    else:
                        try:
                            val = int(float(x.strip()))
                            if 1 <= val <= num_days:
                                current_ng_dict[val] = "全NG"
                        except:
                            pass
            
            for d in range(1, num_days + 1):
                chk_key = f"ng_{doc_name}_{year}_{month}_{d}"
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = current_ng_dict.get(d, "OK")

            saved_strs = []
            for d in range(1, num_days + 1):
                val = st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", "OK")
                if val == "全NG": saved_strs.append(f"{d}日")
                elif val == "日NG": saved_strs.append(f"{d}日(日直NG)")
                elif val == "宿NG": saved_strs.append(f"{d}日(宿直NG)")
                
            if saved_strs:
                st.success(f"✅ **保存済みのNG日:** {', '.join(saved_strs)}")
            else:
                st.info("💡 **保存済みのNG日はありません**")

            with st.form(key=f"ng_form_{original_idx}", border=False):
                if hard_days:
                    st.markdown("<span style='color: #d97706; font-size: 0.9rem; font-weight: bold;'>💡 設定された「入れない曜日」には日付の横に ⚠️ マークが表示されています（自動で宿直が外れますが、翌日が休みの場合は入る可能性があります）。</span>", unsafe_allow_html=True)

                cols = st.columns(7)
                for i, w in enumerate(weekdays_ja):
                    color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
                    cols[i].markdown(f"<div style='color: {color}; font-weight: bold; text-align: center; padding: 4px;'>{w}</div>", unsafe_allow_html=True)
                
                for week in cal_matrix:
                    cols = st.columns(7)
                    for i, day in enumerate(week):
                        if day != 0:
                            date_obj = datetime.date(year, month, day)
                            is_hol_or_sun = jpholiday.is_holiday(date_obj) or date_obj.weekday() == 6 or (day in custom_holidays)
                            is_sat = date_obj.weekday() == 5 and not is_hol_or_sun
                            is_hard = i in hard_days
                            
                            warning_mark = "⚠️" if is_hard else ""
                            
                            with cols[i]:
                                chk_key = f"ng_{doc_name}_{year}_{month}_{day}"
                                opts = ["OK", "全NG", "日NG", "宿NG"]
                                if st.session_state[chk_key] not in opts:
                                    st.session_state[chk_key] = "OK"
                                idx = opts.index(st.session_state[chk_key])
                                current_ng = st.session_state[chk_key]
                                
                                # 文字色の決定
                                if is_hol_or_sun:
                                    text_color = "#ff4b4b"
                                elif is_sat:
                                    text_color = "#1e90ff"
                                else:
                                    text_color = "inherit"

                                # 背景色の決定
                                if current_ng == "全NG":
                                    bg_color = "#ffe6e6"
                                elif current_ng == "日NG":
                                    bg_color = "#fff0e6"
                                elif current_ng == "宿NG":
                                    bg_color = "#e6f2ff"
                                else:
                                    bg_color = "transparent"

                                day_html = f"<div style='background-color: {bg_color}; color: {text_color}; font-weight: bold; font-size: 0.85rem; margin-bottom: 2px; padding: 2px; border-radius: 4px;'>{day}日 {warning_mark}</div>"
                                
                                st.markdown(day_html, unsafe_allow_html=True)
                                st.selectbox(f"{day}日のNG設定", options=opts, index=idx, key=chk_key, label_visibility="collapsed")
                        else:
                            with cols[i]:
                                st.write("")
                
                submitted = st.form_submit_button(f"✨ {doc_name}先生のNG日を確定する", type="primary")
            
            _, col_btn1, col_btn2 = st.columns([6, 1.5, 1.5])
            with col_btn1:
                st.button("全選択(全NG)", key=f"btn_all_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, "全NG"), use_container_width=True)
            with col_btn2:
                st.button("全解除(OK)", key=f"btn_clear_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, "OK"), use_container_width=True)
            
            # DataFrameへ状態を保存
            ng_items = []
            for d in range(1, num_days + 1):
                val = st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", "OK")
                if val == "全NG":
                    ng_items.append(str(d))
                elif val != "OK":
                    ng_items.append(f"{d}:{val}")
                    
            staff_df.at[original_idx, "NG日(半角カンマ区切り)"] = ",".join(ng_items)
            
            if submitted:
                st.toast(f"✅ {doc_name}先生のNG日を保存しました！")

st.divider()
st.markdown("##### 📂 入力途中のデータを一時保存（後で再開したい場合）")
st.write("※途中で入力をやめる場合は、ここまでのデータを保存しておき、次回アップロードすることで続きから再開できます。")

current_csv = staff_df.to_csv(index=False).encode('utf-8-sig')
st.download_button(
    label="📥 現在の医師条件を一時保存する（CSVダウンロード）",
    data=current_csv,
    file_name=f"staff_wip_{year}_{month}.csv",
    mime="text/csv",
    use_container_width=True
)

st.divider()

# ==========================================
# 4. シフト計算ロジック（関数）
# ==========================================
def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None):
    _, num_days = calendar.monthrange(target_year, target_month)
    NIGHT_SHIFTS = ['宿直A', '宿直B', '外来宿直']
    DAY_SHIFTS = ['日直A', '日直B', '外来日直']

    def is_holiday(y, m, d):
        date = datetime.date(y, m, d)
        return date.weekday() >= 5 or jpholiday.is_holiday(date) or (d in custom_holidays)
    
    def safe_int(val, default_val):
        if pd.isna(val): return default_val
        try:
            return int(float(val))
        except:
            return default_val

    doctors = staff_df['先生の名前'].astype(str).tolist()
    ng_days_dict = {}
    req_days = {}          
    req_specific = {}      
    req_priority = {} 
    
    # 入れない曜日の保存用
    hard_weekdays = {}
    
    min_intervals = {}
    min_shifts_total = {}
    max_shifts_total = {}
    max_hol_shifts_per_doc = {}
    max_shifts_per_type = {}
    
    absolute_req_days = {doc: [] for doc in doctors}
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}
    future_worked_dates = {doc: [] for doc in doctors}

    invalid_requests = []

    if fixed_df is not None:
        for _, row in fixed_df.iterrows():
            date_str = str(row.get('日付', ''))
            match = re.search(r'(\d+)\s*[/月\-]\s*(\d+)', date_str)
            if match:
                m = int(match.group(1))
                d = int(match.group(2))
                
                if m == target_month:
                    y = target_year
                elif m == 12 and target_month == 1:
                    y = target_year - 1
                elif m == 1 and target_month == 12:
                    y = target_year + 1
                elif m > target_month and (m - target_month) >= 6: 
                    y = target_year - 1
                elif m < target_month and (target_month - m) >= 6:
                    y = target_year + 1
                else:
                    y = target_year
                
                try:
                    date_obj = datetime.date(y, m, d)
                except ValueError:
                    continue 
                    
                for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
                    if s_type in row and pd.notna(row[s_type]):
                        doc_vals = re.split(r'[、,\s]+', str(row[s_type]))
                        for doc_val in doc_vals:
                            doc_val = doc_val.strip()
                            if doc_val in doctors:
                                if m == target_month:
                                    absolute_req_specific[doc_val].append((d, s_type))
                                elif date_obj < datetime.date(target_year, target_month, 1):
                                    past_worked_dates[doc_val].append(date_obj)
                                else:
                                    future_worked_dates[doc_val].append(date_obj)

    for index, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        
        # 入れない曜日を数値のリストとして取得（月=0, ..., 日=6）
        hard_str = str(row.get('入れない曜日(半角カンマ区切り)', ''))
        hard_days_list = []
        for i, w in enumerate(["月", "火", "水", "木", "金", "土", "日"]):
            if w in hard_str:
                hard_days_list.append(i)
        hard_weekdays[doc] = hard_days_list
        
        # NG日をパース
        ng_str = str(row['NG日(半角カンマ区切り)'])
        ng_dict = {}
        if not pd.isna(row['NG日(半角カンマ区切り)']) and ng_str.strip() != "" and ng_str.lower() not in ["nan", "none"]:
            ng_str = ng_str.translate(str.maketrans('０１２３４５６７８９，．：', '0123456789,.:'))
            for x in ng_str.split(','):
                x = x.strip()
                if not x: continue
                if ':' in x:
                    parts = x.split(':')
                    try:
                        d_val = int(float(parts[0].strip()))
                        ng_dict[d_val] = parts[1].strip()
                    except: pass
                else:
                    try:
                        d_val = int(float(x.strip()))
                        ng_dict[d_val] = "全NG"
                    except: pass
        ng_days_dict[doc] = ng_dict
                
        req_days[doc] = []
        req_specific[doc] = []
        if '希望日(半角カンマ区切り)' in staff_df.columns:
            req_str = str(row['希望日(半角カンマ区切り)'])
            if not (pd.isna(row['希望日(半角カンマ区切り)']) or req_str.strip() == "" or req_str.lower() in ["nan", "none"]):
                req_str = req_str.replace('：', ':')
                items = req_str.split(',')
                for item in items:
                    item = item.strip()
                    if not item: continue
                    if ':' in item:
                        parts = item.split(':')
                        try:
                            d = int(parts[0].strip())
                            s_name = parts[1].strip()
                            req_specific[doc].append((d, s_name))
                        except: pass
                    else:
                        try:
                            req_days[doc].append(int(item))
                        except: pass

        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('最低空ける日数'), 5)
        min_shifts_total[doc] = safe_int(row.get('月間最小回数'), 0)
        max_shifts_total[doc] = safe_int(row.get('月間最大回数'), 5)
        max_hol_shifts_per_doc[doc] = safe_int(row.get('休日最大回数'), 4)

        max_shifts_per_type[doc] = {
            '宿直A': safe_int(row.get('宿直A上限'), 2),
            '宿直B': safe_int(row.get('宿直B上限'), 2),
            '外来宿直': safe_int(row.get('外来宿直上限'), 2),
            '日直A': safe_int(row.get('日直A上限'), 2),
            '日直B': safe_int(row.get('日直B上限'), 2),
            '外来日直': safe_int(row.get('外来日直上限'), 2)
        }
    
    for doc in doctors:
        for d in req_days[doc]:
            if not (1 <= d <= num_days):
                invalid_requests.append(f"❌ **{doc}先生**: {target_month}月にはない日付（{d}日）が希望日に指定されています。")
                
        for d, s_name in req_specific[doc]:
            if not (1 <= d <= num_days):
                invalid_requests.append(f"❌ **{doc}先生**: {target_month}月にはない日付（{d}日）が希望日に指定されています。")

    for doc in doctors:
        if req_priority[doc] >= 100:
            absolute_req_days[doc].extend([d for d in req_days[doc] if 1 <= d <= num_days])
            absolute_req_specific[doc].extend([(d, s) for (d, s) in req_specific[doc] if 1 <= d <= num_days])
            
        all_abs_dates = absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]]
        ng_days_dict[doc] = {d: v for d, v in ng_days_dict[doc].items() if d not in all_abs_dates}

    daily_active_shifts = {}
    for d in range(1, num_days + 1):
        base_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        forced_shifts = [s for doc in doctors for sd, s in absolute_req_specific[doc] if sd == d and s in (NIGHT_SHIFTS + DAY_SHIFTS)]
        daily_active_shifts[d] = list(set(base_shifts + forced_shifts))

    if invalid_requests:
        unique_invalid = list(dict.fromkeys(invalid_requests))
        return None, False, unique_invalid, None, None

    model = cp_model.CpModel()
    shifts = {}
    objective_terms = []

    for d in range(1, num_days + 1):
        for doc in doctors:
            for s in daily_active_shifts[d]:
                shifts[(d, doc, s)] = model.NewBoolVar(f'shift_d{d}_{doc}_{s}')

    over_caps = {}
    for d in range(1, num_days + 1):
        for s in daily_active_shifts[d]:
            over_caps[(d, s)] = model.NewIntVar(0, len(doctors), f'over_cap_d{d}_{s}')
            req_count = multi_slots_dict.get((d, s), 1)
            fixed_docs_count = sum(1 for doc in doctors if (d, s) in absolute_req_specific[doc])
            actual_req_count = max(req_count, fixed_docs_count)
            
            model.Add(sum(shifts[(d, doc, s)] for doc in doctors) == actual_req_count + over_caps[(d, s)])
            objective_terms.append(over_caps[(d, s)] * -50000)

    for doc in doctors:
        for d in range(1, num_days + 1):
            fixed_count = sum(1 for sd, ss in absolute_req_specific[doc] if sd == d and ss in daily_active_shifts[d])
            max_shifts_today = max(1, fixed_count)
            model.Add(sum(shifts[(d, doc, s)] for s in daily_active_shifts[d]) <= max_shifts_today)

    for doc in doctors:
        for d, ng_type in ng_days_dict[doc].items():
            if 1 <= d <= num_days:
                if ng_type == "全NG":
                    for s in daily_active_shifts[d]:
                        model.Add(shifts[(d, doc, s)] == 0)
                elif ng_type == "日NG":
                    for s in daily_active_shifts[d]:
                        if s in DAY_SHIFTS:
                            model.Add(shifts[(d, doc, s)] == 0)
                elif ng_type == "宿NG":
                    for s in daily_active_shifts[d]:
                        if s in NIGHT_SHIFTS:
                            model.Add(shifts[(d, doc, s)] == 0)

    # 入れない曜日は「宿直系」のみNG。ただし【翌日が休日】の場合はOKとする
    for doc in doctors:
        for d in range(1, num_days + 1):
            date_obj = datetime.date(target_year, target_month, d)
            next_date = date_obj + datetime.timedelta(days=1)
            
            next_is_hol = next_date.weekday() >= 5 or jpholiday.is_holiday(next_date)
            if next_date.year == target_year and next_date.month == target_month:
                if next_date.day in custom_holidays:
                    next_is_hol = True
                    
            if date_obj.weekday() in hard_weekdays[doc] and not next_is_hol:
                for s in NIGHT_SHIFTS:
                    if s in daily_active_shifts[d]:
                        model.Add(shifts[(d, doc, s)] == 0)

    for doc in doctors:
        for d in absolute_req_days[doc]:
            specifics_on_d = [s for sd, s in absolute_req_specific[doc] if sd == d]
            if not specifics_on_d:
                model.AddExactlyOne(shifts[(d, doc, s)] for s in daily_active_shifts[d])
            
        for d, s_name in absolute_req_specific[doc]:
            if s_name in daily_active_shifts[d]:
                model.Add(shifts[(d, doc, s_name)] == 1)

    for doc in doctors:
        for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
            worked = [shifts[(d, doc, s_type)] for d in range(1, num_days + 1) if s_type in daily_active_shifts[d]]
            if worked:
                specific_req_count = sum(1 for d, s in absolute_req_specific[doc] if s == s_type)
                actual_max_type = max(max_shifts_per_type[doc][s_type], specific_req_count + len(absolute_req_days[doc]))
                model.Add(sum(worked) <= actual_max_type)

    min_shortfalls = {}
    for doc in doctors:
        min_shortfalls[doc] = model.NewIntVar(0, num_days, f'min_shortfall_{doc}')
        worked_all = []
        for d in range(1, num_days + 1):
            for s in daily_active_shifts[d]:
                worked_all.append(shifts[(d, doc, s)])
        if worked_all:
            all_abs_dates = absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]]
            actual_max_total = max(max_shifts_total[doc], len(all_abs_dates))
            actual_min_total = min(min_shifts_total[doc], actual_max_total)
            
            model.Add(sum(worked_all) <= actual_max_total)
            model.Add(sum(worked_all) + min_shortfalls[doc] >= actual_min_total)
            objective_terms.append(min_shortfalls[doc] * -10000)

    for doc in doctors:
        interval = min_intervals[doc]
        if interval > 0:
            all_abs_dates = set(absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]])
            
            for d in range(1, num_days + 1):
                if d in all_abs_dates:
                    continue
                current_date = datetime.date(target_year, target_month, d)
                for past_date in past_worked_dates[doc]:
                    if 0 < (current_date - past_date).days <= interval:
                        for s in daily_active_shifts[d]:
                            model.Add(shifts[(d, doc, s)] == 0)
                            
                for future_date in future_worked_dates[doc]:
                    if 0 < (future_date - current_date).days <= interval:
                        for s in daily_active_shifts[d]:
                            model.Add(shifts[(d, doc, s)] == 0)
            
            for d1 in range(1, num_days + 1):
                for d2 in range(d1 + 1, min(d1 + interval + 1, num_days + 1)):
                    if d1 in all_abs_dates or d2 in all_abs_dates:
                        continue
                    for s1 in daily_active_shifts[d1]:
                        for s2 in daily_active_shifts[d2]:
                            model.Add(shifts[(d1, doc, s1)] + shifts[(d2, doc, s2)] <= 1)

    holiday_worked = {}
    for doc in doctors:
        hol_shifts = []
        for d in range(1, num_days + 1):
            if is_holiday(target_year, target_month, d):
                for s in daily_active_shifts[d]:
                    if s in NIGHT_SHIFTS + DAY_SHIFTS:
                        hol_shifts.append(shifts[(d, doc, s)])
        holiday_worked[doc] = sum(hol_shifts)
        
        abs_hol_count = sum(1 for d in absolute_req_days[doc] if is_holiday(target_year, target_month, d))
        abs_hol_count += sum(1 for d, s in absolute_req_specific[doc] if is_holiday(target_year, target_month, d))
        actual_hol_max = max(max_hol_shifts_per_doc[doc], abs_hol_count) 
        model.Add(holiday_worked[doc] <= actual_hol_max)
        
    global_max = num_days * 3 
    max_hol_shifts = model.NewIntVar(0, global_max, 'max_hol_shifts')
    for doc in doctors:
        model.Add(holiday_worked[doc] <= max_hol_shifts)
        
    for doc in doctors:
        if req_priority[doc] < 100:  
            weight = req_priority[doc] * 100 
            for d in req_days[doc]:
                if 1 <= d <= num_days:
                    for s in daily_active_shifts[d]:
                        if (d, doc, s) in shifts:
                            objective_terms.append(shifts[(d, doc, s)] * weight)
            for d, s_name in req_specific[doc]:
                if 1 <= d <= num_days:
                    if s_name in daily_active_shifts[d]:
                        if (d, doc, s_name) in shifts:
                            objective_terms.append(shifts[(d, doc, s_name)] * weight)
                    
    if objective_terms:
        model.Maximize(sum(objective_terms) - max_hol_shifts * 1000)
    else:
        model.Minimize(max_hol_shifts)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    solver.parameters.random_seed = random.randint(1, 10000)
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        schedule_list = []
        weekday_ja = ["月", "火", "水", "木", "金", "土", "日"]
        over_cap_warnings = []
        for d in range(1, num_days + 1):
            date_obj = datetime.date(target_year, target_month, d)
            day_str = "休日" if is_holiday(target_year, target_month, d) else "平日"
            row = {"日付": f"{target_month}/{d}({weekday_ja[date_obj.weekday()]})", "平日/休日": day_str}
            for s in NIGHT_SHIFTS + DAY_SHIFTS:
                row[s] = "-"
            for s in daily_active_shifts[d]:
                assigned_docs = []
                for doc in doctors:
                    if solver.Value(shifts[(d, doc, s)]) == 1:
                        assigned_docs.append(doc)
                if assigned_docs:
                    row[s] = "、".join(assigned_docs)
                if solver.Value(over_caps[(d, s)]) > 0:
                    over_cap_warnings.append(f"{target_month}/{d}({weekday_ja[date_obj.weekday()]}) の「{s}」枠")
            schedule_list.append(row)
        warnings = []
        if over_cap_warnings:
            warnings.append("⚠️ **【重要】以下の枠は「決定済みシフト」や「優先度100」が重なったため、AIが自動的に定員を拡張（2名以上配置）してシフトを完成させました:**")
            warnings.extend([f"・{w}" for w in over_cap_warnings])
        return pd.DataFrame(schedule_list), True, warnings, past_worked_dates, future_worked_dates
    else:
        return None, False, ["入力された条件が厳しすぎて、シフトを組むことができませんでした。"], None, None

# ==========================================
# 5. 実行ボタンと結果表示
# ==========================================
st.divider()
st.header("3. シフトの自動生成")
staff_df = staff_df[staff_df['先生の名前'].astype(str).str.strip() != ''].dropna(subset=['先生の名前']).reset_index(drop=True)
fixed_df = edited_fixed_df[edited_fixed_df['日付'].astype(str).str.strip() != ''].dropna(subset=['日付']).reset_index(drop=True)

if len(staff_df) > 0:
    if st.button("🚀 このデータでシフトを自動生成する", type="primary"):
        with st.spinner("AIが最適なシフトを計算中..."):
            try:
                df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df)
                if success:
                    st.session_state['generated_df'] = df_result
                    st.session_state['past_worked_dates'] = past_worked_dates
                    st.session_state['future_worked_dates'] = future_worked_dates
                    st.success("✨ シフトの作成に成功しました！")
                else:
                    st.error(error_reasons[0])
            except Exception as e:
                st.error(f"シフト計算中にエラーが発生しました。詳細: {e}")

    if 'generated_df' in st.session_state:
        df_result = st.session_state['generated_df']
        doctors_list = staff_df['先生の名前'].astype(str).tolist()
        st.subheader("📅 完成したシフト表")
        
        table_container = st.container()
        st.divider()
        st.markdown("<span style='font-size: 0.95rem; font-weight: bold;'>🔍 特定の医師のシフトを色別でハイライト</span>", unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("🟨 **黄色**")
            hl_yellow = st.multiselect("黄色", options=doctors_list, default=[], key="hl_yellow", label_visibility="collapsed")
        with c2:
            st.markdown("🟥 **赤色**")
            hl_red = st.multiselect("赤色", options=doctors_list, default=[], key="hl_red", label_visibility="collapsed")

        shift_columns = ['宿直A', '宿直B', '外来宿直', '日直A', '日直B', '外来日直']

        def highlight_holidays(row):
            styles = [''] * len(row)
            if row['平日/休日'] == '休日':
                for i, col in enumerate(row.index):
                    if col in ['日付', '平日/休日']: styles[i] = 'color: #ff4b4b; font-weight: bold;'
            return styles
        
        def color_highlighted_doctor(val):
            val_str = str(val)
            if val_str == "-" or val_str == "": return ''
            cell_docs = [d.strip() for d in re.split(r'[、,]', val_str)]
            for doc in cell_docs:
                if doc in hl_yellow: return 'background-color: #fff200; color: #000000; font-weight: bold; border: 2px solid #ffcc00;'
                elif doc in hl_red: return 'background-color: #ffcccc; color: #000000; font-weight: bold; border: 2px solid #ff6666;'
            return ''
        
        base_style = df_result.style.apply(highlight_holidays, axis=1)
        
        # ▼▼▼ 修正箇所：Pandasのバージョンに対応する安全な呼び出し ▼▼▼
        if hasattr(base_style, 'map'):
            styled_df = base_style.map(color_highlighted_doctor, subset=shift_columns)
        else:
            styled_df = base_style.applymap(color_highlighted_doctor, subset=shift_columns)
        # ▲▲▲ 修正箇所ここまで ▲▲▲
        
        with table_container:
            st.dataframe(styled_df, use_container_width=True, hide_index=True, height=(len(df_result) * 35 + 40))
        
        csv_result = df_result.to_csv(index=False).encode('utf-8-sig')
        st.download_button(label="📥 完成したシフト表をCSVでダウンロード", data=csv_result, file_name=f"shift_{year}_{month}_result.csv", mime="text/csv")

elif len(staff_df) == 0:
    st.warning("☝️ 表に先生の名前を入力するか、CSVファイルをアップロードしてください。")
