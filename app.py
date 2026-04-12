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
    gap: 4px !important;
    width: 100% !important; 
    box-sizing: border-box !important;
}

/* カレンダーの各マス（セル）のデザイン */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) > div[data-testid="column"] {
    width: 100% !important;
    min-width: 0 !important; 
    box-sizing: border-box !important; 
    border: 1px solid #ddd;
    border-radius: 6px;
    padding: 0px !important; 
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

/* スマホ用に文字サイズを調整 */
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
}

/* プルダウン（Selectbox）をコンパクトに */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) div[data-baseweb="select"] {
    font-size: 0.75rem !important;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) div[data-baseweb="select"] > div {
    min-height: 1.8rem !important;
    border-radius: 4px !important;
}
</style>
""", unsafe_allow_html=True)
# ==============================================================================

# ==========================================
# 1. 上部ダッシュボード：年月と休日の設定
# ==========================================
st.header("📅 作成するシフトの設定")
st.info("💡 **【使い方】** 作成したい年・月を選びます。年末年始やお盆など、平日でも日直が必要な日を「休日扱い」にしたい場合はカレンダーのチェックをオンにしてください。")

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

cols = st.columns(7)
for i, w in enumerate(weekdays_ja):
    color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
    cols[i].markdown(f"<div style='color: {color}; font-weight: bold;'>{w}</div>", unsafe_allow_html=True)

for week in cal_matrix:
    cols = st.columns(7)
    for i, day in enumerate(week):
        if day != 0:
            date_obj = datetime.date(year, month, day)
            is_weekend_or_hol = date_obj.weekday() >= 5 or jpholiday.is_holiday(date_obj)
            with cols[i]:
                if is_weekend_or_hol:
                    st.markdown(f"<div style='color: #ff4b4b; padding-top: 7px;'><b style='font-weight: 600;'>{day}日</b><br><span style='font-size: 0.7rem;'>休</span></div>", unsafe_allow_html=True)
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
            multi_slots_dict[(d_val, s_val)] = int(c_val)
        except: pass

st.divider()

# ==========================================
# 2. 枠数集計
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
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("🌙 宿A", f"{shift_counts['宿直A']}")
c2.metric("🌙 宿B", f"{shift_counts['宿直B']}")
c3.metric("☀️ 日A", f"{shift_counts['日直A']}")
c4.metric("☀️ 日B", f"{shift_counts['日直B']}")
c5.metric("☀️ 外来日", f"{shift_counts['外来日直']}")
c6.metric("🌙 外来宿", f"{shift_counts['外来宿直']}")

st.divider()

# ==========================================
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================
st.header("1. 過去・決定済みシフトの入力（任意）")
fixed_columns = ["日付", "平日/休日", "宿直A", "宿直B", "外来宿直", "日直A", "日直B", "外来日直"]
fixed_file = st.file_uploader("過去・決定済みシフト表（CSV）をアップロード", type="csv", key="fixed_csv")

if fixed_file is not None:
    base_fixed_df = parse_fixed_csv(fixed_file.getvalue())
else:
    base_fixed_df = pd.DataFrame(columns=fixed_columns)
    base_fixed_df.loc[0] = ["" for _ in range(len(fixed_columns))]

if "日付" in base_fixed_df.columns:
    base_fixed_df = base_fixed_df.set_index("日付")

edited_fixed_df_raw = st.data_editor(base_fixed_df, num_rows="dynamic", use_container_width=True, height=200)
edited_fixed_df = edited_fixed_df_raw.reset_index()

st.divider()

st.header("2. 医師条件の読み込み・入力（必須）")
st.info("""
* **入りにくい曜日**: `水,木` のように入力すると自動的に「宿直なし（日直はあり）」となります（翌日が休日の場合は入る可能性があります）。
* **NG日**: 下のカレンダーから「全NG」「日NG」「宿NG」を選んで保存してください。
""")

uploaded_file = st.file_uploader("医師条件CSVをアップロード", type="csv", key="staff_csv")

template_data = {
    "先生の名前": ["Dr. A", "Dr. B", "Dr. C"],
    "入りにくい曜日(半角カンマ区切り)": ["水,木", "", ""],
    "NG日(半角カンマ区切り)": ["", "", ""],
    "希望日(半角カンマ区切り)": ["", "", ""], 
    "希望優先度(数字が大きいほど優先)": [1, 1, 1], 
    "最低空ける日数": [5, 5, 5],  
    "月間最小回数": [0, 0, 0],
    "月間最大回数": [5, 5, 5],
    "休日最大回数": [2, 2, 2],    
    "宿直A上限": [2, 2, 2], "宿直B上限": [2, 2, 2], "外来宿直上限": [2, 2, 2],
    "日直A上限": [2, 2, 2], "日直B上限": [2, 2, 2], "外来日直上限": [2, 2, 2],
    "備考（メモ・説明など自由記入）": ["", "", ""]
}

if uploaded_file is not None:
    if st.session_state.get('last_uploaded_file_id') != uploaded_file.file_id:
        for key in list(st.session_state.keys()):
            if key.startswith("ng_"): del st.session_state[key]
        st.session_state['last_uploaded_file_id'] = uploaded_file.file_id
    base_df = parse_staff_csv(uploaded_file.getvalue())
else:
    base_df = pd.DataFrame(template_data)

if "先生の名前" in base_df.columns:
    base_df = base_df.set_index("先生の名前")

text_cols = ["入りにくい曜日(半角カンマ区切り)", "NG日(半角カンマ区切り)", "希望日(半角カンマ区切り)", "備考（メモ・説明など自由記入）"]
for c in text_cols:
    if c in base_df.columns:
        base_df[c] = base_df[c].apply(lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "none", "<na>"] else str(x))

edited_df = st.data_editor(base_df, num_rows="dynamic", use_container_width=True, height=250)
staff_df = edited_df.reset_index()

st.divider()

st.markdown("##### 🚫 先生ごとのNG日設定")

# カレンダーの視認性向上のためのガイド
st.markdown("""
<div style='display: flex; gap: 15px; margin-bottom: 10px; font-size: 0.85rem;'>
    <div style='display: flex; align-items: center; gap: 5px;'><div style='width: 15px; height: 15px; background: #ffffff; border: 1px solid #ddd; border-radius: 3px;'></div>OK</div>
    <div style='display: flex; align-items: center; gap: 5px;'><div style='width: 15px; height: 15px; background: #ffcccc; border-radius: 3px;'></div><b>全NG</b></div>
    <div style='display: flex; align-items: center; gap: 5px;'><div style='width: 15px; height: 15px; background: #fff2cc; border-radius: 3px;'></div><b>日NG</b></div>
    <div style='display: flex; align-items: center; gap: 5px;'><div style='width: 15px; height: 15px; background: #d9e9ff; border-radius: 3px;'></div><b>宿NG</b></div>
</div>
""", unsafe_allow_html=True)

valid_staff = staff_df[staff_df["先生の名前"].astype(str).str.strip() != ""]
if not valid_staff.empty:
    doctor_names = valid_staff["先生の名前"].astype(str).tolist()
    tabs = st.tabs(doctor_names)
    
    for t_idx, doc_name in enumerate(doctor_names):
        original_idx = valid_staff.index[t_idx]
        with tabs[t_idx]:
            hard_str = str(valid_staff.loc[original_idx].get("入りにくい曜日(半角カンマ区切り)", ""))
            hard_days = [i for i, w in enumerate(["月", "火", "水", "木", "金", "土", "日"]) if w in hard_str]

            current_ng_str = str(valid_staff.loc[original_idx].get("NG日(半角カンマ区切り)", ""))
            current_ng_str = current_ng_str.translate(str.maketrans('０１２３４５６７８９，．：', '0123456789,.:'))
            current_ng_dict = {}
            if current_ng_str and current_ng_str.lower() not in ["nan", "none"]:
                for x in current_ng_str.split(','):
                    x = x.strip()
                    if not x: continue
                    if ':' in x:
                        parts = x.split(':')
                        try:
                            d_val = int(float(parts[0].strip()))
                            current_ng_dict[d_val] = parts[1].strip()
                        except: pass
                    else:
                        try:
                            d_val = int(float(x.strip()))
                            current_ng_dict[d_val] = "全NG"
                        except: pass
            
            for d in range(1, num_days + 1):
                chk_key = f"ng_{doc_name}_{year}_{month}_{d}"
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = current_ng_dict.get(d, "OK")

            with st.form(key=f"ng_form_{original_idx}", border=False):
                for week in cal_matrix:
                    cols = st.columns(7)
                    for i, day in enumerate(week):
                        if day != 0:
                            date_obj = datetime.date(year, month, day)
                            is_hol_or_sun = jpholiday.is_holiday(date_obj) or date_obj.weekday() == 6 or (day in custom_holidays)
                            is_sat = date_obj.weekday() == 5 and not is_hol_or_sun
                            is_hard = i in hard_days
                            chk_key = f"ng_{doc_name}_{year}_{month}_{day}"
                            
                            current_val = st.session_state.get(chk_key, "OK")
                            bg_color = "#ffffff"
                            if current_val == "全NG": bg_color = "#ffcccc"
                            elif current_val == "日NG": bg_color = "#fff2cc"
                            elif current_val == "宿NG": bg_color = "#d9e9ff"
                            
                            with cols[i]:
                                st.markdown(f"""
                                <div style="background-color: {bg_color}; border-radius: 4px; padding: 6px 2px; width: 100%;">
                                    <div style="font-weight: bold; font-size: 0.85rem; color: {'#ff4b4b' if is_hol_or_sun else '#1e90ff' if is_sat else '#333'};">
                                        {day}日 {'⚠️' if is_hard else ''}
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                opts = ["OK", "全NG", "日NG", "宿NG"]
                                idx = opts.index(current_val) if current_val in opts else 0
                                st.selectbox(f"{day}NG", options=opts, index=idx, key=chk_key, label_visibility="collapsed")
                        else:
                            with cols[i]: st.write("")
                
                if st.form_submit_button(f"✨ {doc_name}先生のNG日を確定する", type="primary"):
                    st.toast(f"✅ {doc_name}先生の条件を保存しました！")

            # 保存用文字列作成
            ng_items = []
            for d in range(1, num_days + 1):
                val = st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", "OK")
                if val == "全NG": ng_items.append(str(d))
                elif val != "OK": ng_items.append(f"{d}:{val}")
            staff_df.at[original_idx, "NG日(半角カンマ区切り)"] = ",".join(ng_items)

st.divider()
st.markdown("##### 📂 入力途中のデータを一時保存")
st.write("※途中で入力をやめる場合は、ここまでのデータを保存しておき、次回アップロードすることで続きから再開できます。")

current_csv = staff_df.to_csv(index=False).encode('utf-8-sig')
st.download_button(
    label="📥 現在の医師条件を一時保存する（CSVダウンロード）",
    data=current_csv,
    file_name=f"staff_wip_{year}_{month}.csv",
    mime="text/csv",
    use_container_width=True
)
# ==========================================
# 4. シフト計算ロジック
# ==========================================
def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None):
    _, num_days = calendar.monthrange(target_year, target_month)
    NIGHT_SHIFTS = ['宿直A', '宿直B', '外来宿直']
    DAY_SHIFTS = ['日直A', '日直B', '外来日直']

    def is_holiday(y, m, d):
        dt = datetime.date(y, m, d)
        return dt.weekday() >= 5 or jpholiday.is_holiday(dt) or (d in custom_holidays)

    doctors = staff_df['先生の名前'].astype(str).tolist()
    ng_days_dict = {}
    req_days = {doc: [] for doc in doctors}
    req_specific = {doc: [] for doc in doctors}
    req_priority = {}
    hard_weekdays = {}
    min_intervals = {}
    max_shifts_total = {}
    max_hol_shifts_per_doc = {}
    max_shifts_per_type = {doc: {} for doc in doctors}
    
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}
    future_worked_dates = {doc: [] for doc in doctors}

    # 固定シフトのパース
    if fixed_df is not None:
        for _, row in fixed_df.iterrows():
            date_str = str(row.get('日付', ''))
            match = re.search(r'(\d+)\s*[/月\-]\s*(\d+)', date_str)
            if match:
                m, d = int(match.group(1)), int(match.group(2))
                y = target_year if m == target_month else (target_year-1 if m > target_month else target_year)
                try: dt_obj = datetime.date(y, m, d)
                except: continue
                for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
                    if s_type in row and pd.notna(row[s_type]):
                        for dv in re.split(r'[、,\s]+', str(row[s_type])):
                            dv = dv.strip()
                            if dv in doctors:
                                if m == target_month: absolute_req_specific[dv].append((d, s_type))
                                elif dt_obj < datetime.date(target_year, target_month, 1): past_worked_dates[dv].append(dt_obj)
                                else: future_worked_dates[dv].append(dt_obj)

    for _, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        h_str = str(row.get('入りにくい曜日(半角カンマ区切り)', ''))
        hard_weekdays[doc] = [i for i, w in enumerate(["月", "火", "水", "木", "金", "土", "日"]) if w in h_str]
        
        ng_str = str(row['NG日(半角カンマ区切り)']).translate(str.maketrans('：', ':'))
        ng_dict = {}
        for x in ng_str.split(','):
            x = x.strip()
            if not x: continue
            if ':' in x:
                p = x.split(':')
                try: ng_dict[int(float(p[0]))] = p[1]
                except: pass
            else:
                try: ng_dict[int(float(x))] = "全NG"
                except: pass
        ng_days_dict[doc] = ng_dict

        # 希望日
        r_str = str(row.get('希望日(半角カンマ区切り)', '')).replace('：', ':')
        for item in r_str.split(','):
            item = item.strip()
            if not item: continue
            if ':' in item:
                p = item.split(':')
                try: req_specific[doc].append((int(p[0]), p[1]))
                except: pass
            else:
                try: req_days[doc].append(int(item))
                except: pass

        def s_int(v, d):
            try: return int(float(v))
            except: return d

        req_priority[doc] = s_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = s_int(row.get('最低空ける日数'), 5)
        max_shifts_total[doc] = s_int(row.get('月間最大回数'), 5)
        max_hol_shifts_per_doc[doc] = s_int(row.get('休日最大回数'), 2)
        for s in NIGHT_SHIFTS + DAY_SHIFTS:
            max_shifts_per_type[doc][s] = s_int(row.get(f'{s}上限'), 2)

    model = cp_model.CpModel()
    shifts = {}
    
    daily_active_shifts = {}
    for d in range(1, num_days + 1):
        base = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        forced = [s for doc in doctors for sd, s in absolute_req_specific[doc] if sd == d]
        daily_active_shifts[d] = list(set(base + forced))

    for d in range(1, num_days + 1):
        for doc in doctors:
            for s in daily_active_shifts[d]:
                shifts[(d, doc, s)] = model.NewBoolVar(f's_{d}_{doc}_{s}')

    over_caps = {}
    objective_terms = []
    for d in range(1, num_days + 1):
        for s in daily_active_shifts[d]:
            over_caps[(d, s)] = model.NewIntVar(0, len(doctors), f'o_{d}_{s}')
            req = multi_slots_dict.get((d, s), 1)
            fixed_cnt = sum(1 for doc in doctors if (d, s) in absolute_req_specific[doc])
            model.Add(sum(shifts[(d, doc, s)] for doc in doctors) == max(req, fixed_cnt) + over_caps[(d, s)])
            objective_terms.append(over_caps[(d, s)] * -50000)

    for doc in doctors:
        for d in range(1, num_days + 1):
            fixed_on_d = sum(1 for sd, ss in absolute_req_specific[doc] if sd == d)
            model.Add(sum(shifts[(d, doc, s)] for s in daily_active_shifts[d]) <= max(1, fixed_on_d))

            ng_type = ng_days_dict[doc].get(d)
            if ng_type == "全NG":
                for s in daily_active_shifts[d]: model.Add(shifts[(d, doc, s)] == 0)
            elif ng_type == "日NG":
                for s in daily_active_shifts[d]:
                    if s in DAY_SHIFTS: model.Add(shifts[(d, doc, s)] == 0)
            elif ng_type == "宿NG":
                for s in daily_active_shifts[d]:
                    if s in NIGHT_SHIFTS: model.Add(shifts[(d, doc, s)] == 0)

            # 入りにくい曜日（翌日が平日なら宿直NG）
            dt_obj = datetime.date(target_year, target_month, d)
            nxt = dt_obj + datetime.timedelta(days=1)
            nxt_is_hol = nxt.weekday() >= 5 or jpholiday.is_holiday(nxt) or (nxt.month == target_month and nxt.day in custom_holidays)
            if dt_obj.weekday() in hard_weekdays[doc] and not nxt_is_hol:
                for s in NIGHT_SHIFTS:
                    if s in daily_active_shifts[d]: model.Add(shifts[(d, doc, s)] == 0)

        # 固定/絶対希望
        for d, s in absolute_req_specific[doc]:
            if s in daily_active_shifts[d]: model.Add(shifts[(d, doc, s)] == 1)

        # 間隔ルール
        interval = min_intervals[doc]
        if interval > 0:
            for d in range(1, num_days + 1):
                curr_dt = datetime.date(target_year, target_month, d)
                for past in past_worked_dates[doc]:
                    if 0 < (curr_dt - past).days <= interval:
                        for s in daily_active_shifts[d]: model.Add(shifts[(d, doc, s)] == 0)
            for d1 in range(1, num_days + 1):
                for d2 in range(d1 + 1, min(d1 + interval + 1, num_days + 1)):
                    for s1 in daily_active_shifts[d1]:
                        for s2 in daily_active_shifts[d2]:
                            model.Add(shifts[(d1, doc, s1)] + shifts[(d2, doc, s2)] <= 1)

    model.Maximize(sum(objective_terms))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for d in range(1, num_days + 1):
            dt = datetime.date(target_year, target_month, d)
            row = {"日付": f"{target_month}/{d}", "平日/休日": "休日" if is_holiday(target_year, target_month, d) else "平日"}
            for s in NIGHT_SHIFTS + DAY_SHIFTS:
                docs = [doc for doc in doctors if (d, doc, s) in shifts and solver.Value(shifts[(d, doc, s)]) == 1]
                row[s] = "、".join(docs) if docs else "-"
            res.append(row)
        return pd.DataFrame(res), True
    return None, False

# ==========================================
# 5. 実行
# ==========================================
st.divider()
if st.button("🚀 シフトを自動生成する", type="primary"):
    with st.spinner("計算中..."):
        df_res, success = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, edited_fixed_df)
        if success:
            st.success("作成完了！")
            st.dataframe(df_res, use_container_width=True)
            st.download_button("ダウンロード", df_res.to_csv(index=False).encode('utf-8-sig'), f"shift_{year}_{month}.csv")
        else:
            st.error("作成に失敗しました。条件を緩めてください。")

elif len(staff_df) == 0:
    st.warning("☝️ 表に先生の名前を入力するか、CSVファイルをアップロードしてください。")
