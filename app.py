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

with st.expander("📖 初めての方へ：このアプリの使い方マニュアル（クリックして開く）", expanded=False):
    st.markdown("""
    このアプリは、画面の上から下へ順番に設定していくことで、AIが自動でシフトを作成します。

    ### 1. カレンダーと特別設定（画面上部）
    * **年月の設定**: 作成したいシフトの年と月を選択します。
    * **特別休日の設定**: 平日でも日直が必要な日（年末年始など）は、カレンダーのチェックボックスをオンにして「休日扱い」にします。
    * **複数人シフト（増員）の設定**: GWなどで特定のシフトを「2名以上」に増やしたい場合は、表で日付と枠を指定して人数を変更します。

    ### 2. スタッフ条件の読み込み・入力（必須）
    「📥 ひな形（CSV）」をダウンロードしてExcelで入力しアップロードするか、画面上の表を直接クリックして入力・編集してください。
    
    **【各項目の入力ルール】**
    * **【NG日】** 入力表の下にあるカレンダーから、先生ごとにタブを切り替えて休みたい日をポチポチとクリックして選んでください。
    * **【希望日】** 入りたい日を入力します。
    * **【最低空ける日数】** シフトとシフトの間を最低何日空けるかです。
    * **【月間最大回数】** その月に入るすべてのシフトの「総合計」の上限回数です。

    ### 3. 過去・決定済みシフトの読み込み・入力（任意）
    先月分のシフト表や、今月の「一部だけ確定させたシフト」があれば、読み込ませるか画面に直接入力します。

    ### 4. シフトの自動生成
    設定と入力が終わったら、一番下の「🚀 このデータでシフトを自動生成する」ボタンを押します。
    """)

# === ▼カイゼン：日付は安定の左揃え、曜日だけを見やすく中央揃えにするCSS▼ ===
st.markdown("""
<style>
/* カレンダーの7列ブロックを格子状にする */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) {
    gap: 0 !important;
}

/* 各セルの基本設定（枠線） */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) > div[data-testid="column"] {
    border: 1px solid #d0d0d0 !important;
    margin-right: -1px !important;
    margin-bottom: -1px !important;
    background-color: #ffffff;
    min-height: 60px;
    padding: 10px !important;
    display: flex;
    flex-direction: column;
}

/* ★ポイント：曜日ヘッダー（チェックボックスがない行）だけを中央揃えにする */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)):not(:has(div[data-testid="stCheckbox"])) > div[data-testid="column"] {
    background-color: #f8f9fa;
    justify-content: center !important; /* 上下中央 */
    align-items: center !important;     /* 左右中央 */
    min-height: 40px;
}

div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)):not(:has(div[data-testid="stCheckbox"])) div[data-testid="stMarkdownContainer"] {
    text-align: center !important;
    width: 100%;
}

/* 日付セル（チェックボックスがある行）は安定の左揃えを維持 */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)):has(div[data-testid="stCheckbox"]) > div[data-testid="column"] {
    justify-content: flex-start !important;
    align-items: flex-start !important;
}
</style>
""", unsafe_allow_html=True)
# ==============================================================================

# ==========================================
# 1. 上部ダッシュボード：年月と休日の設定
# ==========================================
st.header("📅 作成するシフトの設定")

today = datetime.date.today()
if today.month == 12:
    default_year = today.year + 1
    default_month = 1
else:
    default_year = today.year
    default_month = today.month + 1

col_y, col_m = st.columns(2)
year = col_y.number_input("年", min_value=2026, value=default_year, step=1)
month = col_m.number_input("月", min_value=1, max_value=12, value=default_month, step=1)

st.divider()

st.subheader(f"📅 カレンダー確認 （特別休日の設定） - {month}月")
st.write("※平日を「休日扱い（日直枠あり）」にしたい場合は、対象の日のチェックボックスをポチッとオンにしてください。")

cal_matrix = calendar.monthcalendar(year, month)
weekdays_ja = ["月", "火", "水", "木", "金", "土", "日"]
custom_holidays = []

# --- 上部カレンダー曜日ヘッダー ---
cols = st.columns(7)
for i, w in enumerate(weekdays_ja):
    color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
    cols[i].markdown(f"<div style='color: {color}; font-weight: bold;'>{w}</div>", unsafe_allow_html=True)

# --- 上部カレンダー日付描画 ---
for week in cal_matrix:
    cols = st.columns(7)
    for i, day in enumerate(week):
        if day != 0:
            date_obj = datetime.date(year, month, day)
            is_weekend_or_hol = date_obj.weekday() >= 5 or jpholiday.is_holiday(date_obj)
            
            with cols[i]:
                if is_weekend_or_hol:
                    st.markdown(f"<div style='color: #ff4b4b; background-color: #ffeeee; padding: 5px; border-radius: 5px; width: 100%;'><b>{day}日</b><br><small>休</small></div>", unsafe_allow_html=True)
                else:
                    if st.checkbox(f"**{day}日**", key=f"hol_{year}_{month}_{day}"):
                        custom_holidays.append(day)
        else:
            with cols[i]:
                st.write("")

st.divider()

st.subheader("👥 複数人シフト（増員）の設定")
st.write("※GWなどで通常1名の枠を「2名以上」に増やしたい場合は、下表に入力してください。")

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
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================
st.header("1. スタッフ条件の読み込み・入力（必須）")

template_data = {
    "先生の名前": ["Dr. A", "Dr. B", "Dr. C", "Dr. D", "Dr. E"],
    "希望日(半角カンマ区切り)": ["", "", "", "", ""], 
    "希望優先度(数字が大きいほど優先)": [1, 1, 1, 1, 1], 
    "最低空ける日数": [5, 5, 5, 5, 5],  
    "月間最小回数": [0, 0, 0, 0, 0],
    "月間最大回数": [5, 5, 5, 5, 5],    
    "宿直A上限": [2, 2, 2, 2, 2],
    "宿直B上限": [2, 2, 2, 2, 2],
    "外来宿直上限": [2, 2, 2, 2, 2],
    "日直A上限": [2, 2, 2, 2, 2],
    "日直B上限": [2, 2, 2, 2, 2],
    "外来日直上限": [2, 2, 2, 2, 2],
    "備考": ["", "", "", "", ""]
}
df_template = pd.DataFrame(template_data)

uploaded_file = st.file_uploader("スタッフ条件（CSV）をアップロード", type="csv", key="staff_csv")

if uploaded_file is not None:
    if st.session_state.get('last_uploaded_file_id') != uploaded_file.file_id:
        for key in list(st.session_state.keys()):
            if key.startswith("ng_"): del st.session_state[key]
        st.session_state['last_uploaded_file_id'] = uploaded_file.file_id
    base_df = parse_staff_csv(uploaded_file.getvalue())
else:
    base_df = df_template.copy()

if "先生の名前" in base_df.columns:
    base_df = base_df.set_index("先生の名前")

if "NG日(半角カンマ区切り)" not in base_df.columns:
    base_df["NG日(半角カンマ区切り)"] = ""

edited_df = st.data_editor(
    base_df, 
    num_rows="dynamic", 
    use_container_width=True, 
    height=300,
    column_config={"NG日(半角カンマ区切り)": None}
)

staff_df = edited_df.reset_index()

st.markdown("##### 🚫 先生ごとのNG日設定（カレンダーでクリック選択）")

valid_staff = staff_df[staff_df["先生の名前"].astype(str).str.strip() != ""]
if not valid_staff.empty:
    doctor_names = valid_staff["先生の名前"].astype(str).tolist()
    tabs = st.tabs(doctor_names)
    
    for t_idx, doc_name in enumerate(doctor_names):
        original_idx = valid_staff.index[t_idx]
        with tabs[t_idx]:
            # 初期復元処理
            current_ng_str = str(valid_staff.loc[original_idx].get("NG日(半角カンマ区切り)", ""))
            current_ng_list = []
            if current_ng_str and current_ng_str.lower() not in ["nan", "none", ""]:
                for x in current_ng_str.split(','):
                    try:
                        val = int(float(x.strip()))
                        if 1 <= val <= num_days: current_ng_list.append(val)
                    except: pass
            
            for d in range(1, num_days + 1):
                chk_key = f"ng_{doc_name}_{year}_{month}_{d}"
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = (d in current_ng_list)

            col_btn1, col_btn2, _ = st.columns([2, 2, 6])
            with col_btn1:
                st.button("✅ 全選択", key=f"btn_all_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, True), use_container_width=True)
            with col_btn2:
                st.button("🗑️ 全解除", key=f"btn_clear_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, False), use_container_width=True)

            with st.form(key=f"ng_form_{original_idx}"):
                st.write(f"※ポチポチ選んだ後、最後に必ず下の**【確定する】**ボタンを押してください。")
                
                cols = st.columns(7)
                for i, w in enumerate(weekdays_ja):
                    color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
                    cols[i].markdown(f"<div style='color: {color}; font-weight: bold;'>{w}</div>", unsafe_allow_html=True)
                
                for week in cal_matrix:
                    cols = st.columns(7)
                    for i, day in enumerate(week):
                        if day != 0:
                            date_obj = datetime.date(year, month, day)
                            is_hol_or_sun = jpholiday.is_holiday(date_obj) or date_obj.weekday() == 6 or (day in custom_holidays)
                            is_sat = date_obj.weekday() == 5 and not is_hol_or_sun
                            day_label = f":red[**{day}日**]" if is_hol_or_sun else (f":blue[**{day}日**]" if is_sat else f"**{day}日**")
                            chk_key = f"ng_{doc_name}_{year}_{month}_{day}"
                            with cols[i]:
                                st.checkbox(day_label, key=chk_key)
                        else:
                            with cols[i]: st.write("")
                st.form_submit_button(f"💾 {doc_name}先生のNG日を確定する")
            
            current_ngs = [str(d) for d in range(1, num_days + 1) if st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", False)]
            staff_df.at[original_idx, "NG日(半角カンマ区切り)"] = ",".join(current_ngs)

current_csv = staff_df.to_csv(index=False).encode('shift_jis')
st.download_button(
    label="📥 現在の入力状況を一時保存する（CSV）",
    data=current_csv,
    file_name=f"staff_wip_{year}_{month}.csv",
    mime="text/csv",
    use_container_width=True
)

st.divider()
st.header("2. 過去・決定済みシフトの入力（任意）")
fixed_columns = ["日付", "平日/休日", "宿直A", "宿直B", "外来宿直", "日直A", "日直B", "外来日直"]
base_fixed_df = pd.DataFrame(columns=fixed_columns)
base_fixed_df.loc[0] = ["" for _ in range(len(fixed_columns))]
edited_fixed_df_raw = st.data_editor(base_fixed_df.set_index("日付"), num_rows="dynamic", use_container_width=True, height=200)
edited_fixed_df = edited_fixed_df_raw.reset_index()

st.divider()
st.header("3. シフトの自動生成")

def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None):
    _, num_days = calendar.monthrange(target_year, target_month)
    NIGHT_SHIFTS = ['宿直A', '宿直B', '外来宿直']
    DAY_SHIFTS = ['日直A', '日直B', '外来日直']
    def is_holiday(y, m, d):
        date = datetime.date(y, m, d)
        return date.weekday() >= 5 or jpholiday.is_holiday(date) or (d in custom_holidays)
    def safe_int(val, default_val):
        try: return int(float(val)) if pd.notna(val) else default_val
        except: return default_val

    doctors = staff_df['先生の名前'].astype(str).tolist()
    ng_days = {}
    min_intervals = {}
    max_shifts_total = {}
    max_shifts_per_type = {}
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}

    for _, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        ng_str = str(row.get('NG日(半角カンマ区切り)', ""))
        ng_days[doc] = [int(x.strip()) for x in ng_str.split(',')] if ng_str and ng_str.lower() != "nan" else []
        min_intervals[doc] = safe_int(row.get('最低空ける日数'), 5)
        max_shifts_total[doc] = safe_int(row.get('月間最大回数'), 5)
        max_shifts_per_type[doc] = {s: safe_int(row.get(f"{s}上限"), 2) for s in NIGHT_SHIFTS + DAY_SHIFTS}

    if fixed_df is not None:
        for _, row in fixed_df.iterrows():
            date_str = str(row.get('日付', ''))
            match = re.match(r'^\s*(\d+)\s*/\s*(\d+)', date_str)
            if match:
                m, d = int(match.group(1)), int(match.group(2))
                if m == target_month:
                    for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
                        if s_type in row and pd.notna(row[s_type]):
                            for doc_val in re.split(r'[、,\s]+', str(row[s_type])):
                                if doc_val.strip() in doctors: absolute_req_specific[doc_val.strip()].append((d, s_type))

    model = cp_model.CpModel()
    shifts = {}
    for d in range(1, num_days + 1):
        active = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        for doc in doctors:
            for s in active: shifts[(d, doc, s)] = model.NewBoolVar(f'd{d}_{doc}_{s}')

    for d in range(1, num_days + 1):
        active = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        for s in active:
            count = multi_slots_dict.get((d, s), 1)
            model.Add(sum(shifts[(d, doc, s)] for doc in doctors) == count)

    for doc in doctors:
        for d in range(1, num_days + 1):
            active = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            model.Add(sum(shifts[(d, doc, s)] for s in active) <= 1)
            if d in ng_days[doc]:
                for s in active: model.Add(shifts[(d, doc, s)] == 0)
        for d, s in absolute_req_specific[doc]:
            if (d, doc, s) in shifts: model.Add(shifts[(d, doc, s)] == 1)
        
        all_worked = [shifts[k] for k in shifts if k[1] == doc]
        model.Add(sum(all_worked) <= max_shifts_total[doc])
        for s in NIGHT_SHIFTS + DAY_SHIFTS:
            s_worked = [shifts[k] for k in shifts if k[1] == doc and k[2] == s]
            if s_worked: model.Add(sum(s_worked) <= max_shifts_per_type[doc][s])

        interval = min_intervals[doc]
        if interval > 0:
            for d1 in range(1, num_days + 1):
                for d2 in range(d1 + 1, min(d1 + interval + 1, num_days + 1)):
                    a1 = [shifts[k] for k in shifts if k[0] == d1 and k[1] == doc]
                    a2 = [shifts[k] for k in shifts if k[0] == d2 and k[1] == doc]
                    if a1 and a2: model.Add(sum(a1) + sum(a2) <= 1)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for d in range(1, num_days + 1):
            date_obj = datetime.date(target_year, target_month, d)
            row = {"日付": f"{target_month}/{d}({weekdays_ja[date_obj.weekday()]})", "平日/休日": "休日" if is_holiday(target_year, target_month, d) else "平日"}
            for s in NIGHT_SHIFTS + DAY_SHIFTS:
                docs = [doc for doc in doctors if (d, doc, s) in shifts and solver.Value(shifts[(d, doc, s)]) == 1]
                row[s] = "、".join(docs) if docs else "-"
            res.append(row)
        return pd.DataFrame(res), True
    return None, False

if st.button("🚀 シフトを自動生成する", type="primary"):
    with st.spinner("AIが計算中..."):
        df_result, success = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, edited_fixed_df)
        if success:
            st.success("✨ 作成成功！")
            st.dataframe(df_result, use_container_width=True, hide_index=True)
            st.download_button("📥 結果をダウンロード", df_result.to_csv(index=False).encode('shift_jis'), f"shift_{year}_{month}.csv", "text/csv")
        else:
            st.error("条件が厳しすぎてシフトが組めませんでした。条件を緩めてください。")
