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
    if "NG日(半角カンマ区切り)" in df.columns:
        df = df.drop(columns=["NG日(半角カンマ区切り)"])
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
    * **【NG日】** 入力表の下にあるカレンダーから、先生ごとにタブを切り替えて休みたい日をポチポチとクリックして選んでください。（※選び終わったら必ず「確定する」ボタンを押してください！）
    * **【希望日】** 入りたい日を入力します。
        * 日付だけを指定（例: `10, 15`）→ その日の「どれかのシフト」に入ります。
        * 種類まで指定（例: `10:宿直A, 15:日直B`）→ その日の「その枠」を狙います。（※コロン `:` は半角/全角どちらでもOK）
    * **【希望優先度】** 希望を通すための「相対的な強さ」です。基本は `1` です。
        * 例：Dr. Aを `10`、Dr. Bを `1` にして同じ日を希望して競合した場合、AIはDr. Aの希望を優先的に叶えます（※ただし間隔などのルール範囲内）。
        * **特例（絶対希望）**：ここを `100` 以上にすると、間隔ルールや上限回数をすべて無視して【確実】にそのシフトに入ります。
    * **【最低空ける日数】** シフトとシフトの間を最低何日空けるかです。（人ごとに設定できます）
    * **【月間最小回数】** その月に入るシフトの「最低保証」回数です。（未入力や空欄の場合は0回扱いになります）
    * **【月間最大回数】** その月に入るすべてのシフトの「総合計」の上限回数です。（人ごとに設定できます）
    * **【各種上限】** 「宿直A」「日直B」など枠ごとの上限回数です。
    * **【備考】** メモや説明などを自由に書き込める欄です。（AIの計算には影響しません）

    ### 3. 過去・決定済みシフトの読み込み・入力（任意）
    先月分のシフト表や、今月の「一部だけ確定させたシフト」があれば、読み込ませるか画面に直接入力します。
    * **先月のシフトを入れた場合**: 前月末の勤務を考慮し、月初の間隔ルールをしっかり守ります。
    * **今月の途中まで作ったシフトを入れた場合**: その部分は「確定」として固定し、残りの空き枠だけをAIが綺麗に埋めてくれます！

    ### 4. シフトの自動生成
    設定と入力が終わったら、一番下の「🚀 このデータでシフトを自動生成する」ボタンを押します。
    💡 **ポイント**: 自動生成ボタンを押すたびに、AIが少しずつ違うパターンのシフトを提案してくれます。完成した表はCSVでダウンロードできます。
    """)

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

# === ▼カイゼン：文字の色は残しつつ、背景だけ薄い黄色にハイライトするCSS▼ ===
st.markdown("""
<style>
div[data-testid="stCheckbox"] {
    padding: 5px;
    border-radius: 5px;
    transition: all 0.2s ease;
}
div[data-testid="stCheckbox"]:has(input:checked) {
    background-color: #fffacc; /* 薄い黄色で選択状態をアピール */
    border: 1px solid #f4d03f;
}
</style>
""", unsafe_allow_html=True)
# =========================================================================

cal_matrix = calendar.monthcalendar(year, month)
weekdays_ja = ["月", "火", "水", "木", "金", "土", "日"]
custom_holidays = []

cols = st.columns(7)
for i, w in enumerate(weekdays_ja):
    color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
    cols[i].markdown(f"<div style='text-align: center; color: {color}; font-weight: bold;'>{w}</div>", unsafe_allow_html=True)

for week in cal_matrix:
    cols = st.columns(7)
    for i, day in enumerate(week):
        if day != 0:
            date_obj = datetime.date(year, month, day)
            is_weekend_or_hol = date_obj.weekday() >= 5 or jpholiday.is_holiday(date_obj)
            
            with cols[i]:
                if is_weekend_or_hol:
                    st.markdown(f"<div style='text-align: center; color: #ff4b4b; background-color: #ffeeee; padding: 5px; border-radius: 5px; margin-bottom: 10px;'><b>{day}日</b><br><small>休</small></div>", unsafe_allow_html=True)
                else:
                    if st.checkbox(f"**{day}日**", key=f"hol_{year}_{month}_{day}", help="クリックで休日扱いに変更"):
                        custom_holidays.append(day)
        else:
            with cols[i]:
                st.write("")

st.divider()

st.subheader("👥 複数人シフト（増員）の設定")
st.write("※GWなどで通常1名の枠を「2名以上」に増やしたい場合は、下表に入力してください。（不要な行は選択してDeleteキーで消せます）")

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
total_night_slots = 0
total_day_slots = 0
for d in range(1, num_days + 1):
    date_obj = datetime.date(year, month, d)
    is_hol = jpholiday.is_holiday(date_obj) or date_obj.weekday() >= 5 or (d in custom_holidays)
    
    for s in NIGHT_SHIFTS_UI:
        total_night_slots += multi_slots_dict.get((d, s), 1)
        
    if is_hol:
        for s in DAY_SHIFTS_UI:
            total_day_slots += multi_slots_dict.get((d, s), 1)

total_slots = total_night_slots + total_day_slots

st.subheader(f"📌 {year}年{month}月の必要シフト枠数")
col1, col2, col3 = st.columns(3)
col1.metric("🌙 宿直枠 (A・B・外来)", f"計 {total_night_slots} 枠")
col2.metric("☀️ 日直枠 (A・B・外来)", f"計 {total_day_slots} 枠")
col3.metric("🏥 月間 総シフト数", f"合計 {total_slots} 枠")

st.divider()

# ==========================================
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================
st.header("1. スタッフ条件の読み込み・入力（必須）")

template_data = {
    "先生の名前": ["Dr. A", "Dr. B", "Dr. C", "Dr. D", "Dr. E"],
    "希望日(半角カンマ区切り)": ["10:宿直A, 15:日直B", "", "8", "20", ""], 
    "希望優先度(数字が大きいほど優先)": [100, 1, 1, 1, 1], 
    "最低空ける日数": [5, 4, 6, 5, 3],  
    "月間最小回数": [1, 2, 0, 1, 0],
    "月間最大回数": [5, 6, 4, 5, 7],    
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
    uploaded_file = st.file_uploader("スタッフ条件（途中保存CSVも可）をアップロード", type="csv", key="staff_csv")

if uploaded_file is not None:
    # 新しいファイルがアップロードされたら、古いカレンダーの記憶を一度リセットする処理
    if st.session_state.get('last_uploaded_file') != uploaded_file.name:
        for key in list(st.session_state.keys()):
            if key.startswith("ng_"):
                del st.session_state[key]
        st.session_state['last_uploaded_file'] = uploaded_file.name
        
    base_df = parse_staff_csv(uploaded_file.getvalue())
else:
    base_df = df_template.copy()

if "先生の名前" in base_df.columns:
    base_df = base_df.set_index("先生の名前")

# NG日を格納する列がなければ裏側で作る
if "NG日(半角カンマ区切り)" not in base_df.columns:
    base_df["NG日(半角カンマ区切り)"] = ""

st.markdown("##### 👩‍⚕️ スタッフ条件の入力・編集")
st.write("※以下の表は直接クリックして文字を入力できます。")

# 表上では「NG日」列を非表示にして、手入力のミスを防ぎます
edited_df = st.data_editor(
    base_df, 
    num_rows="dynamic", 
    use_container_width=True, 
    height=300,
    column_config={
        "NG日(半角カンマ区切り)": None 
    }
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
            
            # --- 復元処理：途中保存CSVなどからNG日を読み取ってチェックを入れる準備 ---
            current_ng_str = str(valid_staff.loc[original_idx].get("NG日(半角カンマ区切り)", ""))
            current_ng_str = current_ng_str.translate(str.maketrans('０１２３４５６７８９，．', '0123456789,.'))
            current_ng_list = []
            if current_ng_str and current_ng_str.lower() not in ["nan", "none", ""]:
                for x in current_ng_str.split(','):
                    try:
                        val = int(float(x.strip()))
                        if 1 <= val <= num_days:
                            current_ng_list.append(val)
                    except:
                        pass
            
            for d in range(1, num_days + 1):
                chk_key = f"ng_{doc_name}_{year}_{month}_{d}"
                # まだ記憶がなければ、CSVのデータを初期値として採用する
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = (d in current_ng_list)

            # === 一括操作ボタン ===
            st.write("▼ **一括操作**（※操作後は下の確定ボタンを押す必要はありません）")
            col_btn1, col_btn2, _ = st.columns([2, 2, 6])
            with col_btn1:
                st.button("✅ 全選択", key=f"btn_all_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, True), use_container_width=True)
            with col_btn2:
                st.button("🗑️ 全解除", key=f"btn_clear_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, False), use_container_width=True)

            # === カレンダー本体（フォームによる完全隔離） ===
            with st.form(key=f"ng_form_{original_idx}"):
                st.write(f"※カレンダーで休みたい日をポチポチ選んだ後、最後に必ず下の**【確定する】**ボタンを押してください。")
                
                new_ng_list = []
                
                cols = st.columns(7)
                for i, w in enumerate(weekdays_ja):
                    color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
                    cols[i].markdown(f"<div style='text-align: center; color: {color}; font-weight: bold;'>{w}</div>", unsafe_allow_html=True)
                
                for week in cal_matrix:
                    cols = st.columns(7)
                    for i, day in enumerate(week):
                        if day != 0:
                            date_obj = datetime.date(year, month, day)
                            is_hol_or_sun = jpholiday.is_holiday(date_obj) or date_obj.weekday() == 6 or (day in custom_holidays)
                            is_sat = date_obj.weekday() == 5 and not is_hol_or_sun
                            
                            if is_hol_or_sun:
                                day_label = f":red[**{day}日**]"
                            elif is_sat:
                                day_label = f":blue[**{day}日**]"
                            else:
                                day_label = f"**{day}日**"
                                
                            chk_key = f"ng_{doc_name}_{year}_{month}_{day}"
                            
                            with cols[i]:
                                if st.checkbox(day_label, key=chk_key):
                                    new_ng_list.append(day)
                        else:
                            with cols[i]:
                                st.write("")
                
                st.form_submit_button(f"💾 {doc_name}先生のNG日を確定する")
            
            # 確定された結果を最終的なデータ（AIに渡す用＆一時保存用）に反映
            current_ngs = [str(d) for d in range(1, num_days + 1) if st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", False)]
            staff_df.at[original_idx, "NG日(半角カンマ区切り)"] = ",".join(current_ngs)

# === 一時保存（WIP）ダウンロードボタン ===
st.markdown("##### 💾 入力状況の保存（後で再開したい場合）")
st.write("※途中で入力をやめる場合は、ここまでのデータを保存しておき、次回アップロードすることで続きから再開できます。")

current_csv = staff_df.to_csv(index=False).encode('shift_jis')
st.download_button(
    label="📥 現在のスタッフ条件を一時保存する（CSVダウンロード）",
    data=current_csv,
    file_name=f"staff_wip_{year}_{month}.csv",
    mime="text/csv",
    use_container_width=True
)

st.divider()

st.header("2. 過去・決定済みシフトの読み込み・入力（任意）")
st.markdown("""
先月分や来月分のシフト表、または今月の「一部だけ人間が確定させたシフト表」があればアップロードしてください。
前後の月のシフト間隔を考慮したり、確定済みの枠を固定して残りの空白をAIに計算させることができます。
""")

fixed_columns = ["日付", "平日/休日", "宿直A", "宿直B", "外来宿直", "日直A", "日直B", "外来日直"]
fixed_template_df = pd.DataFrame(columns=fixed_columns)
fixed_csv_template = fixed_template_df.to_csv(index=False).encode('shift_jis')

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
    ng_days = {}
    req_days = {}          
    req_specific = {}      
    req_priority = {} 
    
    min_intervals = {}
    min_shifts_total = {}
    max_shifts_total = {}
    max_shifts_per_type = {}
    
    absolute_req_days = {doc: [] for doc in doctors}
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}
    future_worked_dates = {doc: [] for doc in doctors}

    invalid_requests = []

    if fixed_df is not None:
        for _, row in fixed_df.iterrows():
            date_str = str(row.get('日付', ''))
            match = re.match(r'^\s*(\d+)\s*/\s*(\d+)', date_str)
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
                                    active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
                                    if s_type not in active_shifts:
                                        day_type = "休日" if is_holiday(target_year, target_month, d) else "平日"
                                        invalid_requests.append(f"❌ **「決定済みシフト(入力欄)」のエラー**: {target_month}月{d}日（{day_type}）には「{s_type}」枠がありませんが、{doc_val}先生が誤って入力されています。")
                                    else:
                                        absolute_req_specific[doc_val].append((d, s_type))
                                elif date_obj < datetime.date(target_year, target_month, 1):
                                    past_worked_dates[doc_val].append(date_obj)
                                else:
                                    future_worked_dates[doc_val].append(date_obj)

    for index, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        
        ng_str = str(row['NG日(半角カンマ区切り)'])
        if pd.isna(row['NG日(半角カンマ区切り)']) or ng_str.strip() == "" or ng_str.lower() in ["nan", "none"]:
            ng_days[doc] = []
        else:
            try:
                ng_days[doc] = [int(x.strip()) for x in ng_str.split(',')]
            except:
                ng_days[doc] = []
                
        req_days[doc] = []
        req_specific[doc] = []
        if '希望日(半角カンマ区切り)' in staff_df.columns:
            req_str = str(row['希望日(半角カンマ区切り)'])
            if not (pd.isna(row['希望日(半角カンマ区切り)']) or req_str.strip() == "" or req_str.lower() in ["nan", "none"]):
                req_str = req_str.replace('：', ':')
                items = req_str.split(',')
                for item in items:
                    item = item.strip()
                    if not item:
                        continue
                    if ':' in item:
                        parts = item.split(':')
                        try:
                            d = int(parts[0].strip())
                            s_name = parts[1].strip()
                            req_specific[doc].append((d, s_name))
                        except:
                            pass
                    else:
                        try:
                            req_days[doc].append(int(item))
                        except:
                            pass

        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('最低空ける日数'), 5)
        min_shifts_total[doc] = safe_int(row.get('月間最小回数'), 0)
        max_shifts_total[doc] = safe_int(row.get('月間最大回数'), 5)

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
            else:
                active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
                if s_name not in active_shifts:
                    day_type = "休日" if is_holiday(target_year, target_month, d) else "平日"
                    invalid_requests.append(f"❌ **{doc}先生**: {target_month}月{d}日（{day_type}）には「{s_name}」というシフト枠はありません。（平日に日直を指定しているか、文字が間違っている可能性があります）")

    for doc in doctors:
        if req_priority[doc] >= 100:
            absolute_req_days[doc].extend([d for d in req_days[doc] if 1 <= d <= num_days])
            absolute_req_specific[doc].extend([(d, s) for (d, s) in req_specific[doc] if 1 <= d <= num_days])
            
            all_abs_dates = absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]]
            ng_days[doc] = [d for d in ng_days[doc] if d not in all_abs_dates]

    for d in range(1, num_days + 1):
        active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        
        for s_name in active_shifts:
            req_docs = [doc for doc in doctors if (d, s_name) in absolute_req_specific[doc]]
            req_count = multi_slots_dict.get((d, s_name), 1)
            if len(req_docs) > req_count:
                invalid_requests.append(f"❌ **{target_month}月{d}日**: 「{s_name}」枠（定員{req_count}名）に、定員を超える先生（{', '.join(req_docs)}）が確定指定（優先度100、または決定済みシフト）されているためパズルが破綻しています。")
                
        abs_req_docs = [doc for doc in doctors if (d in absolute_req_days[doc] or any(sd == d for (sd, ss) in absolute_req_specific[doc]))]
        abs_req_docs = list(set(abs_req_docs))
        daily_req_count = sum(multi_slots_dict.get((d, s), 1) for s in active_shifts)
        if len(abs_req_docs) > daily_req_count:
            invalid_requests.append(f"❌ **{target_month}月{d}日**: その日の総枠数({daily_req_count}枠)に対して、確定指定が{len(abs_req_docs)}名（{', '.join(abs_req_docs)}）もいるため、全員を入れられません。")

    if invalid_requests:
        unique_invalid = list(dict.fromkeys(invalid_requests))
        return None, False, unique_invalid, None, None

    model = cp_model.CpModel()
    shifts = {}

    for d in range(1, num_days + 1):
        active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        for doc in doctors:
            for s in active_shifts:
                shifts[(d, doc, s)] = model.NewBoolVar(f'shift_d{d}_{doc}_{s}')

    for d in range(1, num_days + 1):
        active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        for s in active_shifts:
            req_count = multi_slots_dict.get((d, s), 1)
            model.Add(sum(shifts[(d, doc, s)] for doc in doctors) == req_count)

    for doc in doctors:
        for d in range(1, num_days + 1):
            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            
            fixed_count = sum(1 for sd, ss in absolute_req_specific[doc] if sd == d and ss in active_shifts)
            max_shifts_today = max(1, fixed_count)
            
            model.Add(sum(shifts[(d, doc, s)] for s in active_shifts) <= max_shifts_today)

    for doc in doctors:
        for d in ng_days[doc]:
            if 1 <= d <= num_days:
                active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
                for s in active_shifts:
                    model.Add(shifts[(d, doc, s)] == 0)

    for doc in doctors:
        for d in absolute_req_days[doc]:
            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            model.AddExactlyOne(shifts[(d, doc, s)] for s in active_shifts)
            
        for d, s_name in absolute_req_specific[doc]:
            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            if s_name in active_shifts:
                model.Add(shifts[(d, doc, s_name)] == 1)

    for doc in doctors:
        for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
            worked = [shifts[(d, doc, s_type)] for d in range(1, num_days + 1) 
                      if s_type in (NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS)]
            if worked:
                specific_req_count = sum(1 for d, s in absolute_req_specific[doc] if s == s_type)
                actual_max_type = max(max_shifts_per_type[doc][s_type], specific_req_count)
                model.Add(sum(worked) <= actual_max_type)

    for doc in doctors:
        worked_all = []
        for d in range(1, num_days + 1):
            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            for s in active_shifts:
                worked_all.append(shifts[(d, doc, s)])
        if worked_all:
            all_abs_dates = absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]]
            actual_max_total = max(max_shifts_total[doc], len(all_abs_dates))
            actual_min_total = min(min_shifts_total[doc], actual_max_total)
            model.Add(sum(worked_all) <= actual_max_total)
            model.Add(sum(worked_all) >= actual_min_total)

    for doc in doctors:
        interval = min_intervals[doc]
        if interval > 0:
            all_abs_dates = set(absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]])
            
            for d in range(1, num_days + 1):
                if d in all_abs_dates:
                    continue
                
                current_date = datetime.date(target_year, target_month, d)
                active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
                
                for past_date in past_worked_dates[doc]:
                    if 0 < (current_date - past_date).days <= interval:
                        for s in active_shifts:
                            model.Add(shifts[(d, doc, s)] == 0)
                            
                for future_date in future_worked_dates[doc]:
                    if 0 < (future_date - current_date).days <= interval:
                        for s in active_shifts:
                            model.Add(shifts[(d, doc, s)] == 0)
            
            for d1 in range(1, num_days + 1):
                for d2 in range(d1 + 1, min(d1 + interval + 1, num_days + 1)):
                    if d1 in all_abs_dates and d2 in all_abs_dates:
                        continue
                        
                    active_shifts_d1 = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d1) else NIGHT_SHIFTS
                    active_shifts_d2 = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d2) else NIGHT_SHIFTS
                    
                    for s1 in active_shifts_d1:
                        for s2 in active_shifts_d2:
                            model.Add(shifts[(d1, doc, s1)] + shifts[(d2, doc, s2)] <= 1)

    holiday_worked = {}
    for doc in doctors:
        hol_shifts = []
        for d in range(1, num_days + 1):
            if is_holiday(target_year, target_month, d):
                for s in NIGHT_SHIFTS + DAY_SHIFTS:
                    hol_shifts.append(shifts[(d, doc, s)])
        holiday_worked[doc] = sum(hol_shifts)
        
    global_max = max(max_shifts_total.values()) if max_shifts_total else 15
    max_hol_shifts = model.NewIntVar(0, global_max, 'max_hol_shifts')
    for doc in doctors:
        model.Add(holiday_worked[doc] <= max_hol_shifts)
        
    objective_terms = []
    for doc in doctors:
        if req_priority[doc] < 100:  
            weight = req_priority[doc] * 100 
            
            for d in req_days[doc]:
                if 1 <= d <= num_days:
                    active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
                    for s in active_shifts:
                        objective_terms.append(shifts[(d, doc, s)] * weight)
                        
            for d, s_name in req_specific[doc]:
                if 1 <= d <= num_days:
                    active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
                    if s_name in active_shifts:
                        objective_terms.append(shifts[(d, doc, s_name)] * weight)
                    
    if objective_terms:
        model.Maximize(sum(objective_terms) - max_hol_shifts * 10)
    else:
        model.Minimize(max_hol_shifts)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 45.0
    solver.parameters.random_seed = random.randint(1, 10000)
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        schedule_list = []
        weekday_ja = ["月", "火", "水", "木", "金", "土", "日"]
        
        for d in range(1, num_days + 1):
            date_obj = datetime.date(target_year, target_month, d)
            day_str = "休日" if is_holiday(target_year, target_month, d) else "平日"
            row = {"日付": f"{target_month}/{d}({weekday_ja[date_obj.weekday()]})", "平日/休日": day_str}
            
            for s in NIGHT_SHIFTS + DAY_SHIFTS:
                row[s] = "-"
                
            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            for s in active_shifts:
                assigned_docs = []
                for doc in doctors:
                    if solver.Value(shifts[(d, doc, s)]) == 1:
                        assigned_docs.append(doc)
                if assigned_docs:
                    row[s] = "、".join(assigned_docs)
            schedule_list.append(row)
        return pd.DataFrame(schedule_list), True, [], past_worked_dates, future_worked_dates
    
    else:
        reasons = []
        
        for d in range(1, num_days + 1):
            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            for s_name in active_shifts:
                req_docs = [doc for doc in doctors if (d, s_name) in absolute_req_specific[doc]]
                req_count = multi_slots_dict.get((d, s_name), 1)
                if len(req_docs) > req_count:
                    reasons.append(f"❌ **{target_month}/{d}**: 「{s_name}」枠（定員{req_count}名）に、定員を超える先生（{', '.join(req_docs)}）が確定指定しているため、パズルが破綻しています。")

        for d in range(1, num_days + 1):
            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
            req_slots = sum(multi_slots_dict.get((d, s), 1) for s in active_shifts)
            available = sum(1 for doc in doctors if d not in ng_days[doc])
            
            if available < req_slots:
                reasons.append(f"❌ **{target_month}/{d}**: 必要な枠({req_slots}枠)に対して、出勤可能な先生({available}名)が足りません。（増員設定に対してNG希望者が多すぎます）")
            elif available <= req_slots + 2:
                reasons.append(f"⚠️ **{target_month}/{d}**: 出勤可能な先生が{available}名しかおらず、人ごとの「最低空ける日数」ルールの影響でパズルが詰まっている可能性が高いです。")
                
        for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
            req_total = sum(multi_slots_dict.get((d, s_type), 1) for d in range(1, num_days + 1) if s_type in (NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS))
            max_available = sum(max_shifts_per_type[doc][s_type] for doc in doctors)
            if max_available < req_total:
                reasons.append(f"❌ **「{s_type}」枠**: 月間に必要な総枠数({req_total}枠)に対して、先生全員の「上限回数の合計」({max_available}回)が足りていません。上限を増やす必要があります。")
                
        theoretical_total = sum(max_shifts_total[doc] for doc in doctors)
        req_all_slots = sum(multi_slots_dict.get((d, s), 1) for d in range(1, num_days + 1) for s in (NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS))
        
        if theoretical_total < req_all_slots:
            reasons.append(f"❌ **全体的な人数不足**: 増員を含めて月間に必要な総シフト数({req_all_slots}枠)に対し、先生全員の「月間最大回数」を足し合わせても({theoretical_total}枠分)足りていません。各人の最大回数を増やしてください。")

        theoretical_min_total = sum(min_shifts_total[doc] for doc in doctors)
        if theoretical_min_total > req_all_slots:
            reasons.append(f"❌ **最小回数の設定オーバー**: 先生全員の「月間最小回数」の合計({theoretical_min_total}回)が、月間に必要な総シフト数({req_all_slots}枠)を上回っているため、全員の希望（最低回数）を満たすことができません。「月間最小回数」を下げてください。")

        if not reasons:
            reasons.append("⚠️ 特定の日付に明白な不足は見つかりませんでしたが、人ごとの「最低空ける日数」や「最大回数」ルールの連鎖によってどこかの日程でパズルが破綻しています。条件の厳しい先生の設定を緩めてみてください。")
            
        return None, False, reasons, None, None

# ==========================================
# 5. 実行ボタンと結果表示
# ==========================================
st.divider()
st.header("3. シフトの自動生成")

staff_df = staff_df[staff_df['先生の名前'].astype(str).str.strip() != '']
staff_df = staff_df.dropna(subset=['先生の名前']).reset_index(drop=True)

fixed_df = edited_fixed_df[edited_fixed_df['日付'].astype(str).str.strip() != '']
fixed_df = fixed_df.dropna(subset=['日付']).reset_index(drop=True)

if len(staff_df) > 0 and st.button("🚀 このデータでシフトを自動生成する", type="primary"):
    with st.spinner("AIが最適なシフトを計算中...（最大45秒かかります）"):
        try:
            df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df)
            
            if success:
                st.success("✨ シフトの作成に成功しました！個人のルール（間隔・回数）を厳守し、優先度100以上の絶対希望や確定シフトは全て確約されています。")
                
                def highlight_holidays(row):
                    styles = [''] * len(row)
                    if row['平日/休日'] == '休日':
                        for i, col in enumerate(row.index):
                            if col in ['日付', '平日/休日']: 
                                styles[i] = 'color: #ff4b4b; font-weight: bold;'
                    return styles
                
                styled_df = df_result.style.apply(highlight_holidays, axis=1)
                
                st.subheader("📅 完成したシフト表")
                result_height = len(df_result) * 35 + 40
                st.dataframe(styled_df, use_container_width=True, hide_index=True, height=result_height)
                
                st.subheader("📊 先生ごとのシフト回数（実績）")
                shift_columns = ['宿直A', '宿直B', '外来宿直', '日直A', '日直B', '外来日直']
                summary_list = []
                doctors_list = staff_df['先生の名前'].astype(str).tolist()
                
                req_days_eval = {}
                req_spec_eval = {}
                for index, row in staff_df.iterrows():
                    doc = str(row['先生の名前'])
                    req_days_eval[doc] = []
                    req_spec_eval[doc] = []
                    if '希望日(半角カンマ区切り)' in staff_df.columns:
                        req_str = str(row['希望日(半角カンマ区切り)'])
                        if not (pd.isna(row['希望日(半角カンマ区切り)']) or req_str.strip() == "" or req_str.lower() in ["nan", "none"]):
                            req_str = req_str.replace('：', ':')
                            for item in req_str.split(','):
                                item = item.strip()
                                if not item: continue
                                if ':' in item:
                                    parts = item.split(':')
                                    try:
                                        req_spec_eval[doc].append((int(re.sub(r'\D', '', parts[0].strip())), parts[1].strip()))
                                    except:
                                        pass
                                else:
                                    try:
                                        req_days_eval[doc].append(int(item))
                                    except:
                                        pass
                
                for doc in doctors_list:
                    doc_data = {"先生の名前": doc}
                    total_count = 0
                    hol_count = 0
                    
                    doc_working_dates = set()
                    
                    if past_worked_dates and doc in past_worked_dates:
                        doc_working_dates.update(past_worked_dates[doc])
                    if future_worked_dates and doc in future_worked_dates:
                        doc_working_dates.update(future_worked_dates[doc])
                    
                    for d_idx in range(len(df_result)):
                        row = df_result.iloc[d_idx]
                        is_working = False
                        for s in shift_columns:
                            cell_val = str(row[s])
                            if doc in [x.strip() for x in re.split(r'[、,]', cell_val)]:
                                is_working = True
                                break
                        if is_working:
                            doc_working_dates.add(datetime.date(year, month, d_idx + 1))
                    
                    for s in shift_columns:
                        count = sum(1 for val in df_result[s] if doc in [x.strip() for x in re.split(r'[、,]', str(val))])
                        doc_data[s] = count
                        total_count += count
                        hol_count += sum(1 for val in df_result[df_result['平日/休日'] == '休日'][s] if doc in [x.strip() for x in re.split(r'[、,]', str(val))])
                                
                    doc_data["土日祝の回数"] = hol_count
                    doc_data["総合計"] = total_count
                    
                    sorted_dates = sorted(list(doc_working_dates))
                    if len(sorted_dates) >= 2:
                        intervals = [(sorted_dates[i] - sorted_dates[i-1]).days - 1 for i in range(1, len(sorted_dates))]
                        doc_data["最小間隔"] = min(intervals)
                        doc_data["平均間隔"] = sum(intervals) / len(intervals)
                    else:
                        doc_data["最小間隔"] = None
                        doc_data["平均間隔"] = None
                        
                    total_reqs = len(req_days_eval[doc]) + len(req_spec_eval[doc])
                    if total_reqs > 0:
                        current_month_days = [d.day for d in sorted_dates if d.month == month and d.year == year]
                        granted = sum(1 for d in req_days_eval[doc] if d in current_month_days)
                        for req_d, req_s in req_spec_eval[doc]:
                            if req_d - 1 < len(df_result):
                                row_result = df_result.iloc[req_d - 1]
                                if req_s in row_result and doc in [x.strip() for x in re.split(r'[、,]', str(row_result[req_s]))]:
                                    granted += 1
                        doc_data["希望日の達成"] = f"{granted} / {total_reqs} 回"
                    else:
                        doc_data["希望日の達成"] = "-"
                    
                    summary_list.append(doc_data)
                    
                df_summary = pd.DataFrame(summary_list)
                df_summary = df_summary[['先生の名前', '宿直A', '宿直B', '外来宿直', '日直A', '日直B', '外来日直', '土日祝の回数', '総合計', '希望日の達成', '最小間隔', '平均間隔']]
                
                df_summary = df_summary.set_index('先生の名前')
                
                styled_summary = df_summary.style.format(
                    {"最小間隔": "{:.0f}", "平均間隔": "{:.1f}"}, na_rep="-"
                ).set_properties(
                    subset=['総合計'], **{'font-weight': 'bold'}
                ).set_properties(
                    subset=['希望日の達成'], **{'text-align': 'center'}
                )
                
                summary_height = len(df_summary) * 35 + 40
                st.dataframe(styled_summary, use_container_width=True, height=summary_height)
                
                csv_result = df_result.to_csv(index=False).encode('shift_jis')
                st.download_button(
                    label="📥 完成したシフト表をCSVでダウンロード",
                    data=csv_result,
                    file_name=f"shift_{year}_{month}_result.csv",
                    mime="text/csv",
                )
            else:
                st.error("入力された条件に誤りがあるか、条件が厳しすぎてシフトが組めませんでした。")
                st.warning("💡 **以下の原因が考えられます。Excelの入力や設定画面を見直してください。**")
                for reason in error_reasons:
                    st.write(reason)
                    
        except Exception as e:
            st.error(f"シフト計算中にエラーが発生しました。詳細: {e}")
elif len(staff_df) == 0:
    st.warning("☝️ 表に先生の名前を入力するか、CSVファイルをアップロードしてください。")
