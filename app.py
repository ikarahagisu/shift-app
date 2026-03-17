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

# ページ設定
st.set_page_config(page_title="シフト作成アプリ", layout="wide")
st.title("当直・日直 自動シフト作成アプリ")

with st.expander("📖 初めての方へ：このアプリの使い方マニュアル（クリックして開く）", expanded=False):
    st.markdown("""
    ### 1. 準備する（スタッフ条件データの作成）
    まずは画面中央の「📥 ひな形（CSV）をダウンロード」ボタンを押して、専用のファイルを入手します。
    Excelなどで開き、先生ごとの条件を入力して上書き保存（CSV形式）してください。
    ※【便利機能】CSVを使わず、画面上の表を直接クリックして入力・編集することも可能です！
    
    **【各項目の入力ルール】**
    * **【NG日】** 入れない日を半角数字で入力します。（例: `5,12,20`）
    * **【希望日】** 入りたい日を入力します。
        * 日付だけを指定（例: `10, 15`）→ その日の「どれかのシフト」に入ります。
        * 種類まで指定（例: `10:宿直A, 15:日直B`）→ その日の「その枠」を狙います。（※コロン `:` は半角/全角どちらでもOK）
    * **【希望優先度】** 希望を通すための「相対的な強さ」です。基本は `1` です。
        * 例：Dr. Aを `10`、Dr. Bを `1` にして同じ日を希望して競合した場合、AIはDr. Aの希望を優先的に叶えます（※ただし間隔などのルール範囲内）。
        * **特例（絶対希望）**：ここを `100` 以上にすると、間隔ルールや上限回数をすべて無視して【確実】にそのシフトに入ります。
    * **【最低空ける日数】** シフトとシフトの間を最低何日空けるかです。（人ごとに設定できます）
    * **【月間最大回数】** その月に入るすべてのシフトの「総合計」の上限回数です。（人ごとに設定できます）
    * **【各種上限】** 「宿直A」「日直B」など枠ごとの上限回数です。
    * **【備考】** メモや説明などを自由に書き込める欄です。（AIの計算には影響しません）

    ### 2. アプリで設定する
    画面上部で、作成したい「年」と「月」を選びます。
    年末年始やお盆など、カレンダー上は平日でも日直が必要な（休日扱いにする）日がある場合は「特別休日の設定」に入力します。
    GWなどで特定のシフトを「2名以上」の体制にしたい場合は、「複数人シフトの設定」に増員したい枠を入力してください。

    ### 3. シフトを自動生成する
    条件を入力し、「🚀 自動生成」ボタンを押します。
    💡 **ポイント**: 自動生成ボタンを押すたびに、AIが少しずつ違うパターンのシフトを提案してくれます。
    
    👑 **【超便利機能：過去や決定済みのシフトも考慮する】**
    「2. 過去・決定済みシフトの読み込み・入力（任意）」に、このアプリで出力したシフト表をそのままアップロードできます。
    * **先月のシフトを入れた場合**: 前月末の勤務を考慮し、月初の間隔ルールをしっかり守ります。
    * **今月の途中まで作ったシフトを入れた場合**: その部分は「確定」として固定し、残りの空き枠（空白の部分）だけをAIが綺麗に埋めてくれます！
    * ※CSVを読み込まずに、画面上の表に直接「日付」と「先生の名前」を手入力することも可能です。
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

st.markdown("##### 特別休日の設定")
st.write("※カレンダー上は平日でも、休日（日直枠あり）として扱いたい日を入力してください。")
custom_holidays_str = st.text_input("祝日扱いにする日（半角カンマ区切り。例: 29,30,31）", "")

custom_holidays = []
if custom_holidays_str:
    try:
        custom_holidays = [int(x.strip()) for x in custom_holidays_str.split(',') if x.strip().isdigit()]
    except:
        st.error("数字と半角カンマで入力してください。")

st.markdown("##### 複数人シフト（増員）の設定")
st.write("※GWや年末年始など、通常1名の枠を「2名以上」に増やしたい場合に入力してください。")
multi_slots_str = st.text_input("増員する「日:枠名:人数」（例: 3日:宿直A:2名, 4日:日直B:3名）※カンマ区切り", "")

multi_slots_dict = {}
if multi_slots_str:
    for item in multi_slots_str.split(','):
        parts = item.strip().split(':')
        if len(parts) == 3:
            try:
                d_val = int(re.sub(r'\D', '', parts[0]))
                s_val = parts[1].strip()
                c_val = int(re.sub(r'\D', '', parts[2]))
                multi_slots_dict[(d_val, s_val)] = c_val
            except:
                pass

st.divider()

# ==========================================
# 2. 枠数とカレンダー表示
# ==========================================
_, num_days = calendar.monthrange(year, month)

NIGHT_SHIFTS_UI = ['宿直A', '宿直B', '外来宿直']
DAY_SHIFTS_UI = ['日直A', '日直B', '外来日直']

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

st.subheader(f"📅 カレンダー確認（{month}月）")

cal_matrix = calendar.monthcalendar(year, month)
df_cal = pd.DataFrame(cal_matrix, columns=["月", "火", "水", "木", "金", "土", "日"])
df_cal = df_cal.astype(str).replace("0", "")

def color_calendar(val):
    if val == "":
        return ""
    d = int(val)
    date_obj = datetime.date(year, month, d)
    if date_obj.weekday() == 6 or jpholiday.is_holiday(date_obj) or (d in custom_holidays):
        return "color: #ff4b4b; font-weight: bold; background-color: #ffeeee;"
    elif date_obj.weekday() == 5:
        return "color: #1e90ff; font-weight: bold; background-color: #eef5ff;"
    return ""

if hasattr(df_cal.style, "map"):
    styled_cal = df_cal.style.map(color_calendar)
else:
    styled_cal = df_cal.style.applymap(color_calendar)

cal_height = len(df_cal) * 35 + 40
st.dataframe(styled_cal, use_container_width=True, hide_index=True, height=cal_height)
st.divider()

# ==========================================
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================
st.header("1. スタッフ条件の読み込み・入力（必須）")

template_data = {
    "先生の名前": ["Dr. A", "Dr. B", "Dr. C", "Dr. D", "Dr. E"],
    "NG日(半角カンマ区切り)": ["5,12,20", "10", "", "3,4,5", "25,26"],
    "希望日(半角カンマ区切り)": ["10:宿直A, 15:日直B", "", "8", "20", ""], 
    "希望優先度(数字が大きいほど優先)": [100, 1, 1, 1, 1], 
    "最低空ける日数": [5, 4, 6, 5, 3],  
    "月間最大回数": [5, 6, 4, 5, 7],    
    "宿直A上限": [2, 2, 2, 2, 2],
    "宿直B上限": [2, 2, 2, 2, 2],
    "外来宿直上限": [2, 2, 2, 2, 2],
    "日直A上限": [2, 2, 2, 2, 2],
    "日直B上限": [2, 2, 2, 2, 2],
    "外来日直上限": [2, 2, 2, 2, 2],
    "備考（メモ・説明など自由記入）": ["学会のためNG多め", "15日は午後休", "", "当直明け休み希望", ""]
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
    uploaded_file = st.file_uploader("スタッフ条件（CSV）をアップロード", type="csv", key="staff_csv")

if uploaded_file is not None:
    try:
        base_df = pd.read_csv(io.BytesIO(uploaded_file.getvalue()), encoding='shift_jis')
    except UnicodeDecodeError:
        base_df = pd.read_csv(io.BytesIO(uploaded_file.getvalue()), encoding='utf-8')
else:
    base_df = df_template.copy()

if "先生の名前" in base_df.columns:
    base_df = base_df.set_index("先生の名前")

st.markdown("##### 👩‍⚕️ スタッフ条件の入力・編集")
st.write("※以下の表は**直接クリックして文字を入力**できます。")
edited_df = st.data_editor(base_df, num_rows="dynamic", use_container_width=True, height=300)

staff_df = edited_df.reset_index()

st.divider()

st.header("2. 過去・決定済みシフトの読み込み・入力（任意）")
st.markdown("""
先月分や来月分のシフト表、または今月の「一部だけ人間が確定させたシフト表」があればアップロードしてください。
前後の月のシフト間隔を考慮したり、確定済みの枠を固定して残りの空白をAIに計算させることができます。
""")

fixed_file = st.file_uploader("過去・決定済みシフト表（CSV）をアップロード", type="csv", key="fixed_csv")

fixed_columns = ["日付", "区分", "宿直A", "宿直B", "外来宿直", "日直A", "日直B", "外来日直"]
if fixed_file is not None:
    try:
        f_bytes = fixed_file.getvalue()
        try:
            base_fixed_df = pd.read_csv(io.BytesIO(f_bytes), encoding='shift_jis')
        except UnicodeDecodeError:
            base_fixed_df = pd.read_csv(io.BytesIO(f_bytes), encoding='utf-8')
    except Exception as e:
        st.warning(f"過去シフトファイルの読み込みに失敗しました。詳細: {e}")
        base_fixed_df = pd.DataFrame(columns=fixed_columns)
else:
    # ファイルがない場合は空のデータフレームを作り、1行目だけ空欄の行を用意しておく
    base_fixed_df = pd.DataFrame(columns=fixed_columns)
    base_fixed_df.loc[0] = ["" for _ in range(len(fixed_columns))]

if "日付" in base_fixed_df.columns:
    base_fixed_df = base_fixed_df.set_index("日付")

st.markdown("##### 📅 決定済みシフトの入力・編集")
st.write("※CSVを使わずに、下の表へ直接クリックして「4/1」のように日付と先生の名前を手打ちすることもできます。")
edited_fixed_df_raw = st.data_editor(base_fixed_df, num_rows="dynamic", use_container_width=True, height=200)

# 計算用に見えないところで元の形に戻す
edited_fixed_df = edited_fixed_df_raw.reset_index()
# ======================================================================================

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
        if pd.isna(row['NG日(半角カンマ区切り)']) or ng_str.strip() == "" or ng_str == "nan":
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
            if not (pd.isna(row['希望日(半角カンマ区切り)']) or req_str.strip() == "" or req_str == "nan"):
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
            model.Add(sum(worked_all) <= actual_max_total)

    for doc in doctors:
        interval = min_intervals[doc]
        if interval > 0:
            for start_d in range(1, num_days + 1):
                current_date = datetime.date(target_year, target_month, start_d)
                all_abs_dates = absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]]
                
                if start_d not in all_abs_dates:
                    for past_date in past_worked_dates[doc]:
                        if 0 < (current_date - past_date).days <= interval:
                            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, start_d) else NIGHT_SHIFTS
                            for s in active_shifts:
                                model.Add(shifts[(start_d, doc, s)] == 0)
                                
                    for future_date in future_worked_dates[doc]:
                        if 0 < (future_date - current_date).days <= interval:
                            active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, start_d) else NIGHT_SHIFTS
                            for s in active_shifts:
                                model.Add(shifts[(start_d, doc, s)] == 0)

                window_days = range(start_d, min(start_d + interval + 1, num_days + 1))
                if any(d in all_abs_dates for d in window_days):
                    continue
                
                window_shifts = []
                for d in window_days:
                    active_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
                    for s in active_shifts:
                        window_shifts.append(shifts[(d, doc, s)])
                if window_shifts:
                    model.Add(sum(window_shifts) <= 1)

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
            row = {"日付": f"{target_month}/{d}({weekday_ja[date_obj.weekday()]})", "区分": day_str}
            
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

# === ▼NEW：画面入力された「決定済みシフト」の空行を取り除く処理▼ ===
fixed_df = edited_fixed_df[edited_fixed_df['日付'].astype(str).str.strip() != '']
fixed_df = fixed_df.dropna(subset=['日付']).reset_index(drop=True)
# ====================================================================

if len(staff_df) > 0 and st.button("🚀 このデータでシフトを自動生成する", type="primary"):
    with st.spinner("AIが最適なシフトを計算中...（最大45秒かかります）"):
        try:
            df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df)
            
            if success:
                st.success("✨ シフトの作成に成功しました！個人のルール（間隔・回数）を厳守し、優先度100以上の絶対希望や確定シフトは全て確約されています。")
                
                def highlight_holidays(row):
                    styles = [''] * len(row)
                    if row['区分'] == '休日':
                        for i, col in enumerate(row.index):
                            if col in ['日付', '区分']: 
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
                        if not (pd.isna(row['希望日(半角カンマ区切り)']) or req_str.strip() == "" or req_str == "nan"):
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
                        hol_count += sum(1 for val in df_result[df_result['区分'] == '休日'][s] if doc in [x.strip() for x in re.split(r'[、,]', str(val))])
                                
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
                
                styled_summary = df_summary.style.format(
                    {"最小間隔": "{:.0f}", "平均間隔": "{:.1f}"}, na_rep="-"
                ).set_properties(
                    subset=['総合計'], **{'font-weight': 'bold'}
                ).set_properties(
                    subset=['希望日の達成'], **{'text-align': 'center'}
                )
                
                summary_height = len(df_summary) * 35 + 40
                st.dataframe(styled_summary, use_container_width=True, hide_index=True, height=summary_height)
                
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
