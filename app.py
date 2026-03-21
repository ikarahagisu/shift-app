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

with st.expander("📖 初めての方へ：このアプリの特徴と使い方（クリックして開く）", expanded=False):
    st.markdown("""
    このアプリは、各先生の希望や複雑な勤務ルールを考慮し、AIが自動で最適な当直・日直シフトを作成するシステムです。

    ### ✨ アプリの特徴
    * **スマホでも快適操作**: カレンダーからのNG日選択など、スマホからでもレイアウトが崩れずサクサク入力できます。
    * **柔軟なシフト要件**: 「平日を休日扱いにしたい」「GWだけ2名体制に増員したい」などのイレギュラーにも対応。
    * **月を跨いだ間隔調整**: 先月の確定シフトを読み込ませることで、月末月初を跨いだ「最低空ける日数」も正確に計算します。

    ### 🧠 AIがシフトを組むときのルール
    AIは以下のルールを守りながら、何万通りもの組み合わせの中から最適なパズルを解き明かします。
    * **【絶対に守るルール】**: 「NG日には入れない」「最低空ける日数を守る（連投禁止）」「月間の最小・最大回数、休日最大回数、各枠の上限回数を守る」「優先度100以上の絶対希望日は確実に通す」。
    * **【なるべく叶えるルール】**: 「土日祝日の勤務回数が一部の人に偏らないよう、全体で公平に分担する（最優先）」「優先度付きの希望日を可能な限り多く通す」。

    ---
    ### 📝 使い方ステップ
    
    #### 1. カレンダーと特別設定（画面上部）
    * **年月の設定**: 作成したいシフトの年と月を選択します。
    * **特別休日の設定**: 平日でも日直が必要な日（年末年始やお盆など）は、カレンダーのチェックボックスをオンにして「休日扱い」にしてください。
    * **複数人シフト（増員）の設定**: 特定のシフトを「2名以上」に増やしたい場合は、表で日付と枠を指定して人数を変更します。

    #### 2. スタッフ条件の読み込み・入力（必須）
    まずは「📥 ひな形（CSV）」をダウンロードしてExcelで基本情報を入力し、アップロードするのがおすすめです。もちろん、画面上の表を直接クリックして編集することも可能です。
    
    💡 **ポイント**: 「NG日」の設定は、CSVで数字を入力するよりも、**後からWEBアプリ上のカレンダーでポチポチ直感的にクリックする方が圧倒的に楽**です！CSVでは空欄にしておくことをお勧めします。

    * **【入りにくい曜日】** 「水,木」のように入力しておくと、NGカレンダーでその曜日に ⚠️マーク が表示され、休みたい日を選ぶ際の目印になります。
    * **【NG日】** 先生ごとにタブを切り替え、カレンダーから休みたい日を選んで、最後に赤い**【✨ (先生の名前)のNG日を確定する】**ボタンを押します。
    * **【希望日】** 入りたい日を半角カンマ区切りで入力します。
        * 日付だけ指定（例: `10, 15`）→ その日のどれかのシフトに入ります。
        * 枠まで指定（例: `10:宿直A, 15:日直B`）→ その日のその枠を狙います。
    * **【希望優先度】** 希望を通す「強さ」です。基本は `1` です。どうしても外せない希望がある場合は `100` 以上の数字を入れると、回数上限などのルールを無視して**【確実】**にそのシフトに入ります。
    * **【休日最大回数】** 土日祝日（特別休日含む）に入るシフトの上限回数です。人ごとに休日の負担をコントロールできます。
    * **【各種ルール】** 他にも「最低空ける日数（シフト間隔）」「月間最小/最大回数」「各シフト枠の上限」を設定できます。

    #### 3. 過去・決定済みシフトの入力（任意）
    先月分のシフト表や、今月の「一部だけ確定させたシフト」があれば入力します。
    * 先月のデータを入れれば、前月末の勤務を考慮して月初の間隔ルールをしっかり守ります。
    * 今月の確定データを入れれば、残りの空き枠だけをAIが綺麗に埋めてくれます。

    #### 4. シフトの自動生成
    設定が終わったら、一番下の「🚀 このデータでシフトを自動生成する」ボタンを押します。

    💡 **ポイント**: 
    * ボタンを押すたびに、AIが少しずつ違うパターンのシフトを提案してくれます。
    * 条件が厳しすぎてシフトが組めない場合は、エラーメッセージが出ます。その場合は、各先生の「最大回数」を増やしたり、「最低空ける日数」を減らしたりして条件を少し緩めてみてください。
    * 完成したシフト表や、勤務回数の実績レポートはCSVでダウンロードできます。
    """)

# === 🌟改修：スマホ＆フォーム内で絶対に崩れないカレンダー用CSS ===
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
    padding: 8px 0px !important; 
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: flex-start !important; 
    background-color: #ffffff;
    overflow: hidden; 
}

/* Streamlit特有の余計なマージンを消去して高さを統一 */
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

/* チェックボックスをセルの真ん中に配置 */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) div[data-testid="stCheckbox"] {
    display: flex;
    justify-content: center !important;
    align-items: flex-start !important; 
    width: 100% !important;
}

/* 四角と文字を「縦並び」にする */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(7)) div[data-testid="stCheckbox"] label {
    display: flex !important;
    flex-direction: column-reverse !important; 
    justify-content: flex-start !important;
    align-items: center !important;
    width: 100% !important;
    margin: 0 auto !important; 
    padding: 0 !important;
    gap: 6px !important; 
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
    "入りにくい曜日(半角カンマ区切り)": ["水,木", "", "土,日", "", ""],
    "NG日(半角カンマ区切り)": ["", "", "", "", ""],
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
    uploaded_file = st.file_uploader("スタッフ条件（途中保存CSVも可）をアップロード", type="csv", key="staff_csv")

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

# エラー対策：データ型の統一
if "希望優先度(数字が大きいほど優先)" in base_df.columns:
    base_df["希望優先度(数字が大きいほど優先)"] = pd.to_numeric(base_df["希望優先度(数字が大きいほど優先)"], errors='coerce')

text_cols = ["入りにくい曜日(半角カンマ区切り)", "NG日(半角カンマ区切り)", "希望日(半角カンマ区切り)", "備考（メモ・説明など自由記入）"]
for c in text_cols:
    if c in base_df.columns:
        base_df[c] = base_df[c].apply(lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "none", "<na>"] else str(x))

st.markdown("##### 👩‍⚕️ スタッフ条件の入力・編集")
st.write("※以下の表は直接クリックして文字を入力できます。（ヘッダーの列名にマウスを合わせると入力のヒントが出ます）")

edited_df = st.data_editor(
    base_df, 
    num_rows="dynamic", 
    use_container_width=True, 
    height=300,
    column_config={
        "入りにくい曜日(半角カンマ区切り)": st.column_config.TextColumn(
            "入りにくい曜日",
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

st.markdown("##### 🚫 先生ごとのNG日設定（カレンダーでクリック選択）")

valid_staff = staff_df[staff_df["先生の名前"].astype(str).str.strip() != ""]
if not valid_staff.empty:
    doctor_names = valid_staff["先生の名前"].astype(str).tolist()
    tabs = st.tabs(doctor_names)
    
    for t_idx, doc_name in enumerate(doctor_names):
        original_idx = valid_staff.index[t_idx]
        with tabs[t_idx]:
            
            hard_str = str(valid_staff.loc[original_idx].get("入りにくい曜日(半角カンマ区切り)", ""))
            hard_days = []
            for i, w in enumerate(["月", "火", "水", "木", "金", "土", "日"]):
                if w in hard_str:
                    hard_days.append(i)

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
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = (d in current_ng_list)

            current_ngs = [d for d in range(1, num_days + 1) if st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", False)]
            
            if current_ngs:
                saved_dates_str = ", ".join([f"{d}日" for d in current_ngs])
                st.success(f"✅ **保存済みのNG日:** {saved_dates_str}")
            else:
                st.info("💡 **保存済みのNG日はありません**")

            with st.form(key=f"ng_form_{original_idx}", border=False):
                st.write(f"※カレンダーで休みたい日を複数選んだ後、最後に必ず下の赤い**【✨ {doc_name}先生のNG日を確定する】**ボタンを押して保存してください。")
                if hard_days:
                    st.markdown("<span style='color: #d97706; font-size: 0.9rem; font-weight: bold;'>💡 設定された「入りにくい曜日」には日付の横に ⚠️ マークが表示されています。休みたい場合はチェックを入れてください。</span>", unsafe_allow_html=True)

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
                                if is_hol_or_sun:
                                    day_label = f":red[**{day}日**] {warning_mark}"
                                elif is_sat:
                                    day_label = f":blue[**{day}日**] {warning_mark}"
                                else:
                                    day_label = f"**{day}日** {warning_mark}"
                                    
                                chk_key = f"ng_{doc_name}_{year}_{month}_{day}"
                                st.checkbox(day_label, key=chk_key)
                        else:
                            with cols[i]:
                                st.write("")
                
                submitted = st.form_submit_button(f"✨ {doc_name}先生のNG日を確定する", type="primary")
            
            _, col_btn1, col_btn2 = st.columns([6, 1.5, 1.5])
            with col_btn1:
                st.button("全選択", key=f"btn_all_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, True), use_container_width=True)
            with col_btn2:
                st.button("全解除", key=f"btn_clear_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, False), use_container_width=True)
            
            # 最新の状態を常に staff_df に反映
            current_ngs_str = [str(d) for d in range(1, num_days + 1) if st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", False)]
            staff_df.at[original_idx, "NG日(半角カンマ区切り)"] = ",".join(current_ngs_str)
            
            if submitted:
                st.toast(f"✅ {doc_name}先生のNG日を保存しました！")

st.divider()
st.markdown("##### 📂 入力途中のデータを一時保存（後で再開したい場合）")
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
        
        abs_hol_count = sum(1 for d in absolute_req_days[doc] if is_holiday(target_year, target_month, d))
        abs_hol_count += sum(1 for d, s in absolute_req_specific[doc] if is_holiday(target_year, target_month, d))
        actual_hol_max = max(max_hol_shifts_per_doc[doc], abs_hol_count) 
        model.Add(holiday_worked[doc] <= actual_hol_max)
        
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
        model.Maximize(sum(objective_terms) - max_hol_shifts * 1000)
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

        req_hol_total = 0
        for d in range(1, num_days + 1):
            if is_holiday(target_year, target_month, d):
                active_shifts = NIGHT_SHIFTS + DAY_SHIFTS
                req_hol_total += sum(multi_slots_dict.get((d, s), 1) for s in active_shifts)
        max_hol_available = sum(max_hol_shifts_per_doc[doc] for doc in doctors)
        if max_hol_available < req_hol_total:
            reasons.append(f"❌ **休日シフト枠の不足**: 月間に必要な休日の総枠数({req_hol_total}枠)に対して、先生全員の「休日最大回数」の合計({max_hol_available}回分)が足りていません。各人の休日最大回数を増やしてください。")

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
                
                shift_columns = ['宿直A', '宿直B', '外来宿直', '日直A', '日直B', '外来日直']
                doctors_list = staff_df['先生の名前'].astype(str).tolist()
                
                def highlight_holidays(row):
                    styles = [''] * len(row)
                    if row['平日/休日'] == '休日':
                        for i, col in enumerate(row.index):
                            if col in ['日付', '平日/休日']: 
                                styles[i] = 'color: #ff4b4b; font-weight: bold;'
                    return styles
                
                # ▼ 変更：40名以上の医師を区別するため、50色の異なるカラーパレットを用意 ▼
                distinct_colors = [
                    '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
                    '#800000', '#008000', '#000080', '#808000', '#800080', '#008080',
                    '#FFA500', '#FF1493', '#00FA9A', '#1E90FF', '#FF69B4', '#8A2BE2',
                    '#A52A2A', '#DEB887', '#5F9EA0', '#7FFF00', '#D2691E', '#FF7F50',
                    '#6495ED', '#DC143C', '#00008B', '#008B8B', '#B8860B', '#006400',
                    '#BDB76B', '#8B008B', '#556B2F', '#FF8C00', '#9932CC', '#8B0000',
                    '#E9967A', '#8FBC8F', '#483D8B', '#2F4F4F', '#00CED1', '#9400D3',
                    '#FFD700', '#DAA520', '#ADFF2F', '#CD5C5C', '#4B0082', '#F08080',
                    '#20B2AA', '#778899'
                ]
                doc_colors = {doc: distinct_colors[i % len(distinct_colors)] for i, doc in enumerate(doctors_list)}
                
                def color_doctors(val):
                    val_str = str(val)
                    if val_str == "-" or val_str == "":
                        return ''
                    for doc, color in doc_colors.items():
                        if doc in val_str:
                            return f'background-color: {color}; color: #000000; font-weight: 500;'
                    return ''
                
                base_style = df_result.style.apply(highlight_holidays, axis=1)
                if hasattr(base_style, 'map'):
                    styled_df = base_style.map(color_doctors, subset=shift_columns)
                else:
                    styled_df = base_style.applymap(color_doctors, subset=shift_columns)
                
                st.subheader("📅 完成したシフト表")
                
                legend_html = "<div style='display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 12px;'>"
                for doc, color in doc_colors.items():
                    legend_html += f"<div style='display: flex; align-items: center; gap: 4px;'><div style='width: 16px; height: 16px; background-color: {color}; border: 1px solid #ccc; border-radius: 3px;'></div><span style='font-size: 0.9rem;'>{doc}</span></div>"
                legend_html += "</div>"
                st.markdown(legend_html, unsafe_allow_html=True)
                
                result_height = len(df_result) * 35 + 40
                st.dataframe(styled_df, use_container_width=True, hide_index=True, height=result_height)
                
                st.subheader("📊 先生ごとのシフト回数（実績）")
                summary_list = []
                
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
                                
                    doc_data["宿直回数"] = doc_data.get("宿直A", 0) + doc_data.get("宿直B", 0) + doc_data.get("外来宿直", 0)
                    doc_data["日直回数"] = doc_data.get("日直A", 0) + doc_data.get("日直B", 0) + doc_data.get("外来日直", 0)
                    doc_data["休日回数"] = hol_count
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
                        doc_data["希望日達成"] = f"{granted} / {total_reqs} 回"
                    else:
                        doc_data["希望日達成"] = "-"
                    
                    summary_list.append(doc_data)
                    
                df_summary = pd.DataFrame(summary_list)
                df_summary = df_summary[['先生の名前', '宿直A', '宿直B', '外来宿直', '日直A', '日直B', '外来日直', '宿直回数', '日直回数', '休日回数', '総合計', '希望日達成', '最小間隔', '平均間隔']]
                
                df_summary = df_summary.set_index('先生の名前')
                
                styled_summary = df_summary.style.format(
                    {"最小間隔": "{:.0f}", "平均間隔": "{:.1f}"}, na_rep="-"
                ).set_properties(
                    subset=['総合計', '宿直回数', '日直回数'], **{'font-weight': 'bold'}
                ).set_properties(
                    subset=['希望日達成'], **{'text-align': 'center'}
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
