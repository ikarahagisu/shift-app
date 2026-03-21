import streamlit as st
import pandas as pd
import datetime
import calendar
import pulp
import re
import jpholiday
import io

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🏥 医師シフト自動作成アプリ")

# ==========================================
# 1. サイドバー：年月と枠数の設定
# ==========================================
st.sidebar.header("📅 作成する年月")
today = datetime.date.today()
year = st.sidebar.number_input("年", min_value=2020, max_value=2030, value=today.year)
month = st.sidebar.number_input("月", min_value=1, max_value=12, value=today.month)

st.sidebar.divider()
st.sidebar.header("🎌 独自の祝日・休日設定")
custom_holidays_str = st.sidebar.text_input("休日にする日付（半角カンマ区切り）", placeholder="例: 15, 16")
custom_holidays = []
if custom_holidays_str:
    for item in custom_holidays_str.split(','):
        try:
            day = int(item.strip())
            custom_holidays.append(datetime.date(year, month, day))
        except:
            pass

st.sidebar.divider()
st.sidebar.header("🔢 1日あたりの必要人数（枠数）")
shift_types = ['宿直A', '宿直B', '外来宿直', '日直A', '日直B', '外来日直']
multi_slots_dict = {}

for s in shift_types:
    st.sidebar.subheader(f"【{s}】")
    c1, c2 = st.sidebar.columns(2)
    wd_val = c1.number_input(f"{s} (平日)", min_value=0, max_value=5, value=1 if '宿直' in s else 0, key=f"wd_{s}")
    we_val = c2.number_input(f"{s} (休日)", min_value=0, max_value=5, value=1, key=f"we_{s}")
    multi_slots_dict[s] = {'平日': wd_val, '休日': we_val}

# ==========================================
# 2. シフト生成関数（エラー原因特定機能つき）
# ==========================================
def generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df):
    start_date = datetime.date(year, month, 1)
    if month == 12:
        end_date = datetime.date(year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        end_date = datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)
    
    dates = [start_date + datetime.timedelta(days=i) for i in range((end_date - start_date).days + 1)]
    days_in_month = len(dates)
    
    cal_data = []
    for d in dates:
        day_type = '平日'
        if d.weekday() >= 5 or d in custom_holidays or jpholiday.is_holiday(d):
            day_type = '休日'
        cal_data.append({'日付': d.strftime('%Y/%m/%d'), '曜日': ['月', '火', '水', '木', '金', '土', '日'][d.weekday()], '平日/休日': day_type})
    df_cal = pd.DataFrame(cal_data)
    
    doctors = staff_df['先生の名前'].tolist()
    num_doctors = len(doctors)
    
    doc_conditions = {}
    for index, row in staff_df.iterrows():
        doc = row['先生の名前']
        max_shifts = row['月間最大回数(回)'] if not pd.isna(row['月間最大回数(回)']) else 99
        min_interval = row['勤務間隔(日)'] if not pd.isna(row['勤務間隔(日)']) else 0
        
        req_dates = []
        req_shifts = []
        ng_dates = []
        ng_shifts = []
        
        if '希望日(半角カンマ区切り)' in staff_df.columns:
            req_str = str(row['希望日(半角カンマ区切り)'])
            if req_str.lower() not in ["nan", "none", ""]:
                req_str = req_str.replace('：', ':')
                for item in req_str.split(','):
                    item = item.strip()
                    if not item: continue
                    if ':' in item:
                        parts = item.split(':')
                        try:
                            req_shifts.append((int(re.sub(r'\D', '', parts[0].strip())), parts[1].strip()))
                        except: pass
                    else:
                        try:
                            req_dates.append(int(item))
                        except: pass
        
        if 'NG日(半角カンマ区切り)' in staff_df.columns:
            ng_str = str(row['NG日(半角カンマ区切り)'])
            if ng_str.lower() not in ["nan", "none", ""]:
                ng_str = ng_str.replace('：', ':')
                for item in ng_str.split(','):
                    item = item.strip()
                    if not item: continue
                    if ':' in item:
                        parts = item.split(':')
                        try:
                            ng_shifts.append((int(re.sub(r'\D', '', parts[0].strip())), parts[1].strip()))
                        except: pass
                    else:
                        try:
                            ng_dates.append(int(item))
                        except: pass
        
        doc_conditions[doc] = {
            'max_shifts': max_shifts,
            'min_interval': min_interval,
            'req_dates': req_dates,
            'req_shifts': req_shifts,
            'ng_dates': ng_dates,
            'ng_shifts': ng_shifts,
            'max_syuku_a': row.get('宿直A上限', 99) if pd.notna(row.get('宿直A上限')) else 99,
            'max_syuku_b': row.get('宿直B上限', 99) if pd.notna(row.get('宿直B上限')) else 99,
            'max_gai_syuku': row.get('外来宿直上限', 99) if pd.notna(row.get('外来宿直上限')) else 99,
            'max_nichi_a': row.get('日直A上限', 99) if pd.notna(row.get('日直A上限')) else 99,
            'max_nichi_b': row.get('日直B上限', 99) if pd.notna(row.get('日直B上限')) else 99,
            'max_gai_nichi': row.get('外来日直上限', 99) if pd.notna(row.get('外来日直上限')) else 99,
            'priority': row.get('希望優先度', 1) if pd.notna(row.get('希望優先度')) else 1
        }

    prob = pulp.LpProblem("Shift_Scheduling", pulp.LpMaximize)
    
    x = pulp.LpVariable.dicts("shift",
                              ((d, s, doc) for d in range(days_in_month) for s in shift_types for doc in doctors),
                              cat='Binary')

    # 各枠の必要人数を満たす
    for d in range(days_in_month):
        day_type = df_cal.iloc[d]['平日/休日']
        for s in shift_types:
            req_num = multi_slots_dict.get(s, {}).get(day_type, 0)
            date_str = df_cal.iloc[d]['日付']
            fixed_assigned = []
            if len(fixed_df) > 0:
                fixed_row = fixed_df[fixed_df['日付'] == date_str]
                if not fixed_row.empty and s in fixed_row.columns:
                    val = str(fixed_row.iloc[0][s])
                    if val != "nan" and val != "":
                        fixed_assigned = [doc.strip() for doc in re.split(r'[、,]', val)]
            for fixed_doc in fixed_assigned:
                if fixed_doc in doctors:
                    prob += x[(d, s, fixed_doc)] == 1
            prob += pulp.lpSum(x[(d, s, doc)] for doc in doctors) == req_num

    # 先生ごとの制約
    for doc in doctors:
        cond = doc_conditions[doc]
        prob += pulp.lpSum(x[(d, s, doc)] for d in range(days_in_month) for s in shift_types) <= cond['max_shifts']
        prob += pulp.lpSum(x[(d, '宿直A', doc)] for d in range(days_in_month)) <= cond['max_syuku_a']
        prob += pulp.lpSum(x[(d, '宿直B', doc)] for d in range(days_in_month)) <= cond['max_syuku_b']
        prob += pulp.lpSum(x[(d, '外来宿直', doc)] for d in range(days_in_month)) <= cond['max_gai_syuku']
        prob += pulp.lpSum(x[(d, '日直A', doc)] for d in range(days_in_month)) <= cond['max_nichi_a']
        prob += pulp.lpSum(x[(d, '日直B', doc)] for d in range(days_in_month)) <= cond['max_nichi_b']
        prob += pulp.lpSum(x[(d, '外来日直', doc)] for d in range(days_in_month)) <= cond['max_gai_nichi']
        
        for d in range(days_in_month):
            prob += pulp.lpSum(x[(d, s, doc)] for s in shift_types) <= 1
            
        min_int = cond['min_interval']
        if min_int > 0:
            for d in range(days_in_month):
                for offset in range(1, min_int + 1):
                    if d + offset < days_in_month:
                        prob += pulp.lpSum(x[(d, s, doc)] for s in shift_types) + pulp.lpSum(x[(d + offset, s, doc)] for s in shift_types) <= 1
            
        for req_d in cond['ng_dates']:
            if 1 <= req_d <= days_in_month:
                prob += pulp.lpSum(x[(req_d - 1, s, doc)] for s in shift_types) == 0
        for req_d, req_s in cond['ng_shifts']:
            if 1 <= req_d <= days_in_month and req_s in shift_types:
                prob += x[(req_d - 1, req_s, doc)] == 0

        if cond['priority'] >= 100:
            for req_d in cond['req_dates']:
                if 1 <= req_d <= days_in_month:
                    prob += pulp.lpSum(x[(req_d - 1, s, doc)] for s in shift_types) >= 1
            for req_d, req_s in cond['req_shifts']:
                if 1 <= req_d <= days_in_month and req_s in shift_types:
                    prob += x[(req_d - 1, req_s, doc)] == 1

    # 目的関数
    obj_terms = []
    for doc in doctors:
        cond = doc_conditions[doc]
        weight = cond['priority']
        if weight < 100:
            for req_d in cond['req_dates']:
                if 1 <= req_d <= days_in_month:
                    obj_terms.append(weight * pulp.lpSum(x[(req_d - 1, s, doc)] for s in shift_types))
            for req_d, req_s in cond['req_shifts']:
                if 1 <= req_d <= days_in_month and req_s in shift_types:
                    obj_terms.append(weight * x[(req_d - 1, req_s, doc)])
    prob += pulp.lpSum(obj_terms)

    status = prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=45))
    
    # ▼ エラー時のボトルネック特定機能 ▼
    if pulp.LpStatus[status] != 'Optimal':
        error_reasons = []
        prob_relax = pulp.LpProblem("Shift_Relaxed", pulp.LpMaximize)
        x_r = pulp.LpVariable.dicts("shift_r", ((d, s, doc) for d in range(days_in_month) for s in shift_types for doc in doctors), cat='Binary')
        dummy = pulp.LpVariable.dicts("dummy", ((d, s) for d in range(days_in_month) for s in shift_types), lowBound=0, cat='Continuous')
        
        for d in range(days_in_month):
            day_type = df_cal.iloc[d]['平日/休日']
            for s in shift_types:
                req_num = multi_slots_dict.get(s, {}).get(day_type, 0)
                date_str = df_cal.iloc[d]['日付']
                fixed_assigned = []
                if len(fixed_df) > 0:
                    fixed_row = fixed_df[fixed_df['日付'] == date_str]
                    if not fixed_row.empty and s in fixed_row.columns:
                        val = str(fixed_row.iloc[0][s])
                        if val != "nan" and val != "":
                            fixed_assigned = [doc.strip() for doc in re.split(r'[、,]', val)]
                for fixed_doc in fixed_assigned:
                    if fixed_doc in doctors:
                        prob_relax += x_r[(d, s, fixed_doc)] == 1
                prob_relax += pulp.lpSum(x_r[(d, s, doc)] for doc in doctors) + dummy[(d, s)] == req_num
                
        for doc in doctors:
            cond = doc_conditions[doc]
            prob_relax += pulp.lpSum(x_r[(d, s, doc)] for d in range(days_in_month) for s in shift_types) <= cond['max_shifts']
            for d in range(days_in_month):
                prob_relax += pulp.lpSum(x_r[(d, s, doc)] for s in shift_types) <= 1
            min_int = cond['min_interval']
            if min_int > 0:
                for d in range(days_in_month):
                    for offset in range(1, min_int + 1):
                        if d + offset < days_in_month:
                            prob_relax += pulp.lpSum(x_r[(d, s, doc)] for s in shift_types) + pulp.lpSum(x_r[(d + offset, s, doc)] for s in shift_types) <= 1
            for req_d in cond['ng_dates']:
                if 1 <= req_d <= days_in_month:
                    prob_relax += pulp.lpSum(x_r[(req_d - 1, s, doc)] for s in shift_types) == 0
            for req_d, req_s in cond['ng_shifts']:
                if 1 <= req_d <= days_in_month and req_s in shift_types:
                    prob_relax += x_r[(req_d - 1, req_s, doc)] == 0
            if cond['priority'] >= 100:
                for req_d in cond['req_dates']:
                    if 1 <= req_d <= days_in_month:
                        prob_relax += pulp.lpSum(x_r[(req_d - 1, s, doc)] for s in shift_types) >= 1
                for req_d, req_s in cond['req_shifts']:
                    if 1 <= req_d <= days_in_month and req_s in shift_types:
                        prob_relax += x_r[(req_d - 1, req_s, doc)] == 1

        prob_relax += -10000 * pulp.lpSum(dummy[(d, s)] for d in range(days_in_month) for s in shift_types)
        prob_relax.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=20))
        
        bottlenecks = []
        for d in range(days_in_month):
            for s in shift_types:
                if dummy[(d, s)].varValue and dummy[(d, s)].varValue > 0.1:
                    bottlenecks.append(f"・{d+1}日の「{s}」（あと {int(round(dummy[(d, s)].varValue))} 人足りません）")
        
        if bottlenecks:
            error_reasons.append("🚨 **以下のシフト枠を埋める人が見つかりませんでした。**")
            error_reasons.append("（※間隔ルール、NG日、他シフトとの被り、回数上限などが原因です）")
            error_reasons.extend(bottlenecks)
        else:
            error_reasons.append("条件が複雑すぎて（絶対希望の衝突など）パズルが解けませんでした。優先度100の希望や固定シフトを見直してください。")
            
        return None, False, error_reasons, {}, {}

    df_shift = df_cal.copy()
    for s in shift_types:
        df_shift[s] = ""
        
    for d in range(days_in_month):
        for s in shift_types:
            assigned = [doc for doc in doctors if x[(d, s, doc)].varValue == 1.0]
            if assigned:
                df_shift.at[d, s] = ", ".join(assigned)
            else:
                df_shift.at[d, s] = "-"

    return df_shift, True, [], {}, {}

# ==========================================
# 3. 先生の条件・希望入力
# ==========================================
st.header("1. 先生の条件・希望入力")
uploaded_file = st.file_uploader("CSVファイルをアップロード（任意）", type=['csv'])

default_columns = ['先生の名前', '月間最大回数(回)', '勤務間隔(日)', '希望日(半角カンマ区切り)', 'NG日(半角カンマ区切り)', '希望優先度', '宿直A上限', '宿直B上限', '外来宿直上限', '日直A上限', '日直B上限', '外来日直上限']
if uploaded_file is not None:
    try:
        staff_df = pd.read_csv(uploaded_file, encoding='shift_jis')
    except:
        staff_df = pd.read_csv(uploaded_file, encoding='utf-8')
    for col in default_columns:
        if col not in staff_df.columns:
            staff_df[col] = None
else:
    staff_df = pd.DataFrame(columns=default_columns)
    for i in range(5):
        staff_df.loc[i] = ["" if col == '先生の名前' else None for col in default_columns]

staff_df = st.data_editor(staff_df, num_rows="dynamic", use_container_width=True)

# ==========================================
# 4. 確定シフトの入力
# ==========================================
st.header("2. 確定シフトの入力（任意）")
st.markdown("あらかじめ決まっているシフトがあれば入力してください。（例：1日の宿直Aに「安藤」）")

start_date = datetime.date(year, month, 1)
if month == 12:
    end_date = datetime.date(year + 1, 1, 1) - datetime.timedelta(days=1)
else:
    end_date = datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)

dates = [start_date + datetime.timedelta(days=i) for i in range((end_date - start_date).days + 1)]
cal_data = []
for d in dates:
    cal_data.append({'日付': d.strftime('%Y/%m/%d'), '曜日': ['月', '火', '水', '木', '金', '土', '日'][d.weekday()]})
fixed_df_base = pd.DataFrame(cal_data)
for s in shift_types:
    fixed_df_base[s] = ""

edited_fixed_df = st.data_editor(fixed_df_base, use_container_width=True, hide_index=True)

# ==========================================
# 5. 実行ボタンと結果表示
# ==========================================
st.divider()
st.header("3. シフトの自動生成")

staff_df = staff_df[staff_df['先生の名前'].astype(str).str.strip() != '']
staff_df = staff_df.dropna(subset=['先生の名前']).reset_index(drop=True)

fixed_df = edited_fixed_df[edited_fixed_df['日付'].astype(str).str.strip() != '']
fixed_df = fixed_df.dropna(subset=['日付']).reset_index(drop=True)

if len(staff_df) > 0:
    if st.button("🚀 このデータでシフトを自動生成する", type="primary"):
        with st.spinner("AIが最適なシフトを計算中...（最大45秒かかります）"):
            try:
                df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df)
                
                if success:
                    st.session_state['generated_df'] = df_result
                    st.session_state['past_worked_dates'] = past_worked_dates
                    st.session_state['future_worked_dates'] = future_worked_dates
                    st.success("✨ シフトの作成に成功しました！個人のルール（間隔・回数）を厳守し、優先度100以上の絶対希望や確定シフトは全て確約されています。")
                else:
                    if 'generated_df' in st.session_state:
                        del st.session_state['generated_df']
                    st.error("入力された条件に誤りがあるか、条件が厳しすぎてシフトが組めませんでした。")
                    st.warning("💡 **以下の原因が考えられます。Excelの入力や設定画面を見直してください。**")
                    for reason in error_reasons:
                        st.write(reason)
            except Exception as e:
                st.error(f"シフト計算中にエラーが発生しました。詳細: {e}")

    if 'generated_df' in st.session_state:
        df_result = st.session_state['generated_df']
        past_worked_dates = st.session_state.get('past_worked_dates', {})
        future_worked_dates = st.session_state.get('future_worked_dates', {})
        
        doctors_list = staff_df['先生の名前'].astype(str).tolist()
        
        st.subheader("📅 完成したシフト表")
        
        table_container = st.container()
        
        st.divider()
        st.markdown("<span style='font-size: 0.95rem; font-weight: bold;'>🔍 特定の先生のシフトを色別でハイライト</span>", unsafe_allow_html=True)
        st.write("※各色のすぐ下にあるメモ欄に「神経内科」「呼吸器内科」など自由に書き込めます。")
        
        empty_lbl = "\u200B"
        
        # 1段目
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("🟨 **黄色**")
            hl_yellow = st.multiselect("黄色", options=doctors_list, default=[], key="hl_yellow", label_visibility="collapsed")
            st.text_input("黄色メモ", key="memo_y", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
        with c2:
            st.markdown("🟥 **赤色**")
            hl_red = st.multiselect("赤色", options=doctors_list, default=[], key="hl_red", label_visibility="collapsed")
            st.text_input("赤色メモ", key="memo_r", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
        st.write("")
        
        # 2段目
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("🟦 **水色**")
            hl_blue = st.multiselect("水色", options=doctors_list, default=[], key="hl_blue", label_visibility="collapsed")
            st.text_input("水色メモ", key="memo_b", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
        with c2:
            st.markdown("🟩 **緑色**")
            hl_green = st.multiselect("緑色", options=doctors_list, default=[], key="hl_green", label_visibility="collapsed")
            st.text_input("緑色メモ", key="memo_g", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
        st.write("")
        
        # 3段目
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("🟧 **オレンジ**")
            hl_orange = st.multiselect("オレンジ", options=doctors_list, default=[], key="hl_orange", label_visibility="collapsed")
            st.text_input("オレンジメモ", key="memo_o", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
        with c2:
            st.markdown("🟫 **茶色**")
            hl_brown = st.multiselect("茶色", options=doctors_list, default=[], key="hl_brown", label_visibility="collapsed")
            st.text_input("茶色メモ", key="memo_br", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
        st.write("")
        
        # 4段目
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("🟪 **紫色**")
            hl_purple = st.multiselect("紫色", options=doctors_list, default=[], key="hl_purple", label_visibility="collapsed")
            st.text_input("紫色メモ", key="memo_p", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
        with c2:
            st.markdown("💗 **ピンク**")
            hl_pink = st.multiselect("ピンク", options=doctors_list, default=[], key="hl_pink", label_visibility="collapsed")
            st.text_input("ピンクメモ", key="memo_pi", placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
            
        st.write("")

        def highlight_holidays(row):
            styles = [''] * len(row)
            if row['平日/休日'] == '休日':
                for i, col in enumerate(row.index):
                    if col in ['日付', '平日/休日']: 
                        styles[i] = 'color: #ff4b4b; font-weight: bold;'
            return styles
        
        def color_highlighted_doctor(val):
            val_str = str(val)
            if val_str == "-" or val_str == "":
                return ''
            
            cell_docs = [d.strip() for d in re.split(r'[、,]', val_str)]
            
            for doc in cell_docs:
                if doc in hl_yellow:
                    return 'background-color: #fff200; color: #000000; font-weight: bold; border: 2px solid #ffcc00;'
                elif doc in hl_red:
                    return 'background-color: #ffcccc; color: #000000; font-weight: bold; border: 2px solid #ff6666;'
                elif doc in hl_blue:
                    return 'background-color: #cce5ff; color: #000000; font-weight: bold; border: 2px solid #66b2ff;'
                elif doc in hl_green:
                    return 'background-color: #ccffcc; color: #000000; font-weight: bold; border: 2px solid #66ff66;'
                elif doc in hl_orange:
                    return 'background-color: #ffe5b4; color: #000000; font-weight: bold; border: 2px solid #ffb347;'
                elif doc in hl_brown:
                    return 'background-color: #e6ccb3; color: #000000; font-weight: bold; border: 2px solid #c68c53;'
                elif doc in hl_purple:
                    return 'background-color: #e6ccff; color: #000000; font-weight: bold; border: 2px solid #b366ff;'
                elif doc in hl_pink:
                    return 'background-color: #ffccff; color: #000000; font-weight: bold; border: 2px solid #ff66ff;'
            return ''
        
        base_style = df_result.style.apply(highlight_holidays, axis=1)
        
        if hasattr(base_style, 'map'):
            styled_df = base_style.map(color_highlighted_doctor, subset=shift_types)
        else:
            styled_df = base_style.applymap(color_highlighted_doctor, subset=shift_types)
        
        result_height = len(df_result) * 35 + 40
        
        with table_container:
            st.dataframe(styled_df, use_container_width=True, hide_index=True, height=result_height)
        
        st.divider()
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
                for s in shift_types:
                    cell_val = str(row[s])
                    if doc in [x.strip() for x in re.split(r'[、,]', cell_val)]:
                        is_working = True
                        break
                if is_working:
                    doc_working_dates.add(datetime.date(year, month, d_idx + 1))
            
            for s in shift_types:
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

elif len(staff_df) == 0:
    st.warning("☝️ 表に先生の名前を入力するか、CSVファイルをアップロードしてください。")
