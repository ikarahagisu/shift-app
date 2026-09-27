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
def _read_csv_any_encoding(file_bytes):
    """
    Shift-JIS / UTF-8 BOM / UTF-8 / latin-1 の順に試してDataFrameを返す。
    io.BytesIO はread後にポインタが末尾へ移動するため、
    エンコーディングごとに必ず新しいインスタンスを生成する。
    latin-1 は任意の1バイト列を読めるため実質フォールバックとして機能する。
    """
    for encoding in ('cp932', 'shift_jis', 'utf-8-sig', 'utf-8', 'latin-1'):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSVのエンコーディングを判定できませんでした。")

def normalize_shift_names(df):
    """以前のCSVの枠名・上限列・希望枠を現在の名称へ変換する。"""
    aliases = {"宿直A": "A宿直", "宿直B": "B宿直", "A当直": "A宿直", "B当直": "B宿直", "日直A": "A日直", "日直B": "B日直"}
    for old, new in aliases.items():
        for suffix in ("", "上限"):
            old_col, new_col = old + suffix, new + suffix
            if old_col not in df.columns:
                continue
            if new_col not in df.columns:
                df = df.rename(columns={old_col: new_col})
            else:
                blank = df[new_col].fillna("").astype(str).str.strip().eq("")
                df.loc[blank, new_col] = df.loc[blank, old_col]
                df = df.drop(columns=[old_col])
    request_col = "希望日(半角カンマ区切り)"
    if request_col in df.columns:
        def normalize_request(value):
            if not isinstance(value, str):
                return value
            for old, new in aliases.items():
                value = value.replace(old, new)
            return value
        df[request_col] = df[request_col].map(normalize_request)
    return df

@st.cache_data
def parse_staff_csv(file_bytes):
    df = normalize_shift_names(_read_csv_any_encoding(file_bytes))
    # 旧CSVの列名も受け付け、画面・出力CSVでは新名称に統一する。
    weekday_column = "翌日PM duty"
    for old_column in ("原則、宿直を外す曜日", "入れない曜日(半角カンマ区切り)", "入れない曜日"):
        if old_column not in df.columns:
            continue
        if weekday_column not in df.columns:
            df = df.rename(columns={old_column: weekday_column})
        else:
            # 両方の列がある場合は新名称の値を優先し、空欄だけ旧列で補う。
            blank = df[weekday_column].fillna("").astype(str).str.strip().eq("")
            df.loc[blank, weekday_column] = df.loc[blank, old_column]
            df = df.drop(columns=[old_column])
    return df

@st.cache_data
def parse_fixed_csv(file_bytes):
    df = normalize_shift_names(_read_csv_any_encoding(file_bytes))
    if '区分' in df.columns:
        df = df.rename(columns={'区分': '平日/休日'})
    return df

def parse_shift_date(date_value, target_year, target_month):
    if pd.isna(date_value): return None
    if isinstance(date_value, datetime.datetime): return date_value.date()
    if isinstance(date_value, datetime.date): return date_value
    text=str(date_value).strip().translate(str.maketrans('０１２３４５６７８９／－', '0123456789/-'))
    text=re.sub(r'\([月火水木金土日]\)$','',text).strip()
    full=re.fullmatch(r'(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?',text)
    short=re.fullmatch(r'(\d{1,2})[月/.-](\d{1,2})日?',text)
    if full:
        y,m,d=map(int,full.groups())
    elif short:
        m,d=map(int,short.groups());y=target_year
        if m-target_month>=6:y-=1
        elif target_month-m>=6:y+=1
    else:return None
    try:return datetime.date(y,m,d)
    except ValueError:return None


# ==========================================
# カレンダー一括操作用の裏側ロジック
# ==========================================
def set_all_ng(doc_name, y, m, ndays, val, custom_hols=[]):
    for d in range(1, ndays + 1):
        if val == "OK":
            st.session_state[f"ng_{doc_name}_{y}_{m}_{d}"] = "OK"
        else:
            # 「全選択」が押された場合、休日は「全NG」、平日は「宿NG」にする
            date_obj = datetime.date(y, m, d)
            is_hol = date_obj.weekday() >= 5 or jpholiday.is_holiday(date_obj) or (d in custom_hols)
            st.session_state[f"ng_{doc_name}_{y}_{m}_{d}"] = "全NG" if is_hol else "宿NG"

# ==========================================
# 勤務間隔制約：自動割当は確定勤務・月外勤務との間隔も守る
# ==========================================
def add_interval_constraints(
    model,
    shifts,
    doctors,
    daily_active_shifts,
    num_days,
    target_year,
    target_month,
    min_intervals,
    past_worked_dates,
    future_worked_dates,
    absolute_req_days,
    absolute_req_specific,
):
    """確定同士は維持し、自動割当と全勤務の間隔を制限する。"""
    for doc in doctors:
        gap=min_intervals[doc]
        fixed_days=set(absolute_req_days[doc]) | {d for d,s in absolute_req_specific[doc]}
        external=set(past_worked_dates.get(doc,[])) | set(future_worked_dates.get(doc,[]))
        worked={}
        for d in range(1,num_days+1):
            variables=[shifts[d,doc,s] for s in daily_active_shifts.get(d,[]) if (d,doc,s) in shifts]
            if not variables:continue
            worked[d]=model.NewBoolVar(f'working_{doc}_{d}')
            model.AddMaxEquality(worked[d],variables)
            dt=datetime.date(target_year,target_month,d)
            if d not in fixed_days and any(abs((dt-ext).days)<=gap for ext in external):
                model.Add(worked[d]==0)
        for d in worked:
            for other in range(d+1,min(num_days,d+gap)+1):
                if other in worked and not(d in fixed_days and other in fixed_days):
                    model.Add(worked[d]+worked[other]<=1)


# 月間・横一列を共通のHTMLコンポーネントで描画。選択はブラウザ内で処理し、保存時だけ送信する。
@st.cache_resource
def horizontal_ng_component(html_source):
    import tempfile
    from pathlib import Path
    import streamlit.components.v1 as components
    directory = Path(tempfile.mkdtemp(prefix="shift_ng_calendar_"))
    (directory / "index.html").write_text(html_source, encoding="utf-8")
    import hashlib
    asset_id = hashlib.sha256(html_source.encode("utf-8")).hexdigest()[:16]
    return components.declare_component(f"shift_ng_{asset_id}", path=str(directory))

HORIZONTAL_NG_HTML = '<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>\n*{box-sizing:border-box}body{margin:0;font-family:system-ui,sans-serif;color:#243247;background:white;font-size:14px}.strip{display:flex;gap:8px;overflow-x:auto;width:100%;padding:6px 2px 16px;align-items:stretch;scrollbar-width:auto}.cell{flex:0 0 112px;width:112px;min-width:112px;border:2px solid #dfe3ea;border-radius:9px;padding:6px;background:#f8fafc}.head{height:78px;display:flex;flex-direction:column;justify-content:center;align-items:center;border-radius:5px;gap:5px;font-weight:750;white-space:nowrap}.day{font-size:16px}.state{font-size:14px}.day.holiday{color:#c92336}.day.saturday{color:#1670c5}select{width:100%;height:36px;margin-top:6px;font-size:14px;font-weight:650;border:1px solid #8b95a5;border-radius:5px;background:white;color:#243247;padding:2px}button{background:#ff4b4b;color:white;border:0;border-radius:7px;padding:12px 18px;font:600 14px system-ui;cursor:pointer}button:focus-visible,select:focus-visible{outline:3px solid #4789ff;outline-offset:2px}.note{margin:8px 0;font-size:13px;color:#566174;min-height:20px}\n\n.strip.month{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:6px;overflow:visible;padding-bottom:8px}\n.month .cell{width:auto;min-width:0;padding:5px;flex:none}\n.month .head{height:72px}.weekday{text-align:center;font-weight:700;padding:4px}.blank{min-height:126px;background:#f5f6f8;border-radius:9px}\n@media(max-width:600px){.strip.month{gap:3px}.month .cell{padding:2px;border-width:1px}.month .day{font-size:12px}.month .state{font-size:11px}.month select{font-size:11px;padding:0;height:30px}.month .head{height:66px}.month .blank{min-height:108px}.weekday{font-size:12px}}\n\n/* 上段が日直、下段が宿直。不可の勤務帯だけ塗る。 */\n.cell,.month .cell{background:#fff;border-color:#cbd2dc}\n.head,.month .head{height:124px;gap:4px;justify-content:flex-start;padding-top:2px;color:#243247;background:transparent}\n.state{width:100%;display:grid;grid-template-rows:repeat(2,29px);gap:0;border:1px solid #d6dce5;border-radius:5px;overflow:hidden;order:3}\n.band{display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:750;background:#fff;color:#425268;white-space:nowrap}\n.band + .band{border-top:1px solid #d6dce5}.band.off{background:#fcfcfd;color:#d8dde5;font-weight:400}\n.cell[data-state="全NG"] .band{background:#b42332;color:#fff}\n.cell[data-state="日NG"] .day-band{background:#9a4700;color:#fff}\n.cell[data-state="宿NG"] .night-band{background:#1856a4;color:#fff}\n.warning-slot{height:22px;min-height:22px;display:flex;align-items:center;justify-content:center}\n.weekday-warning{color:#bf5700;background:#fff0c2;border:1px solid #ef9b20;border-radius:4px;padding:0 4px;font-size:14px;font-weight:900;line-height:20px}\n.blank{min-height:176px}\n@media(max-width:600px){.month .head{height:120px}.month .band{font-size:10px}.month .weekday-warning{font-size:15px;padding:0 3px}.month .blank{min-height:160px}}\n\n.weekday-warning{display:inline-flex;align-items:center;justify-content:center;gap:3px;max-width:100%}\n.duty-label{font-size:9px;font-weight:650;line-height:1.15;white-space:nowrap}\n@media(max-width:600px){.month .weekday-warning{gap:1px;padding:0 1px;font-size:12px}.month .duty-label{font-size:8px;white-space:normal;max-width:30px;overflow-wrap:anywhere}}\n\n/* 日付の文字・休日色・警告の有無によらず各段の位置を固定する。 */\n.head,.month .head{\n display:grid;\n grid-template-columns:minmax(0,1fr);\n grid-template-rows:26px 30px 60px;\n justify-content:stretch;\n width:100%;\n min-width:0;\n align-content:end;\n align-items:center;\n justify-items:center;\n gap:4px;\n padding-top:0;\n padding-bottom:0;\n}\n.head > .day{line-height:24px;margin:0;align-self:center}\n.head > .warning-slot{height:30px;min-height:30px;width:100%;margin:0}\n.head > .state{height:60px;min-height:60px;margin:0;align-self:end}\n@media(max-width:600px){\n .month .head{grid-template-rows:22px 30px 60px}\n}\n/* 休日はNGの勤務帯をすべて同じ赤にする。平日の宿NGは従来の青。 */\n.cell[data-holiday="true"][data-state="日NG"] .day-band,\n.cell[data-holiday="true"][data-state="宿NG"] .night-band,\n.cell[data-holiday="true"][data-state="全NG"] .band{background:#b42332;color:#fff}\n</style><body><div id="strip" class="strip" aria-label="NG日カレンダー"></div><div id="note" class="note">左右へスクロールして選択できます。</div><button id="apply" type="button">NG日を保存する</button><script>\nlet version=null,days=[],draft=[],seen=null;let mode=\'row\';let frameHeight=0;function resize(){let h=document.body.scrollHeight+10;if(h!==frameHeight){frameHeight=h;send(\'streamlit:setFrameHeight\',{height:h})}}const labels={OK:\'OK\',\'全NG\':\'✕ 全NG\',\'日NG\':\'☀ 日NG\',\'宿NG\':\'☾ 宿NG\'};\nfunction send(type,data={}){window.parent.postMessage({isStreamlitMessage:true,type,...data},\'*\')}\nfunction paint(cell,value){cell.dataset.state=value;let holiday=cell.dataset.holiday===\'true\';cell.querySelector(\'.day-band\').textContent=holiday?(value===\'全NG\'||value===\'日NG\'?\'日直NG\':\'日直可\'):\'日直なし\';cell.querySelector(\'.night-band\').textContent=value===\'全NG\'||value===\'宿NG\'?\'宿直NG\':\'宿直可\';}\nwindow.addEventListener(\'message\',(event)=>{if(event.data.type!==\'streamlit:render\')return;let a=event.data.args;document.getElementById(\'apply\').textContent=a.doctor+\'先生のNG日を保存する\';if(a.version!==version){version=a.version;mode=a.mode||\'row\';days=a.days;draft=days.map(d=>d.value);let strip=document.getElementById(\'strip\');strip.replaceChildren();strip.className=\'strip \'+(mode===\'month\'?\'month\':\'\');if(mode===\'month\'){[\'月\',\'火\',\'水\',\'木\',\'金\',\'土\',\'日\'].forEach((w,i)=>{let el=document.createElement(\'div\');el.className=\'weekday\';el.textContent=w;el.style.color=i===6?\'#c92336\':i===5?\'#1670c5\':\'#243247\';strip.append(el)});for(let i=0;i<a.offset;i++){let el=document.createElement(\'div\');el.className=\'blank\';strip.append(el)}}days.forEach((d,i)=>{let cell=document.createElement(\'div\');cell.className=\'cell\';cell.dataset.holiday=String(d.options.includes(\'日NG\'));let head=document.createElement(\'div\');head.className=\'head\';let day=document.createElement(\'div\');day.className=\'day \'+d.kind;day.textContent=mode===\'month\'?d.day+\'日\':d.day+\'日（\'+d.weekday+\'）\';let state=document.createElement(\'div\');state.className=\'state\';let dayBand=document.createElement(\'div\');dayBand.className=\'band day-band\'+(d.options.includes(\'日NG\')?\'\':\' off\');let nightBand=document.createElement(\'div\');nightBand.className=\'band night-band\';state.append(dayBand,nightBand);let warningSlot=document.createElement(\'div\');warningSlot.className=\'warning-slot\';if(d.warning){let warning=document.createElement(\'span\');warning.className=\'weekday-warning\';warning.textContent=\'⚠︎\';let duty=document.createElement(\'small\');duty.className=\'duty-label\';duty.textContent=\'翌日PM duty\';warning.append(duty);warning.title=\'翌日PM duty（翌日が休日なら例外あり）\';warning.setAttribute(\'aria-label\',warning.title);warningSlot.append(warning)}head.append(day,warningSlot,state);let sel=document.createElement(\'select\');sel.setAttribute(\'aria-label\',d.day+\'日のNG設定\');d.options.forEach(v=>{let o=document.createElement(\'option\');o.value=v;o.textContent=labels[v];sel.append(o)});sel.value=d.value;sel.onchange=()=>{draft[i]=sel.value;paint(cell,sel.value);document.getElementById(\'note\').textContent=\'未保存の変更があります。選び終わったら下のボタンを押してください。\'};cell.append(head,sel);paint(cell,d.value);strip.append(cell)});if(mode===\'month\'){let blanks=(7-(a.offset+days.length)%7)%7;for(let i=0;i<blanks;i++){let el=document.createElement(\'div\');el.className=\'blank\';strip.append(el)}}document.getElementById(\'note\').textContent=\'NGを選ぶと色が変わります。選び終わったら保存してください。\';}resize()});\ndocument.getElementById(\'apply\').onclick=()=>{send(\'streamlit:setComponentValue\',{value:{version,values:draft,token:Date.now().toString()+\'-\'+Math.random()},dataType:\'json\'});document.getElementById(\'note\').textContent=\'保存しています…\';};new ResizeObserver(resize).observe(document.body);send(\'streamlit:componentReady\',{apiVersion:1});resize();\n</script></body></html>'


def build_hover_schedule_html(df, shift_columns, doctors, color_style):
    """医師名はエスケープし、完全一致のIDで同名を強調する。"""
    import html
    ids = {name: str(i) for i, name in enumerate(dict.fromkeys(doctors))}
    parts = ['<!doctype html><html lang="ja"><meta charset="utf-8"><style>',
        """
        *{box-sizing:border-box}body{margin:0;font:14px system-ui,sans-serif;color:#243247;background:white}
        .scroll{overflow-x:auto;border:1px solid #d6dce5;border-radius:6px}
        table{border-collapse:separate;border-spacing:0;width:100%;min-width:720px;table-layout:fixed}
        th,td{overflow-wrap:anywhere}
        th{white-space:nowrap}
        .doctor{max-width:100%;white-space:normal}
        th,td{padding:7px 9px;border-bottom:1px solid #e4e7ed;border-right:1px solid #e4e7ed;text-align:left}
        th{position:sticky;top:0;background:#f3f5f8;z-index:1}
        .doctor{display:inline-block;border-radius:4px;padding:2px 4px;outline-offset:1px;cursor:default}
        .doctor.match{background:#ffeb70!important;color:#111!important;outline:2px solid #bf7100;font-weight:800}
        .doctor:focus-visible{outline:2px solid #bf7100}
        .note{font-size:16px;line-height:1.7;font-weight:500;color:#243247;margin:0 0 12px;overflow-wrap:anywhere}
        """, '</style><body><p class="note">名前にカーソルを合わせると同じ医師を強調します。クリックするとカーソルを外しても強調表示が続き、同じ名前をもう一度クリックすると解除します。スマートフォンではタップで同じ操作ができます。</p><div class="scroll"><table><colgroup><col style="width:12%"><col style="width:10%"><col span="6" style="width:13%"></colgroup><thead><tr>']
    parts.extend('<th>'+html.escape(str(c))+'</th>' for c in df.columns)
    parts.append('</tr></thead><tbody>')
    for _, row in df.iterrows():
        parts.append('<tr>')
        for col in df.columns:
            value=str(row[col])
            style=color_style(value) if col in shift_columns else (
                'color:#ff4b4b;font-weight:bold;' if row.get('平日/休日')=='休日' else '')
            parts.append('<td style="'+html.escape(style,quote=True)+'">')
            if col in shift_columns:
                tokens=[x.strip() for x in re.split(r'[、,]',value)]
                labels=[]
                for name in tokens:
                    escaped=html.escape(name)
                    labels.append('<span class="doctor" tabindex="0" data-doctor="'+ids[name]+'">'+escaped+'</span>' if name in ids else escaped)
                parts.append('、'.join(labels))
            else:
                parts.append(html.escape(value))
            parts.append('</td>')
        parts.append('</tr>')
    parts.append("""</tbody></table></div><script>
    function notify(type,data={}) {
        window.parent.postMessage({isStreamlitMessage:true,type,...data},'*');
    }
    let lastHeight=0;
    function resizeFrame() {
        const height=Math.ceil(document.body.getBoundingClientRect().height)+8;
        if(height!==lastHeight){lastHeight=height;notify('streamlit:setFrameHeight',{height});}
    }
    notify('streamlit:componentReady',{apiVersion:1});
    window.addEventListener('message',e=>{
        if(e.data && e.data.type==='streamlit:render')requestAnimationFrame(resizeFrame);
    });
    new ResizeObserver(resizeFrame).observe(document.body);
    if(document.fonts)document.fonts.ready.then(resizeFrame);
    requestAnimationFrame(resizeFrame);
    const names=[...document.querySelectorAll('.doctor')];
    let pinned=null;
    function highlight(id){names.forEach(el=>el.classList.toggle('match',el.dataset.doctor===id));}
    names.forEach(el=>{
        el.addEventListener('mouseenter',()=>highlight(el.dataset.doctor));
        el.addEventListener('mouseleave',()=>highlight(pinned));
        el.addEventListener('focus',()=>highlight(el.dataset.doctor));
        el.addEventListener('blur',()=>highlight(pinned));
        el.addEventListener('click',()=>{pinned=pinned===el.dataset.doctor?null:el.dataset.doctor;highlight(pinned);});
        el.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();el.click();}});
    });
    document.addEventListener('keydown',e=>{if(e.key==='Escape'){pinned=null;highlight(null);}});
    </script></body></html>""")
    return ''.join(parts)

ALL_SHIFT_TYPES = ['A宿直', 'B宿直', '外来宿直', 'A日直', 'B日直', '外来日直']
NUMERIC_DEFAULTS = {'希望優先度(数字が大きいほど優先)':1, '最低空ける日数':5, '月間最小回数':0, '月間最大回数':5, '休日最大回数':4, **{s+'上限':2 for s in ALL_SHIFT_TYPES}}

def clean_text(value):
    return '' if pd.isna(value) else str(value).strip()

def checked_day_items(value, year, month, ng=False):
    text = clean_text(value).translate(str.maketrans('０１２３４５６７８９，：', '0123456789,:'))
    items = []
    for token in text.split(','):
        if not token.strip():
            continue
        parts = token.strip().split(':')
        if len(parts) > 2 or not re.fullmatch(r'\d+', parts[0].strip()):
            raise ValueError(f'「{token}」：日付は整数、複数は半角カンマで入力してください。')
        d = int(parts[0].strip())
        if not 1 <= d <= calendar.monthrange(year, month)[1]:
            raise ValueError(f'「{token}」：{year}年{month}月にない日付です。')
        kind = parts[1].strip() if len(parts)==2 else ('全NG' if ng else None)
        if kind not in (['全NG','日NG','宿NG','OK'] if ng else ALL_SHIFT_TYPES+[None]):
            raise ValueError(f'「{token}」：枠名・NGの種類を確認してください。')
        items.append((d,kind))
    if ng:
        by_day={}
        for d,kind in items:
            if d in by_day and by_day[d]!=kind:
                raise ValueError(f'{d}日のNG指定が重複して異なっています。1種類にしてください。')
            by_day[d]=kind
    return list(dict.fromkeys(items))

def validate_staff_inputs(df, year, month):
    df = df.copy()
    errors = []
    if '先生の名前' not in df:
        return df, ['医師条件CSVに「先生の名前」の列がありません。']
    df['先生の名前'] = df['先生の名前'].map(clean_text)
    blank = df['先生の名前'].eq('')
    for i, row in df[blank].iterrows():
        if any(clean_text(v) for k,v in row.items() if k!='先生の名前'):
            errors.append(f'医師条件 {i+1}行目：条件が入力されていますが医師名が空欄です。')
    df = df[~blank].reset_index(drop=True)
    for name in df.loc[df['先生の名前'].duplicated(), '先生の名前'].unique():
        errors.append(f'医師名「{name}」が重複しています。')
    for col, default in NUMERIC_DEFAULTS.items():
        if col not in df:
            df[col] = default
    for col in ['NG日(半角カンマ区切り)','希望日(半角カンマ区切り)','翌日PM duty']:
        if col not in df:
            df[col] = ''
    for i,row in df.iterrows():
        name = row['先生の名前']
        if any(c in name for c in ',、') or name=='-' or '⚠️不足' in name:
            errors.append(f'医師名「{name}」：カンマ・読点・不足表示・単独の「-」は使用できません。')
        for col, default in NUMERIC_DEFAULTS.items():
            raw = row[col]
            try:
                n = default if not clean_text(raw) else float(raw)
                if not math.isfinite(n) or n<0 or n!=int(n) or n>1000000:
                    raise ValueError()
                df.at[i,col] = int(n)
            except (TypeError,ValueError,OverflowError):
                errors.append(f'{name}：{col}は0〜1000000の整数にしてください。')
        try:
            if float(df.at[i,'月間最小回数']) > float(df.at[i,'月間最大回数']):
                errors.append(f'{name}：月間最小回数は月間最大回数以下にしてください。')
        except (ValueError,TypeError):
            pass
        for col, ng in [('NG日(半角カンマ区切り)',True),('希望日(半角カンマ区切り)',False)]:
            try:
                items = checked_day_items(row[col],year,month,ng)
                df.at[i,col] = ','.join(str(d) if kind is None or (ng and kind=='全NG') else f'{d}:{kind}' for d,kind in items)
            except ValueError as e:
                errors.append(f'{name}：{col.split("(")[0]} {e}')
        weekdays = clean_text(row['翌日PM duty']).replace('，',',')
        if any(w.strip() not in list('月火水木金土日') for w in weekdays.split(',') if w.strip()):
            errors.append(f'{name}：翌日PM dutyは「水,木」のように曜日を入力してください。')
    return df, errors

def validate_fixed_inputs(df, doctors, year, month):
    errors=[]
    if df is None:
        return None,errors
    df=df.copy()
    if '日付' not in df:
        return df,['確定当直CSVに「日付」の列がありません。']
    unknown = [c for c in df if c not in ['日付','平日/休日']+ALL_SHIFT_TYPES]
    for c in unknown:
        if df[c].map(clean_text).ne('').any():
            errors.append(f'確定当直：未対応の列「{c}」に入力があります。枠名を確認してください。')
    for i,row in df.iterrows():
        raw=clean_text(row['日付'])
        populated=any(clean_text(row.get(s,'')) not in ('','-') for s in ALL_SHIFT_TYPES)
        if not raw and not populated:
            continue
        dt=parse_shift_date(raw,year,month)
        if dt is None:
            errors.append(f'確定当直 {i+1}行目：日付「{raw}」を確認してください。年付きの年月日で指定できます。')
            continue
        df.at[i,'日付']=dt.isoformat()
        for s in ALL_SHIFT_TYPES:
            val=clean_text(row.get(s,''))
            if val in ('','-'):
                continue
            names=list(dict.fromkeys(n.strip() for n in re.split('[、,]',val)))
            for name in names:
                if name not in doctors:
                    errors.append(f'確定当直 {dt} {s}：医師名「{name}」が名簿と一致しません。')
            df.at[i,s]='、'.join(names)
    return df,errors

def input_signature(year, month, staff, holidays, multi, fixed):
    import hashlib
    import json
    payload=[year,month,staff.to_csv(index=False),sorted(holidays), sorted((d,s,c) for (d,s),c in multi.items()),fixed.to_csv(index=False)]
    return hashlib.sha256(json.dumps(payload,ensure_ascii=False).encode()).hexdigest()


def audit_schedule(result, staff, year, month, holidays, multi, past, future):
    warnings=[]
    for _,row in staff.iterrows():
        name=row['先生の名前']
        worked=[]; counts=dict.fromkeys(ALL_SHIFT_TYPES,0); hol_count=0
        for i,r in result.iterrows():
            d=parse_shift_date(r['日付'],year,month)
            for s in ALL_SHIFT_TYPES:
                if name in [n.strip() for n in re.split('[、,]',str(r[s]))]:
                    worked.append(d);counts[s]+=1
                    hol_count+=int(r['平日/休日']=='休日')
        ng = dict(checked_day_items(row.get('NG日(半角カンマ区切り)', ''), year, month, True))
        for _, r in result.iterrows():
            dt = parse_shift_date(r['日付'], year, month)
            next_dt = dt + datetime.timedelta(days=1)
            next_holiday = next_dt.weekday() >= 5 or jpholiday.is_holiday(next_dt) or ((next_dt.year,next_dt.month)==(year,month) and next_dt.day in holidays)
            for slot in ALL_SHIFT_TYPES:
                if name not in [n.strip() for n in re.split('[、,]', str(r[slot]))]: continue
                kind = ng.get(dt.day, 'OK')
                night = slot in ['A宿直','B宿直','外来宿直']
                if kind=='全NG' or (kind=='宿NG' and night) or (kind=='日NG' and not night):
                    warnings.append(f'{name}：{dt} {slot}はNG指定より確定指定を優先しました。')
                if night and '月火水木金土日'[dt.weekday()] in clean_text(row.get('翌日PM duty','')) and not next_holiday:
                    warnings.append(f'{name}：{dt} {slot}は翌日PM dutyの制限より確定指定を優先しました。')
        total=sum(counts.values())
        if total < int(row['月間最小回数']):
            warnings.append(f'{name}：月間最小回数の目標{int(row["月間最小回数"])}回に対し、{total}回です。')
        for label,value,cap in [('月間最大回数',total,int(row['月間最大回数'])),('休日最大回数',hol_count,int(row['休日最大回数']))]+[(s+'上限',counts[s],int(row[s+'上限'])) for s in ALL_SHIFT_TYPES]:
            if value>cap:
                warnings.append(f'{name}：確定指定に伴い{label}を超えています（設定{cap}回／結果{value}回）。')
        dates=sorted(set(worked)|set((past or {}).get(name,[]))|set((future or {}).get(name,[])))
        for a,b in zip(dates,dates[1:]):
            if (a.year,a.month)!=(year,month) and (b.year,b.month)!=(year,month):continue
            gap=(b-a).days-1
            if gap<int(row['最低空ける日数']):
                warnings.append(f'{name}：確定勤務同士の間隔が設定未満です（{a}→{b}、空き{gap}日／設定{int(row["最低空ける日数"])}日）。確定勤務を維持しました。')
    for _,r in result.iterrows():
        d=parse_shift_date(r['日付'],year,month).day
        for s in ALL_SHIFT_TYPES:
            names=[n.strip() for n in re.split('[、,]',str(r[s])) if n.strip() not in ('','-') and '⚠️不足' not in n]
            need=multi.get((d,s),1)
            if len(names)>need:
                warnings.append(f'{month}/{d} {s}：設定{need}名に対して{len(names)}名を配置しています。')
    return warnings


def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None):
    staff_df,errors=validate_staff_inputs(staff_df,target_year,target_month)
    fixed_df,fixed_errors=validate_fixed_inputs(fixed_df,staff_df.get('先生の名前',pd.Series(dtype=str)).tolist(),target_year,target_month)
    errors+=fixed_errors
    if staff_df.empty:errors.append('医師を1名以上入力してください。')
    for (d,s),count in multi_slots_dict.items():
        if not isinstance(d,int) or not 1<=d<=calendar.monthrange(target_year,target_month)[1] or s not in ALL_SHIFT_TYPES or not isinstance(count,int) or not 2<=count<=10:
            errors.append(f'増員設定「{d}日 {s} {count}名」を確認してください。')
        elif s in ['A日直','B日直','外来日直']:
            dt=datetime.date(target_year,target_month,d)
            if not(dt.weekday()>=5 or jpholiday.is_holiday(dt) or d in custom_holidays):
                errors.append(f'{d}日の日直増員：先に特別休日を指定してください。')
    if errors:return None,False,errors,None,None
    result,success,warnings,past,future=_generate_shift_core(target_year,target_month,staff_df,custom_holidays,multi_slots_dict,fixed_df)
    if result is not None:
        warnings+=audit_schedule(result,staff_df,target_year,target_month,custom_holidays,multi_slots_dict,past,future)
        success=not any('⚠️不足' in str(v) for s in ALL_SHIFT_TYPES for v in result[s])
    return result,success,list(dict.fromkeys(warnings)),past,future



def show_input_errors(messages):
    messages = list(dict.fromkeys(str(m).strip() for m in messages if str(m).strip()))
    st.error(f"入力内容を確認してください（{len(messages)}件）")
    st.dataframe(pd.DataFrame({"修正が必要な項目": messages}), hide_index=True, use_container_width=True)


def show_schedule_notice(result, messages):
    messages = list(dict.fromkeys(str(m).strip() for m in messages if str(m).strip()))
    shortage_rows = []
    if result is not None:
        for _, row in result.iterrows():
            for slot in ALL_SHIFT_TYPES:
                match = re.search(r'⚠️不足\((\d+)名\)', str(row.get(slot, '')))
                if match:
                    shortage_rows.append(int(match.group(1)))
    # 最重要の結果を、説明より先に表示する。
    if result is None:
        st.error("当直案を作成できませんでした。")
    elif shortage_rows:
        st.warning(f"不足が残っています：{len(shortage_rows)}枠、合計{sum(shortage_rows)}名分")
    else:
        st.success("必要人数を満たす当直案ができました。")

    details = []
    for message in messages:
        if message.startswith(('🚨 **以下の枠', '📊 **【不足している枠')):
            continue
        if re.fullmatch(r'・\d+/\d+ の「.+」', message) or re.fullmatch(r'・.+： 計 \d+ 枠不足', message):
            continue
        details.append(message)
    if details or result is None or shortage_rows:
        with st.container(border=True):
            st.markdown("**計算状況・条件の注意事項と対処方法**")
            for message in details:
                st.markdown(message)
            if result is None or shortage_rows:
                st.markdown("**条件を見直す場合**")
                st.write("入力内容を確認し、実際に調整できる範囲で、月間・休日・枠別の上限や勤務間隔を見直して再作成してください。")
            if shortage_rows:
                st.write("不足箇所は当直案の「⚠️不足」で確認できます。手動調整する場合はCSVをダウンロードし、Excelなどで編集してください。")


# ページ設定
st.set_page_config(page_title="当直作成アプリ", layout="wide")
st.title("当直作成アプリ")
st.caption("年月・勤務条件・NG日を入力して、当直表の案を作成します。")
with st.expander("初めて使う方へ：入力からダウンロードまで", expanded=False):
    st.markdown("""
1. **年月・休日・必要人数を設定**します。
2. 必要に応じて、**先月・今月・来月の確定当直**を入力します。
3. **医師ごとの回数・勤務間隔・希望日**を入力します。
4. 各医師のカレンダーで**NG日を選び、最後に「NG日を保存する」を押します。**
5. **「当直案を作成する」**を押し、結果を確認してCSVをダウンロードします。

途中で終了する場合は、下の「医師条件をCSVで保存」をご利用ください。
特別休日・増員設定・確定済み当直は、そのCSVには含まれません。
    """)


def calendar_cell_container():
    """
    上部の「特別休日の設定」カレンダー用。
    対応する日付とチェックボックスの関係が分かりやすいよう、
    1日ごとに枠線付きコンテナを使う。
    border=True 未対応環境では通常コンテナにフォールバックする。
    """
    try:
        return st.container(border=True)
    except TypeError:
        return st.container()


# === スマホ＆フォーム内で絶対に崩れないカレンダー用CSS ===
st.markdown("""
<style>

/* =========================================================
   フォーム内の7列カレンダー
   ========================================================= */

/* 7列のブロックを画面幅の中に必ず収める */
div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7)) {
    display: grid !important;
    grid-template-columns: repeat(7, minmax(0, 1fr)) !important;
    gap: 4px !important;
    width: 100% !important;
    max-width: 100% !important;
    margin-left: 0 !important;
    margin-right: 0 !important;
    padding-left: 0 !important;
    padding-right: 0 !important;
    box-sizing: border-box !important;
    overflow: hidden !important;
}

/* =========================================================
   カレンダー各セル
   ========================================================= */

div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
> div[data-testid="column"] {
    flex: none !important;
    width: auto !important;
    min-width: 0 !important;
    max-width: 100% !important;
    margin: 0 !important;
    box-sizing: border-box !important;
    border: 1px solid #eee;
    border-radius: 4px;
    padding: 6px 2px !important;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: flex-start !important;
    background-color: #ffffff;
    overflow: hidden !important;
}

/* =========================================================
   Streamlit特有の余白を削除
   ========================================================= */

div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
.element-container {
    margin: 0 !important;
    padding: 0 !important;
    width: 100% !important;
    max-width: 100% !important;
    display: flex;
    justify-content: center;
    box-sizing: border-box !important;
}

/* =========================================================
   日付・曜日などの文字
   ========================================================= */

div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7)) p,
div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7)) label,
div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
div[data-testid="stMarkdownContainer"],
div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7)) span,
div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7)) b {
    font-size: 0.8rem !important;
    text-align: center;
    margin: 0 !important;
    white-space: nowrap !important;
    word-break: keep-all !important;
    line-height: 1.5 !important;
}

/* =========================================================
   Selectbox
   ========================================================= */

div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
div[data-testid="stSelectbox"] {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
}

div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
div[data-baseweb="select"] {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    font-size: 0.75rem !important;
    box-sizing: border-box !important;
}

div[data-testid="stForm"]:not(:has([data-ng-layout="row"]))
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
div[data-baseweb="select"] > div {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    padding-left: 2px !important;
    padding-right: 2px !important;
    min-height: 1.8rem !important;
    box-sizing: border-box !important;
}


/* =========================================================
   上部「特別休日の設定」カレンダー
   格子セルの高さを平日・休日・空欄ですべて統一
   ========================================================= */
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
div[data-testid="stVerticalBlockBorderWrapper"] {
    height: 112px !important;
    min-height: 112px !important;
    max-height: 112px !important;
    box-sizing: border-box !important;
    overflow: hidden !important;
}

div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
div[data-testid="stVerticalBlockBorderWrapper"] > div {
    height: 100% !important;
    box-sizing: border-box !important;
}

/* チェックボックスのラベルをチェックのすぐ横に表示 */
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="column"]:nth-child(7))
div[data-testid="stVerticalBlockBorderWrapper"]
div[data-testid="stCheckbox"] label {
    justify-content: center !important;
    gap: 6px !important;
    width: 100% !important;
}


/* 医師別NGカレンダー：選択内容をセル全体で強調 */
div[data-testid="stForm"]:has([data-ng-state]) div[data-testid="stHorizontalBlock"] {
    display: grid !important;
    grid-template-columns: repeat(7, minmax(0, 1fr)) !important;
    gap: 4px !important;
}
div[data-testid="stForm"]:has([data-ng-state]) div[data-testid="stHorizontalBlock"] > div {
    width: auto !important; min-width: 0 !important;
    padding: 4px !important; border: 2px solid #E5E7EB !important;
    border-radius: 8px; background: #FFFFFF; box-sizing: border-box;
}
div[data-testid="stForm"]:has([data-ng-state]) div[data-testid="stHorizontalBlock"] > div:has([data-ng-state="全NG"]) {
    background: #FDE8EC; border-color: #B42332 !important;
    box-shadow: inset 0 0 0 1px #B42332;
}
div[data-testid="stForm"]:has([data-ng-state]) div[data-testid="stHorizontalBlock"] > div:has([data-ng-state="日NG"]) {
    background: #FFF0D9; border-color: #9A4700 !important;
    box-shadow: inset 0 0 0 1px #9A4700;
}
div[data-testid="stForm"]:has([data-ng-state]) div[data-testid="stHorizontalBlock"] > div:has([data-ng-state="宿NG"]) {
    background: #E3EFFF; border-color: #1856A4 !important;
    box-shadow: inset 0 0 0 1px #1856A4;
}
div[data-testid="stForm"]:has([data-ng-state]) [data-baseweb="select"] > div {
    min-height: 1.9rem !important; padding: 0 2px !important;
    font-size: 0.75rem !important; font-weight: 700 !important;
}
div[data-testid="stForm"]:has([data-ng-state]) [data-baseweb="select"] { min-width: 0 !important; }
div[data-testid="stForm"]:has([data-ng-state]) [data-testid="stVerticalBlock"] { gap: 5px !important; }

</style>
""", unsafe_allow_html=True)
# ==============================================================================

# ==========================================
# 1. 上部ダッシュボード：年月と休日の設定
# ==========================================
st.markdown('\n<style>\n/* 本文の見出しだけに適用。カレンダーや結果表の部品には影響しない。 */\n[data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h1,\n.main [data-testid="stHeadingWithActionElements"] h1 {\n    font-size: 1.9rem !important;\n    line-height: 1.4 !important;\n    font-weight: 750 !important;\n    padding: 0 0 .6rem !important;\n}\n[data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h2,\n.main [data-testid="stHeadingWithActionElements"] h2 {\n    font-size: 1.3rem !important;\n    line-height: 1.5 !important;\n    font-weight: 700 !important;\n    background: var(--secondary-background-color, #f0f3f7);\n    border-left: 5px solid #47749b;\n    border-radius: 0 8px 8px 0;\n    padding: .85rem 1rem !important;\n    margin: .4rem 0 .35rem !important;\n}\n[data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h3,\n.main [data-testid="stHeadingWithActionElements"] h3 {\n    font-size: 1.08rem !important;\n    line-height: 1.55 !important;\n    font-weight: 700 !important;\n    padding: .45rem 0 .6rem !important;\n    border-bottom: 1px solid rgba(128, 144, 160, .3);\n    margin: .2rem 0 .15rem !important;\n}\n[data-testid="stMain"] [data-testid="stMarkdownContainer"] hr,\n.main [data-testid="stMarkdownContainer"] hr {\n    margin: 1.4rem 0 !important;\n    border: 0;\n    border-top: 1px solid rgba(128, 144, 160, .35);\n}\n[data-testid="stMetricLabel"] p {font-size: .9rem !important; line-height: 1.5 !important;}\n[data-testid="stMetricValue"] {font-size: 1.8rem !important; line-height: 1.25 !important;}\n@media (max-width:600px) {\n    [data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h1 {font-size:1.6rem !important;}\n    [data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h2 {font-size:1.15rem !important; padding:.7rem .75rem !important;}\n    [data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h3 {font-size:1rem !important;}\n}\n</style>\n', unsafe_allow_html=True)

st.header("📅 年月・休日・必要人数の設定")
st.info("作成する年月を選んでください。土日祝日は自動で休日扱いになります。平日にも日直を設けたい場合は、その日の「休日にする」にチェックを入れます。")

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

st.subheader(f"📅 平日に日直を設ける日（特別休日） - {month}月")
st.caption("年末年始やお盆など、平日でも日直が必要な日を指定します。チェックした日は、日直・宿直ともに休日回数の集計対象になります。")

cal_matrix = calendar.monthcalendar(year, month)
weekdays_ja = ["月", "火", "水", "木", "金", "土", "日"]
custom_holidays = []

st.markdown('\n<style>\n.st-key-special_holiday_calendar [data-testid="stHorizontalBlock"] {\n display:grid !important;\n grid-template-columns:repeat(7,minmax(0,1fr)) !important;\n gap:6px !important;\n width:100% !important;\n}\n.st-key-special_holiday_calendar [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"],[data-testid="column"]) {\n width:100% !important; min-width:0 !important;\n flex:none !important;\n}\n.st-key-special_holiday_calendar [class*="st-key-special_day_"] {\n height:112px !important; min-height:112px !important;\n border:1px solid #d6dce5; border-radius:7px;\n padding:8px 3px !important; gap:6px !important;\n box-sizing:border-box;\n}\n.st-key-special_holiday_calendar [data-testid="stElementContainer"]:has([data-testid="stCheckbox"]),\n.st-key-special_holiday_calendar .element-container:has([data-testid="stCheckbox"]) {\n width:100% !important; align-self:stretch !important;\n}\n.st-key-special_holiday_calendar [data-testid="stCheckbox"] {\n width:100% !important; display:flex !important; justify-content:center !important;\n}\n.st-key-special_holiday_calendar [data-testid="stCheckbox"] label {\n display:flex !important; justify-content:center !important; gap:3px; width:fit-content !important; max-width:100%; margin-left:auto !important; margin-right:auto !important;\n min-width:0;\n}\n.st-key-special_holiday_calendar [data-testid="stCheckbox"] label p {\n font-size:12px; line-height:1.2; overflow-wrap:anywhere;\n}\n.st-key-special_holiday_calendar [data-testid="stCheckbox"] label > span {\n flex-shrink:0;\n}\n@media(max-width:600px) {\n .st-key-special_holiday_calendar [data-testid="stHorizontalBlock"] {gap:3px !important}\n .st-key-special_holiday_calendar [class*="st-key-special_day_"] {\n  height:100px !important; min-height:100px !important; padding:6px 1px !important;\n }\n .st-key-special_holiday_calendar [data-testid="stCheckbox"] label p {font-size:10px}\n .st-key-special_holiday_calendar [data-testid="stCheckbox"] label {gap:1px}\n}\n</style>\n', unsafe_allow_html=True)
with st.container(key="special_holiday_calendar"):
    # 曜日のヘッダー行
    cols = st.columns(7)
    for i, w in enumerate(weekdays_ja):
        color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
        cols[i].markdown(
            f"<div style='color: {color}; font-weight: bold; text-align: center; padding: 4px 0;'>{w}</div>",
            unsafe_allow_html=True
        )

    # 日付とチェックボックス
    # 1日ごとの格子セルは平日・休日・空欄すべて同じ高さにそろえる
    for week in cal_matrix:
        cols = st.columns(7)

        for i, day in enumerate(week):
            with cols[i]:
                with st.container(key=f"special_day_{year}_{month}_{week[0]}_{i}"):

                    if day == 0:
                        # 空欄セルも他の日と同じ高さを保つ
                        st.markdown(
                            "<div style='height: 72px;'></div>",
                            unsafe_allow_html=True
                        )
                        continue

                    date_obj = datetime.date(year, month, day)
                    is_weekend_or_hol = (
                        date_obj.weekday() >= 5
                        or jpholiday.is_holiday(date_obj)
                    )
                    day_color = "#ff4b4b" if is_weekend_or_hol else "inherit"

                    # 日付
                    st.markdown(
                        f"""
                        <div style='
                            width: 100%;
                            text-align: center;
                            color: {day_color};
                            font-weight: 600;
                            font-size: 0.95rem;
                            line-height: 1.35;
                            margin: 0 0 8px 0;
                            padding: 0;
                        '>
                            {day}日
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    if is_weekend_or_hol:
                        # 休日はチェック欄と同じ高さの領域に「休」を表示
                        st.markdown(
                            """
                            <div style='
                                width: 100%;
                                height: 2.35rem;
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                color: #ff4b4b;
                                font-size: 0.85rem;
                                line-height: 1.2;
                                margin: 0;
                                padding: 0;
                            '>
                                休
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                    else:
                        # 「休日にする」をチェックボックスのすぐ横に表示
                        if st.checkbox(
                            "休日にする",
                            key=f"hol_{year}_{month}_{day}"
                        ):
                            custom_holidays.append(day)

holiday_total_placeholder = st.empty()

st.divider()

st.subheader("👥 1つの枠を2名以上にする設定（任意）")
st.info("通常は各枠1名です。増員する場合だけ行を追加し、日付・枠・合計人数を選んでください。2名体制にしたい場合は「2」を入力します。")
st.caption("日直の増員は休日扱いの日に設定してください。平日に日直を設ける場合は、先に上のカレンダーで「休日にする」にチェックを入れます。")

_, num_days = calendar.monthrange(year, month)
NIGHT_SHIFTS_UI = ['A宿直', 'B宿直', '外来宿直']
DAY_SHIFTS_UI = ['A日直', 'B日直', '外来日直']

date_options = [f"{d}日" for d in range(1, num_days + 1)]
shift_options = NIGHT_SHIFTS_UI + DAY_SHIFTS_UI

multi_df_template = pd.DataFrame(columns=["日付", "当直枠", "人数"])
edited_multi_df = st.data_editor(
    multi_df_template,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    height=150,
    column_config={
        "日付": st.column_config.SelectboxColumn("日付を選択", options=date_options, required=True),
        "当直枠": st.column_config.SelectboxColumn("増員する枠を選択", options=shift_options, required=True),
        "人数": st.column_config.NumberColumn("合計人数", help="追加人数ではなく、その枠に配置する合計人数です。例：1名から2名体制にする場合は2。", min_value=2, max_value=10, step=1, required=True)
    }
)

multi_slots_dict = {}
for _, row in edited_multi_df.iterrows():
    d_str = str(row.get("日付", ""))
    s_val = str(row.get("当直枠", ""))
    c_val = row.get("人数")
    
    if d_str and s_val and pd.notna(c_val):
        try:
            d_val = int(re.sub(r'\D', '', d_str))
            c_val = int(c_val)
            multi_slots_dict[(d_val, s_val)] = c_val
        except (ValueError, TypeError):
            pass

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
holiday_total_placeholder.metric("🏥 必要な総当直枠数", f"{total_slots} 枠")

st.divider()

# ==========================================
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================

st.header("1. 先月今月来月の確定当直を入力（任意）")
st.info("先月・今月・来月の、担当者が決まっている当直を入力してください。先月末の勤務は、月初の勤務間隔の確認に使います。入力がなければ、この項目は飛ばせます。")
with st.expander("確定済み当直の入力例と扱い", expanded=False):
    st.markdown("""
- CSVのアップロードと、下の表への直接入力のどちらでも入力できます。
- 日付は `2026/10/1` または `10/1` の形式で入力します。
- 担当する枠の欄に、医師条件と同じ名前を入力します。複数人の場合は `佐藤、鈴木` のように「、」で区切ります。
- 医師名の区切りには「、」を使ってください。名前に含まれる空白は区切りません。医師条件と同じ表記にしてください。
- 未確定の枠は空欄で構いません。「平日/休日」欄ではなく、上部のカレンダー設定で休日を判定します。
- 今月の確定勤務はNG日・曜日制限より優先され、回数上限も必要に応じて緩められます。確定勤務を基準に追加勤務の間隔を守ります。確定勤務同士が近すぎる場合は、確定内容を維持して警告します。

**今月の同じ日・同じ枠を指定した場合、確定当直表への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
    """)

fixed_columns = ["日付", "平日/休日", "A宿直", "B宿直", "外来宿直", "A日直", "B日直", "外来日直"]
fixed_template_df = pd.DataFrame(columns=fixed_columns)
fixed_csv_template = fixed_template_df.to_csv(index=False).encode('utf-8-sig')

col_dl_fixed, col_ul_fixed = st.columns(2)
with col_dl_fixed:
    st.write("▼ Excelで一括入力したい場合")
    st.download_button(
        label="📥 ひな形（CSV）をダウンロード",
        data=fixed_csv_template,
        file_name="確定当直_ひな形.csv",
        mime="text/csv",
    )
with col_ul_fixed:
    fixed_file = st.file_uploader("決定済み当直表（CSV）をアップロード", type="csv", key="fixed_csv")

if fixed_file is not None:
    try:
        base_fixed_df = parse_fixed_csv(fixed_file.getvalue())
    except Exception as e:
        st.error(f"確定当直ファイルの読み込みに失敗しました。詳細: {e}")
        st.stop()
else:
    base_fixed_df = pd.DataFrame(columns=fixed_columns)
    base_fixed_df.loc[0] = ["" for _ in range(len(fixed_columns))]

if "日付" not in base_fixed_df.columns:
    st.error("確定当直CSVに「日付」の列がありません。ひな形の列名を確認してください。")
    st.stop()
if "日付" in base_fixed_df.columns:
    base_fixed_df = base_fixed_df.set_index("日付")

st.markdown("### 📅 先月今月来月の確定当直")
st.write("表のセルをクリックして、日付と担当医師名を入力・編集できます。")
edited_fixed_df_raw = st.data_editor(base_fixed_df, num_rows="dynamic", use_container_width=True, height=200)

edited_fixed_df = edited_fixed_df_raw.reset_index()

st.divider()

st.header("2. 医師条件の読み込み・入力（必須）")
st.info("下の表に医師ごとの条件を入力してください。CSVを使う場合は、ひな形をダウンロードして編集し、アップロードします。NG日は、この後の医師別カレンダーで設定します。")
st.caption("最初に表示される5名は入力例です。実際の医師名・条件に置き換えてください。希望優先度は通常「1」を使用します。")
with st.expander("入力例：曜日・希望日・備考", expanded=False):
    st.markdown("""
| 項目 | 入力方法・意味 |
| --- | --- |
| 翌日PM duty | `水,木` のように半角カンマで区切ります。翌日PMにdutyがある曜日を入力します（例：木曜PMにdutyがある場合は「水」）。原則としてその曜日の宿直を外しますが、翌日が休日なら宿直に入る場合があります。日直は対象外です。 |
| 希望日 | `10,15` はその日のいずれかの枠、`10:A宿直` はその枠を希望します。複数の希望は半角カンマで区切ります。 |
| 指定できる枠 | A宿直・B宿直・外来宿直・A日直・B日直・外来日直。表記を一致させてください。 |
| 備考 | 管理用のメモです。「学会」などと書いても計算条件には反映されません。休みはNG日で指定してください。 |

曜日にかかわらず勤務できない日は、カレンダーでNGを指定します。
平日は「宿NG」、休日に日直・宿直とも勤務できない場合は「全NG」を選んでください。
    """)
with st.expander("回数・勤務間隔の数え方", expanded=False):
    st.markdown("""
| 項目 | 意味・入力例 |
| --- | --- |
| 最低空ける日数 | 勤務と次の勤務の間に空ける日数です。5日なら、10日の次は16日以降です。 |
| 月間最小回数 | できるだけ確保したい回数です。条件によっては、この回数に届かないことがあります。 |
| 月間最大回数 | 日直と宿直を合わせた月間の上限です。 |
| 休日最大回数 | 土日祝日・特別休日に担当する日直と宿直の合計上限です。 |
| 各枠の上限 | A宿直など、それぞれの枠を担当する月間の上限です。 |

1つの枠を1回と数えます。確定指定により同じ日に日直と宿直を担当する場合は2回です。
月間最小回数は、月間最大回数以下に設定してください。
確定済み当直や優先度100以上の希望がある場合は、上限の例外があります。自動追加勤務の間隔は守り、確定勤務同士の間隔違反は警告します。
    """)
with st.expander("希望優先度：通常の希望と、100以上の特別な設定", expanded=False):
    st.markdown("""
- **通常は「1」**を使用します。1〜99は、数字が大きいほど希望を優先しますが、NG日・回数上限・勤務間隔などの範囲内で割り当てます。
- **100以上は、その医師の希望日すべてを確定扱いにする設定**です。「できれば入りたい」という用途には使わないでください。
- 確定扱いの日はNG日・曜日制限より優先され、回数上限が必要に応じて緩められます。その日と自動追加勤務の間隔は守ります。確定勤務同士の間隔違反は警告します。
- 一部の勤務だけを確定させたい場合は、上の「確定済み当直」へ入力し、希望優先度は通常の値にしてください。
- 指定の誤りや条件の組み合わせによっては作成できない場合があります。作成後に確定勤務が反映されているか確認してください。

**今月の同じ日・同じ枠を指定した場合、確定当直表への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
    """)

template_data = {
    "先生の名前": ["Dr. A", "Dr. B", "Dr. C", "Dr. D", "Dr. E"],
    "翌日PM duty": ["水,木", "", "土,日", "", ""],
    "NG日(半角カンマ区切り)": ["", "15:日NG", "10:宿NG", "", ""],
    "希望日(半角カンマ区切り)": ["10:A宿直, 15:B日直", "", "8", "20", ""], 
    "希望優先度(数字が大きいほど優先)": [100, 1, 1, 1, 1], 
    "最低空ける日数": [5, 4, 6, 5, 3],  
    "月間最小回数": [1, 2, 0, 1, 0],
    "月間最大回数": [5, 6, 4, 5, 7],
    "休日最大回数": [2, 2, 2, 2, 2],    
    "A宿直上限": [2, 2, 2, 2, 2],
    "B宿直上限": [2, 2, 2, 2, 2],
    "外来宿直上限": [2, 2, 2, 2, 2],
    "A日直上限": [2, 2, 2, 2, 2],
    "B日直上限": [2, 2, 2, 2, 2],
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
        file_name="医師条件_ひな形.csv",
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
        
    try:
        base_df = parse_staff_csv(uploaded_file.getvalue())
    except Exception as e:
        st.error(f"医師条件CSVを読み込めませんでした。詳細: {e}")
        st.stop()
else:
    base_df = df_template.copy()

if "先生の名前" not in base_df.columns:
    st.error("医師条件CSVに「先生の名前」の列がありません。ひな形の列名を確認してください。")
    st.stop()
if "先生の名前" in base_df.columns:
    base_df = base_df.set_index("先生の名前")

if "希望優先度(数字が大きいほど優先)" in base_df.columns:
    base_df["希望優先度(数字が大きいほど優先)"] = pd.to_numeric(base_df["希望優先度(数字が大きいほど優先)"], errors='coerce')

text_cols = ["翌日PM duty", "NG日(半角カンマ区切り)", "希望日(半角カンマ区切り)", "備考（メモ・説明など自由記入）"]
for c in text_cols:
    if c in base_df.columns:
        base_df[c] = base_df[c].apply(lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "none", "<na>"] else str(x))

st.markdown("### 👩‍⚕️ 医師条件の入力・編集")
st.write("セルをクリックして編集できます。列名にマウスを合わせると、入力例や説明が表示されます。医師名は重複しない表記にしてください。")

edited_df = st.data_editor(
    base_df, 
    num_rows="dynamic", 
    use_container_width=True, 
    height=300,
    column_config={
        "翌日PM duty": st.column_config.TextColumn(
            "翌日PM duty",
            help="例：木曜PMにdutyがある場合は水。複数は水,木のように入力。翌日が休日なら宿直に入る場合があります。日直は対象外です。確実に外す日はカレンダーでNGを指定してください。"
        ),
        "最低空ける日数": st.column_config.NumberColumn("最低空ける日数", help="勤務間の空き日数。5日なら10日の次は16日以降。確定勤務と追加勤務の間隔も守ります。確定勤務同士の違反は警告します。"),
        "月間最小回数": st.column_config.NumberColumn("月間最小回数（目標）", help="できるだけ確保したい回数です。条件によっては未達になります。月間最大回数以下にしてください。"),
        "月間最大回数": st.column_config.NumberColumn("月間最大回数", help="日直・宿直を合わせた上限です。確定指定がある場合は例外があります。"),
        "休日最大回数": st.column_config.NumberColumn("休日最大回数", help="土日祝・特別休日の日直と宿直の合計上限です。1枠を1回と数えます。"),
        "A宿直上限": st.column_config.NumberColumn("A宿直上限", help="A宿直を担当する月間の上限です。確定指定がある場合は例外があります。"),
        "B宿直上限": st.column_config.NumberColumn("B宿直上限", help="B宿直を担当する月間の上限です。確定指定がある場合は例外があります。"),
        "外来宿直上限": st.column_config.NumberColumn("外来宿直上限", help="外来宿直を担当する月間の上限です。確定指定がある場合は例外があります。"),
        "A日直上限": st.column_config.NumberColumn("A日直上限", help="A日直を担当する月間の上限です。確定指定がある場合は例外があります。"),
        "B日直上限": st.column_config.NumberColumn("B日直上限", help="B日直を担当する月間の上限です。確定指定がある場合は例外があります。"),
        "外来日直上限": st.column_config.NumberColumn("外来日直上限", help="外来日直を担当する月間の上限です。確定指定がある場合は例外があります。"),
        "NG日(半角カンマ区切り)": None, 
        "希望日(半角カンマ区切り)": st.column_config.TextColumn(
            "希望日",
            help="例：10,15 または 10:A宿直。複数は半角カンマ区切り。通常の希望は各種条件の範囲内で割り当てます。"
        ),
        "希望優先度(数字が大きいほど優先)": st.column_config.NumberColumn(
            "希望優先度",
            help="通常は1。1〜99は数字が大きいほど優先。100以上はこの医師の希望日すべてが確定扱いになり、NG・上限の例外になります。自動追加勤務との間隔は守ります。"
        ),
        "備考（メモ・説明など自由記入）": st.column_config.TextColumn(
            "備考",
            help="管理用メモです。ここに書いた内容は計算に反映されません。"
        )
    }
)

staff_df = edited_df.reset_index()
staff_df, staff_input_errors = validate_staff_inputs(staff_df, year, month)
if staff_input_errors:
    show_input_errors(staff_input_errors)
    st.stop()

st.markdown("### ⚖️ 必要枠数と担当可能回数の目安")
st.caption("月間最大回数の合計と必要枠数を比較しています。プラスでも、NG日・勤務間隔・枠別上限などによっては埋まらない場合があります。確定指定で追加される枠や上限の例外は、この目安に含まれません。")

if "月間最大回数" in staff_df.columns:
    total_max_capacity = pd.to_numeric(staff_df["月間最大回数"], errors='coerce').fillna(0).sum()
    total_max_capacity = int(total_max_capacity)
    
    margin = total_max_capacity - total_slots
    
    c1, c2, c3 = st.columns(3)
    c1.metric("🏥 必要な総当直枠数", f"{total_slots} 枠")
    c2.metric("👩‍⚕️ 医師の月間最大回数の合計", f"{total_max_capacity} 回分")
    
    if margin >= 0:
        c3.metric("✨ 担当可能回数 − 必要枠数", f"+{margin} 回分")
    else:
        c3.metric("🚨 担当可能回数 − 必要枠数", f"{margin} 回分", delta_color="inverse")
st.divider()

st.markdown("### 🚫 先生ごとのNG日設定（カレンダーで詳細選択）")
st.info("医師名のタブを選び、当直NGを選択してください。最後に「NG日を保存する」を押してください。")
with st.expander("NGの種類・一括操作について", expanded=False):
    st.markdown("""
| 選択肢 | 意味 |
| --- | --- |
| OK | この日についてNGを指定しません。曜日・回数など、ほかの条件は適用されます。 |
| 全NG | 日直・宿直ともに勤務できません。休日で選べます。 |
| 日NG | 日直に勤務できません。宿直は候補になります。休日で選べます。 |
| 宿NG | 宿直に勤務できません。休日なら日直は候補になります。 |

平日は「OK」「宿NG」から選びます。
「全日NGにする」は平日を宿NG、休日を全NGにし、「すべてOKに戻す」はNG指定を解除します。
一括操作は押すと保存されます。確定済み当直・優先度100以上の希望は、NGより優先されます。

**「保存」は、今開いている画面の計算条件への保存です。**
次回も使う場合は、下の「医師条件をCSVで保存」をご利用ください。
    """)

ng_layout = st.radio("NGカレンダーの表示", ["月間カレンダー", "1日〜月末を横一列"], horizontal=True, key="ng_calendar_layout")
ng_horizontal = ng_layout == "1日〜月末を横一列"
st.caption("表示を切り替える前に、選択中のNG日を保存してください。保存済みの内容は、どちらの表示でも共通です。")
if ng_horizontal:
    st.caption("左右へスクロールして日付を選べます。選択中の色はすぐに変わります。最後に「NG日を保存する」を押してください。")

valid_staff = staff_df[staff_df["先生の名前"].astype(str).str.strip() != ""]
if not valid_staff.empty:
    doctor_names = valid_staff["先生の名前"].astype(str).tolist()
    tabs = st.tabs(doctor_names)
    
    for t_idx, doc_name in enumerate(doctor_names):
        original_idx = valid_staff.index[t_idx]
        with tabs[t_idx]:
            
            hard_str = str(valid_staff.loc[original_idx].get("翌日PM duty", ""))
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
                        except (ValueError, TypeError):
                            pass
                    else:
                        try:
                            val = int(float(x.strip()))
                            if 1 <= val <= num_days:
                                current_ng_dict[val] = "全NG"
                        except (ValueError, TypeError):
                            pass
            
            for d in range(1, num_days + 1):
                chk_key = f"ng_{doc_name}_{year}_{month}_{d}"
                st.session_state[chk_key] = st.session_state.get(chk_key, current_ng_dict.get(d, "OK"))

            saved_strs = []
            for d in range(1, num_days + 1):
                val = st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", "OK")
                if val == "全NG": saved_strs.append(f"{d}日")
                elif val == "日NG": saved_strs.append(f"{d}日(日直NG)")
                elif val == "宿NG": saved_strs.append(f"{d}日(宿直NG)")
                
            if saved_strs:
                st.success(f"✅ **保存済みのNG日:** {', '.join(saved_strs)}")
            else:
                st.info("💡 **現在、保存されているNG日はありません**")

            import hashlib
            import json
            component_days = []
            for d in range(1, num_days + 1):
                dt = datetime.date(year, month, d)
                hol = dt.weekday() >= 5 or jpholiday.is_holiday(dt) or d in custom_holidays
                options = ["OK", "全NG", "日NG", "宿NG"] if hol else ["OK", "宿NG"]
                k = f"ng_{doc_name}_{year}_{month}_{d}"
                value = st.session_state.get(k, "OK")
                if value not in options:
                    value = "宿NG" if value == "全NG" else "OK"
                st.session_state[k] = value
                component_days.append({"day": d, "weekday": weekdays_ja[dt.weekday()],
                    "options": options, "value": value, "warning": dt.weekday() in hard_days,
                    "kind": "saturday" if dt.weekday() == 5 and not jpholiday.is_holiday(dt) and d not in custom_holidays else ("holiday" if hol else "weekday")})
            revision = hashlib.sha256(json.dumps([year, month, doc_name, ng_layout, component_days], ensure_ascii=False).encode()).hexdigest()
            component_key = f"ng_editor_{'row' if ng_horizontal else 'month'}_{doc_name}_{year}_{month}"
            st.markdown("<div style='color:#bf5700;background:#fff0c2;border:1px solid #ef9b20;border-radius:6px;padding:8px 10px;font-size:0.9rem;font-weight:600;'>⚠は、翌日PMにdutyがあるため、原則としてその日の宿直を外すことを示します。ただし、翌日が休日の場合は宿直に入ることがあります。</div>", unsafe_allow_html=True)
            response = horizontal_ng_component(HORIZONTAL_NG_HTML)(days=component_days, mode="row" if ng_horizontal else "month", offset=datetime.date(year, month, 1).weekday(), doctor=doc_name, version=revision, key=component_key, default=None)
            seen_key = component_key + "_last_token"
            if isinstance(response, dict) and response.get("token") != st.session_state.get(seen_key):
                st.session_state[seen_key] = response.get("token")
                values = response.get("values")
                if response.get("version") == revision and isinstance(values, list) and len(values) == num_days:
                    if all(v in item["options"] for v, item in zip(values, component_days)):
                        for d, value in enumerate(values, 1):
                            st.session_state[f"ng_{doc_name}_{year}_{month}_{d}"] = value
                        st.rerun()

            _, col_btn1, col_btn2 = st.columns([6, 1.5, 1.5])
            with col_btn1:
                st.button("全日NGにする", key=f"btn_all_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, "全NG", custom_holidays), use_container_width=True)
            with col_btn2:
                st.button("すべてOKに戻す", key=f"btn_clear_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, "OK", custom_holidays), use_container_width=True)
            
            # DataFrameへ状態を保存
            ng_items = []
            for d in range(1, num_days + 1):
                val = st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", "OK")
                if val == "全NG":
                    ng_items.append(str(d))
                elif val != "OK":
                    ng_items.append(f"{d}:{val}")
                    
            staff_df.at[original_idx, "NG日(半角カンマ区切り)"] = ",".join(ng_items)
            


st.divider()
st.markdown("### 📂 医師条件をCSVで保存（次回も使う場合）")
st.write("医師名・回数・勤務間隔・希望日・保存済みのNG日・備考を保存します。次回は「医師条件」のアップロード欄から読み込んでください。")
st.caption("保存前に、各医師の「NG日を保存する」を押してください。対象年月・特別休日・増員設定・確定済み当直・生成結果・色分けとそのメモは、このCSVには含まれません。")

current_csv = staff_df.to_csv(index=False).encode('utf-8-sig')
st.download_button(
    label="📥 医師条件をCSVで保存",
    data=current_csv,
    file_name=f"医師条件_{year}年{month}月.csv",
    mime="text/csv",
    use_container_width=True
)


# ==========================================
# 4. 当直計算ロジック（関数）
# ==========================================
def add_type_cap(model, worked, forced_vars, cap, bound):
    # 上限を緩めるのは、その枠で実際に確定した回数が上限を超える分だけ。
    extra=model.NewIntVar(0,bound,'fixed_type_extra')
    model.AddMaxEquality(extra,[0,sum(forced_vars)-cap])
    model.Add(sum(worked)<=cap+extra)


def _generate_shift_core(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None):
    _, num_days = calendar.monthrange(target_year, target_month)
    NIGHT_SHIFTS = ['A宿直', 'B宿直', '外来宿直']
    DAY_SHIFTS = ['A日直', 'B日直', '外来日直']

    def is_holiday(y, m, d):
        date = datetime.date(y, m, d)
        return date.weekday() >= 5 or jpholiday.is_holiday(date) or (d in custom_holidays)
    
    def safe_int(val, default_val):
        if pd.isna(val): return default_val
        try:
            return int(float(val))
        except (ValueError, TypeError):
            return default_val

    doctors = staff_df['先生の名前'].astype(str).tolist()
    ng_days_dict = {}
    req_days = {}          
    req_specific = {}      
    req_priority = {} 
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
            date_obj = parse_shift_date(row.get('日付', ''), target_year, target_month)
            if date_obj is None:
                continue

            m = date_obj.month
            d = date_obj.day

            for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
                if s_type in row and pd.notna(row[s_type]):
                    doc_vals = re.split(r'[、,]+', str(row[s_type]))
                    for doc_val in doc_vals:
                        doc_val = doc_val.strip()
                        if doc_val in doctors:
                            if date_obj.year == target_year and m == target_month:
                                absolute_req_specific[doc_val].append((d, s_type))
                            elif date_obj < datetime.date(target_year, target_month, 1):
                                past_worked_dates[doc_val].append(date_obj)
                            else:
                                future_worked_dates[doc_val].append(date_obj)

    for index, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        
        hard_str = str(row.get('翌日PM duty', ''))
        hard_days_list = []
        for i, w in enumerate(["月", "火", "水", "木", "金", "土", "日"]):
            if w in hard_str:
                hard_days_list.append(i)
        hard_weekdays[doc] = hard_days_list
        
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
                    except (ValueError, TypeError):
                        pass
                else:
                    try:
                        d_val = int(float(x.strip()))
                        ng_dict[d_val] = "全NG"
                    except (ValueError, TypeError):
                        pass
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
                    if not item:
                        continue
                    if ':' in item:
                        parts = item.split(':')
                        try:
                            d = int(parts[0].strip())
                            s_name = parts[1].strip()
                            req_specific[doc].append((d, s_name))
                        except (ValueError, TypeError):
                            pass
                    else:
                        try:
                            req_days[doc].append(int(item))
                        except (ValueError, TypeError):
                            pass

        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('最低空ける日数'), 5)
        min_shifts_total[doc] = safe_int(row.get('月間最小回数'), 0)
        max_shifts_total[doc] = safe_int(row.get('月間最大回数'), 5)
        max_hol_shifts_per_doc[doc] = safe_int(row.get('休日最大回数'), 4)

        max_shifts_per_type[doc] = {
            'A宿直': safe_int(row.get('A宿直上限'), 2),
            'B宿直': safe_int(row.get('B宿直上限'), 2),
            '外来宿直': safe_int(row.get('外来宿直上限'), 2),
            'A日直': safe_int(row.get('A日直上限'), 2),
            'B日直': safe_int(row.get('B日直上限'), 2),
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
            
        absolute_req_specific[doc] = sorted(set(absolute_req_specific[doc]))
        specified_days = {d for d,s in absolute_req_specific[doc]}
        absolute_req_days[doc] = sorted(set(absolute_req_days[doc]) - specified_days)
        past_worked_dates[doc] = sorted(set(past_worked_dates[doc]))
        future_worked_dates[doc] = sorted(set(future_worked_dates[doc]))
        all_abs_dates = absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]]
        ng_days_dict[doc] = {d: v for d, v in ng_days_dict[doc].items() if d not in all_abs_dates}

    daily_active_shifts = {}
    for d in range(1, num_days + 1):
        base_shifts = NIGHT_SHIFTS + DAY_SHIFTS if is_holiday(target_year, target_month, d) else NIGHT_SHIFTS
        forced_shifts = [s for doc in doctors for sd, s in absolute_req_specific[doc] if sd == d and s in (NIGHT_SHIFTS + DAY_SHIFTS)]
        daily_active_shifts[d] = list(set(base_shifts + forced_shifts))

    for doc in doctors:
        for d,slot in req_specific[doc]:
            if 1 <= d <= num_days and slot not in daily_active_shifts[d]:
                invalid_requests.append(f'{doc}：{d}日の{slot}は設定されていません。日直なら特別休日を設定するか希望日を修正してください。')

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

    # NG日の処理（全NG / 日NG / 宿NG を区別してブロック）
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

    # 翌日PM dutyは「宿直系」のみNG。ただし【翌日が休日】の場合はOKとする
    for doc in doctors:
        for d in range(1, num_days + 1):
            date_obj = datetime.date(target_year, target_month, d)
            next_date = date_obj + datetime.timedelta(days=1)
            
            next_is_hol = next_date.weekday() >= 5 or jpholiday.is_holiday(next_date)
            if next_date.year == target_year and next_date.month == target_month:
                if next_date.day in custom_holidays:
                    next_is_hol = True
                    
            if (
                date_obj.weekday() in hard_weekdays[doc]
                and not next_is_hol
                and d not in absolute_req_days[doc]
                and not any(sd == d for sd, _ in absolute_req_specific[doc])
            ):
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
                forced_vars = [shifts[(d, doc, s_type)] for d in range(1, num_days + 1) if s_type in daily_active_shifts[d] and ((d, s_type) in absolute_req_specific[doc] or d in absolute_req_days[doc])]
                add_type_cap(model, worked, forced_vars, max_shifts_per_type[doc][s_type], num_days)

    min_shortfalls = {}
    for doc in doctors:
        min_shortfalls[doc] = model.NewIntVar(0, min_shifts_total[doc], f'min_shortfall_{doc}')
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

    # ──────────────────────────────────────────────────────────
    # 勤務間隔制約：確定同士を維持し、自動割当の間隔を守る
    # ──────────────────────────────────────────────────────────
    add_interval_constraints(
        model=model,
        shifts=shifts,
        doctors=doctors,
        daily_active_shifts=daily_active_shifts,
        num_days=num_days,
        target_year=target_year,
        target_month=target_month,
        min_intervals=min_intervals,
        past_worked_dates=past_worked_dates,
        future_worked_dates=future_worked_dates,
        absolute_req_days=absolute_req_days,
        absolute_req_specific=absolute_req_specific,
    )

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
        
    global_max = num_days * 6 
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
        if status == cp_model.FEASIBLE:
            warnings.append("必要人数を満たす案です。時間内に最適化を完了したことまでは確認できていません。")
        if over_cap_warnings:
            warnings.append("⚠️ **【重要】以下の枠は「決定済み当直」や「優先度100」が重なったため、AIが自動的に定員を拡張（2名以上配置）して当直を完成させました:**")
            warnings.extend([f"・{w}" for w in over_cap_warnings])
            
        return pd.DataFrame(schedule_list, columns=["日付", "平日/休日", "A日直", "A宿直", "B日直", "B宿直", "外来日直", "外来宿直"]), True, warnings, past_worked_dates, future_worked_dates
    
    else:
        # =========================================================
        # バックアップ（緩和モデル）
        # =========================================================
        reasons = []
        if status == cp_model.MODEL_INVALID:
            return None, False, ["計算モデルが無効です。入力条件と数値の範囲を確認してください。"], None, None
        if status == cp_model.UNKNOWN:
            reasons.append("通常計算では時間内に案を見つけられませんでした。条件が不可能と確定したわけではありません。")
        try:
            relax_model = cp_model.CpModel()
            r_shifts = {}
            dummies = {}
            r_excess = []

            for d in range(1, num_days + 1):
                for s in daily_active_shifts[d]:
                    dummies[(d, s)] = relax_model.NewIntVar(0, 10, f'dummy_d{d}_{s}')
                    for doc in doctors:
                        r_shifts[(d, doc, s)] = relax_model.NewBoolVar(f'r_shift_d{d}_{doc}_{s}')

            for d in range(1, num_days + 1):
                for s in daily_active_shifts[d]:
                    req_count = multi_slots_dict.get((d, s), 1)
                    fixed_docs_count = sum(1 for doc in doctors if (d, s) in absolute_req_specific[doc])
                    actual_req_count = max(req_count, fixed_docs_count)
                    overflow = relax_model.NewIntVar(0, len(doctors), f'r_over_{d}_{s}')
                    r_excess.append(overflow)
                    relax_model.Add(sum(r_shifts[(d, doc, s)] for doc in doctors) + dummies[(d, s)] == actual_req_count + overflow)

            for doc in doctors:
                for d in range(1, num_days + 1):
                    fixed_count = sum(1 for sd, ss in absolute_req_specific[doc] if sd == d and ss in daily_active_shifts[d])
                    max_shifts_today = max(1, fixed_count)
                    relax_model.Add(sum(r_shifts[(d, doc, s)] for s in daily_active_shifts[d]) <= max_shifts_today)

                for d, ng_type in ng_days_dict[doc].items():
                    if 1 <= d <= num_days:
                        if ng_type == "全NG":
                            for s in daily_active_shifts[d]:
                                relax_model.Add(r_shifts[(d, doc, s)] == 0)
                        elif ng_type == "日NG":
                            for s in daily_active_shifts[d]:
                                if s in DAY_SHIFTS:
                                    relax_model.Add(r_shifts[(d, doc, s)] == 0)
                        elif ng_type == "宿NG":
                            for s in daily_active_shifts[d]:
                                if s in NIGHT_SHIFTS:
                                    relax_model.Add(r_shifts[(d, doc, s)] == 0)

                for d in range(1, num_days + 1):
                    date_obj = datetime.date(target_year, target_month, d)
                    next_date = date_obj + datetime.timedelta(days=1)
                    
                    next_is_hol = next_date.weekday() >= 5 or jpholiday.is_holiday(next_date)
                    if next_date.year == target_year and next_date.month == target_month:
                        if next_date.day in custom_holidays:
                            next_is_hol = True
                            
                    if (
                        date_obj.weekday() in hard_weekdays[doc]
                        and not next_is_hol
                        and d not in absolute_req_days[doc]
                        and not any(sd == d for sd, _ in absolute_req_specific[doc])
                    ):
                        for s in NIGHT_SHIFTS:
                            if s in daily_active_shifts[d]:
                                relax_model.Add(r_shifts[(d, doc, s)] == 0)

                for d in absolute_req_days[doc]:
                    specifics_on_d = [s for sd, s in absolute_req_specific[doc] if sd == d]
                    if not specifics_on_d:
                        relax_model.AddExactlyOne(r_shifts[(d, doc, s)] for s in daily_active_shifts[d])

                for d, s_name in absolute_req_specific[doc]:
                    if s_name in daily_active_shifts[d]:
                        relax_model.Add(r_shifts[(d, doc, s_name)] == 1)

                for s_type in NIGHT_SHIFTS + DAY_SHIFTS:
                    worked = [r_shifts[(d, doc, s_type)] for d in range(1, num_days + 1) if s_type in daily_active_shifts[d]]
                    if worked:
                        specific_req_count = sum(1 for d, s in absolute_req_specific[doc] if s == s_type)
                        forced_vars = [r_shifts[(d, doc, s_type)] for d in range(1, num_days + 1) if s_type in daily_active_shifts[d] and ((d, s_type) in absolute_req_specific[doc] or d in absolute_req_days[doc])]
                        add_type_cap(relax_model, worked, forced_vars, max_shifts_per_type[doc][s_type], num_days)

                worked_all = []
                for d in range(1, num_days + 1):
                    for s in daily_active_shifts[d]:
                        worked_all.append(r_shifts[(d, doc, s)])
                if worked_all:
                    all_abs_dates = absolute_req_days[doc] + [d for (d, s) in absolute_req_specific[doc]]
                    actual_max_total = max(max_shifts_total[doc], len(all_abs_dates))
                    relax_model.Add(sum(worked_all) <= actual_max_total)

                hol_shifts = []
                for d in range(1, num_days + 1):
                    if is_holiday(target_year, target_month, d):
                        for s in daily_active_shifts[d]:
                            if s in NIGHT_SHIFTS + DAY_SHIFTS:
                                hol_shifts.append(r_shifts[(d, doc, s)])
                abs_hol_count = sum(1 for d in absolute_req_days[doc] if is_holiday(target_year, target_month, d))
                abs_hol_count += sum(1 for d, s in absolute_req_specific[doc] if is_holiday(target_year, target_month, d))
                actual_hol_max = max(max_hol_shifts_per_doc[doc], abs_hol_count)
                relax_model.Add(sum(hol_shifts) <= actual_hol_max)

            # ──────────────────────────────────────────────────────────
            # 【改善】緩和モデルも同じ関数で間隔制約を追加
            # ──────────────────────────────────────────────────────────
            add_interval_constraints(
                model=relax_model,
                shifts=r_shifts,
                doctors=doctors,
                daily_active_shifts=daily_active_shifts,
                num_days=num_days,
                target_year=target_year,
                target_month=target_month,
                min_intervals=min_intervals,
                past_worked_dates=past_worked_dates,
                future_worked_dates=future_worked_dates,
                absolute_req_days=absolute_req_days,
                absolute_req_specific=absolute_req_specific,
            )

            # 不足人数を優先群ごとに最小化する。
            # 第1群：A/B宿直・A/B日直（同順位）、第2群：外来宿直、第3群：外来日直。
            # 各不足変数の上限は10。下位群全体の最大損失より大きい重みを使う。
            primary_missing = [v for (d, s), v in dummies.items()
                               if s in ("A宿直", "B宿直", "A日直", "B日直")]
            outpatient_night_missing = [v for (d, s), v in dummies.items() if s == "外来宿直"]
            outpatient_day_missing = [v for (d, s), v in dummies.items() if s == "外来日直"]
            outpatient_night_weight = 10 * len(outpatient_day_missing) + 1
            primary_weight = (
                10 * len(outpatient_night_missing) * outpatient_night_weight
                + 10 * len(outpatient_day_missing) + 1
            )
            excess_bound = len(r_excess) * len(doctors)
            relax_model.Minimize(
                (primary_weight * sum(primary_missing)
                + outpatient_night_weight * sum(outpatient_night_missing)
                + sum(outpatient_day_missing)) * (excess_bound + 1)
                + sum(r_excess)
            )

            relax_solver = cp_model.CpSolver()
            relax_solver.parameters.max_time_in_seconds = 15.0
            relax_status = relax_solver.Solve(relax_model)

            if relax_status == cp_model.OPTIMAL or relax_status == cp_model.FEASIBLE:
                bottlenecks = []
                missing_by_shift = {s: 0 for s in NIGHT_SHIFTS + DAY_SHIFTS}
                
                partial_schedule_list = []
                weekday_ja = ["月", "火", "水", "木", "金", "土", "日"]
                
                for d in range(1, num_days + 1):
                    date_obj = datetime.date(target_year, target_month, d)
                    day_str = "休日" if is_holiday(target_year, target_month, d) else "平日"
                    row_dict = {"日付": f"{target_month}/{d}({weekday_ja[date_obj.weekday()]})", "平日/休日": day_str}
                    
                    for s in NIGHT_SHIFTS + DAY_SHIFTS:
                        row_dict[s] = "-"
                        
                    for s in daily_active_shifts[d]:
                        val = relax_solver.Value(dummies[(d, s)])
                        if val > 0:
                            bottlenecks.append(f"・{target_month}/{d} の「{s}」")
                            missing_by_shift[s] += val

                        assigned_docs = []
                        for doc in doctors:
                            if relax_solver.Value(r_shifts[(d, doc, s)]) == 1:
                                assigned_docs.append(doc)
                        
                        if val > 0:
                            assigned_docs.append(f"⚠️不足({val}名)")
                            
                        if assigned_docs:
                            row_dict[s] = "、".join(assigned_docs)
                            
                    partial_schedule_list.append(row_dict)

                partial_df = pd.DataFrame(partial_schedule_list, columns=["日付", "平日/休日", "A日直", "A宿直", "B日直", "B宿直", "外来日直", "外来宿直"])

                if bottlenecks:
                    reasons.append("A宿直・B宿直・A日直・B日直を同順位で最優先とし、次に外来宿直、最後に外来日直の不足を減らす方針で作成しました。条件によっては優先枠にも不足が残ります。")
                    if relax_status == cp_model.FEASIBLE:
                        reasons.append("計算時間内に得られた案です。優先順位に沿った不足の最小化が完了したことまでは確認できていません。")
                    reasons.append("🚨 **以下の枠で必要人数が不足しています:**")
                    reasons.extend(bottlenecks)
                    reasons.append("")
                    reasons.append("📊 **【不足している枠の合計】**")
                    sorted_missing = sorted(missing_by_shift.items(), key=lambda x: x[1], reverse=True)
                    for s, count in sorted_missing:
                        if count > 0:
                            reasons.append(f"・{s}： 計 {count} 枠不足")
                    
                return partial_df, False, reasons, past_worked_dates, future_worked_dates
            if relax_status == cp_model.INFEASIBLE:
                reasons.append("不足枠を許容しても条件が矛盾し、当直案を作成できませんでした。確定指定と上限を確認してください。")
            elif relax_status == cp_model.MODEL_INVALID:
                reasons.append("不足枠計算のモデルが無効です。入力値を確認してください。")
            else:
                reasons.append("不足枠を含む案も時間内に見つかりませんでした。条件が不可能と確定したわけではありません。")
        except Exception as e:
            reasons.append(f"⚠️ 部分的な当直表の作成中にもエラーが発生しました。詳細: {e}")

        return None, False, reasons, None, None

# ==========================================
# 5. 実行ボタンと結果表示
# ==========================================
st.divider()
st.markdown("""
<style>
.st-key-create_duty_action button {
    width: 100%;
    min-height: 60px;
}
.st-key-create_duty_action button p {
    font-size: 1.15rem;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)
create_button_container = st.container(key="create_duty_action")
st.header("3. 当直案の作成・確認")
result_notice_container = st.container()

staff_df = staff_df[staff_df['先生の名前'].astype(str).str.strip() != '']
staff_df = staff_df.dropna(subset=['先生の名前']).reset_index(drop=True)

fixed_df, fixed_input_errors = validate_fixed_inputs(edited_fixed_df, staff_df['先生の名前'].tolist(), year, month)
if fixed_input_errors:
    show_input_errors(fixed_input_errors)
    st.stop()
current_signature = input_signature(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df)
if 'generated_df' in st.session_state and st.session_state.get('generated_signature') != current_signature:
    for key in ['generated_df','past_worked_dates','future_worked_dates','generated_warnings','generated_signature','generated_year','generated_month']:
        st.session_state.pop(key, None)
    st.session_state['result_needs_refresh'] = True
if st.session_state.get('result_needs_refresh'):
    st.warning("条件が変更されています。再作成してください。前回の結果表示とダウンロードを停止しました。")

if len(staff_df) > 0:
    with create_button_container:
        create_clicked = st.button("🚀 この条件で当直案を作成する", type="primary", use_container_width=True)
    if create_clicked:
        with st.spinner("当直案を計算中…（通常は最大60秒、不足枠の確認を含む場合は計算時間が最大75秒です）"):
            try:
                df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df)
                
                if df_result is not None:
                    st.session_state['generated_signature'] = current_signature
                    st.session_state['generated_year'] = year
                    st.session_state['generated_month'] = month
                    st.session_state['generated_warnings'] = error_reasons
                    st.session_state['result_needs_refresh'] = False
                if df_result is not None and not df_result.empty:
                    st.session_state['generated_df'] = df_result
                    st.session_state['past_worked_dates'] = past_worked_dates or {}
                    st.session_state['future_worked_dates'] = future_worked_dates or {}
                else:
                    st.session_state.pop('generated_df', None)
                    with result_notice_container:
                        show_schedule_notice(None, error_reasons)
            except Exception as e:
                st.session_state.pop('generated_df', None)
                with result_notice_container:
                    show_schedule_notice(None, [f"当直計算中にエラーが発生しました。詳細: {e}"])

    if 'generated_df' in st.session_state:
        year = st.session_state['generated_year']
        month = st.session_state['generated_month']
        with result_notice_container:
            show_schedule_notice(st.session_state['generated_df'], st.session_state.get('generated_warnings', []))
        df_result = st.session_state['generated_df'].reindex(columns=["日付", "平日/休日", "A日直", "A宿直", "B日直", "B宿直", "外来日直", "外来宿直"])
        past_worked_dates = st.session_state.get('past_worked_dates', {})
        future_worked_dates = st.session_state.get('future_worked_dates', {})
        
        shift_columns = ['A宿直', 'B宿直', '外来宿直', 'A日直', 'B日直', '外来日直']
        doctors_list = staff_df['先生の名前'].astype(str).tolist()
        
        st.subheader("📅 作成した当直案")
        
        table_container = st.container()
        
        st.divider()
        st.subheader("🔍 特定の医師の当直を色別でハイライト")
        st.write("※各色のすぐ下にあるメモ欄に「神経内科」「呼吸器内科」など自由に書き込めます。")
        
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
            
            if "⚠️不足" in val_str:
                return 'background-color: #ffe6e6; color: #cc0000; font-weight: bold; border: 2px solid #cc0000;'
            
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
        styled_df = base_style.map(color_highlighted_doctor, subset=shift_columns)
        
        result_height = len(df_result) * 35 + 40
        
        with table_container:
            # 表の実測高を通知するコンポーネントで余分な空白を作らない。
            horizontal_ng_component(
                build_hover_schedule_html(df_result, shift_columns, doctors_list, color_highlighted_doctor)
            )(key="hover_schedule_table", default=None)

            csv_result = df_result.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 表示中の当直案をCSVでダウンロード",
                data=csv_result,
                file_name=f"当直案_{year}年{month}月.csv",
                mime="text/csv",
            )
        
        st.divider()
        st.subheader("📊 当直案の担当回数・希望日・勤務間隔")
        st.caption("回数は作成した当直案の集計です。最小・平均間隔は読み込んだ月外の勤務も含む、勤務日と次の勤務日の間の日数です。同じ日の複数勤務は、間隔の集計では1日として扱います。")
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
                            except (ValueError, TypeError):
                                pass
                        else:
                            try:
                                req_days_eval[doc].append(int(item))
                            except (ValueError, TypeError):
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
                        
            doc_data["宿直回数"] = doc_data.get("A宿直", 0) + doc_data.get("B宿直", 0) + doc_data.get("外来宿直", 0)
            doc_data["日直回数"] = doc_data.get("A日直", 0) + doc_data.get("B日直", 0) + doc_data.get("外来日直", 0)
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
        df_summary = df_summary[['先生の名前', 'A宿直', 'B宿直', '外来宿直', 'A日直', 'B日直', '外来日直', '宿直回数', '日直回数', '休日回数', '総合計', '希望日達成', '最小間隔', '平均間隔']]
        
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
        


elif len(staff_df) == 0:
    st.warning("☝️ 表に先生の名前を入力するか、CSVファイルをアップロードしてください。")
