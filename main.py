import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import datetime
import jpholiday
import calendar
import io
import unicodedata

# 1. ページ基本設定
st.set_page_config(page_title="看護師勤務表作成AI", layout="wide")
st.title("看護師勤務表作成AI🩺✨（新人ペア禁止・完全版）")

def clean_text(text):
    if not isinstance(text, str): return str(text)
    text = text.replace(" ", "").replace("　", "")
    return unicodedata.normalize('NFKC', text).strip()

# --- 2. テンプレート配布機能 ---
st.sidebar.header("📁 ステップ1：名簿の準備")

def create_template():
    base_cols = ["名前", "役職", "区分", "交代", "前月最終"]
    hope_cols = [f"{i}日希望" for i in range(1, 32)]
    cols = base_cols + hope_cols
    data = []
    for i in range(1, 30):
        yaku = "主任" if i <= 7 else "一般"
        kubun = "既卒" if i <= 26 else "新人"
        kotai = 3 if 14 <= i <= 25 else 2
        data.append([i, yaku, kubun, kotai, ""] + [""] * 31)
    template_df = pd.DataFrame(data, columns=cols)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False)
    return output.getvalue()

st.sidebar.download_button(
    label="👉 サンプル入りExcel(29名分)をダウンロード",
    data=create_template(),
    file_name="meibo_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.markdown("---")
uploaded_file = st.sidebar.file_uploader("編集した名簿(Excel)をアップロードしてください", type="xlsx")

# --- 3. メイン計算ロジック ---
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [clean_text(c) for c in df.columns]
        st.success(f"名簿（{len(df)}名）の読み込みに成功しました。")

        # 設定
        st.sidebar.header("📅 ステップ2：設定")
        year = st.sidebar.number_input("作成年", value=2026)
        month = st.sidebar.number_input("作成月", value=1, min_value=1, max_value=12)
        _, num_days = calendar.monthrange(year, month)
        
        h_dates = [datetime.date(year, month, d) for d in range(1, num_days + 1)]
        h_count = sum(1 for dt in h_dates if dt.weekday() >= 5 or jpholiday.is_holiday(dt))
        st.sidebar.info(f"💡 今月の基本公休数: {h_count}日")

        st.sidebar.header("👥 人数設定")
        req_day_wk = st.sidebar.slider("平日日勤（目標）", 1, 20, 10)
        req_day_hol = st.sidebar.slider("休日日勤（完全固定）", 1, 20, 4)
        req_semi = st.sidebar.slider("準夜（固定）", 1, 5, 2)
        req_late = st.sidebar.slider("深夜（固定）", 1, 5, 2)

        if st.button("最強ルールで勤務表を生成する"):
            model = cp_model.CpModel()
            num_nurses = len(df)
            shifts = ["日勤", "準夜", "深夜", "休み"]
            
            x = {}
            for n in range(num_nurses):
                for d in range(num_days):
                    for s in shifts:
                        x[n, d, s] = model.NewBoolVar(f'n{n}_d{d}_s{s}')

            penalties = []
            hopes_map = {}

            # --- 既卒・新人のリストを作成 ---
            novice_indices = [n for n in range(num_nurses) if clean_text(str(df.iloc[n].get('区分', ''))) == "新人"]

            # --- 基本制約 ---
            for n in range(num_nurses):
                for d in range(num_days):
                    model.Add(sum(x[n, d, s] for s in shifts) == 1)
                    col = f"{d+1}日希望"
                    if col in df.columns:
                        val = clean_text(str(df.iloc[n][col]))
                        mapping = {"休":"休み", "日":"日勤", "準":"準夜", "深":"深夜"}
                        if val in mapping:
                            model.Add(x[n, d, mapping[val]] == 1)
                            hopes_map[(n, d)] = mapping[val]

            # --- 最強ルール：新人同士の夜勤ペア禁止（ここを追加！） ---
            for d in range(num_days):
                if len(novice_indices) > 0:
                    # 同じ日の準夜に新人は最大1人まで
                    model.Add(sum(x[n, d, "準夜"] for n in novice_indices) <= 1)
                    # 同じ日の深夜に新人は最大1人まで
                    model.Add(sum(x[n, d, "深夜"] for n in novice_indices) <= 1)

            # --- 最強ルール：前月継続性 ---
            for n in range(num_nurses):
                if '前月最終' in df.columns:
                    last = clean_text(str(df.iloc[n]['前月最終']))
                    is_3 = (int(df.iloc[n].get('交代', 2)) == 3)
                    if is_3:
                        if "深夜" in last:
                            model.Add(x[n, 0, "準夜"] == 1); model.Add(x[n, 1, "休み"] == 1)
                    else:
                        if "準夜" in last:
                            model.Add(x[n, 0, "深夜"] == 1); model.Add(x[n, 1, "休み"] == 1)
                        elif "深夜" in last:
                            model.Add(x[n, 0, "休み"] == 1)

            # --- 最強ルール：セット勤務とインターバル ---
            for n in range(num_nurses):
                is_3 = (int(df.iloc[n].get('交代', 2)) == 3)
                for d in range(num_days):
                    if is_3: # 3交代
                        if d < num_days - 1:
                            model.Add(x[n, d+1, "準夜"] == 1).OnlyEnforceIf(x[n, d, "深夜"])
                            model.Add(x[n, d, "深夜"] == 1).OnlyEnforceIf(x[n, d+1, "準夜"])
                        if d < num_days - 2:
                            model.Add(x[n, d+2, "休み"] == 1).OnlyEnforceIf(x[n, d, "深夜"])
                        if d > 0:
                            model.Add(x[n, d-1, "日勤"] == 1).OnlyEnforceIf(x[n, d, "深夜"])
                        if d < num_days - 6:
                            for i in range(1, 6):
                                model.Add(x[n, d+i, "深夜"] == 0).OnlyEnforceIf(x[n, d, "深夜"])
                    else: # 2交代
                        if d < num_days - 1:
                            model.Add(x[n, d+1, "深夜"] == 1).OnlyEnforceIf(x[n, d, "準夜"])
                            model.Add(x[n, d, "準夜"] == 1).OnlyEnforceIf(x[n, d+1, "深夜"])
                        if d < num_days - 2:
                            model.Add(x[n, d+2, "休み"] == 1).OnlyEnforceIf(x[n, d+1, "深夜"])
                        if d < num_days - 5:
                            for i in range(1, 5):
                                model.Add(x[n, d+i, "準夜"] == 0).OnlyEnforceIf(x[n, d, "準夜"])

            # --- 公平性と公休 ---
            for n in range(num_nurses):
                model.Add(sum(x[n, d, "休み"] for d in range(num_days)) == h_count)
                f_h = sum(x[n, d, "準夜"] + x[n, d, "深夜"] for d in range(min(15, num_days)))
                s_h = sum(x[n, d, "準夜"] + x[n, d, "深夜"] for d in range(min(15, num_days), num_days))
                diff = model.NewIntVar(0, 5, f'df_{n}')
                model.Add(diff >= f_h - s_h); model.Add(diff >= s_h - f_h); model.Add(diff <= 2)

            nt = [model.NewIntVar(0, num_days, f'nt_{n}') for n in range(num_nurses)]
            for n in range(num_nurses):
                model.Add(nt[n] == sum(x[n, d, "準夜"] + x[n, d, "深夜"] for d in range(num_days)))
            mi, ma = model.NewIntVar(0, num_days, 'mi'), model.NewIntVar(0, num_days, 'ma')
            for n in range(num_nurses):
                model.Add(mi <= nt[n]); model.Add(ma >= nt[n])
            model.Add(ma - mi <= 4)

            # --- 人数制限 ---
            for d in range(num_days):
                model.Add(sum(x[n, d, "準夜"] for n in range(num_nurses)) == req_semi)
                model.Add(sum(x[n, d, "深夜"] for n in range(num_nurses)) == req_late)
                is_h = (datetime.date(year, month, d+1).weekday() >= 5) or jpholiday.is_holiday(datetime.date(year, month, d+1))
                t = req_day_hol if is_h else req_day_wk
                if is_h:
                    model.Add(sum(x[n, d, "日勤"] for n in range(num_nurses)) == t)
                else:
                    u, o = model.NewIntVar(0, num_nurses, f'u{d}'), model.NewIntVar(0, num_nurses, f'o{d}')
                    model.Add(sum(x[n, d, "日勤"] for n in range(num_nurses)) + u - o == t)
                    penalties.append(u * 100); penalties.append(o * 10)

            # 解決
            model.Minimize(sum(penalties))
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 60.0
            status = solver.Solve(model)

            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                d_cls = [f"{d+1}({['月','火','水','木','金','土','日'][datetime.date(year,month,d+1).weekday()]})" for d in range(num_days)]
                
                # A. 毎日の集計
                st.subheader("📊 毎日の合計人数（縦の合計）")
                summ_list = []
                for s in ["日勤", "準夜", "深夜", "休み"]:
                    row = {"シフト": s}
                    for d in range(num_days):
                        row[d_cls[d]] = sum(solver.Value(x[n, d, s]) for n in range(num_nurses))
                    summ_list.append(row)
                st.table(pd.DataFrame(summ_list))

                # B. 詳細表示
                st.subheader("📋 勤務表詳細")
                
                def style_output(res):
                    styled = pd.DataFrame('', index=res.index, columns=res.columns)
                    for r in range(len(res)):
                        for di, col in enumerate(d_cls):
                            val = res.iloc[r][col]
                            bg = ""
                            if val == '深夜': bg = "background-color: #ffcccc; color: #900; font-weight: bold;"
                            elif val == '準夜': bg = "background-color: #fff0cc; color: #960; font-weight: bold;"
                            elif val == '休み': bg = "color: #bbb;"
                            if (r, di) in hopes_map:
                                bg += "border: 2px solid #00acc1; background-color: #e0f7fa;"
                            styled.iloc[r, styled.columns.get_loc(col)] = bg
                    return styled

                final_data = []
                for n in range(num_nurses):
                    c = {s: sum(solver.Value(x[n, d, s]) for d in range(num_days)) for s in shifts}
                    row = [df.iloc[n]['名前'], df.iloc[n].get('役職',''), df.iloc[n].get('区分',''), f"{df.iloc[n].get('交代',2)}交代", c["日勤"], c["準夜"], c["深夜"], c["休み"]]
                    for d in range(num_days):
                        for s in shifts:
                            if solver.Value(x[n, d, s]): row.append(s)
                    final_data.append(row)
                
                res_df = pd.DataFrame(final_data, columns=["名前", "役職", "区分", "交代", "日勤", "準", "深", "休"] + d_cls)
                st.dataframe(res_df.style.apply(style_output, axis=None), height=600)

                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as wr: res_df.to_excel(wr, index=False)
                st.download_button("Excelで保存", data=out.getvalue(), file_name=f"kimmubyo_{year}_{month}.xlsx")
            else:
                st.error("❌ 条件が厳しすぎます（新人の夜勤枠が足りない可能性があります）。設定を調整してください。")
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")