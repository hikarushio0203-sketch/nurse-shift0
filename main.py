import streamlit as st
import pandas as pd  
from ortools.sat.python import cp_model
import datetime
import jpholiday
import calendar
import io
import unicodedata

# --- 1. ページ基本設定 ---
st.set_page_config(page_title="看護師勤務表作成AI", layout="wide")
st.title("勤務表自動作成ソフト🩺✨")
st.markdown("### ★2交代セット間隔制限・修正ハイライト・全ルール徹底版★")

def clean_text(text):
    if not isinstance(text, str): return str(text)
    text = text.replace(" ", "").replace("　", "")
    return unicodedata.normalize('NFKC', text).strip()

# --- 2. セッション状態（データの保持）の初期化 ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None
if 'hopes_map' not in st.session_state:
    st.session_state.hopes_map = {}
if 'modified_map' not in st.session_state:
    st.session_state.modified_map = {}

# --- 3. テンプレート配布・アップロード ---
st.sidebar.header("📁 ステップ1：名簿の準備")
def create_template():
    base_cols = ["名前", "役職", "区分", "交代", "前月最終"]
    hope_cols = [f"{i}日希望" for i in range(1, 32)]
    cols = base_cols + hope_cols
    data = [[f"看護師{i}", "一般", "既卒", 2, ""] + [""] * 31 for i in range(1, 21)]
    template_df = pd.DataFrame(data, columns=cols)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False)
    return output.getvalue()

st.sidebar.download_button("👉 テンプレートExcelをダウンロード", data=create_template(), file_name="meibo_template.xlsx")
uploaded_file = st.sidebar.file_uploader("編集した名簿をアップロードしてください", type="xlsx")

# --- 4. メイン計算ロジック ---
if uploaded_file is not None:
    try:
        df_meibo = pd.read_excel(uploaded_file)
        df_meibo.columns = [clean_text(c) for c in df_meibo.columns]
        
        # 作成設定
        st.sidebar.header("📅 ステップ2：作成設定")
        year = st.sidebar.number_input("作成年", value=2026)
        month = st.sidebar.number_input("作成月", value=1, min_value=1, max_value=12)
        _, num_days = calendar.monthrange(year, month)
        h_dates = [datetime.date(year, month, d) for d in range(1, num_days + 1)]
        h_count = sum(1 for dt in h_dates if dt.weekday() >= 5 or jpholiday.is_holiday(dt))
        d_cls = [f"{d+1}({['月','火','水','木','金','土','日'][h_dates[d].weekday()]})" for d in range(num_days)]

        st.sidebar.header("👥 人数設定")
        req_day_wk = st.sidebar.slider("平日日勤（目標）", 1, 20, 10)
        req_day_hol = st.sidebar.slider("休日日勤（必ず固定）", 1, 20, 4)
        req_semi = st.sidebar.slider("準夜（固定）", 1, 5, 2)
        req_late = st.sidebar.slider("深夜（固定）", 1, 5, 2)
        night_diff_limit = st.sidebar.slider("個人間の夜勤合計回数差（許容）", 0, 5, 2)

        # --- AI生成エンジン ---
        if st.sidebar.button("最強ルールでAI生成を開始"):
            with st.spinner("AIが全ルールを検証しながら勤務表を構築中..."):
                model = cp_model.CpModel()
                num_nurses = len(df_meibo)
                shifts = ["日勤", "準夜", "深夜", "休み"]
                x = {(n, d, s): model.NewBoolVar(f'n{n}_d{d}_s{s}') for n in range(num_nurses) for d in range(num_days) for s in shifts}

                st.session_state.hopes_map = {}
                st.session_state.modified_map = {}
                novice_indices = [n for n in range(num_nurses) if "新人" in str(df_meibo.iloc[n].get('区分', ''))]

                for n in range(num_nurses):
                    rotation_type = int(df_meibo.iloc[n].get('交代', 2))
                    is_3 = (rotation_type == 3)
                    
                    # 1. 1日1勤務 & 公休数遵守
                    model.Add(sum(x[n, d, "休み"] for d in range(num_days)) == h_count)
                    for d in range(num_days):
                        model.Add(sum(x[n, d, s] for s in shifts) == 1)
                        # 希望反映
                        col = f"{d+1}日希望"
                        if col in df_meibo.columns:
                            val = clean_text(str(df_meibo.iloc[n][col]))
                            m = {"休":"休み", "日":"日勤", "準":"準夜", "深":"深夜"}
                            if val in m:
                                model.Add(x[n, d, m[val]] == 1)
                                st.session_state.hopes_map[(n, d)] = m[val]

                    # 2. 6連勤禁止 (最大5連勤まで)
                    for d in range(num_days - 5):
                        model.Add(sum(x[n, d + i, "休み"] for i in range(6)) >= 1)

                    # 3. 前月最終日接続
                    if '前月最終' in df_meibo.columns:
                        last = clean_text(str(df_meibo.iloc[n]['前月最終']))
                        if is_3 and "深夜" in last:
                            model.Add(x[n, 0, "準夜"] == 1); model.Add(x[n, 1, "休み"] == 1)
                        elif not is_3:
                            if "準夜" in last: model.Add(x[n, 0, "深夜"] == 1); model.Add(x[n, 1, "休み"] == 1)
                            elif "深夜" in last: model.Add(x[n, 0, "休み"] == 1)

                    # 4. 交代別セット勤務 & 間隔ルール
                    for d in range(num_days):
                        if is_3: # 3交代：深夜→準夜→休み
                            if d < num_days - 1:
                                model.Add(x[n, d+1, "準夜"] == 1).OnlyEnforceIf(x[n, d, "深夜"])
                                model.Add(x[n, d, "深夜"] == 1).OnlyEnforceIf(x[n, d+1, "準夜"])
                            if d < num_days - 2: model.Add(x[n, d+2, "休み"] == 1).OnlyEnforceIf(x[n, d, "深夜"])
                            if d > 0: model.Add(x[n, d-1, "日勤"] == 1).OnlyEnforceIf(x[n, d, "深夜"])
                        else: # 2交代：準夜→深夜→休み
                            if d < num_days - 1:
                                model.Add(x[n, d+1, "深夜"] == 1).OnlyEnforceIf(x[n, d, "準夜"])
                                model.Add(x[n, d, "準夜"] == 1).OnlyEnforceIf(x[n, d+1, "深夜"])
                            if d < num_days - 2:
                                model.Add(x[n, d+2, "休み"] == 1).OnlyEnforceIf(x[n, d+1, "深夜"])
                            # ★追加：2交代セットの終了後、翌日(d+2)は準夜不可（最低1日開ける）
                            if d < num_days - 2:
                                model.Add(x[n, d+2, "準夜"] == 0).OnlyEnforceIf([x[n, d, "深夜"], x[n, d+1, "休み"]])

                    # 5. 夜勤格差バランス
                    f_h = sum(x[n, d, "準夜"] + x[n, d, "深夜"] for d in range(min(15, num_days)))
                    s_h = sum(x[n, d, "準夜"] + x[n, d, "深夜"] for d in range(min(15, num_days), num_days))
                    diff_half = model.NewIntVar(0, 5, f'diff_half_{n}')
                    model.Add(diff_half >= f_h - s_h); model.Add(diff_half >= s_h - f_h); model.Add(diff_half <= 2)

                # 6. 新人ペア禁止
                for d in range(num_days):
                    if novice_indices:
                        model.Add(sum(x[n, d, "準夜"] for n in novice_indices) <= 1)
                        model.Add(sum(x[n, d, "深夜"] for n in novice_indices) <= 1)

                # 7. 夜勤合計平準化
                night_totals = [model.NewIntVar(0, num_days, f'nt_{n}') for n in range(num_nurses)]
                for n in range(num_nurses):
                    model.Add(night_totals[n] == sum(x[n, d, "準夜"] + x[n, d, "深夜"] for d in range(num_days)))
                mi_n, ma_n = model.NewIntVar(0, num_days, 'mi_n'), model.NewIntVar(0, num_days, 'ma_n')
                for n in range(num_nurses):
                    model.Add(mi_n <= night_totals[n]); model.Add(ma_n >= night_totals[n])
                model.Add(ma_n - mi_n <= night_diff_limit)

                # 8. 土日祝人数固定
                penalties = []
                for d in range(num_days):
                    model.Add(sum(x[n, d, "準夜"] for n in range(num_nurses)) == req_semi)
                    model.Add(sum(x[n, d, "深夜"] for n in range(num_nurses)) == req_late)
                    is_h = (h_dates[d].weekday() >= 5) or jpholiday.is_holiday(h_dates[d])
                    if is_h:
                        model.Add(sum(x[n, d, "日勤"] for n in range(num_nurses)) == req_day_hol)
                    else:
                        u, o = model.NewIntVar(0, num_nurses, f'u{d}'), model.NewIntVar(0, num_nurses, f'o{d}')
                        model.Add(sum(x[n, d, "日勤"] for n in range(num_nurses)) + u - o == req_day_wk)
                        penalties.append(u * 100 + o * 10)

                model.Minimize(sum(penalties))
                solver = cp_model.CpSolver()
                status = solver.Solve(model)

                if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                    final_res = []
                    for n in range(num_nurses):
                        n_sh = [next(s for s in shifts if solver.Value(x[n, d, s])) for d in range(num_days)]
                        row = [df_meibo.iloc[n]['名前'], df_meibo.iloc[n].get('役職',''), df_meibo.iloc[n].get('区分',''), f"{df_meibo.iloc[n].get('交代',2)}交代",
                               n_sh.count("日勤"), n_sh.count("準夜"), n_sh.count("深夜"), n_sh.count("休み")] + n_sh
                        final_res.append(row)
                    st.session_state.df_result = pd.DataFrame(final_res, columns=["名前", "役職", "区分", "交代", "日", "準", "深", "休"] + d_cls)
                else:
                    st.error("❌ 条件が厳しすぎます。夜勤回数差や土日人数を調整してください。")

        # --- 5. 修正パレット（個人合計自動連動） ---
        if st.session_state.df_result is not None:
            st.markdown("---")
            st.subheader("🛠 修正パレット（修正すると本人の合計数も自動で変わります）")
            with st.container():
                c1, c2, c3, c4 = st.columns([2, 3, 2, 2])
                p_sh = c1.selectbox("🎨 変更後の勤務", ["日勤", "準夜", "深夜", "休み"])
                p_na = c2.selectbox("👤 対象スタッフ", st.session_state.df_result["名前"].tolist())
                p_da = c3.selectbox("📅 日付", d_cls)
                if c4.button("⚡ 修正を確定する"):
                    row_idx = st.session_state.df_result[st.session_state.df_result["名前"] == p_na].index[0]
                    day_idx = d_cls.index(p_da)
                    
                    st.session_state.df_result.at[row_idx, p_da] = p_sh
                    
                    # 個人の合計列をリアルタイム更新
                    current_nurse_row = st.session_state.df_result.loc[row_idx, d_cls].tolist()
                    st.session_state.df_result.at[row_idx, "日"] = current_nurse_row.count("日勤")
                    st.session_state.df_result.at[row_idx, "準"] = current_nurse_row.count("準夜")
                    st.session_state.df_result.at[row_idx, "深"] = current_nurse_row.count("深夜")
                    st.session_state.df_result.at[row_idx, "休"] = current_nurse_row.count("休み")
                    
                    # 修正箇所を記録
                    st.session_state.modified_map[(row_idx, day_idx)] = True
                    st.rerun()

            # --- 6. 人数集計表示 ---
            st.subheader("📊 リアルタイム合計人数（日別）")
            sum_df = pd.DataFrame([{"シフト": s, **{d: (st.session_state.df_result[d] == s).sum() for d in d_cls}} for s in ["日勤", "準夜", "深夜", "休み"]])
            st.table(sum_df)

            # --- 7. 勤務表表示（色分け & ハイライト） ---
            st.subheader("📋 勤務表詳細")
            
            # 勤務の色分け
            def style_cell(v):
                if v == '深夜': return 'background-color: #ffcccc; color: #900; font-weight: bold;'
                if v == '準夜': return 'background-color: #fff0cc; color: #960; font-weight: bold;'
                if v == '休み': return 'color: #bbb;'
                return ''

            # 希望日 & 修正日のハイライト（インデックスエラー修正版）
            def style_highlight(data):
                # 元のDataFrameと同じ形状の空のDataFrameを作成
                attr = pd.DataFrame('', index=data.index, columns=data.columns)
                # 希望日の強調
                for (n, d) in st.session_state.hopes_map.keys():
                    if n < len(data) and d < len(d_cls):
                        attr.at[data.index[n], d_cls[d]] = 'border: 2px solid #00acc1; background-color: #e0f7fa;'
                # 修正日の強調
                for (n_idx, d_idx) in st.session_state.modified_map.keys():
                    if n_idx < len(data) and d_idx < len(d_cls):
                        attr.at[data.index[n_idx], d_cls[d_idx]] = 'border: 2px solid #00acc1; background-color: #e0f7fa;'
                return attr

            # スタイルを適用して表示
            st.dataframe(
                st.session_state.df_result.style.applymap(style_cell).apply(style_highlight, axis=None), 
                height=600, 
                use_container_width=True
            )

            # --- 8. ルール違反警告 ---
            violations = []
            for i, row in st.session_state.df_result.iterrows():
                sl = row[d_cls].tolist()
                for di in range(len(sl)-5):
                    if all(s != "休み" for s in sl[di:di+6]):
                        violations.append(f"🚨 {row['名前']} さん：{d_cls[di]}から6連勤以上です")
            if violations:
                with st.expander("🚨 現在のルール違反状況"):
                    for v in violations: st.warning(v)

            # --- 9. 保存 ---
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as wr: st.session_state.df_result.to_excel(wr, index=False)
            st.download_button("💾 最終結果をExcelでダウンロード", data=out.getvalue(), file_name=f"kimmubyo_final.xlsx")

    except Exception as e:
        st.error(f"システムエラーが発生しました: {e}")