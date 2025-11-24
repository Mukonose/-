import streamlit as st
import pandas as pd
import datetime
import os
import smtplib
from email.mime.text import MIMEText
from email.utils import formatdate
import re
from groq import Client
import io

# PDF生成用ライブラリ
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors

# ==========================================
# ⚙️ 【重要】共有アカウント設定
# ==========================================
try:
    SHARED_EMAIL = st.secrets.get("GMAIL_ADDRESS", "")
    SHARED_PASS = st.secrets.get("GMAIL_PASSWORD", "")
    SHARED_GROQ_KEY = st.secrets.get("GROQ_API_KEY", "")
except Exception:
    SHARED_EMAIL = ""
    SHARED_PASS = ""
    SHARED_GROQ_KEY = ""
# ==========================================

# =====================
# デザイン設定（Wideモード）
# =====================
st.set_page_config(page_title="電話対応管理ツール", layout="wide", page_icon="📫")

st.markdown("""
    <div class="main-header-container">
        <h1 class="main-header-text">電話対応管理ツール</h1>
    </div>

    <style>
    .main-header-container {
        background: linear-gradient(90deg, #0052D4, #4364F7, #2E8B57);
        padding: 10px 15px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }

    .main-header-text {
        margin: 0;
        font-weight: bold;
        font-size: clamp(1.5rem, 5vw, 2.5rem); /* 幅に応じて自動調整 */
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    </style>
""", unsafe_allow_html=True)
    /* スマホ用 */
    @media (max-width: 480px) {
        .main-header h1 {
            font-size: 1.8rem;    /* 小さい画面では小さく */
        }
    }

    /* ボタンデザイン */
    div.stButton > button {
        background-color: #2E8B57;
        color: white;
        border: none;
        border-radius: 5px;
    }
    div.stButton > button:hover { 
        background-color: #3CB371; 
        color: white; 
    }

    /* AI表示ボックス */
    .ai-box {
        background-color: #e6fffa;
        border: 1px solid #2E8B57;
        padding: 15px;
        border-radius: 8px;
        margin-top: 15px;
    }

    /* エラーボックス */
    .error-box {
        background-color: #ffe6e6;
        border: 1px solid #ff0000;
        padding: 15px;
        border-radius: 8px;
        color: #cc0000;
        margin-top: 15px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# =====================
# ファイル設定（Excel）
# =====================
DATA_FILE = "history.xlsx"
EMPLOYEE_FILE = "employees.csv"

# =====================
# 関数定義
# =====================

# 1. 安全な履歴読み込み
def safe_load_history():
    cols = ["日時", "From", "To", "CC", "相手", "電話番号", "用件", "詳細"]
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=cols)
    try:
        all_sheets = pd.read_excel(DATA_FILE, sheet_name=None, engine="openpyxl")
        if not all_sheets:
            return pd.DataFrame(columns=cols)
        df_combined = pd.concat(all_sheets.values(), ignore_index=True)
        for c in cols:
            if c not in df_combined.columns: df_combined[c] = ""
        if "日時" in df_combined.columns:
            df_combined["datetime"] = pd.to_datetime(df_combined["日時"], errors='coerce')
            df_combined = df_combined.sort_values("datetime", ascending=False).drop(columns=["datetime"])
        return df_combined
    except Exception as e:
        return pd.DataFrame(columns=cols)

# 2. 履歴保存
def save_history(dt, f, t, c, caller, tel, req, memo):
    new_row = pd.DataFrame({
        "日時":[dt], "From":[f], "To":[t], "CC":[c],
        "相手":[caller], "電話番号":[tel], "用件":[req], "詳細":[memo]
    })
    
    try:
        date_obj = pd.to_datetime(dt)
        sheet_name = date_obj.strftime("%Y-%m")
    except:
        sheet_name = "Unknown"

    if not os.path.exists(DATA_FILE):
        with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
            new_row.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        try:
            existing_df = pd.read_excel(DATA_FILE, sheet_name=sheet_name, engine="openpyxl")
            updated_df = pd.concat([existing_df, new_row], ignore_index=True)
        except:
            updated_df = new_row
            
        with pd.ExcelWriter(DATA_FILE, mode='a', engine="openpyxl", if_sheet_exists='replace') as writer:
            updated_df.to_excel(writer, sheet_name=sheet_name, index=False)

# 3. 従業員管理
def load_employees():
    if os.path.exists(EMPLOYEE_FILE):
        return pd.read_csv(EMPLOYEE_FILE)
    else:
        df = pd.DataFrame({"名前":["田中課長"], "メール":["tanaka@test.com"]})
        df.to_csv(EMPLOYEE_FILE, index=False, encoding="utf-8-sig")
        return df

def save_employee(name, email):
    new_data = pd.DataFrame({"名前":[name], "メール":[email]})
    new_data.to_csv(EMPLOYEE_FILE, mode='a', header=not os.path.exists(EMPLOYEE_FILE), index=False, encoding="utf-8-sig")

def delete_employee(name_to_delete):
    df = load_employees()
    df = df[df["名前"] != name_to_delete]
    df.to_csv(EMPLOYEE_FILE, index=False, encoding="utf-8-sig")

# 4. メール送信
def send_gmail(from_mail, pw, to_mail, cc_mail, subject, body):
    if not pw:
        st.error("⚠️ メール設定（パスワード）がされていません")
        return False
    try:
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = from_mail
        msg['To'] = to_mail
        msg['Cc'] = cc_mail
        msg['Date'] = formatdate()
        recipients = [to_mail]
        if cc_mail: recipients.append(cc_mail)
        smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
        smtpobj.ehlo()
        smtpobj.starttls()
        smtpobj.login(from_mail, pw)
        smtpobj.sendmail(from_mail, recipients, msg.as_string())
        smtpobj.close()
        return True
    except Exception as e:
        st.error(f"送信エラー: {e}")
        return False

# 5. Groq AI分析
def analyze_with_groq(api_key, memo_list, period_label):
    if not api_key: return "⚠️ Groq APIキーを設定してください"
    
    all_text = "\n".join(memo_list)
    if len(all_text) > 6000:
        return "⚠️ データ量が多すぎてTOKENが足りません。\n期間を絞って再実行してください。"

    try:
        client = Client(api_key=api_key)
        prompt = f"""
        あなたはデータアナリストです。
        対象期間: 【{period_label}】
        以下の電話メモデータを分析し、日本語でレポートを作成してください。
        
        【指示】
        - 「明日」「今日」「電話」「お願いします」などの一般的な単語は分析対象から外してください。
        - 業務上の具体的な課題や、頻出する固有名詞に着目してください。

        【フォーマット】
        1. 頻出トピック (3つ)
        2. 傾向の要約 (200文字以内)
        3. 業務改善アドバイス
        
        [データ]
        {all_text}
        """
        completion = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5, max_tokens=1000
        )
        return completion.choices[0].message.content
    except Exception as e:
        err_msg = str(e)
        if "rate_limit_exceeded" in err_msg or "413" in err_msg or "429" in err_msg:
             return "⚠️ データ量が多すぎてTOKENが足りません（API制限）。\n期間を絞ってください。"
        return f"エラー: {e}"

# 6. AIキーワード抽出
def extract_keywords_ai(api_key, memo_list):
    if not api_key: return None
    all_text = "\n".join(memo_list)
    if len(all_text) > 6000:
        st.error("⚠️ データ量が多すぎてTOKENが足りません。期間を絞ってください。")
        return None

    try:
        client = Client(api_key=api_key)
        prompt = f"""
        以下の電話メモから、業務上重要な「キーワード」をトップ10抽出し、その出現回数をカウントしてください。
        【除外ルール】日時、一般的な動詞（電話、連絡、対応など）は除外。名詞を優先。
        【出力】CSV形式（ヘッダー：キーワード,回数）のみ。余計な文字禁止。
        [データ]
        {all_text}
        """
        completion = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0, max_tokens=200
        )
        content = completion.choices[0].message.content
        content = content.replace("```csv", "").replace("```", "").strip()
        clean_lines = [line.strip() for line in content.split('\n') if "," in line and len(line) < 50]
        clean_content = "\n".join(clean_lines)
        if not clean_content: return None
        df_kw = pd.read_csv(io.StringIO(clean_content), on_bad_lines='skip')
        if len(df_kw.columns) >= 2: df_kw.columns = ["キーワード", "回数"]
        return df_kw
    except Exception as e:
        err_msg = str(e)
        if "rate_limit_exceeded" in err_msg or "413" in err_msg or "429" in err_msg:
             st.error("⚠️ データ量が多すぎてTOKENが足りません（API制限）。期間を絞ってください。")
        else:
             st.error(f"AIキーワード抽出エラー: {e}")
        return None

# 7. PDF生成
def create_pdf_report(report_text, period_label, caller_df, keyword_df):
    buffer = io.BytesIO()
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    style_jp = styles["Normal"]
    style_jp.fontName = "HeiseiKakuGo-W5"
    style_jp.fontSize = 10
    style_jp.leading = 14
    style_title = styles["Title"]
    style_title.fontName = "HeiseiKakuGo-W5"
    style_h2 = styles["Heading2"]
    style_h2.fontName = "HeiseiKakuGo-W5"
    
    story = []
    story.append(Paragraph(f"電話対応分析レポート ({period_label})", style_title))
    story.append(Spacer(1, 10*mm))
    story.append(Paragraph("【AI分析サマリー】", style_h2))
    
    if "TOKENが足りません" in report_text:
        story.append(Paragraph(report_text, style_jp))
    else:
        for line in report_text.split('\n'):
            if line.strip() == "": story.append(Spacer(1, 2*mm))
            else: story.append(Paragraph(line, style_jp))
            
    story.append(Spacer(1, 10*mm))
    
    if not caller_df.empty:
        story.append(Paragraph("【相手先件数ランキング（TOP10）】", style_h2))
        story.append(Spacer(1, 3*mm))
        table_data = [['順位', '相手先名', '件数']]
        top10 = caller_df.head(10)
        for idx, (name, count) in enumerate(top10.items(), 1):
            table_data.append([str(idx), str(name), str(count)])
        t = Table(table_data, colWidths=[20*mm, 90*mm, 30*mm])
        t.setStyle(TableStyle([
            ('FONT', (0,0), (-1,-1), 'HeiseiKakuGo-W5'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ]))
        story.append(t)
        story.append(Spacer(1, 10*mm))

    if keyword_df is not None and not keyword_df.empty:
        story.append(Paragraph("【頻出キーワード（AI抽出）】", style_h2))
        story.append(Spacer(1, 3*mm))
        table_data_kw = [['キーワード', '回数']]
        for index, row in keyword_df.iterrows():
            table_data_kw.append([str(row.iloc[0]), str(row.iloc[1])])
        t_kw = Table(table_data_kw, colWidths=[90*mm, 30*mm])
        t_kw.setStyle(TableStyle([
            ('FONT', (0,0), (-1,-1), 'HeiseiKakuGo-W5'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ]))
        story.append(t_kw)

    doc.build(story)
    buffer.seek(0)
    return buffer

# =====================
# コールバック
# =====================
def fix_name_callback():
    if "input_name_val" in st.session_state:
        current_name = st.session_state.input_name_val.strip()
        honorifics = ["様", "御中", "殿", "先生", "さん"]
        if current_name and not any(current_name.endswith(h) for h in honorifics):
            st.session_state.input_name_val = current_name + "様"

# =====================
# メイン画面
# =====================
st.markdown("""
<div class="main-header">
    <h1>📫電話対応管理ツール</h1>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ 設定状況")
    if SHARED_EMAIL and SHARED_PASS:
        st.success(f"✅ 共有メール設定済み\n({SHARED_EMAIL})")
        my_email = SHARED_EMAIL
        my_pass = SHARED_PASS
    else:
        st.info("※個人設定モード")
        my_email = st.text_input("Gmail", placeholder="me@gmail.com")
        my_pass = st.text_input("アプリパスワード", type="password")
    
    st.divider()
    if SHARED_GROQ_KEY:
        st.success("✅ AI設定済み")
        groq_key = SHARED_GROQ_KEY
    else:
        groq_key = st.text_input("Groq API Key", type="password")

tab1, tab2, tab3 = st.tabs(["📝 電話入力", "👥 アドレス帳", "📊 データ分析"])

# --- TAB1: 入力 ---
with tab1:
    emp_df = load_employees()
    emp_options = ["---"] + [f"{row['名前']} : {row['メール']}" for _, row in emp_df.iterrows()]
    if "input_name_val" not in st.session_state: st.session_state.input_name_val = ""

    with st.container(border=True):
        st.subheader("新規登録")
        with st.form("input_form", clear_on_submit=False):
            # 1. From等
            c_f, c_t, c_c = st.columns(3)
            with c_f: from_sel = st.selectbox("From (受付)", emp_options)
            with c_t: to_sel = st.selectbox("To (担当)", emp_options)
            with c_c: cc_sel = st.selectbox("CC (共有)", ["---"] + [x for x in emp_options if x != "---"])
            
            st.divider()

            # 2. 相手の名前等
            c1, c2 = st.columns(2)
            with c1: in_name = st.text_input("相手の名前 / 会社名", key="input_name_val", placeholder="例：田中")
            with c2: in_tel = st.text_input("電話番号")
            
            st.divider()

            # 3. 日付等
            c_date, c_time = st.columns(2)
            with c_date:
                in_date = st.date_input("日付", datetime.datetime.now())
            with c_time:
                in_time = st.time_input("時間", datetime.datetime.now())
            
            # 4. 対応
            req_options = ["---", "伝言のみ", "折り返しのお願い", "また電話します","お問い合わせ", "その他"]
            in_req = st.selectbox("対応", req_options)

            # 5. メール件名
            in_subject = st.text_input("メール件名（空欄の場合は自動生成）", placeholder="例：【至急】田中様より 折り返しのお願い")
            
            # 6. 詳細メモ
            in_memo = st.text_area("詳細メモ", height=100)

            # 7. 送信ボタン
            submitted = st.form_submit_button("送信＆保存", on_click=fix_name_callback)
            
            if submitted:
                if from_sel == "---" or to_sel == "---":
                    st.error("⚠️ From と To を選択してください")
                elif in_req == "---":
                    st.error("⚠️ 用件を選択してください")
                elif not in_name:
                    st.warning("⚠️ 相手の名前を入力してください")
                else:
                    final_name = st.session_state.input_name_val
                    
                    # 日付と時間を結合
                    dt_obj = datetime.datetime.combine(in_date, in_time)
                    input_dt_str = dt_obj.strftime("%Y/%m/%d %H:%M")
                    
                    f_val = from_sel.split(" : ")
                    t_val = to_sel.split(" : ")
                    f_mail, f_name = f_val[1], f_val[0]
                    t_mail, t_name = t_val[1], t_val[0]
                    c_mail, c_name = "", ""
                    if cc_sel != "---":
                        c_val = cc_sel.split(" : ")
                        c_mail, c_name = c_val[1], c_val[0]
                    
                    save_history(input_dt_str, f_name, t_name, c_name, final_name, in_tel, in_req, in_memo)
                    
                    if in_subject.strip(): subject = in_subject
                    else: subject = f"【電話】{final_name}"
                    
                    body = f"{t_name}さん\n\n電話がありました。\n日時: {input_dt_str}\n相手: {final_name} ({in_tel})\n用件: {in_req}\n\n詳細:\n{in_memo}"
                    
                    if send_gmail(my_email, my_pass, t_mail, c_mail, subject, body):
                        st.success(f"✅ 送信完了！ 日時：{input_dt_str} で登録しました。")
                    else:
                        st.success(f"✅ 保存完了！ 日時：{input_dt_str} で記録しました。（メールは未送信）")

# --- TAB2: アドレス帳 ---
with tab2:
    st.subheader("従業員リスト管理")
    with st.expander("➕ 新規追加", expanded=True):
        c1, c2 = st.columns(2)
        with c1: n_name = st.text_input("名前")
        with c2: n_mail = st.text_input("メール")
        if st.button("追加"):
            if n_name and n_mail:
                save_employee(n_name, n_mail)
                st.success("追加しました")
                st.rerun()
    st.divider()
    curr_df = load_employees()
    if not curr_df.empty:
        del_target = st.selectbox("削除する従業員を選択", ["---"] + curr_df["名前"].tolist())
        if st.button("削除実行"):
            if del_target != "---":
                delete_employee(del_target)
                st.warning(f"{del_target} さんを削除しました")
                st.rerun()
    st.dataframe(load_employees(), use_container_width=True)

# --- TAB3: データ分析 ---
with tab3:
    st.subheader("分析レポート")
    
    if "ai_keywords_df" not in st.session_state:
        st.session_state["ai_keywords_df"] = None
    if "report_text" not in st.session_state:
        st.session_state["report_text"] = ""

    df = safe_load_history()
    
    if len(df) == 0:
        st.info("データがありません")
    else:
        df["datetime"] = pd.to_datetime(df["日時"], errors='coerce')
        df = df.dropna(subset=["datetime"])
        df["year"] = df["datetime"].dt.year
        df["month"] = df["datetime"].dt.month
        
        # === フィルター選択部分 ===
        years = sorted(df["year"].astype(int).unique(), reverse=True)
        if not years:
            st.warning("データなし")
        else:
            c_y, c_m = st.columns(2)
            
            # 1. 年の選択
            year_options = ["---"] + list(years)
            sel_year = st.selectbox("対象年", year_options)
            
            # 2. 月の選択
            if sel_year == "---":
                sel_month = "---"
                st.selectbox("対象月", ["--- (全期間)"], disabled=True)
            else:
                months = sorted(df[df["year"] == sel_year]["month"].astype(int).unique())
                month_options = ["---"] + months
                sel_month = st.selectbox("対象月", month_options)

            # 3. データの絞り込みとラベル作成
            if sel_year == "---":
                df_sub = df
                period_label = "全期間"
            elif sel_month == "---":
                df_sub = df[df["year"] == sel_year]
                period_label = f"{sel_year}年 年間"
            else:
                df_sub = df[(df["year"] == sel_year) & (df["month"] == sel_month)]
                period_label = f"{sel_year}年 {sel_month}月"
            
            # === 結果表示 ===
            if len(df_sub) > 0:
                st.success(f"【{period_label}】のデータ: {len(df_sub)}件")
                
                c_left, c_right = st.columns([1, 1])
                with c_left:
                    st.markdown("### 📞 相手先TOP10")
                    caller_counts = df_sub["相手"].value_counts().head(10)
                    st.bar_chart(caller_counts, horizontal=True)
                    rank_df = caller_counts.reset_index()
                    rank_df.columns = ["相手先", "回数"]
                    st.dataframe(rank_df, use_container_width=True, hide_index=True)

                with c_right:
                    st.markdown("### 🔑 AIキーワード")
                    memos = df_sub["詳細"].dropna().astype(str).tolist()
                    
                    if groq_key:
                        if st.button("🤖 AI集計を実行"):
                            with st.spinner(f"集計中..."):
                                kw_df = extract_keywords_ai(groq_key, memos)
                                st.session_state["ai_keywords_df"] = kw_df
                        
                        if st.session_state["ai_keywords_df"] is not None:
                            kw_df = st.session_state["ai_keywords_df"]
                            chart_data = kw_df.set_index("キーワード")
                            st.bar_chart(chart_data["回数"], horizontal=True)
                            st.dataframe(kw_df, use_container_width=True, hide_index=True)
                    else:
                        st.warning("APIキー未設定")

                st.divider()
                
                st.markdown(f"### ⚡ AI総合レポート ({period_label})")
                if st.button("🤖 総合レポート生成"):
                    if groq_key:
                        with st.spinner(f"執筆中..."):
                            memos = df_sub["詳細"].dropna().tolist()
                            report = analyze_with_groq(groq_key, memos, period_label)
                            st.session_state["report_text"] = report
                    else:
                        st.error("APIキー未設定")
                
                if st.session_state["report_text"]:
                    if "TOKENが足りません" in st.session_state["report_text"]:
                         st.markdown(f'<div class="error-box">{st.session_state["report_text"]}</div>', unsafe_allow_html=True)
                    else:
                         st.markdown(f'<div class="ai-box">{st.session_state["report_text"]}</div>', unsafe_allow_html=True)
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            "📄 テキスト保存", 
                            st.session_state["report_text"], 
                            file_name=f"report_{period_label.replace(' ', '_')}.txt"
                        )
                    with c2:
                        caller_series = df_sub["相手"].value_counts()
                        keyword_data = st.session_state.get("ai_keywords_df", None)
                        
                        pdf_file = create_pdf_report(
                            st.session_state["report_text"], 
                            period_label, 
                            caller_series, keyword_data
                        )
                        st.download_button(
                            "📄 PDF保存", 
                            pdf_file, 
                            file_name=f"report_{period_label.replace(' ', '_')}.pdf", 
                            mime="application/pdf"
                        )
            else:

                st.warning("この期間のデータはありません")








