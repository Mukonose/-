import pandas as pd
import random
import datetime
from openpyxl import Workbook

# =====================
# 設定
# =====================
FILE_NAME = "history.xlsx"
NUM_RECORDS = 600  # 生成するデータ件数（600件）
YEARS_RANGE = 5    # 過去何年分作成するか（5年）

# ダミーデータの素材
EMPLOYEES = [
    ("田中課長", "tanaka@test.com"),
    ("佐藤さん", "sato@test.com"),
    ("鈴木さん", "suzuki@test.com"),
    ("高橋さん", "takahashi@test.com")
]

CLIENTS = [
    ("株式会社ABC", "03-1234-5678"),
    ("山田商事", "090-1111-2222"),
    ("テックソリューションズ", "03-9876-5432"),
    ("斎藤様", "080-3333-4444"),
    ("グローバル貿易", "045-111-2222")
]

REQUESTS = ["折り返しのお願い", "伝言のみ", "緊急対応", "見積依頼", "アポイント調整"]

# AI分析テスト用の文章パターン
MEMOS = [
    "サーバーがダウンしており、至急対応をお願いしたいとのことです。",
    "先日送付した見積書の金額について確認したいそうです。",
    "新製品「Alpha-X」のカタログを送ってほしいとの依頼。",
    "請求書がまだ届いていないので再発行をお願いします。",
    "来週の打ち合わせの日程を変更したいそうです。",
    "システムにログインできないトラブルが発生しています。",
    "担当者が不在のため、戻り次第連絡が欲しいとのこと。",
    "契約更新の手続きについて質問があります。",
    "製品の納品日が遅れている件で、少しお怒りの様子でした。",
    "素晴らしい対応ありがとうございましたとお伝えください。"
]

def generate_dummy_data():
    rows = []
    
    # 日付計算の基準
    end_date = datetime.datetime.now()
    # 5年分の日数（うるう年などは概算）
    total_days = 365 * YEARS_RANGE 
    start_date = end_date - datetime.timedelta(days=total_days)
    
    print(f"🔄 過去{YEARS_RANGE}年分 ({start_date.strftime('%Y/%m')} ～ Now) のデータを生成中...")
    print(f"📊 合計 {NUM_RECORDS} 件を作成します...")

    for _ in range(NUM_RECORDS):
        # 0日後 ～ 5年後(約1825日後) の間でランダムに日時を決定
        random_days = random.randint(0, total_days)
        random_minutes = random.randint(0, 60*24)
        
        dt = start_date + datetime.timedelta(days=random_days, minutes=random_minutes)
        dt_str = dt.strftime("%Y/%m/%d %H:%M")
        
        # ランダムな担当者選定
        to_emp = random.choice(EMPLOYEES)
        from_emp = random.choice(EMPLOYEES)
        while from_emp == to_emp: # 自分から自分への電話は避ける
            from_emp = random.choice(EMPLOYEES)
            
        client = random.choice(CLIENTS)
        req = random.choice(REQUESTS)
        memo = random.choice(MEMOS)
        
        # データ行作成
        row = {
            "日時": dt_str,
            "From": from_emp[0],
            "To": to_emp[0],
            "CC": "",
            "相手": client[0],
            "電話番号": client[1],
            "用件": req,
            "詳細": memo,
            # シート振り分け用のdatetime型（後で削除）
            "_dt_obj": dt
        }
        rows.append(row)

    # DataFrame化
    df = pd.DataFrame(rows)
    
    # 日付順にソート（古い順にしておくとシートが見やすい）
    df = df.sort_values("_dt_obj")
    
    # シート名（年月）列を作成 (例: 2021-03)
    df["sheet_name"] = df["_dt_obj"].apply(lambda x: x.strftime("%Y-%m"))
    
    # 不要な作業用列を削除
    df_save = df.drop(columns=["_dt_obj"])
    
    # Excel書き込み（シート分け）
    try:
        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
            # 月（シート名）ごとにグループ化して保存
            # groupbyのキー順（年月順）にシートが作られます
            for sheet_name, group_df in df_save.groupby("sheet_name"):
                # sheet_name列を除外して保存
                final_df = group_df.drop(columns=["sheet_name"])
                final_df.to_excel(writer, sheet_name=sheet_name, index=False)
                # 進捗表示（多すぎるので5件に1回くらい表示でもいいが、今回は全て出す）
                # print(f"✅ シート作成: {sheet_name} ({len(final_df)}件)")
            
        print(f"\n🎉 完了！ '{FILE_NAME}' に5年分のデータを保存しました。")
        print("アプリ(main.py)を再起動して、分析タブで「年」や「月」を切り替えてみてください。")
        
    except PermissionError:
        print(f"\n⚠️ エラー: '{FILE_NAME}' が開かれています。ファイルを閉じてから再実行してください。")

if __name__ == "__main__":
    generate_dummy_data()