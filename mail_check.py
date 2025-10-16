"""
・前回の処理したデータ df_emails.tsvから前回処理日を抽出。
・前回処理日時よりも最新の 受信メール、送信メールを抽出。 （なぜか送信メールは検出できず、下書きメールから抽出） 
・件名の Re: Fw: を除去したうえで、件名ごとに最新のメールのみとする。 
・最新のメールについて、対応要否を判定する。
2025/10/12 ダミーで対応要の判定をしている。LLMで対応要否を判定するようにする必要あり。

"""


import win32com.client
from datetime import datetime, timedelta, timezone
from strands import tool
import json
import polars as pl
import re

now_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
import utils_responses

#file_path ='api_gpt-5-chat.json'              # API接続情報のファイルパス
file_path ='api_gpt5mini.json'              # API接続情報のファイルパス
api_data = utils_responses.load_api_data(file_path)     # API接続情報を取得
client, model = utils_responses.create_client(api_data) # AzureOpenAIクライアント、モデル名を取得

# Outlook定数を定義
olFolderInbox = 6    # 6は受信箱
olFolderSentMail = 5 # 5は下書きフォルダ？  3は送信完了フォルダ
olMail = 43          # 43はメールアイテムを意味する
int_days_kikan = 1  # 過去何日分のメールを対象とするか   0は当日00:00:00から、 1は1日前の00:00:00からを意味する。
pl.Config.set_tbl_rows(30)  # Dataframeの表示行数


from datetime import datetime, timedelta


# 対象外とする件名をテキストファイルから取り込む
file_path = "_対象外とする件名_.txt"
list_taishogai_kenmei = []
try:
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            # 各行の末尾にある改行文字 '\n' を除去し、リストに追加
            # strip() を使うと、前後の空白文字も除去できるので便利
            if len(line) >= 2:
                list_taishogai_kenmei.append(line.strip())
except FileNotFoundError:
    print(f"エラー: ファイル '{file_path}' が見つかりません。")
except Exception as e:
    print(f"ファイルの読み込み中にエラーが発生しました: {e}")
print(list_taishogai_kenmei)


# 対象外とする差出人をテキストファイルから取り込む
file_path = "_対象外とする差出人_.txt"
list_taishogai_sender = []
try:
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            # 各行の末尾にある改行文字 '\n' を除去し、リストに追加
            # strip() を使うと、前後の空白文字も除去できるので便利
            if len(line) >= 2:
                list_taishogai_sender.append(line.strip())
except FileNotFoundError:
    print(f"エラー: ファイル '{file_path}' が見つかりません。")
except Exception as e:
    print(f"ファイルの読み込み中にエラーが発生しました: {e}")
print(list_taishogai_sender)


def get_n_days_ago_date_string(n: int) -> str:
    """
    現在の日付からn日前の日付を 'YYYY-MM-DD 00:00:00' 形式の文字列で返します。

    Args:
        n (int): 現在の日付からの日数。

    Returns:
        str: n日前の日付を表す文字列。例: '2025-01-14 00:00:00'
    """
    # 現在の日付と時刻を取得
    current_datetime = datetime.now()

    # n日前の日付を計算
    # timedelta(days=n) を使うと、指定した日数分の期間を作成できます。
    n_days_ago = current_datetime - timedelta(days=n)

    # 日付部分を 'YYYY-MM-DD' 形式でフォーマット
    # 時刻部分は '00:00:00' で固定
    formatted_date_string = n_days_ago.strftime('%Y-%m-%d') + ' 00:00:00'

    return formatted_date_string

def clean_subject(subject):
    """件名から左端のRe:、Fw:、空白を除去"""
    if not subject:
        return ""
    
    # 左端から Re:, Fw:, Fwd:, 空白を繰り返し除去
    pattern = r'^(\s*((re|fw|fwd):\s*)*)'
    return re.sub(pattern, '', subject, flags=re.IGNORECASE).strip()

@tool
def get_recent_emails(datetime_str):
    import pythoncom
    pythoncom.CoInitialize()
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # 文字列をdatetimeオブジェクトに変換し、UTC時間として扱う
        #cutoff_date = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
        cutoff_date = datetime_str
        recent_emails = []
        
        print(f"基準日時（この日時以降のメールを取得）: {cutoff_date}")
        
        # --- 受信トレイの処理 ---
        inbox = outlook.GetDefaultFolder(olFolderInbox)  # 定数を使用
        inbox_messages = inbox.Items
        inbox_messages.Sort("[ReceivedTime]", True)
        
        for message in inbox_messages:
            try:
                if message.Class == olMail:  # メールアイテムか確認
                    received_time = message.ReceivedTime
                    print(f"受信: メール受信日時: {received_time}  件名：{message.Subject}") 
                    if received_time.strftime("%Y-%m-%d %H:%M:%S") > cutoff_date:
                        body_preview = message.Body[:1000] if message.Body else ""
                        to_recipients = message.To if hasattr(message, 'To') and message.To else ""
                        cc_recipients = message.CC if hasattr(message, 'CC') and message.CC else ""

                        email_info = {
                            "Subject_shusei": (clean_subject(message.Subject) or "").replace('\t', ''),
                            "Subject": (message.Subject or "").replace('\t', ''),
                            "Sender": message.SenderName,
                            "DateTime": received_time.strftime("%Y-%m-%d %H:%M:%S"),  # ReceivedTimeではなくDateTimeに統一
                            "BodyPreview": (body_preview or "").replace('\t', ''),
                            "To": to_recipients,
                            "CC": cc_recipients,
                            "DataType": "J"  # 受信メール
                        }
                        recent_emails.append(email_info)
                    else:
                        break  # ソート済みなのでこれ以上古いものは不要
            except Exception as e:
                print(f"受信メール処理エラー: {e} (件名: {getattr(message, 'Subject', '不明')})")
                continue
        
        # --- 送信済みアイテムの処理 ---
        # デフォルトの送信済みアイテムフォルダを取得
        try:
            sent_items_folder = outlook.GetDefaultFolder(olFolderSentMail)  # 定数を使用
            sent_messages = sent_items_folder.Items
            sent_messages.Sort("[SentOn]", True)  # SentOnでソート
            
            print(f"デフォルト送信済みアイテムフォルダからメールを取得中。メール数: {sent_messages.Count}")

            for message in sent_messages:
                try:
                    if message.Class == olMail:  # メールアイテムか確認
                        sent_time = message.SentOn
                        print(f"送信: メール送信日時: {sent_time}  件名：{message.Subject}") 
                        
                        if sent_time.strftime("%Y-%m-%d %H:%M:%S") > cutoff_date:
                            body_preview = message.Body[:1000] if message.Body else ""
                            to_recipients = message.To if hasattr(message, 'To') and message.To else ""
                            cc_recipients = message.CC if hasattr(message, 'CC') and message.CC else ""

                            email_info = {
                                "Subject_shusei": (clean_subject(message.Subject) or "").replace('\t', ''),
                                "Subject": (message.Subject or "").replace('\t', ''),
                                "Sender": message.SenderName,
                                "DateTime": sent_time.strftime("%Y-%m-%d %H:%M:%S"),  # SentOnをDateTimeに格納
                                "BodyPreview": (body_preview or "").replace('\t', ''),
                                "To": to_recipients,
                                "CC": cc_recipients,
                                "DataType": "S"  # 送信メール
                            }
                            recent_emails.append(email_info)
                        else:
                            break  # ソート済みなのでこれ以上古いものは不要
                except Exception as e:
                    print(f"送信メール処理エラー: {e} (件名: {getattr(message, 'Subject', '不明')})")
                    continue
        except Exception as e:
            print(f"送信済みアイテムフォルダへのアクセスエラー: {e}")

        # polars DataFrameに変換して重複削除処理を行う
        if recent_emails:
            df = pl.DataFrame(recent_emails)
            # DateTimeをdatetime型に変換
            df = df.with_columns(pl.col("DateTime").str.strptime(pl.Datetime, "%Y-%m-%d %H:%M:%S"))
            # Subject_shuseiごとにDateTimeが最大のものを抽出
            df = df.sort("DateTime", descending=True)
            df = df.group_by("Subject_shusei").agg(pl.col("*").first())
            df = df.sort("DateTime", descending=True)
            # DateTimeを文字列に戻す（必要に応じて）
            df = df.with_columns(pl.col("DateTime").dt.strftime("%Y-%m-%d %H:%M:%S"))
            return df
        else:
            return pl.DataFrame({
                "Subject_shusei": [],
                "Subject": [],
                "Sender": [],
                "DateTime": [],  # ReceivedTimeからDateTimeに変更
                "BodyPreview": [],
                "To": [],
                "CC": [],
                "DataType": []
            })
    finally:
        # COM終了処理
        pythoncom.CoUninitialize()


from strands.tools import tool
import pythoncom
import win32com.client
import threading


def read_outlook():
    pythoncom.CoInitialize()  # STA 初期化
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages_list = list(messages)
    subjects = [msg.Subject for msg in messages_list[:5]]
    pythoncom.CoUninitialize()
    return subjects


@tool
def get_outlook_subjects(limit: int = 5):
    """
    Outlookの受信トレイから件名を取得するツール関数。

    引数:
        limit (int): 取得する件名の最大数。デフォルトは5。

    処理:
        - スレッド内でCOMを初期化しOutlookの受信トレイからメッセージを取得。
        - メッセージの件名をlimit件数だけリストに抽出。

    戻り値:
        list: 件名の文字列リスト。
    """
    result = []
    def target():
        nonlocal result
        result = read_outlook()
    t = threading.Thread(target=target)
    t.start()
    t.join()
    return result



# if __name__ == "__main__":

def mail_check1() :

    str_datetime_start_from = get_n_days_ago_date_string(int_days_kikan)  # str_datetime_start_fromは対象期間の先頭 '2025-01-14 00:00:00' のような値
    str_datetime_from = str_datetime_start_from  # 新着メールを取得する範囲。  前回メール受信日が後の場合は後でその値に更新される。

    # 特定の日時以降のメールを取得
    try:
        df_emails_zenkai = pl.read_csv('./df_emails.tsv',separator='\t')
        #df_emails_zenkai = pl.read_excel('./df_emails.xlsx')
        str_datetime_from = df_emails_zenkai[0,"DateTime"]  # 降順ソートされているので1行目が最新の日時。前回以降の新着メールを取得する。
        print('_df_emails.tsvファイルを読み込みました。')

    except:  # 前回のTSVが存在しない場合
        str_datetime_from = get_n_days_ago_date_string(0)  #当日の00:00:00 以降の新着メールを取得する。
        pass

    # 新着メールを抽出。 
    df_emails = get_recent_emails(str_datetime_from)
    int_shinchaku =  len(df_emails)
    print(f"新着メール数: {int_shinchaku}")


    # 前回のTSVと結合
    try:
        if len(df_emails) > 0  and  len(df_emails_zenkai) > 0:
            df_emails = pl.concat([df_emails,df_emails_zenkai], how="diagonal")  # diagonalは、片方のDataFrameに不足している項目（カラム）がある場合に null をセットして結合
        elif len(df_emails) > 0 :
            pass
        elif len(df_emails_zenkai) > 0 : 
            df_emails = df_emails_zenkai
    except:  
        pass  

    # 前回 および 当日もメールがない場合は終了。 
    if len(df_emails) == 0:
        print('新着メールがありません')
        return


    # 扱う範囲は、最大でも str_datetime_start_from（対象期間の先頭）以降とする。 
    # df_emails = df_emails.with_columns(
    #     pl.col('DateTime').str.to_datetime(format="%Y-%m-%d %H:%M:%S").alias('DateTime_parsed')
    #     ).filter(
    #     pl.col('DateTime_parsed') > datetime.strptime(str_datetime_start_from, "%Y-%m-%d %H:%M:%S")
    #     )


    # 前回のTSV と 最新を合わせたものについて、同一件名ごとに最新のもののみを残す。 
    df_emails = df_emails.sort("DateTime", descending=True)
    df_emails = df_emails.group_by("Subject_shusei").agg(pl.col("*").first())
    df_emails = df_emails.sort("DateTime", descending=True)

    # 扱うメール件数は最大でも500件とする
    df_emails = df_emails[:500]

    # 初回処理の場合は、対象と判定カラムがないので追加。
    if "対象" not in df_emails.columns:
        df_emails = df_emails.with_columns( pl.lit("").alias("対象")   )
    if "判定" not in df_emails.columns:
        df_emails = df_emails.with_columns( pl.lit("").alias("判定")   )



    
    # "Subject"列に 対象外キーワード のいずれかが含まれる場合に"判定"列に「×：対象外の差出人」をセット
    search_pattern = "|".join(re.escape(keyword) for keyword in list_taishogai_sender)  # リストの要素を '|' で結合して正規表現パターンを作成  # re.escape() を使うことで、キーワードに正規表現の特殊文字が含まれていても正しく扱える
    print(f"対象外差出人の検索パターン: '{search_pattern}'")
    df_emails = df_emails.with_columns(
        pl.when(pl.col("Sender").str.contains(search_pattern))
        .then(pl.lit("×：対象外の差出人"))
        .otherwise(pl.col("判定")) # 対象外キーワードが含まれない場合は既存の"判定"列の値をそのまま使う
        .alias("判定")
    )
    
    # "Subject"列に 対象外キーワード のいずれかが含まれる場合に"判定"列に「×：対象外の件名」をセット
    search_pattern = "|".join(re.escape(keyword) for keyword in list_taishogai_kenmei)  # リストの要素を '|' で結合して正規表現パターンを作成  # re.escape() を使うことで、キーワードに正規表現の特殊文字が含まれていても正しく扱える
    print(f"対象外件名の検索パターン: '{search_pattern}'")
    df_emails = df_emails.with_columns(
        pl.when(pl.col("Subject").str.contains(search_pattern))
        .then(pl.lit("×：対象外の件名"))
        .otherwise(pl.col("判定")) # 対象外キーワードが含まれない場合は既存の"判定"列の値をそのまま使う
        .alias("判定")
    )

    # 同一件名ごとの最新メールについて、判定未実施であれば判定する。
    for row_idx, row in enumerate(df_emails.iter_rows()):
        print(f'メール処理中 ： {row_idx + 1} / {int_shinchaku}')

        # すでに前回判定済みのメールはスキップ
        hantei_text = df_emails[row_idx,'判定']
        if hantei_text is not None and  hantei_text != '' : 
            print('判定済みのためスキップします')
            continue

        # 同一件名最新メールが 送信の場合は、対象外としスキップ。
        if df_emails[row_idx,'DataType'] == 'S' :
            df_emails[row_idx,'対象'] = '×：送信'
            df_emails[row_idx,'判定'] = '×：送信'
            print('送信メールのためスキップします')
            continue

        # AIへ本文を渡して、対応要否を判定する。  
        print(f'メール処理中 ： AI判定中 : {df_emails[row_idx,'DateTime']} : {df_emails[row_idx,'Sender']} : {df_emails[row_idx,'Subject']}  ')
        body_preview_text = df_emails[row_idx,'BodyPreview']  # 本文を取得
        response = utils_responses.get_response(client, model, body_preview_text)  # AIによる、メール対応要否の判定を実行
        df_emails[row_idx,"判定"] = response
        print(f'メール判定★ ： {response}') 
        print('----------------------------------------------------------------------------------')


        # evaluation_result = ""
        # if "お願い" in body_preview_text:
        #     evaluation_result = "対応要"
        #     df_emails[row_idx,"判定"] = evaluation_result


    print("\nDataFrame情報:")
    df_hyouji = df_emails['DateTime','Sender','Subject','対象','判定']          # 表示する項目のみとする。
    #print(df_hyouji)
    #df_filtered = df_hyouji.filter(pl.col("判定").str.starts_with("○"))  # "判定" 列が '○' で始まる行のみをフィルタリング
    #df_filtered = df_hyouji.filter(    pl.col("判定").str.starts_with("○") | pl.col("判定").str.starts_with("△") )
    #df_filtered = df_hyouji.filter(    (pl.col("判定").str.starts_with("○")) & (~pl.col("対象").str.starts_with("×").fill_null(False))  )
    df_hyouji = df_hyouji[:30]  # 最新30件のみとする
    df_hyouji = df_hyouji.sort(pl.col('DateTime'))  # 日時で昇順ソート

    print(df_hyouji)
    
    # データが取得できた場合のみTSVファイル出力
    if len(df_emails) > 0:
        df_emails.write_csv('df_emails.tsv',separator='\t')
        print("df_emails.tsvファイルに出力しました")
        df_emails.write_excel('df_emails.xlsx')
        print("df_emails.xlsxファイルに出力しました")
    else:
        print("取得されたデータがないため、ファイル出力をスキップしました")

    print(f'メールチェック完了 {now_str} ') 


    # str1 = get_outlook_subjects(5)
    # print(str1)

if __name__ == "__main__":
    mail_check1()


"""
OutlookのGetDefaultFolderメソッドで使われるフォルダ番号はMAPIの標準フォルダを示しており、代表的な番号は以下の通りです。
- 3: 送信済みアイテム (Sent Items)
- 4: 削除済みアイテム (Deleted Items)
- 5: 下書き (Drafts)
- 6: 受信トレイ (Inbox)
- 9: 予定表 (Calendar)
- 10: 連絡先 (Contacts)
- 12: ジャーナル (Journal)
- 13: メモ (Notes)
- 14: タスク (Tasks)

"""
