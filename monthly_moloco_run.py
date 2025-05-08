import pandas as pd  # pandasライブラリをインポート。データ操作に便利。
import openpyxl  # openpyxlライブラリをインポート。Excelファイルの読み書きに使用。
from datetime import datetime, timedelta  # 日付と時間を扱うためのモジュールをインポート。
import time  # 時間に関する機能を提供するモジュールをインポート。
from selenium import webdriver  # Seleniumを使ってブラウザを自動操作するためのモジュールをインポート。
from selenium.webdriver.chrome.service import Service  # ChromeDriverのサービスを管理するためのモジュールをインポート。
from webdriver_manager.chrome import ChromeDriverManager  # ChromeDriverを自動で管理するためのモジュールをインポート。
from selenium.webdriver.common.by import By  # 要素を特定するためのモジュールをインポート。
import shutil  # ファイルやディレクトリの操作を行うためのモジュールをインポート。
import os  # オペレーティングシステムとのインターフェースを提供するモジュールをインポート。
import sys  # Pythonインタープリタとのインターフェースを提供するモジュールをインポート。
import subprocess  # シェルコマンドを実行するためのモジュールをインポート。
import logging  # ログを管理するためのモジュールをインポート.



# カスタムプログレス表示関数
def show_progress(current, total, description="Processing"):
    # 標準出力が利用可能な場合のみ表示
    if sys.stdout is not None:
        progress = (current + 1) * 100 // total  # 現在の進捗をパーセントで計算
        print(f"\r{description}: {progress}% ({current + 1}/{total})", end="")  # 進捗を表示
        if current + 1 == total:  # 最後のアイテムの場合
            print()  # 完了時に改行

# 実行ファイルのディレクトリを取得
if getattr(sys, 'frozen', False):
    # 実行ファイルとして実行されている場合
    base_dir = os.path.dirname(sys.executable)  # 実行ファイルのディレクトリを取得
    # 親ディレクトリのパスを取得
    parent_dir = os.path.dirname(base_dir)
    # 実行ファイルと同じディレクトリにある設定ファイルを参照
    data_path = os.path.join(parent_dir, '海外アドネ_配信先orクライアント管理画面.xlsm')  # Excelファイルのパス
    source_path = os.path.join(os.path.expanduser('~'), 'Downloads')  # ダウンロードフォルダのパス
    destination_path = os.path.join(parent_dir, 'moloco')  # 出力先フォルダのパス
else:
    # スクリプトとして実行されている場合
    base_dir = os.path.dirname(os.path.abspath(__file__))  # スクリプトのディレクトリを取得
    # 親ディレクトリのパスを取得
    parent_dir = os.path.dirname(base_dir)
    data_path = os.path.join(parent_dir, '海外アドネ_配信先orクライアント管理画面.xlsm')  # Excelファイルのパス
    source_path = os.path.join(os.path.expanduser('~'), 'Downloads')  # ダウンロードフォルダのパス
    destination_path = os.path.join(parent_dir, 'moloco')  # 出力先フォルダのパス

# 認証情報
email = '○○'  # ログイン用のメールアドレス
password = '○○'  # ログイン用のパスワード


# 今日の日付を取得
today = datetime.now()  # 現在の日付と時間を取得
last_month = today.replace(day=1) - timedelta(days=1)  # 先月の最終日
last_month_str = last_month.strftime('%Y年%m月分')  # '2025年03月分' の形式に整形

# ログファイルの設定（YYYY年MM月取得分.log に変更）
jp_log_month = last_month.strftime('%Y年%m月')
log_name = f'{jp_log_month}取得分.log'  # 例：2025年03月取得分.log
log_dir = os.path.join(destination_path, 'log')  # logフォルダ（毎日と共通）
log_file_path = os.path.join(log_dir, log_name)  # 明示的にlogフォルダに出力



# ログディレクトリが存在しない場合は作成（destination_pathは親のmolocoフォルダ）
if not os.path.exists(destination_path):
    os.makedirs(destination_path)

logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')



def process_moloco_data():
    try:
        # Excelファイルを読み込む
        print(f"Reading Excel file from: {data_path}")  # 読み込むファイルのパスを表示
        data = pd.read_excel(data_path, sheet_name='molocoレポートURL')  # Excelファイルを読み込む

        # データの前処理
        data.columns = data.iloc[0]  # 最初の行をカラム名として設定
        data = data[1:]  # 最初の行を削除

        # 左側のデータフレームを取得（1-5列目）
        left_df = data.iloc[:, 1:5]  # 1-5列目を取得

        # 右側のデータフレームを取得（6-8列目）
        right_df = data.iloc[:, 6:8]  # 6-8列目を取得

        # NaNを含む行を抽出するデータフレームを作成
        nan_df = left_df[left_df['URL'].isna()]

        # 最初のインデックスを取得。nan_dfが空でない場合に最初のインデックスを取得
        first_index = nan_df.index[0] if not nan_df.empty else None

        # left_dfをfirst_indexで上下に二つに分ける
        if first_index is not None:  # first_indexがNoneでない場合
            # first_indexの上のデータフレームを取得
            upper_df = left_df.iloc[:first_index]
            # first_indexの下のデータフレームを取得
            lower_df = left_df.iloc[first_index:]


        # combined_dataの作成
        combined_data = upper_df[upper_df['URL'].notna()].reset_index(drop=True)  # のデータフレームをcombined_dataとして設定

        print(f"Found {len(combined_data)} valid URLs to process")  # 有効なURLの数を表示
        return combined_data  # combined_dataを返す
    except Exception as e:
        logging.error(f"Error processing Excel file: {str(e)}")  # エラーメッセージをログに記録
        return None  # エラーが発生した場合はNoneを返す

def download_reports(combined_data):
    try:
        # ダウンロードフォルダとアウトプットフォルダの作成
        if not os.path.exists(source_path):  # ダウンロードフォルダが存在しない場合
            os.makedirs(source_path)  # フォルダを作成
        if not os.path.exists(destination_path):  # 出力先フォルダが存在しない場合
            os.makedirs(destination_path)  # フォルダを作成

        # ChromeDriverのサービスを作成し、ドライバーを初期化
        print("Initializing Chrome WebDriver...")  # WebDriverの初期化を表示

        service = Service(ChromeDriverManager().install())  # ChromeDriverのサービスを作成
        driver = webdriver.Chrome(service=service)  # Chromeブラウザを起動

        total_items = len(combined_data)  # 処理するアイテムの総数を取得

        for i in range(total_items):  # 各アイテムに対してループ
            show_progress(i, total_items, "Downloading reports")  # 進捗を表示

            url = combined_data["URL"][i]  # 現在のURLを取得
            title = combined_data["Ad account"][i]  # 現在の広告アカウント名を取得

            print(f"\nProcessing {title} ({i+1}/{total_items})")  # 処理中のタイトルを表示
            logging.info(f"Processing {title}")  # 処理中のタイトルをログに記録

            driver.get(url)  # URLにアクセス
            time.sleep(5)  # ページが読み込まれるまで待機

            # ログイン処理
            if driver.find_elements(By.ID, "email") and driver.find_elements(By.ID, "password"):  # ログインフォームが存在する場合
                print("Logging in...")  # ログイン処理を表示
                Email = driver.find_element(By.ID, "email")  # メールアドレス入力フィールドを取得
                Email.send_keys(email)  # メールアドレスを入力

                Password = driver.find_element(By.ID, "password")  # パスワード入力フィールドを取得
                Password.send_keys(password)  # パスワードを入力

                time.sleep(5)  # 入力後、少し待機
                btn = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[3]/div[1]/form/button")  # ログインボタンを取得
                btn.click()  # ログインボタンをクリック

            time.sleep(10)  # ページが読み込まれるまで待機

            # 検索範囲選択
            print("Setting date range...")  # 日付範囲の設定を表示
            try:
                elem = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div/div[1]/div/div[2]/div[2]/div/form/div/div[1]/div/div[2]/div[1]/div[2]')  # 日付範囲選択フィールドを取得
            except:
                elem = driver.find_element(By.CSS_SELECTOR, '#root > div > div.sc-oTonc.jYkUxE > div > div.sc-jXPgfr.gHFPVn > div > div.sc-qQXoI.kxmigd > div.sc-qZuGl.cezain > div > form > div > div.sc-ptScb.huSRPK > div > div.sc-oVcRo.cYDwpw > div.sc-qOiPt.dGJyiw > div.sc-qWfCM.hZQJeH')  # 日付範囲選択フィールドを取得
            elem.click()  # 日付範囲選択フィールドをクリック
            time.sleep(6)  # 少し待機

            # Last Monthを選択（XPathとCSSセレクターの三段構成 + text一致補強）
            try:
                month = driver.find_element(By.XPATH, "/html/body/div[6]/div/div[1]/ul/li[9]")  # 先月を選択するオプションを取得
            except:
                try:
                    month = driver.find_element(By.XPATH, "/html/body/div[5]/div/div[1]/ul/li[9]")  # 先月を選択するオプションを取得
                except:
                    try:
                        month = driver.find_element(By.CSS_SELECTOR, "body > div:nth-child(13) > div > div.sc-pkHUE.fBZCFz > ul > li:nth-child(9)")  # 先月を選択するオプションを取得(CSSselector)
                    except:
                        # 最終手段としてtext一致探索（ロバスト化）
                        elems = driver.find_elements(By.CSS_SELECTOR, "li[role='button']")
                        found = False
                        for e in elems:
                            if e.text.strip() == "Last Month":
                                e.click()
                                found = True
                                break
                        if not found:
                            logging.error("Last Month ボタンが見つかりませんでした。")
                            print("Failed to select Last Month")
                            continue
                    else:
                        month.click()
                else:
                    month.click()
            else:
                month.click()

            time.sleep(5)  # 少し待機

            # Apply
            try:
                apply = driver.find_element(By.XPATH, "/html/body/div[6]/div/div[3]/div[2]/button[2]")  # 適用ボタンを取得
            except:
                try:
                    apply = driver.find_element(By.XPATH, "/html/body/div[5]/div/div[3]/div[2]/button[2]")  # 適用ボタンを取得(Xpath)
                except:
                    apply = driver.find_element(By.CSS_SELECTOR, "body > div:nth-child(13) > div > div.sc-qYsMX.cfFEZv > div.sc-pspzH.bqvemH > button.sc-oUDcU.juLxk")  # 適用ボタンを取得(CSSselector)
            apply.click()  # 適用ボタンをクリック
            time.sleep(5)  # 少し待機

            # Run
            print("Running report...")  # レポートを実行することを表示
            try:
                run = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div/div[1]/div/div[2]/div[2]/div/form/div/div[6]/div[2]/button[2]')  # レポート実行ボタンを取得
            except:
                run = driver.find_element(By.CSS_SELECTOR, "#root > div > div.sc-oTonc.jYkUxE > div > div.sc-jXPgfr.gHFPVn > div > div.sc-qQXoI.kxmigd > div.sc-qZuGl.cezain > div > form > div > div.sc-pTTZH.leMTtB > div.sc-pspzH.fKcJMN > button.sc-oUDcU.juLxk")  # レポート実行ボタンを取得(CSSselector)
            run.click()  # レポート実行ボタンをクリック
            time.sleep(5)  # 少し待機

            # CSVダウンロード
            print("Downloading CSV...")  # CSVをダウンロードすることを表示
            try:
                csv_button = driver.find_element(By.XPATH, "//button[@data-testid='ReportDownloadButton']")  # CSVダウンロードボタンを取得
            except:
                csv_button = driver.find_element(By.CSS_SELECTOR, "#root > div > div.sc-oTonc.jYkUxE > div > div.sc-jXPgfr.gHFPVn > div > div.sc-qQXoI.kxmigd > div.sc-qZuGl.cezain > div > div.sc-pcKrn.hoIYjT > button")  # CSVダウンロードボタンを取得(CSSselector)
            csv_button.click()  # CSVダウンロードボタンをクリック
            time.sleep(10)  # 少し待機

            # Last Month用の保存先フォルダ（例: moloco/2025年03月分）を作成
            jp_month_folder = last_month.strftime('%Y年%m月分')
            last_month_folder = os.path.join(destination_path, jp_month_folder)  # 月次フォルダを生成

            if not os.path.exists(last_month_folder):
                os.makedirs(last_month_folder)


            # CSVファイルの処理
            #以下に追加20250303
            # CSVファイルの処理
            print("Processing downloaded file...")  # ダウンロードしたファイルの処理を表示
            csv_files = [f for f in os.listdir(source_path) if f.endswith('.csv')]  # ダウンロードフォルダ内のCSVファイルをリストアップ
            csv_files.sort(key=lambda x: os.path.getmtime(os.path.join(source_path, x)), reverse=True)  # 最終更新日時でソート

            retry_count = 0  # リトライカウントを初期化
            while not csv_files and retry_count < 30:  # CSVファイルが見つからない場合、最大30秒までリトライ
                time.sleep(1)  # 1秒待機
                csv_files = [f for f in os.listdir(source_path) if f.endswith('.csv')]  # 再度CSVファイルをリストアップ
                retry_count += 1  # リトライカウントを増加

            if csv_files:  # CSVファイルが見つかった場合
                latest_file = csv_files[0]  # 最も新しいCSVファイルを取得
                latest_file_path = os.path.join(source_path, latest_file)  # ファイルのフルパスを作成

                print(f"Latest file before renaming: {latest_file_path}")  # ログを追加（デバッグ用）

                # ダウンロード中の一時ファイル (.crdownload) がある場合は待機
                while os.path.exists(latest_file_path + ".crdownload"):
                    time.sleep(1)

                # 空のCSV (0KB) の処理
                if os.stat(latest_file_path).st_size == 0:
                    empty_file_path = os.path.join(last_month_folder, f"{title}.empty.csv") # 修正: title を適用。先月分にフォルダパスに変更
                    shutil.move(latest_file_path, empty_file_path)
                    logging.info(f"Empty CSV detected (0KB): {latest_file} → Renamed to {empty_file_path}")
                    print(f"Empty CSV moved: {empty_file_path}")  # 修正: 実際の移動後のファイル名を表示
                    continue  # 次のCSVへ

                # A列のデータがすべて空かチェック
                try:
                    csv_df = pd.read_csv(latest_file_path)
                    if csv_df.empty or csv_df.iloc[:, 0].isna().all():  # A列がすべてNaNなら
                        empty_file_path = os.path.join(last_month_folder, f"{title}.empty.csv")  # 修正: title を適用。先月分にフォルダパスに変更
                        shutil.move(latest_file_path, empty_file_path)
                        logging.info(f"Empty CSV detected (header only): {latest_file} → Renamed to {empty_file_path}")
                        print(f"Empty CSV moved: {empty_file_path}")  # 修正: 実際の移動後のファイル名を表示
                        continue  # 次のCSVへ
                except pd.errors.EmptyDataError:
                    empty_file_path = os.path.join(last_month_folder, f"{title}.empty.csv") # 修正: title を適用。先月分にフォルダパスに変更
                    shutil.move(latest_file_path, empty_file_path)
                    logging.info(f"Empty CSV detected (parse error): {latest_file} → Renamed to {empty_file_path}")
                    print(f"Empty CSV moved: {empty_file_path}")  # 修正: 実際の移動後のファイル名を表示
                    continue  # 次のCSVへ


                #ここまで



                # CSVをExcelに変換
                csv_df = pd.read_csv(latest_file_path)  # CSVファイルを読み込む

                # 出力ファイル名を年月プレフィックス付きに変更
                jp_file_prefix = last_month.strftime('%Y年%m月分')
                output_file_name = f"{jp_file_prefix}_{title}.csv"
                output_file_path = os.path.join(last_month_folder, output_file_name)

                csv_df.to_csv(output_file_path, index=False, encoding='utf-8-sig')  # CSVファイルとして保存

                # 元のCSVファイルを削除
                if os.path.exists(latest_file_path):  # 元のCSVファイルが存在する場合
                    os.remove(latest_file_path)  # CSVファイルを削除
                    print(f"Converted and saved: {title}")  # 変換と保存が完了したことを表示
                    logging.info(f"Successfully converted and saved: {title}")  # 保存成功のログを追加
            else:
                print(f"Warning: No CSV file found for {title}")  # CSVファイルが見つからなかったことを警告

            time.sleep(5)  # 少し待機

        print("\nAll downloads completed successfully!")  # すべてのダウンロードが完了したことを表示
        driver.quit()  # ブラウザを閉じる

    except Exception as e:
        logging.error(f"Error during download process: {str(e)}")  # ダウンロード中のエラーメッセージをログに記録
        if 'driver' in locals():  # driverが存在する場合
            driver.quit()  # ブラウザを閉じる

def main():
    try:
        print("Starting Moloco Report Downloader...")  # プログラムの開始を表示
        print(f"Base directory: {base_dir}")  # 基本ディレクトリを表示
        print(f"Destination path: {destination_path}")  # 出力先パスを表示

        combined_data = process_moloco_data()  # データを処理
        if combined_data is not None:  # データが正常に処理された場合
            download_reports(combined_data)  # レポートをダウンロード
        else:
            print("Failed to process input data. Exiting...")  # データ処理に失敗した場合のメッセージ

        print("\nProgram completed")  # プログラムの完了を表示
        logging.info("Program completed.")  # プログラムの完了をログに記録
        #input()  # ユーザーの入力を待つ

    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")  # エラーが発生した場合のメッセージをログに記録
        #print("\nPress Enter to exit...")  # 終了のためのメッセージ
        #input()  # ユーザーの入力を待つ

if __name__ == "__main__":
    main()  # メイン関数を実行
