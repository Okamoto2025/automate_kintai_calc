from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
from dotenv import load_dotenv
import os
import logging
from datetime import datetime, timedelta
from pathlib import Path

# Selenium ManagerでEdgeドライバーを自動解決
# Edgeオプションを設定
options = Options()
options.add_argument('--start-maximized')

# WebDriverを初期化
driver = webdriver.Edge(options=options)

# .envファイルからIDとパスワードを読み込む
load_dotenv()
USER_ID = os.getenv('USER_ID')
USER_PASSWORD = os.getenv('USER_PASSWORD')
LOGIN_URL = os.getenv('LOGIN_URL')
LOG_DIR = os.getenv('LOG_DIR', 'log')
DOWNLOAD_DIR = os.getenv('DOWNLOAD_DIR', str(Path.home() / 'Downloads'))
WAIT_DOWNLOAD_SEC = int(os.getenv('WAIT_DOWNLOAD_SEC', '60'))

ORG_CODES = [
    '00020B31600000',
    '00020B31100000',
    '00020B31200000',
    '00020B31300000',
    '00020B31400000',
    '00020B31500000',
    '00020B31650000',
    '00020B31700000',
    '00020B31800000',
    '00020B31850000',
    '00020B31900000',
    '00020B31910000',
]


def merge_downloaded_files_in_org_order(download_dir, org_codes):
    base = Path(download_dir)
    merged_frames = []

    for org_code in org_codes:
        org_files = sorted(
            [p for p in base.glob(f'{org_code}*') if p.is_file() and p.suffix.lower() == '.csv'],
            key=lambda p: p.stat().st_mtime
        )
        if not org_files:
            logger.warning(f'[{org_code}] 結合対象ファイルが見つかりませんでした。')
            continue

        # 同一所属で複数ファイルがある場合は最新を採用
        csv_path = org_files[-1]

        read_ok = False
        for enc in ('utf-8-sig', 'cp932', 'utf-8'):
            try:
                df = pd.read_csv(csv_path, encoding=enc, dtype=str)
                df.insert(0, 'org_code', org_code)
                merged_frames.append(df)
                read_ok = True
                break
            except Exception:
                continue

        if not read_ok:
            logger.error(f'[{org_code}] ファイル読み込みに失敗しました: {csv_path}')

    if not merged_frames:
        return None

    merged_df = pd.concat(merged_frames, ignore_index=True)

    def pick_column(columns, keywords, fallback_index=None):
        normalized = {str(c).strip().lower().replace(' ', ''): c for c in columns}
        for key in keywords:
            k = key.lower().replace(' ', '')
            for norm_col, raw_col in normalized.items():
                if k in norm_col:
                    return raw_col
        if fallback_index is not None and len(columns) > fallback_index:
            return columns[fallback_index]
        return None

    org_no_col = pick_column(merged_df.columns, ['所属番号', '所属コード', 'org_code'], fallback_index=0)
    org_name_col = pick_column(merged_df.columns, ['所属名称', '所属名', '組織名称', '組織名'], fallback_index=1)
    date_col = pick_column(merged_df.columns, ['年月日', '日付', '対象日', '勤務日'], fallback_index=6)
    over_col = pick_column(merged_df.columns, ['予実超過'], fallback_index=16)

    summary_df = merged_df[[org_no_col, org_name_col]].drop_duplicates().copy()

    work_df = merged_df.copy()
    date_series = pd.to_datetime(work_df[date_col], format='%Y%m%d', errors='coerce')
    date_series = date_series.fillna(pd.to_datetime(work_df[date_col], errors='coerce'))
    over_series = pd.to_numeric(work_df[over_col], errors='coerce').fillna(0)

    # 日付比較で境界日(7/16, 8/1 など)が漏れないよう時刻を 00:00:00 にそろえる
    today = pd.Timestamp.today().normalize()
    first_day_current_month = today.replace(day=1)
    prev_month_end = first_day_current_month - pd.Timedelta(days=1)
    prev_month_start_16 = prev_month_end.replace(day=16)
    current_month_15 = first_day_current_month.replace(day=15)
    period1_label = f'{prev_month_start_16.month}/{prev_month_start_16.day}-{prev_month_end.month}/{prev_month_end.day}'
    period2_label = f'{first_day_current_month.month}/{first_day_current_month.day}-{current_month_15.month}/{current_month_15.day}'

    date_series = date_series.dt.normalize()
    target_mask = over_series > 0
    period1_mask = (date_series >= prev_month_start_16) & (date_series <= prev_month_end)
    period2_mask = (date_series >= first_day_current_month) & (date_series <= current_month_15)
    group_cols = [org_no_col, org_name_col]

    period1_counts = (
        work_df[target_mask & period1_mask]
        .groupby(group_cols)
        .size()
        .rename(period1_label)
        .reset_index()
    )
    period2_counts = (
        work_df[target_mask & period2_mask]
        .groupby(group_cols)
        .size()
        .rename(period2_label)
        .reset_index()
    )

    summary_df = summary_df.merge(period1_counts, how='left', on=group_cols)
    summary_df = summary_df.merge(period2_counts, how='left', on=group_cols)
    summary_df[period1_label] = summary_df[period1_label].fillna(0).astype(int)
    summary_df[period2_label] = summary_df[period2_label].fillna(0).astype(int)

    summary_df = summary_df.sort_values([org_no_col, org_name_col])

    output_path = base / f'merged_daily_res_{datetime.now().strftime("%Y%m%d")}.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='結合データ', index=False)
        summary_df.to_excel(writer, sheet_name='集計結果', index=False)
    return output_path


def snapshot_daily_res_files(download_dir):
    base = Path(download_dir)
    if not base.exists():
        return set()
    return {str(p.resolve()) for p in base.glob('daily_res*') if p.is_file()}


def wait_new_daily_res_file(download_dir, before_files, timeout_sec=60):
    base = Path(download_dir)
    if not base.exists():
        return None

    end_time = time.time() + timeout_sec
    while time.time() < end_time:
        candidates = [
            p for p in base.glob('daily_res*')
            if p.is_file() and p.suffix.lower() not in {'.crdownload', '.tmp', '.part'}
        ]
        new_files = [p for p in candidates if str(p.resolve()) not in before_files]
        if new_files:
            return max(new_files, key=lambda p: p.stat().st_mtime)
        time.sleep(1)

    return None

# ログディレクトリ作成
os.makedirs(LOG_DIR, exist_ok=True)
log_filename = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y%m%d')}.log")
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

try:

    logger.info(f"勤怠処理開始")

    # 指定したURLにアクセス
    driver.get(LOGIN_URL)

    time.sleep(5)  # ページの読み込みを待機

    # iframeの存在を確認
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    if iframes:
        logger.info(f"Found {len(iframes)} iframe(s). Switching to the first iframe.")
        driver.switch_to.frame(iframes[0])  # 最初のiframeに切り替え

    # フォームの要素を操作（例: テキストボックスに入力）
    text_box = driver.find_element(By.NAME, 'id')
    text_box.send_keys(USER_ID)

    text_box = driver.find_element(By.NAME, 'password')
    text_box.send_keys(USER_PASSWORD)

    # <span>ログイン</span> の親要素である <a> タグを取得
    login_button = driver.find_element(By.CLASS_NAME, 'buttonlogin')
    # ボタンをクリック
    login_button.click()

    # 必要に応じて待機
    time.sleep(5)

    # "ワークフロー" リンクをクリック
    workflow_link = driver.find_element(By.LINK_TEXT, '勤務管理')
    workflow_link.click()

    # 必要に応じて待機
    time.sleep(2)

    # <a>タグを取得してクリック
    application_link = driver.find_element(By.LINK_TEXT, 'ファイル変換')
    application_link.click()

    logger.info(f"ログイン完了")

    # 必要に応じて待機
    time.sleep(5)


    # 前月16日〜当月15日を対象期間に設定
    first_day_current_month = datetime.today().replace(day=1)
    prev_month_end = first_day_current_month - timedelta(days=1)
    prev_month_start_16 = prev_month_end.replace(day=16)
    start_date_text = prev_month_start_16.strftime('%Y/%m/%d')
    end_date_text = first_day_current_month.replace(day=15).strftime('%Y/%m/%d')

    # 「予残超過集計」の選択と対象期間設定は最初の1回のみ
    title_no_select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'titleNo'))
    )
    Select(title_no_select).select_by_value('01')

    condition_search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.t3searchbtn.t3tabkey'))
    )
    driver.execute_script('arguments[0].click();', condition_search_button)

    time.sleep(5)

    start_date_input = WebDriverWait(driver, 10).until(
        lambda d: next((e for e in d.find_elements(By.ID, 'startDate') if e.is_displayed()), None)
    )
    end_date_input = WebDriverWait(driver, 10).until(
        lambda d: next((e for e in d.find_elements(By.ID, 'endDate') if e.is_displayed()), None)
    )

    driver.execute_script("arguments[0].removeAttribute('disabled');", start_date_input)
    driver.execute_script("arguments[0].removeAttribute('disabled');", end_date_input)

    start_date_input.clear()
    start_date_input.send_keys(start_date_text)
    start_date_input.send_keys(Keys.TAB)
    time.sleep(2)

    end_date_input.clear()
    end_date_input.send_keys(end_date_text)
    end_date_input.send_keys(Keys.TAB)
    time.sleep(2)

    for org_code in ORG_CODES:
        logger.info(f'所属コード処理開始: {org_code}')
        try:
            before_files = snapshot_daily_res_files(DOWNLOAD_DIR)

            # 所属を設定
            org_input = WebDriverWait(driver, 10).until(
                lambda d: next((e for e in d.find_elements(By.ID, 'orgId') if e.is_displayed()), None)
            )
            driver.execute_script("arguments[0].removeAttribute('disabled');", org_input)
            org_input.clear()
            org_input.send_keys(org_code)
            org_input.send_keys(Keys.TAB)
            time.sleep(2)

            execute_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.fileConvertExportBtn.t3tabkey'))
            )
            driver.execute_script('arguments[0].click();', execute_button)
            time.sleep(5)

            # 処理結果ダイアログ内の「ファイル変換」リンクをクリック（ヘッダーリンクは除外）
            try:
                dialog_file_convert_link = WebDriverWait(driver, 5).until(
                    lambda d: next(
                        (
                            e for e in d.find_elements(
                                By.XPATH,
                                "//a[normalize-space()='ファイル変換' and not(contains(@class, 't3changescreen'))]"
                            )
                            if e.is_displayed()
                        ),
                        None
                    )
                )
                driver.execute_script('arguments[0].click();', dialog_file_convert_link)
                logger.info(f'[{org_code}] 処理結果ダイアログのファイル変換リンクをクリックしました。')
            except Exception:
                logger.info(f'[{org_code}] 処理結果ダイアログのファイル変換リンクは見つかりませんでした。')

            time.sleep(2)

            # 「確認」ボタンがある場合はクリック（画面内ボタン/確認ダイアログの両方に対応）
            try:
                confirm_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='確認'] | //button[normalize-space()='確認'] | //span[normalize-space()='確認']/ancestor::a[1]"))
                )
                driver.execute_script('arguments[0].click();', confirm_button)
                logger.info(f'[{org_code}] ファイル変換遷移後に確認ボタンをクリックしました。')
            except Exception:
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present())
                    driver.switch_to.alert.accept()
                    logger.info(f'[{org_code}] ファイル変換遷移後に確認ダイアログでOKを押しました。')
                except Exception:
                    logger.info(f'[{org_code}] ファイル変換遷移後の確認操作は不要でした。')

            downloaded_file = wait_new_daily_res_file(DOWNLOAD_DIR, before_files, WAIT_DOWNLOAD_SEC)
            if downloaded_file:
                target_path = Path(DOWNLOAD_DIR) / f'{org_code}{downloaded_file.suffix}'
                if target_path.exists():
                    target_path.unlink()
                    logger.info(f'[{org_code}] 既存ファイルを削除しました: {target_path}')
                os.replace(str(downloaded_file), str(target_path))
                logger.info(f'[{org_code}] ダウンロードファイルをリネームしました: {target_path}')
            else:
                logger.warning(f'[{org_code}] ダウンロードファイルが見つかりませんでした。')

            logger.info(f'所属コード処理完了: {org_code}')
        except Exception as org_error:
            logger.error(f'所属コード {org_code} の処理でエラー: {org_error}')
            continue



    # logger.info(f"予残超過集計を実行しました: 対象期間 {start_date_text} - {end_date_text}")

    # logger.info(f"ログイン完了")

    # 必要に応じて待機
    time.sleep(5)

    # # ページのソースコードを取得
    # page_source = driver.page_source

    # # ソースコードをファイルに保存
    # with open('data/page_source.html', 'w', encoding='utf-8') as file:
    #     file.write(page_source)

    # print("ページソースを 'data/page_source.html' に保存しました。")

    # # ページのソースコードを取得
    # page_source = driver.page_source

    # # ソースコードをファイルに保存
    # with open('data/page_source.html', 'w', encoding='utf-8') as file:
    #     file.write(page_source)

    # print("ページソースを 'data/page_source.html' に保存しました。")

    # # iframeから元のコンテンツに戻る
    # driver.switch_to.default_content()

    # 必要に応じて待機
    time.sleep(5)

    merged_output_path = merge_downloaded_files_in_org_order(DOWNLOAD_DIR, ORG_CODES)
    if merged_output_path:
        logger.info(f'結合ファイルを作成しました: {merged_output_path}')
    else:
        logger.warning('結合対象ファイルがなく、結合ファイルは作成されませんでした。')

    

except Exception as e:
    logger.error(f"エラーが発生しました: {e}")
    
finally:
    # # dataが存在する場合のみExcelに保存
    # if 'data' in locals():
    #     data.to_excel(excel_file_path, sheet_name=SHEET_NAME, index=False)

    #ブラウザを閉じる
    driver.quit()

#ダウンロードしたcsvファイル群を開く(ファイル名はdaily_resで始まる)
#所属番号が00020B31で始まるもののみ抽出して結合
#指定エクセルファイルをコピーして、結合したデータを指定シートに書き込む
#集計用シートの作成処理

