import os
import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ==========================================
# 0. 環境變數與設定
# ==========================================
load_dotenv()

CLICKUP_API_TOKEN = os.getenv("CLICKUP_API_KEY")
SPREADSHEET_ID = "14o-evT1ny3RdmEzeI5a7ianFQVTctz1kdwQuu7ueO9Y"
WEBHOOK_URL = os.getenv("WEBHOOK_URL")  # Chat 軟體的 Webhook 網址

GOOGLE_SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets'
]


def format_timestamp_to_date(timestamp):
    if not timestamp:
        return "N/A"
    return datetime.fromtimestamp(int(timestamp) / 1000).strftime('%Y-%m-%d')


def get_sheet_name_by_date(dt):
    return dt.strftime('%B %Y')


def get_current_month_range():
    return datetime.now().strftime('%Y-%m')


def get_custom_field_value(custom_fields, field_name_keywords=("品牌", "brand name")):
    """從 ClickUp 自定義欄位中尋找指定欄位，並解析下拉選單與標籤的真實文字"""
    for field in custom_fields:
        field_name = str(field.get('name', '')).lower()

        if any(kw in field_name for kw in field_name_keywords):
            value = field.get('value')
            if value is None:
                continue

            field_type = field.get('type')
            type_config = field.get('type_config', {})

            # 1. 處理下拉選單 (Drop down)
            if field_type == 'drop_down':
                options = type_config.get('options', [])
                if isinstance(value, int) and 0 <= value < len(options):
                    return str(options[value].get('name', ''))

            # 2. 處理標籤 (Labels)
            elif field_type == 'labels':
                options = type_config.get('options', [])
                val_list = value if isinstance(value, list) else [value]
                selected_labels = []
                for v_id in val_list:
                    for opt in options:
                        if opt.get('id') == v_id:
                            selected_labels.append(str(opt.get('name', opt.get('label', ''))))
                if selected_labels:
                    return ", ".join(selected_labels)

            if isinstance(value, list):
                return ", ".join([str(v) for v in value])
            return str(value)

    return "N/A"


def send_chat_message(message):
    """推播訊息到 Chat 軟體 (支援 Slack, Discord 等 Webhook)"""
    if not WEBHOOK_URL:
        print("⚠️ 提示: 未偵測到 WEBHOOK_URL，跳過 Chat 訊息推播。")
        return

    payload = {"text": message}

    if "discord.com" in WEBHOOK_URL:
        payload = {"content": message}

    try:
        response = requests.post(WEBHOOK_URL, json=payload, timeout=15)
        if response.status_code in [200, 204]:
            print("🔔 Chat 訊息推播成功！")
        else:
            print(f"⚠️ Chat 訊息推播失敗，狀態碼: {response.status_code}, 回傳內容: {response.text}")
    except requests.RequestException as e:
        print(f"❌ 發送 Chat 訊息時發生網路異常: {e}")


# ==========================================
# Google Sheets 核心邏輯
# ==========================================
def check_and_create_monthly_sheet(sheets_service):
    now = datetime.now()
    current_sheet_name = get_sheet_name_by_date(now)

    first_day_of_current_month = now.replace(day=1)
    last_day_of_last_month = first_day_of_current_month - timedelta(days=1)
    last_sheet_name = get_sheet_name_by_date(last_day_of_last_month)

    try:
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = spreadsheet.get('sheets', [])

        sheet_titles = [s.get('properties', {}).get('title') for s in sheets]
        sheet_dict = {s.get('properties', {}).get('title'): s.get('properties', {}).get('sheetId') for s in sheets}

        if current_sheet_name in sheet_titles:
            print(f"📊 當月工作表 [{current_sheet_name}] 已存在。")
            return current_sheet_name

        print(f"✨ 未偵測到 [{current_sheet_name}]，開始複製模板...")

        template_sheet_id = None
        if last_sheet_name in sheet_titles:
            template_sheet_id = sheet_dict[last_sheet_name]
        elif sheets:
            template_sheet_id = sheets[0].get('properties', {}).get('sheetId')

        if template_sheet_id is None:
            print("❌ 錯誤：找不到可供複製的模板。")
            return None

        # 1. 複製工作表
        batch_update_request = {
            'requests': [
                {
                    'duplicateSheet': {
                        'sourceSheetId': template_sheet_id,
                        'newSheetName': current_sheet_name
                    }
                }
            ]
        }
        copy_response = sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=batch_update_request
        ).execute()

        new_sheet_id = copy_response['replies'][0]['duplicateSheet']['properties']['sheetId']

        # 2. 自動移除多餘的 Filter
        try:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={'requests': [{'clearBasicFilter': {'sheetId': new_sheet_id}}]}
            ).execute()
            print("🧹 成功清除模板殘留的篩選器 (Filter)。")
        except HttpError:
            pass

        # 3. 自動重設並清除第二列以下的所有背景色塊
        try:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={
                    'requests': [
                        {
                            'repeatCell': {
                                'range': {
                                    'sheetId': new_sheet_id,
                                    'startRowIndex': 1
                                },
                                'cell': {
                                    'userEnteredFormat': {
                                        'backgroundColor': {'red': 1, 'green': 1, 'blue': 1}
                                    }
                                },
                                'fields': 'userEnteredFormat.backgroundColor'
                            }
                        }
                    ]
                }
            ).execute()
            print("🧹 成功還原資料區的所有背景色塊為無底色。")
        except HttpError:
            pass

        # 4. 清空舊資料文字
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{current_sheet_name}'!A2:Z"
        ).execute()
        print(f"🧹 已成功清空從前月複製過來的歷史舊資料。")

        print(f"🎉 成功建立乾淨的新月份工作表：[{current_sheet_name}]！")
        return current_sheet_name

    except HttpError as e:
        print(f"❌ Google API 請求發生錯誤: {e}")
        return None
    except (KeyError, IndexError) as e:
        print(f"❌ 資料解析發生錯誤: {e}")
        return None


def get_existing_urls_and_start_row(sheets_service, sheet_name):
    try:
        range_name = f"'{sheet_name}'!A:G"
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=range_name
        ).execute()
        rows = result.get('values', [])

        existing_urls = set()
        start_row = 2
        found_empty_row = False

        if not rows or len(rows) <= 1:
            return existing_urls, start_row

        header = rows[0]
        url_col_idx = 6
        for i, col_name in enumerate(header):
            if "clickup" in str(col_name).lower():
                url_col_idx = i
                break

        for i, row in enumerate(rows):
            if i == 0:
                continue

            is_placeholder = True
            if len(row) > url_col_idx:
                url = str(row[url_col_idx]).strip()
                if url and url.lower() != "clickup" and url != "n/a" and url != "":
                    existing_urls.add(url)
                    is_placeholder = False

            if is_placeholder and not found_empty_row:
                start_row = i + 1
                found_empty_row = True

        if not found_empty_row:
            start_row = len(rows) + 1

        return existing_urls, start_row

    except HttpError as e:
        print(f"⚠️ 讀取 [{sheet_name}] API 失敗: {e}")
        return set(), 2


def write_tasks_to_sheet(sheets_service, sheet_name, tasks_to_append, start_row):
    if not tasks_to_append:
        print("⏭️ 最終統計：沒有任何符合條件的新資料需要寫入。")
        send_chat_message(f"🔄 【ClickUp 同步通知】\n工作表: `{sheet_name}` 更新完畢。\n📊 本次無符合條件的新工單寫入。")
        return

    try:
        end_row = start_row + len(tasks_to_append) - 1
        range_name = f"'{sheet_name}'!A{start_row}:G{end_row}"

        body = {
            'values': tasks_to_append
        }

        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()

        success_msg = (
            f"🚀 【ClickUp 同步成功】\n"
            f"📅 執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"📊 成功更新工作表: `{sheet_name}`\n"
            f"📥 本次精準寫入: `{len(tasks_to_append)}` 筆新月份工單 (第 {start_row} ~ {end_row} 行)\n"
            f"🔗 試算表連結: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        )
        print(success_msg)
        send_chat_message(success_msg)

    except HttpError as e:
        error_msg = f"❌ 【ClickUp 同步失敗】\n在寫入工作表 `{sheet_name}` 時發生 Google Sheets API 錯誤: {e}"
        print(error_msg)
        send_chat_message(error_msg)


# ==========================================
# 1. 主流程
# ==========================================
def process_and_sync_clickup_to_sheet(list_ids):
    if not CLICKUP_API_TOKEN:
        print("❌ 錯誤: 請先在 .env 中設定 CLICKUP_API_KEY")
        return

    creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=GOOGLE_SCOPES)
    sheets_service = build('sheets', 'v4', credentials=creds)

    target_sheet_name = check_and_create_monthly_sheet(sheets_service)
    if not target_sheet_name:
        send_chat_message("❌ 【ClickUp 同步失敗】無法確認或建立目標月份工作表，程式已中斷。")
        return

    existing_urls, target_start_row = get_existing_urls_and_start_row(sheets_service, target_sheet_name)

    headers = {"Authorization": CLICKUP_API_TOKEN}
    params = {
        "archived": "false",
        "include_closed": "true",
        "subtasks": "false"
    }

    current_month = get_current_month_range()
    raw_tasks_with_timestamp = []

    for list_id in list_ids:
        url = f"https://api.clickup.com/api/v2/list/{list_id}/task"
        page = 0

        while True:
            params["page"] = str(page)
            try:
                res = requests.get(url, headers=headers, params=params, timeout=30)
                if res.status_code != 200:
                    break

                page_tasks = res.json().get('tasks', [])
                if not page_tasks:
                    break

                for task_data in page_tasks:
                    task_name = task_data.get('name', '未命名')
                    date_created_ts = task_data.get('date_created')
                    task_date_str = format_timestamp_to_date(date_created_ts)
                    task_url = task_data.get('url', '').strip()

                    if not task_date_str.startswith(current_month):
                        continue

                    if task_url in existing_urls:
                        continue

                    custom_fields = task_data.get('custom_fields', [])
                    brand_name = get_custom_field_value(custom_fields)

                    status = task_data.get('status', {}).get('status', 'N/A').upper()
                    assignees = task_data.get('assignees', [])
                    assignee_names = ", ".join([a.get('username', '') for a in assignees]) if assignees else "未指派"

                    row_data = [task_date_str, brand_name, task_name, status, "", assignee_names, task_url]

                    raw_tasks_with_timestamp.append({
                        'ts': int(date_created_ts) if date_created_ts else 0,
                        'data': row_data
                    })

                    existing_urls.add(task_url)

                page += 1
            except requests.RequestException as e:
                print(f"❌ 網路請求異常: {e}")
                break

    # 💡 修改為升冪排列 (reverse=False)
    raw_tasks_with_timestamp.sort(key=lambda x: x['ts'], reverse=False)
    tasks_to_append = [item['data'] for item in raw_tasks_with_timestamp]

    write_tasks_to_sheet(sheets_service, target_sheet_name, tasks_to_append, target_start_row)


if __name__ == "__main__":
    TARGET_LIST_IDS = [
        "901812353289", "901812355029", "901812355100", "901812355111", "901812355118", "901812355140",
        "901812411643", "901814309248", "901816628488", "901815399584", "901810486442", "901810054199"
    ]
    TARGET_LIST_IDS = list(dict.fromkeys(TARGET_LIST_IDS))

    print(f"🚀 開始執行任務轉 Google Sheet 同步程式...")
    process_and_sync_clickup_to_sheet(TARGET_LIST_IDS)