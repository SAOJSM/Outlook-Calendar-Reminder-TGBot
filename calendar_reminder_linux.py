import os
import json
import datetime
import pytz
import logging
from dotenv import load_dotenv
import msal
import requests
import schedule
import time
import webbrowser
from telegram.ext import Application
from pathlib import Path

# 載入環境變數
load_dotenv()

# 設定日誌
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Microsoft Graph API 設定
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = 'common'  # 改為 'common' 以支援個人帳號
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPE = ['https://graph.microsoft.com/Calendars.Read']  # 使用委派權限
REDIRECT_URI = "http://localhost:53473"  # 本地重定向 URL

# Telegram 設定
TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
TELEGRAM_CHAT_ID = os.getenv('TELEGRAM_CHAT_ID')

# 設定時區常數
TIMEZONE = pytz.timezone('Asia/Taipei')  # GMT+8

# Token 檔案路徑
TOKEN_PATH = Path('token.json')

# 儲存訪問令牌
access_token = None
token_expires_in = None
refresh_token = None

def save_token(token_data):
    """
    將 token 資訊儲存到檔案
    """
    try:
        with open(TOKEN_PATH, 'w') as f:
            json.dump(token_data, f)
        logger.info("Token 已成功儲存")
    except Exception as e:
        logger.error(f"儲存 Token 時發生錯誤: {str(e)}")

def load_token():
    """
    從檔案載入 token 資訊
    """
    try:
        if TOKEN_PATH.exists():
            with open(TOKEN_PATH, 'r') as f:
                return json.load(f)
    except Exception as e:
        logger.error(f"載入 Token 時發生錯誤: {str(e)}")
    return None

def get_auth_app():
    """
    取得 MSAL 認證應用程式實例
    """
    return msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )

def get_auth_url():
    """
    獲取授權 URL
    """
    app = get_auth_app()
    return app.get_authorization_request_url(
        SCOPE,
        redirect_uri=REDIRECT_URI,
        state="12345"
    )

def acquire_token_by_auth_code(auth_code):
    """
    使用授權碼獲取訪問令牌
    """
    global access_token, token_expires_in, refresh_token
    app = get_auth_app()
    result = app.acquire_token_by_authorization_code(
        auth_code,
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
    if "access_token" in result:
        access_token = result['access_token']
        token_expires_in = result.get('expires_in', 3600)
        refresh_token = result.get('refresh_token')
        # 儲存 token 資訊
        save_token({
            'access_token': access_token,
            'refresh_token': refresh_token,
            'expires_in': token_expires_in,
            'timestamp': datetime.datetime.now().timestamp()
        })
        return True
    else:
        logger.error(f"獲取訪問令牌失敗: {result.get('error_description', '未知錯誤')}")
        return False

def refresh_access_token():
    """
    使用 refresh token 更新 access token
    """
    global access_token, token_expires_in, refresh_token
    try:
        app = get_auth_app()
        result = app.acquire_token_by_refresh_token(refresh_token, SCOPE)
        
        if "access_token" in result:
            access_token = result['access_token']
            token_expires_in = result.get('expires_in', 3600)
            refresh_token = result.get('refresh_token', refresh_token)
            
            # 更新儲存的 token 資訊
            save_token({
                'access_token': access_token,
                'refresh_token': refresh_token,
                'expires_in': token_expires_in,
                'timestamp': datetime.datetime.now().timestamp()
            })
            return True
        else:
            logger.error(f"更新 Token 失敗: {result.get('error_description', '未知錯誤')}")
            return False
    except Exception as e:
        logger.error(f"更新 Token 時發生錯誤: {str(e)}")
        return False

def get_access_token():
    """
    取得 Microsoft Graph API 的存取權杖
    """
    global access_token, refresh_token, token_expires_in
    
    # 嘗試載入已儲存的 token
    token_data = load_token()
    if token_data:
        logger.info("使用已存在的 token")
        access_token = token_data['access_token']
        refresh_token = token_data['refresh_token']
        token_expires_in = token_data['expires_in']
        timestamp = token_data['timestamp']
        
        # 檢查 token 是否過期（提前5分鐘更新）
        if datetime.datetime.now().timestamp() - timestamp > token_expires_in - 300:
            logger.info("Token 即將過期，嘗試更新")
            if refresh_access_token():
                logger.info("Token 更新成功")
                return access_token
            else:
                logger.error("Token 更新失敗，請在 Windows 端重新獲取 token")
                return None
        return access_token

    logger.error("找不到有效的 token，請在 Windows 端執行程式獲取 token，並將 token.json 複製到此目錄")
    return None

async def send_telegram_message(message: str):
    """
    發送 Telegram 訊息
    """
    try:
        # 在命令行輸出訊息內容
        print("\n=== 準備發送的訊息 ===")
        print(message)
        print("=====================\n")

        import telegram
        bot = telegram.Bot(token=TELEGRAM_BOT_TOKEN)
        await bot.send_message(
            chat_id=TELEGRAM_CHAT_ID,
            text=message,
            parse_mode='HTML'
        )
        logger.info("訊息發送成功")
    except Exception as e:
        logger.error(f"發送 Telegram 訊息時發生錯誤: {str(e)}")

def get_calendar_events():
    """
    取得今天的行事曆事件，以及昨天到明天的整天活動
    """
    global access_token
    try:
        access_token = get_access_token()
        if not access_token:
            return None

        # 取得今天的日期（GMT+8）
        today = datetime.datetime.now(TIMEZONE)
        
        # 設定查詢範圍：昨天到明天
        yesterday_start = (today - datetime.timedelta(days=1)).replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        tomorrow_end = (today + datetime.timedelta(days=2)).replace(
            hour=0, minute=0, second=0, microsecond=0
        )

        logger.info(f"查詢日期範圍：{yesterday_start.strftime('%Y-%m-%d')} 到 {tomorrow_end.strftime('%Y-%m-%d')}")

        headers = {
            'Authorization': f'Bearer {access_token}',
            'Prefer': 'outlook.timezone="Asia/Taipei"'
        }

        # 呼叫 Microsoft Graph API
        url = 'https://graph.microsoft.com/v1.0/me/calendar/calendarView'
        
        params = {
            '$select': 'subject,start,end,location,isAllDay,recurrence',
            '$orderby': 'start/dateTime asc',
            '$top': 50,
            'startDateTime': yesterday_start.isoformat(),
            'endDateTime': tomorrow_end.isoformat()
        }

        logger.info(f"API 請求 URL: {url}")
        logger.info(f"查詢參數: {params}")

        response = requests.get(url, headers=headers, params=params)

        logger.info(f"API 回應狀態碼: {response.status_code}")
        logger.info(f"API 回應內容: {response.text}")

        if response.status_code == 200:
            events = response.json().get('value', [])
            logger.info(f"找到 {len(events)} 個事件")
            
            # 在這裡過濾事件
            filtered_events = []
            for event in events:
                # 解析事件時間並轉換到正確的時區
                # 移除毫秒部分以確保跨平台相容性
                start_time = event['start']['dateTime'].split('.')[0]
                end_time = event['end']['dateTime'].split('.')[0]
                
                # 如果沒有時區資訊，添加 +08:00
                if not start_time.endswith('Z') and '+' not in start_time:
                    start_time += '+08:00'
                if not end_time.endswith('Z') and '+' not in end_time:
                    end_time += '+08:00'
                
                event_start = datetime.datetime.fromisoformat(start_time).astimezone(TIMEZONE)
                event_end = datetime.datetime.fromisoformat(end_time).astimezone(TIMEZONE)
                is_all_day = event.get('isAllDay', False)
                
                # 所有在查詢範圍內的事件都加入結果
                if event_start <= tomorrow_end and event_end >= yesterday_start:
                    filtered_events.append(event)
            
            logger.info(f"過濾後剩餘 {len(filtered_events)} 個事件")
            return filtered_events
            
        elif response.status_code == 401:
            # Token 過期，提示用戶重新在 Windows 端獲取
            logger.error("Token 已過期，請在 Windows 端重新獲取 token")
            return None
        else:
            logger.error(f"取得行事曆事件失敗: {response.text}")
            return None
    except Exception as e:
        logger.error(f"取得行事曆事件時發生錯誤: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None

async def send_daily_schedule():
    """
    發送每日行程提醒
    """
    try:
        logger.info("開始獲取行事曆事件...")
        events = get_calendar_events()
        
        if not events:
            logger.info("今天沒有行程安排")
            await send_telegram_message("今天沒有行程安排。")
            return

        logger.info(f"找到 {len(events)} 個行程")

        # 製作訊息
        current_date = datetime.datetime.now(TIMEZONE)
        today_start = current_date.replace(hour=0, minute=0, second=0, microsecond=0)
        today_end = today_start + datetime.timedelta(days=1)
        
        message = f"📅 <b>行程提醒 ({current_date.strftime('%Y-%m-%d')} GMT+8)</b>\n\n"

        # 將事件分類
        all_day_events = []
        regular_events = []
        
        for event in events:
            # 解析事件時間
            start_time = event['start']['dateTime'].split('.')[0]
            end_time = event['end']['dateTime'].split('.')[0]
            
            # 如果沒有時區資訊，添加 +08:00
            if not start_time.endswith('Z') and '+' not in start_time:
                start_time += '+08:00'
            if not end_time.endswith('Z') and '+' not in end_time:
                end_time += '+08:00'
            
            event_start = datetime.datetime.fromisoformat(start_time).astimezone(TIMEZONE)
            event_end = datetime.datetime.fromisoformat(end_time).astimezone(TIMEZONE)
            
            # 檢查事件是否在今天
            if event.get('isAllDay', False):
                # 全天事件：檢查事件的開始日期是否為今天
                if event_start.date() == today_start.date():
                    all_day_events.append(event)
            else:
                # 一般事件：檢查事件的時間範圍是否與今天重疊
                if event_start < today_end and event_end > today_start:
                    regular_events.append(event)

        # 先顯示整天活動
        if all_day_events:
            message += "🌟 <b>整天活動</b>\n"
            # 按開始日期排序
            all_day_events.sort(key=lambda x: x.get('start', {}).get('dateTime', ''))
            
            for event in all_day_events:
                subject = event.get('subject', '未命名事件')
                start = event.get('start', {}).get('dateTime', '').split('T')[0]
                end = event.get('end', {}).get('dateTime', '').split('T')[0]
                end_date = datetime.datetime.strptime(end, '%Y-%m-%d').date() - datetime.timedelta(days=1)
                
                if start == end_date.strftime('%Y-%m-%d'):
                    message += f"📅 {start} - {subject}\n"
                else:
                    message += f"📅 {start} 到 {end_date.strftime('%Y-%m-%d')} - {subject}\n"
                
                location = event.get('location', {}).get('displayName', '')
                if location:
                    message += f"📍 {location}\n"
                message += "\n"

        # 再顯示今天的一般活動
        if regular_events:
            message += "🕒 <b>今日活動</b>\n"
            # 按時間排序
            regular_events.sort(key=lambda x: x.get('start', {}).get('dateTime', ''))
            
            for event in regular_events:
                subject = event.get('subject', '未命名事件')
                start = event.get('start', {}).get('dateTime', '')
                end = event.get('end', {}).get('dateTime', '')
                
                if start and end:
                    # 移除毫秒部分並處理時區
                    start = start.split('.')[0]
                    end = end.split('.')[0]
                    
                    # 如果沒有時區資訊，添加 +08:00
                    if not start.endswith('Z') and '+' not in start:
                        start += '+08:00'
                    if not end.endswith('Z') and '+' not in end:
                        end += '+08:00'
                    
                    start_time = datetime.datetime.fromisoformat(start)
                    end_time = datetime.datetime.fromisoformat(end)
                    
                    # 轉換為 GMT+8 時間
                    start_time = start_time.astimezone(TIMEZONE)
                    end_time = end_time.astimezone(TIMEZONE)
                    
                    message += f"⏰ {start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}\n"
                    message += f"📌 {subject}\n"
                    
                    location = event.get('location', {}).get('displayName', '')
                    if location:
                        message += f"📍 {location}\n"
                    message += "\n"

        await send_telegram_message(message)
    except Exception as e:
        logger.error(f"發送每日行程時發生錯誤: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())

def check_and_refresh_token():
    """
    檢查並更新 token，返回下一次需要檢查的分鐘數
    如果沒有 token 或更新失敗，返回 None
    """
    try:
        token_data = load_token()
        if not token_data:
            logger.error("找不到有效的 token，請在 Windows 端執行程式獲取 token，並將 token.json 複製到此目錄")
            return None
            
        timestamp = token_data['timestamp']
        expires_in = token_data['expires_in']
        current_time = datetime.datetime.now().timestamp()
        
        # 計算剩餘時間（秒）
        remaining_time = expires_in - (current_time - timestamp)
        
        # 如果剩餘時間少於 5 分鐘，立即更新
        if remaining_time < 300:
            logger.info("定期檢查：Token 即將過期，嘗試更新")
            if refresh_access_token():
                logger.info("定期檢查：Token 更新成功")
                # 更新成功後，設定下次檢查時間為 55 分鐘後（3300 秒）
                return 55
            else:
                logger.error("定期檢查：Token 更新失敗，請在 Windows 端重新獲取 token")
                return None
        else:
            # 計算下次檢查時間：剩餘時間減去 5 分鐘的緩衝時間，再轉換為分鐘
            next_check = (remaining_time - 300) / 60
            logger.info(f"Token 仍然有效，剩餘 {remaining_time/60:.1f} 分鐘，將在 {next_check:.1f} 分鐘後檢查")
            return next_check
    except Exception as e:
        logger.error(f"檢查 Token 時發生錯誤: {str(e)}")
        return None

def schedule_next_token_check(minutes):
    """
    排程下一次的 token 檢查
    """
    # 清除所有已排程的 token 檢查
    schedule.clear('token_check')
    
    # 計算下次執行時間
    next_time = datetime.datetime.now() + datetime.timedelta(minutes=minutes)
    next_time_str = next_time.strftime("%H:%M:%S")
    
    # 排程下次檢查
    schedule.every().day.at(next_time_str).do(check_and_schedule_token).tag('token_check')
    logger.info(f"已排程下次 Token 檢查時間：{next_time_str}")

def check_and_schedule_token():
    """
    檢查 token 並排程下一次檢查
    如果檢查失敗，終止程式
    """
    minutes = check_and_refresh_token()
    if minutes is None:
        logger.error("Token 驗證失敗，程式終止")
        os._exit(1)
    schedule_next_token_check(minutes)

def main():
    """
    主程式
    """
    # 首次運行時獲取用戶授權
    if not get_access_token():
        logger.error("無法獲取用戶授權，程式終止")
        return

    # 設算 30 秒後的時間
    current_time = datetime.datetime.now()
    test_time = current_time + datetime.timedelta(seconds=30)
    test_time_str = test_time.strftime("%H:%M:%S")

    # 設定測試執行時間（30秒後）
    logger.info(f"將在 {test_time_str} 進行測試推送")
    schedule.every().day.at(test_time_str).do(lambda: asyncio.run(send_daily_schedule()))

    # 設定每天早上 6 點執行（GMT+8）
    schedule.every().day.at("06:00").do(lambda: asyncio.run(send_daily_schedule()))
    
    logger.info("行事曆提醒機器人已啟動")
    
    # 立即執行一次 token 檢查並排程下一次檢查
    check_and_schedule_token()
    
    while True:
        try:
            schedule.run_pending()
            time.sleep(1)  # 改為每秒檢查一次，以支援精確到秒的排程
        except Exception as e:
            logger.error(f"執行排程時發生錯誤: {str(e)}")
            time.sleep(1)

if __name__ == "__main__":
    import asyncio
    main() 