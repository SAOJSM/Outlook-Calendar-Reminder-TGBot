# Microsoft Calendar Telegram 提醒機器人

這個程式可以自動在每天早上 6 點透過 Telegram 發送 Microsoft Calendar 的行程提醒。所有時間均以 GMT+8（台北時間）顯示。

> 🐧 如果您使用的是 Linux 系統，請參考 [Linux 版本說明文件](README_Linux.md)。

## 創作動機

Outlook行事曆每次設定都要單獨設定是否要提醒，加上每天醒來都要打開行事曆確認今天的行程，為了解放雙手，所以打算在網路上找個現成的就好，但找了很久還是找不到，所以決定自己動手做一個，做了一個Windows版，又覺得也有些人用Linux系統，於是修正了iso時間標準格式(因Linux編碼不同)，連同Linux版本一起發布。

## 功能特點

- 自動獲取當天的所有行程
- 顯示每個事件的時間（GMT+8）、標題和地點
- 使用台北時區（GMT+8）
- 支援 HTML 格式的訊息排版
- 完整的錯誤處理和日誌記錄
- 行程按時間順序自動排序
- 支援跨平台認證和 Token 持久化
- 支援重複性活動
- 支援跨日活動
- 支援整天活動

## 修改紀錄

### 2024-01-11 功能更新
1. 改進事件顯示邏輯
   - 修正了事件顯示的時間範圍判斷
   - 確保只顯示當天的事件，避免顯示前一天或後一天的事件
   - 優化了全天事件和一般事件的過濾邏輯

2. 改進重複性活動支援
   - 修正了重複性活動無法正確顯示的問題
   - 即使活動的起始日期在當天之前，只要該活動在查詢範圍內有重複出現，都會被正確顯示

3. 改進跨日活動支援
   - 修正了跨日活動無法正確顯示的問題
   - 優化了日期範圍的處理邏輯，確保所有橫跨查詢時間範圍的活動都能被正確顯示

4. 改進時區處理
   - 優化了時區轉換的邏輯，確保所有時間都正確顯示為 GMT+8
   - 修正了時區比較的問題，確保時間比較的準確性
   - 改進了日期時間解析方式，確保在不同作業系統（Windows/Linux）上都能正確運作

5. API 查詢優化
   - 改用 calendarView 端點來查詢事件，提供更好的事件展開支援
   - 簡化了查詢邏輯，統一處理所有類型的事件
   - 查詢範圍設定為昨天到明天，確保不會遺漏任何相關事件

6. 程式碼優化
   - 簡化了事件過濾邏輯，移除冗餘的時間範圍檢查
   - 改進了錯誤處理和日誌記錄
   - 提升了程式碼的可讀性和維護性
   - 增強了跨平台相容性

## 效果展示

當機器人運行後，每天早上 6 點會發送類似以下格式的訊息：

```
📅 今日行程提醒 (2024-01-20 GMT+8)

🕒 09:00 - 10:30 (GMT+8)
📌 晨會
📍 會議室 A

🕒 14:00 - 15:00 (GMT+8)
📌 客戶會議
📍 線上會議
```

## 安裝需求

1. Python 3.7 或更高版本
2. pip（Python 套件管理器）

## 詳細設定步驟

### 1. 安裝所需套件
```bash
pip install -r requirements.txt
```

### 2. 設定 Microsoft Graph API（Azure）

1. 前往 [Azure Portal](https://portal.azure.com/)
2. 註冊新的應用程式：
   - 點擊「Azure Active Directory」→「應用程式註冊」→「新增註冊」
   - 填寫應用程式名稱
   - 選擇「僅限此組織目錄中的帳戶」
   - 點擊「註冊」

3. 設定應用程式平台：
   - 在應用程式頁面中選擇「驗證」（Authentication）
   - 點擊「新增平台」（Add a platform）
   - 選擇「桌面應用程式」（Desktop application）
   - 在重新導向 URI 中輸入 `http://localhost:53473`
   - 儲存設定

4. 取得必要資訊：
   - `CLIENT_ID`：在應用程式頁面的「概觀」中找到「應用程式（用戶端）識別碼」
   - `TENANT_ID`：在「概觀」中找到「目錄（租用戶）識別碼」

5. 建立 `CLIENT_SECRET`：
   - 前往「憑證與秘密」→「新增用戶端密碼」
   - 記下產生的密碼值（這只會顯示一次）

6. 設定權限：
   - 前往「API 權限」→「新增權限」
   - 選擇「Microsoft Graph」
   - 選擇「委派的權限」（Delegated permissions）
   - 搜尋並選擇 `Calendars.Read` 權限
   - 點擊「新增權限」
   - 點擊「代表 [組織名稱] 授與管理員同意」

### 3. 設定 Telegram Bot

1. 在 Telegram 中搜尋 [@BotFather](https://t.me/botfather)
2. 發送 `/newbot` 命令
3. 依照指示設定機器人名稱
4. 取得 `TELEGRAM_BOT_TOKEN`
5. 取得 `TELEGRAM_CHAT_ID`：
   - 搜尋 [@userinfobot](https://t.me/userinfobot)
   - 發送任何訊息給它
   - 它會回傳你的 Chat ID

### 4. 設定環境變數

1. 複製 `.env.example` 檔案並重新命名為 `.env`：
```bash
cp .env.example .env
```

2. 在 `.env` 檔案中填入剛才取得的資訊：
```
# Microsoft Graph API 設定
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
TENANT_ID=your_tenant_id

# Telegram Bot 設定
TELEGRAM_BOT_TOKEN=your_telegram_bot_token
TELEGRAM_CHAT_ID=your_chat_id
```

## 使用方法

### 首次執行前的準備

1. 確保已安裝 Python 3.7 或更高版本
2. 安裝所需套件：
```bash
pip install -r requirements.txt
```
3. 確認 `.env` 檔案已正確設定（參考上方「設定環境變數」章節）

### 首次執行

1. 執行程式：
```bash
python calendar_reminder.py
```

2. 當提示選擇授權方式時，選擇選項 1
3. 瀏覽器會自動開啟授權頁面
4. 登入 Microsoft 帳號並同意授權
5. 從重定向的 URL 複製完整網址
6. 將網址貼回程式
7. Token 會自動儲存在 `token.json` 檔案中

> 🐧 如果您使用的是 Linux 系統，請參考 [Linux 版本說明文件](README_Linux.md)

### 後續執行

- 程式會自動使用已儲存的 Token
- Token 過期時會自動更新
- 可以將 `token.json` 複製到其他環境使用

## 注意事項

1. 請確保所有環境變數都已正確設定
2. 程式需要持續運行才能發送提醒，建議使用以下方式之一來執行：
   - Windows：使用工作排程器，或使用 `python calendar_reminder.py` 在命令提示字元中執行
   - Linux：使用 screen、tmux 或設定 systemd 服務
   - 雲端：部署到雲端服務（如 Azure App Service）
3. 時區已設定為 GMT+8（台北時間），所有顯示的時間都會以此時區為準
4. 如果需要修改提醒時間，可以在程式碼中修改 `schedule.every().day.at("06:00")` 的時間設定
5. Token 檔案（`token.json`）包含敏感資訊，請妥善保管
6. 首次設定時請確保選擇「桌面應用程式」作為應用程式平台類型 