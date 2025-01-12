# Microsoft Calendar Telegram 提醒機器人 - Linux 版本

這是 Microsoft Calendar Telegram 提醒機器人的 Linux 版本。此版本設計用於在 Linux 伺服器上長期運行，並使用 Windows 版本所產生的認證資訊。

## 運作原理

Linux 版本的設計基於以下考慮：
1. Linux 伺服器通常沒有圖形界面，不便於進行 OAuth 認證
2. 需要長期穩定運行，減少人工干預
3. Token 的自動更新和錯誤處理更為重要

## 使用前準備

1. 首先在 Windows 環境下運行主程式（`calendar_reminder.py`）並完成認證
2. 確認 Windows 環境下生成了 `token.json` 文件
3. 將以下文件複製到 Linux 伺服器：
   - `calendar_reminder_linux.py`
   - `requirements.txt`
   - `.env`（包含您的設定）
   - `token.json`（從 Windows 環境複製）

## 安裝步驟

1. 安裝 Python 3.7 或更高版本：
```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install python3 python3-pip

# CentOS/RHEL
sudo yum install python3 python3-pip
```

2. 安裝所需套件：
```bash
pip3 install -r requirements.txt
```

3. 確認設定檔：
   - 檢查 `.env` 文件中的設定是否正確
   - 確認 `token.json` 已經從 Windows 環境複製過來

## 運行程式

1. 直接運行：
```bash
python3 calendar_reminder_linux.py
```

2. 使用 screen 在背景運行：
```bash
screen -S calendar
python3 calendar_reminder_linux.py
# 按 Ctrl+A 然後按 D 來分離 screen
```

3. 使用 systemd 服務（推薦）：

建立服務文件：
```bash
sudo nano /etc/systemd/system/calendar-reminder.service
```

填入以下內容：
```ini
[Unit]
Description=Microsoft Calendar Telegram Reminder Bot
After=network.target

[Service]
Type=simple
User=your_username
WorkingDirectory=/path/to/your/bot
ExecStart=/usr/bin/python3 /path/to/your/bot/calendar_reminder_linux.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

啟動服務：
```bash
sudo systemctl daemon-reload
sudo systemctl enable calendar-reminder
sudo systemctl start calendar-reminder
```

## Token 管理

Linux 版本的 token 管理機制：
1. 程式啟動時會檢查 `token.json` 是否存在
2. 如果 token 即將過期（剩餘時間少於 5 分鐘），會自動嘗試更新
3. 如果更新失敗或找不到 token，程式會終止並提示在 Windows 端重新獲取
4. Token 更新成功後，會在 55 分鐘後進行下一次檢查

## 錯誤處理

1. Token 相關錯誤：
   - 找不到 token：請確認是否已從 Windows 環境複製 `token.json`
   - Token 更新失敗：需要在 Windows 環境重新運行主程式獲取新的 token

2. 其他常見錯誤：
   - 網路連接問題：程式會自動重試
   - API 限制：已實作適當的延遲和重試機制
   - 時區問題：確保系統時區設置正確（應為 Asia/Taipei）

## 日誌查看

1. 直接運行時的日誌會輸出到終端
2. 使用 systemd 時可以用以下命令查看日誌：
```bash
journalctl -u calendar-reminder -f
```

## 注意事項

1. 此版本依賴於 Windows 版本產生的認證資訊
2. 不支援直接進行 OAuth 認證流程
3. 如果 token 更新失敗，需要在 Windows 環境重新獲取
4. 建議使用 systemd 服務來運行，以確保程式持續運行
5. 定期備份 `token.json` 文件

## 疑難排解

1. 程式無法啟動：
   - 檢查 Python 版本是否符合要求
   - 確認所有依賴套件都已安裝
   - 檢查設定檔是否正確

2. Token 相關問題：
   - 確認 `token.json` 存在且內容正確
   - 在 Windows 環境重新獲取 token
   - 檢查檔案權限是否正確

3. 時區問題：
   - 檢查系統時區設置
   - 確認 Python 時區設置正確

4. 運行權限問題：
   - 確認執行用戶具有適當權限
   - 檢查檔案和目錄的權限設置 