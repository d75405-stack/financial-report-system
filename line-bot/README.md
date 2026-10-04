# 里辦公處 LINE 官方帳號系統

用 LINE 官方帳號管理里長工作團隊，並提供里民服務。

| 功能 | 誰用 | 怎麼用 |
|---|---|---|
| 成員綁定 | 工作人員 | 聊天室輸入 `綁定 邀請碼 姓名`，或開啟綁定頁 |
| 回報狀況 | 工作人員 | 輸入「回報」開表單（可附照片、定位），或直接輸入 `回報 內容` |
| 推播宣達 | 里長 | 管理後台發給全體工作人員、特定組別或所有好友 |
| 公告查詢 | 所有人 | 輸入「公告」（回覆訊息不計推播則數） |
| 律師諮詢 | 里民 | 輸入「律師諮詢」選時段預約；前一天 18:00 自動提醒 |
| 管理後台 | 里長 | 處理回報、發推播、管理律師時段與成員 |

新回報、新預約會自動推播通知職務為「里長」或「管理員」的成員。

## 架構

```
LINE 使用者 ─▶ LINE 官方帳號 ─webhook─▶ Google Apps Script ─▶ Google 試算表（資料）
                     │                         ▲                └▶ Google Drive（照片）
                     └─ LIFF 頁面 ─────────────┤
里長 ─▶ 管理後台（GitHub Pages）───────────────┘
```

- `apps-script/`：後端程式，貼到 Google Apps Script
- `web/liff.html`：在 LINE 裡開啟的表單（綁定、回報、律師預約、公告）
- `web/admin.html`：管理後台
- `web/config.js`：前端設定（API 網址、LIFF ID）

## 設定步驟

### 1. 建立試算表與 Apps Script

1. 用里辦的 Google 帳號新增一份 Google 試算表，例如命名為「里辦 LINE 資料」。
2. 選單 **擴充功能 → Apps Script**。
3. 把 `apps-script/Code.gs` 的內容整份貼到 `程式碼.gs`。
4. 左側 **專案設定** → 勾選「在編輯器中顯示 appsscript.json」→ 回到編輯器，把 `apps-script/appsscript.json` 貼到 `appsscript.json`。

### 2. 設定指令碼屬性

**專案設定 → 指令碼屬性**，新增：

| 屬性 | 值 | 取得位置 |
|---|---|---|
| `LINE_CHANNEL_ACCESS_TOKEN` | Channel access token | LINE Developers → Messaging API channel → Messaging API 頁籤 |
| `LIFF_ID` | 例如 `1651234567-AbCdEfGh` | 第 5 步建立 LIFF 後再填 |

### 3. 執行 setup

在編輯器上方選 `setup` 函式並按 **執行**，第一次會要求授權（試算表、Drive、外部連線、觸發器）。

執行完成後，**執行記錄**會顯示：
- `ADMIN_TOKEN`：管理後台密碼
- `INVITE_CODE`：工作人員邀請碼

這兩個值都存在指令碼屬性，之後可以自行修改。試算表會自動建立「成員、回報、公告、律師時段、諮詢預約」五個工作表。

### 4. 部署成網頁應用程式

1. **部署 → 新增部署作業 → 類型：網頁應用程式**
2. 執行身分：**我**；誰可以存取：**所有人**
3. 複製網址（結尾是 `/exec`）

> 之後修改程式碼時，請用 **部署 → 管理部署作業 → 編輯 → 版本：新版本**，網址才會維持不變。

### 5. LINE 設定

**Messaging API channel（LINE Developers）**
- Webhook URL：貼上第 4 步的 `/exec` 網址
- Use webhook：開啟
- 按「Verify」可能顯示錯誤（Apps Script 會回 302 轉址），只要實際傳訊息有反應就沒問題。

**LINE 官方帳號管理後台 → 設定 → 回應設定**
- Webhook：開啟
- 自動回應訊息：關閉（避免和機器人重複回覆）
- 聊天：開啟（機器人不認得的訊息會留給里辦人員手動回覆）

**LIFF（需要一個 LINE Login channel，和 Messaging API channel 放在同一個 Provider）**
1. LINE Login channel → LIFF 頁籤 → Add
2. Size：Full
3. Endpoint URL：`https://d75405-stack.github.io/financial-report-system/line-bot/web/liff.html`
4. Scopes：勾選 `openid`、`profile`
5. Add friend option：On (Normal)
6. 建立後複製 LIFF ID，回到 Apps Script 指令碼屬性填入 `LIFF_ID`
7. 把 LINE Login channel 設為 **Published**，否則只有開發者能登入

### 6. 前端設定

編輯 `web/config.js`，填入 Apps Script 網址和 LIFF ID，推送到 `master` 分支，GitHub Pages 會自動部署。

- 管理後台：`https://d75405-stack.github.io/financial-report-system/line-bot/web/admin.html`

### 7. 圖文選單（建議）

在 LINE 官方帳號管理後台 → **圖文選單**，建立選單，按鈕動作設為「文字」：

| 按鈕 | 傳送文字 |
|---|---|
| 📝 回報 | `回報` |
| ⚖️ 律師諮詢 | `律師諮詢` |
| 📢 最新公告 | `公告` |
| 📅 我的預約 | `我的預約` |
| ❓ 使用說明 | `說明` |

### 8. 設定里長身分

里長自己也先綁定一次，然後到管理後台 → **成員**，把職務改為「里長」，之後就會收到新回報與新預約通知。也可以在這裡設定組別（例如「第 3 鄰」「環保志工隊」），推播時就能只發給某一組。

## 費用與則數

- Google Apps Script、試算表、GitHub Pages：免費。
- LINE 官方帳號：**回覆訊息不計費**；**推播**（宣達、通知里長、預約提醒、通知回報人）依收件人數計入每月則數。免費方案的則數有限，實際額度請看官方帳號後台；管理後台總覽會顯示本月已用則數。
- 省則數的做法：一般公告請里民在圖文選單點「公告」查詢，只有緊急事項才發給所有好友。

## 安全與個資

- **管理後台**：用 `ADMIN_TOKEN` 登入；只要外流就請立刻在指令碼屬性更換。
- **LIFF 頁面**：每次呼叫都會向 LINE 驗證 ID token，無法冒用他人身分。
- **Webhook**：Apps Script 讀不到 `X-Line-Signature` 標頭，所以無法驗簽。因此會改變資料的動作（綁定、文字回報）都會先回覆訊息，回覆成功才寫入資料。reply token 只有 LINE 平台會發出，偽造的請求會回覆失敗，也就不會寫入。
- **照片**：存在里辦 Google Drive 的私人資料夾，只有里辦帳號能開啟。
- **個資**：律師諮詢會收集姓名、電話與問題內容，試算表請勿共用給不相關的人，並依個人資料保護法妥善保管與定期清除。

## 修改功能

- 回報類別、狀態、通知對象職務：`Code.gs` 開頭的 `REPORT_CATEGORIES`、`REPORT_STATUSES`、`ADMIN_ROLES`
- 關閉新回報／新預約時通知里長：指令碼屬性加 `NOTIFY_ADMINS` = `false`
- 預約提醒時間：刪掉 Apps Script 的觸發器，修改 `setup()` 裡的 `atHour(18)` 後重新執行 `setup`
