# 🔔 PM-MA Warranty Notification System

ระบบแจ้งเตือนสัญญา MA/PM อัตโนมัติ — ผ่าน Outlook Email + Microsoft Teams + LINE

## สถาปัตยกรรม

```
AppSheet (UI/Form สำหรับ Sale)
        │
        ▼
Google Sheets (warranty_contracts / notification_rules / notification_logs)
        │
        ▼
Google Apps Script (Scheduler → 3 ช่องทาง)
        │
        ├──→ 📧 Outlook Email (Microsoft Graph API)
        ├──→ 💬 MS Teams (Incoming Webhook)
        └──→ 💚 LINE (Messaging API)
```

## โครงสร้างโปรเจค

```
PM-MA_notify/
├── apps-script/          ← Google Apps Script (copy ไปวางใน Apps Script Editor)
│   ├── Code.gs           ← Main Scheduler (รันทุกวัน 08:00)
│   ├── email.gs          ← ส่ง Email ผ่าน Microsoft Graph API
│   ├── teams.gs          ← ส่ง MS Teams ผ่าน Incoming Webhook
│   ├── line.gs           ← ส่ง LINE ผ่าน Messaging API (Flex Message)
│   └── utils.gs          ← Helper Functions (ID generator, logging, rules creator)
│
├── dashboard/            ← Web Dashboard Preview (เปิดดูได้เลย)
│   ├── index.html
│   ├── style.css
│   └── app.js
│
└── README.md             ← คู่มือนี้
```

## 🚀 Quick Start

### 1. ดู Dashboard Preview

เปิดไฟล์ `dashboard/index.html` ในเบราว์เซอร์เพื่อดูตัวอย่าง UI

### 2. ตั้งค่า Google Sheets

สร้าง Google Sheets ใหม่ พร้อม 3 sheets:

| Sheet | Columns |
|-------|---------|
| `warranty_contracts` | contract_id, po_number, project_name, customer_name, service_type, start_date, end_date, recipients_sale, recipients_eng, teams_webhook, note, status, **line_group_id**, **notify_line**, created_by, created_at, updated_at |
| `notification_rules` | rule_id, contract_id, alert_days_before, scheduled_date, notify_email, notify_teams, is_sent, sent_at |
| `notification_logs` | log_id, rule_id, contract_id, channel(`email`/`teams`/`line`), status, error_msg, sent_at |

### 3. ตั้งค่า Apps Script

1. เปิด Google Sheets → Extensions → Apps Script
2. สร้าง **5 ไฟล์** แล้ว copy code จากโฟลเดอร์ `apps-script/`
3. ตั้งค่า **Script Properties** (Project Settings → Script Properties):
   - `TENANT_ID` — Azure AD Tenant ID
   - `CLIENT_ID` — Azure App Client ID
   - `CLIENT_SECRET` — Azure App Client Secret
   - `SENDER_EMAIL` — อีเมลที่ใช้ส่ง เช่น notification@company.com
   - `LINE_CHANNEL_TOKEN` — LINE Messaging API Channel Access Token

### 4. ตั้งค่า Azure App Registration

1. ไปที่ [portal.azure.com](https://portal.azure.com) → App registrations → New
2. เพิ่ม API permissions → `Mail.Send` (Application)
3. Grant admin consent
4. สร้าง Client Secret → Copy ค่า Tenant ID, Client ID, Client Secret

### 5. ตั้งค่า MS Teams Webhook

1. Teams → Channel ที่ต้องการ → Connectors
2. เพิ่ม **Incoming Webhook** → Copy URL
3. Sale นำ URL ไปกรอกในฟอร์ม AppSheet

### 6. ตั้งค่า LINE Messaging API

1. ไปที่ [developers.line.biz](https://developers.line.biz) → Create Provider → Messaging API
2. เปิด **"Allow bot to join group chats"**
3. Issue **Channel access token (long-lived)** → Copy
4. เพิ่ม Bot เข้ากลุ่ม LINE ที่ต้องการแจ้งเตือน
5. หา **Group ID** โดยตั้ง Webhook URL ชั่วคราว (เช่น webhook.site) แล้วพิมพ์อะไรก็ได้ในกลุ่ม → ดู payload: `"source": { "groupId": "Cxxxx..." }`
6. นำ Group ID ไปกรอกในฟอร์ม และ Channel Token ใส่ใน Script Properties ชื่อ `LINE_CHANNEL_TOKEN`

### 7. สร้าง AppSheet App

1. ไปที่ [appsheet.com](https://www.appsheet.com) → New App → From Google Sheets
2. เชื่อมกับ Google Sheets ที่สร้างไว้
3. ออกแบบ Form และ View ตาม dashboard preview
4. ตั้ง Automation: เมื่อบันทึกสัญญาใหม่ → เรียก `createNotificationRules()`

### 8. เปิดใช้งาน Daily Trigger

```javascript
// รันครั้งเดียวใน Apps Script Editor
setupDailyTrigger()
```

## 📊 ช่องทางแจ้งเตือน

| | 📧 Outlook Email | 💬 MS Teams | 💚 LINE |
|------|------|------|------|
| **ความยากตั้งค่า** | ⭐⭐⭐ ยาก | ⭐ ง่าย | ⭐⭐ ปานกลาง |
| **ค่าใช้จ่าย** | 🆓 ฟรี | 🆓 ฟรี | 🆓 ฟรี (200 msg/เดือน) |
| **รองรับกลุ่ม** | ✅ | ✅ | ✅ |
| **Rich Message** | ✅ HTML | ✅ Adaptive Card | ✅ Flex Message |

## ⚠️ หมายเหตุสำคัญ

- **Teams Webhook** ใช้งานได้ทันที ไม่ต้องขอ Permission พิเศษ
- **Outlook Email** ต้องขอ IT Admin ตั้งค่า Azure App Registration ก่อน
- **LINE Messaging API** Free Plan = 200 messages/เดือน — ถ้เกินต้อง upgrade
- **Script Properties** ห้ามใส่ Secret ในโค้ดโดยตรงเด็ดขาด
- Scheduler รันทุกวัน **08:00 น.** ตามเวลา timezone ของ Google Apps Script project
