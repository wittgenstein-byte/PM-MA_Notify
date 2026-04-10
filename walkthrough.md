# 🔔 PM-MA Warranty Notification System — Walkthrough

## สิ่งที่สร้างให้

### 1. Google Apps Script (พร้อม copy ไปใช้)

| ไฟล์ | หน้าที่ |
|------|---------|
| [Code.gs](file:///c:/Users/Admin/PM-MA_notify/apps-script/Code.gs) | Main Scheduler — ตรวจและส่งแจ้งเตือนทุกวัน 08:00 |
| [email.gs](file:///c:/Users/Admin/PM-MA_notify/apps-script/email.gs) | ส่ง Outlook Email ผ่าน Microsoft Graph API (OAuth2) |
| [teams.gs](file:///c:/Users/Admin/PM-MA_notify/apps-script/teams.gs) | ส่ง MS Teams ผ่าน Incoming Webhook (สีตามความเร่งด่วน) |
| [utils.gs](file:///c:/Users/Admin/PM-MA_notify/apps-script/utils.gs) | Helper: ID generator, logging, auto-create rules, trigger setup |

---

### 2. Web Dashboard Preview (เปิดได้เลยใน browser)

````carousel
![Dashboard — หน้าสรุปสัญญาทั้งหมด พร้อมสถิติ](C:\Users\Admin\.gemini\antigravity\brain\7a4d64fc-5721-4662-93e0-876116bc12df\.system_generated\click_feedback\click_feedback_1775806846969.png)
<!-- slide -->
![Contracts List — ตารางสัญญาพร้อม filter และ search](C:\Users\Admin\.gemini\antigravity\brain\7a4d64fc-5721-4662-93e0-876116bc12df\.system_generated\click_feedback\click_feedback_1775806859383.png)
````

**ไฟล์ Dashboard:**
- [index.html](file:///c:/Users/Admin/PM-MA_notify/dashboard/index.html) — โครงสร้าง HTML
- [style.css](file:///c:/Users/Admin/PM-MA_notify/dashboard/style.css) — Dark theme + glassmorphism CSS
- [app.js](file:///c:/Users/Admin/PM-MA_notify/dashboard/app.js) — Logic ทั้งหมด (navigation, form, table, modal)

**ฟีเจอร์ Dashboard:**
- 📊 **Dashboard** — สรุป 4 สถิติ (วิกฤต/ปานกลาง/ปกติ/ทั้งหมด) + ตารางใกล้หมดอายุ + feed แจ้งเตือนล่าสุด
- 📄 **สัญญาทั้งหมด** — ตารางพร้อม filter (ทั้งหมด/ใกล้หมดอายุ/ปานกลาง/ปกติ/หมดอายุ) + ค้นหา
- ➕ **เพิ่มสัญญาใหม่** — ฟอร์มกรอกข้อมูลครบ (PO, โครงการ, ลูกค้า, ประเภท, ระยะเวลา, แจ้งเตือน 90/60/30/7 วัน, ผู้รับ, webhook)
- 📊 **ประวัติแจ้งเตือน** — Log การส่งทั้งหมด (email/teams, success/failed)
- ⚙️ **ตั้งค่า** — Azure config, Trigger status, Setup checklist

---

### 3. เอกสาร

- [README.md](file:///c:/Users/Admin/PM-MA_notify/README.md) — คู่มือติดตั้ง 7 ขั้นตอน (Google Sheets → Apps Script → Azure → Teams → AppSheet → Trigger)

---

## ขั้นตอนถัดไป

1. **สร้าง Google Sheets** — สร้าง 3 sheets ตาม schema แล้ว copy header row
2. **วาง Apps Script** — Extensions → Apps Script → สร้างไฟล์ 4 ไฟล์ แล้ว paste code
3. **ตั้งค่า Script Properties** — TENANT_ID, CLIENT_ID, CLIENT_SECRET, SENDER_EMAIL
4. **สร้าง Azure App** — App Registration + Mail.Send permission
5. **สร้าง Teams Webhook** — ใน channel ที่ต้องการ
6. **สร้าง AppSheet App** — เชื่อม Google Sheets + ออกแบบ Form/View
7. **รัน `setupDailyTrigger()`** — เปิดใช้งาน scheduler
