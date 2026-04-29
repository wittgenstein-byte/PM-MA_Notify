# Use Cases — PM-MA Warranty Notification System

## 1. Contract Creation (Sale)

**Actor:** Sale

**Flow:**
1. Sale กรอกข้อมูลสัญญาใน AppSheet Form
2. ระบบสร้าง contract_id และบันทึกลง `warranty_contracts`
3. ระบบสร้าง notification rules อัตโนมัติ (90, 60, 30, 7 วันก่อนหมดอายุ)
4. ระบบบันทึกลง `notification_rules`

---

## 2. Daily Expiry Check (Scheduler)

**Actor:** System (ผ่าน Apps Script Trigger)

**Flow:**
1. Trigger รันทุกวัน 08:00
2. ดึง rules ที่ `scheduled_date = วันนี้` และ `is_sent = false`
3. ส่ง notification ผ่านทุกช่องทางที่เปิดใช้งาน
4. ถ้าทุกช่องทางส่งสำเร็จ → mark `is_sent = true`
5. บันทึก log ลง `notification_logs`

---

## 3. Email Notification

**Actor:** System → Outlook Email

**Flow:**
1. ขอ Access Token จาก Microsoft Graph API (client_credentials)
2. สร้าง HTML email content
3. ส่งไปยัง recipients_sale และ recipients_eng
4. บันทึกผลลัพธ์ (`success`/`failed`)

---

## 4. Teams Notification

**Actor:** System → MS Teams

**Flow:**
1. ดึง teams_webhook URL จาก contract
2. สร้าง Adaptive Card payload
3. POST ไปยัง Incoming Webhook
4. บันทึกผลลัพธ์

---

## 5. LINE Notification

**Actor:** System → LINE

**Flow:**
1. ตรวจสอบว่า `notify_line = true` และมี `line_group_id`
2. สร้าง Flex Message JSON
3. ส่งผ่าน LINE Messaging API (Push Message)
4. บันทึกผลลัพธ์

---

## 6. Manual Rule Creation

**Actor:** Admin

**Flow:**
1. เรียก `createNotificationRules(contractId, endDate, alertDaysArray)`
2. ระบบสร้าง rules ตาม alertDaysArray ที่กำหนด

---

## 7. Log Review

**Actor:** Admin

**Flow:**
1. เปิด Google Sheets → `notification_logs`
2. ดูสถานะการส่งแต่ละช่องทาง
3. ตรวจสอบ error_msg ถ้าส่งไม่สำเร็จ