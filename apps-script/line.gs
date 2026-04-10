// ============================================================
//  SEND LINE MESSAGE via Messaging API
// ============================================================

function sendLineMessage(contract, daysLeft, alertDays) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty("LINE_CHANNEL_TOKEN");

    if (!token || !contract.line_group_id) return false;

    const endDateStr = Utilities.formatDate(
      new Date(contract.end_date), "Asia/Bangkok", "dd/MM/yyyy"
    );

    // กำหนด Emoji ตามความเร่งด่วน
    const urgencyEmoji = daysLeft <= 7  ? "🚨" :
                         daysLeft <= 30 ? "⚠️" : "🔔";

    const message = buildLineMessage(
      contract, daysLeft, endDateStr, urgencyEmoji
    );

    const payload = {
      to: contract.line_group_id,
      messages: [
        {
          type: "flex",
          altText: `${urgencyEmoji} แจ้งเตือน MA/PM — ${contract.project_name} เหลืออีก ${daysLeft} วัน`,
          contents: message,
        }
      ]
    };

    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`,
      },
      payload: JSON.stringify(payload),
    });

    return true;

  } catch (e) {
    Logger.log("LINE Error: " + e.message);
    return false;
  }
}

// ─────────────────────────────────────────────────────────
// LINE Flex Message Template
// ─────────────────────────────────────────────────────────
function buildLineMessage(contract, daysLeft, endDateStr, urgencyEmoji) {

  // กำหนดสีตามความเร่งด่วน
  const headerColor = daysLeft <= 7  ? "#D32F2F" :  // แดง
                      daysLeft <= 30 ? "#F57C00" :  // ส้ม
                                       "#388E3C";   // เขียว

  return {
    type: "bubble",
    header: {
      type: "box",
      layout: "vertical",
      backgroundColor: headerColor,
      contents: [
        {
          type: "text",
          text: `${urgencyEmoji} แจ้งเตือนสัญญา MA/PM`,
          color: "#FFFFFF",
          weight: "bold",
          size: "lg",
        },
        {
          type: "text",
          text: `เหลืออีก ${daysLeft} วัน`,
          color: "#FFFFFF",
          size: "sm",
          margin: "sm",
        },
      ],
    },
    body: {
      type: "box",
      layout: "vertical",
      spacing: "sm",
      contents: [
        buildLineRow("📁 โครงการ",    contract.project_name),
        buildLineRow("🏢 ลูกค้า",     contract.customer_name),
        buildLineRow("📄 PO Number",  contract.po_number),
        buildLineRow("🛡️ ประเภท",    contract.service_type),
        buildLineRow("📅 หมดอายุ",    endDateStr),
        buildLineRow("⏳ เหลืออีก",   `${daysLeft} วัน`),
        {
          type: "separator",
          margin: "md",
        },
        {
          type: "box",
          layout: "vertical",
          margin: "md",
          backgroundColor: "#FFF8E1",
          cornerRadius: "4px",
          paddingAll: "12px",
          contents: [
            {
              type: "text",
              text: "📌 กรุณาดำเนินการ",
              weight: "bold",
              size: "sm",
            },
            {
              type: "text",
              text: "👷 Engineer: นัดหมายต่อสัญญากับลูกค้า",
              size: "sm",
              margin: "sm",
              wrap: true,
            },
            {
              type: "text",
              text: "💼 Sale: ติดตาม PO ต่ออายุ",
              size: "sm",
              margin: "sm",
              wrap: true,
            },
          ],
        },
      ],
    },
    footer: {
      type: "box",
      layout: "vertical",
      backgroundColor: "#F5F5F5",
      contents: [
        {
          type: "text",
          text: "ส่งโดยระบบ Warranty Notification อัตโนมัติ",
          size: "xs",
          color: "#AAAAAA",
          align: "center",
        },
      ],
    },
  };
}

// Helper: สร้างแถวข้อมูล
function buildLineRow(label, value) {
  return {
    type: "box",
    layout: "horizontal",
    contents: [
      {
        type: "text",
        text: label,
        size: "sm",
        color: "#555555",
        flex: 3,
      },
      {
        type: "text",
        text: value || "-",
        size: "sm",
        color: "#111111",
        flex: 5,
        wrap: true,
        weight: "bold",
      },
    ],
  };
}
