// ============================================================
//  SEND MS TEAMS via Incoming Webhook
// ============================================================

function sendTeamsMessage(contract, daysLeft, alertDays) {
  try {
    const webhookUrl = contract.teams_webhook;
    if (!webhookUrl) return false;

    const endDateStr = Utilities.formatDate(
      new Date(contract.end_date), "Asia/Bangkok", "dd/MM/yyyy"
    );

    // กำหนดสีตามความเร่งด่วน
    const color = daysLeft <= 7  ? "FF0000" :   // 🔴 แดง
                  daysLeft <= 30 ? "FF9800" :   // 🟡 ส้ม
                                   "4CAF50";    // 🟢 เขียว

    const payload = {
      "@type": "MessageCard",
      "@context": "http://schema.org/extensions",
      themeColor: color,
      summary: `แจ้งเตือน MA/PM — ${contract.project_name}`,
      sections: [
        {
          activityTitle: `🔔 แจ้งเตือนสัญญา ${contract.service_type} ใกล้หมดอายุ`,
          activitySubtitle: `เหลืออีก **${daysLeft} วัน** | ครบกำหนด ${endDateStr}`,
          facts: [
            { name: "📁 โครงการ",    value: contract.project_name },
            { name: "🏢 ลูกค้า",     value: contract.customer_name },
            { name: "📄 PO Number",  value: contract.po_number },
            { name: "🛡️ ประเภท",    value: contract.service_type },
            { name: "📅 หมดอายุ",    value: endDateStr },
            { name: "⏳ เหลืออีก",   value: `${daysLeft} วัน` },
          ],
          markdown: true,
        },
        {
          title: "📌 Action Required",
          text: "👷 **Engineer** — ติดต่อลูกค้าเพื่อนัดหมายต่อสัญญา  \n" +
                "💼 **Sale** — ติดตาม PO ต่ออายุจากลูกค้า",
          markdown: true,
        },
      ],
    };

    UrlFetchApp.fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      payload: JSON.stringify(payload),
    });

    return true;

  } catch (e) {
    Logger.log("Teams Error: " + e.message);
    return false;
  }
}
