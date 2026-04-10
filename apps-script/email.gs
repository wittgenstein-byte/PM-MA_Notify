// ============================================================
//  SEND EMAIL via Microsoft Graph API (Outlook)
// ============================================================

// ⚠️ เก็บค่าเหล่านี้ใน Apps Script > Project Settings > Script Properties
// CLIENT_ID, CLIENT_SECRET, TENANT_ID, SENDER_EMAIL

function getAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const tenantId     = props.getProperty("TENANT_ID");
  const clientId     = props.getProperty("CLIENT_ID");
  const clientSecret = props.getProperty("CLIENT_SECRET");

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const payload = {
    grant_type:    "client_credentials",
    client_id:     clientId,
    client_secret: clientSecret,
    scope:         "https://graph.microsoft.com/.default",
  };

  const response = UrlFetchApp.fetch(url, {
    method: "POST",
    payload: payload,
  });

  const json = JSON.parse(response.getContentText());
  return json.access_token;
}

// ─────────────────────────────────────────────────────────
function sendOutlookEmail(contract, daysLeft, alertDays) {
  try {
    const token = getAccessToken();
    const props  = PropertiesService.getScriptProperties();
    const sender = props.getProperty("SENDER_EMAIL");

    // รวม recipients
    const allRecipients = [
      ...contract.recipients_sale.split(","),
      ...contract.recipients_eng.split(","),
    ].map(email => ({ emailAddress: { address: email.trim() } }));

    const subject =
      `[แจ้งเตือน] ประกัน ${contract.service_type} ใกล้หมดอายุ` +
      ` — ${contract.project_name} (เหลืออีก ${daysLeft} วัน)`;

    const body = buildEmailBody(contract, daysLeft, alertDays);

    const emailPayload = {
      message: {
        subject: subject,
        body: {
          contentType: "HTML",
          content: body,
        },
        toRecipients: allRecipients,
      },
      saveToSentItems: true,
    };

    const url = `https://graph.microsoft.com/v1.0/users/${sender}/sendMail`;

    UrlFetchApp.fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(emailPayload),
    });

    return true;

  } catch (e) {
    Logger.log("Email Error: " + e.message);
    return false;
  }
}

// ─────────────────────────────────────────────────────────
function buildEmailBody(contract, daysLeft, alertDays) {
  const endDateStr = Utilities.formatDate(
    new Date(contract.end_date), "Asia/Bangkok", "dd/MM/yyyy"
  );

  return `
    <div style="font-family:Sarabun,Arial,sans-serif; max-width:600px;
                border:1px solid #e0e0e0; border-radius:8px; overflow:hidden;">

      <!-- Header -->
      <div style="background:#d32f2f; padding:20px; color:white;">
        <h2 style="margin:0;">🔔 แจ้งเตือนสัญญา MA/PM ใกล้หมดอายุ</h2>
        <p style="margin:5px 0 0;">เหลืออีก <strong>${daysLeft} วัน</strong></p>
      </div>

      <!-- Body -->
      <div style="padding:24px;">
        <table style="width:100%; border-collapse:collapse;">
          <tr style="background:#f5f5f5;">
            <td style="padding:8px; font-weight:bold; width:40%;">📁 โครงการ</td>
            <td style="padding:8px;">${contract.project_name}</td>
          </tr>
          <tr>
            <td style="padding:8px; font-weight:bold;">🏢 ลูกค้า</td>
            <td style="padding:8px;">${contract.customer_name}</td>
          </tr>
          <tr style="background:#f5f5f5;">
            <td style="padding:8px; font-weight:bold;">📄 PO Number</td>
            <td style="padding:8px;"><code>${contract.po_number}</code></td>
          </tr>
          <tr>
            <td style="padding:8px; font-weight:bold;">🛡️ ประเภทสัญญา</td>
            <td style="padding:8px;">${contract.service_type}</td>
          </tr>
          <tr style="background:#f5f5f5;">
            <td style="padding:8px; font-weight:bold;">📅 วันหมดอายุ</td>
            <td style="padding:8px; color:#d32f2f; font-weight:bold;">
              ${endDateStr}
            </td>
          </tr>
          <tr>
            <td style="padding:8px; font-weight:bold;">⏳ เหลืออีก</td>
            <td style="padding:8px; color:#d32f2f; font-weight:bold;">
              ${daysLeft} วัน
            </td>
          </tr>
        </table>

        <!-- Action Box -->
        <div style="background:#fff3e0; border-left:4px solid #ff9800;
                    padding:16px; margin-top:20px; border-radius:4px;">
          <p style="margin:0; font-weight:bold;">📌 กรุณาดำเนินการ:</p>
          <ul style="margin:8px 0 0; padding-left:20px;">
            <li>👷 <strong>Engineer:</strong> ติดต่อลูกค้าเพื่อนัดหมายต่อสัญญา</li>
            <li>💼 <strong>Sale:</strong> ติดตาม PO ต่ออายุจากลูกค้า</li>
          </ul>
        </div>
      </div>

      <!-- Footer -->
      <div style="background:#f5f5f5; padding:12px; text-align:center;
                  font-size:12px; color:#888;">
        ส่งโดยระบบ Warranty Notification อัตโนมัติ
      </div>
    </div>
  `;
}
