/**
 * HUE SOUND EXPRESS - BO NHAN DON HANG & WAITLIST MIEN PHI (0D)
 * Huong dan cai dat (3 phut):
 * 1. Tao 1 Google Sheet moi co 2 tab: "Orders" va "Waitlist"
 * 2. Vao menu: Tien ich mo rong (Extensions) -> Apps Script
 * 3. Dan toan bo ma nay vao, dien TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID va NOTIFY_EMAIL
 * 4. Bam Deploy -> New deployment -> Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Copy URL Web App dan vao bien WEBHOOK_URL trong file index.html
 */

const CONFIG = {
  TELEGRAM_BOT_TOKEN: "DIEN_BOT_TOKEN_TU_BOTFATHER", // Tuy chon: bao ting-ting 1 giay qua Telegram
  TELEGRAM_CHAT_ID: "DIEN_CHAT_ID_CUA_BAN",
  NOTIFY_EMAIL: "email-cua-ban@gmail.com",           // Gui email ngay khi co don / dang ky waitlist
  SHEET_ORDERS: "Orders",
  SHEET_WAITLIST: "Waitlist"
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss");

    if (data.type === "waitlist") {
      let sheet = ss.getSheetByName(CONFIG.SHEET_WAITLIST);
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG.SHEET_WAITLIST);
        sheet.appendRow(["Thoi gian", "Ten khach", "SDT/Zalo", "SKU", "Ten san pham", "Ghi chu"]);
      }
      sheet.appendRow([now, data.name || "", data.phone || "", data.sku || "", data.product || "", data.note || ""]);

      const msg =
        "🔔 [WAITLIST] HUE SOUND EXPRESS\n" +
        "• Khach: " + (data.name || "Khong ro") + "\n" +
        "• SDT/Zalo: " + (data.phone || "") + "\n" +
        "• Quan tam: " + (data.product || "") + " (" + (data.sku || "") + ")\n" +
        "• Luc: " + now;

      sendNotifications("🔔 [Waitlist] " + (data.product || ""), msg);
    } else {
      let sheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG.SHEET_ORDERS);
        sheet.appendRow(["Thoi gian", "Ma don", "Loai", "Ten khach", "SDT/Zalo", "Dia chi", "San pham", "Tong tien", "Goi qua", "Ghi chu"]);
      }
      const orderCode = data.orderCode || ("HSE-" + Date.now().toString().slice(-6));
      sheet.appendRow([
        now,
        orderCode,
        data.mode || "Dat hang",
        data.name || "",
        data.phone || "",
        data.address || "",
        data.itemsText || "",
        data.total || 0,
        data.giftWrap ? "Co" : "Khong",
        data.note || ""
      ]);

      const msg =
        "💿 [DON MOI / BAO GIA] HUE SOUND EXPRESS\n" +
        "• Ma: " + orderCode + " (" + (data.mode || "Dat hang") + ")\n" +
        "• Khach: " + (data.name || "") + "\n" +
        "• SDT/Zalo: " + (data.phone || "") + "\n" +
        "• Dia chi: " + (data.address || "") + "\n" +
        "• San pham:\n" + (data.itemsText || "") + "\n" +
        "• Tong niem yet: " + Number(data.total || 0).toLocaleString("vi-VN") + " d\n" +
        "• Goi qua: " + (data.giftWrap ? "Co" : "Khong") + "\n" +
        "• Ghi chu: " + (data.note || "Khong");

      sendNotifications("💿 [Hue Sound Express] " + orderCode + " - " + (data.name || ""), msg);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendNotifications(subject, text) {
  // 1. Bao Telegram (tuc thi 1 giay ve dien thoai, 0d)
  if (CONFIG.TELEGRAM_BOT_TOKEN && !CONFIG.TELEGRAM_BOT_TOKEN.startsWith("DIEN_")) {
    try {
      UrlFetchApp.fetch("https://api.telegram.org/bot" + CONFIG.TELEGRAM_BOT_TOKEN + "/sendMessage", {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({
          chat_id: CONFIG.TELEGRAM_CHAT_ID,
          text: text
        })
      });
    } catch (e) {}
  }

  // 2. Bao Email (du phong 0d)
  if (CONFIG.NOTIFY_EMAIL && CONFIG.NOTIFY_EMAIL.indexOf("@") > 0 && CONFIG.NOTIFY_EMAIL !== "email-cua-ban@gmail.com") {
    try {
      MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, subject, text);
    } catch (e) {}
  }
}
