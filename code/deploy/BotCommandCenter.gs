// ==========================================
// File: BotCommandCenter.gs (Production Ready)
// ระบบ Telegram Bot Command Center สำหรับ BTT Field Tool
// ==========================================

function tgToken_() {
  return (
    PropertiesService.getScriptProperties().getProperty("TELEGRAM_TOKEN") ||
    CONFIG.TELEGRAM_TOKEN ||
    ""
  ).trim();
}

function isTelegramChatAllowed_(chatId) {
  const idStr = String(chatId || "").trim();
  if (!idStr) return false;

  const propAllowed = PropertiesService.getScriptProperties().getProperty(
    "TELEGRAM_ALLOWED_CHATS",
  );
  let list = [];
  if (propAllowed) {
    list = propAllowed
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);
  } else if (
    CONFIG.TELEGRAM_ALLOWED_CHATS &&
    Array.isArray(CONFIG.TELEGRAM_ALLOWED_CHATS)
  ) {
    list = CONFIG.TELEGRAM_ALLOWED_CHATS.map((s) => String(s).trim());
  }

  if (CONFIG.TELEGRAM_CHAT_ID) {
    const mainId = String(CONFIG.TELEGRAM_CHAT_ID).trim();
    if (list.indexOf(mainId) < 0) list.push(mainId);
  }
  if (CONFIG.TELEGRAM_ADMIN_CHAT_ID) {
    const adminId = String(CONFIG.TELEGRAM_ADMIN_CHAT_ID).trim();
    if (list.indexOf(adminId) < 0) list.push(adminId);
  }

  return list.indexOf(idStr) >= 0;
}

/**
 * รับคำสั่ง Webhook POST จาก Telegram Bot (ตรวจสอบ Secret + Chat Whitelist)
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return HtmlService.createHtmlOutput("OK");
    }

    // ตรวจสอบ Secret Token ใน Query String (?secret=...) หากมีการตั้งค่าไว้
    const expectedSecret = PropertiesService.getScriptProperties().getProperty(
      "TELEGRAM_WEBHOOK_SECRET",
    );
    if (expectedSecret) {
      const incomingSecret = (e.parameter && e.parameter.secret) || "";
      if (incomingSecret !== expectedSecret) {
        console.warn("Unauthorized Telegram Webhook attempt (Invalid Secret)");
        return HtmlService.createHtmlOutput("OK");
      }
    }

    const contents = JSON.parse(e.postData.contents);

    // 1. รับข้อความปกติ (Message / Command)
    if (contents.message) {
      const chatId = contents.message.chat && contents.message.chat.id;
      if (!isTelegramChatAllowed_(chatId)) {
        sendTelegramMessage_(
          chatId,
          "⛔ <b>ไม่มีสิทธิ์เข้าถึงระบบ</b>\nChat ID ของคุณ: <code>" +
            escapeHtml_(chatId) +
            "</code>\nกรุณาแจ้งผู้ดูแลระบบเพื่อเพิ่มสิทธิ์การใช้งาน",
        );
        return HtmlService.createHtmlOutput("OK");
      }
      processIncomingMessage_(contents.message);
    }
    // 2. รับการกดปุ่ม Interactive (Callback Query)
    else if (contents.callback_query) {
      const chatId =
        contents.callback_query.message &&
        contents.callback_query.message.chat &&
        contents.callback_query.message.chat.id;
      if (!isTelegramChatAllowed_(chatId)) {
        answerCallbackQuery_(
          contents.callback_query.id,
          "ไม่มีสิทธิ์ใช้งานระบบ",
        );
        return HtmlService.createHtmlOutput("OK");
      }
      processCallbackQuery_(contents.callback_query);
    }
  } catch (err) {
    console.error("Error in doPost Webhook: " + err.toString());
  }
  return HtmlService.createHtmlOutput("OK");
}

/**
 * ประมวลผลคำสั่งข้อความที่เข้ามา
 */
function processIncomingMessage_(msg) {
  const chatId = msg.chat.id;
  const text = (msg.text || "").trim();
  const userName =
    (msg.from && (msg.from.first_name || msg.from.username)) || "Engineer";

  if (!text) return;

  // ตัด @BotUsername กรณีสั่งงานในกลุ่ม เช่น /menu@BTTBot
  const cleanCmd = text.replace(/^(\/[a-zA-Z0-9_]+)@[a-zA-Z0-9_]+/, "$1");

  if (
    cleanCmd === "/start" ||
    cleanCmd === "/menu" ||
    cleanCmd.toLowerCase() === "menu" ||
    cleanCmd === "เมนู"
  ) {
    sendMainMenu_(chatId, userName);
  } else if (
    cleanCmd === "/summary" ||
    cleanCmd === "📊 สรุปภาพรวมคลัง" ||
    cleanCmd === "สรุป"
  ) {
    sendSummaryReport_(chatId);
  } else if (
    cleanCmd === "/borrowed" ||
    cleanCmd === "📦 อุปกรณ์ที่ถูกยืม" ||
    cleanCmd === "ยืม"
  ) {
    sendBorrowedList_(chatId);
  } else if (
    cleanCmd === "/overdue" ||
    cleanCmd === "⚠️ รายการยืมค้างนาน" ||
    cleanCmd === "ค้างนาน"
  ) {
    sendOverdueReport_(chatId);
  } else if (
    cleanCmd === "/lowstock" ||
    cleanCmd === "🛒 ของใกล้หมด/ต้องสั่งซื้อ" ||
    cleanCmd === "ซื้อเพิ่ม"
  ) {
    sendLowStockReport_(chatId);
  } else if (/^(\/find|ค้นหา)\s+/i.test(cleanCmd)) {
    const query = cleanCmd.replace(/^(\/find|ค้นหา)\s+/i, "").trim();
    searchAssetsBot_(chatId, query);
  } else if (/^(\/item|รหัส)\s+/i.test(cleanCmd)) {
    const code = cleanCmd.replace(/^(\/item|รหัส)\s+/i, "").trim();
    sendItemDetailsBot_(chatId, code);
  } else {
    // ส่งให้ AI ประมวลผลคำถามภาษาธรรมชาติ (Natural Language Query)
    answerWithAI_(chatId, text, userName);
  }
}

/**
 * ประมวลผลการกดปุ่ม (Inline Keyboard Callbacks)
 */
function processCallbackQuery_(cb) {
  const callbackId = cb.id;
  const chatId = cb.message.chat.id;
  const data = cb.data;

  answerCallbackQuery_(callbackId, "กำลังประมวลผล...");

  if (data === "cmd_summary") {
    sendSummaryReport_(chatId);
  } else if (data === "cmd_borrowed") {
    sendBorrowedList_(chatId);
  } else if (data === "cmd_overdue") {
    sendOverdueReport_(chatId);
  } else if (data === "cmd_lowstock") {
    sendLowStockReport_(chatId);
  } else if (data.startsWith("view_")) {
    const code = data.replace("view_", "");
    sendItemDetailsBot_(chatId, code);
  } else if (data === "cmd_menu") {
    sendMainMenu_(chatId, (cb.from && cb.from.first_name) || "Engineer");
  }
}

// ==========================================
// 🚀 COMMAND CENTER FUNCTIONS & REPORTS
// ==========================================

function buildAssetNameMap_(assets) {
  const C = CONFIG.COL;
  const map = {};
  (assets || []).forEach((r) => {
    const id = String(r[C.ID] || "").trim();
    if (id) map[id] = String(r[C.NAME] || "").trim();
  });
  return map;
}

function sendMainMenu_(chatId, userName) {
  const menuText =
    "⚡ <b>BTT ENGINEER BOT - Command Center</b>\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "สวัสดีครับคุณ <b>" +
    escapeHtml_(userName) +
    "</b> 👋\n" +
    "ระบบจัดการเครื่องมืออัจฉริยะ พร้อมปฏิบัติการ!\n\n" +
    "💡 <i>พิมพ์ค้นหาอุปกรณ์ได้ทันที เช่น: <code>ค้นหา สว่าน</code> หรือ <code>/find BTT-TO014</code></i>";

  const keyboard = {
    inline_keyboard: [
      [
        { text: "📊 สรุปภาพรวมคลัง", callback_data: "cmd_summary" },
        { text: "📦 อุปกรณ์ที่ถูกยืม", callback_data: "cmd_borrowed" },
      ],
      [
        { text: "⚠️ รายการยืมค้างนาน", callback_data: "cmd_overdue" },
        { text: "🛒 สต็อกต่ำ/ต้องสั่งซื้อ", callback_data: "cmd_lowstock" },
      ],
      [{ text: "🌐 เปิดหน้าเว็บ BTT Field Tool", url: appUrl_() }],
    ],
  };

  sendTelegramMessage_(chatId, menuText, keyboard);
}

function sendSummaryReport_(chatId) {
  const data = baseData_();
  const summary = data.summary;

  let msg =
    "📊 <b>สรุปภาพรวมคลังเครื่องมือ BTT</b>\n" +
    "📅 ข้อมูล ณ วันที่: " +
    formatThaiDateTime_(new Date()) +
    "\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "🔹 อุปกรณ์ทั้งหมด: <b>" +
    summary.total +
    "</b> รายการ\n" +
    "🟢 พร้อมใช้งาน: <b>" +
    summary.available +
    "</b> รายการ\n" +
    "🟡 กำลังถูกยืม: <b>" +
    summary.borrowed +
    "</b> รายการ\n" +
    "🔴 ไม่พร้อมใช้งาน/ชำรุด: <b>" +
    summary.unavailable +
    "</b> รายการ\n" +
    "⚠️ สต็อกต่ำกว่า Min: <b>" +
    summary.lowStockCount +
    "</b> รายการ\n";

  if (
    summary.unavailableDetail &&
    Object.keys(summary.unavailableDetail).length > 0
  ) {
    msg += "\n<b>🛠️ รายละเอียดไม่พร้อมใช้งาน:</b>\n";
    for (let k in summary.unavailableDetail) {
      msg +=
        "  • " +
        escapeHtml_(k) +
        ": " +
        summary.unavailableDetail[k] +
        " รายการ\n";
    }
  }

  const keyboard = {
    inline_keyboard: [
      [
        { text: "📦 ดูรายการที่ถูกยืม", callback_data: "cmd_borrowed" },
        { text: "🛒 ดูรายการต้องซื้อ", callback_data: "cmd_lowstock" },
      ],
      [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
    ],
  };

  sendTelegramMessage_(chatId, msg, keyboard);
}

function sendBorrowedList_(chatId) {
  const data = baseData_();
  const activeBorrows = data.borrows.filter((r) => r[8] === "กำลังยืม");
  const nameMap = buildAssetNameMap_(data.assets);

  if (activeBorrows.length === 0) {
    sendTelegramMessage_(chatId, "🟢 <b>ขณะนี้ไม่มีอุปกรณ์ค้างยืมในระบบ</b>", {
      inline_keyboard: [
        [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
      ],
    });
    return;
  }

  let msg =
    "📦 <b>รายการอุปกรณ์ที่กำลังถูกยืม (" +
    activeBorrows.length +
    " รายการ)</b>\n━━━━━━━━━━━━━━━━━━━━━━━\n\n";

  activeBorrows.slice(0, 15).forEach((r, idx) => {
    const code = String(r[1] || "").trim();
    const itemName = nameMap[code] || "";
    msg +=
      idx +
      1 +
      ". <b>" +
      escapeHtml_(code) +
      "</b>" +
      (itemName ? " - " + escapeHtml_(itemName) : "") +
      "\n";
    msg +=
      "   👤 ผู้เบิก: <b>" +
      escapeHtml_(r[6]) +
      "</b> | 🏢 โครงการ: " +
      escapeHtml_(r[2]) +
      "\n";
    msg += "   📅 วันที่เบิก: " + escapeHtml_(r[4]) + "\n\n";
  });

  if (activeBorrows.length > 15) {
    msg += "<i>...และอีก " + (activeBorrows.length - 15) + " รายการ</i>\n";
  }

  const keyboard = {
    inline_keyboard: [
      [
        {
          text: "⚠️ ดูรายการยืมค้างนานเกิน 7 วัน",
          callback_data: "cmd_overdue",
        },
      ],
      [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
    ],
  };

  sendTelegramMessage_(chatId, msg, keyboard);
}

function sendOverdueReport_(chatId) {
  const data = baseData_();
  const activeBorrows = data.borrows.filter((r) => r[8] === "กำลังยืม");
  const nameMap = buildAssetNameMap_(data.assets);
  const limitDays = CONFIG.AI_COMMAND_CENTER.OVERDUE_DAYS_ALERT || 7;
  const now = new Date().getTime();

  const overdueList = [];
  activeBorrows.forEach((r) => {
    const d = aiParseDate_(r[4]);
    if (d) {
      const days = Math.floor((now - d.getTime()) / 86400000);
      if (days >= limitDays) {
        overdueList.push({ row: r, days: days });
      }
    }
  });

  if (overdueList.length === 0) {
    sendTelegramMessage_(
      chatId,
      "✅ <b>ไม่มีรายการยืมค้างเกิน " + limitDays + " วัน</b>",
      {
        inline_keyboard: [
          [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
        ],
      },
    );
    return;
  }

  overdueList.sort((a, b) => b.days - a.days);

  let msg =
    "🚨 <b>เตือน! รายการยืมค้างนานเกิน " +
    limitDays +
    " วัน (" +
    overdueList.length +
    " รายการ)</b>\n━━━━━━━━━━━━━━━━━━━━━━━\n\n";

  overdueList.slice(0, 15).forEach((item, idx) => {
    const r = item.row;
    const code = String(r[1] || "").trim();
    const itemName = nameMap[code] || "";
    msg +=
      idx +
      1 +
      ". 🔴 <b>" +
      escapeHtml_(code) +
      "</b>" +
      (itemName ? " (" + escapeHtml_(itemName) + ")" : "") +
      " — <b>ค้าง " +
      item.days +
      " วัน</b>\n";
    msg +=
      "   👤 ผู้เบิก: <b>" +
      escapeHtml_(r[6]) +
      "</b>\n   📍 สถานที่นำไปใช้: " +
      escapeHtml_(r[7]) +
      "\n\n";
  });

  if (overdueList.length > 15) {
    msg += "<i>...และอีก " + (overdueList.length - 15) + " รายการ</i>\n";
  }

  const keyboard = {
    inline_keyboard: [[{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }]],
  };

  sendTelegramMessage_(chatId, msg, keyboard);
}

function sendLowStockReport_(chatId) {
  const C = CONFIG.COL;
  const data = baseData_();
  const lowItems = data.assets.filter(
    (r) =>
      r[C.NEEDBUY] === "ต้องซื้อเพิ่ม" || Number(r[C.TOTAL]) < Number(r[C.MIN]),
  );

  if (lowItems.length === 0) {
    sendTelegramMessage_(
      chatId,
      "✅ <b>สต็อกอุปกรณ์อยู่ในระดับปกติทุกรายการ</b>",
      {
        inline_keyboard: [
          [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
        ],
      },
    );
    return;
  }

  let msg =
    "🛒 <b>รายการสต็อกต่ำกว่า Min / ต้องสั่งซื้อ (" +
    lowItems.length +
    " รายการ)</b>\n━━━━━━━━━━━━━━━━━━━━━━━\n\n";

  lowItems.slice(0, 15).forEach((r, idx) => {
    msg +=
      idx +
      1 +
      ". <b>" +
      escapeHtml_(r[C.ID]) +
      "</b> - " +
      escapeHtml_(r[C.NAME]) +
      "\n";
    msg +=
      "   📉 คงเหลือ: <b>" +
      r[C.TOTAL] +
      "</b> (Min: " +
      r[C.MIN] +
      ") | หมวด: " +
      escapeHtml_(r[C.CAT]) +
      "\n\n";
  });

  if (lowItems.length > 15) {
    msg += "<i>...และอีก " + (lowItems.length - 15) + " รายการ</i>\n";
  }

  const keyboard = {
    inline_keyboard: [[{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }]],
  };

  sendTelegramMessage_(chatId, msg, keyboard);
}

function searchAssetsBot_(chatId, query) {
  const C = CONFIG.COL;
  const data = baseData_();
  const q = aiNorm_(query);

  const hits = data.assets
    .filter((r) => {
      return (
        aiNorm_(r[C.ID]).includes(q) ||
        aiNorm_(r[C.NAME]).includes(q) ||
        aiNorm_(r[C.CAT]).includes(q)
      );
    })
    .slice(0, 6);

  if (hits.length === 0) {
    sendTelegramMessage_(
      chatId,
      "🔍 ไม่พบอุปกรณ์ที่ตรงกับคำค้น: <b>" + escapeHtml_(query) + "</b>",
      {
        inline_keyboard: [
          [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
        ],
      },
    );
    return;
  }

  let msg =
    "🔍 <b>ผลการค้นหา: " +
    escapeHtml_(query) +
    " (" +
    hits.length +
    " รายการ)</b>\n━━━━━━━━━━━━━━━━━━━━━━━\n\n";
  const buttons = [];

  hits.forEach((r) => {
    const statusIcon = r[C.STATUS] === "พร้อมใช้งาน" ? "🟢" : "🔴";
    msg +=
      statusIcon +
      " <b>" +
      escapeHtml_(r[C.ID]) +
      "</b>: " +
      escapeHtml_(r[C.NAME]) +
      "\n";
    msg +=
      "   📍 " +
      escapeHtml_(r[C.LOC]) +
      " | คงเหลือ: " +
      r[C.TOTAL] +
      " (" +
      escapeHtml_(r[C.STATUS]) +
      ")\n\n";

    buttons.push([
      {
        text: "🔍 ดูรายละเอียด " + r[C.ID],
        callback_data: "view_" + r[C.ID],
      },
    ]);
  });

  buttons.push([{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }]);

  sendTelegramMessage_(chatId, msg, { inline_keyboard: buttons });
}

function sendItemDetailsBot_(chatId, code) {
  const C = CONFIG.COL;
  const data = baseData_();
  const item = data.assets.find(
    (r) =>
      String(r[C.ID]).trim().toLowerCase() ===
        String(code).trim().toLowerCase() ||
      codeMatch_(code, String(r[C.ID]).trim()),
  );

  if (!item) {
    sendTelegramMessage_(
      chatId,
      "❌ ไม่พบอุปกรณ์รหัส: <b>" + escapeHtml_(code) + "</b>",
    );
    return;
  }

  let msg =
    "📦 <b>รายละเอียดอุปกรณ์: " +
    escapeHtml_(item[C.ID]) +
    "</b>\n━━━━━━━━━━━━━━━━━━━━━━━\n";
  msg += "🔹 <b>ชื่อรายการ:</b> " + escapeHtml_(item[C.NAME]) + "\n";
  msg += "📁 <b>หมวดหมู่:</b> " + escapeHtml_(item[C.CAT]) + "\n";
  msg += "📍 <b>สถานที่เก็บ:</b> " + escapeHtml_(item[C.LOC]) + "\n";
  msg +=
    "📊 <b>คงเหลือรวม:</b> " +
    item[C.TOTAL] +
    " (BTT: " +
    item[C.BTT] +
    ", TKE: " +
    item[C.TKE] +
    ")\n";
  msg += "🏷️ <b>สถานะ:</b> " + escapeHtml_(item[C.STATUS]) + "\n";
  if (item[C.REMARK]) {
    msg += "📝 <b>หมายเหตุ:</b> " + escapeHtml_(item[C.REMARK]) + "\n";
  }

  const rawImg = item[C.IMG];
  const firstPic = rawImg ? rawImg.split(",")[0].trim() : null;

  const keyboard = {
    inline_keyboard: [
      [
        {
          text: "🌐 ดูในเว็บ / สแกนเบิก",
          url: appUrl_() + "?code=" + encodeURIComponent(item[C.ID]),
        },
      ],
      [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
    ],
  };

  if (firstPic && /^https?:\/\//i.test(firstPic)) {
    sendTelegramPhoto_(chatId, firstPic, msg, keyboard);
  } else {
    sendTelegramMessage_(chatId, msg, keyboard);
  }
}

function formatAiToTelegramHtml_(text) {
  let safe = escapeHtml_(text || "");
  safe = safe.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
  return safe;
}

function answerWithAI_(chatId, userQuery, userName) {
  sendChatAction_(chatId, "typing");

  try {
    const aiRes = askAI(userQuery, "", userName, "telegram_bot_" + chatId, []);

    if (aiRes && aiRes.success && aiRes.answer) {
      const formatted = formatAiToTelegramHtml_(aiRes.answer);
      sendTelegramMessage_(
        chatId,
        "🤖 <b>BTT AI Assistant:</b>\n\n" + formatted,
        {
          inline_keyboard: [
            [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
          ],
        },
      );
    } else {
      const errMsg =
        (aiRes && aiRes.message) || "ไม่สามารถประมวลผลคำตอบได้ในขณะนี้";
      sendTelegramMessage_(chatId, "⚠️ " + escapeHtml_(errMsg), {
        inline_keyboard: [
          [{ text: "🔙 กลับเมนูหลัก", callback_data: "cmd_menu" }],
        ],
      });
    }
  } catch (err) {
    sendTelegramMessage_(
      chatId,
      "⚠️ เกิดข้อผิดพลาด: " + escapeHtml_(err.toString()),
    );
  }
}

// ==========================================
// 📬 TELEGRAM API HELPERS & TRIGGERS
// ==========================================

function sendTelegramMessage_(chatId, htmlText, inlineKeyboard) {
  const token = tgToken_();
  if (!token || token === "YOUR_TELEGRAM_BOT_TOKEN") return;

  const url = "https://api.telegram.org/bot" + token + "/sendMessage";
  const payload = {
    chat_id: chatId,
    text: htmlText,
    parse_mode: "HTML",
    disable_web_page_preview: true,
  };

  if (inlineKeyboard) payload.reply_markup = JSON.stringify(inlineKeyboard);

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  // Fallback หาก HTML Parse Error ให้ส่งแบบข้อความธรรมดาแทน เพื่อไม่ให้บอทเงียบ
  if (res.getResponseCode() !== 200) {
    const plainPayload = {
      chat_id: chatId,
      text: String(htmlText).replace(/<[^>]+>/g, ""),
      disable_web_page_preview: true,
    };
    if (inlineKeyboard)
      plainPayload.reply_markup = JSON.stringify(inlineKeyboard);
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(plainPayload),
      muteHttpExceptions: true,
    });
  }
}

function sendTelegramPhoto_(chatId, photoUrl, captionHtml, inlineKeyboard) {
  const token = tgToken_();
  if (!token || token === "YOUR_TELEGRAM_BOT_TOKEN") return;

  const url = "https://api.telegram.org/bot" + token + "/sendPhoto";
  const payload = {
    chat_id: chatId,
    photo: photoUrl,
    caption: captionHtml,
    parse_mode: "HTML",
  };

  if (inlineKeyboard) payload.reply_markup = JSON.stringify(inlineKeyboard);

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  // ถ้ารูปโหลดไม่ขึ้นหรือเกินขนาด ให้ส่งเป็นข้อความแทน
  if (res.getResponseCode() !== 200) {
    sendTelegramMessage_(chatId, captionHtml, inlineKeyboard);
  }
}

function sendChatAction_(chatId, action) {
  const token = tgToken_();
  if (!token) return;
  UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + token + "/sendChatAction",
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, action: action || "typing" }),
      muteHttpExceptions: true,
    },
  );
}

function answerCallbackQuery_(callbackId, text) {
  const token = tgToken_();
  if (!token) return;
  UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + token + "/answerCallbackQuery",
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ callback_query_id: callbackId, text: text }),
      muteHttpExceptions: true,
    },
  );
}

function escapeHtml_(text) {
  if (text == null) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function formatThaiDateTime_(date) {
  return Utilities.formatDate(date, "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
}

function sendDailySummaryTrigger() {
  const chatId = CONFIG.TELEGRAM_CHAT_ID;
  if (chatId && chatId !== "YOUR_CHAT_ID") {
    sendSummaryReport_(chatId);
    sendOverdueReport_(chatId);
  }
}

/**
 * 🛠️ ตั้งค่า Telegram Token ผ่าน Editor (ปลอดภัยกว่าเก็บในโค้ด)
 */
function setTelegramToken(token) {
  ownerOnly_();
  PropertiesService.getScriptProperties().setProperty(
    "TELEGRAM_TOKEN",
    String(token || "").trim(),
  );
  return "บันทึก TELEGRAM_TOKEN สำเร็จ";
}

/**
 * 🛠️ ตั้งค่า Webhook พร้อมสร้าง Secret Token อัตโนมัติ (รันครั้งเดียวใน Editor)
 */
function setupTelegramWebhook() {
  ownerOnly_();
  const token = tgToken_();
  if (!token)
    throw new Error("ยังไม่ได้ตั้ง TELEGRAM_TOKEN ใน Script Properties");

  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("TELEGRAM_WEBHOOK_SECRET");
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, "");
    props.setProperty("TELEGRAM_WEBHOOK_SECRET", secret);
  }

  const webhookUrl = appUrl_() + "?secret=" + encodeURIComponent(secret);
  const url =
    "https://api.telegram.org/bot" +
    token +
    "/setWebhook?url=" +
    encodeURIComponent(webhookUrl);
  const response = UrlFetchApp.fetch(url);
  Logger.log("Webhook Result: " + response.getContentText());
}

/**
 * ⏰ ตั้งเวลา Trigger รายงานสรุปทุกเช้า 08:00 น.
 */
function setupDailyReportTrigger() {
  ownerOnly_();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction() === "sendDailySummaryTrigger") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("sendDailySummaryTrigger")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.AI_COMMAND_CENTER.AUTO_REPORT_HOUR || 8)
    .create();

  Logger.log("ตั้งค่า Daily Report Trigger เรียบร้อยแล้ว!");
}

/**
 * 🔔 แจ้งเตือนเข้า Telegram เมื่อมีการเบิก/คืน (ไม่ทำให้ระบบหลักล้มถ้าส่งไม่สำเร็จ)
 */
function notifyTelegram_(htmlText) {
  try {
    const ids = [];
    [CONFIG.TELEGRAM_CHAT_ID, CONFIG.TELEGRAM_ADMIN_CHAT_ID].forEach((id) => {
      const s = String(id || "").trim();
      if (s && s !== "YOUR_CHAT_ID" && ids.indexOf(s) < 0) ids.push(s);
    });
    ids.forEach((id) => sendTelegramMessage_(id, htmlText));
  } catch (e) {
    console.error("notifyTelegram_ error: " + e);
  }
}

function itemNameByCode_(code) {
  try {
    const sh = assetSheet_();
    const r = findAssetRow_(sh, code);
    return r > 0 ? String(sh.getRange(r, CONFIG.COL.NAME + 1).getValue()) : "";
  } catch (e) {
    return "";
  }
}

function notifyBorrow_(borrower, project, place, items, operator) {
  let msg =
    "📤 <b>มีการเบิกอุปกรณ์</b>\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "👤 ผู้เบิก: <b>" +
    escapeHtml_(borrower) +
    "</b>\n" +
    "🏢 โครงการ: " +
    escapeHtml_(project || "-") +
    "\n" +
    "📍 สถานที่ใช้งาน: " +
    escapeHtml_(place || "-") +
    "\n" +
    "🕒 เวลา: " +
    formatThaiDateTime_(new Date()) +
    "\n\n" +
    "<b>รายการ (" +
    items.length +
    "):</b>\n";
  items.forEach((it, i) => {
    const nm = itemNameByCode_(it.itemCode);
    msg +=
      i +
      1 +
      ". <code>" +
      escapeHtml_(it.itemCode) +
      "</code>" +
      (nm ? " - " + escapeHtml_(nm) : "") +
      " × <b>" +
      it.quantity +
      "</b>\n";
  });
  if (operator && operator !== borrower) {
    msg += "\n✍️ ทำรายการโดย: " + escapeHtml_(operator);
  }
  notifyTelegram_(msg);
}

function notifyReturn_(code, qty, borrower, returnedBy, retLoc) {
  const nm = itemNameByCode_(code);
  const msg =
    "📥 <b>มีการคืนอุปกรณ์</b>\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "🔹 <code>" +
    escapeHtml_(code) +
    "</code>" +
    (nm ? " - " + escapeHtml_(nm) : "") +
    " × <b>" +
    qty +
    "</b>\n" +
    "👤 ผู้เบิก: " +
    escapeHtml_(borrower) +
    "\n" +
    "↩️ ผู้คืน: <b>" +
    escapeHtml_(returnedBy) +
    "</b>\n" +
    "📍 วางไว้ที่: " +
    escapeHtml_(retLoc) +
    "\n" +
    "🕒 เวลา: " +
    formatThaiDateTime_(new Date());
  notifyTelegram_(msg);
}

function testTelegram() {
  const token = tgToken_();
  Logger.log("ความยาว token: " + token.length);
  if (!token) {
    Logger.log("❌ ยังไม่ได้ตั้ง TELEGRAM_TOKEN ใน Script Properties");
    return;
  }

  const me = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + token + "/getMe",
    { muteHttpExceptions: true },
  );
  Logger.log("getMe: " + me.getContentText());

  const chat = String(CONFIG.TELEGRAM_CHAT_ID).trim();
  const res = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + token + "/sendMessage",
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chat, text: "✅ ทดสอบแจ้งเตือน" }),
      muteHttpExceptions: true,
    },
  );
  Logger.log(
    "sendMessage: " + res.getResponseCode() + " " + res.getContentText(),
  );
}

function findChatIds() {
  const token = tgToken_();
  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/deleteWebhook");
  const r = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + token + "/getUpdates",
  );
  const j = JSON.parse(r.getContentText());
  const seen = {};
  (j.result || []).forEach(function (u) {
    const m = u.message || u.my_chat_member || u.channel_post;
    if (m && m.chat) seen[m.chat.id] = m.chat.title || m.chat.first_name || "";
  });
  Logger.log(JSON.stringify(seen, null, 2));
}
