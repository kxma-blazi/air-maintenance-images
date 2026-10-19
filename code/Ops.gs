// ==================== Ops.gs — ตรวจความพร้อม Production + ล้างข้อมูลตัวอย่าง (v4.1) ====================

// รันจาก Editor ก่อนเปิดใช้งานจริง: ตรวจทุกอย่างแล้วสรุปผลที่ Execution log (และส่งเข้า Telegram ถ้าใช้ได้)
function productionCheck() {
  const out = [];
  const ok = function (m) {
    out.push("✅ " + m);
  };
  const warn = function (m) {
    out.push("⚠️ " + m);
  };
  const bad = function (m) {
    out.push("❌ " + m);
  };

  // 1) Script properties
  [
    "TELEGRAM_TOKEN",
    "TELEGRAM_CHAT_ID",
    "WEBHOOK_SECRET",
    "WEBAPP_URL",
    "GITHUB_TOKEN",
  ].forEach(function (k) {
    prop_(k) ? ok(k + " ตั้งแล้ว") : bad(k + " ยังไม่ได้ตั้ง");
  });
  const dk = prop_("DASH_KEY") || "";
  if (!dk) bad("DASH_KEY ยังไม่ได้ตั้ง (ตอนนี้ใช้ค่า 'free' ใครก็เข้าได้)");
  else if (dk.length < 16 || dk === "free")
    warn(
      "DASH_KEY สั้น/เดาง่าย — ใช้สุ่มยาว ≥16 ตัวอักษร (ลิงก์นี้คือกุญแจเข้าระบบ)",
    );
  else ok("DASH_KEY ยาว " + dk.length + " ตัวอักษร");
  prop_("GROQ_API_KEY") || prop_("GEMINI_API_KEY")
    ? ok("มี AI key")
    : warn("ยังไม่มี AI key (ใช้ /ask ไม่ได้ แต่ระบบหลักไม่กระทบ)");

  // 2) ชีต + ข้อมูล
  try {
    const airs = getAirs_();
    ok("ชีตเครื่องแอร์: " + airs.length + " เครื่อง");
    const noAlert = airs
      .filter(function (a) {
        return !a.alert;
      })
      .map(function (a) {
        return a.id;
      });
    if (noAlert.length)
      warn(
        "ไม่มีสูตรแจ้งเตือน: " + noAlert.join(", ") + " → รัน fixAirFormulas()",
      );
    const gaps = airs
      .filter(function (a) {
        return !a.btu || !a.model;
      })
      .map(function (a) {
        return a.id;
      });
    if (gaps.length)
      warn(
        "ข้อมูลไม่ครบ (รุ่น/BTU): " +
          gaps.join(", ") +
          " (ประมาณการค่าไฟจะไม่รวมเครื่องที่ไม่มี BTU)",
      );
  } catch (e) {
    bad("ชีตเครื่องแอร์: " + e.message);
  }
  try {
    const sh = logSheet_();
    ok(
      "ชีตประวัติซ่อม: " +
        getLogs_().length +
        " รายการ" +
        (sh.getRange(LOG_FIRST_ROW - 1, 9).getValue() ? " (มีคอลัมน์รูป)" : ""),
    );
    const vals = sh
      .getRange(
        LOG_FIRST_ROW,
        8,
        Math.max(1, sh.getLastRow() - LOG_FIRST_ROW + 1),
        1,
      )
      .getValues();
    const n = vals.filter(function (r) {
      return /ตัวอย่าง/.test(String(r[0]));
    }).length;
    n
      ? warn("ยังมีแถวตัวอย่าง " + n + " แถว → รัน cleanSampleData()")
      : ok("ไม่มีแถวตัวอย่างค้าง");
  } catch (e) {
    bad("ชีตประวัติซ่อม: " + e.message);
  }
  try {
    const e = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERROR_SHEET);
    const cnt = e ? Math.max(0, e.getLastRow() - 1) : 0;
    cnt
      ? warn(
          "ชีต _errors มี " +
            cnt +
            " รายการ (ดูข้อความล่าสุดว่ายังเกิดอยู่ไหม)",
        )
      : ok("ไม่มี error สะสม");
  } catch (e) {}

  // 3) บริการภายนอก
  try {
    ok("GitHub: " + ghCheck_());
  } catch (e) {
    bad("GitHub: " + e.message);
  }
  try {
    const r = JSON.parse(
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot" + prop_("TELEGRAM_TOKEN") + "/getMe",
        { muteHttpExceptions: true },
      ).getContentText(),
    );
    r.ok
      ? ok("Telegram bot @" + r.result.username)
      : bad(
          "Telegram: " +
            r.description +
            " (ขอโทเคนใหม่จาก BotFather แล้วรัน setupAll)",
        );
  } catch (e) {
    bad("Telegram: " + e.message);
  }

  // 4) Trigger
  const hs = ScriptApp.getProjectTriggers().map(function (t) {
    return t.getHandlerFunction();
  });
  ["checkAndSendAirNotification", "weeklyReport", "weeklyBackup"].forEach(
    function (h) {
      hs.indexOf(h) !== -1
        ? ok("Trigger " + h)
        : bad("ไม่มี Trigger " + h + " → รัน setupAll()");
    },
  );

  const text = out.join("\n");
  Logger.log(text);
  return text;
}

// รันครั้งเดียวก่อนใช้งานจริง: สำรองไฟล์ → ลบแถวตัวอย่างในประวัติซ่อม ("ตัวอย่าง - ลบได้") → ล้างชีต _errors
// ไม่แตะข้อมูลในชีตเครื่องแอร์
function cleanSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name =
    ss.getName() +
    " - ก่อนล้างตัวอย่าง " +
    Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HHmm");
  DriveApp.getFileById(ss.getId()).makeCopy(name);
  const sh = logSheet_();
  const n = sh.getLastRow() - LOG_FIRST_ROW + 1;
  let cleared = 0;
  if (n > 0) {
    const notes = sh.getRange(LOG_FIRST_ROW, 8, n, 1).getValues();
    notes.forEach(function (r, i) {
      if (/ตัวอย่าง/.test(String(r[0]))) {
        sh.getRange(LOG_FIRST_ROW + i, 1, 1, 9).clearContent();
        cleared++;
      }
    });
  }
  const e = ss.getSheetByName(ERROR_SHEET);
  if (e && e.getLastRow() > 1) e.deleteRows(2, e.getLastRow() - 1);
  SpreadsheetApp.flush();
  Logger.log(
    "สำรองเป็น '" +
      name +
      "' แล้ว | ล้างแถวตัวอย่าง " +
      cleared +
      " แถว | ล้าง _errors",
  );
}
