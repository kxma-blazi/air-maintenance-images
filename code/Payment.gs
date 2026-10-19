// ==================== Payment.gs — ประวัติสลีปโอนเงิน (v2) ====================
// เก็บรูปสลีปขึ้น GitHub (โฟลเดอร์ slips/) + บันทึกแถวในชีต "ประวัติการจ่ายเงิน"
// ต้องมี Script property GITHUB_TOKEN (ตัวเดียวกับที่ใช้อัปโหลดรูปแอร์)
const PAYMENT_SHEET = "ประวัติการจ่ายเงิน";
const PAY_HEADER = [
  "เวลา",
  "รหัสเครื่อง",
  "จำนวนเงิน (บาท)",
  "หมายเหตุ",
  "รูปสลีป (ลิงก์)",
  "ผู้บันทึก",
];
const PAY = { TIME: 0, ID: 1, AMOUNT: 2, NOTE: 3, SLIP: 4, BY: 5 };
const SLIP_LIST_LIMIT = 300;

function paymentSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(PAYMENT_SHEET);
  if (!s) s = ss.insertSheet(PAYMENT_SHEET);
  // ชีตที่มีอยู่แล้วแต่ยังไม่มีหัวตาราง: เติมหัวตาราง (ถ้ามีข้อมูลอยู่แถว 1 จะดันลงไปเป็นแถว 2)
  if (String(s.getRange(1, 1).getValue()) !== PAY_HEADER[0]) {
    if (!s.getRange(1, 1, 1, PAY_HEADER.length).isBlank()) s.insertRowBefore(1);
    s.getRange(1, 1, 1, PAY_HEADER.length)
      .setValues([PAY_HEADER])
      .setFontWeight("bold")
      .setBackground("#e8eef7");
    s.setFrozenRows(1);
    s.getRange("A2:A").setNumberFormat("dd/mm/yyyy hh:mm");
    s.getRange("B2:B").setNumberFormat("@");
    s.getRange("C2:C").setNumberFormat("#,##0.00");
    s.getRange("D2:F").setNumberFormat("@");
    s.setColumnWidth(5, 280);
  }
  return s;
}

function slipUrlOk_(u) {
  return (
    typeof u === "string" &&
    u.indexOf(ghRawBase_() + "slips/") === 0 &&
    /^[A-Za-z0-9_\-\/.:]+$/.test(u)
  );
}

function slipList_() {
  const sh = paymentSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { slips: [], total: 0, count: 0 };
  const first = Math.max(2, last - SLIP_LIST_LIMIT + 1);
  const v = sh
    .getRange(first, 1, last - first + 1, PAY_HEADER.length)
    .getValues();
  const out = [];
  let total = 0;
  v.forEach(function (r, i) {
    const url = String(r[PAY.SLIP] || "").trim();
    const amount = Number(r[PAY.AMOUNT]) || 0;
    if (!slipUrlOk_(url) && !amount) return; // แถวว่าง/แถวเสีย
    total += amount;
    out.push({
      row: first + i,
      time:
        r[PAY.TIME] instanceof Date
          ? Utilities.formatDate(r[PAY.TIME], TZ, "dd/MM/yyyy HH:mm")
          : String(r[PAY.TIME] || ""),
      id: String(r[PAY.ID] || "").trim(),
      amount: amount,
      note: String(r[PAY.NOTE] || ""),
      url: slipUrlOk_(url) ? url : "",
      by: String(r[PAY.BY] || ""),
    });
  });
  out.reverse(); // ใหม่ → เก่า
  return { slips: out, total: total, count: out.length };
}

// o = { data: "data:image/jpeg;base64,...", id?, amount?, note? }
function saveSlip_(o, actor) {
  if (!ghReady_())
    throw userErr_(
      "ยังไม่ได้ตั้ง GITHUB_TOKEN ใน Script properties (อัปโหลดสลีปไม่ได้)",
    );

  const b64 = String(o.data || o.slipData || "").replace(
    /^data:image\/jpeg;base64,/,
    "",
  );
  if (!b64) throw userErr_("กรุณาเลือกรูปสลีปก่อน");
  if (b64.length > PHOTO_MAX_B64)
    throw userErr_("ไฟล์รูปใหญ่เกินไป ลองถ่ายใหม่");
  if (!/^[A-Za-z0-9+\/]+=*$/.test(b64) || b64.indexOf("/9j/") !== 0)
    throw userErr_("รองรับเฉพาะรูป JPEG");

  const id = normId_(clean_(o.id, 20));
  if (
    id &&
    !getAirs_().some(function (a) {
      return a.id === id;
    })
  )
    throw userErr_("ไม่พบเครื่อง " + id);
  const amount = cleanNum_(o.amount, 0.01, 10000000, "จำนวนเงิน", false);
  const note = clean_(o.note, 200);

  // กันกดซ้ำ/อัปโหลดรูปเดิมซ้ำภายใน 10 นาที
  const cache = CacheService.getScriptCache();
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, b64)
    .map(function (b) {
      return ("0" + (b & 0xff).toString(16)).slice(-2);
    })
    .join("");
  if (cache.get("slip_" + hash))
    throw userErr_("สลีปใบนี้เพิ่งอัปโหลดไปแล้ว (กันบันทึกซ้ำ)");

  const stamp = Utilities.formatDate(new Date(), TZ, "yyyyMMdd-HHmmss");
  const rnd = Utilities.getUuid().replace(/-/g, "").substring(0, 4);
  const path = "slips/" + (id ? id + "-" : "") + stamp + "-" + rnd + ".jpg";
  const url = ghPut_(path, b64, "สลีปโอนเงิน " + (id || "general"), false);

  withLock_(function () {
    const sh = paymentSheet_();
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, PAY_HEADER.length)
      .setNumberFormats([["dd/mm/yyyy hh:mm", "@", "#,##0.00", "@", "@", "@"]])
      .setValues([
        [
          new Date(),
          id || "-",
          amount === null ? "" : amount,
          note,
          url,
          actor || "dashboard",
        ],
      ]);
  });
  cache.put("slip_" + hash, "1", 600);
  audit_(
    actor || "dashboard",
    "web",
    "[สลีป] " + (id || "-") + (amount ? " " + amount + " บาท" : ""),
  );
  SpreadsheetApp.flush();
  return url;
}

function apiSlipList(key) {
  return api_("slipList", key, false, slipList_);
}

function apiSlipSave(key, o) {
  return api_("slipSave", key, true, function () {
    saveSlip_(o || {}, "dashboard");
    return slipList_();
  });
}

// ลบแถวบันทึก (รูปยังเก็บอยู่ใน GitHub เป็นหลักฐาน) o = { row, url }
function apiSlipDelete(key, o) {
  return api_("slipDelete", key, true, function () {
    o = o || {};
    const row = Number(o.row);
    if (!(row >= 2)) throw userErr_("แถวไม่ถูกต้อง");
    withLock_(function () {
      const sh = paymentSheet_();
      if (row > sh.getLastRow())
        throw userErr_("ไม่พบรายการ (อาจถูกลบไปแล้ว) กรุณารีเฟรช");
      const url = String(
        sh.getRange(row, PAY.SLIP + 1).getValue() || "",
      ).trim();
      if (url !== String(o.url || ""))
        throw userErr_("รายการถูกแก้ไขไปแล้ว กรุณารีเฟรชแล้วลองใหม่");
      sh.deleteRow(row);
      audit_("dashboard", "web", "[สลีป] ลบแถว " + row);
    });
    return slipList_();
  });
}

// เก็บไว้เพื่อความเข้ากันได้กับโค้ดเดิม
function apiPaymentSave(key, o) {
  return api_("paymentSave", key, true, function () {
    o = o || {};
    saveSlip_(
      { id: o.id, amount: o.amount, note: o.note, data: o.slipData },
      "dashboard",
    );
    return { success: true, message: "บันทึกการชำระเงินเรียบร้อยแล้ว" };
  });
}
