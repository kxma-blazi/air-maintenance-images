// ==================== Photos.gs — อัปโหลดรูปขึ้น GitHub + อัปเกรดชีต (v4) ====================
// Script properties ที่ต้องเพิ่ม:
//  GITHUB_TOKEN  =
//  GITHUB_REPO   =
//  GITHUB_BRANCH = (ไม่บังคับ) ค่าเริ่มต้น main
const GH_DEFAULT_REPO = "kxma-blazi/air-maintenance-images";
const PHOTO_KINDS = ["before", "after", "issue", "cover"];
const PHOTO_MAX_B64 = 2500000; // ~1.8MB หลังถอดรหัส (ฝั่งมือถือย่อรูปเหลือ ~150-400KB อยู่แล้ว)

function ghRepo_() {
  return prop_("GITHUB_REPO") || GH_DEFAULT_REPO;
}
function ghBranch_() {
  return prop_("GITHUB_BRANCH") || "main";
}
function ghRawBase_() {
  return (
    "https://raw.githubusercontent.com/" + ghRepo_() + "/" + ghBranch_() + "/"
  );
}
function ghReady_() {
  return !!prop_("GITHUB_TOKEN");
}

function ghHeaders_() {
  const t = prop_("GITHUB_TOKEN");
  if (!t)
    throw userErr_(
      "ยังไม่ได้ตั้ง GITHUB_TOKEN ใน Script properties (อัปโหลดรูปไม่ได้)",
    );
  return {
    Authorization: "Bearer " + t,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
  };
}

function ghUrl_(path) {
  return (
    "https://api.github.com/repos/" +
    ghRepo_() +
    "/contents/" +
    path.split("/").map(encodeURIComponent).join("/")
  );
}

function ghFail_(code, body) {
  let msg = "";
  try {
    msg = JSON.parse(body).message || "";
  } catch (e) {}
  if (code === 401)
    return userErr_(
      "GitHub token ไม่ถูกต้องหรือหมดอายุ — สร้างใหม่แล้วอัปเดต GITHUB_TOKEN",
    );
  if (code === 403)
    return userErr_("GitHub ปฏิเสธ (สิทธิ์ไม่พอหรือโดนจำกัดความถี่): " + msg);
  if (code === 404)
    return userErr_(
      "ไม่พบ repo " + ghRepo_() + " หรือ token ไม่มีสิทธิ์เข้าถึง repo นี้",
    );
  return new Error("GitHub " + code + ": " + msg);
}

// อัปโหลด/เขียนทับไฟล์ใน repo แล้วคืน URL ดิบ (raw)
function ghPut_(path, b64, message, overwrite) {
  const h = ghHeaders_();
  const url = ghUrl_(path);
  const body = { message: message, content: b64, branch: ghBranch_() };
  if (overwrite) {
    const g = UrlFetchApp.fetch(
      url + "?ref=" + encodeURIComponent(ghBranch_()),
      {
        headers: h,
        muteHttpExceptions: true,
      },
    );
    if (g.getResponseCode() === 200)
      body.sha = JSON.parse(g.getContentText()).sha;
    else if (g.getResponseCode() !== 404)
      throw ghFail_(g.getResponseCode(), g.getContentText());
  }
  let last;
  for (let i = 0; i < 2; i++) {
    const r = UrlFetchApp.fetch(url, {
      method: "put",
      contentType: "application/json",
      headers: h,
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const c = r.getResponseCode();
    if (c === 200 || c === 201) return ghRawBase_() + path;
    last = ghFail_(c, r.getContentText());
    if (c < 500 && c !== 409) break;
    Utilities.sleep(700);
  }
  throw last;
}

function ghCheck_() {
  const r = UrlFetchApp.fetch("https://api.github.com/repos/" + ghRepo_(), {
    headers: ghHeaders_(),
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200)
    throw ghFail_(r.getResponseCode(), r.getContentText());
  return ghRepo_() + "@" + ghBranch_();
}

// รันจาก Editor เพื่อทดสอบการเชื่อมต่อ GitHub
function testGithub() {
  Logger.log("GitHub OK: " + ghCheck_());
}

function imgVer_() {
  try {
    return JSON.parse(prop_("IMG_VER") || "{}");
  } catch (e) {
    return {};
  }
}

// o = { id, kind: before|after|issue|cover, data: "data:image/jpeg;base64,..." }
function apiPhotoUpload(key, o) {
  return api_("photo", key, true, function () {
    o = o || {};
    const id = normId_(clean_(o.id, 20));
    if (!/^[A-Z0-9][A-Z0-9-]{1,19}$/.test(id))
      throw userErr_("รหัสเครื่องไม่ถูกต้อง");
    requireAirUser_(id);
    const kind = String(o.kind || "");
    if (PHOTO_KINDS.indexOf(kind) === -1) throw userErr_("ประเภทรูปไม่ถูกต้อง");
    const b64 = String(o.data || "").replace(/^data:image\/jpeg;base64,/, "");
    if (!b64 || b64.length > PHOTO_MAX_B64)
      throw userErr_("ไฟล์รูปใหญ่เกินไป ลองถ่ายใหม่");
    if (!/^[A-Za-z0-9+\/]+=*$/.test(b64) || b64.indexOf("/9j/") !== 0)
      throw userErr_("รองรับเฉพาะรูป JPEG");

    if (kind === "cover") {
      const url = ghPut_(
        "assets/" + id + ".jpg",
        b64,
        "รูปเครื่อง " + id,
        true,
      );
      const v = imgVer_();
      v[id] = Date.now();
      PropertiesService.getScriptProperties().setProperty(
        "IMG_VER",
        JSON.stringify(v),
      );
      audit_("dashboard", "web", "[รูป] เปลี่ยนรูปเครื่อง " + id);
      return { url: url + "?v=" + v[id], ver: v[id] };
    }
    const stamp = Utilities.formatDate(new Date(), TZ, "yyyyMMdd-HHmmss");
    const rnd = Utilities.getUuid().replace(/-/g, "").substring(0, 4);
    const path = "logs/" + id + "/" + stamp + "-" + kind + "-" + rnd + ".jpg";
    const url = ghPut_(path, b64, "รูป " + kind + " " + id, false);
    audit_("dashboard", "web", "[รูป] " + kind + " " + id);
    return { url: url };
  });
}

function requireAirUser_(id) {
  const a = getAirs_().filter(function (x) {
    return x.id === id;
  })[0];
  if (!a) throw userErr_("ไม่พบเครื่อง " + id);
  return a;
}

// ชีตประวัติซ่อม: บังคับให้มีคอลัมน์ I (รูปถ่าย) พร้อมหัวตาราง
function logSheet_() {
  const sh = sheet_(LOG_SHEET);
  if (sh.getMaxColumns() < 9)
    sh.insertColumnsAfter(sh.getMaxColumns(), 9 - sh.getMaxColumns());
  const h = sh.getRange(LOG_FIRST_ROW - 1, 9);
  if (!h.getValue()) {
    sh.getRange(LOG_FIRST_ROW - 1, 8).copyTo(
      h,
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false,
    );
    h.setValue("รูปถ่าย (ลิงก์)");
    sh.setColumnWidth(9, 240);
  }
  return sh;
}

// รันครั้งเดียวจาก Editor หลังวางโค้ดใหม่: แก้ช่วงสูตร/ดรอปดาวน์ที่เล็กเกินไป + เพิ่มคอลัมน์รูป
function upgradeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];
  try {
    const lg = logSheet_();
    const listIds = SpreadsheetApp.newDataValidation()
      .requireValueInRange(
        ss.getSheetByName(SHEET_NAME).getRange("A5:A200"),
        true,
      )
      .setAllowInvalid(true)
      .build();
    lg.getRange("B4:B1012").setDataValidation(listIds);
    const types = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["ล้างแอร์", "เติมน้ำยา", "ซ่อม", "แจ้งปัญหา", "ตรวจเช็ค", "อื่นๆ"],
        true,
      )
      .setAllowInvalid(true)
      .build();
    lg.getRange("C4:C1012").setDataValidation(types);
    out.push("✅ ประวัติซ่อม: คอลัมน์รูป + ดรอปดาวน์ครอบคลุมถึงแถว 1012");
  } catch (e) {
    out.push("❌ ประวัติซ่อม: " + e.message);
  }
  try {
    const sm = ss.getSheetByName("สรุป");
    if (sm) {
      for (let r = 4; r <= 8; r++)
        sm.getRange(r, 2).setFormula(
          "=COUNTIF('" + SHEET_NAME + "'!$H$5:$H$200,A" + r + ")",
        );
      out.push(
        "✅ สรุป: นับสถานะถึงแถว 200 (เดิมนับแค่แถว 13 ทำให้ AC-10 ขึ้นไปไม่ถูกนับ)",
      );
    }
  } catch (e) {
    out.push("❌ สรุป: " + e.message);
  }
  Logger.log(out.join("\n"));
}

// ==================== ฟังก์ชันจัดระเบียบรูปอัตโนมัติขึ้น GitHub ====================
// รันฟังก์ชันนี้ครั้งเดียวจาก Editor เพื่อย้ายรูป AC-01.jpg... ใน Root เข้าโฟลเดอร์ assets/
function syncGithubAssets() {
  if (!ghReady_())
    throw new Error("ยังไม่ได้ตั้ง GITHUB_TOKEN ใน Script properties");
  const h = ghHeaders_();
  const repo = ghRepo_();
  const branch = ghBranch_();

  // 1. ดึงรายการไฟล์ใน Root ของ Repo
  const res = UrlFetchApp.fetch(
    "https://api.github.com/repos/" + repo + "/contents?ref=" + branch,
    {
      headers: h,
      muteHttpExceptions: true,
    },
  );
  if (res.getResponseCode() !== 200)
    throw new Error("GitHub Error: " + res.getContentText());

  const files = JSON.parse(res.getContentText());
  let moved = 0;

  files.forEach(function (f) {
    // ค้นหาไฟล์รูป AC-xx.jpg ที่อยู่ใน Root
    if (f.type === "file" && /^AC-\d+\.jpg$/i.test(f.name)) {
      Logger.log("กำลังย้าย " + f.name + " ไปยัง assets/" + f.name + "...");

      const fileRes = UrlFetchApp.fetch(f.url, {
        headers: h,
        muteHttpExceptions: true,
      });
      if (fileRes.getResponseCode() === 200) {
        const fileJson = JSON.parse(fileRes.getContentText());
        const b64 = fileJson.content.replace(/\s+/g, "");

        // บันทึกไฟล์ใหม่เข้าไปที่ assets/AC-xx.jpg
        ghPut_(
          "assets/" + f.name,
          b64,
          "Automated sync: move " + f.name + " to assets/",
          true,
        );

        // ลบไฟล์เดิมออกจาก Root
        UrlFetchApp.fetch(f.url, {
          method: "delete",
          headers: h,
          payload: JSON.stringify({
            message: "Automated cleanup: remove " + f.name + " from root",
            sha: f.sha,
            branch: branch,
          }),
          muteHttpExceptions: true,
        });
        moved++;
      }
    }
  });

  Logger.log(
    "✅ ซิงค์ย้ายรูปภาพเข้า assets/ เรียบร้อยแล้วรวม " + moved + " รายการ",
  );
  return "ซิงค์ย้ายรูปสำเร็จ " + moved + " รายการ";
}
