// ==========================================
// File: Pics.gs — ตรวจ/แก้/ซิงก์รูปที่ไม่ตรงรหัส (Pro Max)
// ==========================================
const PIC_AUDIT_SHEET_NAME_ =
  (typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.PIC_AUDIT) ||
  "ตรวจรูป";

function picAuditSheetName_() {
  return (
    (typeof CONFIG !== "undefined" &&
      CONFIG.SHEETS &&
      CONFIG.SHEETS.PIC_AUDIT) ||
    PIC_AUDIT_SHEET_NAME_ ||
    "ตรวจรูป"
  );
}

/**
 * สร้าง URL พรีวิวรูปภาพหลักของรหัสอุปกรณ์ให้ตรงกับ picBase_() ของระบบ
 */
function auditImageUrl_(code) {
  var cleanCode = String(code || "").trim();
  if (!cleanCode) return "";
  var base = typeof picBase_ === "function" ? picBase_() : ghRawBase_();
  var enc = encodeURIComponent(cleanCode);
  return base + enc + "/" + enc + ".jpg";
}

/**
 * 1) สร้างชีต "ตรวจรูป": รหัส | ชื่อ | รูปที่ระบบแสดงตอนนี้ | (กรอก) ไฟล์รูปที่ถูกต้อง
 */
function buildPicAudit() {
  ownerOnly_();
  var ss = getSS_();
  var C = CONFIG.COL;
  var src = ss.getSheetByName(CONFIG.SHEETS.ASSETS);
  if (!src || src.getLastRow() < 2) {
    throw new Error("ไม่พบข้อมูลในชีต " + CONFIG.SHEETS.ASSETS);
  }

  var lastCol = Math.max(src.getLastColumn(), 16);
  var rows = src
    .getRange(2, 1, src.getLastRow() - 1, lastCol)
    .getDisplayValues();

  var sheetName = picAuditSheetName_();
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clear();
  sh.getRange(1, 1, 1, 4)
    .setValues([
      [
        "รหัสในระบบ",
        "ชื่อรายการ",
        "รูปที่ระบบแสดงตอนนี้",
        "รูปที่ถูกต้อง = ไฟล์ของรหัส (กรอกเฉพาะแถวที่ผิด)",
      ],
    ])
    .setFontWeight("bold")
    .setBackground("#e8eef7");
  sh.setFrozenRows(1);

  var data = [];
  var forms = [];
  var seen = {};

  rows.forEach(function (r) {
    var code = String(r[C.ID] || "").trim();
    if (!code || seen[code]) return;
    seen[code] = true;
    data.push([code, String(r[C.NAME] || "").trim()]);
    forms.push(['=IMAGE("' + auditImageUrl_(code) + '")']);
  });

  if (!data.length) {
    Logger.log("ไม่มีรหัสอุปกรณ์ในชีต " + CONFIG.SHEETS.ASSETS);
    return;
  }

  sh.getRange(2, 1, data.length, 2).setValues(data);
  sh.getRange(2, 3, data.length, 1).setFormulas(forms);
  sh.getRange(2, 4, data.length, 1).setNumberFormat("@");
  sh.setRowHeights(2, data.length, 90);
  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 240);

  Logger.log("สร้างชีต " + sheetName + " สำเร็จ " + data.length + " แถว");
}

/**
 * 2) อ่านคอลัมน์ D แล้วย้าย/สลับรูปหลักบน GitHub ใน commit เดียว (ย้อนกลับได้จาก Git History)
 *    เช่น แถว TOL-027 ใส่ D = TOL-023 แปลว่า นำรูปหลักของ TOL-023 มาใส่ให้ TOL-027
 */
function applyPicMap() {
  ownerOnly_();
  var ss = getSS_();
  var gh = CONFIG.GITHUB;
  var sheetName = picAuditSheetName_();
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error("ยังไม่มีชีต " + sheetName + " (รัน buildPicAudit ก่อน)");
  }

  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues();

  var ref = ghApi_("get", "/git/ref/heads/" + encodeURIComponent(gh.BRANCH));
  var commitSha = ref.object.sha;
  var commit = ghApi_("get", "/git/commits/" + commitSha);
  var tree = ghApi_("get", "/git/trees/" + commit.tree.sha + "?recursive=1");

  // เก็บเฉพาะรูปหลักของแต่ละรหัส: โฟลเดอร์/โฟลเดอร์.jpg
  var blob = {};
  (tree.tree || []).forEach(function (t) {
    var m = /^([^\/]+)\/([^\/]+)\.jpg$/i.exec(t.path);
    if (t.type === "blob" && m && m[1].trim() === m[2].trim()) {
      blob[m[1].trim()] = t.sha;
    }
  });

  var entries = [];
  var log = [];
  var seenTo = {};

  v.forEach(function (r, idx) {
    var rowNum = idx + 2;
    var to = String(r[0] || "").trim();
    var from = String(r[3] || "").trim();
    if (!to || !from || from === to) return;

    if (seenTo[to]) {
      throw new Error("รหัสปลายทางซ้ำ: " + to + " (แถว " + rowNum + ")");
    }
    seenTo[to] = true;

    var sha = blob[from];
    if (!sha) {
      throw new Error(
        "ไม่พบไฟล์ " +
          from +
          "/" +
          from +
          ".jpg (แถว " +
          rowNum +
          " / " +
          to +
          ")",
      );
    }

    entries.push({
      path: to + "/" + to + ".jpg",
      mode: "100644",
      type: "blob",
      sha: sha,
    });
    log.push(from + " → " + to);
  });

  if (!entries.length) {
    Logger.log("ไม่มีแถวที่กรอกคอลัมน์ D");
    return;
  }

  var nt = ghApi_("post", "/git/trees", {
    base_tree: commit.tree.sha,
    tree: entries,
  });
  var nc = ghApi_("post", "/git/commits", {
    message: "Fix pic mapping (" + entries.length + " files)",
    tree: nt.sha,
    parents: [commitSha],
  });
  ghApi_("patch", "/git/refs/heads/" + encodeURIComponent(gh.BRANCH), {
    sha: nc.sha,
  });

  cacheBust_();
  logActivity_(
    "แก้รูปไม่ตรงรหัส",
    "อัปเดต " + entries.length + " ไฟล์: " + log.join(", "),
    "Owner (Editor)",
  );
  Logger.log("ย้ายรูปแล้ว " + entries.length + " ไฟล์\n" + log.join("\n"));
}

/**
 * 3) สแกนหารหัสอุปกรณ์ในชีต Asset ที่ยังไม่มีไฟล์รูปหลักบน GitHub
 */
function scanMissingPics() {
  ownerOnly_();
  var ss = getSS_();
  var C = CONFIG.COL;
  var src = ss.getSheetByName(CONFIG.SHEETS.ASSETS);
  if (!src || src.getLastRow() < 2) {
    throw new Error("ไม่พบข้อมูลในชีต " + CONFIG.SHEETS.ASSETS);
  }

  var rows = src.getRange(2, 1, src.getLastRow() - 1, 16).getDisplayValues();
  var files = typeof picIndex_ === "function" ? picIndex_() : [];
  var hasMain = {};

  files.forEach(function (p) {
    var i = p.indexOf("/");
    if (i < 0) return;
    var dir = p.substring(0, i).trim();
    var file = p.substring(i + 1).trim();
    if (file.toLowerCase() === (dir + ".jpg").toLowerCase()) {
      hasMain[dir.toLowerCase()] = true;
    }
  });

  var missing = [];
  rows.forEach(function (r) {
    var code = String(r[C.ID] || "").trim();
    if (!code) return;
    if (!hasMain[code.toLowerCase()]) {
      missing.push(code + " (" + String(r[C.NAME] || "-").trim() + ")");
    }
  });

  Logger.log(
    "รหัสที่ยังไม่มีรูปหลักบน GitHub (" +
      missing.length +
      " รายการ):\n" +
      (missing.length ? missing.join("\n") : "ครบทุกรายการ"),
  );
  return missing;
}

/**
 * 4) เรียก GitHub REST API พร้อมตรวจสิทธิ์/โทเค็นและสถานะ HTTP
 */
function ghApi_(method, path, body) {
  var token =
    PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) throw new Error("ยังไม่ได้ตั้ง GITHUB_TOKEN");
  var gh = CONFIG.GITHUB;
  var opt = {
    method: method,
    headers: {
      Authorization: "token " + token,
      Accept: "application/vnd.github+json",
    },
    muteHttpExceptions: true,
  };
  if (body) {
    opt.contentType = "application/json";
    opt.payload = JSON.stringify(body);
  }
  var r = UrlFetchApp.fetch(
    "https://api.github.com/repos/" + gh.OWNER + "/" + gh.REPO + path,
    opt,
  );
  var c = r.getResponseCode();
  if (c < 200 || c >= 300) {
    throw new Error(
      "GitHub " + c + ": " + r.getContentText().substring(0, 200),
    );
  }
  return JSON.parse(r.getContentText());
}

/**
 * 5) คัดลอกรูปจากโฟลเดอร์เก่า (เช่น 0_เครื่องวัด/TO-008/TO-008.jpg) มาไว้ที่ root (TO-008/TO-008.jpg)
 *    เฉพาะรหัสที่มีในชีต Asset และยังไม่มีรูปที่ root | ใช้ commit เดียว ย้อนกลับได้จาก Git History
 *    ไม่ลบไฟล์เดิม (คัดลอกเท่านั้น)
 *    ลองรัน copyLegacyPicsDryRun() ก่อนเพื่อดูรายการ แล้วค่อยรัน copyLegacyPics()
 */
function copyLegacyPicsDryRun() {
  return copyLegacyPics_(true);
}

function copyLegacyPics() {
  return copyLegacyPics_(false);
}

function copyLegacyPics_(dryRun) {
  ownerOnly_();
  var C = CONFIG.COL;
  var gh = CONFIG.GITHUB;
  var src = getSS_().getSheetByName(CONFIG.SHEETS.ASSETS);
  if (!src || src.getLastRow() < 2)
    throw new Error("ไม่พบข้อมูลในชีต " + CONFIG.SHEETS.ASSETS);

  var codes = {};
  src
    .getRange(2, C.ID + 1, src.getLastRow() - 1, 1)
    .getDisplayValues()
    .forEach(function (r) {
      var c = String(r[0] || "").trim();
      if (c) codes[picKey_(c)] = c;
    });

  var ref = ghApi_("get", "/git/ref/heads/" + encodeURIComponent(gh.BRANCH));
  var commitSha = ref.object.sha;
  var commit = ghApi_("get", "/git/commits/" + commitSha);
  var tree = ghApi_("get", "/git/trees/" + commit.tree.sha + "?recursive=1");
  if (tree.truncated)
    Logger.log("⚠️ GitHub tree ถูกตัด (truncated) ผลอาจไม่ครบ");

  // โฟลเดอร์ที่ root ที่มีรูปอยู่แล้ว
  var hasRoot = {};
  var legacy = {}; // key -> [{path, name, sha}]
  (tree.tree || []).forEach(function (t) {
    if (t.type !== "blob") return;
    var m1 = /^([^\/]+)\/([^\/]+\.(jpe?g|png|webp))$/i.exec(t.path);
    if (m1) {
      hasRoot[picKey_(m1[1])] = true;
      return;
    }
    var m2 = /^[^\/]+\/([^\/]+)\/([^\/]+\.(jpe?g|png|webp))$/i.exec(t.path);
    if (m2) {
      var k = picKey_(m2[1]);
      (legacy[k] = legacy[k] || []).push({
        path: t.path,
        name: m2[2],
        sha: t.sha,
      });
    }
  });

  var entries = [];
  var log = [];
  var seenPath = {};
  Object.keys(codes).forEach(function (k) {
    if (hasRoot[k] || !legacy[k]) return;
    var code = codes[k];
    legacy[k].forEach(function (f) {
      // คัดลอกเฉพาะรูปหลัก (ชื่อไฟล์ = รหัส) ไฟล์อื่นที่ชื่อไทยจะไม่ถูกคัดลอก
      // เพราะระบบตรวจรูปหลักจากชื่อไฟล์ และชื่อแบบ "(1_3)" จะทำให้ลำดับรูปสับสน
      var base = f.name.replace(/\.(jpe?g|png|webp)$/i, "");
      if (picKey_(base) !== k) return;
      if (seenPath[code + "/" + f.name]) return;
      seenPath[code + "/" + f.name] = true;
      entries.push({
        path: code + "/" + f.name,
        mode: "100644",
        type: "blob",
        sha: f.sha,
      });
      log.push(f.path + " → " + code + "/" + f.name);
    });
  });

  if (!entries.length) {
    Logger.log("ไม่มีรูปในโฟลเดอร์เก่าที่ต้องคัดลอก");
    return [];
  }
  if (dryRun) {
    Logger.log(
      "[DRY RUN] จะคัดลอก " + entries.length + " ไฟล์:\n" + log.join("\n"),
    );
    return log;
  }

  var nt = ghApi_("post", "/git/trees", {
    base_tree: commit.tree.sha,
    tree: entries,
  });
  var nc = ghApi_("post", "/git/commits", {
    message: "Copy legacy pics to root (" + entries.length + " files)",
    tree: nt.sha,
    parents: [commitSha],
  });
  ghApi_("patch", "/git/refs/heads/" + encodeURIComponent(gh.BRANCH), {
    sha: nc.sha,
  });

  cacheBust_();
  logActivity_(
    "คัดลอกรูปจากโฟลเดอร์เก่า",
    "คัดลอก " + entries.length + " ไฟล์",
    "Owner (Editor)",
  );
  Logger.log("คัดลอกแล้ว " + entries.length + " ไฟล์\n" + log.join("\n"));
  return log;
}

/**
 * 6) ทับรูปหลักที่ root ที่ "ผิดรหัส" ด้วยรูปจากโฟลเดอร์เก่า (0_..4_) ที่ตั้งชื่อตามรหัสถูกต้อง
 *    ทำเฉพาะรหัสที่มีในชีต, มีทั้ง root และโฟลเดอร์เก่า, และไฟล์ต่างกัน (เทียบ SHA)
 *    ใช้ commit เดียว ไม่ลบไฟล์ใดๆ รูปเดิมกู้คืนได้จาก Git History
 *    ลองรัน fixPicsFromLegacyDryRun() ก่อน แล้วค่อยรัน fixPicsFromLegacy()
 */
function fixPicsFromLegacyDryRun() {
  return fixPicsFromLegacy_(true);
}

function fixPicsFromLegacy() {
  return fixPicsFromLegacy_(false);
}

function fixPicsFromLegacy_(dryRun) {
  ownerOnly_();
  var C = CONFIG.COL;
  var gh = CONFIG.GITHUB;
  var src = getSS_().getSheetByName(CONFIG.SHEETS.ASSETS);
  if (!src || src.getLastRow() < 2)
    throw new Error("ไม่พบข้อมูลในชีต " + CONFIG.SHEETS.ASSETS);

  var codes = {};
  src
    .getRange(2, C.ID + 1, src.getLastRow() - 1, 1)
    .getDisplayValues()
    .forEach(function (r) {
      var c = String(r[0] || "").trim();
      if (c) codes[picKey_(c)] = c;
    });

  var ref = ghApi_("get", "/git/ref/heads/" + encodeURIComponent(gh.BRANCH));
  var commitSha = ref.object.sha;
  var commit = ghApi_("get", "/git/commits/" + commitSha);
  var tree = ghApi_("get", "/git/trees/" + commit.tree.sha + "?recursive=1");
  if (tree.truncated)
    Logger.log("⚠️ GitHub tree ถูกตัด (truncated) ผลอาจไม่ครบ");

  var rootMain = {}; // key -> {path, sha}
  var legacyMain = {}; // key -> {path, sha}
  (tree.tree || []).forEach(function (t) {
    if (t.type !== "blob") return;
    var m1 = /^([^\/]+)\/([^\/]+)\.(jpe?g)$/i.exec(t.path);
    if (m1 && picKey_(m1[1]) === picKey_(m1[2])) {
      rootMain[picKey_(m1[1])] = { path: t.path, sha: t.sha };
      return;
    }
    var m2 = /^[^\/]+\/([^\/]+)\/([^\/]+)\.(jpe?g)$/i.exec(t.path);
    if (m2 && picKey_(m2[1]) === picKey_(m2[2])) {
      legacyMain[picKey_(m2[1])] = { path: t.path, sha: t.sha };
    }
  });

  var entries = [];
  var log = [];
  Object.keys(codes).forEach(function (k) {
    var r = rootMain[k],
      l = legacyMain[k];
    if (!r || !l || r.sha === l.sha) return;
    entries.push({ path: r.path, mode: "100644", type: "blob", sha: l.sha });
    log.push(l.path + " ⇒ ทับ " + r.path);
  });

  if (!entries.length) {
    Logger.log("ไม่มีรูปที่ต้องทับ");
    return [];
  }
  if (dryRun) {
    Logger.log(
      "[DRY RUN] จะทับ " + entries.length + " ไฟล์:\n" + log.join("\n"),
    );
    return log;
  }

  var nt = ghApi_("post", "/git/trees", {
    base_tree: commit.tree.sha,
    tree: entries,
  });
  var nc = ghApi_("post", "/git/commits", {
    message: "Fix wrong root pics from legacy (" + entries.length + " files)",
    tree: nt.sha,
    parents: [commitSha],
  });
  ghApi_("patch", "/git/refs/heads/" + encodeURIComponent(gh.BRANCH), {
    sha: nc.sha,
  });

  cacheBust_();
  logActivity_(
    "แก้รูปผิดรหัสจากโฟลเดอร์เก่า",
    "ทับ " + entries.length + " ไฟล์: " + log.join(", "),
    "Owner (Editor)",
  );
  Logger.log("ทับแล้ว " + entries.length + " ไฟล์\n" + log.join("\n"));
  return log;
}

function cleanRootPicFolders() {
  ownerOnly_();
  var gh = CONFIG.GITHUB;
  var ref = ghApi_("get", "/git/ref/heads/" + encodeURIComponent(gh.BRANCH));
  var commitSha = ref.object.sha;
  var commit = ghApi_("get", "/git/commits/" + commitSha);
  var tree = ghApi_("get", "/git/trees/" + commit.tree.sha + "?recursive=1");

  var deleteMap = {};
  var folders = {};
  (tree.tree || []).forEach(function (t) {
    if (t.type !== "blob") return;
    var m = /^((?:BTT-TO|TO-)[^\/]+)\/.+$/i.exec(t.path);
    if (m) {
      deleteMap[t.path] = null;
      folders[m[1]] = true;
    } else if (t.path === "delete_root_folders.bat") {
      deleteMap[t.path] = null;
    }
  });

  var folderCount = Object.keys(folders).length;
  if (!Object.keys(deleteMap).length) {
    Logger.log("ไม่พบโฟลเดอร์หรือไฟล์ที่ต้องลบ");
    return;
  }

  picCommit_(deleteMap, "Clean root asset folders");
  Logger.log("ลบเรียบร้อย " + folderCount + " โฟลเดอร์");
}
