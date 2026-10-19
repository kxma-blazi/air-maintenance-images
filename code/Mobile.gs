// ==================== Mobile.gs — API สำหรับหน้าเว็บแอปมือถือ (App.html) v4 ====================
// ใช้ฟังก์ชันเดิม (getAirs_, getLogs_, addLog_, setStatus_, api_, clean_ ฯลฯ) + Photos.gs

// สถิติการล้างต่อเครื่อง จากประวัติจริง (logs เรียงใหม่→เก่า)
function washStats_(logs) {
  const now = Date.now();
  const by = {};
  logs.forEach(function (l) {
    if (l.type !== "ล้างแอร์") return;
    (by[l.id] = by[l.id] || []).push(l);
  });
  const out = {};
  Object.keys(by).forEach(function (id) {
    const a = by[id];
    let cost = 0;
    a.forEach(function (x) {
      cost += x.cost;
    });
    out[id] = {
      n: a.length,
      last: fmt_(a[0].date),
      since: Math.max(0, Math.floor((now - a[0].date.getTime()) / 86400000)),
      avg:
        a.length >= 2
          ? Math.round(
              (a[0].date - a[a.length - 1].date) / 86400000 / (a.length - 1),
            )
          : null,
      cost: cost,
      tech: a[0].tech,
    };
  });
  return out;
}

function mobileData_() {
  const all = getLogs_();
  const ws = washStats_(all);
  const logs = all.slice(0, 400).map(function (l) {
    return {
      row: l.row,
      date: fmt_(l.date),
      id: l.id,
      type: l.type,
      detail: l.detail,
      tech: l.tech,
      cost: l.cost,
      photos: l.photos,
    };
  });
  let techs = [];
  try {
    techs = techListData_().techs.map(function (t) {
      return t.name;
    });
  } catch (e) {}
  const airs = getAirs_().map(function (a) {
    return Object.assign({}, a, { w: ws[a.id] || null });
  });
  return {
    updated: Utilities.formatDate(new Date(), TZ, "dd/MM HH:mm"),
    airs: airs,
    logs: logs,
    techs: techs,
    pending: pendingList_(),
    statuses: STATUSES,
    gh: ghReady_(),
    imgBase: ghRawBase_(),
    imgv: imgVer_(),
  };
}

function apiMobile(key) {
  return api_("mobile", key, false, mobileData_);
}

// o = { id, type: wash|issue|fix|status, cost, tech, detail, status, date: "yyyy-MM-dd", photos: [url], force }
function apiAction(key, o) {
  return api_("action", key, true, function () {
    o = o || {};
    ACTOR_ = "dashboard";
    const id = normId_(clean_(o.id, 20));
    const air = getAirs_().filter(function (x) {
      return x.id === id;
    })[0];
    if (!air) throw userErr_("ไม่พบเครื่อง " + id);

    let when = new Date();
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(o.date || ""))) {
      const p = o.date.split("-");
      when = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
    }
    if (when.getTime() > Date.now() + 86400000)
      throw userErr_("วันที่เป็นอนาคต กรุณาตรวจสอบ");
    if (when.getFullYear() < 2015) throw userErr_("วันที่ไม่ถูกต้อง");

    const cost = cleanNum_(o.cost, 0, 1000000, "ค่าใช้จ่าย", false) || "";
    const tech = clean_(o.tech, 80);
    const detail = clean_(o.detail, 200);
    const note = "บันทึกจากแอปมือถือ";
    const base = ghRawBase_() + "logs/" + id + "/";
    const photos = (Array.isArray(o.photos) ? o.photos : [])
      .map(String)
      .filter(function (u) {
        return u.indexOf(base) === 0 && /^[A-Za-z0-9_\-\/.:]+$/.test(u);
      })
      .slice(0, 6);

    switch (o.type) {
      case "wash":
        if (!o.force) {
          const day = fmt_(when);
          const dup = getLogs_().some(function (l) {
            return l.id === id && l.type === "ล้างแอร์" && fmt_(l.date) === day;
          });
          if (dup)
            throw userErr_(
              "[DUP]เครื่องนี้มีบันทึกล้างแอร์วันที่ " +
                day +
                " อยู่แล้ว — กดบันทึกอีกครั้งถ้าต้องการบันทึกซ้ำ",
            );
        }
        addLog_({
          date: when,
          id: id,
          type: "ล้างแอร์",
          detail: detail,
          tech: tech,
          cost: cost,
          next: addMonths_(when, WASH_INTERVAL_MONTHS),
          note: note,
          photos: photos,
        });
        setStatus_(id, "ปกติ");
        break;
      case "issue":
        if (!detail) throw userErr_("ใส่อาการที่พบด้วย");
        addLog_({
          date: when,
          id: id,
          type: "แจ้งปัญหา",
          detail: detail,
          note: note,
          photos: photos,
        });
        setStatus_(id, "รอซ่อม");
        break;
      case "fix":
        addLog_({
          date: when,
          id: id,
          type: "ซ่อม",
          detail: detail || "ซ่อมเสร็จ",
          tech: tech,
          cost: cost,
          note: note,
          photos: photos,
        });
        setStatus_(id, "ปกติ");
        break;
      case "status": {
        const st = clean_(o.status, 30);
        if (STATUSES.indexOf(st) === -1)
          throw userErr_("สถานะต้องเป็น: " + STATUSES.join(" / "));
        setStatus_(id, st);
        break;
      }
      default:
        throw userErr_("ไม่รู้จักประเภทงาน");
    }
    audit_(
      "dashboard",
      "web",
      "[แอป] " +
        o.type +
        " " +
        id +
        (photos.length ? " +" + photos.length + " รูป" : ""),
    );
    SpreadsheetApp.flush();
    return mobileData_();
  });
}

// ลบบันทึกที่กรอกผิด (ล้างแถวในชีต) o = { row, id, type, date: "dd/MM/yyyy" } — ตรวจแถวตรงกันก่อนลบ
function apiLogDelete(key, o) {
  return api_("logDelete", key, true, function () {
    o = o || {};
    const row = Number(o.row);
    if (!(row >= LOG_FIRST_ROW)) throw userErr_("แถวไม่ถูกต้อง");
    withLock_(function () {
      const sh = logSheet_();
      if (row > sh.getLastRow())
        throw userErr_("ไม่พบรายการ (อาจถูกลบไปแล้ว) กรุณารีเฟรช");
      const r = sh.getRange(row, 1, 1, 9).getValues()[0];
      if (
        normId_(r[LOG.ID]) !== normId_(o.id) ||
        String(r[LOG.TYPE]) !== String(o.type) ||
        fmt_(r[LOG.DATE]) !== String(o.date)
      )
        throw userErr_("รายการถูกแก้ไขไปแล้ว กรุณารีเฟรชแล้วลองใหม่");
      sh.getRange(row, 1, 1, 9).clearContent();
      audit_(
        "dashboard",
        "web",
        "[แอป] ลบบันทึก " + o.type + " " + o.id + " " + o.date,
      );
    });
    SpreadsheetApp.flush();
    return mobileData_();
  });
}
