// ==================== Repair.gs — ซ่อมข้อมูลชีตที่เพี้ยน (รันครั้งเดียวจาก Editor) ====================
// แก้ 4 เรื่อง:
//  1) ช่อง "ค่าใช้จ่าย" ในประวัติซ่อมที่ถูกจัดรูปแบบเป็นวันที่ (ทำให้ 480 กลายเป็น 24/04/1901 และยอดรวมติดลบหลายล้านล้าน)
//  2) วันที่ที่ปีผิด (เช่น 16/10/1026 → 16/10/2026)
//  3) สูตรชีต "เครื่องแอร์" คอลัมน์ นัดครั้งถัดไป / จำนวนครั้ง / ค่าใช้จ่ายรวม ที่อ้างผิดคอลัมน์หรือถูกพิมพ์ทับ
//  4) บังคับรูปแบบช่องค่าใช้จ่ายเป็นตัวเลขทั้งคอลัมน์ กันเกิดซ้ำ
// สำรองไฟล์ก่อนแก้ให้อัตโนมัติ
function repairData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bak =
    ss.getName() +
    " - ก่อนซ่อมข้อมูล " +
    Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HHmm");
  DriveApp.getFileById(ss.getId()).makeCopy(bak);
  const out = ["สำรองเป็น '" + bak + "' แล้ว"];

  // ---------- 1+2+4) ชีตประวัติซ่อม ----------
  const lg = logSheet_();
  const n = lg.getLastRow() - LOG_FIRST_ROW + 1;
  let fixedCost = 0,
    fixedDate = 0;
  lg.getRange(LOG_FIRST_ROW, 6, 1009, 1).setNumberFormat("#,##0;(#,##0);-"); // แถว 4–1012
  if (n > 0) {
    const a = lg.getRange(LOG_FIRST_ROW, 1, n, 1).getValues();
    const f = lg.getRange(LOG_FIRST_ROW, 6, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      const row = LOG_FIRST_ROW + i;
      if (f[i][0] instanceof Date) {
        lg.getRange(row, 6).setValue(costNum_(f[i][0]));
        fixedCost++;
      }
      if (a[i][0] instanceof Date) {
        const p = Utilities.formatDate(a[i][0], TZ, "yyyy-MM-dd").split("-");
        const y = Number(p[0]);
        if (y >= 1000 && y < 2000) {
          lg.getRange(row, 1)
            .setValue(
              new Date(y + 1000, Number(p[1]) - 1, Number(p[2]), 12, 0, 0),
            )
            .setNumberFormat("dd/mm/yyyy");
          fixedDate++;
        }
      }
    }
  }
  out.push(
    "ประวัติซ่อม: แก้ค่าใช้จ่ายที่เป็นวันที่ " +
      fixedCost +
      " แถว | แก้ปีผิด " +
      fixedDate +
      " แถว",
  );

  // ---------- 3) ชีตเครื่องแอร์: เขียนสูตรใหม่ให้ถูกคอลัมน์ ----------
  const sh = sheet_(SHEET_NAME);
  const m = sh.getLastRow() - FIRST_DATA_ROW + 1;
  if (m > 0) {
    const ids = sh.getRange(FIRST_DATA_ROW, 1, m, 1).getValues();
    const L = "'" + LOG_SHEET + "'!";
    let fixedRows = 0;
    ids.forEach(function (r, i) {
      if (!String(r[0]).trim()) return;
      const row = FIRST_DATA_ROW + i;
      const next =
        "MAXIFS(" + L + "$G$4:$G$1021," + L + "$B$4:$B$1021,$A" + row + ")";
      sh.getRange(row, COL.NEXT_DATE + 1, 1, 3).setFormulas([
        [
          "=IF($A" + row + '="","",IF(' + next + '=0,"",' + next + "))",
          "=IF($A" +
            row +
            '="","",COUNTIF(' +
            L +
            "$B$4:$B$1021,$A" +
            row +
            "))",
          "=IF($A" +
            row +
            '="","",SUMIF(' +
            L +
            "$B$4:$B$1021,$A" +
            row +
            "," +
            L +
            "$F$4:$F$1021))",
        ],
      ]);
      sh.getRange(row, COL.NEXT_DATE + 1).setNumberFormat("dd/mm/yyyy");
      fixedRows++;
    });
    out.push(
      "เครื่องแอร์: เขียนสูตรใหม่ " +
        fixedRows +
        " แถว (นัดถัดไป/จำนวนครั้ง/ค่าใช้จ่ายรวม)",
    );
  }

  SpreadsheetApp.flush();
  Logger.log(out.join("\n"));
  return out.join("\n");
}
