// ═══════════════════════════════════════════════════════════════════
// ASCG AI Survey 2026 — Google Apps Script (Backend)
// วางโค้ดนี้ใน Google Apps Script แล้ว Deploy เป็น Web App
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAME = "responses";

// ── รับข้อมูลจาก Survey App (POST) ──────────────────────────────
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // สร้าง Sheet + Header row ถ้ายังไม่มี
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        "timestamp", "id", "name", "unit", "level",
        "aiFreq",
        "aiTools",       // JSON array
        "workTypes",     // JSON array
        "skill_1", "skill_2", "skill_3", "skill_4", "skill_5",
        "impact_1", "impact_2", "impact_3", "impact_4", "impact_5",
        "avgSkill", "avgImpact",
        "barriers",      // JSON array
        "supportNeeds",  // JSON array
        "trainingFormat",
        "problemName", "problemDetail",
        "score_business", "score_feasibility", "score_speed", "score_crossGroup",
        "totalScore", "priority",
        "suggestions"
      ];
      sheet.appendRow(headers);

      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#1E4DA1").setFontColor("#FFFFFF").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents);

    // คำนวณ avgSkill, avgImpact
    const skillVals = data.skills.filter(v => v > 0);
    const impactVals = data.impacts.filter(v => v > 0);
    const avgSkill = skillVals.length ? (skillVals.reduce((a, b) => a + b, 0) / skillVals.length).toFixed(2) : 0;
    const avgImpact = impactVals.length ? (impactVals.reduce((a, b) => a + b, 0) / impactVals.length).toFixed(2) : 0;

    // คำนวณ totalScore (weighted)
    const s = data.scoring;
    const totalScore = (s.business * 0.4 + s.feasibility * 0.3 + s.speed * 0.2 + s.crossGroup * 0.1).toFixed(2);
    const priority = parseFloat(totalScore) >= 4.0 ? "P1 - Pilot ทันที" : parseFloat(totalScore) >= 2.5 ? "P2 - คิวถัดไป" : "P3 - บันทึกไว้";

    const row = [
      new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" }),
      data.id,
      data.name || "(ไม่ระบุ)",
      data.unit,
      data.level,
      data.aiFreq,
      JSON.stringify(data.aiTools),
      JSON.stringify(data.workTypes),
      data.skills[0], data.skills[1], data.skills[2], data.skills[3], data.skills[4],
      data.impacts[0], data.impacts[1], data.impacts[2], data.impacts[3], data.impacts[4],
      avgSkill, avgImpact,
      JSON.stringify(data.barriers),
      JSON.stringify(data.supportNeeds),
      data.trainingFormat,
      data.problemName,
      data.problemDetail,
      s.business, s.feasibility, s.speed, s.crossGroup,
      totalScore, priority,
      data.suggestions
    ];

    sheet.appendRow(row);

    // Auto-resize columns
    sheet.autoResizeColumns(1, row.length);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, id: data.id }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({ "Access-Control-Allow-Origin": "*" });

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({ "Access-Control-Allow-Origin": "*" });
  }
}

// ── ส่งข้อมูลกลับ Dashboard (GET) ───────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: [] }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeaders({ "Access-Control-Allow-Origin": "*" });
    }

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const records = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        // Parse JSON arrays back
        if (["aiTools", "workTypes", "barriers", "supportNeeds"].includes(h)) {
          try { obj[h] = JSON.parse(row[i]); } catch { obj[h] = []; }
        } else {
          obj[h] = row[i];
        }
      });
      return obj;
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: records, count: records.length }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({ "Access-Control-Allow-Origin": "*" });

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({ "Access-Control-Allow-Origin": "*" });
  }
}

// ── CORS preflight ───────────────────────────────────────────────
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeaders({
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });
}
