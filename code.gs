/**
 * ================================================================
 * Ventilator Simulator - Backend (Google Apps Script)
 * เวอร์ชันปรับปรุง: v2
 *
 * สิ่งที่แก้ไขจากเวอร์ชันเดิม:
 * 1. เพิ่ม APP_SECRET กันคนนอกยิง POST ปลอมคะแนนเข้ามาตรงๆ
 * 2. เพิ่ม LockService กัน race condition เวลามีคนสอบพร้อมกัน
 * 3. Validate ข้อมูลที่ส่งเข้ามาก่อนบันทึกจริง
 * 4. แยกคะแนนเป็นคอลัมน์ตัวเลข (score, totalQuestions) แทนการเก็บ
 *    เป็น text "8 / 10" เพื่อให้ทำสูตร/Pivot/กราฟใน Sheet ได้
 * 5. เพิ่ม submissionId (UUID) ทุกแถว + endpoint สำหรับให้ frontend
 *    เรียกกลับมาเช็คว่าแถวถูกบันทึกจริงหรือไม่ (verify-after-write)
 *    เพื่อแก้ปัญหาเดิมที่ frontend ใช้ mode:'no-cors' แล้วไม่รู้ว่า
 *    บันทึกสำเร็จจริงหรือเปล่า
 * ================================================================
 */

// ⚠️ เปลี่ยนค่านี้เป็นรหัสของคุณเอง และต้องตรงกับค่า APP_SECRET ใน index.html
// นี่ไม่ใช่ระบบความปลอดภัยระดับสูง (เพราะฝั่ง client มองเห็นค่านี้ได้อยู่ดี)
// แต่ช่วยกันการยิง POST ปลอมแบบสุ่มๆ หรือบอทกวนระบบได้ในระดับหนึ่ง
var APP_SECRET = "VENT-SIM-2026-CHANGE-ME";

var SHEET_NAME = "ExamSummary";

var HEADERS = [
  "วัน-เวลา",
  "SubmissionID",
  "ชื่อ-สกุล",
  "หน่วยงาน",
  "โรงพยาบาล",
  "อายุงาน (ปี)",
  "โหมดการสอบ",
  "คะแนน",
  "เต็ม",
  "คะแนนที่ได้ (แสดงผล)",
  "รายละเอียดการตอบ"
];

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doGet รองรับ 2 อย่าง:
 * 1. ไม่มี parameter -> คืนค่า default settings (พฤติกรรมเดิม)
 * 2. ?action=checkSubmission&id=xxx -> เช็คว่า submissionId นี้ถูกบันทึก
 *    ลง Sheet จริงหรือไม่ (ใช้ตอน frontend verify-after-write)
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  if (params.action === "checkSubmission" && params.id) {
    return jsonOutput_(checkSubmission_(params.id));
  }

  var responseData = {
    status: "success",
    defaultSettings: { mode: "pcv", flowType: "decel", rr: 15, ti: 1.0, pi: 15, peep: 5, slope: 0 }
  };
  return jsonOutput_(responseData);
}

function checkSubmission_(submissionId) {
  try {
    var sheet = getOrCreateSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: "success", found: false };

    // อ่านเฉพาะคอลัมน์ SubmissionID (คอลัมน์ B) เพื่อความเร็ว
    var idRange = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < idRange.length; i++) {
      if (idRange[i][0] === submissionId) {
        return { status: "success", found: true, row: i + 2 };
      }
    }
    return { status: "success", found: false };
  } catch (err) {
    return { status: "error", message: err.message };
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var hasLock = lock.tryLock(15000); // รอสูงสุด 15 วินาที
    if (!hasLock) {
      return jsonOutput_({ status: "error", message: "ระบบไม่ว่าง กรุณาลองใหม่อีกครั้ง (lock timeout)" });
    }

    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput_({ status: "error", message: "ไม่พบข้อมูลที่ส่งมา" });
    }

    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonOutput_({ status: "error", message: "รูปแบบข้อมูลไม่ถูกต้อง (JSON parse error)" });
    }

    // ---- 1. ตรวจสอบ secret กัน spam / ปลอมข้อมูล ----
    if (data.secret !== APP_SECRET) {
      return jsonOutput_({ status: "error", message: "ไม่ได้รับอนุญาต (invalid secret)" });
    }

    if (data.action !== "saveScore") {
      return jsonOutput_({ status: "error", message: "ไม่รู้จัก action นี้" });
    }

    // ---- 2. Validate ข้อมูลที่จำเป็น ----
    var requiredFields = ["name", "department", "hospital", "experience", "examType", "score", "submissionId"];
    for (var f = 0; f < requiredFields.length; f++) {
      var key = requiredFields[f];
      if (data[key] === undefined || data[key] === null || data[key] === "") {
        return jsonOutput_({ status: "error", message: "ข้อมูลไม่ครบ: " + key });
      }
    }

    var scoreNum = Number(data.score);
    var totalNum = Number(data.totalQuestions || 10);
    if (isNaN(scoreNum) || scoreNum < 0 || scoreNum > totalNum) {
      return jsonOutput_({ status: "error", message: "ค่าคะแนนไม่ถูกต้อง" });
    }

    var experienceNum = Number(data.experience);
    if (isNaN(experienceNum) || experienceNum < 0) {
      return jsonOutput_({ status: "error", message: "ค่าอายุงานไม่ถูกต้อง" });
    }

    // ---- 3. ป้องกันการบันทึกซ้ำ (ถ้า frontend ยิงซ้ำเพราะ retry) ----
    var existing = checkSubmission_(data.submissionId);
    if (existing.found) {
      return jsonOutput_({ status: "success", submissionId: data.submissionId, duplicate: true });
    }

    // ---- 4. บันทึกจริง ----
    var sheet = getOrCreateSheet_();
    sheet.appendRow([
      new Date(),
      data.submissionId,
      String(data.name).substring(0, 200),
      String(data.department).substring(0, 200),
      String(data.hospital).substring(0, 200),
      experienceNum,
      String(data.examType).substring(0, 100),
      scoreNum,
      totalNum,
      scoreNum + " / " + totalNum,
      String(data.details || "").substring(0, 10000)
    ]);

    return jsonOutput_({ status: "success", submissionId: data.submissionId, duplicate: false });

  } catch (error) {
    return jsonOutput_({ status: "error", message: error.message });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* ignore */ }
  }
}
