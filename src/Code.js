/**
 * ⚙️ CONFIGURATION: แก้ไขชื่อชีทหรือ Token ที่นี่ที่เดียวจบ
 */
const CONFIG = {
  SHEET: {
    MASTER: "Master",    
    SCHEDULE: "Schedule",
    LEAVES: "Leaves",
    LEAVES_SUM: "Leaves_Sum",
    EVENT: "Event"

    // MASTER: "TEST_Master",    
    // SCHEDULE: "TEST_Schedule",
    // LEAVES: "TEST_Leaves",
    // LEAVES_SUM: "TEST_Leaves_Sum",
    // EVENT: "TEST_Event"
  },
  LINE_ACCESS_TOKEN: "YOUR LINE TOKEN",
  TITLES: {
    INDEX: "ระบบจัดตารางนักดนตรี",
    MEMBER: "แจ้งคิวงานนักดนตรี",
    MYSCHEDULE: "ตารางงานของฉัน"
  },
    CELLS: {
    YEAR: "C1",
    MONTH: "D1",
    DATE_COLUMN: "A3:A33",
    OUTPUT_EVENING: "C3:C33",
    OUTPUT_NIGHT: "D3:D33",
    DATE_COLUMN_START: 3,
    DATE_COLUMN_INDEX: 1 // กะดึก Column E (Index 5) แต่ช่วง Range คือ C3:E33
  }
};

/**
 * 🛠 HELPER FUNCTIONS
 */
const getAppSheet = (nameKey) => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET[nameKey]);
const getDayNum = (val) => (val instanceof Date) ? val.getDate() : parseInt(val);

/**
 * 🚀 MAIN WEB APP
 */
function doGet(e) {
  try {
    const page = (e.parameter.page || 'index').toLowerCase();
    const pages = {
      'index': { file: 'Index', title: CONFIG.TITLES.INDEX },
      'member': { file: 'Member', title: CONFIG.TITLES.MEMBER },
      'myschedule': { file: 'MySchedule', title: CONFIG.TITLES.MYSCHEDULE }
    };

    const config = pages[page] || pages['index'];
    const template = HtmlService.createTemplateFromFile(config.file);

    return template.evaluate()
      .setTitle(config.title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput("❌ Error: " + err.toString());
  }
}

/**
 * 📅 SCHEDULE LOGIC
 */
function runAutoSchedule(year, month) {
  try {
    const sheet = getAppSheet('SCHEDULE');
    sheet.getRange(CONFIG.CELLS.YEAR).setValue(year);
    sheet.getRange(CONFIG.CELLS.MONTH).setValue(month);
    
    // ล้างเฉพาะคอลัมน์ C และ E (31 แถว)
    sheet.getRange(3, 3, 31, 1).clearContent(); 
    sheet.getRange(3, 5, 31, 1).clearContent(); 
    
    autoDraftSchedule();
    return `✅ จัดตาราง ${year}年 ${month}月 เรียบร้อยแล้ว!`;
  } catch (e) { return "❌ Error: " + e.toString(); }
}

function getScheduleData() {
  try {
    const sheet = getAppSheet('SCHEDULE');
    const year = sheet.getRange(CONFIG.CELLS.YEAR).getValue();
    const month = sheet.getRange(CONFIG.CELLS.MONTH).getValue();
    const daysInMonth = new Date(year, month, 0).getDate();
    const data = sheet.getRange(3, 1, 31, 6).getValues();

    return data.filter(row => {
      const dNum = getDayNum(row[0]);
      return !isNaN(dNum) && dNum > 0 && dNum <= daysInMonth;
    }).map(row => ({
      date: getDayNum(row[0]),
      day: row[1] ? row[1].toString().trim() : "",
      slot1: row[2] ? row[2].toString().trim() : "",
      dj1: row[3] ? row[3].toString().trim() : "",   
      slot2: row[4] ? row[4].toString().trim() : "", 
      dj2: row[5] ? row[5].toString().trim() : ""    
    }));
  } catch (e) { return []; }
}

function updateCell(row, col, value) {
  getAppSheet('SCHEDULE').getRange(row + 3, col).setValue(value);
  return "บันทึกแล้ว";
}

/**
 * 👤 MUSICIAN DATA
 */
function getMusicianList() {
  try {
    const sheet = getAppSheet('MASTER');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  } catch (e) { return ["ไม่พบรายชื่อ"]; }
}

function getPersonalSchedule(name, year, month) {
  try {
    const sheet = getAppSheet('SCHEDULE');
    const data = sheet.getRange(3, 1, 31, 5).getValues();
    const searchName = name.toString().trim();

    return data.filter(row => row[0] !== "").map(row => {
      const checkInShift = (cell) => cell?.toString().split(',').some(n => n.trim() === searchName);
      const hasShift1 = checkInShift(row[2]);
      const hasShift2 = checkInShift(row[4]);

      return (hasShift1 || hasShift2) ? {
        day: getDayNum(row[0]),
        dayOfWeek: row[1]?.toString() || "",
        shift1: hasShift1 ? "✓" : "",
        shift2: hasShift2 ? "✓" : ""
      } : null;
    }).filter(Boolean).sort((a, b) => a.day - b.day);
  } catch (e) { throw new Error(e.toString()); }
}

/**
 * 💾 DATA SAVING
 */
function saveBulkData(payload) {
  try {
    const { name, items, isNew, skills } = payload;
    const timestamp = new Date();
    const sheetLeaves = getAppSheet('LEAVES');

    if (isNew) {
      getAppSheet('MASTER').appendRow([name, skills["Voc"], skills["AcG"], skills["ElG"], skills["Bass"], skills["Drum"], skills["Keyboard"], "Active", 1]);
    }

    // --- [ส่วนที่ 1: เตรียมข้อมูลวันที่และเดือน] ---
    const targetDate = items[0].date.split("-");
    const yr = targetDate[0];
    const mo = targetDate[1];
    const daysInMonth = new Date(yr, mo, 0).getDate();

    // --- [ส่วนที่ 2: ลบข้อมูล "ลา" เดิมของคนนี้ในเดือนนี้ออกให้หมดก่อน] ---
    // เพื่อป้องกันการบันทึกซ้ำซ้อน และเพื่อให้เรา Update ข้อมูลใหม่ได้คลีนที่สุด
    const lastRow = sheetLeaves.getLastRow();
    if (lastRow > 1) {
      const allData = sheetLeaves.getRange(2, 1, lastRow - 1, 6).getValues();
      // ลบย้อนหลังเพื่อไม่ให้ Index เคลื่อน
      for (let i = allData.length - 1; i >= 0; i--) {
        const row = allData[i];
        if (row[3] === name && row[1] == yr && row[2] == mo) {
          sheetLeaves.deleteRow(i + 2);
        }
      }
    }

    // --- [ส่วนที่ 3: จัดการข้อมูลที่ได้รับจากหน้าเว็บ] ---
    const workItems = items.filter(item => item.mode === "Work");
    const leaveItems = items.filter(item => item.mode === "Leave");

    // 3.1 บันทึกวันที่ "แจ้งลา" (Leave)
    leaveItems.forEach(item => {
      const dy = item.date.split("-")[2];
      const shifts = (item.shift === "ลาทั้งคืน") ? ["ค่ำ", "ดึก"] : [item.shift];
      shifts.forEach(s => {
        sheetLeaves.appendRow([timestamp, yr, mo, name, dy, s, "แจ้งลาปกติ"]);
      });
    });

    // 3.2 คำนวณ "ลาอัตโนมัติ" (เฉพาะวันที่ไม่ได้กด "ลงงาน")
    // ถ้ามีการกดลงงานมาแม้แต่วันเดียว ระบบจะลงลาให้อัตโนมัติในวันที่เหลือ
    if (workItems.length > 0) {
      const workedDays = workItems.map(item => ({
        day: parseInt(item.date.split("-")[2]),
        shift: item.shift // เก็บไว้ว่าลงงานกะไหน
      }));

      for (let d = 1; d <= daysInMonth; d++) {
        // เช็คว่าวันนั้นๆ มีการลงงานกะไหนบ้าง
        const todayWork = workedDays.filter(w => w.day === d);
        const shiftsWorked = todayWork.map(w => w.shift === "ลงงานทั้งคืน" ? ["ค่ำ", "ดึก"] : [w.shift]).flat();
        
        // ถ้าวันนั้นไม่มีการลงงานกะไหนเลย หรือลงไม่ครบกะ ให้ลงลาในกะที่ว่าง
        ["ค่ำ", "ดึก"].forEach(s => {
          if (!shiftsWorked.includes(s)) {
            // เช็คก่อนว่าวันนั้นเราไม่ได้ "แจ้งลาปกติ" ไว้แล้ว (เพื่อไม่ให้บันทึกซ้ำ)
            const isAlreadyLeave = leaveItems.some(l => parseInt(l.date.split("-")[2]) === d && (l.shift === "ลาทั้งคืน" || l.shift === s));
            if (!isAlreadyLeave) {
              sheetLeaves.appendRow([timestamp, yr, mo, name, d, s, "ลาอัตโนมัติ"]);
            }
          }
        });
      }
    }

    return "✅ อัปเดตข้อมูลเรียบร้อยแล้ว!";
  } catch (err) { return "❌ Error: " + err.toString(); }
}

/**
 * 🤖 AUTOMATION
 */
function autoDraftSchedule() {
  const scheduleSheet = getAppSheet('SCHEDULE');
  const year = scheduleSheet.getRange(CONFIG.CELLS.YEAR).getValue();
  const month = scheduleSheet.getRange(CONFIG.CELLS.MONTH).getValue();
  const daysInMonth = new Date(year, month, 0).getDate();

  const musicians = getMusicianList();
  const leaveData = getAppSheet('LEAVES').getDataRange().getValues();
  const leaveMap = {};
  
  leaveData.slice(1).forEach(row => {
    if (row[1] == year && row[2] == month) {
      leaveMap[`${row[3]}-${row[4]}-${row[5]}`] = true;
    }
  });

  let finalSchedule = [];
  let prevEvening = "", prevNight = "";

  for (let day = 1; day <= daysInMonth; day++) {
    const canWorkEvening = musicians.filter(name => !leaveMap[`${name}-${day}-ค่ำ`]);
    const canWorkNight = musicians.filter(name => !leaveMap[`${name}-${day}-ดึก`]);

    let todayEvening = (prevNight && canWorkEvening.includes(prevNight)) ? prevNight : (canWorkEvening.filter(n => n !== prevEvening)[day % (canWorkEvening.length || 1)] || canWorkEvening[0] || "");
    let filteredNight = canWorkNight.filter(n => n !== todayEvening);
    let todayNight = (prevEvening && filteredNight.includes(prevEvening)) ? prevEvening : (filteredNight.filter(n => n !== prevNight)[0] || filteredNight[0] || "");

    finalSchedule.push([todayEvening, todayNight]);
    prevEvening = todayEvening; prevNight = todayNight;
  }

  scheduleSheet.getRange(3, 3, finalSchedule.length, 1).setValues(finalSchedule.map(r => [r[0]]));
  scheduleSheet.getRange(3, 5, finalSchedule.length, 1).setValues(finalSchedule.map(r => [r[1]]));
}

function onEdit(e) {
  const range = e.range;
  const sheetName = range.getSheet().getName();
  if (sheetName === CONFIG.SHEET.SCHEDULE && range.getRow() === 1 && range.getColumn() === 1 && range.getValue() === true) {
    SpreadsheetApp.getActiveSpreadsheet().toast("กำลังจัดตารางให้ครับ...", "ระบบอัตโนมัติ", 3);
    autoDraftSchedule();
    range.setValue(false);
    SpreadsheetApp.getUi().alert("✅ จัดตารางงานเรียบร้อยแล้วครับ!");
  }
}

/**
 * 💬 LINE BOT
 */
function doPost(e) {
  try {
    const event = JSON.parse(e.postData.contents).events[0];
    if (!event?.message || event.message.type !== "text") return;

    const userMessage = event.message.text.trim();
    const isLeave = userMessage.indexOf("ลา") === 0;
    const isEvent = (userMessage.indexOf("ล็อค") === 0 || userMessage.indexOf("Event") === 0);

    if (isLeave || isEvent) {
      const sheet = getAppSheet(isLeave ? 'LEAVES' : 'EVENT');
      const lines = userMessage.split('\n');
      let totalSaved = 0, summaryList = [], targetName = "-", firstCommand = "";

      lines.forEach(line => {
        const parts = line.trim().split(/\s+/);
        if (!["ลา", "ล็อค", "Event"].includes(parts[0])) return;
        if (!firstCommand) firstCommand = parts[0];
        targetName = parts[1] || "-";

        let now = new Date(), currYr = now.getFullYear(), currMo = now.getMonth() + 1, pending = [];

        for (let i = 2; i < parts.length; i++) {
          let item = parts[i];
          let m;
          if (m = item.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/)) {
            currYr = parseInt(m[1]); currMo = parseInt(m[2]);
            pending.push({ y: currYr, m: currMo, d: parseInt(m[3]) });
          } else if (m = item.match(/^(\d{1,2})[-ถึง](\d{1,2})$/)) {
            for (let d = parseInt(m[1]); d <= parseInt(m[2]); d++) pending.push({ y: currYr, m: currMo, d: d });
          } else if (m = item.match(/^(\d{1,2})$/)) {
            pending.push({ y: currYr, m: currMo, d: parseInt(m[1]) });
          } else if (item === "ค่ำ" || item === "ดึก") {
            pending.forEach(date => {
              sheet.appendRow([new Date(), date.y, date.m, targetName, date.d, item, "บันทึกผ่าน LINE"]);
              summaryList.push(`${date.y}/${date.m}/${date.d} (${item})`);
              totalSaved++;
            });
            pending = [];
          }
        }
      });
      if (totalSaved > 0) replyMessage(event.replyToken, `✅ บันทึกสำเร็จ ${totalSaved} รายการ\n👤 นักดนตรี: ${targetName}\n📅 ${summaryList.join("\n📅 ")}`);
    }
  } catch (err) {}
}

function replyMessage(token, text) {
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    "method": "post",
    "headers": { "Authorization": "Bearer " + CONFIG.LINE_ACCESS_TOKEN, "Content-Type": "application/json" },
    "payload": JSON.stringify({ "replyToken": token, "messages": [{ "type": "text", "text": text }] })
  });
}