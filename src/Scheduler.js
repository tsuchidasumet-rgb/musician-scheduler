/**
 * ฟังก์ชันหลักในการจัดตารางงาน (เรียกจากหน้าเว็บ)
 */
function runAutoSchedule(year, month) {
  try {
    const sheet = getAppSheet('SCHEDULE');
    
    // 1. อัปเดตปีและเดือนใน Sheet ก่อนเริ่มคำนวณ
    sheet.getRange(CONFIG.CELLS.YEAR).setValue(year);
    sheet.getRange(CONFIG.CELLS.MONTH).setValue(month);

    // 2. ล้างข้อมูลเก่าในคอลัมน์นักดนตรี (C และ E)
    sheet.getRange("C3:C33").clearContent();
    sheet.getRange("E3:E33").clearContent();

    // 3. เริ่มฟังก์ชันประมวลผลจัดทีม
    autoDraftSchedule();
    
    return `✅ จัดตาราง ${year} 年 ${month} 月 เรียบร้อยแล้ว!`;
  } catch (e) {
    return "❌ เกิดข้อผิดพลาด: " + e.toString();
  }
}

/**
 * Logic การจัดทีมและหยอดข้อมูลลงชีท
 */
function autoDraftSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    schedule: getAppSheet('SCHEDULE'),
    master: getAppSheet('MASTER'),
    leaves: getAppSheet('LEAVES_SUM'),
    event: getAppSheet('EVENT')
  };

  const selectedYear = Number(sheets.schedule.getRange(CONFIG.CELLS.YEAR).getValue());
  const selectedMonth = Number(sheets.schedule.getRange(CONFIG.CELLS.MONTH).getValue());

  // [1. ดึงข้อมูลการลา (Leave Map)]
  const leaveMap = {};
  sheets.leaves.getRange(2, 1, 31, 3).getValues().forEach(row => {
    if (row[0]) {
      const d = row[0];
      leaveMap[d + "_ค่ำ"] = row[1] ? row[1].toString().split(",").map(s => s.trim()) : [];
      leaveMap[d + "_ดึก"] = row[2] ? row[2].toString().split(",").map(s => s.trim()) : [];
    }
  });

  // [2. ดึงข้อมูลการล็อคคิว (Lock Map)]
  const eventData = sheets.event.getDataRange().getValues();
  const lockMap = {};
  eventData.slice(1).forEach(row => {
    if (Number(row[1]) === selectedYear && Number(row[2]) === selectedMonth) {
      lockMap[row[4] + "_" + row[5]] = row[3];
    }
  });

  // [3. ดึงข้อมูลนักดนตรี (Musicians Pool)]
  const musicians = sheets.master.getDataRange().getValues().slice(1)
    .filter(row => row[0] && (row[7] || "").toString().toLowerCase() === 'active')
    .map(row => ({
      name: row[0],
      roles: [
        row[1] == 1 ? "vocal" : null,
        row[2] == 1 ? "acg" : null,
        row[3] == 1 ? "eig" : null,
        row[4] == 1 ? "bass" : null,
        row[5] == 1 ? "drum" : null,
        row[6] == 1 ? "key" : null
      ].filter(Boolean),
      conditions: (row[9] || "").toString().toLowerCase().split(',').map(s => s.trim())
    }));

  // [4. ประมวลผลตาราง (Core Logic)]
  const scheduleRange = sheets.schedule.getRange(CONFIG.CELLS.DATE_COLUMN).getValues();
  const finalDataC = [];
  const finalDataE = [];
  let prevEveningVocal = "", prevNightVocal = "";

  scheduleRange.forEach((row, i) => {
    const rowDate = row[0];
    if (!rowDate) {
      finalDataC.push([""]); finalDataE.push([""]); return;
    }

    const dateObj = new Date(rowDate);
    const dayNum = dateObj.getDate();
    const dayOfWeek = dateObj.getDay();

    const buildTeam = (shift, forceFull, otherShiftVocal, myPrevVocal, otherPrevVocal) => {
      const lockKey = `${dayNum}_${shift}`;
      if (lockMap[lockKey]) return [lockMap[lockKey]];

      const banned = leaveMap[`${dayNum}_${shift}`] || [];
      const available = musicians.filter(m => !banned.includes(m.name));

      let team = [];
      let selectedVocal = "";
      
      // Note: "Staff_A" and "Staff_B" are core staff members with specific role priorities.
      const vocalPool = available.filter(m => m.roles.includes("vocal") && m.name !== "Staff_A" && m.name !== "Staff_B" && m.name !== otherShiftVocal);

      if (vocalPool.length > 0) {
        const swapVocal = vocalPool.find(m => m.name === otherPrevVocal);
        selectedVocal = swapVocal ? swapVocal.name : vocalPool[i % vocalPool.length].name;
      } else {
        const otherDrummers = available.filter(m => m.roles.includes("drum") && m.name !== "Staff_A");
        const gameAvail = available.find(m => m.name === "Staff_A" && m.name !== otherShiftVocal);
        if (gameAvail && otherDrummers.length > 0) {
          selectedVocal = "Staff_A";
        } else {
          const doAvail = available.find(m => m.name === "Staff_B" && m.name !== otherShiftVocal);
          if (doAvail) selectedVocal = "Staff_B";
        }
      }

      if (selectedVocal) team.push(selectedVocal);

      const pick = (role) => {
        if (role === 'vocal') return;
        let pool = available.filter(m => m.roles.includes(role) && !team.includes(m.name));
        if (role === 'drum') {
          if (selectedVocal === "Staff_A") {
            const altDrum = pool.find(m => m.name !== "Staff_A");
            if (altDrum) team.push(altDrum.name);
          } else {
            const gameDrum = pool.find(m => m.name === "Staff_A");
            if (gameDrum) team.push(gameDrum.name);
            else if (pool.length > 0) team.push(pool[0].name);
          }
        } else if (role === 'acg' || role === 'eig') {
          const doGtr = pool.find(m => m.name === "Staff_B");
          if (doGtr) team.push(doGtr.name);
          else if (pool.length > 0) team.push(pool[0].name);
        } else {
          if (pool.length > 0) team.push(pool[i % pool.length].name);
        }
      };

      pick('drum');
      if (forceFull || dayOfWeek === 5 || dayOfWeek === 6) {
        ['eig', 'bass', 'key'].forEach(pick);
      } else {
        pick('acg');
      }
      return team;
    };

    const teamC = buildTeam("ค่ำ", false, "", prevEveningVocal, prevNightVocal);
    const teamE = buildTeam("ดึก", true, teamC[0] || "", prevNightVocal, prevEveningVocal);

    finalDataC.push([teamC.join(", ")]);
    finalDataE.push([teamE.join(", ")]);
    
    prevEveningVocal = teamC[0] || "";
    prevNightVocal = teamE[0] || "";
  });

  // [5. บันทึกข้อมูลลง Sheet ตามตำแหน่งที่ CONFIG กำหนด]
  sheets.schedule.getRange(3, 3, finalDataC.length, 1).setValues(finalDataC);
  sheets.schedule.getRange(3, 5, finalDataE.length, 1).setValues(finalDataE);
}