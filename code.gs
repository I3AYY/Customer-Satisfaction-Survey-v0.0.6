/**
 * ============================================================================
 * เว็บแอปพลิเคชันระบบประเมินความพึงพอใจผู้รับบริการ (Customer Satisfaction Survey)
 * เวอร์ชัน: 0.0.6 AI Core (Multi-Role Support & Internal Comments)
 * กลุ่มงานพยาธิวิทยาคลินิกและเทคนิคการแพทย์ โรงพยาบาลมะเร็งสุราษฎร์ธานี
 * ============================================================================
 */

// โครงสร้างการตั้งค่าของแต่ละกลุ่มเป้าหมาย
const ROLE_CONFIG = {
  'Patient': { q: 'Patient_Q', a: 'Patient_A', c: null }, // ผู้ป่วย ยังคงใช้ระบบ QR/Google Form สำหรับข้อเสนอแนะ (ตามโครงสร้างเดิม)
  'Doctor': { q: 'Doctor_Q', a: 'Doctor_A', c: 'Doctor_C' }, // แพทย์
  'Nurse': { q: 'Nurse_Q', a: 'Nurse_A', c: 'Nurse_C' },   // พยาบาล
  'NA': { q: 'NA_Q', a: 'NA_A', c: 'NA_C' }            // ผู้ช่วยเหลือคนไข้
};

function doGet(e) {
  setupSheets();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบประเมินความพึงพอใจ - กลุ่มงานพยาธิวิทยาคลินิกและเทคนิคการแพทย์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ติดตั้งและตรวจสอบฐานข้อมูล (Google Sheets) อัตโนมัติสำหรับทุกกลุ่มเป้าหมาย
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // สร้างและจัดการ Sheet ตาม Role ที่ตั้งไว้
  for (const [role, config] of Object.entries(ROLE_CONFIG)) {
    // 1. ชีตเก็บคำถาม (_Q)
    let qSheet = ss.getSheetByName(config.q);
    if (!qSheet) {
      qSheet = ss.insertSheet(config.q);
      qSheet.appendRow(['ID', 'QuestionText', 'ImageURL', 'Status']);
      qSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
      
      const sampleQuestions = [
        [1, 'ความพึงพอใจในบริการข้อที่ 1', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Question+1', true],
        [2, 'ความพึงพอใจในบริการข้อที่ 2', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Question+2', true],
        [3, 'ความพึงพอใจในบริการข้อที่ 3', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Question+3', true]
      ];
      sampleQuestions.forEach(q => qSheet.appendRow(q));
      qSheet.setFrozenRows(1);
      qSheet.autoResizeColumns(1, 4);
    }

    // 2. ชีตเก็บคำตอบ (_A)
    let aSheet = ss.getSheetByName(config.a);
    if (!aSheet) {
      aSheet = ss.insertSheet(config.a);
      const headers = ['Timestamp', 'Year', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5'];
      aSheet.appendRow(headers);
      aSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
      aSheet.setFrozenRows(1);
      
      const range = aSheet.getRange("C2:G2000");
      const ruleSatisfied = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('พอใจ').setBackground('#d1fae5').setFontColor('#065f46').setRanges([range]).build();
      const ruleDissatisfied = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('ไม่พอใจ').setBackground('#ffe4e6').setFontColor('#9f1239').setRanges([range]).build();

      const rules = aSheet.getConditionalFormatRules();
      rules.push(ruleSatisfied, ruleDissatisfied);
      aSheet.setConditionalFormatRules(rules);
      aSheet.autoResizeColumns(1, 7);
    }

    // 3. ชีตเก็บข้อเสนอแนะเพิ่มเติม (_C) (สำหรับ แพทย์, พยาบาล, NA)
    if (config.c) {
      let cSheet = ss.getSheetByName(config.c);
      if (!cSheet) {
        cSheet = ss.insertSheet(config.c);
        cSheet.appendRow(['Timestamp', 'Comment']);
        cSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
        cSheet.setFrozenRows(1);
        cSheet.autoResizeColumns(1, 2);
      }
    }
  }

  // 4. ชีตเก็บการตั้งค่าและโลโก้ (Logo)
  let logoSheet = ss.getSheetByName('Logo');
  if (!logoSheet) {
    logoSheet = ss.insertSheet('Logo');
    logoSheet.getRange('A1:B1').setValues([['Setting', 'Value']]).setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    logoSheet.getRange('A2:B3').setValues([
      ['Logo URL', 'https://drive.google.com/thumbnail?id=1NvjW7zbszUCWPhBcH4th6uVsYqjACdED&sz=w800'],
      ['QR Code URL', 'https://placehold.co/200x200/ffffff/1e3a8a?text=Scan+to+Comment']
    ]);
    logoSheet.setFrozenRows(1);
    logoSheet.autoResizeColumns(1, 2);
  } else {
    const data = logoSheet.getDataRange().getValues();
    let hasQR = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'QR Code URL') hasQR = true;
    }
    if (!hasQR) {
      logoSheet.appendRow(['QR Code URL', 'https://placehold.co/200x200/ffffff/1e3a8a?text=Scan+to+Comment']);
    }
  }
}

/**
 * ดึงข้อมูลเริ่มต้นตามกลุ่มเป้าหมาย (Questions, Settings, Years)
 */
function getInitialData(role = 'Patient') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ROLE_CONFIG[role] || ROLE_CONFIG['Patient'];
  
  // 1. ดึงคำถามตาม Role
  const qSheet = ss.getSheetByName(config.q);
  const activeQuestions = [];
  if (qSheet) {
    const sqData = qSheet.getDataRange().getValues();
    for (let i = 1; i < sqData.length; i++) {
      if (sqData[i][3] === true || sqData[i][3].toString().toUpperCase() === 'TRUE') {
        activeQuestions.push({ id: sqData[i][0], text: sqData[i][1], imageUrl: sqData[i][2] });
      }
    }
  }

  // 2. ดึงการตั้งค่า Logo และ QR
  const logoSheet = ss.getSheetByName('Logo');
  let logoUrl = '';
  let qrUrl = '';
  if (logoSheet) {
    const logoData = logoSheet.getDataRange().getValues();
    for (let i = 1; i < logoData.length; i++) {
      if (logoData[i][0] === 'Logo URL') logoUrl = logoData[i][1].toString().trim();
      if (logoData[i][0] === 'QR Code URL') qrUrl = logoData[i][1].toString().trim();
    }
  }

  // 3. ดึงปีที่มีคนตอบ สำหรับ Role ปัจจุบัน
  const yearSet = new Set();
  const aSheet = ss.getSheetByName(config.a);
  if (aSheet && aSheet.getLastRow() > 1) {
    aSheet.getRange(2, 2, aSheet.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) yearSet.add(r[0].toString()); });
  }

  // สำหรับผู้ป่วย รองรับการดึงปีจากแบบฟอร์มเดิม (ถ้ามี)
  if (role === 'Patient') {
    let formSheet = null;
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sName = sheets[i].getName();
      if (sName.includes('การตอบแบบฟอร์ม') || sName.includes('Form Responses') || sName === 'AdditionalComments') {
        formSheet = sheets[i];
        break;
      }
    }

    if (formSheet && formSheet.getLastRow() > 1) {
      const isGoogleForm = formSheet.getName().includes('การตอบแบบฟอร์ม') || formSheet.getName().includes('Form Responses');
      if (isGoogleForm) {
        formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 1).getValues().forEach(r => {
          if (r[0]) {
             try {
               const d = new Date(r[0]);
               if(!isNaN(d.getTime())) yearSet.add((d.getFullYear() + 543).toString());
             } catch(e) {}
          }
        });
      } else {
        formSheet.getRange(2, 2, formSheet.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) yearSet.add(r[0].toString()); });
      }
    }
  }
  
  const availableYears = Array.from(yearSet).sort((a, b) => b - a);
  const currentYearBE = (new Date().getFullYear() + 543).toString();
  
  return { 
    questions: activeQuestions, 
    logoUrl: logoUrl, 
    qrUrl: qrUrl, 
    availableYears: availableYears, 
    currentYear: currentYearBE,
    currentRole: role
  };
}

/**
 * บันทึกคำตอบแบบประเมินและข้อเสนอแนะลงตาราง ตาม Role
 */
function submitSurveyResponse(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('ระบบกำลังประมวลผลหนาแน่น กรุณาลองกดส่งใหม่อีกครั้ง'); }
  
  try {
    // รองรับการส่งข้อมูลแบบเดิม (Array) หรือแบบใหม่ (Object)
    let role = 'Patient';
    let answers = [];
    let comment = '';

    if (Array.isArray(payload)) {
      answers = payload;
    } else {
      role = payload.role || 'Patient';
      answers = payload.answers || [];
      comment = payload.comment || '';
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ROLE_CONFIG[role] || ROLE_CONFIG['Patient'];
    const now = new Date();
    const timestamp = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    const yearBE = (now.getFullYear() + 543).toString();
    
    // 1. บันทึกคำตอบลง Sheet _A
    const aSheet = ss.getSheetByName(config.a);
    if (aSheet) {
      const rowData = [timestamp, yearBE];
      for (let i = 0; i < 5; i++) rowData.push(answers[i] ? answers[i] : "");
      aSheet.appendRow(rowData);
    }

    // 2. บันทึกข้อเสนอแนะลง Sheet _C (ยกเว้นผู้ป่วย ที่สแกน QR ผ่าน Google Form แทน)
    if (role !== 'Patient' && comment && comment.trim() !== '') {
      if (config.c) {
        const cSheet = ss.getSheetByName(config.c);
        if (cSheet) {
          cSheet.appendRow([timestamp, comment.trim()]);
        }
      }
    }
    
    lock.releaseLock();
    return { success: true };
  } catch (error) {
    lock.releaseLock();
    throw new Error("เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + error.toString());
  }
}

/**
 * ดึงข้อมูลสถิติและคอมเมนต์สำหรับหน้า Dashboard ตาม Role ที่เลือก
 */
function getDashboardStats(filterYear, role = 'Patient') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ROLE_CONFIG[role] || ROLE_CONFIG['Patient'];
  
  const stats = {
    totalRespondents: 0,
    overall: { satisfied: 0, dissatisfied: 0 },
    questions: {
      Q1: { satisfied: 0, dissatisfied: 0 }, Q2: { satisfied: 0, dissatisfied: 0 },
      Q3: { satisfied: 0, dissatisfied: 0 }, Q4: { satisfied: 0, dissatisfied: 0 },
      Q5: { satisfied: 0, dissatisfied: 0 }
    },
    comments: []
  };
  
  // 1. ประมวลผลคำตอบหลัก
  const aSheet = ss.getSheetByName(config.a);
  if (aSheet) {
    const lastRow = aSheet.getLastRow();
    if (lastRow > 1) {
      const data = aSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      data.forEach(row => {
        const rowYear = row[1] ? row[1].toString() : '';
        if (filterYear === 'ทั้งหมด' || filterYear === rowYear) {
          stats.totalRespondents++;
          
          const qKeys = ['Q1', 'Q2', 'Q3', 'Q4', 'Q5'];
          for(let i=0; i<5; i++){
             const ans = row[i+2]; // คอลัมน์ C ถึง G
             if (ans === 'พอใจ') {
               stats.questions[qKeys[i]].satisfied++;
               stats.overall.satisfied++;
             } else if (ans === 'ไม่พอใจ') {
               stats.questions[qKeys[i]].dissatisfied++;
               stats.overall.dissatisfied++;
             }
          }
        }
      });
    }
  }
  
  // 2. ดึงคอมเมนต์
  if (role === 'Patient') {
    // ผู้ป่วย: ใช้ระบบเดิมที่อ่านจาก Google Form หรือ AdditionalComments
    let commSheet = null;
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sName = sheets[i].getName();
      if (sName.includes('การตอบแบบฟอร์ม') || sName.includes('Form Responses') || sName === 'AdditionalComments') {
        commSheet = sheets[i];
        break;
      }
    }
    
    if (commSheet) {
      const cLastRow = commSheet.getLastRow();
      const isGoogleForm = commSheet.getName().includes('การตอบแบบฟอร์ม') || commSheet.getName().includes('Form Responses');

      if (cLastRow > 1) {
        const numCols = isGoogleForm ? commSheet.getLastColumn() : 3;
        const cData = commSheet.getRange(2, 1, cLastRow - 1, Math.max(numCols, 2)).getValues();
        
        cData.forEach(r => {
          const timestamp = r[0];
          let cYear = '';
          let commentText = '';

          if (isGoogleForm) {
             try {
               const dateObj = new Date(timestamp);
               if(!isNaN(dateObj.getTime())) {
                 cYear = (dateObj.getFullYear() + 543).toString();
               }
             } catch(e) {}
             commentText = r[1] ? r[1].toString() : '';
          } else {
             cYear = r[1] ? r[1].toString() : '';
             commentText = r[2] ? r[2].toString() : '';
          }

          if (filterYear === 'ทั้งหมด' || filterYear === cYear) {
             if (commentText) { 
               let formattedTime = "";
               try {
                 formattedTime = Utilities.formatDate(new Date(timestamp), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
               } catch(e) { formattedTime = timestamp.toString(); }
               
               stats.comments.push({ timestamp: formattedTime, text: commentText });
             }
          }
        });
      }
    }
  } else {
    // แพทย์, พยาบาล, NA: อ่านจาก Sheet _C โดยตรง
    if (config.c) {
      const cSheet = ss.getSheetByName(config.c);
      if (cSheet) {
        const cLastRow = cSheet.getLastRow();
        if (cLastRow > 1) {
          const cData = cSheet.getRange(2, 1, cLastRow - 1, 2).getValues(); // Timestamp, Comment
          cData.forEach(r => {
             const timestamp = r[0];
             const commentText = r[1] ? r[1].toString() : '';
             
             // ดึงปีจาก Timestamp มาตรวจสอบเงื่อนไข Filter
             let cYear = '';
             try {
               const dateObj = new Date(timestamp);
               if(!isNaN(dateObj.getTime())) {
                 cYear = (dateObj.getFullYear() + 543).toString();
               }
             } catch(e) {}

             if (filterYear === 'ทั้งหมด' || filterYear === cYear) {
               if (commentText.trim() !== '') {
                 let formattedTime = "";
                 try {
                   formattedTime = Utilities.formatDate(new Date(timestamp), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
                 } catch(e) { formattedTime = timestamp.toString(); }
                 
                 stats.comments.push({ timestamp: formattedTime, text: commentText });
               }
             }
          });
        }
      }
    }
  }
  
  // จัดเรียงคอมเมนต์ล่าสุดขึ้นก่อนเสมอ
  stats.comments.reverse();
  
  return stats;
}
