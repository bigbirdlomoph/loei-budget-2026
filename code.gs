const SPREADSHEET_ID = '1L7jenTSA4Jwmjq5QPsPn4nR-BAam9TfLiYcfrV_B0zU';

// ลำดับอำเภอมาตรฐานตามรหัส 4201 - 4214
const DISTRICT_ORDER = ["เมืองเลย", "นาด้วง", "เชียงคาน", "ปากชม", "ด่านซ้าย", "นาแห้ว", "ภูเรือ", "ท่าลี่", "วังสะพุง", "ภูกระดึง", "ภูหลวง", "ผาขาว", "เอราวัณ", "หนองหิน"];

function doGet() {
  const version = getSystemVersion();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard งบลงทุน สสจ.เลย (V.' + version + ')') // ระบุ Version ใน Title Bar
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSystemVersion() {
  const now = new Date();
  return (now.getFullYear() + 543) + 
         ("0" + (now.getMonth() + 1)).slice(-2) + 
         ("0" + now.getDate()).slice(-2) + "-" + 
         ("0" + now.getHours()).slice(-2) + 
         ("0" + now.getMinutes()).slice(-2);
}

function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hospSheet = ss.getSheetByName('c_hospital');
    const hospitalData = getSheetData(hospSheet);
    
    // สร้าง Mapping ข้อมูลอำเภอและรหัสเพื่อความแม่นยำ
    const hospMap = {};
    hospitalData.forEach(function(h) {
      var id = String(h['รหัสหน่วยบริการ'] || '').trim();
      if (id) {
        hospMap[id] = {
          distCode: h['รหัสอำเภอ'] || 9999,
          unitType: h['ประเภทหน่วย'] || '-',
          amphoe: h['อำเภอ'] || '-'
        };
      }
    });

    const equipment = getSheetData(ss.getSheetByName('bureau_equipment'));
    const building = getSheetData(ss.getSheetByName('bureau_building'));

    return {
      equipment: equipment,
      building: building,
      hospMap: hospMap,
      years: [...new Set([...equipment, ...building].map(d => d['ปีงบประมาณ']))].filter(y => y).sort().reverse(),
      amphoes: DISTRICT_ORDER, // เรียงลำดับอำเภอคงที่ 4201-4214
      version: getSystemVersion()
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data.shift().map(h => String(h).trim());
  return data.map(row => headers.reduce((acc, header, i) => {
    acc[header] = row[i];
    return acc;
  }, {}));
}

// ฟังก์ชันสำหรับดึงข้อมูล Lookup ต่างๆ
function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amphoeSheet = ss.getSheetByName('c_amphoe');
  const hospitalSheet = ss.getSheetByName('c_hospital');
  
  // ดึงข้อมูลปีงบประมาณ (สมมติว่าดึงจาก m_budget_equipment หรือกำหนดเอง)
  const years = [2568, 2569]; 

  const amphoes = amphoeSheet.getDataRange().getValues().slice(1); // ตัด Header ออก
  const hospitals = hospitalSheet.getDataRange().getValues().slice(1);

  return { years, amphoes, hospitals };
}

// ฟังก์ชันค้นหาข้อมูล
function searchProgressReport(criteria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = '';
  
  if (criteria.budgetType === 'ครุภัณฑ์') {
    sheetName = 'm_budget_equipment';
  } else if (criteria.budgetType === 'สิ่งก่อสร้าง') {
    sheetName = 'm_budget_building';
  } else {
    return []; // หรือแจ้งเตือนว่าต้องเลือกประเภทงบลงทุน
  }

  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // กรองข้อมูลตาม Criteria
  const filteredData = rows.filter(row => {
    // ... เขียน Logic การกรองข้อมูลที่นี่ ...
    // เช่น ตรวจสอบปีงบประมาณ, รหัสหน่วยบริการ (ที่ได้จากชื่อหน่วยบริการหรืออำเภอ)
    return true; // ตัวอย่าง
  });

  return { headers, data: filteredData };
}
