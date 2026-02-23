const SPREADSHEET_ID = '1L7jenTSA4Jwmjq5QPsPn4nR-BAam9TfLiYcfrV_B0zU';
const DISTRICT_ORDER = ["เมืองเลย", "นาด้วง", "เชียงคาน", "ปากชม", "ด่านซ้าย", "นาแห้ว", "ภูเรือ", "ท่าลี่", "วังสะพุง", "ภูกระดึง", "ภูหลวง", "ผาขาว", "เอราวัณ", "หนองหิน"];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบงบลงทุน สสจ.เลย')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 1. ดึงข้อมูล Dashboard
// ==========================================
function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hospMap = getHospitalMap(ss);

    const eqData = getSheetData(ss.getSheetByName('bureau_equipment'));
    const bdData = getSheetData(ss.getSheetByName('bureau_building'));

    const mapObj = (list, type) => {
      let result = [];
      for (let i = 0; i < list.length; i++) {
        let d = list[i];
        let status = String(d['การพิจารณา'] || '').trim();
        if (status.indexOf('จัดสรร') !== -1) {
          let hospId = String(d['รหัสหน่วยบริการ'] || '').trim();
          let info = hospMap[hospId] || { name: hospId, amphoe: '-', unitType: '-' };
          result.push({
            year: String(d['ปีงบประมาณ'] || ''),
            status: status,
            totalBudget: parseMoney(d['วงเงินรวม'] || d['วงเงิน']),
            name: String(d['รายการ'] || d['ชื่อรายการ'] || '-'),
            hospId: hospId,
            hospName: info.name,
            amphoe: info.amphoe,
            unitType: info.unitType,
            dataType: type
          });
        }
      }
      return result;
    };

    const equipment = mapObj(eqData, 'ครุภัณฑ์');
    const building = mapObj(bdData, 'สิ่งก่อสร้าง');

    let yearSet = new Set();
    equipment.forEach(x => { if (x.year) yearSet.add(x.year); });
    building.forEach(x => { if (x.year) yearSet.add(x.year); });
    const allYears = Array.from(yearSet).sort().reverse();

    return {
      equipment: equipment,
      building: building,
      years: allYears,
      amphoes: DISTRICT_ORDER,
      version: Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmm")
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ==========================================
// 1.5. ดึงข้อมูลสำหรับหน้าค้นหาประวัติคำขอ (Search)
// ==========================================
function getSearchData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hospMap = getHospitalMap(ss);

    const eqData = getSheetData(ss.getSheetByName('bureau_equipment'));
    const bdData = getSheetData(ss.getSheetByName('bureau_building'));

    const mapObj = (list, type) => {
      let result = [];
      for (let i = 0; i < list.length; i++) {
        let d = list[i];
        let hospId = String(d['รหัสหน่วยบริการ'] || '').trim();
        let info = hospMap[hospId] || { name: hospId, amphoe: '-', unitType: '-' };
        
        result.push({
          type: type,
          year: String(d['ปีงบประมาณ'] || ''),
          hospId: hospId,
          hospName: info.name !== hospId ? info.name : (d['ชื่อหน่วยบริการ'] || hospId),
          amphoe: info.amphoe !== '-' ? info.amphoe : (String(d['อำเภอ'] || '-').trim()),
          unitType: String(d['ประเภทหน่วย'] || info.unitType || '-').trim(),
          hospLevel: String(d['ระดับหน่วยบริการเดิม'] || info.hospLevel || '-').trim(),
          sapLevel: String(d['ระดับ SAP'] || info.sapLevel || '-').trim(),
          name: String(d['รายการ'] || d['ชื่อรายการ'] || '-'),
          price: parseMoney(d['ราคาต่อหน่วย'] || 0),
          budget: parseMoney(d['วงเงินรวม'] || d['วงเงิน'] || 0),
          status: String(d['การพิจารณา'] || '-').trim()
        });
      }
      return result;
    };

    return {
      equipment: mapObj(eqData, 'ครุภัณฑ์'),
      building: mapObj(bdData, 'สิ่งก่อสร้าง')
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ==========================================
// 2. ดึงข้อมูลรายงาน (Report)
// ==========================================
function getReportData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hospMap = getHospitalMap(ss);
    return {
      equipment: getMBudgetList(ss.getSheetByName('m_budget_equipment'), 'ครุภัณฑ์', hospMap),
      building: getMBudgetList(ss.getSheetByName('m_budget_building'), 'สิ่งก่อสร้าง', hospMap)
    };
  } catch (e) { return { error: e.message }; }
}

function getMBudgetList(sheet, type, hospMap) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const rows = data.slice(1);

  return rows.map(row => {
    const getVal = (idx) => {
      let v = row[idx];
      if (v instanceof Date) return Utilities.formatDate(v, "GMT+7", "yyyy-MM-dd");
      return v !== undefined && v !== null ? String(v).trim() : '';
    };

    const hospId = getVal(6);
    const hospInfo = hospMap[hospId] || { name: hospId, amphoe: '-', unitType: '-' };

    if (type === 'ครุภัณฑ์') {
      // Equipment column map (0-based idx):
      // 8 unitPrice, 13 totalBudget, 14 method, 15 signDate, 16 endDate, 17 delDate, 18 inspectDate,
      // 19 contract(alloc), 20 payDate, 21 spentAmount(paid), 22 balance, 23 procStep, 24 status,
      // 25 spentStatus, 26 risk, 27 note(remark)
      return {
        id: getVal(0),
        dataType: 'ครุภัณฑ์',
        year: getVal(1),
        name: getVal(5),
        hospName: hospInfo.name,
        amphoe: hospInfo.amphoe,

        unitPrice: parseMoney(getVal(8)),
        totalBudget: parseMoney(getVal(13)),
        contractAmount: parseMoney(getVal(19)),
        method: getVal(14),

        procStep: getVal(23),
        status: getVal(24),
        spentStatus: getVal(25),
        risk: getVal(26),

        contractSignDate: getVal(15),
        contractEndDate: getVal(16),
        deliveryDate: getVal(17),
        inspectionDate: getVal(18),
        paymentDate: getVal(20),

        spentAmount: parseMoney(getVal(21)),
        balance: parseMoney(getVal(22)),
        note: getVal(27),
        period: '-'
      };
    } else {
      // Building column map (0-based idx):
      // 10 unitPrice, 15 totalBudget, 16 method, 17 signDate, 18 endDate, 19 delDate, 20 inspectDate,
      // 21 contract(alloc), 22 payDate, 23 spentAmount(paid), 24 balance, 31 procStep, 32 status,
      // 33 spentStatus, 34 risk, 35 note(remark)
      return {
        id: getVal(0),
        dataType: 'สิ่งก่อสร้าง',
        year: getVal(1),
        name: getVal(5),
        hospName: hospInfo.name,
        amphoe: hospInfo.amphoe,

        unitPrice: parseMoney(getVal(10)),
        totalBudget: parseMoney(getVal(15)),
        contractAmount: parseMoney(getVal(21)),
        method: getVal(16),

        procStep: getVal(31),
        status: getVal(32),
        spentStatus: getVal(33),
        risk: getVal(34),

        contractSignDate: getVal(17),
        contractEndDate: getVal(18),
        deliveryDate: getVal(19),
        inspectionDate: getVal(20),
        paymentDate: getVal(22),

        spentAmount: parseMoney(getVal(23)),
        balance: parseMoney(getVal(24)),
        note: getVal(35),
        period: (getVal(27) || '-') + '/' + (getVal(25) || '-'),
        totalPeriod: getVal(25),
        yearPeriod: getVal(26),
        currentPeriod: getVal(27),
        delayPeriod: getVal(28),
        delayReason: getVal(29)
      };
    }
  });
}

// ==========================================
// 3. ดึงตัวเลือก (Options) สำหรับฟอร์ม
// ==========================================
function getFormOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const options = {};
    const hospSheet = ss.getSheetByName('c_hospital');
    if (hospSheet && hospSheet.getLastRow() > 1) {
      const data = hospSheet.getDataRange().getValues().slice(1);
      options.hospitals = data.map(r => ({ id: String(r[8] || '').trim(), name: String(r[10] || '').trim(), amphoe: String(r[5] || '').trim() })).filter(h => h.id);
    } else { options.hospitals = []; }

    ['c_risk', 'c_procedure', 'c_status', 'c_spent', 'c_proc_ebid_equipment', 'c_proc_specific_equipment', 'c_proc_ebid_building', 'c_proc_specific_building', 'c_proc_selection'].forEach(name => {
      const s = ss.getSheetByName(name);
      options[name] = (s && s.getLastRow() > 1) ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().flat().map(String).filter(String) : [];
    });
    return options;
  } catch (e) { return { error: e.message }; }
}

// ==========================================
// 4. บันทึกและแก้ไขข้อมูล (Save & Update)
// ==========================================
function saveBudgetRecord(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const isEq = form.budgetType === 'ครุภัณฑ์';
    const sheetName = isEq ? 'm_budget_equipment' : 'm_budget_building';
    const logSheetName = isEq ? 't_equipment_log' : 't_building_log';

    const sheet = ss.getSheetByName(sheetName);
    let logSheet = ss.getSheetByName(logSheetName);
    if (!sheet) return { success: false, error: 'ไม่พบชีต ' + sheetName };
    if (!logSheet) logSheet = ss.insertSheet(logSheetName);

    const lastRow = sheet.getLastRow();
    let newId = (isEq ? 'eq-' : 'bu-') + '0001';
    if (lastRow > 1) {
      const lastId = String(sheet.getRange(lastRow, 1).getValue());
      const parts = lastId.split('-');
      if (parts.length === 2) { newId = parts[0] + '-' + ('0000' + (parseInt(parts[1]) + 1)).slice(-4); }
    }

    let rowData = [];
    if (isEq) {
      rowData = [newId, form.year, 'รอพิจารณา', form.budgetType, form.subType, form.name, form.hospId, form.unit, parseMoney(form.price), parseMoney(form.qty), parseMoney(form.bg1), parseMoney(form.bg2), parseMoney(form.bg3), parseMoney(form.total), form.method, form.signDate, form.endDate, form.delDate, form.inspectDate, parseMoney(form.alloc), form.paidDate, parseMoney(form.paid), parseMoney(form.balance), form.procStep, form.status, form.spentStatus, form.risk, form.remark];
    } else {
      rowData = [newId, form.year, 'รอพิจารณา', form.budgetType, form.subType, form.name, form.hospId, form.duration, form.periodStr, form.unit, parseMoney(form.price), parseMoney(form.qty), parseMoney(form.bg1), parseMoney(form.bg2), parseMoney(form.bg3), parseMoney(form.total), form.method, form.signDate, form.endDate, form.delDate, form.inspectDate, parseMoney(form.alloc), form.paidDate, parseMoney(form.paid), parseMoney(form.balance), form.totalPeriod, form.yearPeriod, form.currentPeriod, form.delayPeriod, form.delayReason, form.contractDueDate, form.procStep, form.status, form.spentStatus, form.risk, form.remark];
    }

    sheet.appendRow(rowData);
    logSheet.appendRow([new Date(), ...rowData.slice(1)]);
    return { success: true, id: newId };
  } catch (e) { return { success: false, error: e.toString() }; } finally { lock.releaseLock(); }
}

function updateBudgetRecord(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const isEq = form.budgetType === 'ครุภัณฑ์';
    const sheetName = isEq ? 'm_budget_equipment' : 'm_budget_building';
    const logSheetName = isEq ? 't_equipment_log' : 't_building_log';

    const sheet = ss.getSheetByName(sheetName);
    let logSheet = ss.getSheetByName(logSheetName);
    if (!sheet) return { success: false, error: 'ไม่พบชีต ' + sheetName };
    if (!logSheet) logSheet = ss.insertSheet(logSheetName);

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(r => String(r[0]) === String(form.id));
    if (rowIndex === -1) return { success: false, error: 'ไม่พบ ID: ' + form.id };
    const rowNum = rowIndex + 1;

    const contract = parseMoney(form.contract);
    const spentAmount = parseMoney(form.spentAmount);
    const balance = contract - spentAmount;

    if (isEq) {
      sheet.getRange(rowNum, 15).setValue(form.method || '');
      sheet.getRange(rowNum, 20).setValue(contract);
      sheet.getRange(rowNum, 24).setValue(form.procStep || '');
      sheet.getRange(rowNum, 25).setValue(form.status || '');
      sheet.getRange(rowNum, 26).setValue(form.spentStatus || '');
      sheet.getRange(rowNum, 27).setValue(form.risk || '');

      if (form.contractSignDate !== undefined) sheet.getRange(rowNum, 16).setValue(form.contractSignDate || '');
      if (form.contractEndDate !== undefined) sheet.getRange(rowNum, 17).setValue(form.contractEndDate || '');
      if (form.deliveryDate !== undefined) sheet.getRange(rowNum, 18).setValue(form.deliveryDate || '');
      if (form.inspectionDate !== undefined) sheet.getRange(rowNum, 19).setValue(form.inspectionDate || '');
      if (form.paymentDate !== undefined) sheet.getRange(rowNum, 21).setValue(form.paymentDate || '');

      sheet.getRange(rowNum, 22).setValue(spentAmount);
      sheet.getRange(rowNum, 23).setValue(balance);

      if (form.note !== undefined) sheet.getRange(rowNum, 28).setValue(form.note || '');
    } else {
      sheet.getRange(rowNum, 17).setValue(form.method || '');
      sheet.getRange(rowNum, 22).setValue(contract);
      sheet.getRange(rowNum, 32).setValue(form.procStep || '');
      sheet.getRange(rowNum, 33).setValue(form.status || '');
      sheet.getRange(rowNum, 34).setValue(form.spentStatus || '');
      sheet.getRange(rowNum, 35).setValue(form.risk || '');

      if (form.contractSignDate !== undefined) sheet.getRange(rowNum, 18).setValue(form.contractSignDate || '');
      if (form.contractEndDate !== undefined) sheet.getRange(rowNum, 19).setValue(form.contractEndDate || '');
      if (form.deliveryDate !== undefined) sheet.getRange(rowNum, 20).setValue(form.deliveryDate || '');
      if (form.inspectionDate !== undefined) sheet.getRange(rowNum, 21).setValue(form.inspectionDate || '');
      if (form.paymentDate !== undefined) sheet.getRange(rowNum, 23).setValue(form.paymentDate || '');

      sheet.getRange(rowNum, 24).setValue(spentAmount);
      sheet.getRange(rowNum, 25).setValue(balance);

      if (form.totalPeriod !== undefined) sheet.getRange(rowNum, 26).setValue(form.totalPeriod || '');
      if (form.yearPeriod !== undefined) sheet.getRange(rowNum, 27).setValue(form.yearPeriod || '');
      if (form.currentPeriod !== undefined) sheet.getRange(rowNum, 28).setValue(form.currentPeriod || '');
      if (form.delayPeriod !== undefined) sheet.getRange(rowNum, 29).setValue(form.delayPeriod || '');
      if (form.delayReason !== undefined) sheet.getRange(rowNum, 30).setValue(form.delayReason || '');

      if (form.note !== undefined) sheet.getRange(rowNum, 36).setValue(form.note || '');
    }

    const updatedRow = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    logSheet.appendRow([new Date(), ...updatedRow.slice(1)]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 5. HELPER FUNCTIONS
// ==========================================
function parseMoney(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;

  let s = String(val).trim();
  let neg = false;
  if (s.startsWith('(') && s.endsWith(')')) { neg = true; s = s.slice(1, -1); }

  s = s.replace(/[, ]/g, '').replace(/[^\d.-]/g, '');
  const n = parseFloat(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}

function getHospitalMap(ss) {
  const sheet = ss.getSheetByName('c_hospital');
  if (!sheet) return {};
  const data = getSheetData(sheet);
  const map = {};
  data.forEach(r => {
    const id = String(r['รหัสหน่วยบริการ'] || r['รหัส'] || '').trim();
    if (id) {
      map[id] = { 
        name: r['ชื่อเต็มหน่วยบริการ'] || r['ชื่อหน่วยบริการ'] || id, 
        amphoe: r['อำเภอ'] || '', 
        unitType: r['ประเภทหน่วย'] || '',
        hospLevel: r['ระดับหน่วยบริการเดิม'] || '',
        sapLevel: r['ระดับ SAP'] || ''
      };
    }
  });
  return map;
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const rows = [];

  for (let r = 1; r < data.length; r++) {
    let rowObj = {};
    for (let c = 0; c < headers.length; c++) {
      let val = data[r][c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
      } else if (val === undefined || val === null) {
        val = '';
      }
      rowObj[headers[c]] = val;
    }
    rows.push(rowObj);
  }
  return rows;
}
