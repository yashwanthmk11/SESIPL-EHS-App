/**
 * Run once from the Apps Script editor: initializeSystem()
 * Creates the database spreadsheet, Drive folders, seed users and demo project.
 */
function initializeSystem() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const configuredId = String(DATABASE_SPREADSHEET_ID || '').trim();
    let ssId = configuredId || props.getProperty('SPREADSHEET_ID');
    let ss;
    if (ssId) {
      try {
        ss = SpreadsheetApp.openById(ssId);
      } catch (e) {
        if (configuredId) {
          throw new Error('Cannot open configured database spreadsheet ' + configuredId + '. Share it with the Apps Script owner and run initializeSystem() again. Details: ' + e);
        }
        Logger.log('Stored database spreadsheet could not be opened; creating a replacement. ' + e);
        ssId = '';
      }
    }
    if (!ss) {
      ss = SpreadsheetApp.create(APP_NAME + ' Database');
      ssId = ss.getId();
      props.setProperty('SPREADSHEET_ID', ssId);
    }
    Logger.log('Database spreadsheet: ' + ss.getUrl());

    Object.keys(HEADERS).forEach(name => {
      ensureSheetSchema_(ss, name, HEADERS[name]);
    });
    const extra = ss.getSheetByName('Sheet1');
    if (extra && ss.getSheets().length > 1) ss.deleteSheet(extra);

    let rootId = props.getProperty('ROOT_FOLDER_ID');
    let root;
    if (rootId) {
      try {
        root = DriveApp.getFolderById(rootId);
      } catch (e) {
        Logger.log('Stored root folder could not be opened; creating a replacement. ' + e);
        rootId = '';
      }
    }
    if (!root) {
      root = DriveApp.createFolder(APP_NAME + ' Files');
      rootId = root.getId();
      props.setProperty('ROOT_FOLDER_ID', rootId);
    }
    Logger.log('Root folder: ' + root.getUrl());

    const p1 = 'PRJ_INTUIT';
    const p2 = 'PRJ_SIEMENS';
    const p3 = 'PRJ_QUALCOMM';
    const p4 = 'PRJ_INFOSYS';

    const userRows = [
      { employeeId: 'EMP001', uan: 'UAN001', name: 'Site Lead (Intuit)', role: ROLES.LEAD, email: 'harish.ehs@sesipl.com', phone: '+91 98450 44004', active: 'TRUE', mappedProjects: p1 },
      { employeeId: 'EMP002', uan: 'UAN002', name: 'Asst EHS Manager', role: ROLES.ASST, email: 'asst.mgr@sesipl.com', phone: '+91 98450 22005', active: 'TRUE', mappedProjects: p1 + ',' + p2 },
      { employeeId: 'EMP003', uan: 'UAN003', name: 'EHS Manager', role: ROLES.MANAGER, email: 'manager.ehs@sesipl.com', phone: '+91 98450 33003', active: 'TRUE', mappedProjects: '' },
      { employeeId: 'EMP004', uan: 'UAN004', name: 'Director', role: ROLES.DIRECTOR, email: 'director@sesipl.com', phone: '+91 98450 11001', active: 'TRUE', mappedProjects: '' },
      { employeeId: 'EMP005', uan: 'UAN005', name: 'Site Lead (Qualcomm)', role: ROLES.LEAD, email: 'murugan.ehs@sesipl.com', phone: '+91 98450 44009', active: 'TRUE', mappedProjects: p3 }
    ];
    userRows.forEach(u => {
      const existing = findOne_(SHEETS.USERS, 'employeeId', u.employeeId);
      if (existing) {
        updateRowByKey_(SHEETS.USERS, 'employeeId', u.employeeId, u);
      } else {
        appendRow_(SHEETS.USERS, u);
      }
    });

    const projects = [
      { id: p1, code: 'INTUIT', name: 'Intuit', client: 'Intuit', pmc: 'CBRE', inCharge: 'Mr.Harish', manager: 'HR Ravikiran', scope: 'Internal Electrical work (Fit Out)', startDate: '2025-01-01', endDate: '2026-08-31', areaSqft: '389175', poNo: 'C 47344', status: 'RUNNING', region: 'Bangalore', projectDuration: '8 Months' },
      { id: p2, code: 'SIEMENS', name: 'Siemens', client: 'Siemens', pmc: 'Cushman & Wakefield', inCharge: 'S. Rajesh', manager: 'K. Sharma', scope: 'Electrical Fit Out & Commissioning', startDate: '2025-06-01', endDate: '2026-05-31', areaSqft: '245000', poNo: 'C 48120', status: 'RUNNING', region: 'Bangalore', projectDuration: '12 Months' },
      { id: p3, code: 'QUALCOMM', name: 'Qualcomm-CH', client: 'Qualcomm', pmc: 'JLL', inCharge: 'V. Murugan', manager: 'HR Ravikiran', scope: 'HV & LV Electrical Installation', startDate: '2025-03-01', endDate: '2026-01-31', areaSqft: '410000', poNo: 'C 49055', status: 'RUNNING', region: 'Chennai', projectDuration: '10 Months' },
      { id: p4, code: 'INFOSYS', name: 'Infosys', client: 'Infosys', pmc: 'Turner & Townsend', inCharge: 'A. Reddy', manager: 'K. Sharma', scope: 'Internal Electrical & Substation', startDate: '2025-02-01', endDate: '2026-04-30', areaSqft: '520000', poNo: 'C 50210', status: 'RUNNING', region: 'Hyderabad', projectDuration: '14 Months' }
    ];
    projects.forEach(p => {
      if (!findOne_(SHEETS.PROJECTS, 'id', p.id)) {
        appendRow_(SHEETS.PROJECTS, p);
      } else {
        updateRowById_(SHEETS.PROJECTS, p.id, p);
      }
      ensureProjectFolder_(p);
    });

    [
      { employeeId: 'EMP001', projectId: p1, role: ROLES.LEAD },
      { employeeId: 'EMP005', projectId: p3, role: ROLES.LEAD },
      { employeeId: 'EMP002', projectId: p1, role: ROLES.ASST },
      { employeeId: 'EMP002', projectId: p2, role: ROLES.ASST }
    ].forEach(r => {
      const mapped = rowsToObjects_(SHEETS.PROJECT_USERS).some(x =>
        x.employeeId === r.employeeId && x.projectId === r.projectId && x.role === r.role
      );
      if (!mapped) appendRow_(SHEETS.PROJECT_USERS, r);
    });

    // Seed Project Organization Chart Hierarchy (Slide 5)
    const orgSeeds = [
      { projectId: p1, level: '1', roleTitle: 'Director', name: 'Mr. Shankar', phone: '+91 98450 11001', email: 'director@sesipl.com' },
      { projectId: p1, level: '2', roleTitle: 'General Manager', name: 'Mr. C. Shekhar', phone: '+91 98450 22002', email: 'gm@sesipl.com' },
      { projectId: p1, level: '3', roleTitle: 'Project Manager', name: 'HR Ravikiran', phone: '+91 98450 33003', email: 'ravikiran@sesipl.com' },
      { projectId: p1, level: '4', roleTitle: 'HSE Officer / In charge', name: 'Mr. Harish', phone: '+91 98450 44004', email: 'harish.ehs@sesipl.com' },
      { projectId: p1, level: '5', roleTitle: 'Project Engineer', name: 'P. Naveen', phone: '+91 98450 55005', email: 'naveen.pe@sesipl.com' },

      { projectId: p2, level: '1', roleTitle: 'Director', name: 'Mr. Shankar', phone: '+91 98450 11001', email: 'director@sesipl.com' },
      { projectId: p2, level: '2', roleTitle: 'General Manager', name: 'Mr. C. Shekhar', phone: '+91 98450 22002', email: 'gm@sesipl.com' },
      { projectId: p2, level: '3', roleTitle: 'Project Manager', name: 'K. Sharma', phone: '+91 98450 33006', email: 'sharma@sesipl.com' },
      { projectId: p2, level: '4', roleTitle: 'HSE Officer / In charge', name: 'S. Rajesh', phone: '+91 98450 44007', email: 'rajesh.ehs@sesipl.com' },
      { projectId: p2, level: '5', roleTitle: 'Project Engineer', name: 'M. Karthik', phone: '+91 98450 55008', email: 'karthik.pe@sesipl.com' },

      { projectId: p3, level: '1', roleTitle: 'Director', name: 'Mr. Shankar', phone: '+91 98450 11001', email: 'director@sesipl.com' },
      { projectId: p3, level: '2', roleTitle: 'General Manager', name: 'Mr. C. Shekhar', phone: '+91 98450 22002', email: 'gm@sesipl.com' },
      { projectId: p3, level: '3', roleTitle: 'Project Manager', name: 'HR Ravikiran', phone: '+91 98450 33003', email: 'ravikiran@sesipl.com' },
      { projectId: p3, level: '4', roleTitle: 'HSE Officer / In charge', name: 'V. Murugan', phone: '+91 98450 44009', email: 'murugan.ehs@sesipl.com' },
      { projectId: p3, level: '5', roleTitle: 'Project Engineer', name: 'S. Vignesh', phone: '+91 98450 55010', email: 'vignesh.pe@sesipl.com' },

      { projectId: p4, level: '1', roleTitle: 'Director', name: 'Mr. Shankar', phone: '+91 98450 11001', email: 'director@sesipl.com' },
      { projectId: p4, level: '2', roleTitle: 'General Manager', name: 'Mr. C. Shekhar', phone: '+91 98450 22002', email: 'gm@sesipl.com' },
      { projectId: p4, level: '3', roleTitle: 'Project Manager', name: 'K. Sharma', phone: '+91 98450 33006', email: 'sharma@sesipl.com' },
      { projectId: p4, level: '4', roleTitle: 'HSE Officer / In charge', name: 'A. Reddy', phone: '+91 98450 44011', email: 'reddy.ehs@sesipl.com' },
      { projectId: p4, level: '5', roleTitle: 'Project Engineer', name: 'T. Venkat', phone: '+91 98450 55012', email: 'venkat.pe@sesipl.com' }
    ];
    orgSeeds.forEach(o => {
      const existing = rowsToObjects_(SHEETS.PROJECT_ORG).find(x => x.projectId === o.projectId && x.roleTitle === o.roleTitle);
      if (existing) {
        updateRowById_(SHEETS.PROJECT_ORG, existing.id, o);
      } else {
        appendRow_(SHEETS.PROJECT_ORG, Object.assign({ id: uid_('ORG') }, o));
      }
    });

    // Seed Slide 8 Standard Daily Logs
    const logs = [
      { projectId: p1, date: '2026-09-15', staff: '6', workers: '48', totalManpower: '54', workingHours: '8', totalManHours: '432', safeManHours: '432', cumSafeManHours: '370840', inductions: '5', indStaff: '1', indWorkers: '4', tbtCount: '3', tbtPersons: '48', trainingTopic: 'Earth pit excavation & isolation', trainingPersons: '18', permitHot: '2', permitElectrical: '4', permitCold: '1', permitGeneral: '3', permitOthers: '0', firstAid: '0', nearMiss: '0', ltiCount: '0', accidentDetails: 'Nil - Safe day', remarks: 'Good site housekeeping maintained', enteredBy: 'EMP001', verifiedBy: 'Ravikiran', approvedBy: 'Mr. Shankar' },
      { projectId: p2, date: '2026-09-15', staff: '4', workers: '38', totalManpower: '42', workingHours: '8', totalManHours: '336', safeManHours: '336', cumSafeManHours: '215420', inductions: '3', indStaff: '0', indWorkers: '3', tbtCount: '2', tbtPersons: '38', trainingTopic: 'Cable laying and tray installation', trainingPersons: '14', permitHot: '1', permitElectrical: '3', permitCold: '0', permitGeneral: '2', permitOthers: '0', firstAid: '0', nearMiss: '0', ltiCount: '0', accidentDetails: 'Nil - Safe day', remarks: 'Cable puller PPE inspected', enteredBy: 'EMP002', verifiedBy: 'K. Sharma', approvedBy: 'Mr. Shankar' },
      { projectId: p3, date: '2026-09-15', staff: '8', workers: '60', totalManpower: '68', workingHours: '8', totalManHours: '544', safeManHours: '544', cumSafeManHours: '410290', inductions: '6', indStaff: '2', indWorkers: '4', tbtCount: '4', tbtPersons: '60', trainingTopic: 'Working at height & full-body harness', trainingPersons: '24', permitHot: '3', permitElectrical: '5', permitCold: '2', permitGeneral: '4', permitOthers: '1', firstAid: '1', nearMiss: '1', ltiCount: '0', accidentDetails: 'Minor finger scrape; treated with first aid kit', remarks: 'First aid kit replenished', enteredBy: 'EMP005', verifiedBy: 'Ravikiran', approvedBy: 'Mr. Shankar' },
      { projectId: p4, date: '2026-09-15', staff: '10', workers: '70', totalManpower: '80', workingHours: '8', totalManHours: '640', safeManHours: '640', cumSafeManHours: '280150', inductions: '8', indStaff: '1', indWorkers: '7', tbtCount: '4', tbtPersons: '70', trainingTopic: 'Substation panel commissioning & LOTO', trainingPersons: '28', permitHot: '2', permitElectrical: '6', permitCold: '1', permitGeneral: '5', permitOthers: '0', firstAid: '0', nearMiss: '0', ltiCount: '0', accidentDetails: 'Nil - Safe day', remarks: 'Substation room entry controlled', enteredBy: 'EMP003', verifiedBy: 'K. Sharma', approvedBy: 'Mr. Shankar' }
    ];
    logs.forEach(l => {
      const existing = rowsToObjects_(SHEETS.DAILY_LOG).find(x => x.projectId === l.projectId && x.date === l.date);
      if (existing) {
        updateRowById_(SHEETS.DAILY_LOG, existing.id, l);
      } else {
        appendRow_(SHEETS.DAILY_LOG, Object.assign({ id: uid_('DL') }, l));
      }
    });

    // Seed Slide 11 Master Tracker of Scaffolding & Ladders
    const scf = [
      { projectId: p1, date: '2026-09-15', region: 'Bangalore', sharavFab: '10', sesipl: '4', vinayaka: '2', hbs: '3', other: '0', totalScaffold: '14', ladderSharav: '12', ladderSesipl: '0', ladderVinayaka: '0', ladderHbs: '0', ladderRental: '0', airportLadder: '2', workstationLadder: '2', frpMsafe: '3', frpYoungman: '2', totalLadder: '9', returnedScaffold: '5', returnedLadder: '3', remarks: '14 Scaffolds in site; 9 Ladders in site', enteredBy: 'EMP001' },
      { projectId: p2, date: '2026-09-15', region: 'Bangalore', sharavFab: '6', sesipl: '3', vinayaka: '1', hbs: '2', other: '0', totalScaffold: '12', ladderSharav: '6', ladderSesipl: '2', ladderVinayaka: '0', ladderHbs: '0', ladderRental: '2', airportLadder: '1', workstationLadder: '1', frpMsafe: '2', frpYoungman: '2', totalLadder: '8', returnedScaffold: '2', returnedLadder: '1', remarks: 'Periodic inspection completed', enteredBy: 'EMP002' },
      { projectId: p3, date: '2026-09-15', region: 'Chennai', sharavFab: '8', sesipl: '4', vinayaka: '1', hbs: '2', other: '0', totalScaffold: '15', ladderSharav: '0', ladderSesipl: '0', ladderVinayaka: '0', ladderHbs: '0', ladderRental: '0', airportLadder: '0', workstationLadder: '0', frpMsafe: '0', frpYoungman: '0', totalLadder: '0', returnedScaffold: '0', returnedLadder: '0', remarks: '15 Scaffolds in site', enteredBy: 'EMP005' },
      { projectId: p4, date: '2026-09-15', region: 'Hyderabad', sharavFab: '5', sesipl: '2', vinayaka: '3', hbs: '0', other: '0', totalScaffold: '10', ladderSharav: '6', ladderSesipl: '0', ladderVinayaka: '0', ladderHbs: '0', ladderRental: '8', airportLadder: '2', workstationLadder: '2', frpMsafe: '4', frpYoungman: '4', totalLadder: '14', returnedScaffold: '0', returnedLadder: '0', remarks: '10 Scaffolds in site; 14 Ladders in site', enteredBy: 'EMP003' }
    ];
    scf.forEach(s => {
      const existing = rowsToObjects_(SHEETS.SCAFFOLD).find(x => x.projectId === s.projectId && x.date === s.date);
      if (existing) {
        updateRowById_(SHEETS.SCAFFOLD, existing.id, s);
      } else {
        appendRow_(SHEETS.SCAFFOLD, Object.assign({ id: uid_('SCF') }, s));
      }
    });

    // Seed Slide 10 PPE Stock Register Data
    const ppeSeeds = [
      { projectId: p1, date: '2026-09-15', contractor: 'Vinayaka Electricals', receivedBy: 'R. Prakash', helmetWhite: '3', helmetGreen: '2', helmetBlue: '3', helmetRed: '2', jacketGreen: '3', jacketOrange: '4', cottonGloves: '20', leatherGloves: '8', goggle: '15', faceShield: '4', mask: '30', apron: '2', shoulderPad: '2', earMuff: '4', dcNo: '274', totalReceived: '10', returnable: 'Yes', balanceStock: '8', remarks: 'Store In-charge: Mr. Harish', enteredBy: 'EMP001' },
      { projectId: p2, date: '2026-09-15', contractor: 'Apex Infra', receivedBy: 'S. Kumar', helmetWhite: '4', helmetGreen: '2', helmetBlue: '2', helmetRed: '0', jacketGreen: '2', jacketOrange: '4', cottonGloves: '15', leatherGloves: '6', goggle: '10', faceShield: '2', mask: '25', apron: '0', shoulderPad: '0', earMuff: '2', dcNo: '281', totalReceived: '12', returnable: 'Yes', balanceStock: '9', remarks: 'Store In-charge: S. Rajesh', enteredBy: 'EMP002' },
      { projectId: p3, date: '2026-09-15', contractor: 'Sharav Power', receivedBy: 'T. Murthy', helmetWhite: '5', helmetGreen: '3', helmetBlue: '4', helmetRed: '2', jacketGreen: '4', jacketOrange: '6', cottonGloves: '30', leatherGloves: '12', goggle: '20', faceShield: '6', mask: '40', apron: '4', shoulderPad: '4', earMuff: '6', dcNo: '295', totalReceived: '18', returnable: 'Yes', balanceStock: '14', remarks: 'Store In-charge: V. Murugan', enteredBy: 'EMP005' },
      { projectId: p4, date: '2026-09-15', contractor: 'Southern Tech', receivedBy: 'G. Naidu', helmetWhite: '4', helmetGreen: '4', helmetBlue: '3', helmetRed: '1', jacketGreen: '5', jacketOrange: '5', cottonGloves: '25', leatherGloves: '10', goggle: '18', faceShield: '4', mask: '35', apron: '2', shoulderPad: '2', earMuff: '5', dcNo: '304', totalReceived: '15', returnable: 'Yes', balanceStock: '11', remarks: 'Store In-charge: A. Reddy', enteredBy: 'EMP003' }
    ];
    ppeSeeds.forEach(p => {
      const existing = rowsToObjects_(SHEETS.PPE).find(x => x.projectId === p.projectId && x.dcNo === p.dcNo);
      if (existing) {
        updateRowById_(SHEETS.PPE, existing.id, p);
      } else {
        appendRow_(SHEETS.PPE, Object.assign({ id: uid_('PPE') }, p));
      }
    });

    // Seed Slide 15 Audit Scores
    const auditSeeds = [
      { projectId: p1, auditDate: '2026-09-10', auditor: 'CBRE Lead Auditor', location: 'Intuit Bellandur', totalScore: '485', maxScore: '500', percent: '97', grade: 'A+', status: 'APPROVED', submittedBy: 'EMP003' },
      { projectId: p2, auditDate: '2026-09-08', auditor: 'Cushman Auditor', location: 'Siemens E-City', totalScore: '460', maxScore: '500', percent: '92', grade: 'A', status: 'APPROVED', submittedBy: 'EMP003' },
      { projectId: p3, auditDate: '2026-09-12', auditor: 'JLL EHS Auditor', location: 'Qualcomm Chennai', totalScore: '475', maxScore: '500', percent: '95', grade: 'A', status: 'APPROVED', submittedBy: 'EMP003' },
      { projectId: p4, auditDate: '2026-09-05', auditor: 'Turner Auditor', location: 'Infosys Hyderabad', totalScore: '450', maxScore: '500', percent: '90', grade: 'A', status: 'APPROVED', submittedBy: 'EMP003' }
    ];
    auditSeeds.forEach(a => {
      const existing = rowsToObjects_(SHEETS.AUDITS).find(x => x.projectId === a.projectId && x.auditDate === a.auditDate);
      if (existing) {
        updateRowById_(SHEETS.AUDITS, existing.id, a);
      } else {
        appendRow_(SHEETS.AUDITS, Object.assign({ id: uid_('ADT') }, a));
      }
    });

    FORM_DEFS.forEach(f => {
      const existing = findOne_(SHEETS.FORM_CATALOG, 'formCode', f.formCode);
      if (existing) {
        updateRowByKey_(SHEETS.FORM_CATALOG, 'formCode', f.formCode, f);
      } else {
        appendRow_(SHEETS.FORM_CATALOG, f);
      }
    });

    // Systematic Deduplication across all database sheets
    deduplicateSheet_(SHEETS.USERS, ['employeeId']);
    deduplicateSheet_(SHEETS.PROJECTS, ['id']);
    deduplicateSheet_(SHEETS.PROJECT_USERS, ['employeeId', 'projectId', 'role']);
    deduplicateSheet_(SHEETS.PROJECT_ORG, ['projectId', 'roleTitle']);
    deduplicateSheet_(SHEETS.FORM_CATALOG, ['formCode']);
    deduplicateSheet_(SHEETS.DAILY_LOG, ['projectId', 'date']);
    deduplicateSheet_(SHEETS.SCAFFOLD, ['projectId', 'date']);
    deduplicateSheet_(SHEETS.PPE, ['projectId', 'dcNo']);
    deduplicateSheet_(SHEETS.AUDITS, ['projectId', 'auditDate']);

    // Format all sheets with professional styling, row heights, column widths, and teal branding
    formatAllSheets_();

    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'checkEscalations')) {
      ScriptApp.newTrigger('checkEscalations').timeBased().everyHours(1).create();
    }

    ss.getRange('A1');
    const result = {
      ok: true,
      spreadsheetUrl: ss.getUrl(),
      spreadsheetId: ssId,
      folderUrl: root.getUrl(),
      sheetCount: ss.getSheets().length,
      userCount: rowsToObjects_(SHEETS.USERS).length,
      projectCount: rowsToObjects_(SHEETS.PROJECTS).length,
      formCount: rowsToObjects_(SHEETS.FORM_CATALOG).length,
      triggerCount: ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'checkEscalations').length,
      demoLogins: userRows.map(u => ({ employeeId: u.employeeId, uan: u.uan, role: u.role, name: u.name }))
    };
    Logger.log(JSON.stringify(result));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function setupStatus() {
  const props = PropertiesService.getScriptProperties();
  const ssId = String(DATABASE_SPREADSHEET_ID || '').trim() || props.getProperty('SPREADSHEET_ID');
  if (ssId) props.setProperty('SPREADSHEET_ID', ssId);
  const rootId = props.getProperty('ROOT_FOLDER_ID');
  const status = {
    spreadsheetId: ssId || '',
    spreadsheetUrl: '',
    rootFolderId: rootId || '',
    rootFolderUrl: '',
    sheets: [],
    users: 0,
    projects: 0,
    forms: 0,
    escalationTriggers: 0
  };
  try {
    if (ssId) {
      const ss = SpreadsheetApp.openById(ssId);
      status.spreadsheetUrl = ss.getUrl();
      status.sheets = ss.getSheets().map(s => s.getName());
      status.users = rowsToObjects_(SHEETS.USERS).length;
      status.projects = rowsToObjects_(SHEETS.PROJECTS).length;
      status.forms = rowsToObjects_(SHEETS.FORM_CATALOG).length;
    }
    if (rootId) status.rootFolderUrl = DriveApp.getFolderById(rootId).getUrl();
    status.escalationTriggers = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'checkEscalations').length;
    status.ok = Boolean(status.spreadsheetId && status.rootFolderId && status.users && status.projects && status.forms);
    status.message = status.ok ? 'System is initialized.' : 'System is incomplete. Run initializeSystem() again.';
  } catch (e) {
    status.ok = false;
    status.message = String(e && e.message ? e.message : e);
  }
  Logger.log(JSON.stringify(status));
  return status;
}

function ensureSheetSchema_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(h => String(h || '').trim());
    const missing = headers.filter(h => current.indexOf(h) < 0);
    if (missing.length) {
      sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  formatSheetProfessionally_(sh, name, headers);
  return sh;
}

function ensureProjectFolder_(project) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
  const name = project.code + ' - ' + project.name;
  const folders = root.getFoldersByName(name);
  const folder = folders.hasNext() ? folders.next() : root.createFolder(name);
  MODULES.forEach(m => {
    const kids = folder.getFoldersByName(m.title);
    if (!kids.hasNext()) folder.createFolder(m.title);
  });
  ['Gallery', 'eLibrary', 'Generated PDFs'].forEach(n => {
    const kids = folder.getFoldersByName(n);
    if (!kids.hasNext()) folder.createFolder(n);
  });
  return folder.getId();
}

function getProjectFolder_(project) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
  const name = project.code + ' - ' + project.name;
  const folders = root.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.getFolderById(ensureProjectFolder_(project));
}

function getNamedSubfolder_(project, name) {
  const pf = getProjectFolder_(project);
  const it = pf.getFoldersByName(name);
  return it.hasNext() ? it.next() : pf.createFolder(name);
}
