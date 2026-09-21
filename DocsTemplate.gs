/* =========================================================
   DOCS TEMPLATE ENGINE & PDF CONVERTER
   Fetches Google Doc templates, substitutes {{placeholders}}
   with form data, and converts to PDF for manager/director viewing.
   ========================================================= */

/**
 * Official Google Doc Template IDs directly configured in the backend.
 * Linked to the user's provided official templates.
 */
const BACKEND_DOCS_TEMPLATES = {
  // Work Permits (Safety Work Clearance)
  WP_HEIGHT: '1t-SjqzbVnBHORMCvk_bAEZYeg3M4GpYl',  // Working At Height Permit-04
  WP_GENERAL: '1xB1mCI7UlmPBT0losFIVL2iXZFrIh6E-', // General work permit-02
  WP_HOT: '122BRc7qu7FWQXs06X3g50DjbusrZWiNq',     // Hot Work Permit-03
  WP_LIFT: '1oxR-cGv7bU6pt1lrl1qCkx8_67X28g5S',    // Lifting Activity Permit-05
  WP_SHAFT: '1tr2XbzrGmtJ86BRE_B9nnRDpV-TSr5KY',   // Shaft work permit-06
  WP_NIGHT: '1g_z99d2ROzrGRsAwK96tUN9Yrg7D-fwy',   // Night Work permit-07

  // Checklists & Equipment Inspections
  CL_WELD: '1iQC5RiaqPexVGQ1E6_8eZfns_nVPxnoy',    // SESIPL-EHS-Welding Machine (1)
  CL_GRIND: '1a1g3LBj887C6NC20DX9W1x7tl04ukSuY',   // SESIPLEHS-Grinding Machine Checklist
  CL_CUT: '',                                      // Cutting Machine Checklist
  CL_DRILL: '1-NMUXI9ACdSJfnuBaRrkvoGIpHEYJ0AO',   // SESIPL-EHS Drilling Machine
  CL_FE: '1GOdBINiB8ZgbRSjGvXWT9GCmf0g0ksw5',      // SESIPL-EHS - Fire Extinguisher
  CL_SCAFFOLD: '1fDFw_miYiWM9AcHKBbGqUtTIP-NPyVW3',// Scaffolding checklist

  // Attendances, Training & Inductions
  CL_TBT: '1qCc1BzpLnof_TTaqVFVlp7WgSsaQlmUr',      // SESIPL-EHS-Tool Box Talk
  CL_JST: '1egxHNPpQ3S8BU5EGHHLPrK0fujWAjD6P',      // SESIPL-EHS-JST Attendance sheet
  CL_INDUCTION: '',                                // EHS Induction

  // Worker Screening & Medical
  CL_SCREENING: '1v7WlDOPl7176yGCLWkWY6xghB3agH0Ky',// Screening of Worker Format
  CL_MEDICAL: '1CwgjUnuPptlNYTaO3abxZEuCQo5BVLvm',  // Medical certificate-xI

  // Stickers / Tags
  TAG_IND: '',
  TAG_TOOL: '',
  TAG_FE: '',
  TAG_RED: ''
};

/**
 * Resolves the Google Doc template ID for a given formCode,
 * checking runtime properties first and falling back to backend defaults.
 */
function getDocTemplateId_(formCode) {
  let id = '';
  try {
    if (typeof getDocTemplateRegistry_ === 'function') {
      const reg = getDocTemplateRegistry_();
      id = (reg && reg[formCode]) || '';
    }
  } catch (e) {}
  if (!id || id.trim() === '') {
    id = BACKEND_DOCS_TEMPLATES[formCode] || '';
  }
  return id ? id.trim() : '';
}

/**
 * Builds a comprehensive dictionary of placeholder replacements.
 * Matches all standard and specific tags for permits, checklists, and worker forms.
 */
function buildPlaceholderMap_(formCode, fields, project, user, submissionId) {
  fields = fields || {};
  project = project || {};
  user = user || {};

  const docCode = (typeof PERMIT_DOC_CODES !== 'undefined' && PERMIT_DOC_CODES[formCode]) || formCode || 'SESIPL-EHS';
  const contractorName = fields.contractorName || fields.contractor || project.client || 'Shankar Electricals Services I Pvt Ltd';
  const permitNo = fields.permitNo || ('SESIPL/' + formCode.replace('WP_', '') + '/' + (submissionId ? String(submissionId).slice(-4) : '001'));
  const emergency1 = fields.emergencyContact1 || '9591461971';
  const emergency2 = fields.emergencyContact2 || '7899650058';

  const map = {
    // Header & Document Info
    contractorName: contractorName,
    projectName: project.name || 'SESIPL Site',
    projectCode: project.code || 'PRJ',
    permitNo: permitNo,
    docCode: docCode,
    emergencyContact1: emergency1,
    emergencyContact2: emergency2,

    // General Work Details
    area: fields.area || '',
    location: fields.location || '',
    date: fields.date ? displayDate_(fields.date) : (fields.workExecutionDate ? displayDate_(fields.workExecutionDate) : todayDisplay_()),
    time: fields.time || '',
    siteEngineerName: fields.siteEngineerName || '',
    siteEngineerSign: fields.siteEngineerSign || (fields.siteEngineerName ? 'Signed' : ''),
    safetyOfficerName: fields.safetyOfficerName || '',
    safetyOfficerSign: fields.safetyOfficerSign || (fields.safetyOfficerName ? 'Signed' : ''),
    contractorSiteIncharge: fields.contractorSiteIncharge || fields.agencySupervisor || '',
    contactNumber: fields.contactNumber || '',
    descriptionOfWork: fields.descriptionOfWork || fields.workDescription || '',
    workExecutionDate: fields.workExecutionDate ? displayDate_(fields.workExecutionDate) : '',
    validFrom: fields.validFrom || '',
    validTo: fields.validTo || '',

    // Permit Specific
    shaftWorkmenNames: fields.workmenNames || fields.shaftWorkmenNames || '',
    hot_other: fields.hot_other || '',
    lift_other: fields.lift_other || '',
    gen_other: fields.gen_other || '',
    height_other: fields.height_other || '',
    night_remarks: fields.night_remarks || '',

    // Reviewed & Approved By
    approvalEhsName: fields.approvalEhsName || '',
    approvalEhsSign: fields.approvalEhsSign || (fields.approvalEhsName ? 'Signed' : ''),
    approvalEhsDate: fields.approvalEhsDate ? displayDate_(fields.approvalEhsDate) : '',
    approvalEhsTime: fields.approvalEhsTime || '',
    approvalSiteEngName: fields.approvalSiteEngineerName || fields.approvalSiteEngName || '',
    approvalSiteEngSign: fields.approvalSiteEngineerSign || fields.approvalSiteEngSign || (fields.approvalSiteEngineerName ? 'Signed' : ''),
    approvalSiteEngDate: fields.approvalSiteEngineerDate ? displayDate_(fields.approvalSiteEngineerDate) : '',
    approvalSiteEngTime: fields.approvalSiteEngineerTime || '',

    // Permit Closing / Cancellation
    closingSiteEngName: fields.closingSiteEngName || '',
    closingSiteEngSign: fields.closingSiteEngSign || (fields.closingSiteEngName ? 'Signed' : ''),
    closingSiteEngDate: fields.closingSiteEngDate ? displayDate_(fields.closingSiteEngDate) : '',
    closingSiteEngTime: fields.closingSiteEngTime || '',
    closingSafetyOfficerName: fields.closingSafetyOfficerName || '',
    closingSafetyOfficerSign: fields.closingSafetyOfficerSign || (fields.closingSafetyOfficerName ? 'Signed' : ''),
    closingSafetyOfficerDate: fields.closingSafetyOfficerDate ? displayDate_(fields.closingSafetyOfficerDate) : '',
    closingSafetyOfficerTime: fields.closingSafetyOfficerTime || '',
    closingPmcSiteEngName: fields.closingPmcSiteEngName || '',
    closingPmcSiteEngSign: fields.closingPmcSiteEngSign || (fields.closingPmcSiteEngName ? 'Signed' : ''),
    closingPmcSiteEngDate: fields.closingPmcSiteEngDate ? displayDate_(fields.closingPmcSiteEngDate) : '',
    closingPmcSiteEngTime: fields.closingPmcSiteEngTime || '',

    // Worker Screening & Medical
    workerId: fields.workerId || '',
    workerFullName: fields.workerName || fields.workerFullName || '',
    fatherOrHusbandName: fields.fatherName || fields.fatherOrHusbandName || '',
    permanentAddress: fields.permanentAddress || '',
    presentAddress: fields.presentAddress || '',
    dob: fields.dateOfBirth ? displayDate_(fields.dateOfBirth) : (fields.dob ? displayDate_(fields.dob) : ''),
    sex: fields.sex || '',
    age: fields.age || '',
    maritalStatus: fields.maritalStatus || '',
    numChildren: fields.childrenCount || fields.numChildren || '',
    motherTongue: fields.motherTongue || '',
    otherLanguages: fields.languages || fields.otherLanguages || '',
    emergencyContactPerson: fields.emergencyContact || fields.emergencyContactPerson || '',
    identificationMark: fields.identificationMark || fields.identificationMarks || '',
    visionStatus: fields.vision || 'Normal',
    visionProblem: fields.visionProblem || '',
    healthStatus: fields.health || 'Normal',
    healthProblem: fields.healthProblem || '',
    weightKg: fields.weightKg || '',
    heightCm: fields.heightCms || fields.heightCm || '',
    bloodGroup: fields.bloodGroup || '',
    suitableEmployment: fields.suitableEmployment || '',
    siteInchargeName: fields.siteInCharge || fields.siteInchargeName || '',
    workerSign: fields.workerDeclarationSignature || 'Signed',
    contractorSign: fields.contractorDeclarationSignature || 'Signed',

    // Form XI Medical
    certSerialNo: fields.certificateNo || fields.certSerialNo || '',
    workerThumbSign: fields.workerDeclarationSignature || 'Thumb Impressed',
    medicalOfficerSign: fields.medicalInspector || 'Verified & Sealed',

    // TBT & Training
    topicDiscussed: fields.topic || fields.topicDiscussed || '',
    tbtConductedBy: fields.conductedBy || fields.tbtConductedBy || '',
    tbtConductedBySign: fields.conductedBy ? 'Signed' : '',
    projectManagerName: fields.projectManager || fields.projectManagerName || '',
    projectManagerSign: fields.projectManager ? 'Signed' : '',
    ehsOfficerSign: 'Signed',

    // Equipment Details
    equipmentId: fields.equipmentId || fields.equipmentNo || fields.machineNo || '',
    equipmentNo: fields.equipmentNo || fields.equipmentId || fields.machineNo || '',
    machineNo: fields.machineNo || fields.equipmentId || fields.equipmentNo || '',
    makeType: fields.makeType || fields.make || '',
    inspectionDate: fields.inspectionDate ? displayDate_(fields.inspectionDate) : '',
    nextInspectionDate: fields.nextInspectionDate ? displayDate_(fields.nextInspectionDate) : (fields.nextDue ? displayDate_(fields.nextDue) : '')
  };

  // Process all keys in fields directly
  for (const k in fields) {
    if (fields[k] != null && !map[k]) {
      map[k] = String(fields[k]);
    }
  }

  // Generate Yes / No / NA / Not Required checkbox indicators
  for (const k in fields) {
    const val = String(fields[k] || '').trim().toLowerCase();
    const isYes = val === 'yes' || val === 'y' || val === 'true';
    const isNo = val === 'no' || val === 'n';
    const isNA = val === 'na' || val === 'n/a';
    const isNR = val === 'not required' || val === 'nr';

    // Form-specific prefix (e.g. height_q1_yes, night_q1_no)
    map[k + '_yes'] = isYes ? '✓' : '';
    map[k + '_no'] = isNo ? '✓' : '';
    map[k + '_na'] = isNA ? '✓' : '';
    map[k + '_nr'] = isNR ? '✓' : '';

    // Generic prefix (e.g. height_q1 -> q1_yes, 1_yes)
    const match = k.match(/^(?:height|gen|hot|lift|shaft|night|wm|grind|cut|drill|fe)_?(q?\d+)/i);
    if (match) {
      const numPart = match[1].toLowerCase().replace('q', '');
      const qTag = 'q' + numPart;
      map[qTag + '_yes'] = isYes ? '✓' : (map[qTag + '_yes'] || '');
      map[qTag + '_no'] = isNo ? '✓' : (map[qTag + '_no'] || '');
      map[qTag + '_na'] = isNA ? '✓' : (map[qTag + '_na'] || '');
      map[qTag + '_nr'] = isNR ? '✓' : (map[qTag + '_nr'] || '');
      map[qTag] = fields[k];

      map[numPart + '_yes'] = isYes ? '✓' : (map[numPart + '_yes'] || '');
      map[numPart + '_no'] = isNo ? '✓' : (map[numPart + '_no'] || '');
      map[numPart + '_na'] = isNA ? '✓' : (map[numPart + '_na'] || '');
      map[numPart + '_nr'] = isNR ? '✓' : (map[numPart + '_nr'] || '');
    }
  }

  // Handle daily inspection checklists like item1Day1
  for (let item = 1; item <= 15; item++) {
    for (let day = 1; day <= 7; day++) {
      const fieldKey = 'item' + item + 'Day' + day;
      if (fields[fieldKey]) {
        const v = String(fields[fieldKey]).trim();
        const isTick = v.toLowerCase() === 'yes' || v.toLowerCase() === 'y' || v === '✓';
        const displayVal = isTick ? '✓' : v;
        map['wm_q' + item + '_d' + day] = displayVal;
        map['grind_q' + item + '_d' + day] = displayVal;
        map['cut_q' + item + '_d' + day] = displayVal;
        map['q' + item + '_d' + day] = displayVal;
        map['item' + item + '_d' + day] = displayVal;
      }
    }
  }

  // Handle participant rows for TBT, JST, Training, and Induction
  for (let i = 1; i <= 30; i++) {
    const name = fields['participant' + i + 'Name'] || fields['tbt_name_' + i] || fields['train_name_' + i] || '';
    const desig = fields['participant' + i + 'Designation'] || fields['tbt_desig_' + i] || fields['train_desig_' + i] || '';
    const agency = fields['participant' + i + 'Company'] || fields['participant' + i + 'Agency'] || fields['tbt_agency_' + i] || fields['train_company_' + i] || '';
    const sign = name ? (fields['participant' + i + 'Signature'] || 'Signed') : '';

    map['tbt_name_' + i] = name;
    map['tbt_desig_' + i] = desig;
    map['tbt_agency_' + i] = agency;
    map['tbt_sign_' + i] = sign;

    map['train_name_' + i] = name;
    map['train_desig_' + i] = desig;
    map['train_company_' + i] = agency;
    map['train_sign_' + i] = sign;

    const indId = fields['attendee' + i + 'Id'] || fields['ind_id_' + i] || '';
    const indName = fields['attendee' + i + 'Name'] || fields['ind_name_' + i] || '';
    const indDesig = fields['attendee' + i + 'Designation'] || fields['ind_desig_' + i] || '';
    const indSign = indName ? (fields['attendee' + i + 'Signature'] || 'Signed') : '';

    map['ind_id_' + i] = indId;
    map['ind_name_' + i] = indName;
    map['ind_desig_' + i] = indDesig;
    map['ind_sign_' + i] = indSign;
  }

  // Generate uppercase, lowercase, and capitalized aliases for every placeholder
  const finalMap = {};
  for (const k in map) {
    const val = map[k];
    finalMap[k] = val;
    finalMap[k.toLowerCase()] = val;
    finalMap[k.toUpperCase()] = val;
    const cap = k.charAt(0).toUpperCase() + k.slice(1);
    finalMap[cap] = val;
  }

  return finalMap;
}

/**
 * Copies a Google Doc template, substitutes {{placeholders}}, and exports to PDF.
 */
function generatePdfFromDocsTemplate_(templateDocId, placeholderMap, title, project) {
  if (!templateDocId) {
    throw new Error("No Google Doc template ID provided.");
  }

  const templateFile = DriveApp.getFileById(templateDocId);
  if (!templateFile) {
    throw new Error("Template Google Doc not accessible with ID: " + templateDocId);
  }

  // Create working copy in designated folder
  const folder = getEhsGeneratedPdfFolder_(project);
  const copyTitle = (title || 'EHS_Document') + '_' + new Date().getTime();
  const docCopy = templateFile.makeCopy(copyTitle, folder);
  const doc = DocumentApp.openById(docCopy.getId());

  // Replace placeholders in Document Body, Header, Footer, and all Table cells
  replacePlaceholdersInDoc_(doc, placeholderMap);
  doc.saveAndClose();

  // Convert populated doc copy to PDF
  const pdfBlob = docCopy.getAs('application/pdf').setName((title || 'EHS_Document') + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);

  try {
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log("Notice: PDF sharing setting: " + err);
  }

  // Remove the temporary doc copy to keep Google Drive uncluttered
  try {
    docCopy.setTrashed(true);
  } catch (cleanErr) {
    Logger.log("Notice: Cleaning up temporary doc copy: " + cleanErr);
  }

  return {
    ok: true,
    fileId: pdfFile.getId(),
    pdfUrl: pdfFile.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
    previewUrl: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/preview'
  };
}

/**
 * Replaces {{placeholder}} tokens across doc body, header, footer, and tables.
 */
function replacePlaceholdersInDoc_(doc, placeholderMap) {
  const replaceInElement = (element) => {
    if (!element) return;
    for (const key in placeholderMap) {
      const val = placeholderMap[key] != null ? String(placeholderMap[key]) : '';
      const regexPattern = '\\{\\{\\s*' + escapeRegex_(key) + '\\s*\\}\\}';
      try {
        element.replaceText(regexPattern, val);
      } catch (e) {
        // Continue
      }
    }
  };

  replaceInElement(doc.getBody());
  if (doc.getHeader()) replaceInElement(doc.getHeader());
  if (doc.getFooter()) replaceInElement(doc.getFooter());

  // Explicitly traverse all tables in the body to guarantee all cells are updated
  try {
    const tables = doc.getBody().getTables();
    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      const numRows = table.getNumRows();
      for (let r = 0; r < numRows; r++) {
        const row = table.getRow(r);
        const numCells = row.getNumCells();
        for (let c = 0; c < numCells; c++) {
          const cell = row.getCell(c);
          replaceInElement(cell);
        }
      }
    }
  } catch (tableErr) {
    Logger.log("Notice: Table traversal in replacePlaceholdersInDoc_: " + tableErr);
  }
}

function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Ensures a dedicated folder for generated PDFs exists in Google Drive.
 */
function getEhsGeneratedPdfFolder_(project) {
  try {
    if (project && project.id) {
      return getNamedSubfolder_(project, 'Generated PDFs');
    }
  } catch (e) {}

  const folderName = 'SESIPL EHS Generated PDFs';
  const existing = DriveApp.getFoldersByName(folderName);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * Resolves or generates the PDF for any submission.
 * If a Google Doc Template ID is configured for this form, generates from Google Docs.
 * If not, falls back gracefully to high-fidelity HTML PDF generation.
 */
function getOrGenerateSubmissionPdf_(submissionId, options) {
  options = options || {};
  const sub = findOne_(SHEETS.SUBMISSIONS, 'id', submissionId);
  if (!sub) throw new Error("Submission not found: " + submissionId);

  const project = findOne_(SHEETS.PROJECTS, 'id', sub.projectId) || { id: sub.projectId, name: 'SESIPL Project Site', code: 'PRJ' };
  const def = FORM_DEFS.find(f => f.formCode === sub.formCode) || { formCode: sub.formCode, title: sub.formCode };
  const fields = (typeof sub.payload === 'object' && sub.payload) ? sub.payload : JSON.parse(sub.payloadJson || '{}');
  const user = { employeeId: sub.submittedBy, name: sub.submittedBy };

  const templateDocId = getDocTemplateId_(sub.formCode);

  // 1. If a Google Docs template ID is configured for this form:
  if (templateDocId && templateDocId.trim() !== '') {
    const placeholderMap = buildPlaceholderMap_(sub.formCode, fields, project, user, submissionId);
    const title = (project.code || 'PRJ') + '_' + sub.formCode + '_' + submissionId;
    const docResult = generatePdfFromDocsTemplate_(templateDocId.trim(), placeholderMap, title, project);

    // Save pdfFileId to submission record for fast subsequent access
    try {
      sub.pdfFileId = docResult.fileId;
      updateOne_(SHEETS.SUBMISSIONS, 'id', sub.id, { pdfFileId: docResult.fileId });
    } catch (dbErr) {
      Logger.log("Notice: Updating submission with pdfFileId: " + dbErr);
    }

    return {
      ok: true,
      source: 'GOOGLE_DOCS',
      title: def.title,
      submission: sub,
      fileId: docResult.fileId,
      pdfUrl: docResult.pdfUrl,
      previewUrl: docResult.previewUrl,
      downloadUrl: docResult.downloadUrl
    };
  }

  // 2. Fallback: High-fidelity HTML PDF generation
  const isPermit = sub.formCode && sub.formCode.indexOf('WP_') === 0;
  const html = isPermit
    ? buildWorkPermitPdfHtml_(project, def, fields, user, sub.version || 1, options)
    : buildFormPdfHtml_(project, def, fields, user, sub.version || 1);

  return {
    ok: true,
    source: 'HTML_FALLBACK',
    title: def.title,
    submission: sub,
    html: html
  };
}
