const APP_NAME = 'SESIPL EHS Digital';
const DATABASE_SPREADSHEET_ID = '1RklASnOOh_gX7sjvfJTt8eBY0JBgj1iPi0wg-G3O4N0';
const SESSION_TTL_SEC = 21600;
const ESCALATE_HOURS = 24;

const ROLES = {
  LEAD: 'EHS_LEAD',
  ASST: 'ASST_EHS_MANAGER',
  MANAGER: 'EHS_MANAGER',
  DIRECTOR: 'DIRECTOR'
};

const STATUS = {
  DRAFT: 'DRAFT',
  SUBMITTED: 'SUBMITTED',
  APPROVED: 'APPROVED',
  REJECTED: 'REJECTED',
  RESUBMITTED: 'RESUBMITTED',
  OPEN: 'OPEN',
  CLOSED: 'CLOSED',
  FALLBACK: 'FALLBACK'
};

const SHEETS = {
  USERS: 'Users',
  PROJECTS: 'Projects',
  PROJECT_USERS: 'ProjectUsers',
  PROJECT_ORG: 'ProjectOrgChart',
  FORM_CATALOG: 'FormCatalog',
  SUBMISSIONS: 'Submissions',
  TEMPLATE_RECORDS: 'TemplateRecords',
  COMMENTS: 'Comments',
  OBSERVATIONS: 'Observations',
  DAILY_LOG: 'DailyLog',
  PPE: 'PpeIssue',
  SCAFFOLD: 'ScaffoldTracker',
  MOVEMENT: 'ScaffoldMovement',
  AUDITS: 'Audits',
  AUDIT_SCORES: 'AuditScores',
  LIBRARY: 'Library',
  GALLERY: 'Gallery',
  NOTIFICATIONS: 'Notifications',
  TRAINING: 'TrainingCalendar',
  ESCALATIONS: 'Escalations',
  AUDIT_LOG: 'AuditLog'
};

const HEADERS = {
  Users: ['employeeId', 'uan', 'name', 'role', 'email', 'phone', 'active', 'mappedProjects'],
  Projects: ['id', 'code', 'name', 'client', 'pmc', 'inCharge', 'manager', 'scope', 'startDate', 'endDate', 'areaSqft', 'poNo', 'status', 'region', 'projectDuration'],
  ProjectUsers: ['employeeId', 'projectId', 'role'],
  ProjectOrgChart: ['id', 'projectId', 'level', 'roleTitle', 'name', 'phone', 'email'],
  FormCatalog: ['formCode', 'module', 'title', 'cadence', 'entryType', 'slaHours'],
  Submissions: ['id', 'requestId', 'projectId', 'formCode', 'version', 'status', 'submittedBy', 'submittedAt', 'reviewedBy', 'reviewedAt', 'pdfFileId', 'docFileId', 'payloadJson'],
  TemplateRecords: ['id', 'submissionId', 'projectId', 'formCode', 'templateKey', 'version', 'status', 'submittedBy', 'submittedAt', 'reviewedBy', 'reviewedAt', 'payloadJson'],
  Comments: ['id', 'entityType', 'entityId', 'projectId', 'byEmployeeId', 'byName', 'role', 'text', 'createdAt', 'broadcast'],
  Observations: ['id', 'projectId', 'reportNo', 'date', 'vendor', 'auditedBy', 'location', 'contractor', 'observation', 'preventive', 'status', 'ownerEmployeeId', 'dueAt', 'escalateAt', 'fallbackNote', 'fallbackAt', 'unsafeFileId', 'rectifiedFileId', 'submittedBy'],
  DailyLog: ['id', 'projectId', 'date', 'staff', 'workers', 'totalManpower', 'workingHours', 'totalManHours', 'safeManHours', 'cumSafeManHours', 'inductions', 'indStaff', 'indWorkers', 'tbtCount', 'tbtPersons', 'trainingTopic', 'trainingPersons', 'permitHot', 'permitElectrical', 'permitCold', 'permitGeneral', 'permitOthers', 'firstAid', 'nearMiss', 'ltiCount', 'accidentDetails', 'remarks', 'enteredBy', 'verifiedBy', 'approvedBy'],
  PpeIssue: ['id', 'projectId', 'date', 'contractor', 'receivedBy', 'helmetWhite', 'helmetGreen', 'helmetBlue', 'helmetRed', 'jacketGreen', 'jacketOrange', 'cottonGloves', 'leatherGloves', 'goggle', 'faceShield', 'mask', 'apron', 'shoulderPad', 'earMuff', 'dcNo', 'totalReceived', 'returnable', 'balanceStock', 'remarks', 'enteredBy'],
  ScaffoldTracker: ['id', 'projectId', 'date', 'region', 'sharavFab', 'sesipl', 'vinayaka', 'hbs', 'other', 'totalScaffold', 'ladderSharav', 'ladderSesipl', 'ladderVinayaka', 'ladderHbs', 'ladderRental', 'airportLadder', 'workstationLadder', 'frpMsafe', 'frpYoungman', 'totalLadder', 'returnedScaffold', 'returnedLadder', 'remarks', 'enteredBy'],
  ScaffoldMovement: ['id', 'date', 'fromProject', 'fromQty', 'toProject', 'toQty', 'remarks', 'enteredBy'],
  Audits: ['id', 'projectId', 'auditDate', 'auditor', 'location', 'totalScore', 'maxScore', 'percent', 'grade', 'status', 'submittedBy'],
  AuditScores: ['id', 'auditId', 'section', 'sn', 'particulars', 'score', 'remarks'],
  Library: ['id', 'projectId', 'module', 'title', 'fileId', 'uploadedBy', 'uploadedAt', 'tags'],
  Gallery: ['id', 'projectId', 'category', 'title', 'fileId', 'mimeType', 'uploadedBy', 'uploadedAt'],
  Notifications: ['id', 'toEmployeeId', 'projectId', 'title', 'body', 'type', 'read', 'createdAt', 'createdBy'],
  TrainingCalendar: ['id', 'projectId', 'date', 'topic', 'owner', 'dept', 'status', 'notes', 'createdBy'],
  Escalations: ['id', 'entityType', 'entityId', 'level', 'toEmployeeId', 'sentAt', 'reason'],
  AuditLog: ['id', 'at', 'employeeId', 'action', 'entityType', 'entityId', 'detail']
};

const MODULES = [
  { id: 'TEST_CERT', title: 'Test Certificates', cadence: 'ONETIME', icon: 'verified' },
  { id: 'WEEKLY', title: 'Weekly / Monthly Report', cadence: 'WEEKLY', icon: 'calendar' },
  { id: 'PERMIT', title: 'Work Permit', cadence: 'AS_REQUIRED', icon: 'permit' },
  { id: 'EHS_DOCS', title: 'EHS Documents', cadence: 'ONETIME', icon: 'folder' },
  { id: 'LEGAL', title: 'Legal Documents', cadence: 'ONETIME', icon: 'legal' },
  { id: 'OBSERVATION', title: 'Observation', cadence: 'AS_REQUIRED', icon: 'eye' },
  { id: 'POLICY', title: 'Policy & Certificate', cadence: 'ONETIME', icon: 'policy' },
  { id: 'SWMS', title: 'SWMS HIRA', cadence: 'ONETIME', icon: 'hazard' },
  { id: 'CHECKLIST', title: 'Checklist & Formats', cadence: 'DAILY', icon: 'check' }
];

const FORM_DEFS = [
  { formCode: 'TC_UPLOAD', module: 'TEST_CERT', title: 'Test Certificate Upload', cadence: 'ONETIME', entryType: 'UPLOAD', slaHours: 0 },
  { formCode: 'WR_WEEKLY', module: 'WEEKLY', title: 'Weekly EHS Report', cadence: 'WEEKLY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'WR_MONTHLY', module: 'WEEKLY', title: 'Monthly EHS Report', cadence: 'MONTHLY', entryType: 'FORM', slaHours: 48 },
  { formCode: 'WP_HEIGHT', module: 'PERMIT', title: 'Working At Height Permit', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 8 },
  { formCode: 'WP_NIGHT', module: 'PERMIT', title: 'Night Work Permit', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 8 },
  { formCode: 'WP_SHAFT', module: 'PERMIT', title: 'Shaft Work Permit', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 8 },
  { formCode: 'WP_LIFT', module: 'PERMIT', title: 'Lifting Activity Permit', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 8 },
  { formCode: 'WP_HOT', module: 'PERMIT', title: 'Hot Work Permit', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 8 },
  { formCode: 'WP_GENERAL', module: 'PERMIT', title: 'General Work Permit', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 8 },
  { formCode: 'EHS_UPLOAD', module: 'EHS_DOCS', title: 'EHS Document Upload', cadence: 'ONETIME', entryType: 'UPLOAD', slaHours: 0 },
  { formCode: 'LEGAL_UPLOAD', module: 'LEGAL', title: 'Legal Document Upload', cadence: 'ONETIME', entryType: 'UPLOAD', slaHours: 0 },
  { formCode: 'OBS_DAILY', module: 'OBSERVATION', title: 'Daily Safety Observation', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'POL_UPLOAD', module: 'POLICY', title: 'Policy & Certificate Upload', cadence: 'ONETIME', entryType: 'UPLOAD', slaHours: 0 },
  { formCode: 'SWMS_UPLOAD', module: 'SWMS', title: 'SWMS / HIRA Upload', cadence: 'ONETIME', entryType: 'UPLOAD', slaHours: 0 },
  { formCode: 'CL_SCREENING', module: 'CHECKLIST', title: 'Screening of Worker Format', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_IDCARD', module: 'CHECKLIST', title: 'Project ID Card', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_DRILL', module: 'CHECKLIST', title: 'Drilling Machine Checklist', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_WELD', module: 'CHECKLIST', title: 'Welding Machine Checklist', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_TBT', module: 'CHECKLIST', title: 'Tool Box Talk', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_JST', module: 'CHECKLIST', title: 'JST Attendance Sheet', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_GRIND', module: 'CHECKLIST', title: 'Grinding Machine Checklist', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_INDUCTION', module: 'CHECKLIST', title: 'EHS Induction', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_CUT', module: 'CHECKLIST', title: 'Cutting Machine Checklist', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_SCAFFOLD', module: 'CHECKLIST', title: 'Scaffolding Checklist', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_MEDICAL', module: 'CHECKLIST', title: 'Medical Certificate', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'TAG_IND', module: 'CHECKLIST', title: 'Induction Sticker', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'TAG_TOOL', module: 'CHECKLIST', title: 'Tools Inspection Sticker', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'TAG_FE', module: 'CHECKLIST', title: 'Fire Extinguisher Tag', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'TAG_RED', module: 'CHECKLIST', title: 'Red Tag', cadence: 'AS_REQUIRED', entryType: 'FORM', slaHours: 24 },
  { formCode: 'TAG_SCAFF', module: 'CHECKLIST', title: 'Scaffold Inspection Tag', cadence: 'DAILY', entryType: 'FORM', slaHours: 24 },
  { formCode: 'CL_FE', module: 'CHECKLIST', title: 'Fire Extinguisher Checklist', cadence: 'WEEKLY', entryType: 'FORM', slaHours: 24 }
];

const AUDIT_SECTIONS = [
  { id: 'A', name: 'GENERAL REQUIREMENTS & POLICY', max: 30 },
  { id: 'B', name: 'HAZARDS, RISK MANAGEMENT & SWMS', max: 45 },
  { id: 'C', name: 'EHS SCREENING, INDUCTION and TRAINING', max: 35 },
  { id: 'D', name: 'Documents & Records', max: 35 },
  { id: 'E', name: 'EMERGENCY PREPAREDNESS', max: 25 },
  { id: 'F', name: 'EHS PERFORMANCE MEASUREMENT & MOTIVATION', max: 10 },
  { id: 'G', name: 'REPORTING AND INVESTIGATION', max: 30 },
  { id: 'H', name: 'Welfare & Health Management', max: 55 },
  { id: 'I', name: 'Mechanical Management System', max: 25 },
  { id: 'J', name: 'Electrical Management System', max: 40 },
  { id: 'K', name: 'Housekeeping Management', max: 15 },
  { id: 'L', name: 'Waste Management', max: 10 },
  { id: 'M', name: 'Store & Material Management', max: 20 },
  { id: 'N', name: 'Working Platforms and Ladders', max: 20 },
  { id: 'O', name: 'FLOOR DB PANELS', max: 60 },
  { id: 'P', name: 'HOT WORK', max: 40 },
  { id: 'Q', name: 'SHAFT WORK', max: 0 },
  { id: 'R', name: 'PPES', max: 30 }
];

const GALLERY_CATEGORIES = ['COMPANY', 'TRAINING', 'MEETING', 'EVENT'];

function getFormFields_(formCode) {
  const common = [
    { key: 'date', label: 'Date', type: 'date', required: true },
    { key: 'location', label: 'Location / Area', type: 'text', required: true },
    { key: 'supervisor', label: 'Supervisor / In-charge', type: 'text', required: true }
  ];
  const permitExtra = [
    { key: 'permitNo', label: 'Permit No', type: 'text', required: true },
    { key: 'contractor', label: 'Contractor', type: 'text', required: true },
    { key: 'workDescription', label: 'Work Description', type: 'textarea', required: true },
    { key: 'startTime', label: 'Start Time', type: 'time', required: true },
    { key: 'endTime', label: 'End Time', type: 'time', required: true },
    { key: 'hazards', label: 'Hazards Identified', type: 'textarea', required: true },
    { key: 'controls', label: 'Control Measures', type: 'textarea', required: true },
    { key: 'ppe', label: 'PPE Required', type: 'text', required: true },
    { key: 'workersCount', label: 'No. of Workers', type: 'number', required: true },
    { key: 'receiverName', label: 'Permit Receiver', type: 'text', required: true }
  ];
  const machine = [
    { key: 'equipmentId', label: 'Equipment / Asset ID', type: 'text', required: true },
    { key: 'make', label: 'Make / Model', type: 'text', required: false },
    { key: 'conditionOk', label: 'Condition Satisfactory?', type: 'select', options: ['Yes', 'No'], required: true },
    { key: 'guardOk', label: 'Guards / Safety devices OK?', type: 'select', options: ['Yes', 'No', 'NA'], required: true },
    { key: 'cableOk', label: 'Cable / Plug / ELCB OK?', type: 'select', options: ['Yes', 'No', 'NA'], required: true },
    { key: 'operatorTrained', label: 'Operator trained?', type: 'select', options: ['Yes', 'No'], required: true },
    { key: 'defects', label: 'Defects / Remarks', type: 'textarea', required: false },
    { key: 'inspectedBy', label: 'Inspected By', type: 'text', required: true }
  ];
  const dailyInspectionFields = (items) => {
    const fields = common.concat([
      { key: 'equipmentId', label: 'Equipment / Machine No.', type: 'text', required: true },
      { key: 'make', label: 'Make / Type', type: 'text', required: false },
      { key: 'contractor', label: 'Contractor Name', type: 'text', required: true }
    ]);
    items.forEach((item, index) => {
      for (let day = 1; day <= 7; day++) {
        fields.push({ key: 'item' + (index + 1) + 'Day' + day, label: 'Item ' + (index + 1) + ' - Day ' + day + ': ' + item, type: 'select', options: ['Yes', 'No', 'NA'], required: true });
      }
    });
    return fields.concat([
      { key: 'supervisorSign', label: 'Supervisor Name / Signature', type: 'text', required: true },
      { key: 'safetyOfficerSign', label: 'Safety Officer Name / Signature', type: 'text', required: true },
      { key: 'engineerSign', label: 'Electrical / Mechanical Engineer Name / Signature', type: 'text', required: true }
    ]);
  };
  const inductionAttendees = [];
  for (let row = 1; row <= 15; row++) {
    inductionAttendees.push({ key: 'attendee' + row + 'Id', label: 'Attendee ' + row + ' ID Card No.', type: 'text', required: false });
    inductionAttendees.push({ key: 'attendee' + row + 'Name', label: 'Attendee ' + row + ' Name', type: 'text', required: false });
    inductionAttendees.push({ key: 'attendee' + row + 'Designation', label: 'Attendee ' + row + ' Designation', type: 'text', required: false });
  }
  const permitSignatures = [
    { key: 'siteEngineer', label: 'Site Engineer / Permit Requesting Authority', type: 'text', required: true },
    { key: 'safetyOfficer', label: 'Safety Officer', type: 'text', required: true },
    { key: 'contractorInCharge', label: 'Contractor Site In-charge', type: 'text', required: true },
    { key: 'contactNumber', label: 'Emergency / Contact Number', type: 'text', required: true },
    { key: 'validFrom', label: 'Valid From', type: 'datetime-local', required: true },
    { key: 'validTo', label: 'Valid To', type: 'datetime-local', required: true },
    { key: 'approvedByEhs', label: 'Approved By EHS', type: 'text', required: false },
    { key: 'approvedBySiteEngineer', label: 'Approved By Site Engineer', type: 'text', required: false },
    { key: 'closingRemarks', label: 'Permit Closing / Cancellation Remarks', type: 'textarea', required: false }
  ];
  const attendanceRows = (count, prefix) => {
    const fields = [];
    for (let row = 1; row <= count; row++) {
      fields.push({ key: prefix + row + 'Name', label: 'Participant ' + row + ' Name', type: 'text', required: false });
      fields.push({ key: prefix + row + 'Designation', label: 'Participant ' + row + ' Designation', type: 'text', required: false });
      fields.push({ key: prefix + row + 'Company', label: 'Participant ' + row + ' Company / Agency', type: 'text', required: false });
    }
    return fields;
  };
  const map = {
    WR_WEEKLY: common.concat([
      { key: 'weekNo', label: 'Week No', type: 'text', required: true },
      { key: 'manpowerAvg', label: 'Average Manpower', type: 'number', required: true },
      { key: 'manHours', label: 'Man Hours', type: 'number', required: true },
      { key: 'tbtCount', label: 'TBT Conducted', type: 'number', required: true },
      { key: 'inductions', label: 'Inductions', type: 'number', required: true },
      { key: 'permits', label: 'Work Permits Issued', type: 'number', required: true },
      { key: 'observations', label: 'Observations Open / Closed', type: 'text', required: true },
      { key: 'incidents', label: 'Incidents / Near Miss', type: 'textarea', required: false },
      { key: 'highlights', label: 'Highlights / Issues', type: 'textarea', required: true }
    ]),
    WR_MONTHLY: common.concat([
      { key: 'month', label: 'Month', type: 'month', required: true },
      { key: 'cumManHours', label: 'Cumulative Man Hours', type: 'number', required: true },
      { key: 'ltifr', label: 'LTIFR / TRIR notes', type: 'text', required: false },
      { key: 'trainings', label: 'Trainings Conducted', type: 'textarea', required: true },
      { key: 'audits', label: 'Audits / Compliance', type: 'textarea', required: true },
      { key: 'summary', label: 'Monthly Summary', type: 'textarea', required: true }
    ]),
    WP_HEIGHT: common.concat(permitExtra).concat(permitSignatures).concat([
      { key: 'heightM', label: 'Working Height (m)', type: 'number', required: true },
      { key: 'harnessOk', label: 'Full body harness inspected?', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'anchorOk', label: 'Anchor / lifeline OK?', type: 'select', options: ['Yes', 'No'], required: true }
    ]),
    WP_NIGHT: common.concat(permitExtra).concat(permitSignatures).concat([
      { key: 'lightingOk', label: 'Adequate lighting?', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'emergencyContact', label: 'Night emergency contact', type: 'text', required: true }
    ]),
    WP_SHAFT: common.concat(permitExtra).concat(permitSignatures).concat([
      { key: 'shaftId', label: 'Shaft / Lift well ID', type: 'text', required: true },
      { key: 'barricadeOk', label: 'Barricade / cover OK?', type: 'select', options: ['Yes', 'No'], required: true }
    ]),
    WP_LIFT: common.concat(permitExtra).concat(permitSignatures).concat([
      { key: 'loadKg', label: 'Load (kg)', type: 'number', required: true },
      { key: 'craneId', label: 'Crane / Chain block ID', type: 'text', required: true },
      { key: 'rigger', label: 'Rigger / Signalman', type: 'text', required: true }
    ]),
    WP_HOT: common.concat(permitExtra).concat(permitSignatures).concat([
      { key: 'hotType', label: 'Hot work type', type: 'select', options: ['Welding', 'Cutting', 'Grinding', 'Other'], required: true },
      { key: 'fireWatcher', label: 'Fire watcher', type: 'text', required: true },
      { key: 'extinguisherOk', label: 'Fire extinguisher available?', type: 'select', options: ['Yes', 'No'], required: true }
    ]),
    WP_GENERAL: common.concat(permitExtra).concat(permitSignatures),
    OBS_DAILY: [
      { key: 'date', label: 'Date', type: 'date', required: true },
      { key: 'reportNo', label: 'Report No', type: 'text', required: true },
      { key: 'vendor', label: 'Vendor', type: 'text', required: true },
      { key: 'auditedBy', label: 'Audited By', type: 'text', required: true },
      { key: 'contractor', label: 'Contractor', type: 'text', required: true },
      { key: 'location', label: 'Location', type: 'text', required: true },
      { key: 'observation', label: 'Observation', type: 'textarea', required: true },
      { key: 'preventive', label: 'Preventive measures', type: 'textarea', required: true },
      { key: 'ownerEmployeeId', label: 'Responsible Lead (Employee ID)', type: 'text', required: true }
    ],
    CL_SCREENING: common.concat([
      { key: 'workerId', label: 'Worker ID No.', type: 'text', required: true },
      { key: 'contractor', label: 'Contractor / Sub-contractor', type: 'text', required: true },
      { key: 'workerName', label: 'Full Name of Workman', type: 'text', required: true },
      { key: 'fatherName', label: 'Father / Husband Name', type: 'text', required: true },
      { key: 'permanentAddress', label: 'Permanent Address', type: 'textarea', required: true },
      { key: 'presentAddress', label: 'Present Address', type: 'textarea', required: true },
      { key: 'photoFile', label: 'Worker Photo File ID / URL', type: 'text', required: false },
      { key: 'dateOfBirth', label: 'Date of Birth', type: 'date', required: false },
      { key: 'sex', label: 'Sex', type: 'select', options: ['Male', 'Female', 'Other'], required: true },
      { key: 'age', label: 'Age', type: 'number', required: true },
      { key: 'maritalStatus', label: 'Marital Status', type: 'select', options: ['Married', 'Single', 'Widow', 'Widower'], required: true },
      { key: 'childrenCount', label: 'Number of Children', type: 'number', required: false },
      { key: 'motherTongue', label: 'Mother Tongue', type: 'text', required: false },
      { key: 'languages', label: 'Other Languages Known', type: 'text', required: false },
      { key: 'emergencyContact', label: 'Emergency Contact Person', type: 'text', required: true },
      { key: 'identificationMark', label: 'Other Identification Mark', type: 'text', required: false },
      { key: 'vision', label: 'Vision', type: 'select', options: ['Normal', 'Not Normal'], required: true },
      { key: 'visionProblem', label: 'Vision Problem Details', type: 'text', required: false },
      { key: 'health', label: 'Health', type: 'select', options: ['Normal', 'Not Normal'], required: true },
      { key: 'healthProblem', label: 'Health Problem Details', type: 'text', required: false },
      { key: 'weightKg', label: 'Weight (KG)', type: 'number', required: false },
      { key: 'heightCms', label: 'Height (CMS)', type: 'number', required: false },
      { key: 'bloodGroup', label: 'Blood Group', type: 'text', required: false },
      { key: 'educationDetails', label: 'Education Details', type: 'textarea', required: false },
      { key: 'workerDeclarationSignature', label: 'Worker Declaration Signature / Thumb Impression', type: 'text', required: true },
      { key: 'contractorDeclarationSignature', label: 'Contractor Signature', type: 'text', required: true },
      { key: 'suitableEmployment', label: 'Suitable for Employment In', type: 'text', required: true },
      { key: 'siteInCharge', label: 'Site In-charge / Site Engineer', type: 'text', required: true }
    ]),
    CL_IDCARD: common.concat([
      { key: 'projectAddress', label: 'Project Address', type: 'textarea', required: true },
      { key: 'workerName', label: 'Name', type: 'text', required: true },
      { key: 'empNo', label: 'ID Number', type: 'text', required: true },
      { key: 'designation', label: 'Designation', type: 'text', required: true },
      { key: 'dobAge', label: 'DOB / Age', type: 'text', required: true },
      { key: 'bloodGroup', label: 'Blood Group', type: 'text', required: true },
      { key: 'issueDate', label: 'Date of Issue', type: 'date', required: true },
      { key: 'validity', label: 'Validity Up To', type: 'date', required: true },
      { key: 'authoritySign', label: 'Authority Signature', type: 'text', required: true },
      { key: 'screeningDate', label: 'Screening Done On', type: 'date', required: true },
      { key: 'screeningDocumentNo', label: 'Screening Document No.', type: 'text', required: true },
      { key: 'safetyViolation', label: 'Safety Violation Status', type: 'select', options: ['Green', 'Yellow', 'Red'], required: true },
      { key: 'sign', label: 'Signature', type: 'text', required: true }
    ]),
    CL_DRILL: dailyInspectionFields([
      'Equipment double insulated', 'Equipment free from defects', 'Electrical cable free from defects',
      'Industrial plug top available', 'Drilling bit undamaged', 'Drilling bit holder available', 'Switch working'
    ]),
    CL_WELD: dailyInspectionFields([
      'ON / OFF knob undamaged', 'Regulator with indicator', 'Welding cables connected with lugs',
      'Welding cable insulation undamaged', 'Electrode and earthing holders undamaged', 'Industrial plug available',
      'No exposed live electrical parts', 'Trolley wheels undamaged', 'Fire extinguisher and sand bucket available'
    ]),
    CL_GRIND: dailyInspectionFields([
      'Handle free from damage', 'Wheel guard covers three-fourths area', 'Grinding wheel free from crack',
      'Rear handle without damage', 'Cord strain reliever present', 'Trigger switch undamaged',
      'Dead man switch present', 'Electrical wire without cut or joint', 'Plug top provided', 'Machine body undamaged'
    ]),
    CL_CUT: dailyInspectionFields([
      'Cutting blade defined and undamaged', 'Safety guard available and good', 'Lock system for plate and guard',
      'Job clamp / fence available', 'Handle available and good', 'Cable connection undamaged',
      'Dust guard / chip deflector available', 'Machine base undamaged'
    ]),
    CL_TBT: common.concat([
      { key: 'topic', label: 'Topic Discussed', type: 'text', required: true },
      { key: 'conductedBy', label: 'Tool Box Talk Conducted By', type: 'text', required: true },
      { key: 'projectManager', label: 'Project Manager Name', type: 'text', required: true },
      { key: 'keyPoints', label: 'Key Points / Hazards Discussed', type: 'textarea', required: true }
    ]).concat(attendanceRows(25, 'participant')),
    CL_JST: common.concat([
      { key: 'topic', label: 'Training Topic', type: 'text', required: true },
      { key: 'trainingTime', label: 'Training Time', type: 'time', required: true },
      { key: 'conductedBy', label: 'Conducted By', type: 'text', required: true },
      { key: 'acknowledgement', label: 'Training acknowledgement / hazards and precautions', type: 'textarea', required: true },
      { key: 'projectManagerSignature', label: 'Project Manager Signature', type: 'text', required: true },
      { key: 'ehsOfficerSignature', label: 'EHS Officer Signature', type: 'text', required: true }
    ]).concat(attendanceRows(20, 'participant')),
    CL_INDUCTION: common.concat([
      { key: 'contractor', label: 'Contractor Name', type: 'text', required: true },
      { key: 'topicsCovered', label: 'Safety Induction Topics Covered', type: 'textarea', required: true },
      { key: 'safetyOfficer', label: 'Safety Officer Name', type: 'text', required: true },
      { key: 'projectManager', label: 'Project Manager Name', type: 'text', required: true }
    ]).concat(inductionAttendees),
    CL_SCAFFOLD: dailyInspectionFields([
      'All coupler hooks properly installed', 'Proper platform and locking system provided',
      'Inspection tag displayed', 'Scaffold erected on firm base', 'Wheel lock provided and working',
      'Toe board provided', 'Access ladder properly installed', 'Double top and mid railing provided'
    ]),
    CL_MEDICAL: common.concat([
      { key: 'certificateNo', label: 'Certificate Serial No.', type: 'text', required: true },
      { key: 'workerName', label: 'Worker Name', type: 'text', required: true },
      { key: 'fatherName', label: 'Father / Husband Name', type: 'text', required: true },
      { key: 'identificationMarks', label: 'Identification Marks', type: 'textarea', required: false },
      { key: 'sex', label: 'Sex', type: 'select', options: ['Male', 'Female', 'Other'], required: true },
      { key: 'residence', label: 'Residence', type: 'textarea', required: true },
      { key: 'dateOfBirth', label: 'Date of Birth', type: 'date', required: false },
      { key: 'physicalFitness', label: 'Physical Fitness / Fit for Employment', type: 'select', options: ['Fit', 'Not Fit'], required: true },
      { key: 'medicalInspector', label: 'Medical Inspector / CMO', type: 'text', required: true }
    ]),
    TAG_IND: common.concat([{ key: 'workerName', label: 'Name', type: 'text', required: true }, { key: 'inductionId', label: 'Induction ID', type: 'text', required: true }, { key: 'inductionDate', label: 'Date of Induction', type: 'date', required: true }, { key: 'designation', label: 'Designation / Trade', type: 'text', required: true }, { key: 'emergencyNumber', label: 'Emergency Number', type: 'text', required: true }]),
    TAG_TOOL: common.concat([{ key: 'machineNo', label: 'M/C No.', type: 'text', required: true }, { key: 'makeType', label: 'Make / Type', type: 'text', required: true }, { key: 'inspectionDate', label: 'Date of Inspection', type: 'date', required: true }, { key: 'nextDue', label: 'Due Date of Inspection', type: 'date', required: true }, { key: 'inspectionResult', label: 'Inspected OK / Not OK', type: 'select', options: ['OK', 'Not OK'], required: true }, { key: 'signature', label: 'Signature', type: 'text', required: true }]),
    TAG_FE: common.concat([{ key: 'location', label: 'Location', type: 'text', required: true }, { key: 'inspectionDate', label: 'Inspection Date', type: 'date', required: true }, { key: 'nextDue', label: 'Next Inspection Date', type: 'date', required: true }, { key: 'inspectionResult', label: 'Safe for Use', type: 'select', options: ['Safe for Use', 'Do Not Use'], required: true }, { key: 'inspectionRemarks', label: 'Inspection Remarks', type: 'textarea', required: false }]),
    TAG_RED: common.concat([{ key: 'item', label: 'Item / Equipment', type: 'text', required: true }, { key: 'defect', label: 'Defect', type: 'textarea', required: true }, { key: 'isolatedBy', label: 'Isolated By', type: 'text', required: true }]),
    TAG_SCAFF: common.concat([{ key: 'scaffoldId', label: 'Scaffold ID', type: 'text', required: true }, { key: 'tagColour', label: 'Tag Colour', type: 'select', options: ['Green - Safe for Use', 'Red - Do Not Use'], required: true }, { key: 'inspectionDate', label: 'Inspection Date', type: 'date', required: true }, { key: 'inspectionRemarks', label: 'Inspection Record / Remarks', type: 'textarea', required: false }]),
    CL_FE: common.concat([
      { key: 'extinguisherNo', label: 'Extinguisher No.', type: 'text', required: true },
      { key: 'ehsName', label: 'EHS Name', type: 'text', required: true },
      { key: 'inspectionDate', label: 'Inspection Date', type: 'date', required: true },
      { key: 'nextInspectionDate', label: 'Next Inspection Date', type: 'date', required: true },
      { key: 'extinguisherSpecification', label: 'Extinguisher Specifications / Type', type: 'select', options: ['ABC Type', 'Water Type', 'CO2 Type'], required: true },
      { key: 'conditionClean', label: 'Extinguisher clean and tidy', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'corroded', label: 'Extinguisher corroded', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'safetyPin', label: 'Safety pin in locking position', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'nozzleCondition', label: 'Discharge nozzle condition', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'hoseCondition', label: 'Hose condition', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'weightMatches', label: 'Weight matches body marking', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'visualBoard', label: 'Visual board available', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'gaugeGreen', label: 'Indicator gauge in green', type: 'select', options: ['Yes', 'No', 'NA'], required: true },
      { key: 'expiryStatus', label: 'Not expired / refill date valid', type: 'select', options: ['Yes', 'No'], required: true },
      { key: 'otherRemarks', label: 'Other Remarks', type: 'textarea', required: false },
      { key: 'checkedBy', label: 'Checked By Safety Officer', type: 'text', required: true },
      { key: 'reviewStatus', label: 'Reviewed Status', type: 'select', options: ['Accepted', 'Rejected'], required: true }
    ]),
  };
  if (formCode.indexOf('_UPLOAD') !== -1) {
    return [
      { key: 'title', label: 'Document Title', type: 'text', required: true },
      { key: 'docDate', label: 'Document Date', type: 'date', required: false },
      { key: 'expiry', label: 'Expiry Date', type: 'date', required: false },
      { key: 'remarks', label: 'Remarks', type: 'textarea', required: false }
    ];
  }
  return map[formCode] || common.concat([{ key: 'remarks', label: 'Remarks', type: 'textarea', required: false }]);
}
