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
  Notifications: ['id', 'toEmployeeId', 'projectId', 'title', 'body', 'type', 'read', 'createdAt', 'createdBy', 'dismissed'],
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

const PERMIT_DOC_CODES = {
  WP_SHAFT: 'SESIPL-EHS.Blr Prj-06',
  WP_NIGHT: 'SESIPL-EHS.Blr Prj-07',
  WP_LIFT: 'SESIPL-EHS.Blr Prj-05',
  WP_HOT: 'SESIPL-EHS.Blr Prj-03',
  WP_HEIGHT: 'SESIPL-EHS.Blr Prj-04',
  WP_GENERAL: 'SESIPL-EHS.Blr Prj-01'
};

function getFormFields_(formCode) {
  const common = [
    { key: 'date', label: 'Date', type: 'date', required: true },
    { key: 'location', label: 'Location / Area', type: 'text', required: true },
    { key: 'supervisor', label: 'Supervisor / In-charge', type: 'text', required: true }
  ];

  // Official SESIPL Work Permit Header & Authority Fields (1:1 with scanned PDFs)
  const permitHeaderFields = [
    { key: 'contractorName', label: 'Contractor Name', type: 'text', required: true },
    { key: 'emergencyContact1', label: 'Emergency Contact No. 1', type: 'text', required: true },
    { key: 'emergencyContact2', label: 'Emergency Contact No. 2', type: 'text', required: false },
    { key: 'permitNo', label: 'Permit No', type: 'text', required: true },
    { key: 'area', label: 'Area', type: 'text', required: true },
    { key: 'location', label: 'Location', type: 'text', required: true },
    { key: 'date', label: 'Date', type: 'date', required: true },
    { key: 'time', label: 'Time', type: 'time', required: true },
    { key: 'siteEngineer', label: 'Name of Site Engineer (Permit Requesting Authority)', type: 'text', required: true },
    { key: 'siteEngineerSign', label: 'Site Engineer Sign / Ack', type: 'text', required: false },
    { key: 'safetyOfficer', label: 'Name of Safety Officer', type: 'text', required: true },
    { key: 'safetyOfficerSign', label: 'Safety Officer Sign / Ack', type: 'text', required: false },
    { key: 'contractorInCharge', label: 'Name of Contractor Site In charge', type: 'text', required: true },
    { key: 'contactNumber', label: 'Contact Number', type: 'text', required: true },
    { key: 'workDescription', label: 'Description of work', type: 'textarea', required: true },
    { key: 'workExecutionDate', label: 'Work Execution Date', type: 'date', required: true },
    { key: 'validFrom', label: 'Valid From', type: 'datetime-local', required: true },
    { key: 'validTo', label: 'Valid To', type: 'datetime-local', required: true }
  ];

  const permitApprovalFields = [
    { key: 'approvalEhsName', label: 'Reviewed & Approved By EHS (Name)', type: 'text', required: false },
    { key: 'approvalEhsSign', label: 'EHS Sign', type: 'text', required: false },
    { key: 'approvalEhsDate', label: 'EHS Approval Date', type: 'date', required: false },
    { key: 'approvalEhsTime', label: 'EHS Approval Time', type: 'time', required: false },
    { key: 'approvalSiteEngineerName', label: 'Reviewed & Approved By Site Engineer (Name)', type: 'text', required: false },
    { key: 'approvalSiteEngineerSign', label: 'Site Engineer Sign', type: 'text', required: false },
    { key: 'approvalSiteEngineerDate', label: 'Site Engineer Approval Date', type: 'date', required: false },
    { key: 'approvalSiteEngineerTime', label: 'Site Engineer Approval Time', type: 'time', required: false }
  ];

  const permitClosingFields = [
    { key: 'closingSiteEngName', label: 'Closing Site Engineer (Requesting Authority)', type: 'text', required: false },
    { key: 'closingSiteEngSign', label: 'Site Engineer Closing Sign', type: 'text', required: false },
    { key: 'closingSiteEngDate', label: 'Closing Date', type: 'date', required: false },
    { key: 'closingSiteEngTime', label: 'Closing Time', type: 'time', required: false },
    { key: 'closingSafetyOfficerName', label: 'Closing Safety Officer', type: 'text', required: false },
    { key: 'closingSafetyOfficerSign', label: 'Safety Officer Closing Sign', type: 'text', required: false },
    { key: 'closingSafetyOfficerDate', label: 'Safety Officer Closing Date', type: 'date', required: false },
    { key: 'closingSafetyOfficerTime', label: 'Safety Officer Closing Time', type: 'time', required: false },
    { key: 'closingPmcSiteEngName', label: 'Closing PMC Site Engineer (Issuing Authority)', type: 'text', required: false },
    { key: 'closingPmcSiteEngSign', label: 'PMC Engineer Closing Sign', type: 'text', required: false },
    { key: 'closingPmcSiteEngDate', label: 'PMC Engineer Closing Date', type: 'date', required: false },
    { key: 'closingPmcSiteEngTime', label: 'PMC Engineer Closing Time', type: 'time', required: false }
  ];

  // Specific Precautions Checklists from Scanned SESIPL Forms
  const shaftPrecautions = [
    { key: 'shaft_q1', label: '1. Proper Access/ Exit available', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'shaft_q2', label: '2. Proper ventilation and / or lighting provided', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'shaft_q3', label: '3. Proper & Safe platform provided', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'shaft_q4', label: '4. Workers have been briefed about hazardous', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'shaft_q5', label: '5. All Electrical Tools and machinery checked prior to use.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'shaft_q6', label: '6. Shaft area Properly barricaded.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'shaft_q7', label: '7. Conducted JST for all workers who are all engaging to shaft work.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'workmenNames', label: 'Names of workmen entering shaft', type: 'textarea', required: true }
  ];

  const nightPrecautions = [
    { key: 'night_q1', label: '1. Is dedicated Night shift in charge available?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q2', label: '2. Is supervisor available in night shift to supervise the task?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q3', label: '3. Is First aider available?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q4', label: '4. Is Ambulance available for emergency?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q5', label: '5. Are the workers working continuously for last 12 hours?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q6', label: '6. Is the work area safe for work?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q7', label: '7. Is there proper illumination provided at the work area?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q8', label: '8. Are the hazards related with the work identified and assessed at workplace?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q9', label: '9. Is toolbox talk / pre-start briefing carried out prior to start night shift?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q10', label: '10. Are the workers having specific PPE’s according to the requirement of the task?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_q11', label: '11. Is there any high-risk activity like working at height, work in penetration and shafts, electrical testing and commissioning, hot work, Mechanical lifting operation, excavation etc. to be carried out in night shift?', type: 'select', options: ['Yes', 'NO', 'NA'], required: true },
    { key: 'night_remarks', label: 'Night Shift Precautions Remarks', type: 'textarea', required: false }
  ];

  const liftPrecautions = [
    { key: 'lift_q1', label: '1. Crane used for lifting activity tested, certified and approved for rated lifting works.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q2', label: '2. All lifting tackles, gears/ appliances are tested and certified for lifting works.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q3', label: '3. Crane operator is trained and competent for lifting operation.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q4', label: '4. Lifting belt protected against sharp edge of jobs to be lifted.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q5', label: '5. Access and exist marked and without obstruction.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q6', label: '6. Lighting arrangement adequate.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q7', label: '7. Unwanted and rubbish material removed from working platform.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q8', label: '8. Guidelines has provided for balancing & guiding jobs to be lifted.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q9', label: '9. Periphery area of crane booms as well lifting job is barricaded .', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q10', label: '10. Rigger and signal man is trained and competent for lifting work.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q11', label: '11. No lifting activity to be carried during lightening, heavy wind /rain.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q12', label: '12. If scaffolding to be used during lift , Scaffolding with valid tag available for use', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q13', label: '13. Double lanyards Safety Harness/belt checked and in working condition', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_q14', label: '14. Safety shoes (nonslip), Helmet with chin strip available with employees.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'lift_other', label: '15. Other Precautions / Notes', type: 'text', required: false }
  ];

  const hotPrecautions = [
    { key: 'hot_q1', label: '1. Proper Access/ Exit available', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q2', label: '2. Proper ventilation and / or lighting provided', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q3', label: '3. Proper & Safe scaffolding, platform, ladder provided', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q4', label: '4. Welding machine located in a clean and dry area', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q5', label: '5. Welding machine grounded at the equipment & proper leakage current protection device (ELCB) provided for welding machine.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q6', label: '6. Competent and Trained personnel deployed to carry the work.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q7', label: '7. Welding machine, Input / Output Cables, welding holder and weld return clamp (Holder ) insulated & in good condition', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q8', label: '8. Welder and fitter trained to connect ground / work return clamps (Holder) to the work piece prior to energization of Welding machine.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q9', label: '9. Gas Cylinders stacked vertically and not below the welding/cutting area. Regulator Key is available with cylinders.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q10', label: '10. Work Area Isolated with barricading and caution sign', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q11', label: '11. Personal Protective Equipment. Minimum applicable - Safety helmet, safety goggles, welding helmet, safety shoes, leather gloves, long sleeve and nose mask provided.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q12', label: '12. In case of pits, water removed from the pit & wood /rubber insulation provided.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q13', label: '13. Adequate & suitable nos. of fire fighting extinguisher provided.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q14', label: '14. Near by combustible material removed. Housekeeping Done.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_q15', label: '15. Fire watch as standby is in place.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'hot_other', label: '16. Other Precautions / Notes', type: 'text', required: false }
  ];

  const heightPrecautions = [
    { key: 'height_q1', label: '1. Scaffolding with valid tag available for use', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q2', label: '2. Conducted JST/TBT conducted', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q3', label: '3. Safety shoes (nonslip), Helmet with chin strip available with employees.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q4', label: '4. All tightening tools, hand tools /equipment checked and in good condition.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q5', label: '5. Access and exist marked and without obstruction.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q6', label: '6. Lighting arrangement adequate.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q7', label: '7. Unwanted and rubbish material removed from working platform.', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q8', label: '8. Electrical cable in good condition', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q9', label: '9. Signboards provided', type: 'select', options: ['Yes', 'Not Required'], required: true },
    { key: 'height_q10', label: '10. Employees aware about hazards and safe working practices while working at height.', type: 'select', options: ['Yes', 'Not Required'], required: true }
  ];

  const assemblePermit = (precautions) => {
    return permitHeaderFields
      .concat(precautions)
      .concat(permitApprovalFields)
      .concat(permitClosingFields);
  };

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
    WP_SHAFT: assemblePermit(shaftPrecautions),
    WP_NIGHT: assemblePermit(nightPrecautions),
    WP_LIFT: assemblePermit(liftPrecautions),
    WP_HOT: assemblePermit(hotPrecautions),
    WP_HEIGHT: assemblePermit(heightPrecautions),
    WP_GENERAL: assemblePermit([]),
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
