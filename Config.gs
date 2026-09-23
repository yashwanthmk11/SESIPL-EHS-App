const APP_NAME = "SESIPL EHS Digital";
const DATABASE_SPREADSHEET_ID = "1RklASnOOh_gX7sjvfJTt8eBY0JBgj1iPi0wg-G3O4N0";
const SESSION_TTL_SEC = 21600;
const ESCALATE_HOURS = 24;

const ROLES = {
  LEAD: "EHS_LEAD",
  ASST: "ASST_EHS_MANAGER",
  MANAGER: "EHS_MANAGER",
  DIRECTOR: "DIRECTOR",
};

const STATUS = {
  DRAFT: "DRAFT",
  SUBMITTED: "SUBMITTED",
  APPROVED: "APPROVED",
  REJECTED: "REJECTED",
  RESUBMITTED: "RESUBMITTED",
  OPEN: "OPEN",
  CLOSED: "CLOSED",
  FALLBACK: "FALLBACK",
};

const SHEETS = {
  USERS: "Users",
  PROJECTS: "Projects",
  PROJECT_USERS: "ProjectUsers",
  PROJECT_ORG: "ProjectOrgChart",
  FORM_CATALOG: "FormCatalog",
  SUBMISSIONS: "Submissions",
  TEMPLATE_RECORDS: "TemplateRecords",
  COMMENTS: "Comments",
  OBSERVATIONS: "Observations",
  DAILY_LOG: "DailyLog",
  PPE: "PpeIssue",
  SCAFFOLD: "ScaffoldTracker",
  MOVEMENT: "ScaffoldMovement",
  AUDITS: "Audits",
  AUDIT_SCORES: "AuditScores",
  LIBRARY: "Library",
  GALLERY: "Gallery",
  NOTIFICATIONS: "Notifications",
  TRAINING: "TrainingCalendar",
  ESCALATIONS: "Escalations",
  AUDIT_LOG: "AuditLog",
};

const HEADERS = {
  Users: [
    "employeeId",
    "uan",
    "name",
    "role",
    "email",
    "phone",
    "active",
    "mappedProjects",
  ],
  Projects: [
    "id",
    "code",
    "name",
    "client",
    "pmc",
    "inCharge",
    "manager",
    "scope",
    "startDate",
    "endDate",
    "areaSqft",
    "poNo",
    "status",
    "region",
    "projectDuration",
  ],
  ProjectUsers: ["employeeId", "projectId", "role"],
  ProjectOrgChart: [
    "id",
    "projectId",
    "level",
    "roleTitle",
    "name",
    "phone",
    "email",
  ],
  FormCatalog: [
    "formCode",
    "module",
    "title",
    "cadence",
    "entryType",
    "slaHours",
  ],
  Submissions: [
    "id",
    "requestId",
    "projectId",
    "formCode",
    "version",
    "status",
    "submittedBy",
    "submittedAt",
    "reviewedBy",
    "reviewedAt",
    "pdfFileId",
    "docFileId",
    "payloadJson",
  ],
  TemplateRecords: [
    "id",
    "submissionId",
    "projectId",
    "formCode",
    "templateKey",
    "version",
    "status",
    "submittedBy",
    "submittedAt",
    "reviewedBy",
    "reviewedAt",
    "payloadJson",
  ],
  Comments: [
    "id",
    "entityType",
    "entityId",
    "projectId",
    "byEmployeeId",
    "byName",
    "role",
    "text",
    "createdAt",
    "broadcast",
  ],
  Observations: [
    "id",
    "projectId",
    "reportNo",
    "date",
    "vendor",
    "auditedBy",
    "location",
    "contractor",
    "observation",
    "preventive",
    "status",
    "ownerEmployeeId",
    "dueAt",
    "escalateAt",
    "fallbackNote",
    "fallbackAt",
    "unsafeFileId",
    "rectifiedFileId",
    "submittedBy",
  ],
  DailyLog: [
    "id",
    "projectId",
    "date",
    "staff",
    "workers",
    "totalManpower",
    "workingHours",
    "totalManHours",
    "safeManHours",
    "cumSafeManHours",
    "inductions",
    "indStaff",
    "indWorkers",
    "tbtCount",
    "tbtPersons",
    "trainingTopic",
    "trainingPersons",
    "permitHot",
    "permitElectrical",
    "permitCold",
    "permitGeneral",
    "permitOthers",
    "firstAid",
    "nearMiss",
    "ltiCount",
    "accidentDetails",
    "remarks",
    "enteredBy",
    "verifiedBy",
    "approvedBy",
  ],
  PpeIssue: [
    "id",
    "projectId",
    "date",
    "itemCategory",
    "contractor",
    "receivedBy",
    "helmetWhite",
    "helmetGreen",
    "helmetBlue",
    "helmetRed",
    "jacketGreen",
    "jacketOrange",
    "cottonGloves",
    "leatherGloves",
    "goggle",
    "faceShield",
    "mask",
    "apron",
    "shoulderPad",
    "earMuff",
    "dcNo",
    "totalReceived",
    "issuedQty",
    "returnable",
    "balanceStock",
    "storeIncharge",
    "remarks",
    "enteredBy",
  ],
  ScaffoldTracker: [
    "id",
    "projectId",
    "date",
    "region",
    "sharavFab",
    "sesipl",
    "vinayaka",
    "hbs",
    "other",
    "totalScaffold",
    "ladderSharav",
    "ladderSesipl",
    "ladderVinayaka",
    "ladderHbs",
    "ladderRental",
    "airportLadder",
    "workstationLadder",
    "frpMsafe",
    "frpYoungman",
    "totalLadder",
    "returnedScaffold",
    "returnedLadder",
    "remarks",
    "enteredBy",
  ],
  ScaffoldMovement: [
    "id",
    "date",
    "fromProject",
    "fromQty",
    "toProject",
    "toQty",
    "remarks",
    "enteredBy",
  ],
  Audits: [
    "id",
    "projectId",
    "auditDate",
    "auditor",
    "location",
    "totalScore",
    "maxScore",
    "percent",
    "grade",
    "status",
    "submittedBy",
  ],
  AuditScores: [
    "id",
    "auditId",
    "section",
    "sn",
    "particulars",
    "score",
    "remarks",
  ],
  Library: [
    "id",
    "projectId",
    "module",
    "title",
    "fileId",
    "uploadedBy",
    "uploadedAt",
    "tags",
  ],
  Gallery: [
    "id",
    "projectId",
    "category",
    "title",
    "fileId",
    "mimeType",
    "uploadedBy",
    "uploadedAt",
  ],
  Notifications: [
    "id",
    "toEmployeeId",
    "projectId",
    "title",
    "body",
    "type",
    "read",
    "createdAt",
    "createdBy",
    "dismissed",
  ],
  TrainingCalendar: [
    "id",
    "projectId",
    "date",
    "topic",
    "owner",
    "dept",
    "status",
    "notes",
    "createdBy",
  ],
  Escalations: [
    "id",
    "entityType",
    "entityId",
    "level",
    "toEmployeeId",
    "sentAt",
    "reason",
  ],
  AuditLog: [
    "id",
    "at",
    "employeeId",
    "action",
    "entityType",
    "entityId",
    "detail",
  ],
};

/* =========================================================
   GOOGLE DOCS TEMPLATE REGISTRY
   Maps form codes to Google Doc template IDs for PDF generation.
   ========================================================= */
const DEFAULT_DOC_TEMPLATE_REGISTRY = {
  // Work Permits
  WP_HEIGHT: "1RAX7rxlZ7fY3ec3sH9OGOcZyRGD4nRKkUDkgjpgOy5w",
  WP_GENERAL: "1wekm-P4Jv3iz85gBld-dWNuSzVLjDlr05j4HzQeWamY",
  WP_HOT: "1XmYhVlBiJtK32y-IXRsX-5rxP_8ykACny4C-MuAl0HU",
  WP_LIFT: "1EZEqr-YwsyFNdUI191MV-lehWlgQmVmb9z-oMxTygs0",
  WP_SHAFT: "1dp9UV5F0OEnxq7KarUhIJUXwzZQCFH-rZTgcLWTp8mE",
  WP_NIGHT: "1MYVyt8LRIoBIDQRjX8W_pS7NdpVm88Me1wRWawUJSGE",

  // Checklists & Equipment Inspections
  CL_WELD: "1IYwFad3phxQ3VeZTbsu6INBvDfJWhSIEdBmKRcQjS6Q", // Welding Machine
  CL_GRIND: "1c7T1ip1jnnz9Tc0b61fASMYOBKC2z2zonVN1SUC3aVA", // Grinding Machine Checklist
  CL_DRILL: "1x18zs1fBXgPJqlejQqUgMIy6juMhQb_rQSGMZOdt9ug", // Drilling Machine
  CL_FE: "1H87LMWzmqgIWHPpm54O4vQUrq6rvh4yAi6u5ZPDelns", // Fire Extinguisher
  CL_SCAFFOLD: "10J-l0KyOAHOAr7ulkv3TPgHOJqF3vaywUvD9mncVF1s", // Scaffolding checklist
  CL_CUT: "",

  // Attendances & Inductions
  CL_TBT: "1vt3LBwuRnbo_Av2zaLVPsDK00ZT1-hiNN09fntRAYXE", // Tool Box Talk
  CL_JST: "1eCSSEMa7-QRyklGQudxOqkaXMwfI78Gl9HH_vKC4xi4", // JST Attendance sheet
  CL_INDUCTION: "",

  // Worker Screening & Medical
  CL_SCREENING: "1py6zgX0lv9NI_Cbe9aDMPctYEey0_3wS0RW_yE0S9Z0", // Screening of Worker Format
  CL_MEDICAL: "1rceiB1zO3IKEO7IHWfT1jNjSC3vZ5eFGjUEWRYC4Znc", // Medical certificate-xI

  // Stickers / Tags
  TAG_IND: "",
  TAG_TOOL: "",
  TAG_FE: "",
  TAG_RED: "",
};

function getDocTemplateRegistry_() {
  const merged = Object.assign({}, DEFAULT_DOC_TEMPLATE_REGISTRY);
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(
      "DOC_TEMPLATE_REGISTRY",
    );
    if (raw) {
      const parsed = JSON.parse(raw);
      for (const k in parsed) {
        if (parsed[k] && String(parsed[k]).trim() !== "") {
          merged[k] = String(parsed[k]).trim();
        }
      }
    }
  } catch (e) {
    // fallback
  }
  return merged;
}

function setDocTemplateRegistry_(registry) {
  const merged = Object.assign({}, getDocTemplateRegistry_(), registry || {});
  PropertiesService.getScriptProperties().setProperty(
    "DOC_TEMPLATE_REGISTRY",
    JSON.stringify(merged),
  );
  return merged;
}

const MODULES = [
  {
    id: "TEST_CERT",
    title: "Test Certificates",
    cadence: "ONETIME",
    icon: "verified",
  },
  {
    id: "WEEKLY",
    title: "Weekly / Monthly Report",
    cadence: "WEEKLY",
    icon: "calendar",
  },
  {
    id: "PERMIT",
    title: "Work Permit",
    cadence: "AS_REQUIRED",
    icon: "permit",
  },
  {
    id: "EHS_DOCS",
    title: "EHS Documents",
    cadence: "ONETIME",
    icon: "folder",
  },
  { id: "LEGAL", title: "Legal Documents", cadence: "ONETIME", icon: "legal" },
  {
    id: "OBSERVATION",
    title: "Observation",
    cadence: "AS_REQUIRED",
    icon: "eye",
  },
  {
    id: "POLICY",
    title: "Policy & Certificate",
    cadence: "ONETIME",
    icon: "policy",
  },
  { id: "SWMS", title: "SWMS HIRA", cadence: "ONETIME", icon: "hazard" },
  {
    id: "CHECKLIST",
    title: "Checklist & Formats",
    cadence: "DAILY",
    icon: "check",
  },
];

const FORM_DEFS = [
  {
    formCode: "TC_UPLOAD",
    module: "TEST_CERT",
    title: "Test Certificate Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "WR_WEEKLY",
    module: "WEEKLY",
    title: "Weekly EHS Report",
    cadence: "WEEKLY",
    entryType: "UPLOAD",
    slaHours: 24,
  },
  {
    formCode: "WR_MONTHLY",
    module: "WEEKLY",
    title: "Monthly EHS Report",
    cadence: "MONTHLY",
    entryType: "UPLOAD",
    slaHours: 48,
  },
  {
    formCode: "WP_HEIGHT",
    module: "PERMIT",
    title: "Working At Height Permit",
    cadence: "AS_REQUIRED",
    entryType: "FORM",
    slaHours: 8,
  },
  {
    formCode: "WP_NIGHT",
    module: "PERMIT",
    title: "Night Work Permit",
    cadence: "AS_REQUIRED",
    entryType: "FORM",
    slaHours: 8,
  },
  {
    formCode: "WP_SHAFT",
    module: "PERMIT",
    title: "Shaft Work Permit",
    cadence: "AS_REQUIRED",
    entryType: "FORM",
    slaHours: 8,
  },
  {
    formCode: "WP_LIFT",
    module: "PERMIT",
    title: "Lifting Activity Permit",
    cadence: "AS_REQUIRED",
    entryType: "FORM",
    slaHours: 8,
  },
  {
    formCode: "WP_HOT",
    module: "PERMIT",
    title: "Hot Work Permit",
    cadence: "AS_REQUIRED",
    entryType: "FORM",
    slaHours: 8,
  },
  {
    formCode: "WP_GENERAL",
    module: "PERMIT",
    title: "General Work Permit",
    cadence: "AS_REQUIRED",
    entryType: "FORM",
    slaHours: 8,
  },
  {
    formCode: "WP_UPLOAD",
    module: "PERMIT",
    title: "Work Permit File Upload (Signed PTW Scan / Copy)",
    cadence: "AS_REQUIRED",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "EHS_UPLOAD",
    module: "EHS_DOCS",
    title: "EHS Document Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "LEGAL_UPLOAD",
    module: "LEGAL",
    title: "Legal Document Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "OBS_DAILY",
    module: "OBSERVATION",
    title: "Daily Safety Observation",
    cadence: "AS_REQUIRED",
    entryType: "FORM",
    slaHours: 24,
  },
  {
    formCode: "POL_UPLOAD",
    module: "POLICY",
    title: "Policy & Certificate Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "SWMS_UPLOAD",
    module: "SWMS",
    title: "SWMS / HIRA Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "CL_DRILL",
    module: "CHECKLIST",
    title: "Drilling Machine Checklist",
    cadence: "DAILY",
    entryType: "FORM",
    slaHours: 24,
  },
  {
    formCode: "CL_WELD",
    module: "CHECKLIST",
    title: "Welding Machine Checklist",
    cadence: "DAILY",
    entryType: "FORM",
    slaHours: 24,
  },
  {
    formCode: "CL_TBT",
    module: "CHECKLIST",
    title: "Tool Box Talk",
    cadence: "DAILY",
    entryType: "FORM",
    slaHours: 24,
  },
  {
    formCode: "CL_JST",
    module: "CHECKLIST",
    title: "JST Attendance Sheet",
    cadence: "DAILY",
    entryType: "FORM",
    slaHours: 24,
  },
  {
    formCode: "CL_GRIND",
    module: "CHECKLIST",
    title: "Grinding Machine Checklist",
    cadence: "DAILY",
    entryType: "FORM",
    slaHours: 24,
  },
  {
    formCode: "CL_SCAFFOLD",
    module: "CHECKLIST",
    title: "Scaffolding Checklist",
    cadence: "DAILY",
    entryType: "FORM",
    slaHours: 24,
  },
  {
    formCode: "CL_FE",
    module: "CHECKLIST",
    title: "Fire Extinguisher Checklist",
    cadence: "DAILY",
    entryType: "FORM",
    slaHours: 24,
  },

  // One-time Document Uploads (Remaining Checklists, Screening, Medical & Tags)
  {
    formCode: "CL_SCREENING",
    module: "EHS_DOCS",
    title: "Screening of Worker Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "CL_MEDICAL",
    module: "EHS_DOCS",
    title: "Medical Certificate (Form XI) Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "CL_IDCARD",
    module: "EHS_DOCS",
    title: "Project ID Card Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "CL_CUT",
    module: "EHS_DOCS",
    title: "Cutting Machine Checklist Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "CL_INDUCTION",
    module: "EHS_DOCS",
    title: "EHS Induction Record Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "TAG_IND",
    module: "EHS_DOCS",
    title: "Induction Sticker Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "TAG_TOOL",
    module: "EHS_DOCS",
    title: "Tools Inspection Sticker Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "TAG_FE",
    module: "EHS_DOCS",
    title: "Fire Extinguisher Tag Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "TAG_RED",
    module: "EHS_DOCS",
    title: "Red Tag Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
  {
    formCode: "TAG_SCAFF",
    module: "EHS_DOCS",
    title: "Scaffold Inspection Tag Upload",
    cadence: "ONETIME",
    entryType: "UPLOAD",
    slaHours: 0,
  },
];

const AUDIT_SECTIONS = [
  { id: "A", name: "GENERAL REQUIREMENTS & POLICY", max: 30 },
  { id: "B", name: "HAZARDS, RISK MANAGEMENT & SWMS", max: 45 },
  { id: "C", name: "EHS SCREENING, INDUCTION and TRAINING", max: 35 },
  { id: "D", name: "Documents & Records", max: 35 },
  { id: "E", name: "EMERGENCY PREPAREDNESS", max: 25 },
  { id: "F", name: "EHS PERFORMANCE MEASUREMENT & MOTIVATION", max: 10 },
  { id: "G", name: "REPORTING AND INVESTIGATION", max: 30 },
  { id: "H", name: "Welfare & Health Management", max: 55 },
  { id: "I", name: "Mechanical Management System", max: 25 },
  { id: "J", name: "Electrical Management System", max: 40 },
  { id: "K", name: "Housekeeping Management", max: 15 },
  { id: "L", name: "Waste Management", max: 10 },
  { id: "M", name: "Store & Material Management", max: 20 },
  { id: "N", name: "Working Platforms and Ladders", max: 20 },
  { id: "O", name: "FLOOR DB PANELS", max: 60 },
  { id: "P", name: "HOT WORK", max: 40 },
  { id: "Q", name: "SHAFT WORK", max: 0 },
  { id: "R", name: "PPES", max: 30 },
];

const AUDIT_PERFORMANCE_BANDS = [
  {
    name: "Platinum",
    min: 85,
    max: 100,
    color: "#f8fafc",
    textColor: "#0f172a",
    desc: "85 - 100 %",
  },
  {
    name: "Gold",
    min: 71,
    max: 84,
    color: "#fef08a",
    textColor: "#854d0e",
    desc: "71 - 84 %",
  },
  {
    name: "Silver",
    min: 55,
    max: 70,
    color: "#bbf7d0",
    textColor: "#166534",
    desc: "55 - 70 %",
  },
  {
    name: "Blue",
    min: 0,
    max: 54,
    color: "#bfdbfe",
    textColor: "#1e40af",
    desc: "< 54 %",
  },
];

const AUDIT_CHECKLIST_SCHEMA = [
  {
    id: "A",
    name: "GENERAL REQUIREMENTS & POLICY",
    max: 30,
    page: 1,
    items: [
      {
        sn: 1,
        text: "A Project Specific EHS plan is available ?",
        defaultScore: 5,
      },
      {
        sn: 2,
        text: "Has EHS plan is approved by PMC? and updated?",
        defaultScore: 5,
      },
      {
        sn: 3,
        text: "Is the policy and EHS Plan availabe and explained to all employees, including temporary staff? Records available?",
        defaultScore: 5,
      },
      { sn: 4, text: "Is statutory documents are available?", defaultScore: 5 },
      {
        sn: 5,
        text: "Individual EHS responsibilities defined, documented, known and understood?",
        defaultScore: 5,
      },
      {
        sn: 6,
        text: "Safety Committee formed & Meeting MOM available",
        defaultScore: 5,
      },
    ],
  },
  {
    id: "B",
    name: "HAZARDS, RISK MANAGEMENT & SWMS",
    max: 45,
    page: 1,
    items: [
      { sn: 1, text: "Are all foreseeable risk assessed?", defaultScore: 3 },
      {
        sn: 2,
        text: "Is updated risk Assessment register available?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Are such discovered risks fully evaluated & documented?",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "Are risk assessments and evaluations communicated to all concerned?",
        defaultScore: 3,
      },
      {
        sn: 5,
        text: "Is projects environmental aspect evaluation documented?",
        defaultScore: "NA",
      },
      {
        sn: 6,
        text: "Do those engaged in high risk activities receive guidance and training on risk mitigation measures?",
        defaultScore: 3,
      },
      {
        sn: 7,
        text: "Have method statements available for all jobs/tasks?",
        defaultScore: 3,
      },
      {
        sn: 8,
        text: "Do they adequately reflect legal standards and requirements?",
        defaultScore: 3,
      },
      {
        sn: 9,
        text: "Are Safe Work Practices (SWMS) reviewed?",
        defaultScore: 3,
      },
      {
        sn: 10,
        text: "Are planned job observations conducted to ensure compliance with safe work practices? Are these observations documented?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "C",
    name: "EHS SCREENING, INDUCTION and TRAINING",
    max: 35,
    page: 1,
    items: [
      {
        sn: 1,
        text: "Has a Induction plan and module available?",
        defaultScore: 5,
      },
      {
        sn: 2,
        text: "Has a Identified Induction area available?",
        defaultScore: 5,
      },
      { sn: 3, text: "Are the induction details recorded?", defaultScore: 5 },
      {
        sn: 4,
        text: "Has a training plan been formulated and documented?",
        defaultScore: 5,
      },
      {
        sn: 5,
        text: "Is the progress of actual training monitored against the training plan/matrix?",
        defaultScore: 3,
      },
      {
        sn: 6,
        text: "Are records of actual training achieved, maintained and analyzed? Do training matrix provided for refresher training",
        defaultScore: 5,
      },
      { sn: 7, text: "Do competency of people ensured?", defaultScore: 5 },
    ],
  },
  {
    id: "D",
    name: "Documents & Records",
    max: 35,
    page: 1,
    items: [
      {
        sn: 1,
        text: "Has a program of planned inspections been established?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Do itemized inspection checklists exist and used?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Are EHS inspections, Records of observation & their Compliance status maintained?",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "Are Checklists implemented and Recorded on daily basis?",
        defaultScore: 3,
      },
      {
        sn: 5,
        text: "Are Implementation of Permit to Work & their effectiveness maintained?",
        defaultScore: 3,
      },
      {
        sn: 6,
        text: "Are EHS Daily, Weekly & Monthly reports maintained?",
        defaultScore: 3,
      },
      {
        sn: 7,
        text: "Is pre health check-ups done for all new labours and certificates are available? And remedial checkups done at least once in a year and records are available?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "E",
    name: "EMERGENCY PREPADENESS",
    max: 25,
    page: 1,
    items: [
      {
        sn: 1,
        text: "Is emergency preparedness and response plan available?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Emergency Rescue team (ERT) is developed and followed?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Emergency Rescue team is identified and effectively communicated?",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "All potential emergency situations identified, a plan for mock drills developed and followed?",
        defaultScore: 3,
      },
      {
        sn: 5,
        text: "Is Emergency contact Details displayed?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "F",
    name: "EHS PERFORMANCE MEASUREMENT & MOTIVATION",
    max: 10,
    page: 1,
    items: [
      {
        sn: 1,
        text: "Performance monitoring and Evaluation? Result published and Records available?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Is continuing publicity given to EHS objectives and any motivation programs being conducted?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "G",
    name: "REPORTING AND INVESTIGATION",
    max: 30,
    page: 2,
    items: [
      {
        sn: 1,
        text: "Internal audits conducted, recommendations raised, follow up & Close out?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Has an appropriate reporting procedure been established for all Incidents? SESIPL procedure is being ensured?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Is the reporting procedure correctly used?",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "Are all incidents including unsafe conditions/unsafe acts/near miss occurrences properly investigated to find out route cause?",
        defaultScore: 3,
      },
      {
        sn: 5,
        text: "Is adequate preventative follow up action taken as a result of reports received/ raised?",
        defaultScore: 3,
      },
      {
        sn: 6,
        text: "Is personal injury/property damage/near miss/ Environment Incident data suitably classified? Recorded & communicated?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "H",
    name: "Welfare & Health Management",
    max: 55,
    page: 2,
    items: [
      {
        sn: 1,
        text: "Is the first AID provision available and details desplayed?",
        defaultScore: 3,
      },
      { sn: 2, text: "Is Qualified First Aider available", defaultScore: 3 },
      {
        sn: 3,
        text: "Are List of First Aid Equipment's & Stock status available at site?",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "Medical Check-up of workers and staff (Pre & Post)",
        defaultScore: 3,
      },
      {
        sn: 5,
        text: "Is Tie up with near by HOSPITAL is made?",
        defaultScore: 3,
      },
      {
        sn: 6,
        text: "Are the records of first aid treatment / First aid Register available?",
        defaultScore: 3,
      },
      {
        sn: 7,
        text: "Record of accident / incident investigation and their preventive action are available?",
        defaultScore: 3,
      },
      {
        sn: 8,
        text: "Records of first aid training is available?",
        defaultScore: 3,
      },
      {
        sn: 9,
        text: "Is canteen or Lunch area available at site for workers?",
        defaultScore: 3,
      },
      {
        sn: 10,
        text: '"Welfare facilities (Power supply, Drinking Water, Toilets in good conditions',
        defaultScore: 3,
      },
      {
        sn: 11,
        text: "Drinking water provided & Test certificate available ?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "I",
    name: "Mechanical Management System",
    max: 25,
    page: 2,
    items: [
      {
        sn: 1,
        text: "Is the List of Plant & Machineries available ?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Are Competency Certificate (Experience and qualification certificate) available at site?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Third Party Inspection Report available",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "Is the condition of machines and their operation procedures are as per safety system? And in good condition",
        defaultScore: 3,
      },
      {
        sn: 5,
        text: "Are all moving / rotator parts of machinery adequately guarded?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "J",
    name: "Electrical Management System",
    max: 40,
    page: 2,
    items: [
      {
        sn: 1,
        text: "Is only persons having valid company authorization certificate are employed for carrying out electrical work and repair of electrical equipment, installation and maintenance at site?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Is electrical distribution boards are mounted above ground level?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Is cables are kept away from the passage to prevent slip, trip or fall hazard and kept at height wherever possible?",
        defaultScore: 3,
      },
      { sn: 4, text: "Is LOTO Available at site?", defaultScore: 3 },
      {
        sn: 5,
        text: "Is the condition of electric cable and connectivity to the board is in acceptable condition?",
        defaultScore: 3,
      },
      {
        sn: 6,
        text: '"Is DGs have the folowing Specifications? a)installed is provided with platform approved by PMC. b)Acoustic type DG is provided / DG is protected from rain, etc. c)DG maintained in good condition (Eco Friendly & CEIG Approval)',
        defaultScore: "NA",
      },
      {
        sn: 7,
        text: "Competency certificate of electrician is available at site?",
        defaultScore: 3,
      },
      {
        sn: 8,
        text: "Nomenclature of electrical earth pit and record of resistance is available?",
        defaultScore: 3,
      },
      {
        sn: 9,
        text: "Is Periodic Electrical Inspection carried out as per schedule?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "K",
    name: "Housekeeping Management",
    max: 15,
    page: 2,
    items: [
      {
        sn: 1,
        text: "Are all the passages, floors, and the stairways free from obstruction?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Is material stacked properly at all site locations?",
        defaultScore: 3,
      },
      { sn: 3, text: "Is scrap area properly identified?", defaultScore: 3 },
    ],
  },
  {
    id: "L",
    name: "Waste Management",
    max: 10,
    page: 2,
    items: [
      {
        sn: 1,
        text: "Are wastes bin colour coded as per waste disposal?",
        defaultScore: 3,
      },
      { sn: 2, text: "Sagrigation of waste and scrape", defaultScore: 3 },
    ],
  },
  {
    id: "M",
    name: "Store & Material Management",
    max: 20,
    page: 2,
    items: [
      {
        sn: 1,
        text: "MSDS of Various chemical available at site?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Status of stacking of materials is in good condition?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Have inflammable materials such as paints, Oil, diesel, petrol, chemicals etc., been stored in separate areas out side the main stores?",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "Are firefighting equipment is easily accessible, i.e. without any obstruction within 15 ft. of “Fire Points”?",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "N",
    name: "Working Platforms and Ladders",
    max: 20,
    page: 4,
    items: [
      {
        sn: 1,
        text: "Is scaffolding material is free from bends, cuts, rust, coated by antirust paint and Working platform deck & its access is free from all debris and loose materials?",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Is working platform has proper and safe access. Scaffolds have the proper guard rails at 1.0 m height with middle rails at 0.5 m height from the platform?",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Is tagging system for scaffolds (stating fitness of the scaffolding) is being maintained",
        defaultScore: 3,
      },
      { sn: 4, text: "Is propper agrement is available", defaultScore: 3 },
    ],
  },
  {
    id: "O",
    name: "FLOOR DB PANELS:",
    max: 60,
    page: 4,
    items: [
      {
        sn: 1,
        text: "Is sufficient illumination provided in all work areas",
        defaultScore: 5,
      },
      {
        sn: 2,
        text: "Are Floor DB panels are placed in an accessible location",
        defaultScore: 5,
      },
      {
        sn: 3,
        text: "Are Floor DB panels are equipped with RCCB/ELCB",
        defaultScore: 5,
      },
      { sn: 4, text: "Are the RCCB/ELCB are functioning", defaultScore: 5 },
      {
        sn: 5,
        text: "Is any ELCB/RCCB Testing has been conducted at site",
        defaultScore: 5,
      },
      {
        sn: 6,
        text: "Is Earth connection available with the DB Panel",
        defaultScore: 5,
      },
      { sn: 7, text: "Is Rubber mat available", defaultScore: 5 },
      {
        sn: 8,
        text: "Is CO2/DCP Fire extinguisher has been provided for each DB Panel",
        defaultScore: 5,
      },
      {
        sn: 9,
        text: "Is physical condition of the Panel is good",
        defaultScore: 5,
      },
      {
        sn: 10,
        text: "Are there any emergency contact details displayed on the Panels",
        defaultScore: 5,
      },
      {
        sn: 11,
        text: "Is the temporary cable routing is in a proper manner",
        defaultScore: 5,
      },
      {
        sn: 12,
        text: "Is any weekly/Monthly inspection is being carried & checklist is made available",
        defaultScore: 5,
      },
    ],
  },
  {
    id: "P",
    name: "HOT WORK:",
    max: 40,
    page: 4,
    items: [
      {
        sn: 1,
        text: "Is any specific fabrication/welding yard identified",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Is the area of hot work is free from combustibles",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Is required PPEs is being worn by the welder",
        defaultScore: 3,
      },
      { sn: 4, text: "Is the welding machine is fit for use", defaultScore: 3 },
      {
        sn: 5,
        text: "Is used welding electrodes are being collected separately",
        defaultScore: 3,
      },
      {
        sn: 6,
        text: "Is there is availability of fire extinguisher near the hot work area",
        defaultScore: 3,
      },
      {
        sn: 7,
        text: "Is there any signage available near hot work area",
        defaultScore: 3,
      },
      {
        sn: 8,
        text: "Are there any training provided for workmen involving in hot work",
        defaultScore: 3,
      },
    ],
  },
  {
    id: "Q",
    name: "SHAFT WORK:",
    max: 0,
    page: 4,
    items: [
      {
        sn: 1,
        text: "Is there any proper fall arresting system available",
        defaultScore: "NA",
      },
      {
        sn: 2,
        text: "Is there any safe work platform available",
        defaultScore: "NA",
      },
      {
        sn: 3,
        text: "Is the method of shaft work could be considered as safe",
        defaultScore: "NA",
      },
      {
        sn: 4,
        text: "Is the workman entering a shaft has undergone medical checkup prior to employment & is not having vertigo",
        defaultScore: "NA",
      },
      {
        sn: 5,
        text: "Are special safety precautions are being taken for any hot work being carried inside of the shaft",
        defaultScore: "NA",
      },
      {
        sn: 6,
        text: "Is there any continuous supervision being provided",
        defaultScore: "NA",
      },
      {
        sn: 7,
        text: "Are there any training provided for workmen involving in work inside shaft",
        defaultScore: "NA",
      },
      {
        sn: 8,
        text: "Are all workmen working at height (more than 1.8/2 Mtrs) are working without any fall arresting system",
        defaultScore: "NA",
      },
      {
        sn: 9,
        text: "Are there any workmen working in an unsafe working platform to support work at height",
        defaultScore: "NA",
      },
      {
        sn: 10,
        text: "Are the lanyard hooks of the full body harness has been anchored to an secured anchoring point",
        defaultScore: "NA",
      },
      {
        sn: 11,
        text: "Are there any training provided for workmen involving in height work",
        defaultScore: "NA",
      },
      {
        sn: 12,
        text: "Is there any proper climbing access & egress on the working platform being used for work at height",
        defaultScore: "NA",
      },
      {
        sn: 13,
        text: "Is any workmen with vertigo issue has been deployed in height work",
        defaultScore: "NA",
      },
    ],
  },
  {
    id: "R",
    name: "PPES",
    max: 30,
    page: 4,
    items: [
      {
        sn: 1,
        text: "Are all workmen are using Job specific PPEs",
        defaultScore: 3,
      },
      {
        sn: 2,
        text: "Are there any workmen found with faulty PPE",
        defaultScore: 3,
      },
      {
        sn: 3,
        text: "Are all full body harness available at site are fit for use",
        defaultScore: 3,
      },
      {
        sn: 4,
        text: "Are the damaged PPEs are being collected back to avoid use of damaged PPEs at site",
        defaultScore: 3,
      },
      {
        sn: 5,
        text: "Is there is any register maintained for the issue of PPEs",
        defaultScore: 3,
      },
      {
        sn: 6,
        text: "Are all critical work activities such as welding, grinding, cutting, chipping, drilling etc. are being carried with job specific PPEs",
        defaultScore: 3,
      },
    ],
  },
];

function getDefaultAuditSeedData_(project) {
  const p = project || {
    id: "PRJ001",
    code: "BLR-01",
    name: "Intuit Bellandur",
    areaSqft: "2,45,000 sq.ft",
  };
  const sectionScores = {};
  let totalScore = 0;
  let totalMax = 0;

  AUDIT_CHECKLIST_SCHEMA.forEach((sec) => {
    let secActual = 0;
    sec.items.forEach((item) => {
      const val = item.defaultScore;
      if (typeof val === "number") {
        secActual += val;
      }
    });
    const pct = sec.max > 0 ? Math.round((secActual / sec.max) * 100) : null;
    sectionScores[sec.id] = {
      id: sec.id,
      name: sec.name,
      max: sec.max,
      actual: secActual,
      percent: pct,
      percentText: pct !== null ? pct + "%" : "#DIV/0!",
    };
    totalScore += secActual;
    totalMax += sec.max;
  });

  const overallPercent =
    totalMax > 0 ? Math.round((totalScore / totalMax) * 100) : 0;
  let grade = "Silver";
  if (overallPercent >= 85) grade = "Platinum";
  else if (overallPercent >= 71) grade = "Gold";
  else if (overallPercent >= 55) grade = "Silver";
  else grade = "Blue";

  return {
    id: "ADT-" + (p.code || "BLR-01") + "-20260910",
    projectId: p.id,
    projectName: p.name || "Intuit Bellandur",
    projectLocation: p.areaSqft
      ? p.name + ", " + p.areaSqft
      : "Pritech Park, Bellandur, Bengaluru",
    auditDate: "2026-09-10",
    auditor: "CBRE Lead Auditor",
    auditorEmail: "info@shankarelectricals.com",
    auditorWebsite: "www.shankarelectricals.com",
    totalScore: totalScore,
    maxScore: totalMax,
    percent: overallPercent,
    grade: grade,
    status: "APPROVED",
    sectionScores: sectionScores,
    schema: AUDIT_CHECKLIST_SCHEMA,
  };
}

const GALLERY_CATEGORIES = ["COMPANY", "TRAINING", "MEETING", "EVENT"];

const PERMIT_DOC_CODES = {
  WP_SHAFT: "SESIPL-EHS.Blr Prj-06",
  WP_NIGHT: "SESIPL-EHS.Blr Prj-07",
  WP_LIFT: "SESIPL-EHS.Blr Prj-05",
  WP_HOT: "SESIPL-EHS.Blr Prj-03",
  WP_HEIGHT: "SESIPL-EHS.Blr Prj-04",
  WP_GENERAL: "SESIPL-EHS.Blr Prj-01",
};

function getFormFields_(formCode) {
  const common = [
    { key: "date", label: "Date", type: "date", required: true },
    { key: "location", label: "Location / Area", type: "text", required: true },
    {
      key: "supervisor",
      label: "Supervisor / In-charge",
      type: "text",
      required: true,
    },
  ];

  // Official SESIPL Work Permit Header & Authority Fields (1:1 with scanned PDFs)
  const permitHeaderFields = [
    {
      key: "contractorName",
      label: "Contractor Name",
      type: "text",
      required: true,
      placeholder: "e.g. Vinayaka Electricals / Apex Engineering Pvt. Ltd.",
    },
    {
      key: "emergencyContact1",
      label: "Emergency Contact No. 1",
      type: "text",
      required: true,
      placeholder: "e.g. +91 98450 12345 (Site Safety In-charge)",
    },
    {
      key: "emergencyContact2",
      label: "Emergency Contact No. 2",
      type: "text",
      required: false,
      placeholder: "e.g. +91 98450 67890 (Site Medical / Ambulance)",
    },
    {
      key: "permitNo",
      label: "Permit No",
      type: "text",
      required: true,
      placeholder: "e.g. SESIPL/WP/2026/042",
    },
    {
      key: "area",
      label: "Area",
      type: "text",
      required: true,
      placeholder: "e.g. Building 2, Floor 4, AHU Shaft",
    },
    {
      key: "location",
      label: "Location",
      type: "text",
      required: true,
      placeholder: "e.g. North Wing Grid C12-D14",
    },
    {
      key: "date",
      label: "Date",
      type: "date",
      required: true,
      placeholder: "YYYY-MM-DD",
    },
    {
      key: "time",
      label: "Time",
      type: "time",
      required: true,
      placeholder: "HH:mm",
    },
    {
      key: "siteEngineer",
      label: "Name of Site Engineer (Permit Requesting Authority)",
      type: "text",
      required: true,
      placeholder: "e.g. R. Prakash (Site In-Charge)",
    },
    {
      key: "siteEngineerSign",
      label: "Site Engineer Sign / Ack",
      type: "text",
      required: false,
      placeholder: "e.g. Signed - R. Prakash",
    },
    {
      key: "safetyOfficer",
      label: "Name of Safety Officer",
      type: "text",
      required: true,
      placeholder: "e.g. K. Harish (EHS Lead)",
    },
    {
      key: "safetyOfficerSign",
      label: "Safety Officer Sign / Ack",
      type: "text",
      required: false,
      placeholder: "e.g. Verified & Signed - K. Harish",
    },
    {
      key: "contractorInCharge",
      label: "Name of Contractor Site In charge",
      type: "text",
      required: true,
      placeholder: "e.g. S. Kumar (Agency Lead)",
    },
    {
      key: "contactNumber",
      label: "Contact Number",
      type: "text",
      required: true,
      placeholder: "e.g. +91 94480 55667",
    },
    {
      key: "workDescription",
      label: "Description of work",
      type: "textarea",
      required: true,
      placeholder:
        "e.g. Cable tray bracket erection and electrical conduits installation at 4.5m height using mobile scaffold and double lanyard harness...",
    },
    {
      key: "workExecutionDate",
      label: "Work Execution Date",
      type: "date",
      required: true,
      placeholder: "YYYY-MM-DD",
    },
    {
      key: "validFrom",
      label: "Valid From",
      type: "datetime-local",
      required: true,
      placeholder: "YYYY-MM-DDTHH:mm",
    },
    {
      key: "validTo",
      label: "Valid To",
      type: "datetime-local",
      required: true,
      placeholder: "YYYY-MM-DDTHH:mm",
    },
  ];

  const permitApprovalFields = [
    {
      key: "approvalEhsName",
      label: "Reviewed & Approved By EHS (Name)",
      type: "text",
      required: false,
      placeholder: "e.g. HR Ravikiran (EHS Manager)",
    },
    {
      key: "approvalEhsSign",
      label: "EHS Sign",
      type: "text",
      required: false,
      placeholder: "e.g. Approved - HR Ravikiran",
    },
    {
      key: "approvalEhsDate",
      label: "EHS Approval Date",
      type: "date",
      required: false,
      placeholder: "YYYY-MM-DD",
    },
    {
      key: "approvalEhsTime",
      label: "EHS Approval Time",
      type: "time",
      required: false,
      placeholder: "HH:mm",
    },
    {
      key: "approvalSiteEngineerName",
      label: "Reviewed & Approved By Site Engineer (Name)",
      type: "text",
      required: false,
      placeholder: "e.g. S. Sharma (Project Manager)",
    },
    {
      key: "approvalSiteEngineerSign",
      label: "Site Engineer Sign",
      type: "text",
      required: false,
      placeholder: "e.g. Approved - S. Sharma",
    },
    {
      key: "approvalSiteEngineerDate",
      label: "Site Engineer Approval Date",
      type: "date",
      required: false,
      placeholder: "YYYY-MM-DD",
    },
    {
      key: "approvalSiteEngineerTime",
      label: "Site Engineer Approval Time",
      type: "time",
      required: false,
      placeholder: "HH:mm",
    },
  ];

  const permitClosingFields = [
    {
      key: "closingSiteEngName",
      label: "Closing Site Engineer (Requesting Authority)",
      type: "text",
      required: false,
      placeholder: "e.g. R. Prakash",
    },
    {
      key: "closingSiteEngSign",
      label: "Site Engineer Closing Sign",
      type: "text",
      required: false,
      placeholder: "e.g. Work Completed - R. Prakash",
    },
    {
      key: "closingSiteEngDate",
      label: "Closing Date",
      type: "date",
      required: false,
      placeholder: "YYYY-MM-DD",
    },
    {
      key: "closingSiteEngTime",
      label: "Closing Time",
      type: "time",
      required: false,
      placeholder: "HH:mm",
    },
    {
      key: "closingSafetyOfficerName",
      label: "Closing Safety Officer",
      type: "text",
      required: false,
      placeholder: "e.g. K. Harish",
    },
    {
      key: "closingSafetyOfficerSign",
      label: "Safety Officer Closing Sign",
      type: "text",
      required: false,
      placeholder: "e.g. Housekeeping Inspected - K. Harish",
    },
    {
      key: "closingSafetyOfficerDate",
      label: "Safety Officer Closing Date",
      type: "date",
      required: false,
      placeholder: "YYYY-MM-DD",
    },
    {
      key: "closingSafetyOfficerTime",
      label: "Safety Officer Closing Time",
      type: "time",
      required: false,
      placeholder: "HH:mm",
    },
    {
      key: "closingPmcSiteEngName",
      label: "Closing PMC Site Engineer (Issuing Authority)",
      type: "text",
      required: false,
      placeholder: "e.g. CBRE Site Incharge",
    },
    {
      key: "closingPmcSiteEngSign",
      label: "PMC Engineer Closing Sign",
      type: "text",
      required: false,
      placeholder: "e.g. Verified & Closed - PMC",
    },
    {
      key: "closingPmcSiteEngDate",
      label: "PMC Engineer Closing Date",
      type: "date",
      required: false,
      placeholder: "YYYY-MM-DD",
    },
    {
      key: "closingPmcSiteEngTime",
      label: "PMC Engineer Closing Time",
      type: "time",
      required: false,
      placeholder: "HH:mm",
    },
  ];

  // Specific Precautions Checklists from Scanned SESIPL Forms
  const shaftPrecautions = [
    {
      key: "shaft_q1",
      label: "1. Proper Access/ Exit available",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "shaft_q2",
      label: "2. Proper ventilation and / or lighting provided",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "shaft_q3",
      label: "3. Proper & Safe platform provided",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "shaft_q4",
      label: "4. Workers have been briefed about hazardous",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "shaft_q5",
      label: "5. All Electrical Tools and machinery checked prior to use.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "shaft_q6",
      label: "6. Shaft area Properly barricaded.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "shaft_q7",
      label:
        "7. Conducted JST for all workers who are all engaging to shaft work.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "workmenNames",
      label: "Names of workmen entering shaft",
      type: "textarea",
      required: true,
    },
  ];

  const nightPrecautions = [
    {
      key: "night_q1",
      label: "1. Is dedicated Night shift in charge available?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q2",
      label: "2. Is supervisor available in night shift to supervise the task?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q3",
      label: "3. Is First aider available?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q4",
      label: "4. Is Ambulance available for emergency?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q5",
      label: "5. Are the workers working continuously for last 12 hours?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q6",
      label: "6. Is the work area safe for work?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q7",
      label: "7. Is there proper illumination provided at the work area?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q8",
      label:
        "8. Are the hazards related with the work identified and assessed at workplace?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q9",
      label:
        "9. Is toolbox talk / pre-start briefing carried out prior to start night shift?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q10",
      label:
        "10. Are the workers having specific PPE’s according to the requirement of the task?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_q11",
      label:
        "11. Is there any high-risk activity like working at height, work in penetration and shafts, electrical testing and commissioning, hot work, Mechanical lifting operation, excavation etc. to be carried out in night shift?",
      type: "select",
      options: ["Yes", "NO", "NA"],
      required: true,
    },
    {
      key: "night_remarks",
      label: "Night Shift Precautions Remarks",
      type: "textarea",
      required: false,
    },
  ];

  const liftPrecautions = [
    {
      key: "lift_q1",
      label:
        "1. Crane used for lifting activity tested, certified and approved for rated lifting works.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q2",
      label:
        "2. All lifting tackles, gears/ appliances are tested and certified for lifting works.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q3",
      label:
        "3. Crane operator is trained and competent for lifting operation.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q4",
      label:
        "4. Lifting belt protected against sharp edge of jobs to be lifted.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q5",
      label: "5. Access and exist marked and without obstruction.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q6",
      label: "6. Lighting arrangement adequate.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q7",
      label: "7. Unwanted and rubbish material removed from working platform.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q8",
      label:
        "8. Guidelines has provided for balancing & guiding jobs to be lifted.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q9",
      label:
        "9. Periphery area of crane booms as well lifting job is barricaded .",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q10",
      label:
        "10. Rigger and signal man is trained and competent for lifting work.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q11",
      label:
        "11. No lifting activity to be carried during lightening, heavy wind /rain.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q12",
      label:
        "12. If scaffolding to be used during lift , Scaffolding with valid tag available for use",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q13",
      label:
        "13. Double lanyards Safety Harness/belt checked and in working condition",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_q14",
      label:
        "14. Safety shoes (nonslip), Helmet with chin strip available with employees.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "lift_other",
      label: "15. Other Precautions / Notes",
      type: "text",
      required: false,
    },
  ];

  const hotPrecautions = [
    {
      key: "hot_q1",
      label: "1. Proper Access/ Exit available",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q2",
      label: "2. Proper ventilation and / or lighting provided",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q3",
      label: "3. Proper & Safe scaffolding, platform, ladder provided",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q4",
      label: "4. Welding machine located in a clean and dry area",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q5",
      label:
        "5. Welding machine grounded at the equipment & proper leakage current protection device (ELCB) provided for welding machine.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q6",
      label: "6. Competent and Trained personnel deployed to carry the work.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q7",
      label:
        "7. Welding machine, Input / Output Cables, welding holder and weld return clamp (Holder ) insulated & in good condition",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q8",
      label:
        "8. Welder and fitter trained to connect ground / work return clamps (Holder) to the work piece prior to energization of Welding machine.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q9",
      label:
        "9. Gas Cylinders stacked vertically and not below the welding/cutting area. Regulator Key is available with cylinders.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q10",
      label: "10. Work Area Isolated with barricading and caution sign",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q11",
      label:
        "11. Personal Protective Equipment. Minimum applicable - Safety helmet, safety goggles, welding helmet, safety shoes, leather gloves, long sleeve and nose mask provided.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q12",
      label:
        "12. In case of pits, water removed from the pit & wood /rubber insulation provided.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q13",
      label:
        "13. Adequate & suitable nos. of fire fighting extinguisher provided.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q14",
      label: "14. Near by combustible material removed. Housekeeping Done.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_q15",
      label: "15. Fire watch as standby is in place.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "hot_other",
      label: "16. Other Precautions / Notes",
      type: "text",
      required: false,
    },
  ];

  const heightPrecautions = [
    {
      key: "height_q1",
      label: "1. Scaffolding with valid tag available for use",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q2",
      label: "2. Conducted JST/TBT conducted",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q3",
      label:
        "3. Safety shoes (nonslip), Helmet with chin strip available with employees.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q4",
      label:
        "4. All tightening tools, hand tools /equipment checked and in good condition.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q5",
      label: "5. Access and exist marked and without obstruction.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q6",
      label: "6. Lighting arrangement adequate.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q7",
      label: "7. Unwanted and rubbish material removed from working platform.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q8",
      label: "8. Electrical cable in good condition",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q9",
      label: "9. Signboards provided",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "height_q10",
      label:
        "10. Employees aware about hazards and safe working practices while working at height.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
  ];

  const generalPrecautions = [
    {
      key: "gen_q1",
      label: "1. Proper Access/ Exit available.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q2",
      label: "2. Proper & Safe scaffolding, platform, ladder provided",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q3",
      label: "3. Daily housekeeping of the work area completed",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q4",
      label:
        "4. Identification & protection of any utility services like electric cables, pipes etc. nearby before start of work.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q5",
      label: "5. Checked safe condition of hand tools/ Power tools.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q6",
      label: "6. Conducted JST/TBT conducted",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q7",
      label: "7. Emergency precautionary measures are in place",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q8",
      label: "8. Proper illumination",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q9",
      label:
        "9. Personal Protective Equipment provided. Minimum applicable are safety helmet, safety goggles, safety shoes, Hand gloves, dust mask, Etc.",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q10",
      label:
        "10. Solid & strong barricade provided around excavation/work area",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_q11",
      label: "11. Safety Sign board are in place .",
      type: "select",
      options: ["Yes", "Not Required"],
      required: true,
    },
    {
      key: "gen_other",
      label: "12-16. Other Precautions / Notes",
      type: "text",
      required: false,
    },
  ];

  const assemblePermit = (precautions) => {
    return permitHeaderFields
      .concat(precautions)
      .concat(permitApprovalFields)
      .concat(permitClosingFields);
  };

  const dailyInspectionFields = (items) => {
    const fields = common.concat([
      {
        key: "equipmentId",
        label: "Equipment / Machine No.",
        type: "text",
        required: true,
      },
      { key: "make", label: "Make / Type", type: "text", required: false },
      {
        key: "contractor",
        label: "Contractor Name",
        type: "text",
        required: true,
      },
    ]);
    items.forEach((item, index) => {
      for (let day = 1; day <= 7; day++) {
        fields.push({
          key: "item" + (index + 1) + "Day" + day,
          label: "Item " + (index + 1) + " - Day " + day + ": " + item,
          type: "select",
          options: ["Yes", "No", "NA"],
          required: true,
        });
      }
    });
    return fields.concat([
      {
        key: "supervisorSign",
        label: "Supervisor Name / Signature",
        type: "text",
        required: true,
      },
      {
        key: "safetyOfficerSign",
        label: "Safety Officer Name / Signature",
        type: "text",
        required: true,
      },
      {
        key: "engineerSign",
        label: "Electrical / Mechanical Engineer Name / Signature",
        type: "text",
        required: true,
      },
    ]);
  };

  const inductionAttendees = [];
  for (let row = 1; row <= 15; row++) {
    inductionAttendees.push({
      key: "attendee" + row + "Id",
      label: "Attendee " + row + " ID Card No.",
      type: "text",
      required: false,
    });
    inductionAttendees.push({
      key: "attendee" + row + "Name",
      label: "Attendee " + row + " Name",
      type: "text",
      required: false,
    });
    inductionAttendees.push({
      key: "attendee" + row + "Designation",
      label: "Attendee " + row + " Designation",
      type: "text",
      required: false,
    });
  }

  const attendanceRows = (count, prefix) => {
    const fields = [];
    for (let row = 1; row <= count; row++) {
      fields.push({
        key: prefix + row + "Name",
        label: "Participant " + row + " Name",
        type: "text",
        required: false,
      });
      fields.push({
        key: prefix + row + "Designation",
        label: "Participant " + row + " Designation",
        type: "text",
        required: false,
      });
      fields.push({
        key: prefix + row + "Company",
        label: "Participant " + row + " Company / Agency",
        type: "text",
        required: false,
      });
    }
    return fields;
  };

  const map = {
    WR_WEEKLY: common.concat([
      { key: "weekNo", label: "Week No", type: "text", required: true },
      {
        key: "manpowerAvg",
        label: "Average Manpower",
        type: "number",
        required: true,
      },
      { key: "manHours", label: "Man Hours", type: "number", required: true },
      {
        key: "tbtCount",
        label: "TBT Conducted",
        type: "number",
        required: true,
      },
      {
        key: "inductions",
        label: "Inductions",
        type: "number",
        required: true,
      },
      {
        key: "permits",
        label: "Work Permits Issued",
        type: "number",
        required: true,
      },
      {
        key: "observations",
        label: "Observations Open / Closed",
        type: "text",
        required: true,
      },
      {
        key: "incidents",
        label: "Incidents / Near Miss",
        type: "textarea",
        required: false,
      },
      {
        key: "highlights",
        label: "Highlights / Issues",
        type: "textarea",
        required: true,
      },
    ]),
    WR_MONTHLY: common.concat([
      { key: "month", label: "Month", type: "month", required: true },
      {
        key: "cumManHours",
        label: "Cumulative Man Hours",
        type: "number",
        required: true,
      },
      {
        key: "ltifr",
        label: "LTIFR / TRIR notes",
        type: "text",
        required: false,
      },
      {
        key: "trainings",
        label: "Trainings Conducted",
        type: "textarea",
        required: true,
      },
      {
        key: "audits",
        label: "Audits / Compliance",
        type: "textarea",
        required: true,
      },
      {
        key: "summary",
        label: "Monthly Summary",
        type: "textarea",
        required: true,
      },
    ]),
    WP_SHAFT: assemblePermit(shaftPrecautions),
    WP_NIGHT: assemblePermit(nightPrecautions),
    WP_LIFT: assemblePermit(liftPrecautions),
    WP_HOT: assemblePermit(hotPrecautions),
    WP_HEIGHT: assemblePermit(heightPrecautions),
    WP_GENERAL: assemblePermit(generalPrecautions),
    OBS_DAILY: [
      { key: "date", label: "Date", type: "date", required: true },
      { key: "reportNo", label: "Report No", type: "text", required: true },
      { key: "vendor", label: "Vendor", type: "text", required: true },
      { key: "auditedBy", label: "Audited By", type: "text", required: true },
      { key: "contractor", label: "Contractor", type: "text", required: true },
      { key: "location", label: "Location", type: "text", required: true },
      {
        key: "observation",
        label: "Observation",
        type: "textarea",
        required: true,
      },
      {
        key: "preventive",
        label: "Preventive measures",
        type: "textarea",
        required: true,
      },
      {
        key: "ownerEmployeeId",
        label: "Responsible Lead (Employee ID)",
        type: "text",
        required: true,
      },
    ],
    CL_SCREENING: common.concat([
      { key: "workerId", label: "Worker ID No.", type: "text", required: true },
      {
        key: "contractor",
        label: "Contractor / Sub-contractor",
        type: "text",
        required: true,
      },
      {
        key: "workerName",
        label: "Full Name of Workman",
        type: "text",
        required: true,
      },
      {
        key: "fatherName",
        label: "Father / Husband Name",
        type: "text",
        required: true,
      },
      {
        key: "permanentAddress",
        label: "Permanent Address",
        type: "textarea",
        required: true,
      },
      {
        key: "presentAddress",
        label: "Present Address",
        type: "textarea",
        required: true,
      },
      {
        key: "photoFile",
        label: "Worker Photo File ID / URL",
        type: "text",
        required: false,
      },
      {
        key: "dateOfBirth",
        label: "Date of Birth",
        type: "date",
        required: false,
      },
      {
        key: "sex",
        label: "Sex",
        type: "select",
        options: ["Male", "Female", "Other"],
        required: true,
      },
      { key: "age", label: "Age", type: "number", required: true },
      {
        key: "maritalStatus",
        label: "Marital Status",
        type: "select",
        options: ["Married", "Single", "Widow", "Widower"],
        required: true,
      },
      {
        key: "childrenCount",
        label: "Number of Children",
        type: "number",
        required: false,
      },
      {
        key: "motherTongue",
        label: "Mother Tongue",
        type: "text",
        required: false,
      },
      {
        key: "languages",
        label: "Other Languages Known",
        type: "text",
        required: false,
      },
      {
        key: "emergencyContact",
        label: "Emergency Contact Person",
        type: "text",
        required: true,
      },
      {
        key: "identificationMark",
        label: "Other Identification Mark",
        type: "text",
        required: false,
      },
      {
        key: "vision",
        label: "Vision",
        type: "select",
        options: ["Normal", "Not Normal"],
        required: true,
      },
      {
        key: "visionProblem",
        label: "Vision Problem Details",
        type: "text",
        required: false,
      },
      {
        key: "health",
        label: "Health",
        type: "select",
        options: ["Normal", "Not Normal"],
        required: true,
      },
      {
        key: "healthProblem",
        label: "Health Problem Details",
        type: "text",
        required: false,
      },
      {
        key: "weightKg",
        label: "Weight (KG)",
        type: "number",
        required: false,
      },
      {
        key: "heightCms",
        label: "Height (CMS)",
        type: "number",
        required: false,
      },
      {
        key: "bloodGroup",
        label: "Blood Group",
        type: "text",
        required: false,
      },
      {
        key: "educationDetails",
        label: "Education Details",
        type: "textarea",
        required: false,
      },
      {
        key: "workerDeclarationSignature",
        label: "Worker Declaration Signature / Thumb Impression",
        type: "text",
        required: true,
      },
      {
        key: "contractorDeclarationSignature",
        label: "Contractor Signature",
        type: "text",
        required: true,
      },
      {
        key: "suitableEmployment",
        label: "Suitable for Employment In",
        type: "text",
        required: true,
      },
      {
        key: "siteInCharge",
        label: "Site In-charge / Site Engineer",
        type: "text",
        required: true,
      },
    ]),
    CL_IDCARD: common.concat([
      {
        key: "projectAddress",
        label: "Project Address",
        type: "textarea",
        required: true,
      },
      { key: "workerName", label: "Name", type: "text", required: true },
      { key: "empNo", label: "ID Number", type: "text", required: true },
      {
        key: "designation",
        label: "Designation",
        type: "text",
        required: true,
      },
      { key: "dobAge", label: "DOB / Age", type: "text", required: true },
      { key: "bloodGroup", label: "Blood Group", type: "text", required: true },
      {
        key: "issueDate",
        label: "Date of Issue",
        type: "date",
        required: true,
      },
      {
        key: "validity",
        label: "Validity Up To",
        type: "date",
        required: true,
      },
      {
        key: "authoritySign",
        label: "Authority Signature",
        type: "text",
        required: true,
      },
      {
        key: "screeningDate",
        label: "Screening Done On",
        type: "date",
        required: true,
      },
      {
        key: "screeningDocumentNo",
        label: "Screening Document No.",
        type: "text",
        required: true,
      },
      {
        key: "safetyViolation",
        label: "Safety Violation Status",
        type: "select",
        options: ["Green", "Yellow", "Red"],
        required: true,
      },
      { key: "sign", label: "Signature", type: "text", required: true },
    ]),
    CL_DRILL: common.concat([
      {
        key: "equipmentNo",
        label: "Equipment No.",
        type: "text",
        required: true,
      },
      {
        key: "contractor",
        label: "Contractor Name",
        type: "text",
        required: true,
      },
      {
        key: "inspectionDate",
        label: "Inspection Date",
        type: "date",
        required: true,
      },
      { key: "drill_q1_choice", label: "1. Equipment double insulated", type: "select", options: ["Yes", "No"], required: true },
      { key: "drill_q1_remarks", label: "1. Remarks", type: "text", required: false },
      { key: "drill_q2_choice", label: "2. Equipment free from any defect", type: "select", options: ["Yes", "No"], required: true },
      { key: "drill_q2_remarks", label: "2. Remarks", type: "text", required: false },
      { key: "drill_q3_choice", label: "3. Electrical cable free from defects", type: "select", options: ["Yes", "No"], required: true },
      { key: "drill_q3_remarks", label: "3. Remarks", type: "text", required: false },
      { key: "drill_q4_choice", label: "4. Industrial plug top available and in working condition", type: "select", options: ["Yes", "No"], required: true },
      { key: "drill_q4_remarks", label: "4. Remarks", type: "text", required: false },
      { key: "drill_q5_choice", label: "5. Drilling Bit without any damage", type: "select", options: ["Yes", "No"], required: true },
      { key: "drill_q5_remarks", label: "5. Remarks", type: "text", required: false },
      { key: "drill_q6_choice", label: "6. Drilling bit holder is available and in working condition", type: "select", options: ["Yes", "No"], required: true },
      { key: "drill_q6_remarks", label: "6. Remarks", type: "text", required: false },
      { key: "drill_q7_choice", label: "7. Switch in working condition", type: "select", options: ["Yes", "No"], required: true },
      { key: "drill_q7_remarks", label: "7. Remarks", type: "text", required: false },
      { key: "checkedByName", label: "Checked by Name & Sign", type: "text", required: true },
      { key: "ehsName", label: "EHS SESIPL Name & Sign", type: "text", required: true },
      { key: "reviewStatus", label: "PMC Review Status", type: "select", options: ["Accepted", "Rejected"], required: true },
    ]),
    CL_WELD: dailyInspectionFields([
      "ON / OFF knob undamaged",
      "Regulator with indicator",
      "Welding cables connected with lugs",
      "Welding cable insulation undamaged",
      "Electrode and earthing holders undamaged",
      "Industrial plug available",
      "No exposed live electrical parts",
      "Trolley wheels undamaged",
      "Fire extinguisher and sand bucket available",
    ]),
    CL_GRIND: dailyInspectionFields([
      "Handle free from damage",
      "Wheel guard covers three-fourths area",
      "Grinding wheel free from crack",
      "Rear handle without damage",
      "Cord strain reliever present",
      "Trigger switch undamaged",
      "Dead man switch present",
      "Electrical wire without cut or joint",
      "Plug top provided",
      "Machine body undamaged",
    ]),
    CL_CUT: dailyInspectionFields([
      "Cutting blade defined and undamaged",
      "Safety guard available and good",
      "Lock system for plate and guard",
      "Job clamp / fence available",
      "Handle available and good",
      "Cable connection undamaged",
      "Dust guard / chip deflector available",
      "Machine base undamaged",
    ]),
    CL_TBT: common
      .concat([
        {
          key: "topic",
          label: "Topic Discussed",
          type: "text",
          required: true,
        },
        {
          key: "conductedBy",
          label: "Tool Box Talk Conducted By",
          type: "text",
          required: true,
        },
        {
          key: "projectManager",
          label: "Project Manager Name",
          type: "text",
          required: true,
        },
        {
          key: "keyPoints",
          label: "Key Points / Hazards Discussed",
          type: "textarea",
          required: true,
        },
      ])
      .concat(attendanceRows(25, "participant")),
    CL_JST: common
      .concat([
        { key: "topic", label: "Training Topic", type: "text", required: true },
        {
          key: "trainingTime",
          label: "Training Time",
          type: "time",
          required: true,
        },
        {
          key: "conductedBy",
          label: "Conducted By",
          type: "text",
          required: true,
        },
        {
          key: "acknowledgement",
          label: "Training acknowledgement / hazards and precautions",
          type: "textarea",
          required: true,
        },
        {
          key: "projectManagerSignature",
          label: "Project Manager Signature",
          type: "text",
          required: true,
        },
        {
          key: "ehsOfficerSignature",
          label: "EHS Officer Signature",
          type: "text",
          required: true,
        },
      ])
      .concat(attendanceRows(20, "participant")),
    CL_INDUCTION: common
      .concat([
        {
          key: "contractor",
          label: "Contractor Name",
          type: "text",
          required: true,
        },
        {
          key: "topicsCovered",
          label: "Safety Induction Topics Covered",
          type: "textarea",
          required: true,
        },
        {
          key: "safetyOfficer",
          label: "Safety Officer Name",
          type: "text",
          required: true,
        },
        {
          key: "projectManager",
          label: "Project Manager Name",
          type: "text",
          required: true,
        },
      ])
      .concat(inductionAttendees),
    CL_SCAFFOLD: dailyInspectionFields([
      "All coupler hooks are properly installed",
      "Proper platform has been provided & is having proper locking system",
      "Proper platform has been provided & is having proper locking system",
      "Inspection tag has been displayed",
      "Scaffolding has been erected on a firm base",
      "Wheel lock has been provided & is in working condition",
      "Toe board has been provided",
      "Access ladder has been provided & installed properly",
      "Double (top & mid) railing has been provided",
    ]),
    CL_MEDICAL: common.concat([
      {
        key: "certificateNo",
        label: "Certificate Serial No.",
        type: "text",
        required: true,
      },
      { key: "workerName", label: "Worker Name", type: "text", required: true },
      {
        key: "fatherName",
        label: "Father / Husband Name",
        type: "text",
        required: true,
      },
      {
        key: "identificationMarks",
        label: "Identification Marks",
        type: "textarea",
        required: false,
      },
      {
        key: "sex",
        label: "Sex",
        type: "select",
        options: ["Male", "Female", "Other"],
        required: true,
      },
      {
        key: "residence",
        label: "Residence",
        type: "textarea",
        required: true,
      },
      {
        key: "dateOfBirth",
        label: "Date of Birth",
        type: "date",
        required: false,
      },
      {
        key: "physicalFitness",
        label: "Physical Fitness / Fit for Employment",
        type: "select",
        options: ["Fit", "Not Fit"],
        required: true,
      },
      {
        key: "medicalInspector",
        label: "Medical Inspector / CMO",
        type: "text",
        required: true,
      },
    ]),
    TAG_IND: common.concat([
      { key: "workerName", label: "Name", type: "text", required: true },
      {
        key: "inductionId",
        label: "Induction ID",
        type: "text",
        required: true,
      },
      {
        key: "inductionDate",
        label: "Date of Induction",
        type: "date",
        required: true,
      },
      {
        key: "designation",
        label: "Designation / Trade",
        type: "text",
        required: true,
      },
      {
        key: "emergencyNumber",
        label: "Emergency Number",
        type: "text",
        required: true,
      },
    ]),
    TAG_TOOL: common.concat([
      { key: "machineNo", label: "M/C No.", type: "text", required: true },
      { key: "makeType", label: "Make / Type", type: "text", required: true },
      {
        key: "inspectionDate",
        label: "Date of Inspection",
        type: "date",
        required: true,
      },
      {
        key: "nextDue",
        label: "Due Date of Inspection",
        type: "date",
        required: true,
      },
      {
        key: "inspectionResult",
        label: "Inspected OK / Not OK",
        type: "select",
        options: ["OK", "Not OK"],
        required: true,
      },
      { key: "signature", label: "Signature", type: "text", required: true },
    ]),
    TAG_FE: common.concat([
      { key: "location", label: "Location", type: "text", required: true },
      {
        key: "inspectionDate",
        label: "Inspection Date",
        type: "date",
        required: true,
      },
      {
        key: "nextDue",
        label: "Next Inspection Date",
        type: "date",
        required: true,
      },
      {
        key: "inspectionResult",
        label: "Safe for Use",
        type: "select",
        options: ["Safe for Use", "Do Not Use"],
        required: true,
      },
      {
        key: "inspectionRemarks",
        label: "Inspection Remarks",
        type: "textarea",
        required: false,
      },
    ]),
    TAG_RED: common.concat([
      { key: "item", label: "Item / Equipment", type: "text", required: true },
      { key: "defect", label: "Defect", type: "textarea", required: true },
      { key: "isolatedBy", label: "Isolated By", type: "text", required: true },
    ]),
    TAG_SCAFF: common.concat([
      { key: "scaffoldId", label: "Scaffold ID", type: "text", required: true },
      {
        key: "tagColour",
        label: "Tag Colour",
        type: "select",
        options: ["Green - Safe for Use", "Red - Do Not Use"],
        required: true,
      },
      {
        key: "inspectionDate",
        label: "Inspection Date",
        type: "date",
        required: true,
      },
      {
        key: "inspectionRemarks",
        label: "Inspection Record / Remarks",
        type: "textarea",
        required: false,
      },
    ]),
    CL_FE: common.concat([
      {
        key: "extinguisherNo",
        label: "Extinguisher No.",
        type: "text",
        required: true,
      },
      { key: "ehsName", label: "EHS Name", type: "text", required: true },
      {
        key: "inspectionDate",
        label: "Inspection Date",
        type: "date",
        required: true,
      },
      {
        key: "nextInspectionDate",
        label: "Next Inspection Date",
        type: "date",
        required: true,
      },
      {
        key: "extinguisherSpecification",
        label: "Extinguisher Specifications / Type",
        type: "select",
        options: ["ABC Type", "Water Type", "CO2 Type"],
        required: true,
      },
      {
        key: "conditionClean",
        label: "Extinguisher clean and tidy",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "corroded",
        label: "Extinguisher corroded",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "safetyPin",
        label: "Safety pin in locking position",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "nozzleCondition",
        label: "Discharge nozzle condition",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "hoseCondition",
        label: "Hose condition",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "weightMatches",
        label: "Weight matches body marking",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "visualBoard",
        label: "Visual board available",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "gaugeGreen",
        label: "Indicator gauge in green",
        type: "select",
        options: ["Yes", "No", "NA"],
        required: true,
      },
      {
        key: "expiryStatus",
        label: "Not expired / refill date valid",
        type: "select",
        options: ["Yes", "No"],
        required: true,
      },
      {
        key: "otherRemarks",
        label: "Other Remarks",
        type: "textarea",
        required: false,
      },
      {
        key: "checkedBy",
        label: "Checked By Safety Officer",
        type: "text",
        required: true,
      },
      {
        key: "reviewStatus",
        label: "Reviewed Status",
        type: "select",
        options: ["Accepted", "Rejected"],
        required: true,
      },
    ]),
  };
  if (formCode.indexOf("_UPLOAD") !== -1) {
    return [
      { key: "title", label: "Document Title", type: "text", required: true },
      { key: "docDate", label: "Document Date", type: "date", required: false },
      { key: "expiry", label: "Expiry Date", type: "date", required: false },
      { key: "remarks", label: "Remarks", type: "textarea", required: false },
    ];
  }
  return (
    map[formCode] ||
    common.concat([
      { key: "remarks", label: "Remarks", type: "textarea", required: false },
    ])
  );
}
