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
  WP_HEIGHT: "1RAX7rxlZ7fY3ec3sH9OGOcZyRGD4nRKkUDkgjpgOy5w", // Working At Height Permit-04
  WP_GENERAL: "1wekm-P4Jv3iz85gBld-dWNuSzVLjDlr05j4HzQeWamY", // General work permit-02
  WP_HOT: "1XmYhVlBiJtK32y-IXRsX-5rxP_8ykACny4C-MuAl0HU", // Hot Work Permit-03
  WP_LIFT: "1EZEqr-YwsyFNdUI191MV-lehWlgQmVmb9z-oMxTygs0", // Lifting Activity Permit-05
  WP_SHAFT: "1dp9UV5F0OEnxq7KarUhIJUXwzZQCFH-rZTgcLWTp8mE", // Shaft work permit-06
  WP_NIGHT: "1MYVyt8LRIoBIDQRjX8W_pS7NdpVm88Me1wRWawUJSGE", // Night Work permit-07

  // Checklists & Equipment Inspections
  CL_WELD: "1IYwFad3phxQ3VeZTbsu6INBvDfJWhSIEdBmKRcQjS6Q", // Welding Machine
  CL_GRIND: "1c7T1ip1jnnz9Tc0b61fASMYOBKC2z2zonVN1SUC3aVA", // Grinding Machine Checklist
  CL_DRILL: "1x18zs1fBXgPJqlejQqUgMIy6juMhQb_rQSGMZOdt9ug", // Drilling Machine
  CL_FE: "1H87LMWzmqgIWHPpm54O4vQUrq6rvh4yAi6u5ZPDelns", // Fire Extinguisher
  CL_SCAFFOLD: "l0KyOAHOAr7ulkv3TPgHOJqF3vaywUvD9mncVF1s", // Scaffolding checklist
  CL_CUT: "",

  // Attendances, Training & Inductions
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

/**
 * Resolves the Google Doc template ID for a given formCode,
 * checking runtime properties first and falling back to backend defaults.
 */
function getDocTemplateId_(formCode) {
  if (
    BACKEND_DOCS_TEMPLATES[formCode] &&
    BACKEND_DOCS_TEMPLATES[formCode].trim() !== ""
  ) {
    return BACKEND_DOCS_TEMPLATES[formCode].trim();
  }
  let id = "";
  try {
    if (typeof getDocTemplateRegistry_ === "function") {
      const reg = getDocTemplateRegistry_();
      id = (reg && reg[formCode]) || "";
    }
  } catch (e) {}
  return id ? id.trim() : "";
}

/**
 * Builds a comprehensive dictionary of placeholder replacements.
 * Matches all standard and specific tags for permits, checklists, and worker forms.
 */
function buildPlaceholderMap_(formCode, fields, project, user, submissionId) {
  fields = fields || {};
  project = project || {};
  user = user || {};

  const docCode =
    (typeof PERMIT_DOC_CODES !== "undefined" && PERMIT_DOC_CODES[formCode]) ||
    formCode ||
    "SESIPL-EHS";
  const contractorName =
    fields.contractorName ||
    fields.contractor ||
    project.client ||
    "Shankar Electricals Services I Pvt Ltd";
  const permitNo =
    fields.permitNo ||
    "SESIPL/" +
      formCode.replace("WP_", "") +
      "/" +
      (submissionId ? String(submissionId).slice(-4) : "001");
  const emergency1 = fields.emergencyContact1 || "9591461971";
  const emergency2 = fields.emergencyContact2 || "7899650058";

  const map = {
    // Header & Document Info
    contractorName: contractorName,
    ContractorName: contractorName,
    contractor: contractorName,
    projectName: project.name || "SESIPL Site",
    projectCode: project.code || "PRJ",
    permitNo: permitNo,
    PermitNo: permitNo,
    docCode: docCode,
    emergencyContact1: emergency1,
    emergencyContact2: emergency2,

    // General Work Details
    area: fields.area || "",
    location: fields.location || "",
    date: fields.date
      ? displayDate_(fields.date)
      : fields.workExecutionDate
        ? displayDate_(fields.workExecutionDate)
        : todayDisplay_(),
    time: fields.time || "",
    siteEngineerName: fields.siteEngineer || fields.siteEngineerName || "",
    SiteEngineerName: fields.siteEngineer || fields.siteEngineerName || "",
    siteEngineer: fields.siteEngineer || fields.siteEngineerName || "",
    requestingSiteEngineerName: fields.siteEngineer || fields.siteEngineerName || "",
    RequestingSiteEngineerName: fields.siteEngineer || fields.siteEngineerName || "",
    requestingSiteEngineer: fields.siteEngineer || fields.siteEngineerName || "",
    siteEngineerSign:
      fields.siteEngineerSign ||
      (fields.siteEngineer
        ? "Signed - " + fields.siteEngineer
        : fields.siteEngineerName
          ? "Signed"
          : ""),
    SiteEngineerSign:
      fields.siteEngineerSign ||
      (fields.siteEngineer
        ? "Signed - " + fields.siteEngineer
        : fields.siteEngineerName
          ? "Signed"
          : ""),
    requestingSiteEngineerSign:
      fields.siteEngineerSign ||
      (fields.siteEngineer
        ? "Signed - " + fields.siteEngineer
        : fields.siteEngineerName
          ? "Signed"
          : ""),
    RequestingSiteEngineerSign:
      fields.siteEngineerSign ||
      (fields.siteEngineer
        ? "Signed - " + fields.siteEngineer
        : fields.siteEngineerName
          ? "Signed"
          : ""),
    siteEngineerDate: fields.siteEngineerDate
      ? displayDate_(fields.siteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    SiteEngineerDate: fields.siteEngineerDate
      ? displayDate_(fields.siteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    requestingSiteEngineerDate: fields.siteEngineerDate
      ? displayDate_(fields.siteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    RequestingSiteEngineerDate: fields.siteEngineerDate
      ? displayDate_(fields.siteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    siteEngineerTime: fields.siteEngineerTime || fields.time || "",
    SiteEngineerTime: fields.siteEngineerTime || fields.time || "",
    requestingSiteEngineerTime: fields.siteEngineerTime || fields.time || "",
    RequestingSiteEngineerTime: fields.siteEngineerTime || fields.time || "",

    safetyOfficerName: fields.safetyOfficer || fields.safetyOfficerName || "",
    SafetyOfficerName: fields.safetyOfficer || fields.safetyOfficerName || "",
    safetyOfficer: fields.safetyOfficer || fields.safetyOfficerName || "",
    safetyOfficerSign:
      fields.safetyOfficerSign ||
      (fields.safetyOfficer
        ? "Signed - " + fields.safetyOfficer
        : fields.safetyOfficerName
          ? "Signed"
          : ""),
    SafetyOfficerSign:
      fields.safetyOfficerSign ||
      (fields.safetyOfficer
        ? "Signed - " + fields.safetyOfficer
        : fields.safetyOfficerName
          ? "Signed"
          : ""),
    safetyOfficerDate: fields.safetyOfficerDate
      ? displayDate_(fields.safetyOfficerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    SafetyOfficerDate: fields.safetyOfficerDate
      ? displayDate_(fields.safetyOfficerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    safetyOfficerTime: fields.safetyOfficerTime || fields.time || "",
    SafetyOfficerTime: fields.safetyOfficerTime || fields.time || "",

    contractorSiteIncharge:
      fields.contractorInCharge ||
      fields.contractorSiteIncharge ||
      fields.agencySupervisor ||
      "",
    contractorInCharge:
      fields.contractorInCharge ||
      fields.contractorSiteIncharge ||
      fields.agencySupervisor ||
      "",
    contactNumber: fields.contactNumber || "",
    descriptionOfWork: fields.descriptionOfWork || fields.workDescription || "",
    workDescription: fields.descriptionOfWork || fields.workDescription || "",
    workExecutionDate: fields.workExecutionDate
      ? displayDate_(fields.workExecutionDate)
      : "",
    validFrom: fields.validFrom
      ? String(fields.validFrom).replace("T", " ")
      : "",
    validTo: fields.validTo ? String(fields.validTo).replace("T", " ") : "",
    PMCsiteEngineer:
      fields.closingPmcSiteEngName ||
      fields.pmcSiteEngineer ||
      fields.PMCsiteEngineer ||
      "",
    pmcSiteEngineer:
      fields.closingPmcSiteEngName ||
      fields.pmcSiteEngineer ||
      fields.PMCsiteEngineer ||
      "",

    // Permit Specific
    shaftWorkmenNames: fields.workmenNames || fields.shaftWorkmenNames || "",
    workmenNames: fields.workmenNames || fields.shaftWorkmenNames || "",
    hot_other: fields.hot_other || "",
    lift_other: fields.lift_other || "",
    gen_other: fields.gen_other || "",
    height_other: fields.height_other || "",
    other: fields.lift_other || fields.hot_other || fields.gen_other || fields.height_other || fields.night_remarks || "",
    night_remarks: fields.night_remarks || "",
    remarks: fields.night_remarks || fields.lift_other || "",

    // Reviewed & Approved By (Permit Issuing Authority)
    approvalEhsName: fields.approvalEhsName || "",
    approvalEhsOfficerName: fields.approvalEhsName || "",
    ApprovalEhsName: fields.approvalEhsName || "",
    approvalEhsSign:
      fields.approvalEhsSign || (fields.approvalEhsName ? "Signed" : ""),
    ApprovalEhsSign:
      fields.approvalEhsSign || (fields.approvalEhsName ? "Signed" : ""),
    approvalEhsDate: fields.approvalEhsDate
      ? displayDate_(fields.approvalEhsDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    ApprovalEhsDate: fields.approvalEhsDate
      ? displayDate_(fields.approvalEhsDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    approvalEhsTime: fields.approvalEhsTime || fields.time || "",
    ApprovalEhsTime: fields.approvalEhsTime || fields.time || "",

    approvalSiteEngName:
      fields.approvalSiteEngineerName || fields.approvalSiteEngName || "",
    approvalSiteEngineerName:
      fields.approvalSiteEngineerName || fields.approvalSiteEngName || "",
    ApprovalSiteEngineerName:
      fields.approvalSiteEngineerName || fields.approvalSiteEngName || "",
    issuingSiteEngineerName:
      fields.approvalSiteEngineerName || fields.approvalSiteEngName || fields.siteEngineer || fields.siteEngineerName || "",
    IssuingSiteEngineerName:
      fields.approvalSiteEngineerName || fields.approvalSiteEngName || fields.siteEngineer || fields.siteEngineerName || "",
    issuingSiteEngineer:
      fields.approvalSiteEngineerName || fields.approvalSiteEngName || fields.siteEngineer || fields.siteEngineerName || "",
    approvalSiteEngSign:
      fields.approvalSiteEngineerSign ||
      fields.approvalSiteEngSign ||
      (fields.approvalSiteEngineerName ? "Signed" : ""),
    approvalSiteEngineerSign:
      fields.approvalSiteEngineerSign ||
      fields.approvalSiteEngSign ||
      (fields.approvalSiteEngineerName ? "Signed" : ""),
    ApprovalSiteEngineerSign:
      fields.approvalSiteEngineerSign ||
      fields.approvalSiteEngSign ||
      (fields.approvalSiteEngineerName ? "Signed" : ""),
    issuingSiteEngineerSign:
      fields.approvalSiteEngineerSign ||
      fields.approvalSiteEngSign ||
      (fields.approvalSiteEngineerName ? "Signed" : (fields.siteEngineerSign ? fields.siteEngineerSign : "Signed")),
    IssuingSiteEngineerSign:
      fields.approvalSiteEngineerSign ||
      fields.approvalSiteEngSign ||
      (fields.approvalSiteEngineerName ? "Signed" : (fields.siteEngineerSign ? fields.siteEngineerSign : "Signed")),
    approvalSiteEngDate: fields.approvalSiteEngineerDate
      ? displayDate_(fields.approvalSiteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    approvalSiteEngineerDate: fields.approvalSiteEngineerDate
      ? displayDate_(fields.approvalSiteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    ApprovalSiteEngineerDate: fields.approvalSiteEngineerDate
      ? displayDate_(fields.approvalSiteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    issuingSiteEngineerDate: fields.approvalSiteEngineerDate
      ? displayDate_(fields.approvalSiteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    IssuingSiteEngineerDate: fields.approvalSiteEngineerDate
      ? displayDate_(fields.approvalSiteEngineerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    approvalSiteEngTime: fields.approvalSiteEngineerTime || fields.time || "",
    approvalSiteEngineerTime: fields.approvalSiteEngineerTime || fields.time || "",
    ApprovalSiteEngineerTime: fields.approvalSiteEngineerTime || fields.time || "",
    issuingSiteEngineerTime: fields.approvalSiteEngineerTime || fields.time || "",
    IssuingSiteEngineerTime: fields.approvalSiteEngineerTime || fields.time || "",

    // Permit Closing / Cancellation
    // Requesting Authority
    closingSiteEngName: fields.closingSiteEngName || fields.siteEngineer || fields.siteEngineerName || "",
    closingSiteEngineerName: fields.closingSiteEngName || fields.siteEngineer || fields.siteEngineerName || "",
    ClosingSiteEngineerName: fields.closingSiteEngName || fields.siteEngineer || fields.siteEngineerName || "",
    closingRequestingSiteEngineerName: fields.closingSiteEngName || fields.siteEngineer || fields.siteEngineerName || "",
    closingSiteEngSign:
      fields.closingSiteEngSign || (fields.closingSiteEngName ? "Signed" : (fields.siteEngineer ? "Signed - " + fields.siteEngineer : "Signed")),
    closingSiteEngineerSign:
      fields.closingSiteEngSign || (fields.closingSiteEngName ? "Signed" : (fields.siteEngineer ? "Signed - " + fields.siteEngineer : "Signed")),
    ClosingSiteEngineerSign:
      fields.closingSiteEngSign || (fields.closingSiteEngName ? "Signed" : (fields.siteEngineer ? "Signed - " + fields.siteEngineer : "Signed")),
    closingRequestingSiteEngineerSign:
      fields.closingSiteEngSign || (fields.closingSiteEngName ? "Signed" : (fields.siteEngineer ? "Signed - " + fields.siteEngineer : "Signed")),
    closingSiteEngDate: fields.closingSiteEngDate
      ? displayDate_(fields.closingSiteEngDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    closingSiteEngineerDate: fields.closingSiteEngDate
      ? displayDate_(fields.closingSiteEngDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    ClosingSiteEngineerDate: fields.closingSiteEngDate
      ? displayDate_(fields.closingSiteEngDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    closingRequestingSiteEngineerDate: fields.closingSiteEngDate
      ? displayDate_(fields.closingSiteEngDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    closingSiteEngTime: fields.closingSiteEngTime || fields.time || "",
    closingSiteEngineerTime: fields.closingSiteEngTime || fields.time || "",
    ClosingSiteEngineerTime: fields.closingSiteEngTime || fields.time || "",
    closingRequestingSiteEngineerTime: fields.closingSiteEngTime || fields.time || "",

    closingSafetyOfficerName:
      fields.closingSafetyOfficerName ||
      fields.safetyOfficer ||
      fields.safetyOfficerName ||
      "",
    closingSafetyOfficerSign:
      fields.closingSafetyOfficerSign ||
      (fields.closingSafetyOfficerName
        ? "Signed"
        : (fields.safetyOfficer
          ? "Signed - " + fields.safetyOfficer
          : (fields.safetyOfficerName ? "Signed" : ""))),
    closingSafetyOfficerDate: fields.closingSafetyOfficerDate
      ? displayDate_(fields.closingSafetyOfficerDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    closingSafetyOfficerTime: fields.closingSafetyOfficerTime || fields.time || "",

    // Issuing Authority (PMC / Site Engineer)
    closingPmcSiteEngName:
      fields.closingPmcSiteEngName ||
      fields.pmcSiteEngineer ||
      fields.PMCsiteEngineer ||
      fields.approvalSiteEngineerName ||
      fields.approvalSiteEngName ||
      "",
    closingPmcSiteEngineerName:
      fields.closingPmcSiteEngName ||
      fields.pmcSiteEngineer ||
      fields.PMCsiteEngineer ||
      fields.approvalSiteEngineerName ||
      fields.approvalSiteEngName ||
      "",
    ClosingPmcSiteEngineerName:
      fields.closingPmcSiteEngName ||
      fields.pmcSiteEngineer ||
      fields.PMCsiteEngineer ||
      fields.approvalSiteEngineerName ||
      fields.approvalSiteEngName ||
      "",
    closingIssuingSiteEngineerName:
      fields.closingPmcSiteEngName ||
      fields.pmcSiteEngineer ||
      fields.PMCsiteEngineer ||
      fields.approvalSiteEngineerName ||
      fields.approvalSiteEngName ||
      "",
    ClosingIssuingSiteEngineerName:
      fields.closingPmcSiteEngName ||
      fields.pmcSiteEngineer ||
      fields.PMCsiteEngineer ||
      fields.approvalSiteEngineerName ||
      fields.approvalSiteEngName ||
      "",
    closingPmcSiteEngSign:
      fields.closingPmcSiteEngSign ||
      (fields.closingPmcSiteEngName ? "Signed" : "Signed"),
    closingPmcSiteEngineerSign:
      fields.closingPmcSiteEngSign ||
      (fields.closingPmcSiteEngName ? "Signed" : "Signed"),
    ClosingPmcSiteEngineerSign:
      fields.closingPmcSiteEngSign ||
      (fields.closingPmcSiteEngName ? "Signed" : "Signed"),
    closingIssuingSiteEngineerSign:
      fields.closingPmcSiteEngSign ||
      (fields.closingPmcSiteEngName ? "Signed" : "Signed"),
    closingPmcSiteEngDate: fields.closingPmcSiteEngDate
      ? displayDate_(fields.closingPmcSiteEngDate)
      : (fields.approvalSiteEngineerDate
        ? displayDate_(fields.approvalSiteEngineerDate)
        : (fields.date ? displayDate_(fields.date) : todayDisplay_())),
    closingPmcSiteEngineerDate: fields.closingPmcSiteEngDate
      ? displayDate_(fields.closingPmcSiteEngDate)
      : (fields.approvalSiteEngineerDate
        ? displayDate_(fields.approvalSiteEngineerDate)
        : (fields.date ? displayDate_(fields.date) : todayDisplay_())),
    ClosingPmcSiteEngineerDate: fields.closingPmcSiteEngDate
      ? displayDate_(fields.closingPmcSiteEngDate)
      : (fields.approvalSiteEngineerDate
        ? displayDate_(fields.approvalSiteEngineerDate)
        : (fields.date ? displayDate_(fields.date) : todayDisplay_())),
    closingIssuingSiteEngineerDate: fields.closingPmcSiteEngDate
      ? displayDate_(fields.closingPmcSiteEngDate)
      : (fields.approvalSiteEngineerDate
        ? displayDate_(fields.approvalSiteEngineerDate)
        : (fields.date ? displayDate_(fields.date) : todayDisplay_())),
    closingPmcSiteEngTime: fields.closingPmcSiteEngTime || fields.approvalSiteEngineerTime || fields.time || "",
    closingPmcSiteEngineerTime: fields.closingPmcSiteEngTime || fields.approvalSiteEngineerTime || fields.time || "",
    ClosingPmcSiteEngineerTime: fields.closingPmcSiteEngTime || fields.approvalSiteEngineerTime || fields.time || "",
    closingIssuingSiteEngineerTime: fields.closingPmcSiteEngTime || fields.approvalSiteEngineerTime || fields.time || "",

    // Worker Screening & Medical
    workerId: fields.workerId || "",
    workerFullName: fields.workerName || fields.workerFullName || "",
    fatherOrHusbandName: fields.fatherName || fields.fatherOrHusbandName || "",
    permanentAddress: fields.permanentAddress || "",
    presentAddress: fields.presentAddress || "",
    dob: fields.dateOfBirth
      ? displayDate_(fields.dateOfBirth)
      : fields.dob
        ? displayDate_(fields.dob)
        : "",
    sex: fields.sex || "",
    age: fields.age || "",
    maritalStatus: fields.maritalStatus || "",
    numChildren: fields.childrenCount || fields.numChildren || "",
    motherTongue: fields.motherTongue || "",
    otherLanguages: fields.languages || fields.otherLanguages || "",
    emergencyContactPerson:
      fields.emergencyContact || fields.emergencyContactPerson || "",
    identificationMark:
      fields.identificationMark || fields.identificationMarks || "",
    visionStatus: fields.vision || "Normal",
    visionProblem: fields.visionProblem || "",
    healthStatus: fields.health || "Normal",
    healthProblem: fields.healthProblem || "",
    weightKg: fields.weightKg || "",
    heightCm: fields.heightCms || fields.heightCm || "",
    bloodGroup: fields.bloodGroup || "",
    suitableEmployment: fields.suitableEmployment || "",
    siteInchargeName: fields.siteInCharge || fields.siteInchargeName || "",
    workerSign: fields.workerDeclarationSignature || "Signed",
    contractorSign: fields.contractorDeclarationSignature || "Signed",

    // Form XI Medical
    certSerialNo: fields.certificateNo || fields.certSerialNo || "",
    workerThumbSign: fields.workerDeclarationSignature || "Thumb Impressed",
    medicalOfficerSign: fields.medicalInspector || "Verified & Sealed",

    // TBT & Training (JST)
    topicDiscussed: fields.topic || fields.topicDiscussed || "",
    topic: fields.topic || fields.topicDiscussed || "",
    jstTopic: fields.topic || fields.topicDiscussed || "",
    trainingTime: fields.trainingTime || fields.time || "",
    conductedBy: fields.conductedBy || fields.tbtConductedBy || "",
    tbtConductedBy: fields.conductedBy || fields.tbtConductedBy || "",
    conductedBySign:
      fields.conductedBy ? "Signed - " + fields.conductedBy : "Signed",
    tbtConductedBySign:
      fields.conductedBy ? "Signed - " + fields.conductedBy : "Signed",
    projectManager:
      fields.projectManager || fields.projectManagerName || "",
    projectManagerName:
      fields.projectManager || fields.projectManagerName || "",
    projectManagerSign:
      fields.projectManagerSignature ||
      fields.projectManagerSign ||
      (fields.projectManager ? "Signed - " + fields.projectManager : "Signed"),
    projectManagerSignature:
      fields.projectManagerSignature ||
      fields.projectManagerSign ||
      (fields.projectManager ? "Signed - " + fields.projectManager : "Signed"),
    keyPoints: fields.keyPoints || fields.keyPointsDiscussed || "",
    keyPointsDiscussed: fields.keyPoints || fields.keyPointsDiscussed || "",
    acknowledgement:
      fields.acknowledgement || fields.keyPoints || fields.keyPointsDiscussed || "",
    ehsOfficerSign: fields.ehsOfficerSignature || fields.ehsOfficerSign || "Signed",
    ehsOfficerSignature: fields.ehsOfficerSignature || fields.ehsOfficerSign || "Signed",

    // Equipment Details
    equipmentId:
      fields.equipmentId || fields.equipmentNo || fields.machineNo || "",
    equipmentNo:
      fields.equipmentNo || fields.equipmentId || fields.machineNo || "",
    machineNo:
      fields.machineNo || fields.equipmentId || fields.equipmentNo || "",
    make: fields.make || fields.makeType || "",
    makeType: fields.makeType || fields.make || "",
    supervisor: fields.supervisor || "",
    supervisorSign:
      fields.supervisorSign ||
      (fields.supervisor ? "Signed - " + fields.supervisor : "Signed"),
    engineerSign:
      fields.engineerSign ||
      (fields.engineer ? "Signed - " + fields.engineer : "Signed"),
    checkedByName:
      fields.checkedByName || fields.checkedBy || fields.supervisor || "",
    checkedBySign:
      fields.checkedByName ||
      fields.supervisorSign ||
      (fields.checkedBy ? "Signed - " + fields.checkedBy : "Signed"),
    ehsName:
      fields.ehsName ||
      fields.safetyOfficer ||
      fields.approvalEhsName ||
      "Ravikiran - EHS SESIPL",
    ehsSign:
      fields.ehsName ||
      fields.safetyOfficerSign ||
      fields.approvalEhsSign ||
      "Signed",
    reviewStatus: fields.reviewStatus || "Accepted",
    accepted:
      String(fields.reviewStatus || "").toLowerCase() === "accepted" ||
      String(fields.accepted || "").toLowerCase() === "yes"
        ? "✓"
        : "",
    rejected:
      String(fields.reviewStatus || "").toLowerCase() === "rejected" ||
      String(fields.rejected || "").toLowerCase() === "yes"
        ? "✓"
        : "",
    inspectionDate: fields.inspectionDate
      ? displayDate_(fields.inspectionDate)
      : (fields.date ? displayDate_(fields.date) : todayDisplay_()),
    nextInspectionDate: fields.nextInspectionDate
      ? displayDate_(fields.nextInspectionDate)
      : fields.nextDue
        ? displayDate_(fields.nextDue)
        : "",
  };

  // Process all keys in fields directly
  for (const k in fields) {
    if (fields[k] != null && !map[k]) {
      map[k] = String(fields[k]);
    }
  }

  // Generate Yes / No / NA / Not Required checkbox indicators
  for (const k in fields) {
    const val = String(fields[k] || "")
      .trim()
      .toLowerCase();
    const isYes = val === "yes" || val === "y" || val === "true";
    const isNo = val === "no" || val === "n";
    const isNA = val === "na" || val === "n/a";
    const isNR = val === "not required" || val === "nr";

    // Form-specific prefix (e.g. height_q1_yes, night_q1_no)
    map[k + "_yes"] = isYes ? "✓" : "";
    map[k + "_no"] = isNo ? "✓" : "";
    map[k + "_na"] = isNA ? "✓" : "";
    map[k + "_nr"] = isNR ? "✓" : "";

    // Generic prefix (e.g. height_q1 -> q1_yes, 1_yes)
    const match = k.match(
      /^(?:height|gen|hot|lift|shaft|night|wm|grind|cut|drill|fe)_?(q?\d+)/i,
    );
    if (match) {
      const numPart = match[1].toLowerCase().replace("q", "");
      const qTag = "q" + numPart;
      if (k.toLowerCase().indexOf("choice") !== -1 || isYes || isNo || isNA || isNR) {
        map[qTag + "_yes"] = isYes ? "✓" : map[qTag + "_yes"] || "";
        map[qTag + "_no"] = isNo ? "✓" : map[qTag + "_no"] || "";
        map[qTag + "_na"] = isNA ? "✓" : map[qTag + "_na"] || "";
        map[qTag + "_nr"] = isNR ? "✓" : map[qTag + "_nr"] || "";

        map[numPart + "_yes"] = isYes ? "✓" : map[numPart + "_yes"] || "";
        map[numPart + "_no"] = isNo ? "✓" : map[numPart + "_no"] || "";
        map[numPart + "_na"] = isNA ? "✓" : map[numPart + "_na"] || "";
        map[numPart + "_nr"] = isNR ? "✓" : map[numPart + "_nr"] || "";
      }
      if (k.toLowerCase().indexOf("remarks") !== -1) {
        map[qTag + "_remarks"] = fields[k] || "";
        map["remarks_" + numPart] = fields[k] || "";
      } else {
        map[qTag] = fields[k];
      }
    }
  }

  // Handle daily inspection checklists like item1Day1
  for (let item = 1; item <= 15; item++) {
    for (let day = 1; day <= 7; day++) {
      const fieldKey = "item" + item + "Day" + day;
      if (fields[fieldKey]) {
        const v = String(fields[fieldKey]).trim();
        const isTick =
          v.toLowerCase() === "yes" || v.toLowerCase() === "y" || v === "✓";
        const displayVal = isTick ? "✓" : v;
        map["wm_q" + item + "_d" + day] = displayVal;
        map["grind_q" + item + "_d" + day] = displayVal;
        map["cut_q" + item + "_d" + day] = displayVal;
        map["drill_q" + item + "_d" + day] = displayVal;
        map["q" + item + "_d" + day] = displayVal;
        map["item" + item + "_d" + day] = displayVal;
        map["item" + item + "Day" + day] = displayVal;
        map["item" + item + "_day" + day] = displayVal;
        map["i" + item + "_d" + day] = displayVal;
        map["q" + item + "d" + day] = displayVal;
        map["d" + day + "_q" + item] = displayVal;
        map["day" + day + "_q" + item] = displayVal;

        if (day === 1) {
          const isYes = v.toLowerCase() === "yes" || v.toLowerCase() === "y" || v === "✓";
          const isNo = v.toLowerCase() === "no" || v.toLowerCase() === "n";
          const isNA = v.toLowerCase() === "na" || v.toLowerCase() === "n/a";
          map["q" + item + "_yes"] = isYes ? "✓" : map["q" + item + "_yes"] || "";
          map["q" + item + "_no"] = isNo ? "✓" : map["q" + item + "_no"] || "";
          map["q" + item + "_na"] = isNA ? "✓" : map["q" + item + "_na"] || "";
          map["item" + item + "_yes"] = isYes ? "✓" : "";
          map["item" + item + "_no"] = isNo ? "✓" : "";
          map["item" + item + "_na"] = isNA ? "✓" : "";
        }
      }
    }
  }

  // Handle participant rows for TBT, JST, Training, and Induction
  for (let i = 1; i <= 30; i++) {
    const name =
      fields["participant" + i + "Name"] ||
      fields["tbt_name_" + i] ||
      fields["train_name_" + i] ||
      "";
    const desig =
      fields["participant" + i + "Designation"] ||
      fields["tbt_desig_" + i] ||
      fields["train_desig_" + i] ||
      "";
    const agency =
      fields["participant" + i + "Company"] ||
      fields["participant" + i + "Agency"] ||
      fields["tbt_agency_" + i] ||
      fields["train_company_" + i] ||
      "";
    const sign = name
      ? fields["participant" + i + "Signature"] || "Signed"
      : "";

    map["tbt_name_" + i] = name;
    map["tbt_desig_" + i] = desig;
    map["tbt_agency_" + i] = agency;
    map["tbt_sign_" + i] = sign;

    map["jst_name_" + i] = name;
    map["jst_desig_" + i] = desig;
    map["jst_agency_" + i] = agency;
    map["jst_company_" + i] = agency;
    map["jst_sign_" + i] = sign;

    map["participant" + i + "Name"] = name;
    map["participant" + i + "Designation"] = desig;
    map["participant" + i + "Company"] = agency;
    map["participant" + i + "Agency"] = agency;
    map["participant" + i + "Sign"] = sign;
    map["participant" + i + "Signature"] = sign;

    map["name_" + i] = name;
    map["desig_" + i] = desig;
    map["agency_" + i] = agency;
    map["company_" + i] = agency;
    map["sign_" + i] = sign;

    map["name" + i] = name;
    map["desig" + i] = desig;
    map["agency" + i] = agency;
    map["company" + i] = agency;
    map["sign" + i] = sign;

    map["train_name_" + i] = name;
    map["train_desig_" + i] = desig;
    map["train_company_" + i] = agency;
    map["train_sign_" + i] = sign;

    const indId = fields["attendee" + i + "Id"] || fields["ind_id_" + i] || "";
    const indName =
      fields["attendee" + i + "Name"] || fields["ind_name_" + i] || "";
    const indDesig =
      fields["attendee" + i + "Designation"] || fields["ind_desig_" + i] || "";
    const indSign = indName
      ? fields["attendee" + i + "Signature"] || "Signed"
      : "";

    map["ind_id_" + i] = indId;
    map["ind_name_" + i] = indName;
    map["ind_desig_" + i] = indDesig;
    map["ind_sign_" + i] = indSign;
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
function generatePdfFromDocsTemplate_(
  templateDocId,
  placeholderMap,
  title,
  project,
) {
  if (!templateDocId) {
    throw new Error("No Google Doc template ID provided.");
  }

  let templateFile;
  try {
    templateFile = DriveApp.getFileById(templateDocId);
  } catch (fileErr) {
    throw new Error(
      "Cannot access template file with ID '" +
        templateDocId +
        "'. " +
        "Please ensure the file exists and is shared with 'Anyone with the link can view' or with your Google account. " +
        "Details: " +
        fileErr.message,
    );
  }

  // Check MIME type: DocumentApp can only open native Google Docs (application/vnd.google-apps.document)
  const mimeType = templateFile.getMimeType();
  if (mimeType !== MimeType.GOOGLE_DOCS) {
    throw new Error(
      "The template (ID: " +
        templateDocId +
        ") is of type '" +
        mimeType +
        "', not a native Google Doc. " +
        "If this is an uploaded Word .docx or PDF, open it in Google Drive and select File > 'Save as Google Docs', then copy the new document's ID.",
    );
  }

  // Create working copy in designated folder
  const folder = getEhsGeneratedPdfFolder_(project);
  const copyTitle = (title || "EHS_Document") + "_" + new Date().getTime();
  const docCopy = folder
    ? templateFile.makeCopy(copyTitle, folder)
    : templateFile.makeCopy(copyTitle);

  // Retry opening the copied document with backoff to handle Google Drive indexing propagation
  let doc = null;
  let lastErr = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      if (attempt > 1) {
        Utilities.sleep(600);
      }
      doc = DocumentApp.openById(docCopy.getId());
      if (doc) break;
    } catch (e) {
      lastErr = e;
    }
  }

  if (!doc) {
    try {
      docCopy.setTrashed(true);
    } catch (t) {}
    throw new Error(
      "The document is inaccessible. Please ensure the template file is shared with 'Anyone with the link can view' and is a native Google Doc. " +
        "Details: " +
        (lastErr ? lastErr.message : "openById failed"),
    );
  }

  try {
    // Replace placeholders in Document Body, Header, Footer, and all Table cells
    replacePlaceholdersInDoc_(doc, placeholderMap);
    doc.saveAndClose();

    // Convert populated doc copy to PDF
    const pdfBlob = docCopy
      .getAs("application/pdf")
      .setName((title || "EHS_Document") + ".pdf");
    const pdfFile = folder
      ? folder.createFile(pdfBlob)
      : DriveApp.createFile(pdfBlob);

    try {
      pdfFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (err) {
      Logger.log("Notice: PDF sharing setting: " + err);
    }

    return {
      ok: true,
      fileId: pdfFile.getId(),
      pdfUrl: pdfFile.getUrl(),
      downloadUrl:
        "https://drive.google.com/uc?export=download&id=" + pdfFile.getId(),
      previewUrl:
        "https://drive.google.com/file/d/" + pdfFile.getId() + "/preview",
    };
  } finally {
    // Remove the temporary doc copy to keep Google Drive uncluttered
    try {
      docCopy.setTrashed(true);
    } catch (cleanErr) {
      Logger.log("Notice: Cleaning up temporary doc copy: " + cleanErr);
    }
  }
}

/**
 * Replaces {{placeholder}} tokens across doc body, header, footer, and tables.
 */
function replacePlaceholdersInDoc_(doc, placeholderMap) {
  const body = doc.getBody();
  if (!body) return;

  const header = doc.getHeader();
  const footer = doc.getFooter();

  // Combine text to quickly filter only placeholders present in this document
  let fullDocText = body.getText() || "";
  if (header) fullDocText += " " + (header.getText() || "");
  if (footer) fullDocText += " " + (footer.getText() || "");
  const fullDocTextLower = fullDocText.toLowerCase();

  // Only perform replaceText for keys that actually appear in the document
  const activeKeys = [];
  for (const key in placeholderMap) {
    if (fullDocTextLower.indexOf(key.toLowerCase()) !== -1) {
      activeKeys.push(key);
    }
  }

  const replaceInContainer = (container) => {
    if (!container) return;
    for (let i = 0; i < activeKeys.length; i++) {
      const key = activeKeys[i];
      const val =
        placeholderMap[key] != null ? String(placeholderMap[key]) : "";
      const regexPattern = "\\{\\{\\s*" + escapeRegex_(key) + "\\s*\\}\\}";
      try {
        container.replaceText(regexPattern, val);
      } catch (e) {}
    }
  };

  // Single-pass replacement: Body.replaceText natively updates all paragraphs, tables, and cells!
  replaceInContainer(body);
  if (header) replaceInContainer(header);
  if (footer) replaceInContainer(footer);
}

function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function todayDisplay_() {
  if (typeof displayDate_ === "function" && typeof todayIso_ === "function") {
    return displayDate_(todayIso_());
  }
  const d = new Date();
  const pad = function (n) {
    return n < 10 ? "0" + n : "" + n;
  };
  return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear();
}

/**
 * Ensures a dedicated folder for generated PDFs exists in Google Drive.
 */
function getEhsGeneratedPdfFolder_(project) {
  try {
    if (project && project.id) {
      return getNamedSubfolder_(project, "Generated PDFs");
    }
  } catch (e) {}

  const folderName = "SESIPL EHS Generated PDFs";
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
  const sub = findOne_(SHEETS.SUBMISSIONS, "id", submissionId);
  if (!sub) throw new Error("Submission not found: " + submissionId);

  const project = findOne_(SHEETS.PROJECTS, "id", sub.projectId) || {
    id: sub.projectId,
    name: "SESIPL Project Site",
    code: "PRJ",
  };
  const def = FORM_DEFS.find((f) => f.formCode === sub.formCode) || {
    formCode: sub.formCode,
    title: sub.formCode,
  };
  const fields =
    typeof sub.payload === "object" && sub.payload
      ? sub.payload
      : JSON.parse(sub.payloadJson || "{}");
  const user = { employeeId: sub.submittedBy, name: sub.submittedBy };

  const templateDocId = getDocTemplateId_(sub.formCode);

  let docErrorMsg = "";
  // 1. If a Google Docs template ID is configured for this form:
  if (templateDocId && templateDocId.trim() !== "") {
    try {
      const placeholderMap = buildPlaceholderMap_(
        sub.formCode,
        fields,
        project,
        user,
        submissionId,
      );
      const title =
        (project.code || "PRJ") + "_" + sub.formCode + "_" + submissionId;
      const docResult = generatePdfFromDocsTemplate_(
        templateDocId.trim(),
        placeholderMap,
        title,
        project,
      );

      // Save pdfFileId to submission record for fast subsequent access
      try {
        sub.pdfFileId = docResult.fileId;
        updateOne_(SHEETS.SUBMISSIONS, "id", sub.id, {
          pdfFileId: docResult.fileId,
        });
      } catch (dbErr) {
        Logger.log("Notice: Updating submission with pdfFileId: " + dbErr);
      }

      return {
        ok: true,
        source: "GOOGLE_DOCS",
        title: def.title,
        submission: sub,
        fileId: docResult.fileId,
        pdfUrl: docResult.pdfUrl,
        previewUrl: docResult.previewUrl,
        downloadUrl: docResult.downloadUrl,
      };
    } catch (docErr) {
      docErrorMsg = docErr && docErr.message ? docErr.message : String(docErr);
      Logger.log(
        "Notice: Error generating PDF from Docs template (" +
          sub.formCode +
          "): " +
          docErrorMsg +
          ". Falling back to HTML generator.",
      );
    }
  }

  // 2. Fallback: High-fidelity HTML PDF generation
  const isPermit = sub.formCode && sub.formCode.indexOf("WP_") === 0;
  const html = isPermit
    ? buildWorkPermitPdfHtml_(
        project,
        def,
        fields,
        user,
        sub.version || 1,
        options,
      )
    : buildFormPdfHtml_(project, def, fields, user, sub.version || 1);

  return {
    ok: true,
    source: "HTML_FALLBACK",
    title: def.title,
    submission: sub,
    html: html,
    docError: docErrorMsg,
    templateDocId: templateDocId,
  };
}
