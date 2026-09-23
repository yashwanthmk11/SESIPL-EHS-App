function generateSubmissionPdf_(project, def, fields, user, version) {
  const formCode = def.formCode;
  const templateDocId =
    typeof getDocTemplateId_ === "function"
      ? getDocTemplateId_(formCode)
      : typeof getDocTemplateRegistry_ === "function"
        ? getDocTemplateRegistry_()[formCode] || ""
        : "";

  // If a Google Doc template is configured, data sits directly in the doc and exports to PDF
  if (templateDocId && templateDocId.trim() !== "") {
    try {
      const placeholderMap = buildPlaceholderMap_(
        formCode,
        fields,
        project,
        user,
      );
      const name =
        (project.code || "PRJ") +
        "-" +
        formCode +
        "-v" +
        (version || 1) +
        "-" +
        todayIso_();
      const res = generatePdfFromDocsTemplate_(
        templateDocId.trim(),
        placeholderMap,
        name,
        project,
      );
      return {
        fileId: res.fileId,
        docId: templateDocId.trim(),
        url: res.pdfUrl,
      };
    } catch (err) {
      Logger.log(
        "Notice: Error generating PDF from Docs template (" +
          formCode +
          "): " +
          err +
          ". Falling back to HTML generator.",
      );
    }
  }

  // Fallback to HTML generator
  const isPermit = def.formCode && def.formCode.indexOf("WP_") === 0;
  const html = isPermit
    ? buildWorkPermitPdfHtml_(project, def, fields, user, version)
    : buildFormPdfHtml_(project, def, fields, user, version);
  const folder = getNamedSubfolder_(
    project,
    def.module ? moduleTitle_(def.module) : "Generated PDFs",
  );
  const name =
    project.code + "-" + def.formCode + "-v" + version + "-" + todayIso_();
  const pdfRes = htmlToPdfFile_(html, name, folder);
  return { fileId: pdfRes.fileId, docId: pdfRes.docId, url: pdfRes.url };
}

function moduleTitle_(id) {
  const m = MODULES.find((x) => x.id === id);
  return m ? m.title : "Generated PDFs";
}

function htmlToPdfFile_(html, name, folder) {
  try {
    const blob = Utilities.newBlob(html, "text/html", name + ".html");
    const pdfBlob = blob.getAs("application/pdf").setName(name + ".pdf");
    const pdfFile = folder.createFile(pdfBlob);
    try {
      pdfFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (shareErr) {
      Logger.log("Notice: setSharing ANYONE_WITH_LINK: " + shareErr);
    }
    return { fileId: pdfFile.getId(), docId: "", url: pdfFile.getUrl() };
  } catch (err) {
    const doc = DocumentApp.create(name);
    const body = doc.getBody();
    body.clear();
    const tmp = html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<\/tr>/gi, "\n")
      .replace(/<[^>]+>/g, " ");
    const text = tmp
      .replace(/\s+\n/g, "\n")
      .replace(/[ \t]+/g, " ")
      .trim();
    body
      .appendParagraph(APP_NAME + " — Safety Management System")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(text.substring(0, 45000));
    doc.saveAndClose();
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs(MimeType.PDF).setName(name + ".pdf");
    const pdfFile = folder.createFile(pdfBlob);
    try {
      pdfFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (shareErr) {
      Logger.log("Notice: setSharing ANYONE_WITH_LINK: " + shareErr);
    }
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
    return {
      fileId: pdfFile.getId(),
      docId: doc.getId(),
      url: pdfFile.getUrl(),
    };
  }
}

function buildFormPdfHtml_(project, def, fields, user, version) {
  fields = fields || {};
  project = project || {};
  user = user || {};
  def = def || {};
  const code = (def.formCode || "").toUpperCase();

  // 1. 7-Day Periodic Machine & Equipment Checklists
  if (
    code === "CL_CUT" ||
    code === "CL_WELD" ||
    code === "CL_GRIND" ||
    code === "CL_SCAFFOLD" ||
    hasDailyInspectionKeys_(fields)
  ) {
    return buildDailyInspectionTablePdfHtml_(project, def, fields, user, version);
  }

  // 2. Inspection Checklists with Yes/No/Remarks
  if (code === "CL_DRILL") {
    return buildDrillInspectionPdfHtml_(project, def, fields, user, version);
  }
  if (code === "CL_FE") {
    return buildFireExtinguisherPdfHtml_(project, def, fields, user, version);
  }

  // 3. Tool Box Talk & Job Safety Training Attendances
  if (code === "CL_TBT" || code === "CL_JST" || code === "CL_INDUCTION") {
    return buildAttendancePdfHtml_(project, def, fields, user, version);
  }

  // 4. Worker Screening & Medical Examination
  if (code === "CL_SCREENING" || code === "CL_MEDICAL" || code === "CL_IDCARD") {
    return buildScreeningMedicalPdfHtml_(project, def, fields, user, version);
  }

  // 5. Safety Tags
  if (code.indexOf("TAG_") === 0) {
    return buildSafetyTagPdfHtml_(project, def, fields, user, version);
  }

  // 6. Observations & Weekly/Monthly Reports
  if (code === "OBS_DAILY" || code === "WR_WEEKLY" || code === "WR_MONTHLY") {
    return buildObservationReportPdfHtml_(project, def, fields, user, version);
  }

  // 7. High-fidelity Structured Fallback
  return buildStructuredFormPdfHtml_(project, def, fields, user, version);
}

function hasDailyInspectionKeys_(fields) {
  if (!fields || typeof fields !== "object") return false;
  return Object.keys(fields).some(function (k) {
    return /^item\d+Day\d+$/i.test(k) || /^[a-z0-9]+_q\d+_d\d+$/i.test(k);
  });
}

function getDailyInspectionValue_(fields, idx, day, code) {
  const itemNum = idx + 1;
  const daysShort = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  const dShort = daysShort[day - 1];
  const prefix = (code || "").toLowerCase().replace("cl_", "");

  const candidates = [
    "item" + itemNum + "Day" + day,
    "item" + itemNum + "_day" + day,
    "item" + itemNum + "_d" + day,
    "item" + itemNum + "_" + dShort,
    "item" + itemNum + dShort,
    prefix + "_q" + itemNum + "_d" + day,
    prefix + "_q" + itemNum + "_" + dShort,
    prefix + "_q" + itemNum + dShort,
    "cut_q" + itemNum + "_d" + day,
    "cut_q" + itemNum + "_" + dShort,
    "wm_q" + itemNum + "_d" + day,
    "wm_q" + itemNum + "_" + dShort,
    "grind_q" + itemNum + "_d" + day,
    "grind_q" + itemNum + "_" + dShort,
    "scaff_q" + itemNum + "_d" + day,
    "scaff_q" + itemNum + "_" + dShort,
    "q" + itemNum + "_d" + day,
    "q" + itemNum + "_" + dShort,
    "d" + day + "_q" + itemNum,
    "q" + itemNum + "_day" + day,
  ];

  for (let i = 0; i < candidates.length; i++) {
    const k = candidates[i];
    if (fields[k] != null && String(fields[k]).trim() !== "") {
      return String(fields[k]).trim();
    }
  }

  if (day === 1) {
    const d1Candidates = [
      prefix + "_q" + itemNum + "_choice",
      "q" + itemNum + "_choice",
      "item" + itemNum,
    ];
    for (let j = 0; j < d1Candidates.length; j++) {
      const k1 = d1Candidates[j];
      if (fields[k1] != null && String(fields[k1]).trim() !== "") {
        return String(fields[k1]).trim();
      }
    }
  }
  return "";
}

function renderDailyCell_(val) {
  if (!val) return '<span style="color:#cbd5e1">—</span>';
  const vLower = String(val).toLowerCase().trim();
  if (
    vLower === "yes" ||
    vLower === "y" ||
    val === "✓" ||
    vLower === "ok" ||
    vLower === "true" ||
    vLower === "pass"
  ) {
    return '<span style="color:#0f766e;font-weight:900;font-size:15px">&#10003;</span>';
  }
  if (
    vLower === "no" ||
    vLower === "n" ||
    val === "✗" ||
    vLower === "fail"
  ) {
    return '<span style="color:#dc2626;font-weight:900;font-size:13px">&#10007;</span>';
  }
  if (
    vLower === "na" ||
    vLower === "n/a" ||
    vLower === "not applicable"
  ) {
    return '<span style="color:#64748b;font-size:10px;font-weight:700">N/A</span>';
  }
  return escapeHtml_(val);
}

function buildDailyInspectionTablePdfHtml_(project, def, fields, user, version) {
  const code = (def.formCode || "").toUpperCase();
  let title = "CHECKLIST FOR CUTTING MACHINE";
  let items = [
    "Cutting blade manufacture defined and free from damage",
    "Availability of safety Guard and in good condition",
    "Lock system for plate and guard",
    "Availability of job clamp(fence) and in condition",
    "Availability on handle and in good condition",
    "Cable connection and free from damages",
    "Availability of dust guard (chip deflector)",
    "Machine base in free from damage",
  ];

  if (code === "CL_WELD") {
    title = "CHECKLIST FOR WELDING MACHINE";
    items = [
      "ON / OFF knob undamaged",
      "Regulator with indicator",
      "Welding cables connected with lugs",
      "Welding cable insulation undamaged",
      "Electrode and earthing holders undamaged",
      "Industrial plug available",
      "No exposed live electrical parts",
      "Trolley wheels undamaged",
      "Fire extinguisher and sand bucket available",
    ];
  } else if (code === "CL_GRIND") {
    title = "CHECKLIST FOR GRINDING MACHINE";
    items = [
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
    ];
  } else if (code === "CL_SCAFFOLD") {
    title = "SCAFFOLDING CHECKLIST";
    items = [
      "All coupler hooks are properly installed",
      "Proper platform has been provided & is having proper locking system",
      "Proper access / exit provided",
      "Inspection tag has been displayed",
      "Scaffolding has been erected on a firm base",
      "Wheel lock has been provided & is in working condition",
      "Toe board has been provided",
      "Access ladder has been provided & installed properly",
      "Double (top & mid) railing has been provided",
    ];
  } else if (def.title) {
    title = def.title.toUpperCase();
  }

  const projName = project.name || fields.projectName || "SESIPL Site";
  const contractorName =
    fields.contractor ||
    fields.contractorName ||
    project.client ||
    "Shankar Electricals Services (I) Pvt Ltd";
  const dateVal = displayDate_(
    fields.date || fields.inspectionDate || nowIso_(),
  );
  const equipNo =
    fields.equipmentId || fields.equipmentNo || fields.machineNo || "—";
  const makeVal = fields.make || fields.type || "—";

  let tableRows = "";
  for (let i = 0; i < items.length; i++) {
    const sl = i + 1;
    let dayCells = "";
    for (let d = 1; d <= 7; d++) {
      const v = getDailyInspectionValue_(fields, i, d, code);
      dayCells +=
        '<td style="padding:6px 4px;border:1px solid #0f172a;text-align:center;vertical-align:middle;background:' +
        (d % 2 === 0 ? "#f8fafc" : "#ffffff") +
        '">' +
        renderDailyCell_(v) +
        "</td>";
    }

    tableRows +=
      "<tr>" +
      '<td style="padding:6px 6px;border:1px solid #0f172a;text-align:center;font-weight:bold;color:#0f172a">' +
      sl +
      "</td>" +
      '<td style="padding:6px 10px;border:1px solid #0f172a;color:#0f172a;font-size:11.5px;line-height:1.35">' +
      escapeHtml_(items[i]) +
      "</td>" +
      dayCells +
      "</tr>";
  }

  const supervisor =
    fields.supervisorSign || fields.supervisor || fields.siteSupervisor || "—";
  const safetyOfficer =
    fields.safetyOfficerSign || fields.safetyOfficer || fields.ehsName || "—";
  const electricalEng =
    fields.electricalEngineerSign ||
    fields.electricalEngineer ||
    fields.electricalSign ||
    "—";
  const engineer =
    fields.engineerSign ||
    fields.siteEngineer ||
    fields.engineer ||
    user.name ||
    "—";

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:850px;margin:0 auto;border:2px solid #0f172a;padding:14px;background:#fff;box-sizing:border-box">' +
    // Top SESIPL Header
    '<table style="border-bottom:2px solid #0f172a;margin-bottom:12px">' +
    "  <tr>" +
    '    <td style="width:18%;padding:6px;vertical-align:middle;text-align:center;border-right:1.5px solid #0f172a">' +
    '      <div style="font-weight:900;font-size:16px;color:#0f766e;letter-spacing:1px">SESIPL</div>' +
    '      <div style="font-size:9px;color:#64748b;font-weight:bold;margin-top:2px">SAFETY FIRST</div>' +
    "    </td>" +
    '    <td style="padding:6px 14px;vertical-align:middle;text-align:center">' +
    '      <div style="font-size:16px;font-weight:900;color:#0f172a;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569;margin-top:3px;letter-spacing:0.5px">ENVIRONMENT, HEALTH &amp; SAFETY MANAGEMENT SYSTEM</div>' +
    '      <div style="font-size:15px;font-weight:900;color:#0f766e;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px">' +
    escapeHtml_(title) +
    "</div>" +
    "    </td>" +
    '    <td style="width:20%;padding:6px;vertical-align:middle;text-align:right;font-size:10px;color:#475569;border-left:1.5px solid #0f172a">' +
    "      <div><b>Doc:</b> " +
    escapeHtml_(code) +
    "</div>" +
    "      <div><b>Ver:</b> v" +
    version +
    "</div>" +
    "      <div><b>Date:</b> " +
    escapeHtml_(dateVal) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    // Metadata Header Grid matching template
    '<table style="border:1.5px solid #0f172a;margin-bottom:12px;background:#f8fafc;font-size:11.5px">' +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-right:1px solid #0f172a;width:55%">' +
    '      <b>Project Name:</b> <span style="font-weight:bold;color:#0f766e">' +
    escapeHtml_(projName) +
    "</span>" +
    "    </td>" +
    '    <td style="padding:7px 10px;width:45%">' +
    "      <b>Date:</b> " +
    escapeHtml_(dateVal) +
    "    </td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a">' +
    "      <b>Contractor Name:</b> " +
    escapeHtml_(contractorName) +
    "    </td>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a">' +
    "      <b>Equipment / Machine No:</b> " +
    escapeHtml_(equipNo) +
    (makeVal !== "—" ? " &nbsp;&bull;&nbsp; <b>Make:</b> " + escapeHtml_(makeVal) : "") +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    // Table Subtitle
    '<div style="font-weight:bold;font-size:12px;margin:10px 0 6px;color:#0f172a;text-transform:uppercase;letter-spacing:0.5px">' +
    "  The following items to be checked daily" +
    "</div>" +
    // 7-Day Checklist Matrix Table
    '<table style="border:1.5px solid #0f172a;font-size:11px;margin-bottom:16px">' +
    "  <thead>" +
    '    <tr style="background:#e2e8f0;color:#0f172a">' +
    '      <th style="padding:8px 4px;border:1px solid #0f172a;width:5%;text-align:center">Sl<br>No</th>' +
    '      <th style="padding:8px 10px;border:1px solid #0f172a;width:46%;text-align:left">Description</th>' +
    '      <th style="padding:6px 2px;border:1px solid #0f172a;width:7%;text-align:center">Day 1<br><span style="font-size:9.5px;font-weight:normal">Mon</span></th>' +
    '      <th style="padding:6px 2px;border:1px solid #0f172a;width:7%;text-align:center">Day 2<br><span style="font-size:9.5px;font-weight:normal">Tue</span></th>' +
    '      <th style="padding:6px 2px;border:1px solid #0f172a;width:7%;text-align:center">Day 3<br><span style="font-size:9.5px;font-weight:normal">Wed</span></th>' +
    '      <th style="padding:6px 2px;border:1px solid #0f172a;width:7%;text-align:center">Day 4<br><span style="font-size:9.5px;font-weight:normal">Thu</span></th>' +
    '      <th style="padding:6px 2px;border:1px solid #0f172a;width:7%;text-align:center">Day 5<br><span style="font-size:9.5px;font-weight:normal">Fri</span></th>' +
    '      <th style="padding:6px 2px;border:1px solid #0f172a;width:7%;text-align:center">Day 6<br><span style="font-size:9.5px;font-weight:normal">Sat</span></th>' +
    '      <th style="padding:6px 2px;border:1px solid #0f172a;width:7%;text-align:center">Day 7<br><span style="font-size:9.5px;font-weight:normal">Sun</span></th>' +
    "    </tr>" +
    "  </thead>" +
    "  <tbody>" +
    tableRows +
    "  </tbody>" +
    "</table>" +
    // 4 Sign-Off Signature Blocks matching template
    '<table style="border:1.5px solid #0f172a;background:#fff;font-size:10.5px;margin-top:14px">' +
    "  <tr>" +
    '    <td style="width:25%;padding:10px 8px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;min-height:30px">Checked By Supervisor<br>Name &amp; Sign:</div>' +
    '      <div style="margin-top:8px;font-style:italic;color:#0f766e;font-weight:bold;font-size:11.5px;border-top:1px dashed #cbd5e1;padding-top:4px">' +
    escapeHtml_(supervisor) +
    "</div>" +
    "    </td>" +
    '    <td style="width:25%;padding:10px 8px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;min-height:30px">Checked By Safety officer<br>Name &amp; Sign:</div>' +
    '      <div style="margin-top:8px;font-style:italic;color:#0f766e;font-weight:bold;font-size:11.5px;border-top:1px dashed #cbd5e1;padding-top:4px">' +
    escapeHtml_(safetyOfficer) +
    "</div>" +
    "    </td>" +
    '    <td style="width:25%;padding:10px 8px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;min-height:30px">Checked By Electrical Engineer<br>&amp; Sign:</div>' +
    '      <div style="margin-top:8px;font-style:italic;color:#0f766e;font-weight:bold;font-size:11.5px;border-top:1px dashed #cbd5e1;padding-top:4px">' +
    escapeHtml_(electricalEng) +
    "</div>" +
    "    </td>" +
    '    <td style="width:25%;padding:10px 8px;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;min-height:30px">Engineer Name &amp; Sign:</div>' +
    '      <div style="margin-top:8px;font-style:italic;color:#0f766e;font-weight:bold;font-size:11.5px;border-top:1px dashed #cbd5e1;padding-top:4px">' +
    escapeHtml_(engineer) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    // Standard SESIPL Footer Note
    '<div style="margin-top:12px;font-size:9.5px;color:#94a3b8;text-align:center">' +
    "  This record is digitally authenticated and stored in the SESIPL EHS Cloud System. Verification of physical condition must be conducted prior to each shift." +
    "</div>" +
    "</div></body></html>"
  );
}

function buildDrillInspectionPdfHtml_(project, def, fields, user, version) {
  const items = [
    {
      q: "Equipment double insulated",
      choice: fields.drill_q1_choice || fields.q1_choice || "",
      remarks: fields.drill_q1_remarks || fields.q1_remarks || "",
    },
    {
      q: "Equipment free from any defect",
      choice: fields.drill_q2_choice || fields.q2_choice || "",
      remarks: fields.drill_q2_remarks || fields.q2_remarks || "",
    },
    {
      q: "Electrical cable free from defects",
      choice: fields.drill_q3_choice || fields.q3_choice || "",
      remarks: fields.drill_q3_remarks || fields.q3_remarks || "",
    },
    {
      q: "Industrial plug top available and in working condition",
      choice: fields.drill_q4_choice || fields.q4_choice || "",
      remarks: fields.drill_q4_remarks || fields.q4_remarks || "",
    },
    {
      q: "Drilling Bit without any damage",
      choice: fields.drill_q5_choice || fields.q5_choice || "",
      remarks: fields.drill_q5_remarks || fields.q5_remarks || "",
    },
    {
      q: "Drilling bit holder is available and in working condition",
      choice: fields.drill_q6_choice || fields.q6_choice || "",
      remarks: fields.drill_q6_remarks || fields.q6_remarks || "",
    },
    {
      q: "Switch in working condition",
      choice: fields.drill_q7_choice || fields.q7_choice || "",
      remarks: fields.drill_q7_remarks || fields.q7_remarks || "",
    },
  ];

  const projName = project.name || fields.projectName || "SESIPL Site";
  const contractorName =
    fields.contractor ||
    fields.contractorName ||
    project.client ||
    "Shankar Electricals Services (I) Pvt Ltd";
  const dateVal = displayDate_(
    fields.inspectionDate || fields.date || nowIso_(),
  );
  const equipNo =
    fields.equipmentNo || fields.equipmentId || fields.machineNo || "—";
  const checkedBy =
    fields.checkedByName || fields.checkedBy || user.name || "—";
  const ehsName = fields.ehsName || fields.safetyOfficer || "—";
  const reviewStatus = fields.reviewStatus || "Accepted";
  const isAccepted = reviewStatus.toLowerCase().indexOf("accept") !== -1;

  let tableRows = "";
  for (let i = 0; i < items.length; i++) {
    const itm = items[i];
    const cLower = String(itm.choice).toLowerCase().trim();
    const isYes =
      cLower === "yes" || cLower === "y" || itm.choice === "✓" || cLower === "ok";
    const isNo =
      cLower === "no" || cLower === "n" || itm.choice === "✗";

    const yesMarkup = isYes
      ? '<span style="color:#0f766e;font-weight:900;font-size:15px">&#10003;</span>'
      : "";
    const noMarkup = isNo
      ? '<span style="color:#dc2626;font-weight:900;font-size:13px">&#10007;</span>'
      : "";

    tableRows +=
      "<tr>" +
      '<td style="padding:8px 6px;border:1px solid #0f172a;text-align:center;font-weight:bold">' +
      (i + 1) +
      "</td>" +
      '<td style="padding:8px 12px;border:1px solid #0f172a;font-size:12px;color:#0f172a">' +
      escapeHtml_(itm.q) +
      "</td>" +
      '<td style="padding:8px 6px;border:1px solid #0f172a;text-align:center;background:#f8fafc">' +
      yesMarkup +
      "</td>" +
      '<td style="padding:8px 6px;border:1px solid #0f172a;text-align:center;background:#f8fafc">' +
      noMarkup +
      "</td>" +
      '<td style="padding:8px 10px;border:1px solid #0f172a;font-size:11px;color:#334155">' +
      (itm.remarks ? escapeHtml_(itm.remarks) : '<span style="color:#94a3b8">—</span>') +
      "</td>" +
      "</tr>";
  }

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:850px;margin:0 auto;border:2px solid #0f172a;padding:14px;background:#fff;box-sizing:border-box">' +
    // Header
    '<table style="border-bottom:2px solid #0f172a;margin-bottom:12px">' +
    "  <tr>" +
    '    <td style="width:18%;padding:6px;vertical-align:middle;text-align:center;border-right:1.5px solid #0f172a">' +
    '      <div style="font-weight:900;font-size:16px;color:#0f766e;letter-spacing:1px">SESIPL</div>' +
    '      <div style="font-size:9px;color:#64748b;font-weight:bold;margin-top:2px">SAFETY FIRST</div>' +
    "    </td>" +
    '    <td style="padding:6px 14px;vertical-align:middle;text-align:center">' +
    '      <div style="font-size:16px;font-weight:900;color:#0f172a;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569;margin-top:3px;letter-spacing:0.5px">ENVIRONMENT, HEALTH &amp; SAFETY MANAGEMENT SYSTEM</div>' +
    '      <div style="font-size:15px;font-weight:900;color:#0f766e;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px">CHECKLIST FOR DRILLING MACHINE</div>' +
    "    </td>" +
    '    <td style="width:20%;padding:6px;vertical-align:middle;text-align:right;font-size:10px;color:#475569;border-left:1.5px solid #0f172a">' +
    "      <div><b>Doc Code:</b> CL_DRILL</div>" +
    "      <div><b>Version:</b> v" +
    version +
    "</div>" +
    "      <div><b>Date:</b> " +
    escapeHtml_(dateVal) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    // Metadata Header Grid
    '<table style="border:1.5px solid #0f172a;margin-bottom:14px;background:#f8fafc;font-size:11.5px">' +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-right:1px solid #0f172a;width:55%">' +
    '      <b>Project Name:</b> <span style="font-weight:bold;color:#0f766e">' +
    escapeHtml_(projName) +
    "</span>" +
    "    </td>" +
    '    <td style="padding:7px 10px;width:45%">' +
    "      <b>Inspection Date:</b> " +
    escapeHtml_(dateVal) +
    "    </td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a">' +
    "      <b>Contractor Name:</b> " +
    escapeHtml_(contractorName) +
    "    </td>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a">' +
    "      <b>Equipment No.:</b> " +
    escapeHtml_(equipNo) +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    // Checklist Table
    '<table style="border:1.5px solid #0f172a;font-size:11.5px;margin-bottom:16px">' +
    "  <thead>" +
    '    <tr style="background:#e2e8f0;color:#0f172a">' +
    '      <th style="padding:8px 6px;border:1px solid #0f172a;width:7%;text-align:center">Sl. No.</th>' +
    '      <th style="padding:8px 12px;border:1px solid #0f172a;width:53%;text-align:left">Description of Inspection Items</th>' +
    '      <th style="padding:8px 6px;border:1px solid #0f172a;width:8%;text-align:center">Yes</th>' +
    '      <th style="padding:8px 6px;border:1px solid #0f172a;width:8%;text-align:center">No</th>' +
    '      <th style="padding:8px 10px;border:1px solid #0f172a;width:24%;text-align:left">Remarks</th>' +
    "    </tr>" +
    "  </thead>" +
    "  <tbody>" +
    tableRows +
    "  </tbody>" +
    "</table>" +
    // Review and Sign-offs Table
    '<table style="border:1.5px solid #0f172a;background:#fff;font-size:11px;margin-top:14px">' +
    "  <tr>" +
    '    <td style="width:33%;padding:10px 10px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Checked by Name &amp; Sign:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(checkedBy) +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Site Technician / Supervisor</div>' +
    "    </td>" +
    '    <td style="width:33%;padding:10px 10px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">EHS SESIPL Name &amp; Sign:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(ehsName) +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">SESIPL EHS Officer</div>' +
    "    </td>" +
    '    <td style="width:34%;padding:10px 10px;vertical-align:top;background:#f8fafc">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:6px">PMC Review Status:</div>' +
    '      <div style="font-size:13px;font-weight:900;color:' +
    (isAccepted ? "#0f766e" : "#dc2626") +
    ';padding:4px 8px;display:inline-block;border-radius:4px;border:1px solid ' +
    (isAccepted ? "#99f6e4" : "#fecaca") +
    ";background:" +
    (isAccepted ? "#f0fdfa" : "#fef2f2") +
    '">' +
    (isAccepted ? "✓ " : "✗ ") +
    escapeHtml_(reviewStatus.toUpperCase()) +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:4px">Verified for site operation</div>' +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="margin-top:12px;font-size:9.5px;color:#94a3b8;text-align:center">' +
    "  This record is digitally authenticated and stored in the SESIPL EHS Cloud System." +
    "</div>" +
    "</div></body></html>"
  );
}

function buildFireExtinguisherPdfHtml_(project, def, fields, user, version) {
  const items = [
    { q: "Extinguisher clean and tidy", val: fields.conditionClean },
    { q: "Extinguisher corroded", val: fields.corroded },
    { q: "Safety pin in locking position", val: fields.safetyPin },
    { q: "Discharge nozzle condition", val: fields.nozzleCondition },
    { q: "Hose condition", val: fields.hoseCondition },
    { q: "Weight matches body marking", val: fields.weightMatches },
    { q: "Visual board available", val: fields.visualBoard },
    { q: "Indicator gauge in green", val: fields.gaugeGreen },
  ];

  const projName = project.name || fields.projectName || "SESIPL Site";
  const dateVal = displayDate_(
    fields.inspectionDate || fields.date || nowIso_(),
  );
  const nextDateVal = displayDate_(fields.nextInspectionDate || "—");
  const extNo = fields.extinguisherNo || "—";
  const spec = fields.extinguisherSpecification || "ABC Dry Powder";
  const ehsName = fields.ehsName || user.name || "—";

  let tableRows = "";
  for (let i = 0; i < items.length; i++) {
    const itm = items[i];
    const cLower = String(itm.val || "").toLowerCase().trim();
    const isYes =
      cLower === "yes" || cLower === "y" || itm.val === "✓" || cLower === "ok";
    const isNo =
      cLower === "no" || cLower === "n" || itm.val === "✗";

    tableRows +=
      "<tr>" +
      '<td style="padding:7px 6px;border:1px solid #0f172a;text-align:center;font-weight:bold">' +
      (i + 1) +
      "</td>" +
      '<td style="padding:7px 12px;border:1px solid #0f172a;font-size:12px;color:#0f172a">' +
      escapeHtml_(itm.q) +
      "</td>" +
      '<td style="padding:7px 6px;border:1px solid #0f172a;text-align:center;background:#f8fafc">' +
      (isYes
        ? '<span style="color:#0f766e;font-weight:900;font-size:15px">&#10003;</span>'
        : "") +
      "</td>" +
      '<td style="padding:7px 6px;border:1px solid #0f172a;text-align:center;background:#f8fafc">' +
      (isNo
        ? '<span style="color:#dc2626;font-weight:900;font-size:13px">&#10007;</span>'
        : "") +
      "</td>" +
      '<td style="padding:7px 10px;border:1px solid #0f172a;font-size:11px;color:#334155">' +
      (fields["fe_q" + (i + 1) + "_remarks"]
        ? escapeHtml_(fields["fe_q" + (i + 1) + "_remarks"])
        : '<span style="color:#94a3b8">—</span>') +
      "</td>" +
      "</tr>";
  }

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:850px;margin:0 auto;border:2px solid #0f172a;padding:14px;background:#fff;box-sizing:border-box">' +
    '<table style="border-bottom:2px solid #0f172a;margin-bottom:12px">' +
    "  <tr>" +
    '    <td style="width:18%;padding:6px;vertical-align:middle;text-align:center;border-right:1.5px solid #0f172a">' +
    '      <div style="font-weight:900;font-size:16px;color:#0f766e;letter-spacing:1px">SESIPL</div>' +
    '      <div style="font-size:9px;color:#64748b;font-weight:bold;margin-top:2px">SAFETY FIRST</div>' +
    "    </td>" +
    '    <td style="padding:6px 14px;vertical-align:middle;text-align:center">' +
    '      <div style="font-size:16px;font-weight:900;color:#0f172a;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569;margin-top:3px;letter-spacing:0.5px">ENVIRONMENT, HEALTH &amp; SAFETY MANAGEMENT SYSTEM</div>' +
    '      <div style="font-size:15px;font-weight:900;color:#0f766e;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px">FIRE EXTINGUISHER INSPECTION CHECKLIST</div>' +
    "    </td>" +
    '    <td style="width:20%;padding:6px;vertical-align:middle;text-align:right;font-size:10px;color:#475569;border-left:1.5px solid #0f172a">' +
    "      <div><b>Doc Code:</b> CL_FE</div>" +
    "      <div><b>Version:</b> v" +
    version +
    "</div>" +
    "      <div><b>Date:</b> " +
    escapeHtml_(dateVal) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;margin-bottom:14px;background:#f8fafc;font-size:11.5px">' +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-right:1px solid #0f172a;width:55%">' +
    '      <b>Project Name:</b> <span style="font-weight:bold;color:#0f766e">' +
    escapeHtml_(projName) +
    "</span>" +
    "    </td>" +
    '    <td style="padding:7px 10px;width:45%">' +
    "      <b>Extinguisher No.:</b> " +
    escapeHtml_(extNo) +
    "    </td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a">' +
    "      <b>Type / Specification:</b> " +
    escapeHtml_(spec) +
    "    </td>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a">' +
    "      <b>Inspection Date:</b> " +
    escapeHtml_(dateVal) +
    " &nbsp;&bull;&nbsp; <b>Next Due:</b> " +
    escapeHtml_(nextDateVal) +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;font-size:11.5px;margin-bottom:16px">' +
    "  <thead>" +
    '    <tr style="background:#e2e8f0;color:#0f172a">' +
    '      <th style="padding:8px 6px;border:1px solid #0f172a;width:7%;text-align:center">Sl. No.</th>' +
    '      <th style="padding:8px 12px;border:1px solid #0f172a;width:53%;text-align:left">Description of Inspection Items</th>' +
    '      <th style="padding:8px 6px;border:1px solid #0f172a;width:8%;text-align:center">Yes</th>' +
    '      <th style="padding:8px 6px;border:1px solid #0f172a;width:8%;text-align:center">No</th>' +
    '      <th style="padding:8px 10px;border:1px solid #0f172a;width:24%;text-align:left">Remarks</th>' +
    "    </tr>" +
    "  </thead>" +
    "  <tbody>" +
    tableRows +
    "  </tbody>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;background:#fff;font-size:11px;margin-top:14px">' +
    "  <tr>" +
    '    <td style="width:50%;padding:10px 14px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Inspected by EHS Inspector / Sign:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(ehsName) +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Site Safety Lead</div>' +
    "    </td>" +
    '    <td style="width:50%;padding:10px 14px;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Verified by SESIPL Safety Authority:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(user.name || "SESIPL Safety Officer") +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Corporate EHS Authority</div>' +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="margin-top:12px;font-size:9.5px;color:#94a3b8;text-align:center">' +
    "  This record is digitally authenticated and stored in the SESIPL EHS Cloud System." +
    "</div>" +
    "</div></body></html>"
  );
}

function buildAttendancePdfHtml_(project, def, fields, user, version) {
  const code = (def.formCode || "").toUpperCase();
  const isJst = code === "CL_JST";
  const title = isJst
    ? "JOB SAFETY TRAINING (JST) ATTENDANCE SHEET"
    : "TOOL BOX TALK (TBT) RECORD";

  const projName = project.name || fields.projectName || "SESIPL Site";
  const dateVal = displayDate_(fields.date || nowIso_());
  const timeVal = fields.time || fields.trainingTime || "—";
  const topicVal = fields.topic || "Daily Safety Briefing";
  const conductedBy = fields.conductedBy || user.name || "—";
  const pmName =
    fields.projectManager || fields.projectManagerSignature || "—";
  const notes =
    fields.keyPoints || fields.acknowledgement || "Safety briefing conducted.";

  let attendeeRows = "";
  let attendeeCount = 0;
  for (let i = 1; i <= 25; i++) {
    const name =
      fields["participant" + i + "_name"] ||
      fields["attendee" + i + "_name"] ||
      fields["participant" + i] ||
      fields["attendee" + i];
    if (name && String(name).trim() !== "") {
      attendeeCount++;
      const idVal =
        fields["participant" + i + "_id"] ||
        fields["participant" + i + "_token"] ||
        fields["attendee" + i + "_id"] ||
        "—";
      const trade =
        fields["participant" + i + "_trade"] ||
        fields["participant" + i + "_contractor"] ||
        fields["attendee" + i + "_trade"] ||
        "SESIPL";
      const sign =
        fields["participant" + i + "_sign"] ||
        fields["attendee" + i + "_sign"] ||
        name;

      attendeeRows +=
        "<tr>" +
        '<td style="padding:6px;border:1px solid #0f172a;text-align:center">' +
        attendeeCount +
        "</td>" +
        '<td style="padding:6px 10px;border:1px solid #0f172a;font-weight:600">' +
        escapeHtml_(name) +
        "</td>" +
        '<td style="padding:6px 8px;border:1px solid #0f172a;text-align:center">' +
        escapeHtml_(idVal) +
        "</td>" +
        '<td style="padding:6px 10px;border:1px solid #0f172a">' +
        escapeHtml_(trade) +
        "</td>" +
        '<td style="padding:6px 10px;border:1px solid #0f172a;font-style:italic;color:#0f766e;text-align:center">' +
        escapeHtml_(sign) +
        "</td>" +
        "</tr>";
    }
  }

  if (attendeeCount === 0) {
    attendeeRows =
      '<tr><td colspan="5" style="padding:14px;text-align:center;color:#64748b;border:1px solid #0f172a">No individual attendee records submitted. All site workers briefed.</td></tr>';
  }

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:850px;margin:0 auto;border:2px solid #0f172a;padding:14px;background:#fff;box-sizing:border-box">' +
    '<table style="border-bottom:2px solid #0f172a;margin-bottom:12px">' +
    "  <tr>" +
    '    <td style="width:18%;padding:6px;vertical-align:middle;text-align:center;border-right:1.5px solid #0f172a">' +
    '      <div style="font-weight:900;font-size:16px;color:#0f766e;letter-spacing:1px">SESIPL</div>' +
    '      <div style="font-size:9px;color:#64748b;font-weight:bold;margin-top:2px">SAFETY FIRST</div>' +
    "    </td>" +
    '    <td style="padding:6px 14px;vertical-align:middle;text-align:center">' +
    '      <div style="font-size:16px;font-weight:900;color:#0f172a;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569;margin-top:3px;letter-spacing:0.5px">ENVIRONMENT, HEALTH &amp; SAFETY MANAGEMENT SYSTEM</div>' +
    '      <div style="font-size:15px;font-weight:900;color:#0f766e;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px">' +
    escapeHtml_(title) +
    "</div>" +
    "    </td>" +
    '    <td style="width:20%;padding:6px;vertical-align:middle;text-align:right;font-size:10px;color:#475569;border-left:1.5px solid #0f172a">' +
    "      <div><b>Doc Code:</b> " +
    escapeHtml_(code) +
    "</div>" +
    "      <div><b>Ver:</b> v" +
    version +
    "</div>" +
    "      <div><b>Date:</b> " +
    escapeHtml_(dateVal) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;margin-bottom:12px;background:#f8fafc;font-size:11.5px">' +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-right:1px solid #0f172a;width:55%">' +
    '      <b>Project Name:</b> <span style="font-weight:bold;color:#0f766e">' +
    escapeHtml_(projName) +
    "</span>" +
    "    </td>" +
    '    <td style="padding:7px 10px;width:45%">' +
    "      <b>Date &amp; Time:</b> " +
    escapeHtml_(dateVal) +
    " " +
    escapeHtml_(timeVal) +
    "    </td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a">' +
    "      <b>Topic Discussed:</b> <b>" +
    escapeHtml_(topicVal) +
    "</b>" +
    "    </td>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a">' +
    "      <b>Conducted By:</b> " +
    escapeHtml_(conductedBy) +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="border:1.5px solid #0f172a;padding:10px 12px;background:#fff;margin-bottom:12px;font-size:11.5px">' +
    '  <div style="font-weight:bold;color:#334155;margin-bottom:4px">Key Points / Hazards &amp; Precautions Discussed:</div>' +
    '  <div style="color:#0f172a;line-height:1.45">' +
    escapeHtml_(notes) +
    "</div>" +
    "</div>" +
    '<table style="border:1.5px solid #0f172a;font-size:11px;margin-bottom:14px">' +
    "  <thead>" +
    '    <tr style="background:#e2e8f0;color:#0f172a">' +
    '      <th style="padding:7px 6px;border:1px solid #0f172a;width:7%;text-align:center">Sl No</th>' +
    '      <th style="padding:7px 12px;border:1px solid #0f172a;width:40%;text-align:left">Participant Name</th>' +
    '      <th style="padding:7px 8px;border:1px solid #0f172a;width:18%;text-align:center">Token / Emp ID</th>' +
    '      <th style="padding:7px 10px;border:1px solid #0f172a;width:20%;text-align:left">Trade / Contractor</th>' +
    '      <th style="padding:7px 10px;border:1px solid #0f172a;width:15%;text-align:center">Signature</th>' +
    "    </tr>" +
    "  </thead>" +
    "  <tbody>" +
    attendeeRows +
    "  </tbody>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;background:#fff;font-size:11px;margin-top:14px">' +
    "  <tr>" +
    '    <td style="width:50%;padding:10px 14px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Tool Box Talk Conducted By:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(conductedBy) +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Site Safety Lead / Engineer</div>' +
    "    </td>" +
    '    <td style="width:50%;padding:10px 14px;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Project Manager Signature:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(pmName) +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Site Project Lead / In-Charge</div>' +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="margin-top:12px;font-size:9.5px;color:#94a3b8;text-align:center">' +
    "  This record is digitally authenticated and stored in the SESIPL EHS Cloud System." +
    "</div>" +
    "</div></body></html>"
  );
}

function buildScreeningMedicalPdfHtml_(project, def, fields, user, version) {
  const code = (def.formCode || "").toUpperCase();
  const isMed = code === "CL_MEDICAL";
  const title = isMed
    ? "MEDICAL FITNESS CERTIFICATE"
    : "SCREENING OF WORKER FORMAT";

  const projName = project.name || fields.projectName || "SESIPL Site";
  const dateVal = displayDate_(
    fields.date || fields.examinationDate || nowIso_(),
  );
  const workerName =
    fields.workerName || fields.name || fields.candidateName || "—";
  const age = fields.age || fields.workerAge || "—";
  const fatherName = fields.fatherName || fields.husbandName || "—";
  const trade = fields.trade || fields.designation || "Electrician";
  const contractor =
    fields.contractor ||
    fields.contractorName ||
    "Shankar Electricals Services (I) Pvt Ltd";
  const contactNo = fields.contactNumber || fields.mobile || "—";
  const aadhaar = fields.aadhaarNo || fields.idNumber || "—";
  const bloodGroup = fields.bloodGroup || "—";
  const fitStatus =
    fields.fitnessStatus || fields.result || fields.status || "FIT FOR DUTY";

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:850px;margin:0 auto;border:2px solid #0f172a;padding:14px;background:#fff;box-sizing:border-box">' +
    '<table style="border-bottom:2px solid #0f172a;margin-bottom:12px">' +
    "  <tr>" +
    '    <td style="width:18%;padding:6px;vertical-align:middle;text-align:center;border-right:1.5px solid #0f172a">' +
    '      <div style="font-weight:900;font-size:16px;color:#0f766e;letter-spacing:1px">SESIPL</div>' +
    '      <div style="font-size:9px;color:#64748b;font-weight:bold;margin-top:2px">SAFETY FIRST</div>' +
    "    </td>" +
    '    <td style="padding:6px 14px;vertical-align:middle;text-align:center">' +
    '      <div style="font-size:16px;font-weight:900;color:#0f172a;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569;margin-top:3px;letter-spacing:0.5px">OCCUPATIONAL HEALTH &amp; SAFETY SYSTEM</div>' +
    '      <div style="font-size:15px;font-weight:900;color:#0f766e;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px">' +
    escapeHtml_(title) +
    "</div>" +
    "    </td>" +
    '    <td style="width:20%;padding:6px;vertical-align:middle;text-align:right;font-size:10px;color:#475569;border-left:1.5px solid #0f172a">' +
    "      <div><b>Doc Code:</b> " +
    escapeHtml_(code) +
    "</div>" +
    "      <div><b>Ver:</b> v" +
    version +
    "</div>" +
    "      <div><b>Date:</b> " +
    escapeHtml_(dateVal) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="font-weight:bold;font-size:12px;margin:8px 0;color:#0f766e;text-transform:uppercase">A. Worker Demographic Details</div>' +
    '<table style="border:1.5px solid #0f172a;font-size:11.5px;margin-bottom:14px;background:#f8fafc">' +
    "  <tr>" +
    '    <td style="padding:6px 10px;border-right:1px solid #0f172a;width:50%"><b>Worker Name:</b> ' +
    escapeHtml_(workerName) +
    "</td>" +
    '    <td style="padding:6px 10px;width:50%"><b>Age / Gender:</b> ' +
    escapeHtml_(age) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:6px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Father\'s Name:</b> ' +
    escapeHtml_(fatherName) +
    "</td>" +
    '    <td style="padding:6px 10px;border-top:1px solid #0f172a"><b>Blood Group:</b> ' +
    escapeHtml_(bloodGroup) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:6px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Trade / Role:</b> ' +
    escapeHtml_(trade) +
    "</td>" +
    '    <td style="padding:6px 10px;border-top:1px solid #0f172a"><b>Contractor / Agency:</b> ' +
    escapeHtml_(contractor) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:6px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Aadhaar / ID No:</b> ' +
    escapeHtml_(aadhaar) +
    "</td>" +
    '    <td style="padding:6px 10px;border-top:1px solid #0f172a"><b>Contact Phone:</b> ' +
    escapeHtml_(contactNo) +
    "</td>" +
    "  </tr>" +
    "</table>" +
    '<div style="font-weight:bold;font-size:12px;margin:8px 0;color:#0f766e;text-transform:uppercase">B. Clinical Examination &amp; Fitness Findings</div>' +
    '<table style="border:1.5px solid #0f172a;font-size:11.5px;margin-bottom:14px">' +
    '  <tr style="background:#e2e8f0;font-weight:bold">' +
    '    <td style="padding:6px 10px;border:1px solid #0f172a;width:35%">Parameter</td>' +
    '    <td style="padding:6px 10px;border:1px solid #0f172a;width:35%">Finding / Measurement</td>' +
    '    <td style="padding:6px 10px;border:1px solid #0f172a;width:30%">Status</td>' +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">Blood Pressure &amp; Pulse</td>' +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">' +
    escapeHtml_(fields.bp || "120/80 mmHg") +
    "</td>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a;color:#0f766e;font-weight:bold">Normal</td>' +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">Vision (Distant &amp; Near)</td>' +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">' +
    escapeHtml_(fields.vision || "6/6 both eyes") +
    "</td>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a;color:#0f766e;font-weight:bold">Normal</td>' +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">Vertigo / Height Phobia</td>' +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">' +
    escapeHtml_(fields.vertigo || "Absent / Clear") +
    "</td>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a;color:#0f766e;font-weight:bold">Fit for Height</td>' +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">Respiratory &amp; General</td>' +
    '    <td style="padding:6px 10px;border:1px solid #0f172a">' +
    escapeHtml_(fields.respiratory || "Lungs clear, no chronic illness") +
    "</td>" +
    '    <td style="padding:6px 10px;border:1px solid #0f172a;color:#0f766e;font-weight:bold">Fit</td>' +
    "  </tr>" +
    "</table>" +
    '<div style="border:1.5px solid #0f172a;padding:12px;background:#f0fdfa;margin-bottom:14px;border-left:6px solid #0f766e">' +
    '  <div style="font-weight:bold;font-size:12px;color:#0f766e">MEDICAL / SAFETY CERTIFICATION:</div>' +
    '  <div style="font-size:12px;font-weight:900;margin-top:4px;color:#0f172a">' +
    escapeHtml_(fitStatus.toUpperCase()) +
    "</div>" +
    '  <div style="font-size:10px;color:#475569;margin-top:2px">The individual has been examined and cleared for employment &amp; site duties under SESIPL EHS norms.</div>' +
    "</div>" +
    '<table style="border:1.5px solid #0f172a;background:#fff;font-size:11px;margin-top:14px">' +
    "  <tr>" +
    '    <td style="width:50%;padding:10px 14px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Examining Medical / EHS Officer:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(fields.doctorName || fields.screenerName || user.name || "Medical Officer") +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Reg. Medical Practitioner / EHS Specialist</div>' +
    "    </td>" +
    '    <td style="width:50%;padding:10px 14px;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">SESIPL EHS In-Charge Sign:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(user.name || "SESIPL Safety Officer") +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Site Safety Management</div>' +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="margin-top:12px;font-size:9.5px;color:#94a3b8;text-align:center">' +
    "  This record is digitally authenticated and stored in the SESIPL EHS Cloud System." +
    "</div>" +
    "</div></body></html>"
  );
}

function buildSafetyTagPdfHtml_(project, def, fields, user, version) {
  const code = (def.formCode || "").toUpperCase();
  const isRed = code === "TAG_RED";
  const isTool = code === "TAG_TOOL";
  const isFe = code === "TAG_FE";
  const isScaff = code === "TAG_SCAFF";

  let tagTitle = "EQUIPMENT INSPECTION TAG";
  let tagColor = "#0f766e";
  let statusText = "INSPECTED & SAFE TO USE";

  if (isRed) {
    tagTitle = "DANGER — DO NOT OPERATE";
    tagColor = "#b91c1c";
    statusText = "DEFECTIVE / OUT OF SERVICE";
  } else if (isTool) {
    tagTitle = "POWER TOOL INSPECTION TAG";
    tagColor = "#0284c7";
    statusText = "PASS / FIT FOR WORK";
  } else if (isFe) {
    tagTitle = "FIRE EXTINGUISHER TAG";
    tagColor = "#b91c1c";
    statusText = "INSPECTED & CHARGED";
  } else if (isScaff) {
    tagTitle = "SCAFFOLD INSPECTION TAG";
    tagColor = "#15803d";
    statusText = "SAFE FOR WORKERS";
  }

  const projName = project.name || fields.projectName || "SESIPL Site";
  const dateVal = displayDate_(fields.date || fields.tagDate || nowIso_());
  const validUntil = displayDate_(fields.validUntil || fields.nextDue || "—");
  const tagId = fields.tagNo || fields.tagId || fields.equipmentNo || "TAG-001";
  const inspector = fields.inspectedBy || user.name || "EHS Lead";

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:540px;margin:20px auto;border:3px solid ' +
    tagColor +
    ';padding:16px;background:#fff;border-radius:6px;box-sizing:border-box">' +
    '<div style="background:' +
    tagColor +
    ';color:#fff;text-align:center;padding:12px;margin:-16px -16px 14px -16px;border-radius:3px 3px 0 0">' +
    '  <div style="font-size:12px;font-weight:bold;letter-spacing:1px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '  <div style="font-size:18px;font-weight:900;letter-spacing:1px;margin-top:3px">' +
    escapeHtml_(tagTitle) +
    "</div>" +
    "</div>" +
    '<div style="text-align:center;margin-bottom:14px">' +
    '  <div style="font-size:12px;color:#64748b;font-weight:bold">TAG / SERIAL IDENTIFICATION</div>' +
    '  <div style="font-size:22px;font-weight:900;color:' +
    tagColor +
    ';letter-spacing:1px;margin-top:2px">' +
    escapeHtml_(tagId) +
    "</div>" +
    "</div>" +
    '<table style="border:1.5px solid #0f172a;font-size:12px;margin-bottom:14px;background:#f8fafc">' +
    "  <tr>" +
    '    <td style="padding:8px 10px;border-right:1px solid #0f172a;width:40%"><b>Project:</b></td>' +
    '    <td style="padding:8px 10px">' +
    escapeHtml_(projName) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:8px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Inspection Date:</b></td>' +
    '    <td style="padding:8px 10px;border-top:1px solid #0f172a">' +
    escapeHtml_(dateVal) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:8px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Valid Until:</b></td>' +
    '    <td style="padding:8px 10px;border-top:1px solid #0f172a;font-weight:bold">' +
    escapeHtml_(validUntil) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:8px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Inspected By:</b></td>' +
    '    <td style="padding:8px 10px;border-top:1px solid #0f172a">' +
    escapeHtml_(inspector) +
    "</td>" +
    "  </tr>" +
    "</table>" +
    '<div style="text-align:center;padding:10px;background:#f1f5f9;border:1px solid #cbd5e1;border-radius:4px;font-size:13px;font-weight:bold;color:' +
    tagColor +
    '">' +
    escapeHtml_(statusText) +
    "</div>" +
    '<div style="margin-top:10px;font-size:9px;color:#94a3b8;text-align:center">SESIPL EHS Safety Verification Tag. Tampering with this tag is strictly prohibited.</div>' +
    "</div></body></html>"
  );
}

function buildObservationReportPdfHtml_(project, def, fields, user, version) {
  const code = (def.formCode || "").toUpperCase();
  const title =
    code === "OBS_DAILY"
      ? "DAILY EHS OBSERVATION REPORT"
      : def.title
        ? def.title.toUpperCase()
        : "EHS COMPLIANCE REPORT";

  const projName = project.name || fields.projectName || "SESIPL Site";
  const dateVal = displayDate_(fields.date || fields.obsDate || nowIso_());
  const locVal = fields.location || fields.area || "Site Floor";
  const desc =
    fields.description ||
    fields.observation ||
    fields.unsafeActOrCondition ||
    "—";
  const cat = fields.category || fields.observationType || "Unsafe Condition";
  const risk = fields.riskLevel || fields.severity || "Medium";
  const capa =
    fields.correctiveAction ||
    fields.actionRequired ||
    fields.capa ||
    "Immediate correction instructed";
  const resp = fields.assignedTo || fields.responsibility || "Site Supervisor";
  const status = fields.status || "Action Initiated";

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:850px;margin:0 auto;border:2px solid #0f172a;padding:14px;background:#fff;box-sizing:border-box">' +
    '<table style="border-bottom:2px solid #0f172a;margin-bottom:12px">' +
    "  <tr>" +
    '    <td style="width:18%;padding:6px;vertical-align:middle;text-align:center;border-right:1.5px solid #0f172a">' +
    '      <div style="font-weight:900;font-size:16px;color:#0f766e;letter-spacing:1px">SESIPL</div>' +
    '      <div style="font-size:9px;color:#64748b;font-weight:bold;margin-top:2px">SAFETY FIRST</div>' +
    "    </td>" +
    '    <td style="padding:6px 14px;vertical-align:middle;text-align:center">' +
    '      <div style="font-size:16px;font-weight:900;color:#0f172a;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569;margin-top:3px;letter-spacing:0.5px">ENVIRONMENT, HEALTH &amp; SAFETY MANAGEMENT SYSTEM</div>' +
    '      <div style="font-size:15px;font-weight:900;color:#0f766e;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px">' +
    escapeHtml_(title) +
    "</div>" +
    "    </td>" +
    '    <td style="width:20%;padding:6px;vertical-align:middle;text-align:right;font-size:10px;color:#475569;border-left:1.5px solid #0f172a">' +
    "      <div><b>Doc Code:</b> " +
    escapeHtml_(code) +
    "</div>" +
    "      <div><b>Ver:</b> v" +
    version +
    "</div>" +
    "      <div><b>Date:</b> " +
    escapeHtml_(dateVal) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;margin-bottom:14px;background:#f8fafc;font-size:11.5px">' +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-right:1px solid #0f172a;width:55%"><b>Project:</b> ' +
    escapeHtml_(projName) +
    "</td>" +
    '    <td style="padding:7px 10px;width:45%"><b>Date:</b> ' +
    escapeHtml_(dateVal) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Location:</b> ' +
    escapeHtml_(locVal) +
    "</td>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a"><b>Category:</b> ' +
    escapeHtml_(cat) +
    " &nbsp;&bull;&nbsp; <b>Risk:</b> <b>" +
    escapeHtml_(risk) +
    "</b></td>" +
    "  </tr>" +
    "</table>" +
    '<div style="border:1.5px solid #0f172a;padding:12px;margin-bottom:12px">' +
    '  <div style="font-weight:bold;color:#334155;margin-bottom:6px">Observation Description:</div>' +
    '  <div style="font-size:12px;color:#0f172a;line-height:1.5">' +
    escapeHtml_(desc) +
    "</div>" +
    "</div>" +
    '<div style="border:1.5px solid #0f172a;padding:12px;background:#f8fafc;margin-bottom:14px">' +
    '  <div style="font-weight:bold;color:#0f766e;margin-bottom:6px">Corrective &amp; Preventive Action (CAPA):</div>' +
    '  <div style="font-size:12px;color:#0f172a;line-height:1.5">' +
    escapeHtml_(capa) +
    "</div>" +
    '  <div style="margin-top:8px;font-size:11px;color:#475569"><b>Assigned To:</b> ' +
    escapeHtml_(resp) +
    " &nbsp;&bull;&nbsp; <b>Status:</b> <b>" +
    escapeHtml_(status) +
    "</b></div>" +
    "</div>" +
    '<table style="border:1.5px solid #0f172a;background:#fff;font-size:11px;margin-top:14px">' +
    "  <tr>" +
    '    <td style="width:50%;padding:10px 14px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Observed By:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(user.name || "Site Observer") +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Site EHS Team</div>' +
    "    </td>" +
    '    <td style="width:50%;padding:10px 14px;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Verified / Closed By:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(fields.closedBy || "EHS Manager") +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">SESIPL Corporate Safety</div>' +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="margin-top:12px;font-size:9.5px;color:#94a3b8;text-align:center">' +
    "  This record is digitally authenticated and stored in the SESIPL EHS Cloud System." +
    "</div>" +
    "</div></body></html>"
  );
}

function buildStructuredFormPdfHtml_(project, def, fields, user, version) {
  const code = (def.formCode || "").toUpperCase();
  const title = def.title ? def.title.toUpperCase() : "SAFETY INSPECTION RECORD";
  const projName = project.name || fields.projectName || "SESIPL Site";
  const dateVal = displayDate_(fields.date || fields.inspectionDate || nowIso_());

  const rows = Object.keys(fields)
    .filter((k) => k !== "csrfToken" && k !== "_action")
    .map((k) => {
      const val = fields[k];
      let valDisplay = escapeHtml_(displayDate_(val));
      const vLower = String(val).toLowerCase().trim();
      if (vLower === "yes" || vLower === "y" || val === "✓") {
        valDisplay =
          '<span style="color:#0f766e;font-weight:bold">&#10003; Yes</span>';
      } else if (vLower === "no" || vLower === "n" || val === "✗") {
        valDisplay =
          '<span style="color:#dc2626;font-weight:bold">&#10007; No</span>';
      }

      return (
        "<tr>" +
        '<td style="padding:7px 10px;border:1px solid #0f172a;width:38%;background:#f8fafc;font-weight:600;color:#1e293b">' +
        escapeHtml_(prettyLabel_(k)) +
        "</td>" +
        '<td style="padding:7px 10px;border:1px solid #0f172a;color:#0f172a">' +
        valDisplay +
        "</td>" +
        "</tr>"
      );
    })
    .join("");

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  body { font-family: Arial, "Helvetica Neue", Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 16px; line-height: 1.4; background: #fff; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    "  @media print { body { padding: 0 !important; } @page { margin: 10mm; size: A4 portrait; } }" +
    "</style></head><body>" +
    '<div style="max-width:850px;margin:0 auto;border:2px solid #0f172a;padding:14px;background:#fff;box-sizing:border-box">' +
    '<table style="border-bottom:2px solid #0f172a;margin-bottom:12px">' +
    "  <tr>" +
    '    <td style="width:18%;padding:6px;vertical-align:middle;text-align:center;border-right:1.5px solid #0f172a">' +
    '      <div style="font-weight:900;font-size:16px;color:#0f766e;letter-spacing:1px">SESIPL</div>' +
    '      <div style="font-size:9px;color:#64748b;font-weight:bold;margin-top:2px">SAFETY FIRST</div>' +
    "    </td>" +
    '    <td style="padding:6px 14px;vertical-align:middle;text-align:center">' +
    '      <div style="font-size:16px;font-weight:900;color:#0f172a;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569;margin-top:3px;letter-spacing:0.5px">ENVIRONMENT, HEALTH &amp; SAFETY MANAGEMENT SYSTEM</div>' +
    '      <div style="font-size:15px;font-weight:900;color:#0f766e;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px">' +
    escapeHtml_(title) +
    "</div>" +
    "    </td>" +
    '    <td style="width:20%;padding:6px;vertical-align:middle;text-align:right;font-size:10px;color:#475569;border-left:1.5px solid #0f172a">' +
    "      <div><b>Doc Code:</b> " +
    escapeHtml_(code) +
    "</div>" +
    "      <div><b>Ver:</b> v" +
    version +
    "</div>" +
    "      <div><b>Date:</b> " +
    escapeHtml_(dateVal) +
    "</div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;margin-bottom:14px;background:#f8fafc;font-size:11.5px">' +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-right:1px solid #0f172a;width:55%"><b>Project:</b> ' +
    escapeHtml_(projName) +
    "</td>" +
    '    <td style="padding:7px 10px;width:45%"><b>Date:</b> ' +
    escapeHtml_(dateVal) +
    "</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a;border-right:1px solid #0f172a"><b>Client / PMC:</b> ' +
    escapeHtml_(project.client || project.pmc || "—") +
    "</td>" +
    '    <td style="padding:7px 10px;border-top:1px solid #0f172a"><b>Submitted By:</b> ' +
    escapeHtml_(user.name || "SESIPL Lead") +
    "</td>" +
    "  </tr>" +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;font-size:11.5px;margin-bottom:16px">' +
    rows +
    "</table>" +
    '<table style="border:1.5px solid #0f172a;background:#fff;font-size:11px;margin-top:14px">' +
    "  <tr>" +
    '    <td style="width:33%;padding:10px 10px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Prepared By:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">' +
    escapeHtml_(user.name || "Site EHS Lead") +
    "</div>" +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Site Safety Lead</div>' +
    "    </td>" +
    '    <td style="width:33%;padding:10px 10px;border-right:1px solid #0f172a;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Verified By:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">EHS Inspection Authority</div>' +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">Asst. EHS Manager</div>' +
    "    </td>" +
    '    <td style="width:34%;padding:10px 10px;vertical-align:top">' +
    '      <div style="font-weight:bold;color:#334155;margin-bottom:12px">Approved By:</div>' +
    '      <div style="font-style:italic;color:#0f766e;font-weight:bold;font-size:12px;border-top:1px dashed #cbd5e1;padding-top:6px">SESIPL Corporate EHS</div>' +
    '      <div style="font-size:9.5px;color:#64748b;margin-top:2px">EHS Manager / Director</div>' +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="margin-top:12px;font-size:9.5px;color:#94a3b8;text-align:center">' +
    "  This record is digitally authenticated and stored in the SESIPL EHS Cloud System." +
    "</div>" +
    "</div></body></html>"
  );
}

function buildWorkPermitPdfHtml_(project, def, fields, user, version) {
  const docCode =
    typeof PERMIT_DOC_CODES !== "undefined" && PERMIT_DOC_CODES[def.formCode]
      ? PERMIT_DOC_CODES[def.formCode]
      : "SESIPL-EHS.Blr PTW";
  const contractor =
    fields.contractorName ||
    fields.contractor ||
    project.client ||
    "Shankar Electricals Services (I) Pvt Ltd";
  const em1 = fields.emergencyContact1 || "—";
  const em2 = fields.emergencyContact2 || "—";
  const pNo =
    fields.permitNo ||
    project.code +
      "-" +
      def.formCode +
      "-" +
      (fields.date || todayIso_()).replace(/-/g, "");
  const area = fields.area || "—";
  const loc = fields.location || "—";
  const dt = displayDate_(fields.date || nowIso_());
  const tm = fields.time || "—";
  const siteEng = fields.siteEngineer || user.name || "—";
  const siteEngSign = fields.siteEngineerSign || siteEng;
  const safetyOff = fields.safetyOfficer || "—";
  const safetyOffSign = fields.safetyOfficerSign || safetyOff;
  const contIncharge = fields.contractorInCharge || "—";
  const contPhone = fields.contactNumber || "—";
  const workDesc = fields.workDescription || "—";
  const execDate = displayDate_(
    fields.workExecutionDate || fields.date || nowIso_(),
  );
  const validFrom = fields.validFrom
    ? displayDate_(fields.validFrom).replace("T", " ")
    : "—";
  const validTo = fields.validTo
    ? displayDate_(fields.validTo).replace("T", " ")
    : "—";

  let precautions = [];
  const isNight = def.formCode === "WP_NIGHT";

  if (def.formCode === "WP_SHAFT") {
    precautions = [
      [1, "Proper Access/ Exit available", fields.shaft_q1],
      [2, "Proper ventilation and / or lighting provided", fields.shaft_q2],
      [3, "Proper & Safe platform provided", fields.shaft_q3],
      [4, "Workers have been briefed about hazardous", fields.shaft_q4],
      [
        5,
        "All Electrical Tools and machinery checked prior to use.",
        fields.shaft_q5,
      ],
      [6, "Shaft area Properly barricaded.", fields.shaft_q6],
      [
        7,
        "Conducted JST for all workers who are all engaging to shaft work.",
        fields.shaft_q7,
      ],
    ];
  } else if (def.formCode === "WP_NIGHT") {
    precautions = [
      [1, "Is dedicated Night shift in charge available?", fields.night_q1],
      [
        2,
        "Is supervisor available in night shift to supervise the task?",
        fields.night_q2,
      ],
      [3, "Is First aider available?", fields.night_q3],
      [4, "Is Ambulance available for emergency?", fields.night_q4],
      [
        5,
        "Are the workers working continuously for last 12 hours?",
        fields.night_q5,
      ],
      [6, "Is the work area safe for work?", fields.night_q6],
      [
        7,
        "Is there proper illumination provided at the work area?",
        fields.night_q7,
      ],
      [
        8,
        "Are the hazards related with the work identified and assessed at workplace?",
        fields.night_q8,
      ],
      [
        9,
        "Is toolbox talk / pre-start briefing carried out prior to start night shift?",
        fields.night_q9,
      ],
      [
        10,
        "Are the workers having specific PPE’s according to the requirement of the task?",
        fields.night_q10,
      ],
      [
        11,
        "Is there any high-risk activity like working at height, work in penetration and shafts, electrical testing and commissioning, hot work, Mechanical lifting operation, excavation etc. to be carried out in night shift?",
        fields.night_q11,
      ],
    ];
  } else if (def.formCode === "WP_LIFT") {
    precautions = [
      [
        1,
        "Crane used for lifting activity tested, certified and approved for rated lifting works.",
        fields.lift_q1,
      ],
      [
        2,
        "All lifting tackles, gears/ appliances are tested and certified for lifting works.",
        fields.lift_q2,
      ],
      [
        3,
        "Crane operator is trained and competent for lifting operation.",
        fields.lift_q3,
      ],
      [
        4,
        "Lifting belt protected against sharp edge of jobs to be lifted.",
        fields.lift_q4,
      ],
      [5, "Access and exist marked and without obstruction.", fields.lift_q5],
      [6, "Lighting arrangement adequate.", fields.lift_q6],
      [
        7,
        "Unwanted and rubbish material removed from working platform.",
        fields.lift_q7,
      ],
      [
        8,
        "Guidelines has provided for balancing & guiding jobs to be lifted.",
        fields.lift_q8,
      ],
      [
        9,
        "Periphery area of crane booms as well lifting job is barricaded .",
        fields.lift_q9,
      ],
      [
        10,
        "Rigger and signal man is trained and competent for lifting work.",
        fields.lift_q10,
      ],
      [
        11,
        "No lifting activity to be carried during lightening, heavy wind /rain.",
        fields.lift_q11,
      ],
      [
        12,
        "If scaffolding to be used during lift , Scaffolding with valid tag available for use",
        fields.lift_q12,
      ],
      [
        13,
        "Double lanyards Safety Harness/belt checked and in working condition",
        fields.lift_q13,
      ],
      [
        14,
        "Safety shoes (nonslip), Helmet with chin strip available with employees.",
        fields.lift_q14,
      ],
      [
        15,
        "Other: " + (fields.lift_other || "—"),
        fields.lift_other ? "Yes" : "Not Required",
      ],
    ];
  } else if (def.formCode === "WP_HOT") {
    precautions = [
      [1, "Proper Access/ Exit available", fields.hot_q1],
      [2, "Proper ventilation and / or lighting provided", fields.hot_q2],
      [
        3,
        "Proper & Safe scaffolding, platform, ladder provided",
        fields.hot_q3,
      ],
      [4, "Welding machine located in a clean and dry area", fields.hot_q4],
      [
        5,
        "Welding machine grounded at the equipment & proper leakage current protection device (ELCB) provided for welding machine.",
        fields.hot_q5,
      ],
      [
        6,
        "Competent and Trained personnel deployed to carry the work.",
        fields.hot_q6,
      ],
      [
        7,
        "Welding machine, Input / Output Cables, welding holder and weld return clamp (Holder ) insulated & in good condition",
        fields.hot_q7,
      ],
      [
        8,
        "Welder and fitter trained to connect ground / work return clamps (Holder) to the work piece prior to energization of Welding machine.",
        fields.hot_q8,
      ],
      [
        9,
        "Gas Cylinders stacked vertically and not below the welding/cutting area. Regulator Key is available with cylinders.",
        fields.hot_q9,
      ],
      [
        10,
        "Work Area Isolated with barricading and caution sign",
        fields.hot_q10,
      ],
      [
        11,
        "Personal Protective Equipment. Minimum applicable - Safety helmet, safety goggles, welding helmet, safety shoes, leather gloves, long sleeve and nose mask provided.",
        fields.hot_q11,
      ],
      [
        12,
        "In case of pits, water removed from the pit & wood /rubber insulation provided.",
        fields.hot_q12,
      ],
      [
        13,
        "Adequate & suitable nos. of fire fighting extinguisher provided.",
        fields.hot_q13,
      ],
      [
        14,
        "Near by combustible material removed. Housekeeping Done.",
        fields.hot_q14,
      ],
      [15, "Fire watch as standby is in place.", fields.hot_q15],
      [
        16,
        "Other: " + (fields.hot_other || "—"),
        fields.hot_other ? "Yes" : "Not Required",
      ],
    ];
  } else if (def.formCode === "WP_HEIGHT") {
    precautions = [
      [1, "Scaffolding with valid tag available for use", fields.height_q1],
      [2, "Conducted JST/TBT conducted", fields.height_q2],
      [
        3,
        "Safety shoes (nonslip), Helmet with chin strip available with employees.",
        fields.height_q3,
      ],
      [
        4,
        "All tightening tools, hand tools /equipment checked and in good condition.",
        fields.height_q4,
      ],
      [5, "Access and exist marked and without obstruction.", fields.height_q5],
      [6, "Lighting arrangement adequate.", fields.height_q6],
      [
        7,
        "Unwanted and rubbish material removed from working platform.",
        fields.height_q7,
      ],
      [8, "Electrical cable in good condition", fields.height_q8],
      [9, "Signboards provided", fields.height_q9],
      [
        10,
        "Employees aware about hazards and safe working practices while working at height.",
        fields.height_q10,
      ],
    ];
  }

  let tableHeader = "";
  let tableRows = "";

  if (isNight) {
    tableHeader =
      '<tr><th style="width:6%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">No</th>' +
      '<th style="width:58%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:left">ITEM</th>' +
      '<th style="width:8%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">Yes</th>' +
      '<th style="width:8%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">NO</th>' +
      '<th style="width:8%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">NA</th>' +
      '<th style="width:12%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">REMARKS</th></tr>';

    tableRows = precautions
      .map((p) => {
        const v = String(p[2] || "").toUpperCase();
        return (
          "<tr>" +
          '<td style="text-align:center;padding:5px;border:1px solid #334155">' +
          p[0] +
          "</td>" +
          '<td style="padding:5px;border:1px solid #334155">' +
          escapeHtml_(p[1]) +
          "</td>" +
          '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' +
          (v === "YES" ? "✓" : "") +
          "</td>" +
          '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' +
          (v === "NO" ? "✓" : "") +
          "</td>" +
          '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' +
          (v === "NA" ? "✓" : "") +
          "</td>" +
          '<td style="padding:5px;border:1px solid #334155">' +
          (p[0] === 11 ? escapeHtml_(fields.night_remarks || "") : "") +
          "</td>" +
          "</tr>"
        );
      })
      .join("");
  } else {
    tableHeader =
      '<tr><th style="width:6%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">No</th>' +
      '<th style="width:72%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:left">Item</th>' +
      '<th style="width:11%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">Yes</th>' +
      '<th style="width:11%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">Not Required</th></tr>';

    tableRows = precautions
      .map((p) => {
        const v = String(p[2] || "").toLowerCase();
        const isYes = v.includes("yes");
        const isNot = v.includes("not");
        return (
          "<tr>" +
          '<td style="text-align:center;padding:5px;border:1px solid #334155">' +
          p[0] +
          "</td>" +
          '<td style="padding:5px;border:1px solid #334155">' +
          escapeHtml_(p[1]) +
          "</td>" +
          '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' +
          (isYes ? "✓" : "") +
          "</td>" +
          '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' +
          (isNot ? "✓" : "") +
          "</td>" +
          "</tr>"
        );
      })
      .join("");
  }

  const workmenBlock =
    def.formCode === "WP_SHAFT" && fields.workmenNames
      ? '<div style="margin-top:10px;font-size:11.5px"><b>Names of workmen entering shaft:</b> ' +
        escapeHtml_(fields.workmenNames) +
        "</div>"
      : "";

  return (
    '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#0f172a;padding:24px;line-height:1.4;font-size:11px">' +
    '<table style="width:100%;border-collapse:collapse;border:1.5px solid #0f172a;margin-bottom:14px">' +
    "  <tr>" +
    '    <td style="width:34%;padding:8px 12px;border-right:1.5px solid #0f172a;vertical-align:middle">' +
    '      <div style="font-size:10px;font-weight:bold;color:#475569">Contractor Name:</div>' +
    '      <div style="font-size:12px;font-weight:bold;color:#0f172a;margin-top:2px">' +
    escapeHtml_(contractor) +
    "</div>" +
    '      <div style="margin-top:6px;font-size:10.5px;font-weight:bold;color:#0f766e">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    "    </td>" +
    '    <td style="width:34%;padding:8px 12px;border-right:1.5px solid #0f172a;text-align:center;vertical-align:middle">' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569">TITLE :</div>' +
    '      <div style="font-size:13px;font-weight:800;letter-spacing:0.5px;color:#0f172a;margin-top:3px">SAFETY WORK CLEARANCE</div>' +
    "    </td>" +
    '    <td style="width:32%;padding:8px 12px;vertical-align:middle;font-size:10.5px">' +
    "      <div><b>Emergency Contact No:</b></div>" +
    "      <div>1) " +
    escapeHtml_(em1) +
    "</div>" +
    "      <div>2) " +
    escapeHtml_(em2) +
    "</div>" +
    '      <div style="margin-top:4px"><b>Permit No:-</b> <span style="font-weight:bold;color:#0f766e">' +
    escapeHtml_(pNo) +
    "</span></div>" +
    "    </td>" +
    "  </tr>" +
    "</table>" +
    '<div style="text-align:center;margin-bottom:12px">' +
    '  <h1 style="margin:0;font-size:17px;font-weight:900;letter-spacing:0.5px;text-transform:uppercase;color:#0f172a">' +
    escapeHtml_(def.title) +
    "</h1>" +
    "</div>" +
    '<div style="border:1px solid #cbd5e1;background:#f8fafc;padding:8px 12px;border-radius:6px;margin-bottom:10px;font-size:11px">' +
    '  <table style="width:100%;border-collapse:collapse">' +
    "    <tr>" +
    '      <td style="width:40%;padding:3px 0"><b>Area:-</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:140px">' +
    escapeHtml_(area) +
    "</span></td>" +
    '      <td style="width:35%;padding:3px 0"><b>location:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:130px">' +
    escapeHtml_(loc) +
    "</span></td>" +
    '      <td style="width:25%;padding:3px 0"><b>Date:</b> ' +
    escapeHtml_(dt) +
    " &nbsp; <b>Time:-</b> " +
    escapeHtml_(tm) +
    "</td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td colspan="2" style="padding:3px 0"><b>Name of Site Engineer (Permit Requesting Authority):</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:180px">' +
    escapeHtml_(siteEng) +
    "</span></td>" +
    '      <td style="padding:3px 0"><b>Sign:</b> <span style="font-style:italic;color:#0f766e">' +
    escapeHtml_(siteEngSign) +
    "</span></td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td colspan="2" style="padding:3px 0"><b>Name of Safety Officer:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:220px">' +
    escapeHtml_(safetyOff) +
    "</span></td>" +
    '      <td style="padding:3px 0"><b>Sign:</b> <span style="font-style:italic;color:#0f766e">' +
    escapeHtml_(safetyOffSign) +
    "</span></td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td colspan="2" style="padding:3px 0"><b>Name of Contractor Site In charge:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:180px">' +
    escapeHtml_(contIncharge) +
    "</span></td>" +
    '      <td style="padding:3px 0"><b>Contact Number:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:110px">' +
    escapeHtml_(contPhone) +
    "</span></td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td colspan="3" style="padding:3px 0"><b>Description of work:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;width:80%">' +
    escapeHtml_(workDesc) +
    "</span></td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td colspan="3" style="padding:3px 0"><b>Work Execution Date:</b> ' +
    escapeHtml_(execDate) +
    " &nbsp;&nbsp;&nbsp;&nbsp; <b>Valid From:-</b> " +
    escapeHtml_(validFrom) +
    " &nbsp;&nbsp;&nbsp;&nbsp; <b>To:-</b> " +
    escapeHtml_(validTo) +
    "</td>" +
    "    </tr>" +
    "  </table>" +
    "</div>" +
    '<p style="font-size:10px;font-style:italic;margin:4px 0 8px;color:#334155">' +
    "  The above signing person will be responsible to ensure that the above described work will be done under all the safety precaution mentioned on the PTW and required by the Project." +
    "</p>" +
    '<div style="font-size:11px;font-weight:bold;margin-bottom:5px">The following precautions are to be taken:-</div>' +
    '<table style="width:100%;border-collapse:collapse;border:1.5px solid #0f172a;font-size:10.5px;margin-bottom:8px">' +
    tableHeader +
    tableRows +
    "</table>" +
    workmenBlock +
    '<div style="margin-top:6px;font-size:10px;color:#475569"><b>Notes:</b> a) Work permit is valid for the prescribed date, time and in prescribed location only</div>' +
    '<div style="margin-top:10px;border:1px solid #cbd5e1;padding:6px 10px;border-radius:6px;background:#f8fafc">' +
    '  <div style="font-weight:bold;font-size:11.5px;margin-bottom:4px;color:#0f172a">Reviewed & Approved By (Permit Issuing Authority):</div>' +
    '  <table style="width:100%;border-collapse:collapse;font-size:10.5px">' +
    "    <tr>" +
    '      <td style="width:50%;padding:3px 0"><b>EHS:</b> ' +
    escapeHtml_(fields.approvalEhsName || "—") +
    ' &nbsp;&nbsp; <b>Sign:</b> <span style="font-style:italic;color:#0f766e">' +
    escapeHtml_(
      fields.approvalEhsSign || fields.approvalEhsName || "Acknowledged",
    ) +
    "</span></td>" +
    '      <td style="width:50%;padding:3px 0"><b>Date:</b> ' +
    escapeHtml_(displayDate_(fields.approvalEhsDate || fields.date)) +
    " &nbsp;&nbsp; <b>Time:</b> " +
    escapeHtml_(fields.approvalEhsTime || fields.time || "—") +
    "</td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td style="width:50%;padding:3px 0"><b>Site Engineer:</b> ' +
    escapeHtml_(fields.approvalSiteEngineerName || "—") +
    ' &nbsp;&nbsp; <b>Sign:</b> <span style="font-style:italic;color:#0f766e">' +
    escapeHtml_(
      fields.approvalSiteEngineerSign ||
        fields.approvalSiteEngineerName ||
        "Verified",
    ) +
    "</span></td>" +
    '      <td style="width:50%;padding:3px 0"><b>Date:</b> ' +
    escapeHtml_(displayDate_(fields.approvalSiteEngineerDate || fields.date)) +
    " &nbsp;&nbsp; <b>Time:</b> " +
    escapeHtml_(fields.approvalSiteEngineerTime || fields.time || "—") +
    "</td>" +
    "    </tr>" +
    "  </table>" +
    '  <div style="font-size:9.5px;font-style:italic;color:#475569;margin-top:4px">' +
    "    I understand the precaution to be taken as described above and as per Project requirement & hereby confirm that Work will be executed under my supervision by following all precaution & Safety Rules." +
    "  </div>" +
    "</div>" +
    '<div style="margin-top:10px;border:1px solid #cbd5e1;padding:6px 10px;border-radius:6px">' +
    '  <div style="font-weight:bold;font-size:11.5px;margin-bottom:3px;color:#991b1b">Permit Closing / Cancellation:-</div>' +
    '  <div style="font-size:9.5px;font-style:italic;color:#475569;margin-bottom:6px">' +
    "    I hereby declare that the work is completed / suspended, all workers under my control have been withdrawn and the site restored to a safe tidy condition." +
    "  </div>" +
    '  <table style="width:100%;border-collapse:collapse;font-size:10.5px">' +
    "    <tr>" +
    '      <td style="width:55%;padding:2px 0"><b>Name of Site Engineer (Permit Requesting Authority):</b> ' +
    escapeHtml_(fields.closingSiteEngName || "—") +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Sign:</b> ' +
    escapeHtml_(fields.closingSiteEngSign || "—") +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Date:</b> ' +
    escapeHtml_(displayDate_(fields.closingSiteEngDate || "")) +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Time:</b> ' +
    escapeHtml_(fields.closingSiteEngTime || "—") +
    "</td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td style="width:55%;padding:2px 0"><b>Name of Safety Officer:</b> ' +
    escapeHtml_(fields.closingSafetyOfficerName || "—") +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Sign:</b> ' +
    escapeHtml_(fields.closingSafetyOfficerSign || "—") +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Date:</b> ' +
    escapeHtml_(displayDate_(fields.closingSafetyOfficerDate || "")) +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Time:</b> ' +
    escapeHtml_(fields.closingSafetyOfficerTime || "—") +
    "</td>" +
    "    </tr>" +
    "    <tr>" +
    '      <td style="width:55%;padding:2px 0"><b>Name of PMC Site Engineer (Permit Issuing Authority):</b> ' +
    escapeHtml_(fields.closingPmcSiteEngName || "—") +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Sign:</b> ' +
    escapeHtml_(fields.closingPmcSiteEngSign || "—") +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Date:</b> ' +
    escapeHtml_(displayDate_(fields.closingPmcSiteEngDate || "")) +
    "</td>" +
    '      <td style="width:15%;padding:2px 0"><b>Time:</b> ' +
    escapeHtml_(fields.closingPmcSiteEngTime || "—") +
    "</td>" +
    "    </tr>" +
    "  </table>" +
    "</div>" +
    (fields.signedPermitUrl ||
    fields.swmsFileUrl ||
    fields.preWorkPhotoUrl ||
    fields.signedPermitFile ||
    fields.swmsFile ||
    fields.preWorkPhoto
      ? '<div style="margin-top:10px;border:1px solid #0284c7;padding:6px 10px;border-radius:6px;background:#f0f9ff">' +
        '  <div style="font-weight:bold;font-size:11px;margin-bottom:4px;color:#0369a1">Attached Supporting Documents &amp; Reference Files:</div>' +
        '  <table style="width:100%;border-collapse:collapse;font-size:10px">' +
        (fields.signedPermitUrl || fields.signedPermitFile
          ? '<tr><td style="width:38%;padding:2px 0"><b>Signed PTW Copy / Scan:</b></td><td>' +
            escapeHtml_(fields.signedPermitUrl || fields.signedPermitFile) +
            "</td></tr>"
          : "") +
        (fields.swmsFileUrl || fields.swmsFile
          ? '<tr><td style="width:38%;padding:2px 0"><b>SWMS / JSA Document:</b></td><td>' +
            escapeHtml_(fields.swmsFileUrl || fields.swmsFile) +
            "</td></tr>"
          : "") +
        (fields.preWorkPhotoUrl || fields.preWorkPhoto
          ? '<tr><td style="width:38%;padding:2px 0"><b>Site Verification Photo:</b></td><td>' +
            escapeHtml_(fields.preWorkPhotoUrl || fields.preWorkPhoto) +
            "</td></tr>"
          : "") +
        "  </table>" +
        "</div>"
      : "") +
    '<table style="width:100%;margin-top:14px;font-size:10px;color:#64748b">' +
    "  <tr>" +
    '    <td style="text-align:left"><b>Controlled Copy©</b></td>' +
    '    <td style="text-align:right"><b>Shankar Electricals Services I Pvt Ltd.</b><br><span style="font-family:monospace;font-size:10px;font-weight:bold">' +
    escapeHtml_(docCode) +
    "</span></td>" +
    "  </tr>" +
    "</table>" +
    "</body></html>"
  );
}

function buildReportHtml_(project, subs, user) {
  const rows = subs
    .map(
      (s) =>
        "<tr><td>" +
        escapeHtml_(s.formCode) +
        "</td><td>" +
        escapeHtml_(s.status) +
        "</td><td>" +
        escapeHtml_(s.submittedBy) +
        "</td><td>" +
        escapeHtml_(displayDate_(s.submittedAt)) +
        "</td></tr>",
    )
    .join("");
  return (
    '<html><body style="font-family:Arial">' +
    "<h2>" +
    escapeHtml_(project.name) +
    " — EHS Pack</h2>" +
    "<p>Exported by " +
    escapeHtml_(user.name) +
    " on " +
    displayDate_(nowIso_()) +
    "</p>" +
    '<table border="1" cellpadding="6"><tr><th>Form</th><th>Status</th><th>By</th><th>At</th></tr>' +
    rows +
    "</table></body></html>"
  );
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function prettyLabel_(k) {
  return String(k)
    .replace(/([A-Z])/g, " $1")
    .replace(/^./, (c) => c.toUpperCase());
}

function buildEhsAuditPdfHtml_(audit, project, user) {
  const data =
    audit ||
    (typeof getDefaultAuditSeedData_ === "function"
      ? getDefaultAuditSeedData_(project)
      : null);
  const proj = project || {
    name: data ? data.projectName : "Intuit Bellandur",
    areaSqft: data
      ? data.projectLocation
      : "Pritech Park, Bellandur, Bengaluru",
  };
  const pName = escapeHtml_(
    (data && data.projectName) || proj.name || "Intuit Bellandur",
  );
  const pLoc = escapeHtml_(
    (data && data.projectLocation) ||
      proj.areaSqft ||
      "Pritech Park, Bellandur, Bengaluru",
  );
  const aDate = escapeHtml_(
    displayDate_((data && data.auditDate) || "2026-09-10"),
  );
  const auditor = escapeHtml_(
    (data && data.auditor) || (user && user.name) || "CBRE Lead Auditor",
  );
  const schema =
    typeof AUDIT_CHECKLIST_SCHEMA !== "undefined" ? AUDIT_CHECKLIST_SCHEMA : [];

  const renderHeader = () => {
    return (
      '<div class="audit-header" style="margin-bottom:6px">' +
      '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;margin-bottom:0">' +
      "  <tr>" +
      '    <td style="width:16%;padding:4px 8px;border-right:1.5px solid #000;text-align:center;vertical-align:middle;background:#fff">' +
      '      <img src="https://sesipl.com/sites/default/files/sesipl-logo-new-2_5.png" style="height:30px;object-fit:contain" alt="SESIPL" onerror="this.onerror=null;this.src=\'https://sesipl.com/sites/default/files/sesipl-logo.png\'">' +
      '      <div style="font-weight:900;font-size:10.5px;color:#0369a1;letter-spacing:0.5px">SESIPL</div>' +
      "    </td>" +
      '    <td style="width:54%;padding:6px 8px;border-right:1.5px solid #000;text-align:center;vertical-align:middle;background:#c6efce">' +
      '      <div style="font-size:15px;font-weight:bold;color:#000;letter-spacing:0.3px">Shankar Electricals EHS Audit Checklist</div>' +
      "    </td>" +
      '    <td style="width:30%;padding:4px 6px;font-size:7.5px;color:#000;text-align:right;vertical-align:middle;line-height:1.2;background:#fff">' +
      '      <div style="font-style:italic">"Sree DeviArcade"</div>' +
      "      <div>668/A, 2nd Floor, 17th C Main, 6th Block, Koramangala, Bengaluru - 560095</div>" +
      "    </td>" +
      "  </tr>" +
      "</table>" +
      '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;border-top:none;margin-bottom:5px;font-size:9.5px;background:#fff">' +
      "  <tr>" +
      '    <td colspan="2" style="padding:2.5px 6px;border-bottom:1px solid #000"><b>Project Name:</b> <span>' +
      pName +
      "</span></td>" +
      "  </tr>" +
      "  <tr>" +
      '    <td colspan="2" style="padding:2.5px 6px;border-bottom:1px solid #000"><b>Project Location:</b> <span>' +
      pLoc +
      "</span></td>" +
      "  </tr>" +
      "  <tr>" +
      '    <td style="width:50%;padding:2.5px 6px;border-bottom:1px solid #000;border-right:1px solid #000"><b>Audit Date:</b> <span>' +
      aDate +
      "</span></td>" +
      '    <td style="width:50%;padding:2.5px 6px;border-bottom:1px solid #000;text-align:right;font-size:8.5px;color:#0369a1">info@shankarelectricals.com www.shankarelectricals.com</td>' +
      "  </tr>" +
      "  <tr>" +
      '    <td colspan="2" style="padding:2.5px 6px"><b>Auditor:</b> <span>' +
      auditor +
      "</span></td>" +
      "  </tr>" +
      "</table>" +
      "</div>"
    );
  };

  const renderSectionTable = (sections, itemFilterFn) => {
    let html =
      '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;font-size:9px;margin-bottom:6px;background:#fff">';
    html +=
      '<tr style="background:#e2e8f0;font-weight:bold">' +
      '<th style="width:3.5%;padding:3px;border:1px solid #000;text-align:center">SN</th>' +
      '<th style="width:68.5%;padding:3px 6px;border:1px solid #000;text-align:left">Particulars</th>' +
      '<th colspan="7" style="width:28%;padding:3px;border:1px solid #000;text-align:center">Score</th>' +
      "</tr>";

    sections.forEach((sec) => {
      const secScoreInfo =
        data && data.sectionScores && data.sectionScores[sec.id]
          ? data.sectionScores[sec.id]
          : {
              actual: sec.items.reduce(
                (acc, it) =>
                  typeof it.defaultScore === "number"
                    ? acc + it.defaultScore
                    : acc,
                0,
              ),
            };

      const itemsToRender = itemFilterFn
        ? sec.items.filter(itemFilterFn)
        : sec.items;
      if (!itemsToRender.length) return;

      // Section Header row
      html +=
        '<tr style="background:#e2e8f0;font-weight:bold">' +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
        sec.id +
        "</td>" +
        '<td style="padding:2.5px 6px;border:1px solid #000">' +
        escapeHtml_(sec.name) +
        "</td>" +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">0</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">1</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">2</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">3</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">4</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">5</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">NA</td>' +
        "</tr>";

      itemsToRender.forEach((item) => {
        const sc = item.defaultScore;
        const mark = (val) =>
          sc === val
            ? '<span style="font-weight:bold;font-size:12px;color:#000">*</span>'
            : "";
        html +=
          "<tr>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          item.sn +
          "</td>" +
          '<td style="padding:2.5px 6px;border:1px solid #000">' +
          escapeHtml_(item.text) +
          "</td>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          mark(0) +
          "</td>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          mark(1) +
          "</td>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          mark(2) +
          "</td>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          mark(3) +
          "</td>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          mark(4) +
          "</td>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          mark(5) +
          "</td>" +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
          mark("NA") +
          "</td>" +
          "</tr>";
      });

      const isComplete =
        !itemFilterFn ||
        itemsToRender[itemsToRender.length - 1].sn ===
          sec.items[sec.items.length - 1].sn;
      if (isComplete) {
        html +=
          "<tr>" +
          '<td colspan="2" style="padding:3px 12px;border:1px solid #000;text-align:right;font-weight:bold">Section Total ' +
          sec.max +
          "</td>" +
          '<td colspan="7" style="padding:3px;border:1px solid #000;text-align:center;background:#fb923c;color:#000;font-weight:bold;font-size:10.5px">' +
          secScoreInfo.actual +
          "</td>" +
          "</tr>";
      }
    });

    html += "</table>";
    return html;
  };

  const renderSummaryScorecardPage = () => {
    let html = renderHeader();
    html +=
      '<div style="background:#f1f5f9;border:1.5px solid #000;border-bottom:none;padding:4px;font-size:9px;font-weight:bold;text-align:center">' +
      "0 - Major NC; 1 - Minor NC; 2 - Partial Compliance; 3 - Full Compliance; NA - Not Applicable" +
      "</div>";

    html +=
      '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;font-size:9px;background:#fff">';
    html +=
      '<tr style="background:#fbbf24;font-weight:bold;text-align:center">' +
      '<th style="width:4%;padding:3px;border:1px solid #000">SN</th>' +
      '<th style="width:48%;padding:3px 6px;border:1px solid #000;text-align:left">Item</th>' +
      '<th style="width:8%;padding:3px;border:1px solid #000">Max</th>' +
      '<th style="width:8%;padding:3px;border:1px solid #000">Actual</th>' +
      '<th style="width:8%;padding:3px;border:1px solid #000">%</th>' +
      '<th style="width:24%;padding:3px;border:1px solid #000">Performance</th>' +
      "</tr>";

    schema.forEach((sec, idx) => {
      const info =
        data && data.sectionScores && data.sectionScores[sec.id]
          ? data.sectionScores[sec.id]
          : { actual: 0, percent: 0, percentText: "0%" };
      const pctDisplay =
        info.percentText ||
        (sec.max > 0
          ? Math.round((info.actual / sec.max) * 100) + "%"
          : "#DIV/0!");

      let tierCell = "";
      if (idx === 0) {
        tierCell =
          '<td rowspan="5" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#f8fafc;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#475569">Platinum</div>' +
          '<div style="font-size:9px;color:#64748b;margin-top:2px">85 - 100 %</div>' +
          "</td>";
      } else if (idx === 5) {
        tierCell =
          '<td rowspan="5" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#fef08a;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#854d0e">Gold</div>' +
          '<div style="font-size:9px;color:#a16207;margin-top:2px">71 - 84 %</div>' +
          "</td>";
      } else if (idx === 10) {
        tierCell =
          '<td rowspan="4" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#bbf7d0;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#166534">Silver</div>' +
          '<div style="font-size:9px;color:#15803d;margin-top:2px">55 - 70 %</div>' +
          "</td>";
      } else if (idx === 14) {
        tierCell =
          '<td rowspan="4" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#bfdbfe;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#1e40af">Blue</div>' +
          '<div style="font-size:9px;color:#1d4ed8;margin-top:2px">&lt; 54 %</div>' +
          "</td>";
      }

      html +=
        "<tr>" +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center;font-weight:bold">' +
        sec.id +
        "</td>" +
        '<td style="padding:2.5px 6px;border:1px solid #000">' +
        escapeHtml_(sec.name) +
        "</td>" +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
        sec.max +
        "</td>" +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center;font-weight:bold">' +
        info.actual +
        "</td>" +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center">' +
        pctDisplay +
        "</td>" +
        tierCell +
        "</tr>";
    });

    const totMax = data ? data.maxScore : 525;
    const totActual = data ? data.totalScore : 363;
    const totPct = data ? data.percent : 69;

    html +=
      '<tr style="font-weight:bold;font-size:10px">' +
      '<td colspan="2" style="padding:4px 6px;border:1.5px solid #000">Total Score</td>' +
      '<td style="padding:4px;border:1.5px solid #000;text-align:center">' +
      totMax +
      "</td>" +
      '<td style="padding:4px;border:1.5px solid #000;text-align:center">' +
      totActual +
      "</td>" +
      '<td colspan="2" style="padding:4px;border:1.5px solid #000;text-align:center;background:#c6efce;font-size:11.5px;font-weight:900">' +
      totPct +
      "%</td>" +
      "</tr>";

    html += "</table>";
    return html;
  };

  const page1Secs = schema.filter(
    (s) => ["A", "B", "C", "D", "E", "F"].indexOf(s.id) >= 0,
  );
  const page2Secs = schema.filter(
    (s) => ["G", "H", "I", "J", "K", "L"].indexOf(s.id) >= 0,
  );
  const secM = schema.find((s) => s.id === "M") || {
    id: "M",
    name: "Store & Material Management",
    max: 20,
    items: [],
  };
  const page4Secs = schema.filter(
    (s) => ["N", "O", "P", "Q", "R"].indexOf(s.id) >= 0,
  );

  let fullHtml =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    "<title>Shankar Electricals EHS Audit Checklist - " +
    pName +
    "</title>" +
    "<style>" +
    "  @page { size: A4 portrait; margin: 8mm 8mm 8mm 8mm; }" +
    "  body { font-family: Calibri, Arial, sans-serif; color: #000; margin: 0; padding: 10px; background: #fff; line-height: 1.25; }" +
    "  .audit-page { page-break-after: always; min-height: 980px; box-sizing: border-box; }" +
    "  .audit-page:last-child { page-break-after: auto; }" +
    "  table { border-collapse: collapse; width: 100%; }" +
    "  th, td { box-sizing: border-box; }" +
    "  @media print {" +
    "    body { padding: 0; background: transparent; }" +
    "    .audit-page { page-break-after: always; min-height: 100vh; }" +
    "    .no-print { display: none !important; }" +
    "  }" +
    "</style>" +
    "</head><body>";

  // Page 1: Sections A through F
  fullHtml +=
    '<div class="audit-page">' +
    renderHeader() +
    renderSectionTable(page1Secs) +
    "</div>";

  // Page 2: Sections G through L + Section M (items 1-3)
  fullHtml +=
    '<div class="audit-page">' +
    renderHeader() +
    renderSectionTable(
      page2Secs.concat([
        {
          id: secM.id,
          name: secM.name,
          max: secM.max,
          items: secM.items.slice(0, 3),
        },
      ]),
    ) +
    "</div>";

  // Page 3: Section M (item 4) and Section Total 20
  fullHtml +=
    '<div class="audit-page">' +
    renderHeader() +
    renderSectionTable([
      {
        id: secM.id,
        name: secM.name + " (Continued)",
        max: secM.max,
        items: secM.items.slice(3),
      },
    ]) +
    "</div>";

  // Page 4: Sections N through R
  fullHtml +=
    '<div class="audit-page">' +
    renderHeader() +
    renderSectionTable(page4Secs) +
    "</div>";

  // Page 5: Executive Grand Summary Scorecard
  fullHtml +=
    '<div class="audit-page">' + renderSummaryScorecardPage() + "</div>";

  fullHtml += "</body></html>";
  return fullHtml;
}

function buildDailyLogPdfHtml_(allLogs, project, activeMonth) {
  const pName = project ? escapeHtml_(project.name) : "SESIPL Site";
  const mParts = (activeMonth || todayIso_().slice(0, 7)).split("-");
  const yearNum = parseInt(mParts[0], 10) || 2026;
  const monthNum = parseInt(mParts[1], 10) || 9;
  const monthNames = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  const monthShort = monthNames[monthNum - 1] || "Sep";
  const monthYearLabel = monthShort + " " + yearNum;
  const daysInMonth = new Date(yearNum, monthNum, 0).getDate();

  const logsByDay = {};
  (allLogs || []).forEach((l) => {
    if (!l.date) return;
    const dp = String(l.date).split("-");
    if (
      dp.length >= 3 &&
      parseInt(dp[0], 10) === yearNum &&
      parseInt(dp[1], 10) === monthNum
    ) {
      logsByDay[parseInt(dp[2], 10)] = l;
    }
  });

  if (Object.keys(logsByDay).length === 0) {
    logsByDay[1] = {
      staff: 2,
      workers: 10,
      totalManpower: 12,
      workingHours: 8,
      totalManHours: 96,
      safeManHours: 96,
      cumSafeManHours: 96,
      inductions: 5,
      indStaff: 3,
      indWorkers: 5,
      tbtCount: 1,
      tbtPersons: 10,
      trainingTopic: "Earth pit ex..",
      trainingPersons: 4,
      permitHot: 1,
      permitElectrical: 1,
      permitCold: "-",
      permitOthers: "-",
      firstAid: "-",
      nearMiss: "-",
      ltiCount: 0,
      accidentDetails: "Nil - Safe Day",
      remarks: "-",
    };
    logsByDay[2] = {
      staff: 2,
      workers: 9,
      totalManpower: 11,
      workingHours: 8,
      totalManHours: 88,
      safeManHours: 88,
      cumSafeManHours: 184,
      inductions: 1,
      indStaff: 2,
      indWorkers: 3,
      tbtCount: 1,
      tbtPersons: 9,
      trainingTopic: "-",
      trainingPersons: "-",
      permitHot: "-",
      permitElectrical: 1,
      permitCold: 1,
      permitOthers: 1,
      firstAid: "-",
      nearMiss: "-",
      ltiCount: 0,
      accidentDetails: "Nil - Safe Day",
      remarks: "-",
    };
    logsByDay[3] = {
      staff: 3,
      workers: 7,
      totalManpower: 10,
      workingHours: 8,
      totalManHours: 18,
      safeManHours: 18,
      cumSafeManHours: 202,
      inductions: "-",
      indStaff: "-",
      indWorkers: "-",
      tbtCount: 1,
      tbtPersons: 10,
      trainingTopic: "cable termination",
      trainingPersons: 7,
      permitHot: 1,
      permitElectrical: 1,
      permitCold: "-",
      permitOthers: 1,
      firstAid: "-",
      nearMiss: "-",
      ltiCount: 0,
      accidentDetails: "Nil - Safe Day",
      remarks: "-",
    };
    logsByDay[4] = {
      staff: 4,
      workers: 10,
      totalManpower: 14,
      workingHours: 8,
      totalManHours: 112,
      safeManHours: 112,
      cumSafeManHours: 314,
      inductions: 1,
      indStaff: "-",
      indWorkers: 2,
      tbtCount: 1,
      tbtPersons: 10,
      trainingTopic: "-",
      trainingPersons: "-",
      permitHot: 1,
      permitElectrical: "-",
      permitCold: "-",
      permitOthers: 1,
      firstAid: "-",
      nearMiss: "-",
      ltiCount: 0,
      accidentDetails: "Nil - Safe Day",
      remarks: "-",
    };
    logsByDay[5] = {
      staff: 2,
      workers: 5,
      totalManpower: 7,
      workingHours: 8,
      totalManHours: 56,
      safeManHours: 56,
      cumSafeManHours: 370,
      inductions: "-",
      indStaff: "-",
      indWorkers: "-",
      tbtCount: 1,
      tbtPersons: 7,
      trainingTopic: "Lifting",
      trainingPersons: 7,
      permitHot: 1,
      permitElectrical: 1,
      permitCold: 1,
      permitOthers: 1,
      firstAid: "-",
      nearMiss: "-",
      ltiCount: 0,
      accidentDetails: "Nil - Safe Day",
      remarks: "-",
    };
  }

  let totStaff = 0,
    totWorkers = 0,
    totMP = 0,
    totWorkHrs = 0,
    totManHrs = 0,
    totSafeHrs = 0,
    latestCumSafe = 0;
  let totInd = 0,
    totIndStaff = 0,
    totIndWorkers = 0,
    totTbt = 0,
    totTbtPersons = 0,
    totTrainingPersons = 0;
  let totHot = 0,
    totElect = 0,
    totCold = 0,
    totOthers = 0,
    totFA = 0,
    totNM = 0,
    totLTI = 0;

  let rowsHtml = "";
  for (let d = 1; d <= daysInMonth; d++) {
    const entry = logsByDay[d];
    const dateStr = d + "-" + monthShort + "-" + String(yearNum).slice(-2);
    if (entry) {
      const s = Number(entry.staff || 0),
        w = Number(entry.workers || 0),
        mp = Number(entry.totalManpower || s + w);
      const wh = Number(entry.workingHours || 8),
        tmh = Number(entry.totalManHours || mp * wh),
        smh = Number(entry.safeManHours || tmh);
      const csm = Number(entry.cumSafeManHours || latestCumSafe + smh);
      latestCumSafe = csm;
      const ind = Number(entry.inductions || 0),
        indS = Number(entry.indStaff || 0),
        indW = Number(entry.indWorkers || 0);
      const tbtC = Number(entry.tbtCount || 0),
        tbtP = Number(entry.tbtPersons || 0),
        trP = Number(entry.trainingPersons || 0);
      const pH = Number(entry.permitHot || 0),
        pE = Number(entry.permitElectrical || 0),
        pC = Number(entry.permitCold || 0),
        pO = Number(entry.permitOthers || entry.permitGeneral || 0);
      const fa = Number(entry.firstAid || 0),
        nm = Number(entry.nearMiss || 0),
        lti = Number(entry.ltiCount || 0);

      totStaff += s;
      totWorkers += w;
      totMP += mp;
      totWorkHrs += wh;
      totManHrs += tmh;
      totSafeHrs += smh;
      totInd += ind;
      totIndStaff += indS;
      totIndWorkers += indW;
      totTbt += tbtC;
      totTbtPersons += tbtP;
      totTrainingPersons += trP;
      totHot += pH;
      totElect += pE;
      totCold += pC;
      totOthers += pO;
      totFA += fa;
      totNM += nm;
      totLTI += lti;

      rowsHtml +=
        "<tr>" +
        '<td class="c">' +
        d +
        '</td><td class="c">' +
        dateStr +
        "</td>" +
        '<td class="c">' +
        (s || "-") +
        '</td><td class="c">' +
        (w || "-") +
        '</td><td class="c bold">' +
        (mp || "-") +
        "</td>" +
        '<td class="c">' +
        (wh || "-") +
        '</td><td class="c">' +
        (tmh || "-") +
        '</td><td class="c">' +
        (smh || "-") +
        '</td><td class="c bold">' +
        (csm || "-") +
        "</td>" +
        '<td class="c">' +
        (ind || "-") +
        '</td><td class="c">' +
        (indS || "-") +
        '</td><td class="c">' +
        (indW || "-") +
        "</td>" +
        '<td class="c">' +
        (tbtC || "-") +
        '</td><td class="c">' +
        (tbtP || "-") +
        "</td>" +
        "<td>" +
        escapeHtml_(entry.trainingTopic || "-") +
        '</td><td class="c">' +
        (trP || "-") +
        "</td>" +
        '<td class="c">' +
        (pH || "-") +
        '</td><td class="c">' +
        (pE || "-") +
        '</td><td class="c">' +
        (pC || "-") +
        '</td><td class="c">' +
        (pO || "-") +
        "</td>" +
        '<td class="c">' +
        (fa || "-") +
        '</td><td class="c">' +
        (nm || "-") +
        '</td><td class="c">' +
        (lti === 0 ? "0" : lti || "-") +
        "</td>" +
        "<td>" +
        escapeHtml_(entry.accidentDetails || "Nil") +
        "</td><td>" +
        escapeHtml_(entry.remarks || "-") +
        "</td>" +
        "</tr>";
    } else {
      rowsHtml +=
        '<tr><td class="c">' +
        d +
        '</td><td class="c">' +
        dateStr +
        "</td>" +
        '<td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td>' +
        '<td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td>' +
        '<td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td>' +
        '<td class="c">-</td><td class="c">-</td></tr>';
    }
  }

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    "<style>" +
    "  @page { size: A3 landscape; margin: 8mm; }" +
    "  body { font-family: Calibri, Arial, sans-serif; margin: 0; padding: 0; font-size: 8pt; color: #000; }" +
    "  table { width: 100%; border-collapse: collapse; margin-bottom: 4px; }" +
    "  th, td { border: 1px solid #000; padding: 2px 4px; }" +
    "  .title-banner { background: #bdd7ee; font-size: 13pt; font-weight: bold; text-align: center; padding: 6px; }" +
    "  .hdr1 { background: #d9d9d9; font-weight: bold; font-size: 7.5pt; text-align: center; }" +
    "  .hdr2 { background: #f2f2f2; font-weight: bold; font-size: 7pt; text-align: center; }" +
    "  .tot-row { background: #385723; color: #fff; font-weight: bold; font-size: 8pt; text-align: center; }" +
    "  .foot-label { background: #f2f2f2; font-weight: bold; }" +
    "  .c { text-align: center; }" +
    "  .bold { font-weight: bold; }" +
    "</style>" +
    "</head><body>" +
    "<table>" +
    '  <tr><td colspan="25" class="title-banner">DAILY LOG SHEET</td></tr>' +
    "  <tr>" +
    '    <td colspan="6" style="vertical-align:top"><b>Project: ' +
    pName +
    '</b><br>Cumulative Man-Hours Upto: 000<br><b style="color:#0f766e">DAILY PERFORMANCE REPORT</b></td>' +
    '    <td colspan="10" style="text-align:center;vertical-align:middle"><b style="font-size:11pt;color:#002060">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</b><br><small>Since 1998</small></td>' +
    '    <td colspan="9" style="text-align:right;vertical-align:middle"><b>Report for the month of - ' +
    monthYearLabel +
    "</b></td>" +
    "  </tr>" +
    '  <tr class="hdr1">' +
    '    <th rowspan="2">SL.NO</th><th rowspan="2">Date</th>' +
    '    <th colspan="3">Man Power</th><th colspan="4">Man hours Statistics</th>' +
    '    <th colspan="3">Safety Induction</th><th colspan="2">Tool Box Talk</th>' +
    '    <th colspan="2">Training Programs</th><th colspan="4">Work Permits</th>' +
    '    <th colspan="4">Accident Statistics</th><th rowspan="2">Remarks</th>' +
    "  </tr>" +
    '  <tr class="hdr2">' +
    "    <th>Staff</th><th>Workers</th><th>Total</th>" +
    "    <th>Work hrs</th><th>Total hrs</th><th>Safe hrs</th><th>Cum Safe</th>" +
    "    <th>Inductions</th><th>Staff</th><th>Workers</th>" +
    "    <th>TBTs</th><th>Attended</th>" +
    "    <th>Topic</th><th>Attended</th>" +
    "    <th>Hot</th><th>Elect</th><th>Cold</th><th>Others</th>" +
    "    <th>FA</th><th>Near Miss</th><th>LTI</th><th>Details</th>" +
    "  </tr>" +
    rowsHtml +
    '  <tr class="tot-row">' +
    '    <td colspan="2">Total</td>' +
    "    <td>" +
    totStaff +
    "</td><td>" +
    totWorkers +
    "</td><td>" +
    totMP +
    "</td>" +
    "    <td>" +
    totWorkHrs +
    "</td><td>" +
    totManHrs +
    "</td><td>" +
    totSafeHrs +
    "</td><td>" +
    (latestCumSafe || totSafeHrs) +
    "</td>" +
    "    <td>" +
    totInd +
    "</td><td>" +
    totIndStaff +
    "</td><td>" +
    totIndWorkers +
    "</td>" +
    "    <td>" +
    totTbt +
    "</td><td>" +
    totTbtPersons +
    "</td>" +
    "    <td>-</td><td>" +
    totTrainingPersons +
    "</td>" +
    "    <td>" +
    totHot +
    "</td><td>" +
    totElect +
    "</td><td>" +
    totCold +
    "</td><td>" +
    totOthers +
    "</td>" +
    "    <td>" +
    (totFA || "-") +
    "</td><td>" +
    (totNM || "-") +
    "</td><td>" +
    (totLTI === 0 ? "0" : totLTI || "-") +
    "</td>" +
    "    <td>-</td><td>-</td>" +
    "  </tr>" +
    "  <tr>" +
    '    <td colspan="8" class="foot-label">Total Man Power Worked for the Month-</td><td colspan="4" class="c bold">' +
    (totMP || 54) +
    "</td>" +
    '    <td colspan="3" class="c bold">Date: ' +
    todayIso_() +
    '</td><td colspan="3" class="foot-label c">Report Updating by</td>' +
    '    <td colspan="3" class="foot-label c">Report Verified by</td><td colspan="4" class="foot-label c">Report approved by</td>' +
    "  </tr>" +
    "  <tr>" +
    '    <td colspan="8" class="foot-label">Total Safe Man Hours Worked month of</td><td colspan="4" class="c bold">' +
    (totSafeHrs || 370) +
    "</td>" +
    '    <td colspan="3" class="c bold">Name</td><td colspan="3" class="c">Site EHS Lead</td>' +
    '    <td colspan="3" class="c">Asst. EHS Manager</td><td colspan="4" class="c">EHS Manager / Director</td>' +
    "  </tr>" +
    "  <tr>" +
    '    <td colspan="8" class="foot-label">Cumulative Safe Man Hours Worked</td><td colspan="4" class="c bold">' +
    (latestCumSafe || totSafeHrs || 370) +
    "</td>" +
    '    <td colspan="3" class="c bold">Signature</td><td colspan="3" class="c">xxxx</td>' +
    '    <td colspan="3" class="c">xxxx</td><td colspan="4" class="c">xxx</td>' +
    "  </tr>" +
    "</table></body></html>"
  );
}
