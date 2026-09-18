function generateSubmissionPdf_(project, def, fields, user, version) {
  const isPermit = def.formCode && def.formCode.indexOf('WP_') === 0;
  const html = isPermit
    ? buildWorkPermitPdfHtml_(project, def, fields, user, version)
    : buildFormPdfHtml_(project, def, fields, user, version);
  const folder = getNamedSubfolder_(project, def.module ? moduleTitle_(def.module) : 'Generated PDFs');
  const name = project.code + '-' + def.formCode + '-v' + version + '-' + todayIso_();
  const pdfRes = htmlToPdfFile_(html, name, folder);
  return { fileId: pdfRes.fileId, docId: pdfRes.docId, url: pdfRes.url };
}

function moduleTitle_(id) {
  const m = MODULES.find(x => x.id === id);
  return m ? m.title : 'Generated PDFs';
}

function htmlToPdfFile_(html, name, folder) {
  const doc = DocumentApp.create(name);
  const body = doc.getBody();
  body.clear();
  const tmp = html.replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<\/tr>/gi, '\n').replace(/<[^>]+>/g, ' ');
  const text = tmp.replace(/\s+\n/g, '\n').replace(/[ \t]+/g, ' ').trim();
  body.appendParagraph(APP_NAME + ' — Safety Management System').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(text.substring(0, 45000));
  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF).setName(name + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);
  return { fileId: pdfFile.getId(), docId: doc.getId(), url: pdfFile.getUrl() };
}

function buildFormPdfHtml_(project, def, fields, user, version) {
  const rows = Object.keys(fields).map(k => {
    return '<tr><td style="padding:7px 10px;border:1px solid #d0d7de;width:34%;background:#f8fafc;font-weight:600;color:#1e293b">' + escapeHtml_(prettyLabel_(k)) + '</td>' +
      '<td style="padding:7px 10px;border:1px solid #d0d7de;color:#0f172a">' + escapeHtml_(displayDate_(fields[k])) + '</td></tr>';
  }).join('');
  return '<html><body style="font-family:Arial,sans-serif;color:#1e293b;padding:20px;line-height:1.5">' +
    '<div style="border-bottom:2px solid #147d6f;padding-bottom:12px;margin-bottom:16px">' +
    '  <h2 style="margin:0;color:#147d6f;font-size:20px;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</h2>' +
    '  <p style="margin:4px 0 0;font-size:12px;color:#64748b;font-weight:bold;letter-spacing:1px">ENVIRONMENT, HEALTH & SAFETY MANAGEMENT SYSTEM</p>' +
    '</div>' +
    '<div style="background:#f1f5f9;border-left:4px solid #147d6f;padding:10px 14px;margin-bottom:16px">' +
    '  <h3 style="margin:0;font-size:16px;color:#0f172a">' + escapeHtml_(def.title) + ' <span style="font-size:12px;color:#64748b;font-weight:normal">[' + escapeHtml_(def.formCode) + ']</span></h3>' +
    '  <p style="margin:4px 0 0;font-size:12px;color:#475569">' +
    '    Project: <b>' + escapeHtml_(project.name) + ' (' + escapeHtml_(project.code) + ')</b> &bull; ' +
    '    Client: <b>' + escapeHtml_(project.client || '—') + '</b> &bull; ' +
    '    PMC: <b>' + escapeHtml_(project.pmc || '—') + '</b> &bull; ' +
    '    Version: <b>v' + version + '</b> &bull; ' +
    '    Date: <b>' + displayDate_(nowIso_()) + '</b><br>' +
    '    Submitted By: <b>' + escapeHtml_(user.name) + ' (' + escapeHtml_(user.employeeId) + ')</b>' +
    '  </p>' +
    '</div>' +
    '<table style="border-collapse:collapse;width:100%;font-size:12px;margin-bottom:24px">' + rows + '</table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:30px;font-size:11px;border-top:1px solid #cbd5e1;padding-top:12px">' +
    '  <tr>' +
    '    <td style="width:33%;padding:10px;vertical-align:top;border:1px solid #e2e8f0">' +
    '      <b>Prepared By:</b><br>' + escapeHtml_(user.name) + '<br><small style="color:#64748b">Site EHS Lead</small>' +
    '    </td>' +
    '    <td style="width:33%;padding:10px;vertical-align:top;border:1px solid #e2e8f0">' +
    '      <b>Verified By:</b><br>EHS Inspection Authority<br><small style="color:#64748b">Asst. EHS Manager</small>' +
    '    </td>' +
    '    <td style="width:33%;padding:10px;vertical-align:top;border:1px solid #e2e8f0">' +
    '      <b>Approved By:</b><br>SESIPL Corporate EHS<br><small style="color:#64748b">EHS Manager / Director</small>' +
    '    </td>' +
    '  </tr>' +
    '</table>' +
    '<p style="margin-top:18px;font-size:10px;color:#94a3b8;text-align:center">' +
    '  This record is digitally captured & stored in the SESIPL EHS cloud archive. Any alteration invalidates this record.' +
    '</p>' +
    '</body></html>';
}

function buildWorkPermitPdfHtml_(project, def, fields, user, version) {
  const docCode = (typeof PERMIT_DOC_CODES !== 'undefined' && PERMIT_DOC_CODES[def.formCode])
    ? PERMIT_DOC_CODES[def.formCode]
    : 'SESIPL-EHS.Blr PTW';
  const contractor = fields.contractorName || fields.contractor || project.client || 'Shankar Electricals Services (I) Pvt Ltd';
  const em1 = fields.emergencyContact1 || '—';
  const em2 = fields.emergencyContact2 || '—';
  const pNo = fields.permitNo || (project.code + '-' + def.formCode + '-' + (fields.date || todayIso_()).replace(/-/g, ''));
  const area = fields.area || '—';
  const loc = fields.location || '—';
  const dt = displayDate_(fields.date || nowIso_());
  const tm = fields.time || '—';
  const siteEng = fields.siteEngineer || user.name || '—';
  const siteEngSign = fields.siteEngineerSign || siteEng;
  const safetyOff = fields.safetyOfficer || '—';
  const safetyOffSign = fields.safetyOfficerSign || safetyOff;
  const contIncharge = fields.contractorInCharge || '—';
  const contPhone = fields.contactNumber || '—';
  const workDesc = fields.workDescription || '—';
  const execDate = displayDate_(fields.workExecutionDate || fields.date || nowIso_());
  const validFrom = fields.validFrom ? displayDate_(fields.validFrom).replace('T', ' ') : '—';
  const validTo = fields.validTo ? displayDate_(fields.validTo).replace('T', ' ') : '—';

  let precautions = [];
  const isNight = def.formCode === 'WP_NIGHT';

  if (def.formCode === 'WP_SHAFT') {
    precautions = [
      [1, "Proper Access/ Exit available", fields.shaft_q1],
      [2, "Proper ventilation and / or lighting provided", fields.shaft_q2],
      [3, "Proper & Safe platform provided", fields.shaft_q3],
      [4, "Workers have been briefed about hazardous", fields.shaft_q4],
      [5, "All Electrical Tools and machinery checked prior to use.", fields.shaft_q5],
      [6, "Shaft area Properly barricaded.", fields.shaft_q6],
      [7, "Conducted JST for all workers who are all engaging to shaft work.", fields.shaft_q7]
    ];
  } else if (def.formCode === 'WP_NIGHT') {
    precautions = [
      [1, "Is dedicated Night shift in charge available?", fields.night_q1],
      [2, "Is supervisor available in night shift to supervise the task?", fields.night_q2],
      [3, "Is First aider available?", fields.night_q3],
      [4, "Is Ambulance available for emergency?", fields.night_q4],
      [5, "Are the workers working continuously for last 12 hours?", fields.night_q5],
      [6, "Is the work area safe for work?", fields.night_q6],
      [7, "Is there proper illumination provided at the work area?", fields.night_q7],
      [8, "Are the hazards related with the work identified and assessed at workplace?", fields.night_q8],
      [9, "Is toolbox talk / pre-start briefing carried out prior to start night shift?", fields.night_q9],
      [10, "Are the workers having specific PPE’s according to the requirement of the task?", fields.night_q10],
      [11, "Is there any high-risk activity like working at height, work in penetration and shafts, electrical testing and commissioning, hot work, Mechanical lifting operation, excavation etc. to be carried out in night shift?", fields.night_q11]
    ];
  } else if (def.formCode === 'WP_LIFT') {
    precautions = [
      [1, "Crane used for lifting activity tested, certified and approved for rated lifting works.", fields.lift_q1],
      [2, "All lifting tackles, gears/ appliances are tested and certified for lifting works.", fields.lift_q2],
      [3, "Crane operator is trained and competent for lifting operation.", fields.lift_q3],
      [4, "Lifting belt protected against sharp edge of jobs to be lifted.", fields.lift_q4],
      [5, "Access and exist marked and without obstruction.", fields.lift_q5],
      [6, "Lighting arrangement adequate.", fields.lift_q6],
      [7, "Unwanted and rubbish material removed from working platform.", fields.lift_q7],
      [8, "Guidelines has provided for balancing & guiding jobs to be lifted.", fields.lift_q8],
      [9, "Periphery area of crane booms as well lifting job is barricaded .", fields.lift_q9],
      [10, "Rigger and signal man is trained and competent for lifting work.", fields.lift_q10],
      [11, "No lifting activity to be carried during lightening, heavy wind /rain.", fields.lift_q11],
      [12, "If scaffolding to be used during lift , Scaffolding with valid tag available for use", fields.lift_q12],
      [13, "Double lanyards Safety Harness/belt checked and in working condition", fields.lift_q13],
      [14, "Safety shoes (nonslip), Helmet with chin strip available with employees.", fields.lift_q14],
      [15, "Other: " + (fields.lift_other || "—"), fields.lift_other ? 'Yes' : 'Not Required']
    ];
  } else if (def.formCode === 'WP_HOT') {
    precautions = [
      [1, "Proper Access/ Exit available", fields.hot_q1],
      [2, "Proper ventilation and / or lighting provided", fields.hot_q2],
      [3, "Proper & Safe scaffolding, platform, ladder provided", fields.hot_q3],
      [4, "Welding machine located in a clean and dry area", fields.hot_q4],
      [5, "Welding machine grounded at the equipment & proper leakage current protection device (ELCB) provided for welding machine.", fields.hot_q5],
      [6, "Competent and Trained personnel deployed to carry the work.", fields.hot_q6],
      [7, "Welding machine, Input / Output Cables, welding holder and weld return clamp (Holder ) insulated & in good condition", fields.hot_q7],
      [8, "Welder and fitter trained to connect ground / work return clamps (Holder) to the work piece prior to energization of Welding machine.", fields.hot_q8],
      [9, "Gas Cylinders stacked vertically and not below the welding/cutting area. Regulator Key is available with cylinders.", fields.hot_q9],
      [10, "Work Area Isolated with barricading and caution sign", fields.hot_q10],
      [11, "Personal Protective Equipment. Minimum applicable - Safety helmet, safety goggles, welding helmet, safety shoes, leather gloves, long sleeve and nose mask provided.", fields.hot_q11],
      [12, "In case of pits, water removed from the pit & wood /rubber insulation provided.", fields.hot_q12],
      [13, "Adequate & suitable nos. of fire fighting extinguisher provided.", fields.hot_q13],
      [14, "Near by combustible material removed. Housekeeping Done.", fields.hot_q14],
      [15, "Fire watch as standby is in place.", fields.hot_q15],
      [16, "Other: " + (fields.hot_other || "—"), fields.hot_other ? 'Yes' : 'Not Required']
    ];
  } else if (def.formCode === 'WP_HEIGHT') {
    precautions = [
      [1, "Scaffolding with valid tag available for use", fields.height_q1],
      [2, "Conducted JST/TBT conducted", fields.height_q2],
      [3, "Safety shoes (nonslip), Helmet with chin strip available with employees.", fields.height_q3],
      [4, "All tightening tools, hand tools /equipment checked and in good condition.", fields.height_q4],
      [5, "Access and exist marked and without obstruction.", fields.height_q5],
      [6, "Lighting arrangement adequate.", fields.height_q6],
      [7, "Unwanted and rubbish material removed from working platform.", fields.height_q7],
      [8, "Electrical cable in good condition", fields.height_q8],
      [9, "Signboards provided", fields.height_q9],
      [10, "Employees aware about hazards and safe working practices while working at height.", fields.height_q10]
    ];
  }

  let tableHeader = '';
  let tableRows = '';

  if (isNight) {
    tableHeader = '<tr><th style="width:6%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">No</th>' +
      '<th style="width:58%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:left">ITEM</th>' +
      '<th style="width:8%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">Yes</th>' +
      '<th style="width:8%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">NO</th>' +
      '<th style="width:8%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">NA</th>' +
      '<th style="width:12%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">REMARKS</th></tr>';

    tableRows = precautions.map(p => {
      const v = String(p[2] || '').toUpperCase();
      return '<tr>' +
        '<td style="text-align:center;padding:5px;border:1px solid #334155">' + p[0] + '</td>' +
        '<td style="padding:5px;border:1px solid #334155">' + escapeHtml_(p[1]) + '</td>' +
        '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' + (v === 'YES' ? '✓' : '') + '</td>' +
        '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' + (v === 'NO' ? '✓' : '') + '</td>' +
        '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' + (v === 'NA' ? '✓' : '') + '</td>' +
        '<td style="padding:5px;border:1px solid #334155">' + (p[0] === 11 ? escapeHtml_(fields.night_remarks || '') : '') + '</td>' +
      '</tr>';
    }).join('');
  } else {
    tableHeader = '<tr><th style="width:6%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">No</th>' +
      '<th style="width:72%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:left">Item</th>' +
      '<th style="width:11%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">Yes</th>' +
      '<th style="width:11%;padding:6px;border:1.5px solid #0f172a;background:#f1f5f9;text-align:center">Not Required</th></tr>';

    tableRows = precautions.map(p => {
      const v = String(p[2] || '').toLowerCase();
      const isYes = v.includes('yes');
      const isNot = v.includes('not');
      return '<tr>' +
        '<td style="text-align:center;padding:5px;border:1px solid #334155">' + p[0] + '</td>' +
        '<td style="padding:5px;border:1px solid #334155">' + escapeHtml_(p[1]) + '</td>' +
        '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' + (isYes ? '✓' : '') + '</td>' +
        '<td style="text-align:center;padding:5px;border:1px solid #334155;font-weight:bold">' + (isNot ? '✓' : '') + '</td>' +
      '</tr>';
    }).join('');
  }

  const workmenBlock = (def.formCode === 'WP_SHAFT' && fields.workmenNames) ?
    '<div style="margin-top:10px;font-size:11.5px"><b>Names of workmen entering shaft:</b> ' + escapeHtml_(fields.workmenNames) + '</div>' : '';

  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#0f172a;padding:24px;line-height:1.4;font-size:11px">' +
    '<table style="width:100%;border-collapse:collapse;border:1.5px solid #0f172a;margin-bottom:14px">' +
    '  <tr>' +
    '    <td style="width:34%;padding:8px 12px;border-right:1.5px solid #0f172a;vertical-align:middle">' +
    '      <div style="font-size:10px;font-weight:bold;color:#475569">Contractor Name:</div>' +
    '      <div style="font-size:12px;font-weight:bold;color:#0f172a;margin-top:2px">' + escapeHtml_(contractor) + '</div>' +
    '      <div style="margin-top:6px;font-size:10.5px;font-weight:bold;color:#0f766e">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</div>' +
    '    </td>' +
    '    <td style="width:34%;padding:8px 12px;border-right:1.5px solid #0f172a;text-align:center;vertical-align:middle">' +
    '      <div style="font-size:11px;font-weight:bold;color:#475569">TITLE :</div>' +
    '      <div style="font-size:13px;font-weight:800;letter-spacing:0.5px;color:#0f172a;margin-top:3px">SAFETY WORK CLEARANCE</div>' +
    '    </td>' +
    '    <td style="width:32%;padding:8px 12px;vertical-align:middle;font-size:10.5px">' +
    '      <div><b>Emergency Contact No:</b></div>' +
    '      <div>1) ' + escapeHtml_(em1) + '</div>' +
    '      <div>2) ' + escapeHtml_(em2) + '</div>' +
    '      <div style="margin-top:4px"><b>Permit No:-</b> <span style="font-weight:bold;color:#0f766e">' + escapeHtml_(pNo) + '</span></div>' +
    '    </td>' +
    '  </tr>' +
    '</table>' +

    '<div style="text-align:center;margin-bottom:12px">' +
    '  <h1 style="margin:0;font-size:17px;font-weight:900;letter-spacing:0.5px;text-transform:uppercase;color:#0f172a">' + escapeHtml_(def.title) + '</h1>' +
    '</div>' +

    '<div style="border:1px solid #cbd5e1;background:#f8fafc;padding:8px 12px;border-radius:6px;margin-bottom:10px;font-size:11px">' +
    '  <table style="width:100%;border-collapse:collapse">' +
    '    <tr>' +
    '      <td style="width:40%;padding:3px 0"><b>Area:-</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:140px">' + escapeHtml_(area) + '</span></td>' +
    '      <td style="width:35%;padding:3px 0"><b>location:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:130px">' + escapeHtml_(loc) + '</span></td>' +
    '      <td style="width:25%;padding:3px 0"><b>Date:</b> ' + escapeHtml_(dt) + ' &nbsp; <b>Time:-</b> ' + escapeHtml_(tm) + '</td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td colspan="2" style="padding:3px 0"><b>Name of Site Engineer (Permit Requesting Authority):</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:180px">' + escapeHtml_(siteEng) + '</span></td>' +
    '      <td style="padding:3px 0"><b>Sign:</b> <span style="font-style:italic;color:#0f766e">' + escapeHtml_(siteEngSign) + '</span></td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td colspan="2" style="padding:3px 0"><b>Name of Safety Officer:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:220px">' + escapeHtml_(safetyOff) + '</span></td>' +
    '      <td style="padding:3px 0"><b>Sign:</b> <span style="font-style:italic;color:#0f766e">' + escapeHtml_(safetyOffSign) + '</span></td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td colspan="2" style="padding:3px 0"><b>Name of Contractor Site In charge:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:180px">' + escapeHtml_(contIncharge) + '</span></td>' +
    '      <td style="padding:3px 0"><b>Contact Number:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;min-width:110px">' + escapeHtml_(contPhone) + '</span></td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td colspan="3" style="padding:3px 0"><b>Description of work:</b> <span style="border-bottom:1px solid #0f172a;display:inline-block;width:80%">' + escapeHtml_(workDesc) + '</span></td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td colspan="3" style="padding:3px 0"><b>Work Execution Date:</b> ' + escapeHtml_(execDate) + ' &nbsp;&nbsp;&nbsp;&nbsp; <b>Valid From:-</b> ' + escapeHtml_(validFrom) + ' &nbsp;&nbsp;&nbsp;&nbsp; <b>To:-</b> ' + escapeHtml_(validTo) + '</td>' +
    '    </tr>' +
    '  </table>' +
    '</div>' +

    '<p style="font-size:10px;font-style:italic;margin:4px 0 8px;color:#334155">' +
    '  The above signing person will be responsible to ensure that the above described work will be done under all the safety precaution mentioned on the PTW and required by the Project.' +
    '</p>' +

    '<div style="font-size:11px;font-weight:bold;margin-bottom:5px">The following precautions are to be taken:-</div>' +

    '<table style="width:100%;border-collapse:collapse;border:1.5px solid #0f172a;font-size:10.5px;margin-bottom:8px">' +
    tableHeader +
    tableRows +
    '</table>' +

    workmenBlock +

    '<div style="margin-top:6px;font-size:10px;color:#475569"><b>Notes:</b> a) Work permit is valid for the prescribed date, time and in prescribed location only</div>' +

    '<div style="margin-top:10px;border:1px solid #cbd5e1;padding:6px 10px;border-radius:6px;background:#f8fafc">' +
    '  <div style="font-weight:bold;font-size:11.5px;margin-bottom:4px;color:#0f172a">Reviewed & Approved By (Permit Issuing Authority):</div>' +
    '  <table style="width:100%;border-collapse:collapse;font-size:10.5px">' +
    '    <tr>' +
    '      <td style="width:50%;padding:3px 0"><b>EHS:</b> ' + escapeHtml_(fields.approvalEhsName || '—') + ' &nbsp;&nbsp; <b>Sign:</b> <span style="font-style:italic;color:#0f766e">' + escapeHtml_(fields.approvalEhsSign || fields.approvalEhsName || 'Acknowledged') + '</span></td>' +
    '      <td style="width:50%;padding:3px 0"><b>Date:</b> ' + escapeHtml_(displayDate_(fields.approvalEhsDate || fields.date)) + ' &nbsp;&nbsp; <b>Time:</b> ' + escapeHtml_(fields.approvalEhsTime || fields.time || '—') + '</td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td style="width:50%;padding:3px 0"><b>Site Engineer:</b> ' + escapeHtml_(fields.approvalSiteEngineerName || '—') + ' &nbsp;&nbsp; <b>Sign:</b> <span style="font-style:italic;color:#0f766e">' + escapeHtml_(fields.approvalSiteEngineerSign || fields.approvalSiteEngineerName || 'Verified') + '</span></td>' +
    '      <td style="width:50%;padding:3px 0"><b>Date:</b> ' + escapeHtml_(displayDate_(fields.approvalSiteEngineerDate || fields.date)) + ' &nbsp;&nbsp; <b>Time:</b> ' + escapeHtml_(fields.approvalSiteEngineerTime || fields.time || '—') + '</td>' +
    '    </tr>' +
    '  </table>' +
    '  <div style="font-size:9.5px;font-style:italic;color:#475569;margin-top:4px">' +
    '    I understand the precaution to be taken as described above and as per Project requirement & hereby confirm that Work will be executed under my supervision by following all precaution & Safety Rules.' +
    '  </div>' +
    '</div>' +

    '<div style="margin-top:10px;border:1px solid #cbd5e1;padding:6px 10px;border-radius:6px">' +
    '  <div style="font-weight:bold;font-size:11.5px;margin-bottom:3px;color:#991b1b">Permit Closing / Cancellation:-</div>' +
    '  <div style="font-size:9.5px;font-style:italic;color:#475569;margin-bottom:6px">' +
    '    I hereby declare that the work is completed / suspended, all workers under my control have been withdrawn and the site restored to a safe tidy condition.' +
    '  </div>' +
    '  <table style="width:100%;border-collapse:collapse;font-size:10.5px">' +
    '    <tr>' +
    '      <td style="width:55%;padding:2px 0"><b>Name of Site Engineer (Permit Requesting Authority):</b> ' + escapeHtml_(fields.closingSiteEngName || '—') + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Sign:</b> ' + escapeHtml_(fields.closingSiteEngSign || '—') + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Date:</b> ' + escapeHtml_(displayDate_(fields.closingSiteEngDate || '')) + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Time:</b> ' + escapeHtml_(fields.closingSiteEngTime || '—') + '</td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td style="width:55%;padding:2px 0"><b>Name of Safety Officer:</b> ' + escapeHtml_(fields.closingSafetyOfficerName || '—') + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Sign:</b> ' + escapeHtml_(fields.closingSafetyOfficerSign || '—') + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Date:</b> ' + escapeHtml_(displayDate_(fields.closingSafetyOfficerDate || '')) + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Time:</b> ' + escapeHtml_(fields.closingSafetyOfficerTime || '—') + '</td>' +
    '    </tr>' +
    '    <tr>' +
    '      <td style="width:55%;padding:2px 0"><b>Name of PMC Site Engineer (Permit Issuing Authority):</b> ' + escapeHtml_(fields.closingPmcSiteEngName || '—') + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Sign:</b> ' + escapeHtml_(fields.closingPmcSiteEngSign || '—') + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Date:</b> ' + escapeHtml_(displayDate_(fields.closingPmcSiteEngDate || '')) + '</td>' +
    '      <td style="width:15%;padding:2px 0"><b>Time:</b> ' + escapeHtml_(fields.closingPmcSiteEngTime || '—') + '</td>' +
    '    </tr>' +
    '  </table>' +
    '</div>' +

    '<table style="width:100%;margin-top:14px;font-size:10px;color:#64748b">' +
    '  <tr>' +
    '    <td style="text-align:left"><b>Controlled Copy©</b></td>' +
    '    <td style="text-align:right"><b>Shankar Electricals Services I Pvt Ltd.</b><br><span style="font-family:monospace;font-size:10px;font-weight:bold">' + escapeHtml_(docCode) + '</span></td>' +
    '  </tr>' +
    '</table>' +

    '</body></html>';
}

function buildReportHtml_(project, subs, user) {
  const rows = subs.map(s => '<tr><td>' + escapeHtml_(s.formCode) + '</td><td>' + escapeHtml_(s.status) + '</td><td>' + escapeHtml_(s.submittedBy) + '</td><td>' + escapeHtml_(displayDate_(s.submittedAt)) + '</td></tr>').join('');
  return '<html><body style="font-family:Arial">' +
    '<h2>' + escapeHtml_(project.name) + ' — EHS Pack</h2>' +
    '<p>Exported by ' + escapeHtml_(user.name) + ' on ' + displayDate_(nowIso_()) + '</p>' +
    '<table border="1" cellpadding="6"><tr><th>Form</th><th>Status</th><th>By</th><th>At</th></tr>' + rows + '</table></body></html>';
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function prettyLabel_(k) {
  return String(k).replace(/([A-Z])/g, ' $1').replace(/^./, c => c.toUpperCase());
}
