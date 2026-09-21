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
  try {
    const blob = Utilities.newBlob(html, 'text/html', name + '.html');
    const pdfBlob = blob.getAs('application/pdf').setName(name + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("Notice: setSharing ANYONE_WITH_LINK: " + shareErr);
    }
    return { fileId: pdfFile.getId(), docId: '', url: pdfFile.getUrl() };
  } catch (err) {
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
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("Notice: setSharing ANYONE_WITH_LINK: " + shareErr);
    }
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
    return { fileId: pdfFile.getId(), docId: doc.getId(), url: pdfFile.getUrl() };
  }
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

function buildWorkPermitPdfHtml_(project, def, fields, user, version, options) {
  const isTagMode = options && options.usePlaceholders;
  const f = fields || {};
  const formCode = (def && def.formCode) || 'WP_HEIGHT';
  const docCode = (typeof PERMIT_DOC_CODES !== 'undefined' && PERMIT_DOC_CODES[formCode])
    ? PERMIT_DOC_CODES[formCode]
    : 'SESIPL-EHS.Blr PTW';

  const contractor = f.contractorName || f.contractor || '';
  const em1 = f.emergencyContact1 || '';
  const em2 = f.emergencyContact2 || '';
  const pNo = f.permitNo || '';
  const area = f.area || '';
  const loc = f.location || '';
  const dt = f.date ? displayDate_(f.date) : '';
  const tm = f.time || '';
  const siteEng = f.siteEngineer || '';
  const siteEngSign = f.siteEngineerSign || (f.siteEngineer ? 'Signed' : '');
  const safetyOff = f.safetyOfficer || '';
  const safetyOffSign = f.safetyOfficerSign || (f.safetyOfficer ? 'Signed' : '');
  const contIncharge = f.contractorInCharge || '';
  const contPhone = f.contactNumber || '';
  const workDesc = f.workDescription || '';
  const execDate = f.workExecutionDate ? displayDate_(f.workExecutionDate) : (f.date ? displayDate_(f.date) : '');
  const validFrom = f.validFrom ? displayDate_(f.validFrom).replace('T', ' ') : '';
  const validTo = f.validTo ? displayDate_(f.validTo).replace('T', ' ') : '';

  function renderLine_(val, placeholderTag, minWidth, fullWidth) {
    if (val && String(val).trim() && String(val).trim() !== '—') {
      return '<span style="border-bottom:1.5px solid #000;font-weight:bold;padding:0 6px;color:#0f172a;display:inline-block;' + (fullWidth ? 'width:calc(100% - 150px);vertical-align:bottom;' : ('min-width:' + minWidth + ';')) + '">' + escapeHtml_(val) + '</span>';
    }
    if (isTagMode && placeholderTag) {
      return '<span style="border-bottom:1.5px solid #000;font-family:monospace;color:#0369a1;font-weight:bold;padding:0 4px;display:inline-block;' + (fullWidth ? 'width:calc(100% - 150px);' : ('min-width:' + minWidth + ';')) + '">{{' + placeholderTag + '}}</span>';
    }
    return '<span style="border-bottom:1.5px solid #000;display:inline-block;' + (fullWidth ? 'width:calc(100% - 150px);' : ('min-width:' + minWidth + ';')) + '">&nbsp;</span>';
  }

  function renderDateLine_(val, placeholderTag) {
    if (val && String(val).trim() && String(val).trim() !== '—') {
      return '<span style="border-bottom:1.5px solid #000;font-weight:bold;padding:0 6px;color:#0f172a;display:inline-block;min-width:100px;text-align:center">' + escapeHtml_(val) + '</span>';
    }
    if (isTagMode && placeholderTag) {
      return '<span style="border-bottom:1.5px solid #000;font-family:monospace;color:#0369a1;font-weight:bold;padding:0 4px;display:inline-block;min-width:100px;text-align:center">{{' + placeholderTag + '}}</span>';
    }
    return '<span style="border-bottom:1.5px solid #000;display:inline-block;min-width:100px;text-align:center">&nbsp;&nbsp;&nbsp;&nbsp; / &nbsp;&nbsp;&nbsp;&nbsp; / &nbsp;&nbsp;&nbsp;&nbsp;</span>';
  }

  let precautions = [];
  const isNight = formCode === 'WP_NIGHT';

  if (formCode === 'WP_SHAFT') {
    precautions = [
      [1, "Proper Access/ Exit available", f.shaft_q1],
      [2, "Proper ventilation and / or lighting provided", f.shaft_q2],
      [3, "Proper & Safe platform provided", f.shaft_q3],
      [4, "Workers have been briefed about hazardous", f.shaft_q4],
      [5, "All Electrical Tools and machinery checked prior to use.", f.shaft_q5],
      [6, "Shaft area Properly barricaded.", f.shaft_q6],
      [7, "Conducted JST for all workers who are all engaging to shaft work.", f.shaft_q7]
    ];
  } else if (formCode === 'WP_NIGHT') {
    precautions = [
      [1, "Is dedicated Night shift in charge available?", f.night_q1],
      [2, "Is supervisor available in night shift to supervise the task?", f.night_q2],
      [3, "Is First aider available?", f.night_q3],
      [4, "Is Ambulance available for emergency?", f.night_q4],
      [5, "Are the workers working continuously for last 12 hours?", f.night_q5],
      [6, "Is the work area safe for work?", f.night_q6],
      [7, "Is there proper illumination provided at the work area?", f.night_q7],
      [8, "Are the hazards related with the work identified and assessed at workplace?", f.night_q8],
      [9, "Is toolbox talk / pre-start briefing carried out prior to start night shift?", f.night_q9],
      [10, "Are the workers having specific PPE’s according to the requirement of the task?", f.night_q10],
      [11, "Is there any high-risk activity like working at height, work in penetration and shafts, electrical testing and commissioning, hot work, Mechanical lifting operation, excavation etc. to be carried out in night shift?", f.night_q11]
    ];
  } else if (formCode === 'WP_LIFT') {
    precautions = [
      [1, "Crane used for lifting activity tested, certified and approved for rated lifting works.", f.lift_q1],
      [2, "All lifting tackles, gears/ appliances are tested and certified for lifting works.", f.lift_q2],
      [3, "Crane operator is trained and competent for lifting operation.", f.lift_q3],
      [4, "Lifting belt protected against sharp edge of jobs to be lifted.", f.lift_q4],
      [5, "Access and exist marked and without obstruction.", f.lift_q5],
      [6, "Lighting arrangement adequate.", f.lift_q6],
      [7, "Unwanted and rubbish material removed from working platform.", f.lift_q7],
      [8, "Guidelines has provided for balancing & guiding jobs to be lifted.", f.lift_q8],
      [9, "Periphery area of crane booms as well lifting job is barricaded .", f.lift_q9],
      [10, "Rigger and signal man is trained and competent for lifting work.", f.lift_q10],
      [11, "No lifting activity to be carried during lightening, heavy wind /rain.", f.lift_q11],
      [12, "If scaffolding to be used during lift , Scaffolding with valid tag available for use", f.lift_q12],
      [13, "Double lanyards Safety Harness/belt checked and in working condition", f.lift_q13],
      [14, "Safety shoes (nonslip), Helmet with chin strip available with employees.", f.lift_q14],
      [15, "Other: " + (f.lift_other || ""), f.lift_other ? 'Yes' : '']
    ];
  } else if (formCode === 'WP_HOT') {
    precautions = [
      [1, "Proper Access/ Exit available", f.hot_q1],
      [2, "Proper ventilation and / or lighting provided", f.hot_q2],
      [3, "Proper & Safe scaffolding, platform, ladder provided", f.hot_q3],
      [4, "Welding machine located in a clean and dry area", f.hot_q4],
      [5, "Welding machine grounded at the equipment & proper leakage current protection device (ELCB) provided for welding machine.", f.hot_q5],
      [6, "Competent and Trained personnel deployed to carry the work.", f.hot_q6],
      [7, "Welding machine, Input / Output Cables, welding holder and weld return clamp (Holder ) insulated & in good condition", f.hot_q7],
      [8, "Welder and fitter trained to connect ground / work return clamps (Holder) to the work piece prior to energization of Welding machine.", f.hot_q8],
      [9, "Gas Cylinders stacked vertically and not below the welding/cutting area. Regulator Key is available with cylinders.", f.hot_q9],
      [10, "Work Area Isolated with barricading and caution sign", f.hot_q10],
      [11, "Personal Protective Equipment. Minimum applicable - Safety helmet, safety goggles, welding helmet, safety shoes, leather gloves, long sleeve and nose mask provided.", f.hot_q11],
      [12, "In case of pits, water removed from the pit & wood /rubber insulation provided.", f.hot_q12],
      [13, "Adequate & suitable nos. of fire fighting extinguisher provided.", f.hot_q13],
      [14, "Near by combustible material removed. Housekeeping Done.", f.hot_q14],
      [15, "Fire watch as standby is in place.", f.hot_q15],
      [16, "Other: " + (f.hot_other || ""), f.hot_other ? 'Yes' : '']
    ];
  } else if (formCode === 'WP_HEIGHT') {
    precautions = [
      [1, "Scaffolding with valid tag available for use", f.height_q1],
      [2, "Conducted JST/TBT conducted", f.height_q2],
      [3, "Safety shoes (nonslip), Helmet with chin strip available with employees.", f.height_q3],
      [4, "All tightening tools, hand tools /equipment checked and in good condition.", f.height_q4],
      [5, "Access and exist marked and without obstruction.", f.height_q5],
      [6, "Lighting arrangement adequate.", f.height_q6],
      [7, "Unwanted and rubbish material removed from working platform.", f.height_q7],
      [8, "Electrical cable in good condition", f.height_q8],
      [9, "Signboards provided", f.height_q9],
      [10, "Employees aware about hazards and safe working practices while working at height.", f.height_q10],
      [11, "", ""],
      [12, "", ""],
      [13, "", ""],
      [14, "", ""]
    ];
  } else {
    // WP_GENERAL
    precautions = [
      [1, "Area inspected and free from obvious hazards", f.gen_q1],
      [2, "Personnel briefed, inducted and wearing required PPE", f.gen_q2],
      [3, "Hand tools, electrical equipment and machinery checked prior to use", f.gen_q3],
      [4, "Safe access, emergency exit and walkways marked without obstruction", f.gen_q4],
      [5, "Adequate ventilation and illumination provided at workplace", f.gen_q5],
      [6, "Housekeeping done & unwanted combustible materials removed", f.gen_q6],
      [7, "Caution signboards provided and emergency numbers posted", f.gen_q7],
      [8, "", ""],
      [9, "", ""],
      [10, "", ""]
    ];
  }

  let tableHeader = '';
  let tableRows = '';

  if (isNight) {
    tableHeader = '<tr>' +
      '<th style="width:6%;padding:4px 2px;border:1.5px solid #000;background:#f8fafc;text-align:center">No</th>' +
      '<th style="width:58%;padding:4px 8px;border:1.5px solid #000;background:#f8fafc;text-align:left">ITEM</th>' +
      '<th style="width:8%;padding:4px 2px;border:1.5px solid #000;background:#f8fafc;text-align:center">Yes</th>' +
      '<th style="width:8%;padding:4px 2px;border:1.5px solid #000;background:#f8fafc;text-align:center">NO</th>' +
      '<th style="width:8%;padding:4px 2px;border:1.5px solid #000;background:#f8fafc;text-align:center">NA</th>' +
      '<th style="width:12%;padding:4px 4px;border:1.5px solid #000;background:#f8fafc;text-align:center">REMARKS</th>' +
    '</tr>';

    tableRows = precautions.map(p => {
      const v = String(p[2] || '').trim().toUpperCase();
      const isYes = v === 'YES';
      const isNo = v === 'NO';
      const isNa = v === 'NA';
      return '<tr>' +
        '<td style="text-align:center;padding:4px 2px;border:1px solid #000;font-weight:bold">' + p[0] + '</td>' +
        '<td style="padding:4px 8px;border:1px solid #000">' + escapeHtml_(p[1]) + '</td>' +
        '<td style="text-align:center;padding:4px 2px;border:1px solid #000;font-weight:bold;font-size:12px">' + (isYes ? '✓' : (isTagMode ? '{{q' + p[0] + '_yes}}' : '')) + '</td>' +
        '<td style="text-align:center;padding:4px 2px;border:1px solid #000;font-weight:bold;font-size:12px">' + (isNo ? '✓' : (isTagMode ? '{{q' + p[0] + '_no}}' : '')) + '</td>' +
        '<td style="text-align:center;padding:4px 2px;border:1px solid #000;font-weight:bold;font-size:12px">' + (isNa ? '✓' : (isTagMode ? '{{q' + p[0] + '_na}}' : '')) + '</td>' +
        '<td style="padding:4px 6px;border:1px solid #000;font-size:10px">' + (p[0] === 11 ? escapeHtml_(f.night_remarks || '') : '') + '</td>' +
      '</tr>';
    }).join('');
  } else {
    tableHeader = '<tr>' +
      '<th style="width:6%;padding:4px 2px;border:1.5px solid #000;background:#f8fafc;text-align:center">No</th>' +
      '<th style="width:72%;padding:4px 8px;border:1.5px solid #000;background:#f8fafc;text-align:center">Item</th>' +
      '<th style="width:11%;padding:4px 2px;border:1.5px solid #000;background:#f8fafc;text-align:center">Yes</th>' +
      '<th style="width:11%;padding:4px 2px;border:1.5px solid #000;background:#f8fafc;text-align:center">Not Required</th>' +
    '</tr>';

    tableRows = precautions.map(p => {
      const v = String(p[2] || '').trim().toLowerCase();
      const isYes = v === 'yes';
      const isNot = v.includes('not') || v === 'no';
      return '<tr>' +
        '<td style="text-align:center;padding:4px 2px;border:1px solid #000;font-weight:bold">' + (p[0] || '') + '</td>' +
        '<td style="padding:4px 8px;border:1px solid #000">' + (p[1] ? escapeHtml_(p[1]) : '&nbsp;') + '</td>' +
        '<td style="text-align:center;padding:4px 2px;border:1px solid #000;font-weight:bold;font-size:12px">' + (isYes ? '✓' : (isTagMode && p[0] ? '{{q' + p[0] + '_yes}}' : '')) + '</td>' +
        '<td style="text-align:center;padding:4px 2px;border:1px solid #000;font-weight:bold;font-size:12px">' + (isNot ? '✓' : (isTagMode && p[0] ? '{{q' + p[0] + '_nr}}' : '')) + '</td>' +
      '</tr>';
    }).join('');
  }

  const workmenBlock = (formCode === 'WP_SHAFT' && (f.workmenNames || isTagMode)) ?
    '<div style="margin:6px 0 10px;font-size:11px"><b>Names of workmen entering shaft:</b> ' + renderLine_(f.workmenNames, 'workmenNames', '300px') + '</div>' : '';

  const permitTitle = (def && def.title) ? def.title : (
    formCode === 'WP_HEIGHT' ? 'WORKING AT HEIGHT PERMIT' :
    formCode === 'WP_HOT' ? 'HOT WORK PERMIT' :
    formCode === 'WP_NIGHT' ? 'NIGHT WORK CHECKLIST PERMIT' :
    formCode === 'WP_SHAFT' ? 'SHAFT WORK & CONFINED SPACE PERMIT' :
    formCode === 'WP_LIFT' ? 'LIFTING WORK PERMIT' : 'GENERAL WORK CLEARANCE PERMIT'
  );

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>' + escapeHtml_(permitTitle) + '</title>' +
    '<style>' +
    '  @page { size: A4 portrait; margin: 8mm 10mm 8mm 10mm; }' +
    '  body { font-family: Arial, Helvetica, sans-serif; color: #000; padding: 14px 18px; line-height: 1.35; font-size: 11px; background: #fff; max-width: 820px; margin: 0 auto; box-sizing: border-box; }' +
    '  table { border-collapse: collapse; }' +
    '  @media print { body { padding: 0 !important; margin: 0 !important; max-width: 100% !important; } .no-print { display: none !important; } }' +
    '</style></head><body>' +

    '<!-- 1. Header Table (1:1 Paper Twin) -->' +
    '<table style="width:100%;border-collapse:collapse;border:2px solid #000;margin-bottom:10px">' +
    '  <tr>' +
    '    <td style="width:34%;border-right:2px solid #000;padding:6px 10px;vertical-align:top">' +
    '      <div style="font-size:11px;font-weight:normal;margin-bottom:2px">Contractor Name:</div>' +
    '      <div style="text-align:center;margin-top:2px">' +
    '        <img src="https://sesipl.com/sites/default/files/sesipl-logo-new-2_5.png" alt="SESIPL Logo" style="height:30px;max-width:160px;object-fit:contain;margin-bottom:2px" onerror="this.onerror=null;this.src=\'https://sesipl.com/sites/default/files/sesipl-logo.png\'">' +
    '        <div style="font-size:10.5px;font-weight:900;color:#003399;letter-spacing:0.3px;line-height:1.2">SHANKAR ELECTRICALS</div>' +
    '        <div style="font-size:8px;font-weight:700;color:#003399;letter-spacing:0.2px;line-height:1.2">SERVICES (I) PRIVATE LIMITED</div>' +
    (contractor ? '<div style="font-size:10px;font-weight:bold;margin-top:3px;color:#0f172a">Agency: ' + escapeHtml_(contractor) + '</div>' : (isTagMode ? '<div style="font-size:10px;font-family:monospace;color:#0284c7;margin-top:2px">{{contractorName}}</div>' : '')) +
    '      </div>' +
    '    </td>' +
    '    <td style="width:34%;border-right:2px solid #000;padding:6px 10px;text-align:center;vertical-align:middle">' +
    '      <div style="font-size:11px;font-weight:bold;margin-bottom:4px">TITLE :</div>' +
    '      <div style="font-size:13px;font-weight:900;letter-spacing:0.5px;color:#000">SAFETY WORK CLEARANCE</div>' +
    '    </td>' +
    '    <td style="width:32%;padding:0;vertical-align:top">' +
    '      <table style="width:100%;height:100%;border-collapse:collapse">' +
    '        <tr>' +
    '          <td style="width:55%;padding:4px 6px;border-bottom:2px solid #000;border-right:1.5px solid #000;font-size:10px;vertical-align:middle">' +
    '            <b>Emergency Contact No.</b>' +
    '          </td>' +
    '          <td style="width:45%;padding:4px 6px;border-bottom:2px solid #000;font-size:9.5px;vertical-align:middle;line-height:1.3">' +
    '            1) ' + renderLine_(em1, 'emergencyContact1', '70px') + '<br>' +
    '            2) ' + renderLine_(em2, 'emergencyContact2', '70px') +
    '          </td>' +
    '        </tr>' +
    '        <tr>' +
    '          <td colspan="2" style="padding:5px 6px;font-size:10.5px;vertical-align:middle">' +
    '            <b>Permit No:-</b> ' + renderLine_(pNo, 'permitNo', '120px') +
    '          </td>' +
    '        </tr>' +
    '      </table>' +
    '    </td>' +
    '  </tr>' +
    '</table>' +

    '<!-- 2. Centered Permit Title -->' +
    '<div style="text-align:center;margin:8px 0 10px">' +
    '  <h2 style="margin:0;font-size:15px;font-weight:900;letter-spacing:0.5px;text-transform:uppercase;color:#000">' + escapeHtml_(permitTitle) + '</h2>' +
    '</div>' +

    '<!-- 3. Work Details & Location with Underline Placeholders -->' +
    '<div style="font-size:10.5px;line-height:1.8;margin-bottom:6px">' +
    '  <div>' +
    '    <b>Area:-</b> ' + renderLine_(area, 'area', '170px') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>location</b> ' + renderLine_(loc, 'location', '170px') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>Date:</b> ' + renderDateLine_(dt, 'date') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>Time:-</b> ' + renderLine_(tm, 'time', '75px') +
    '  </div>' +
    '  <div>' +
    '    <b>Name of Site Engineer (Permit Requesting Authority):</b> ' + renderLine_(siteEng, 'siteEngineer', '230px') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>Sign:</b> ' + renderLine_(siteEngSign, 'siteEngineerSign', '120px') +
    '  </div>' +
    '  <div>' +
    '    <b>Name of Safety Officer:</b> ' + renderLine_(safetyOff, 'safetyOfficer', '250px') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>Sign:</b> ' + renderLine_(safetyOffSign, 'safetyOfficerSign', '140px') +
    '  </div>' +
    '  <div>' +
    '    <b>Name of Contractor Site In charge:</b> ' + renderLine_(contIncharge, 'contractorInCharge', '220px') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>Contact Number:</b> ' + renderLine_(contPhone, 'contactNumber', '130px') +
    '  </div>' +
    '  <div>' +
    '    <b>Description of work:</b> ' + renderLine_(workDesc, 'workDescription', '80%', true) +
    '  </div>' +
    '  <div>' +
    '    <b>Work Execution Date:</b> ' + renderDateLine_(execDate, 'workExecutionDate') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>Valid From:-</b> ' + renderLine_(validFrom, 'validFrom', '130px') +
    '    &nbsp;&nbsp;&nbsp;&nbsp;<b>To:-</b> ' + renderLine_(validTo, 'validTo', '130px') +
    '  </div>' +
    '</div>' +

    '<!-- Declaration Paragraph -->' +
    '<p style="font-size:9.5px;margin:4px 0 3px;color:#111;line-height:1.35">' +
    '  The above signing person will be responsible to ensure that the above described work will be done under all the safety precaution mentioned on the PTW and required by the Project.' +
    '</p>' +
    '<div style="font-size:10px;margin-bottom:5px">The following precautions are to be taken:-</div>' +

    '<!-- 4. Precautions Checklist Table -->' +
    '<table style="width:100%;border-collapse:collapse;border:2px solid #000;font-size:9.5px;margin-bottom:8px">' +
    '  <thead>' +
    tableHeader +
    '  </thead>' +
    '  <tbody>' +
    tableRows +
    '  </tbody>' +
    '</table>' +

    workmenBlock +

    '<!-- 5. Reviewed & Approved By -->' +
    '<div style="margin-top:6px;font-size:10px">' +
    '  <div style="font-weight:bold;margin-bottom:3px">Reviewed &amp; Approved By(Permit Issuing Authority):</div>' +
    '  <div style="margin-bottom:3px">' +
    '    <b>EHS:</b> ' + renderLine_(f.approvalEhsName, 'approvalEhsName', '210px') +
    '    &nbsp;&nbsp;<b>Sign:</b> ' + renderLine_(f.approvalEhsSign || (f.approvalEhsName ? 'Signed' : ''), 'approvalEhsSign', '100px') +
    '    &nbsp;&nbsp;<b>Date:</b> ' + renderDateLine_(f.approvalEhsDate ? displayDate_(f.approvalEhsDate) : '', 'approvalEhsDate') +
    '    &nbsp;&nbsp;<b>Time</b> ' + renderLine_(f.approvalEhsTime, 'approvalEhsTime', '65px') +
    '  </div>' +
    '  <div style="margin-bottom:3px">' +
    '    <b>Site Engineer:</b> ' + renderLine_(f.approvalSiteEngineerName, 'approvalSiteEngineerName', '170px') +
    '    &nbsp;&nbsp;<b>Sign :</b> ' + renderLine_(f.approvalSiteEngineerSign || (f.approvalSiteEngineerName ? 'Signed' : ''), 'approvalSiteEngineerSign', '100px') +
    '    &nbsp;&nbsp;<b>Date:</b> ' + renderDateLine_(f.approvalSiteEngineerDate ? displayDate_(f.approvalSiteEngineerDate) : '', 'approvalSiteEngineerDate') +
    '    &nbsp;&nbsp;<b>Time</b> ' + renderLine_(f.approvalSiteEngineerTime, 'approvalSiteEngineerTime', '65px') +
    '  </div>' +
    '  <p style="font-size:9px;margin:2px 0 6px;color:#222;line-height:1.25">' +
    '    I understand the precaution to be taken as described above and as per Project requirement &amp; here by confirm that Work will be executed under my supervision by following all precaution &amp; Safety Rules.' +
    '  </p>' +
    '</div>' +

    '<!-- 6. Permit Closing / Cancellation -->' +
    '<div style="margin-top:4px;font-size:10px">' +
    '  <div style="font-weight:bold;margin-bottom:2px">Permit Closing / Cancellation:-</div>' +
    '  <p style="font-size:9px;margin:1px 0 4px;color:#222;line-height:1.25">' +
    '    I hereby declare that the work is completed / suspended, all workers under my control have been withdrawn and the site restored to a safe tidy condition.' +
    '  </p>' +
    '  <div style="margin-bottom:3px">' +
    '    <b>Name of Site Engineer (Permit Requesting Authority)</b> ' + renderLine_(f.closingSiteEngName, 'closingSiteEngName', '160px') +
    '    &nbsp;&nbsp;<b>Sign</b> ' + renderLine_(f.closingSiteEngSign || (f.closingSiteEngName ? 'Signed' : ''), 'closingSiteEngSign', '85px') +
    '    &nbsp;&nbsp;<b>Date</b> ' + renderDateLine_(f.closingSiteEngDate ? displayDate_(f.closingSiteEngDate) : '', 'closingSiteEngDate') +
    '    &nbsp;&nbsp;<b>Time</b> ' + renderLine_(f.closingSiteEngTime, 'closingSiteEngTime', '65px') +
    '  </div>' +
    '  <div style="margin-bottom:3px">' +
    '    <b>Name of PMC Site Engineer (Permit Issuing Authority)</b> ' + renderLine_(f.closingPmcSiteEngName, 'closingPmcSiteEngName', '160px') +
    '    &nbsp;&nbsp;<b>Sign</b> ' + renderLine_(f.closingPmcSiteEngSign || (f.closingPmcSiteEngName ? 'Signed' : ''), 'closingPmcSiteEngSign', '85px') +
    '    &nbsp;&nbsp;<b>Date</b> ' + renderDateLine_(f.closingPmcSiteEngDate ? displayDate_(f.closingPmcSiteEngDate) : '', 'closingPmcSiteEngDate') +
    '    &nbsp;&nbsp;<b>Time</b> ' + renderLine_(f.closingPmcSiteEngTime, 'closingPmcSiteEngTime', '65px') +
    '  </div>' +
    '</div>' +

    '<!-- 7. Supporting File Attachments & Placeholders -->' +
    ((f.signedPermitUrl || f.swmsFileUrl || f.preWorkPhotoUrl || f.signedPermitFile || f.swmsFile || f.preWorkPhoto) ?
      '<div style="margin-top:6px;border:1.5px solid #0284c7;padding:5px 8px;border-radius:4px;background:#f0f9ff;font-size:9.5px">' +
      '  <div style="font-weight:bold;color:#0369a1;margin-bottom:2px">📎 Attached Supporting Documents &amp; Reference Files:</div>' +
      '  <table style="width:100%;border-collapse:collapse">' +
      (f.signedPermitUrl || f.signedPermitFile ? '<tr><td style="width:38%;padding:2px 0"><b>Signed PTW Copy / Scan:</b></td><td>' + escapeHtml_(f.signedPermitUrl || f.signedPermitFile) + '</td></tr>' : '') +
      (f.swmsFileUrl || f.swmsFile ? '<tr><td style="width:38%;padding:2px 0"><b>SWMS / JSA Document:</b></td><td>' + escapeHtml_(f.swmsFileUrl || f.swmsFile) + '</td></tr>' : '') +
      (f.preWorkPhotoUrl || f.preWorkPhoto ? '<tr><td style="width:38%;padding:2px 0"><b>Site Verification Photo:</b></td><td>' + escapeHtml_(f.preWorkPhotoUrl || f.preWorkPhoto) + '</td></tr>' : '') +
      '  </table>' +
      '</div>' : '') +

    '<!-- 8. Footer Notation -->' +
    '<table style="width:100%;margin-top:8px;font-size:9px;color:#475569">' +
    '  <tr>' +
    '    <td style="text-align:left"><b>Controlled Copy©</b></td>' +
    '    <td style="text-align:right"><b>Shankar Electricals Services I Pvt Ltd.</b> &nbsp;|&nbsp; <b>' + escapeHtml_(docCode) + '</b></td>' +
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

function buildEhsAuditPdfHtml_(audit, project, user) {
  const data = audit || (typeof getDefaultAuditSeedData_ === 'function' ? getDefaultAuditSeedData_(project) : null);
  const proj = project || { name: (data ? data.projectName : 'Intuit Bellandur'), areaSqft: (data ? data.projectLocation : 'Pritech Park, Bellandur, Bengaluru') };
  const pName = escapeHtml_((data && data.projectName) || proj.name || 'Intuit Bellandur');
  const pLoc = escapeHtml_((data && data.projectLocation) || proj.areaSqft || 'Pritech Park, Bellandur, Bengaluru');
  const aDate = escapeHtml_(displayDate_((data && data.auditDate) || '2026-09-10'));
  const auditor = escapeHtml_((data && data.auditor) || (user && user.name) || 'CBRE Lead Auditor');
  const schema = (typeof AUDIT_CHECKLIST_SCHEMA !== 'undefined') ? AUDIT_CHECKLIST_SCHEMA : [];

  const renderHeader = () => {
    return '<div class="audit-header" style="margin-bottom:6px">' +
      '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;margin-bottom:0">' +
      '  <tr>' +
      '    <td style="width:16%;padding:4px 8px;border-right:1.5px solid #000;text-align:center;vertical-align:middle;background:#fff">' +
      '      <img src="https://sesipl.com/sites/default/files/sesipl-logo-new-2_5.png" style="height:30px;object-fit:contain" alt="SESIPL" onerror="this.onerror=null;this.src=\'https://sesipl.com/sites/default/files/sesipl-logo.png\'">' +
      '      <div style="font-weight:900;font-size:10.5px;color:#0369a1;letter-spacing:0.5px">SESIPL</div>' +
      '    </td>' +
      '    <td style="width:54%;padding:6px 8px;border-right:1.5px solid #000;text-align:center;vertical-align:middle;background:#c6efce">' +
      '      <div style="font-size:15px;font-weight:bold;color:#000;letter-spacing:0.3px">Shankar Electricals EHS Audit Checklist</div>' +
      '    </td>' +
      '    <td style="width:30%;padding:4px 6px;font-size:7.5px;color:#000;text-align:right;vertical-align:middle;line-height:1.2;background:#fff">' +
      '      <div style="font-style:italic">"Sree DeviArcade"</div>' +
      '      <div>668/A, 2nd Floor, 17th C Main, 6th Block, Koramangala, Bengaluru - 560095</div>' +
      '    </td>' +
      '  </tr>' +
      '</table>' +
      '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;border-top:none;margin-bottom:5px;font-size:9.5px;background:#fff">' +
      '  <tr>' +
      '    <td colspan="2" style="padding:2.5px 6px;border-bottom:1px solid #000"><b>Project Name:</b> <span>' + pName + '</span></td>' +
      '  </tr>' +
      '  <tr>' +
      '    <td colspan="2" style="padding:2.5px 6px;border-bottom:1px solid #000"><b>Project Location:</b> <span>' + pLoc + '</span></td>' +
      '  </tr>' +
      '  <tr>' +
      '    <td style="width:50%;padding:2.5px 6px;border-bottom:1px solid #000;border-right:1px solid #000"><b>Audit Date:</b> <span>' + aDate + '</span></td>' +
      '    <td style="width:50%;padding:2.5px 6px;border-bottom:1px solid #000;text-align:right;font-size:8.5px;color:#0369a1">info@shankarelectricals.com www.shankarelectricals.com</td>' +
      '  </tr>' +
      '  <tr>' +
      '    <td colspan="2" style="padding:2.5px 6px"><b>Auditor:</b> <span>' + auditor + '</span></td>' +
      '  </tr>' +
      '</table>' +
      '</div>';
  };

  const renderSectionTable = (sections, itemFilterFn) => {
    let html = '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;font-size:9px;margin-bottom:6px;background:#fff">';
    html += '<tr style="background:#e2e8f0;font-weight:bold">' +
      '<th style="width:3.5%;padding:3px;border:1px solid #000;text-align:center">SN</th>' +
      '<th style="width:68.5%;padding:3px 6px;border:1px solid #000;text-align:left">Particulars</th>' +
      '<th colspan="7" style="width:28%;padding:3px;border:1px solid #000;text-align:center">Score</th>' +
      '</tr>';

    sections.forEach(sec => {
      const secScoreInfo = (data && data.sectionScores && data.sectionScores[sec.id])
        ? data.sectionScores[sec.id]
        : { actual: sec.items.reduce((acc, it) => typeof it.defaultScore === 'number' ? acc + it.defaultScore : acc, 0) };

      const itemsToRender = itemFilterFn ? sec.items.filter(itemFilterFn) : sec.items;
      if (!itemsToRender.length) return;

      // Section Header row
      html += '<tr style="background:#e2e8f0;font-weight:bold">' +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + sec.id + '</td>' +
        '<td style="padding:2.5px 6px;border:1px solid #000">' + escapeHtml_(sec.name) + '</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">0</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">1</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">2</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">3</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">4</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">5</td>' +
        '<td style="width:4%;padding:2.5px;border:1px solid #000;text-align:center">NA</td>' +
        '</tr>';

      itemsToRender.forEach(item => {
        const sc = item.defaultScore;
        const mark = (val) => (sc === val ? '<span style="font-weight:bold;font-size:12px;color:#000">*</span>' : '');
        html += '<tr>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + item.sn + '</td>' +
          '<td style="padding:2.5px 6px;border:1px solid #000">' + escapeHtml_(item.text) + '</td>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + mark(0) + '</td>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + mark(1) + '</td>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + mark(2) + '</td>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + mark(3) + '</td>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + mark(4) + '</td>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + mark(5) + '</td>' +
          '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + mark('NA') + '</td>' +
          '</tr>';
      });

      const isComplete = !itemFilterFn || itemsToRender[itemsToRender.length - 1].sn === sec.items[sec.items.length - 1].sn;
      if (isComplete) {
        html += '<tr>' +
          '<td colspan="2" style="padding:3px 12px;border:1px solid #000;text-align:right;font-weight:bold">Section Total ' + sec.max + '</td>' +
          '<td colspan="7" style="padding:3px;border:1px solid #000;text-align:center;background:#fb923c;color:#000;font-weight:bold;font-size:10.5px">' + secScoreInfo.actual + '</td>' +
          '</tr>';
      }
    });

    html += '</table>';
    return html;
  };

  const renderSummaryScorecardPage = () => {
    let html = renderHeader();
    html += '<div style="background:#f1f5f9;border:1.5px solid #000;border-bottom:none;padding:4px;font-size:9px;font-weight:bold;text-align:center">' +
      '0 - Major NC; 1 - Minor NC; 2 - Partial Compliance; 3 - Full Compliance; NA - Not Applicable' +
      '</div>';

    html += '<table style="width:100%;border-collapse:collapse;border:1.5px solid #000;font-size:9px;background:#fff">';
    html += '<tr style="background:#fbbf24;font-weight:bold;text-align:center">' +
      '<th style="width:4%;padding:3px;border:1px solid #000">SN</th>' +
      '<th style="width:48%;padding:3px 6px;border:1px solid #000;text-align:left">Item</th>' +
      '<th style="width:8%;padding:3px;border:1px solid #000">Max</th>' +
      '<th style="width:8%;padding:3px;border:1px solid #000">Actual</th>' +
      '<th style="width:8%;padding:3px;border:1px solid #000">%</th>' +
      '<th style="width:24%;padding:3px;border:1px solid #000">Performance</th>' +
      '</tr>';

    schema.forEach((sec, idx) => {
      const info = (data && data.sectionScores && data.sectionScores[sec.id])
        ? data.sectionScores[sec.id]
        : { actual: 0, percent: 0, percentText: '0%' };
      const pctDisplay = info.percentText || (sec.max > 0 ? Math.round((info.actual / sec.max) * 100) + '%' : '#DIV/0!');

      let tierCell = '';
      if (idx === 0) {
        tierCell = '<td rowspan="5" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#f8fafc;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#475569">Platinum</div>' +
          '<div style="font-size:9px;color:#64748b;margin-top:2px">85 - 100 %</div>' +
          '</td>';
      } else if (idx === 5) {
        tierCell = '<td rowspan="5" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#fef08a;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#854d0e">Gold</div>' +
          '<div style="font-size:9px;color:#a16207;margin-top:2px">71 - 84 %</div>' +
          '</td>';
      } else if (idx === 10) {
        tierCell = '<td rowspan="4" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#bbf7d0;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#166534">Silver</div>' +
          '<div style="font-size:9px;color:#15803d;margin-top:2px">55 - 70 %</div>' +
          '</td>';
      } else if (idx === 14) {
        tierCell = '<td rowspan="4" style="border:1.5px solid #000;text-align:center;vertical-align:middle;background:#bfdbfe;padding:4px">' +
          '<div style="font-size:10.5px;font-weight:bold;color:#1e40af">Blue</div>' +
          '<div style="font-size:9px;color:#1d4ed8;margin-top:2px">&lt; 54 %</div>' +
          '</td>';
      }

      html += '<tr>' +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center;font-weight:bold">' + sec.id + '</td>' +
        '<td style="padding:2.5px 6px;border:1px solid #000">' + escapeHtml_(sec.name) + '</td>' +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + sec.max + '</td>' +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center;font-weight:bold">' + info.actual + '</td>' +
        '<td style="padding:2.5px;border:1px solid #000;text-align:center">' + pctDisplay + '</td>' +
        tierCell +
        '</tr>';
    });

    const totMax = data ? data.maxScore : 525;
    const totActual = data ? data.totalScore : 363;
    const totPct = data ? data.percent : 69;

    html += '<tr style="font-weight:bold;font-size:10px">' +
      '<td colspan="2" style="padding:4px 6px;border:1.5px solid #000">Total Score</td>' +
      '<td style="padding:4px;border:1.5px solid #000;text-align:center">' + totMax + '</td>' +
      '<td style="padding:4px;border:1.5px solid #000;text-align:center">' + totActual + '</td>' +
      '<td colspan="2" style="padding:4px;border:1.5px solid #000;text-align:center;background:#c6efce;font-size:11.5px;font-weight:900">' + totPct + '%</td>' +
      '</tr>';

    html += '</table>';
    return html;
  };

  const page1Secs = schema.filter(s => ['A','B','C','D','E','F'].indexOf(s.id) >= 0);
  const page2Secs = schema.filter(s => ['G','H','I','J','K','L'].indexOf(s.id) >= 0);
  const secM = schema.find(s => s.id === 'M') || { id: 'M', name: 'Store & Material Management', max: 20, items: [] };
  const page4Secs = schema.filter(s => ['N','O','P','Q','R'].indexOf(s.id) >= 0);

  let fullHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<title>Shankar Electricals EHS Audit Checklist - ' + pName + '</title>' +
    '<style>' +
    '  @page { size: A4 portrait; margin: 8mm 8mm 8mm 8mm; }' +
    '  body { font-family: Calibri, Arial, sans-serif; color: #000; margin: 0; padding: 10px; background: #fff; line-height: 1.25; }' +
    '  .audit-page { page-break-after: always; min-height: 980px; box-sizing: border-box; }' +
    '  .audit-page:last-child { page-break-after: auto; }' +
    '  table { border-collapse: collapse; width: 100%; }' +
    '  th, td { box-sizing: border-box; }' +
    '  @media print {' +
    '    body { padding: 0; background: transparent; }' +
    '    .audit-page { page-break-after: always; min-height: 100vh; }' +
    '    .no-print { display: none !important; }' +
    '  }' +
    '</style>' +
    '</head><body>';

  // Page 1: Sections A through F
  fullHtml += '<div class="audit-page">' + renderHeader() + renderSectionTable(page1Secs) + '</div>';

  // Page 2: Sections G through L + Section M (items 1-3)
  fullHtml += '<div class="audit-page">' + renderHeader() + renderSectionTable(page2Secs.concat([{
    id: secM.id,
    name: secM.name,
    max: secM.max,
    items: secM.items.slice(0, 3)
  }])) + '</div>';

  // Page 3: Section M (item 4) and Section Total 20
  fullHtml += '<div class="audit-page">' + renderHeader() + renderSectionTable([{
    id: secM.id,
    name: secM.name + ' (Continued)',
    max: secM.max,
    items: secM.items.slice(3)
  }]) + '</div>';

  // Page 4: Sections N through R
  fullHtml += '<div class="audit-page">' + renderHeader() + renderSectionTable(page4Secs) + '</div>';

  // Page 5: Executive Grand Summary Scorecard
  fullHtml += '<div class="audit-page">' + renderSummaryScorecardPage() + '</div>';

  fullHtml += '</body></html>';
  return fullHtml;
}

function buildDailyLogPdfHtml_(allLogs, project, activeMonth) {
  const pName = project ? escapeHtml_(project.name) : 'SESIPL Site';
  const mParts = (activeMonth || todayIso_().slice(0, 7)).split('-');
  const yearNum = parseInt(mParts[0], 10) || 2026;
  const monthNum = parseInt(mParts[1], 10) || 9;
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthShort = monthNames[monthNum - 1] || 'Sep';
  const monthYearLabel = monthShort + ' ' + yearNum;
  const daysInMonth = new Date(yearNum, monthNum, 0).getDate();

  const logsByDay = {};
  (allLogs || []).forEach(l => {
    if (!l.date) return;
    const dp = String(l.date).split('-');
    if (dp.length >= 3 && parseInt(dp[0], 10) === yearNum && parseInt(dp[1], 10) === monthNum) {
      logsByDay[parseInt(dp[2], 10)] = l;
    }
  });

  if (Object.keys(logsByDay).length === 0) {
    logsByDay[1] = { staff: 2, workers: 10, totalManpower: 12, workingHours: 8, totalManHours: 96, safeManHours: 96, cumSafeManHours: 96, inductions: 5, indStaff: 3, indWorkers: 5, tbtCount: 1, tbtPersons: 10, trainingTopic: 'Earth pit ex..', trainingPersons: 4, permitHot: 1, permitElectrical: 1, permitCold: '-', permitOthers: '-', firstAid: '-', nearMiss: '-', ltiCount: 0, accidentDetails: 'Nil - Safe Day', remarks: '-' };
    logsByDay[2] = { staff: 2, workers: 9, totalManpower: 11, workingHours: 8, totalManHours: 88, safeManHours: 88, cumSafeManHours: 184, inductions: 1, indStaff: 2, indWorkers: 3, tbtCount: 1, tbtPersons: 9, trainingTopic: '-', trainingPersons: '-', permitHot: '-', permitElectrical: 1, permitCold: 1, permitOthers: 1, firstAid: '-', nearMiss: '-', ltiCount: 0, accidentDetails: 'Nil - Safe Day', remarks: '-' };
    logsByDay[3] = { staff: 3, workers: 7, totalManpower: 10, workingHours: 8, totalManHours: 18, safeManHours: 18, cumSafeManHours: 202, inductions: '-', indStaff: '-', indWorkers: '-', tbtCount: 1, tbtPersons: 10, trainingTopic: 'cable termination', trainingPersons: 7, permitHot: 1, permitElectrical: 1, permitCold: '-', permitOthers: 1, firstAid: '-', nearMiss: '-', ltiCount: 0, accidentDetails: 'Nil - Safe Day', remarks: '-' };
    logsByDay[4] = { staff: 4, workers: 10, totalManpower: 14, workingHours: 8, totalManHours: 112, safeManHours: 112, cumSafeManHours: 314, inductions: 1, indStaff: '-', indWorkers: 2, tbtCount: 1, tbtPersons: 10, trainingTopic: '-', trainingPersons: '-', permitHot: 1, permitElectrical: '-', permitCold: '-', permitOthers: 1, firstAid: '-', nearMiss: '-', ltiCount: 0, accidentDetails: 'Nil - Safe Day', remarks: '-' };
    logsByDay[5] = { staff: 2, workers: 5, totalManpower: 7, workingHours: 8, totalManHours: 56, safeManHours: 56, cumSafeManHours: 370, inductions: '-', indStaff: '-', indWorkers: '-', tbtCount: 1, tbtPersons: 7, trainingTopic: 'Lifting', trainingPersons: 7, permitHot: 1, permitElectrical: 1, permitCold: 1, permitOthers: 1, firstAid: '-', nearMiss: '-', ltiCount: 0, accidentDetails: 'Nil - Safe Day', remarks: '-' };
  }

  let totStaff = 0, totWorkers = 0, totMP = 0, totWorkHrs = 0, totManHrs = 0, totSafeHrs = 0, latestCumSafe = 0;
  let totInd = 0, totIndStaff = 0, totIndWorkers = 0, totTbt = 0, totTbtPersons = 0, totTrainingPersons = 0;
  let totHot = 0, totElect = 0, totCold = 0, totOthers = 0, totFA = 0, totNM = 0, totLTI = 0;

  let rowsHtml = '';
  for (let d = 1; d <= daysInMonth; d++) {
    const entry = logsByDay[d];
    const dateStr = d + '-' + monthShort + '-' + String(yearNum).slice(-2);
    if (entry) {
      const s = Number(entry.staff || 0), w = Number(entry.workers || 0), mp = Number(entry.totalManpower || (s + w));
      const wh = Number(entry.workingHours || 8), tmh = Number(entry.totalManHours || (mp * wh)), smh = Number(entry.safeManHours || tmh);
      const csm = Number(entry.cumSafeManHours || (latestCumSafe + smh));
      latestCumSafe = csm;
      const ind = Number(entry.inductions || 0), indS = Number(entry.indStaff || 0), indW = Number(entry.indWorkers || 0);
      const tbtC = Number(entry.tbtCount || 0), tbtP = Number(entry.tbtPersons || 0), trP = Number(entry.trainingPersons || 0);
      const pH = Number(entry.permitHot || 0), pE = Number(entry.permitElectrical || 0), pC = Number(entry.permitCold || 0), pO = Number(entry.permitOthers || entry.permitGeneral || 0);
      const fa = Number(entry.firstAid || 0), nm = Number(entry.nearMiss || 0), lti = Number(entry.ltiCount || 0);

      totStaff += s; totWorkers += w; totMP += mp; totWorkHrs += wh; totManHrs += tmh; totSafeHrs += smh;
      totInd += ind; totIndStaff += indS; totIndWorkers += indW; totTbt += tbtC; totTbtPersons += tbtP; totTrainingPersons += trP;
      totHot += pH; totElect += pE; totCold += pC; totOthers += pO; totFA += fa; totNM += nm; totLTI += lti;

      rowsHtml += '<tr>' +
        '<td class="c">' + d + '</td><td class="c">' + dateStr + '</td>' +
        '<td class="c">' + (s || '-') + '</td><td class="c">' + (w || '-') + '</td><td class="c bold">' + (mp || '-') + '</td>' +
        '<td class="c">' + (wh || '-') + '</td><td class="c">' + (tmh || '-') + '</td><td class="c">' + (smh || '-') + '</td><td class="c bold">' + (csm || '-') + '</td>' +
        '<td class="c">' + (ind || '-') + '</td><td class="c">' + (indS || '-') + '</td><td class="c">' + (indW || '-') + '</td>' +
        '<td class="c">' + (tbtC || '-') + '</td><td class="c">' + (tbtP || '-') + '</td>' +
        '<td>' + escapeHtml_(entry.trainingTopic || '-') + '</td><td class="c">' + (trP || '-') + '</td>' +
        '<td class="c">' + (pH || '-') + '</td><td class="c">' + (pE || '-') + '</td><td class="c">' + (pC || '-') + '</td><td class="c">' + (pO || '-') + '</td>' +
        '<td class="c">' + (fa || '-') + '</td><td class="c">' + (nm || '-') + '</td><td class="c">' + (lti === 0 ? '0' : (lti || '-')) + '</td>' +
        '<td>' + escapeHtml_(entry.accidentDetails || 'Nil') + '</td><td>' + escapeHtml_(entry.remarks || '-') + '</td>' +
        '</tr>';
    } else {
      rowsHtml += '<tr><td class="c">' + d + '</td><td class="c">' + dateStr + '</td>' +
        '<td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td>' +
        '<td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td>' +
        '<td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td><td class="c">-</td>' +
        '<td class="c">-</td><td class="c">-</td></tr>';
    }
  }

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    '  @page { size: A3 landscape; margin: 8mm; }' +
    '  body { font-family: Calibri, Arial, sans-serif; margin: 0; padding: 0; font-size: 8pt; color: #000; }' +
    '  table { width: 100%; border-collapse: collapse; margin-bottom: 4px; }' +
    '  th, td { border: 1px solid #000; padding: 2px 4px; }' +
    '  .title-banner { background: #bdd7ee; font-size: 13pt; font-weight: bold; text-align: center; padding: 6px; }' +
    '  .hdr1 { background: #d9d9d9; font-weight: bold; font-size: 7.5pt; text-align: center; }' +
    '  .hdr2 { background: #f2f2f2; font-weight: bold; font-size: 7pt; text-align: center; }' +
    '  .tot-row { background: #385723; color: #fff; font-weight: bold; font-size: 8pt; text-align: center; }' +
    '  .foot-label { background: #f2f2f2; font-weight: bold; }' +
    '  .c { text-align: center; }' +
    '  .bold { font-weight: bold; }' +
    '</style>' +
    '</head><body>' +
    '<table>' +
    '  <tr><td colspan="25" class="title-banner">DAILY LOG SHEET</td></tr>' +
    '  <tr>' +
    '    <td colspan="6" style="vertical-align:top"><b>Project: ' + pName + '</b><br>Cumulative Man-Hours Upto: 000<br><b style="color:#0f766e">DAILY PERFORMANCE REPORT</b></td>' +
    '    <td colspan="10" style="text-align:center;vertical-align:middle"><b style="font-size:11pt;color:#002060">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</b><br><small>Since 1998</small></td>' +
    '    <td colspan="9" style="text-align:right;vertical-align:middle"><b>Report for the month of - ' + monthYearLabel + '</b></td>' +
    '  </tr>' +
    '  <tr class="hdr1">' +
    '    <th rowspan="2">SL.NO</th><th rowspan="2">Date</th>' +
    '    <th colspan="3">Man Power</th><th colspan="4">Man hours Statistics</th>' +
    '    <th colspan="3">Safety Induction</th><th colspan="2">Tool Box Talk</th>' +
    '    <th colspan="2">Training Programs</th><th colspan="4">Work Permits</th>' +
    '    <th colspan="4">Accident Statistics</th><th rowspan="2">Remarks</th>' +
    '  </tr>' +
    '  <tr class="hdr2">' +
    '    <th>Staff</th><th>Workers</th><th>Total</th>' +
    '    <th>Work hrs</th><th>Total hrs</th><th>Safe hrs</th><th>Cum Safe</th>' +
    '    <th>Inductions</th><th>Staff</th><th>Workers</th>' +
    '    <th>TBTs</th><th>Attended</th>' +
    '    <th>Topic</th><th>Attended</th>' +
    '    <th>Hot</th><th>Elect</th><th>Cold</th><th>Others</th>' +
    '    <th>FA</th><th>Near Miss</th><th>LTI</th><th>Details</th>' +
    '  </tr>' +
    rowsHtml +
    '  <tr class="tot-row">' +
    '    <td colspan="2">Total</td>' +
    '    <td>' + totStaff + '</td><td>' + totWorkers + '</td><td>' + totMP + '</td>' +
    '    <td>' + totWorkHrs + '</td><td>' + totManHrs + '</td><td>' + totSafeHrs + '</td><td>' + (latestCumSafe || totSafeHrs) + '</td>' +
    '    <td>' + totInd + '</td><td>' + totIndStaff + '</td><td>' + totIndWorkers + '</td>' +
    '    <td>' + totTbt + '</td><td>' + totTbtPersons + '</td>' +
    '    <td>-</td><td>' + totTrainingPersons + '</td>' +
    '    <td>' + totHot + '</td><td>' + totElect + '</td><td>' + totCold + '</td><td>' + totOthers + '</td>' +
    '    <td>' + (totFA || '-') + '</td><td>' + (totNM || '-') + '</td><td>' + (totLTI === 0 ? '0' : (totLTI || '-')) + '</td>' +
    '    <td>-</td><td>-</td>' +
    '  </tr>' +
    '  <tr>' +
    '    <td colspan="8" class="foot-label">Total Man Power Worked for the Month-</td><td colspan="4" class="c bold">' + (totMP || 54) + '</td>' +
    '    <td colspan="3" class="c bold">Date: ' + todayIso_() + '</td><td colspan="3" class="foot-label c">Report Updating by</td>' +
    '    <td colspan="3" class="foot-label c">Report Verified by</td><td colspan="4" class="foot-label c">Report approved by</td>' +
    '  </tr>' +
    '  <tr>' +
    '    <td colspan="8" class="foot-label">Total Safe Man Hours Worked month of</td><td colspan="4" class="c bold">' + (totSafeHrs || 370) + '</td>' +
    '    <td colspan="3" class="c bold">Name</td><td colspan="3" class="c">Site EHS Lead</td>' +
    '    <td colspan="3" class="c">Asst. EHS Manager</td><td colspan="4" class="c">EHS Manager / Director</td>' +
    '  </tr>' +
    '  <tr>' +
    '    <td colspan="8" class="foot-label">Cumulative Safe Man Hours Worked</td><td colspan="4" class="c bold">' + (latestCumSafe || totSafeHrs || 370) + '</td>' +
    '    <td colspan="3" class="c bold">Signature</td><td colspan="3" class="c">xxxx</td>' +
    '    <td colspan="3" class="c">xxxx</td><td colspan="4" class="c">xxx</td>' +
    '  </tr>' +
    '</table></body></html>';
}
