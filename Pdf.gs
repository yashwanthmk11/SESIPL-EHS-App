function generateSubmissionPdf_(project, def, fields, user, version) {
  const html = buildFormPdfHtml_(project, def, fields, user, version);
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
