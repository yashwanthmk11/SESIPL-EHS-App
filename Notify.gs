function notifyApprovers_(submitter, project, def, submission) {
  const targets = rowsToObjects_(SHEETS.USERS).filter(u => {
    if (String(u.active).toUpperCase() !== 'TRUE') return false;
    if (u.employeeId === submitter.employeeId) return false;
    if (u.role === ROLES.MANAGER || u.role === ROLES.DIRECTOR) return true;
    if (u.role === ROLES.ASST) {
      return String(u.mappedProjects || '').split(',').map(s => s.trim()).indexOf(project.id) >= 0;
    }
    return false;
  });
  const title = 'Pending approval: ' + def.title;
  const body = submitter.name + ' submitted ' + def.title + ' for ' + project.name + ' (v' + submission.version + ').';
  targets.forEach(t => {
    pushNotify_(t.employeeId, project.id, submitter, title, body, 'APPROVAL');
    mailUser_(t.employeeId, title, body);
  });
}

function broadcastGallery_(user, project, title, category) {
  const pName = project ? project.name : 'SESIPL';
  const pId = project ? project.id : '*';
  const body = user.name + ' (' + user.role + ') uploaded a ' + category.toLowerCase() + ' file: "' + title + '" for ' + pName + '. Please do check the gallery!';
  
  if (user.role === ROLES.MANAGER || user.role === ROLES.DIRECTOR) {
    // Managerial upload: WhatsApp-style broadcast to all active users across company
    const allUsers = rowsToObjects_(SHEETS.USERS).filter(u => String(u.active).toUpperCase() === 'TRUE');
    allUsers.forEach(u => {
      if (u.employeeId !== user.employeeId) {
        pushNotify_(u.employeeId, pId, user, 'New ' + category + ' in Gallery', body, 'GALLERY');
        mailUser_(u.employeeId, 'Gallery Update: ' + title, body);
      }
    });
  } else {
    scopedUsersForProject_(project ? project.id : '').forEach(u => {
      if (u.employeeId !== user.employeeId) {
        pushNotify_(u.employeeId, pId, user, 'Gallery upload', body, 'GALLERY');
        mailUser_(u.employeeId, 'Gallery: ' + title, body);
      }
    });
  }
}

function pushNotify_(toEmployeeId, projectId, fromUser, title, body, type) {
  if (!toEmployeeId) return;
  appendRow_(SHEETS.NOTIFICATIONS, {
    id: uid_('NTF'),
    toEmployeeId: toEmployeeId,
    projectId: projectId || '',
    title: title,
    body: body,
    type: type || 'INFO',
    read: 'FALSE',
    createdAt: nowIso_(),
    createdBy: fromUser ? fromUser.employeeId : 'SYSTEM'
  });
}

function mailUser_(employeeId, subject, body) {
  const u = findOne_(SHEETS.USERS, 'employeeId', employeeId);
  if (!u || !u.email) return;
  try {
    const html =
      '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;background:#ffffff">' +
      '  <div style="background:#147d6f;padding:16px 20px;color:#ffffff">' +
      '    <h3 style="margin:0;font-size:16px;letter-spacing:0.5px">SHANKAR ELECTRICALS SERVICES (I) PVT. LTD.</h3>' +
      '    <div style="font-size:11px;opacity:0.85;margin-top:2px">ENVIRONMENT • HEALTH • SAFETY MANAGEMENT SYSTEM</div>' +
      '  </div>' +
      '  <div style="padding:20px 24px;color:#1e293b">' +
      '    <h4 style="margin:0 0 12px;color:#0f172a;font-size:15px">' + escapeHtml_(subject) + '</h4>' +
      '    <div style="font-size:13px;line-height:1.6;color:#334155;background:#f8fafc;padding:14px;border-left:4px solid #147d6f;border-radius:4px;white-space:pre-wrap">' +
      escapeHtml_(body) +
      '    </div>' +
      '    <div style="margin-top:16px;font-size:12px;color:#64748b">' +
      '      Recipient: <b>' + escapeHtml_(u.name) + '</b> (' + escapeHtml_(u.employeeId) + ')<br>' +
      '      Role: <b>' + escapeHtml_(u.role) + '</b> &bull; Timestamp: ' + nowIso_() +
      '    </div>' +
      '  </div>' +
      '  <div style="background:#f1f5f9;padding:10px 20px;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;text-align:center">' +
      '    SESIPL EHS Cloud Automated Notification &bull; Confidential &bull; Authorized Personnel Only' +
      '  </div>' +
      '</div>';

    MailApp.sendEmail({
      to: u.email,
      subject: '[' + APP_NAME + '] ' + subject,
      body: body,
      htmlBody: html
    });
  } catch (e) {
    writeAudit_('SYSTEM', 'MAIL_FAIL', 'User', employeeId, String(e));
  }
}

function notifyHigherManager_(obs, title, body, fromUser) {
  const managers = rowsToObjects_(SHEETS.USERS).filter(u => u.role === ROLES.MANAGER || u.role === ROLES.DIRECTOR);
  managers.forEach(m => {
    pushNotify_(m.employeeId, obs.projectId, fromUser || { employeeId: 'SYSTEM' }, title, body, 'ESCALATION');
    mailUser_(m.employeeId, title, body + '\nObservation: ' + (obs.observation || '') + '\nProject: ' + obs.projectId);
  });
}

function higherManagerEmails_() {
  return rowsToObjects_(SHEETS.USERS).filter(u => u.role === ROLES.MANAGER || u.role === ROLES.DIRECTOR);
}
