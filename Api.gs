function sanitizeForClient_(obj) {
  if (obj === null || obj === undefined) return null;
  return JSON.parse(JSON.stringify(obj, (key, value) => {
    if (typeof value === 'number' && (isNaN(value) || !isFinite(value))) {
      return 0;
    }
    if (value === undefined) {
      return null;
    }
    return value;
  }));
}

function apiBootstrap(token) {
  try {
    const user = requireUser_(token);
    const data = buildBootstrap_(user);
    return sanitizeForClient_({ ok: true, user: user, data: data });
  } catch (e) {
    Logger.log("apiBootstrap error: " + e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

function buildBootstrap_(user) {
  let allProjects = rowsToObjects_(SHEETS.PROJECTS).filter(
    (p) => String(p.status).toUpperCase() !== "ARCHIVED",
  );
  if (!allProjects.length) {
    allProjects = ensureDemoProjects_();
  }
  const ids = scopedProjectIds_(user);
  const projects = allProjects.filter((p) => ids.indexOf(p.id) >= 0);
  let selectedProjectId = "";
  if (user.role === ROLES.LEAD && projects.length)
    selectedProjectId = projects[0].id;
  const submissions = rowsToObjects_(SHEETS.SUBMISSIONS).filter(
    (s) => ids.indexOf(s.projectId) >= 0,
  );
  const observations = rowsToObjects_(SHEETS.OBSERVATIONS).filter(
    (s) => ids.indexOf(s.projectId) >= 0,
  );
  const notifications = rowsToObjects_(SHEETS.NOTIFICATIONS)
    .filter(
      (n) =>
        (n.toEmployeeId === user.employeeId ||
          n.toEmployeeId === "*" ||
          n.toEmployeeId === "ROLE:" + user.role) &&
        String(n.dismissed || "").toUpperCase() !== "TRUE",
    )
    .slice(-80)
    .reverse();
  const comments = rowsToObjects_(SHEETS.COMMENTS).filter(
    (c) => ids.indexOf(c.projectId) >= 0,
  );
  const daily = rowsToObjects_(SHEETS.DAILY_LOG).filter(
    (s) => ids.indexOf(s.projectId) >= 0,
  );
  let gallery = rowsToObjects_(SHEETS.GALLERY).filter(
    (s) => ids.indexOf(s.projectId) >= 0,
  );
  if (!gallery.length && projects.length) {
    gallery = ensureDemoGallery_(projects);
  }
  const library = rowsToObjects_(SHEETS.LIBRARY).filter(
    (s) => !s.projectId || ids.indexOf(s.projectId) >= 0,
  );
  const training = rowsToObjects_(SHEETS.TRAINING).filter(
    (s) => !s.projectId || s.projectId === "*" || ids.indexOf(s.projectId) >= 0,
  );
  const ppe = rowsToObjects_(SHEETS.PPE).filter(
    (s) => ids.indexOf(s.projectId) >= 0,
  );
  const scaffold = rowsToObjects_(SHEETS.SCAFFOLD).filter(
    (s) => ids.indexOf(s.projectId) >= 0,
  );
  const movement = rowsToObjects_(SHEETS.MOVEMENT).filter((row) => {
    if (canSeeAllProjects_(user.role)) return true;
    return projects.some(
      (project) =>
        sameProjectValue_(row.fromProject, project) ||
        sameProjectValue_(row.toProject, project),
    );
  });
  const audits = rowsToObjects_(SHEETS.AUDITS).filter(
    (s) => ids.indexOf(s.projectId) >= 0,
  );
  const orgChart = rowsToObjects_(SHEETS.PROJECT_ORG).filter(
    (o) => ids.indexOf(o.projectId) >= 0,
  );
  return {
    projects: projects,
    selectedProjectId: selectedProjectId,
    modules: MODULES,
    catalog: FORM_DEFS,
    submissions: submissions.map(decorateSubmission_),
    observations: observations,
    notifications: notifications,
    comments: comments,
    dailyLog: daily,
    gallery: gallery.map(withFileUrl_),
    library: library.map(withFileUrl_),
    training: training,
    ppe: ppe,
    scaffold: scaffold,
    movement: movement,
    audits: audits,
    orgChart: orgChart,
    users: canMutateAll_(user.role)
      ? rowsToObjects_(SHEETS.USERS).map((u) => ({
          employeeId: u.employeeId,
          name: u.name,
          role: u.role,
          active: u.active,
          mappedProjects: u.mappedProjects,
          email: u.email,
        }))
      : [],
    stats: computeStats_(projects, submissions, observations, daily, user, scaffold, ppe, audits),
  };
}

function decorateSubmission_(s) {
  const def = FORM_DEFS.find((f) => f.formCode === s.formCode) || {};
  return Object.assign({}, s, {
    title: def.title || s.formCode,
    module: def.module || "",
    payload: jsonSafe_(s.payloadJson),
  });
}

function withFileUrl_(row) {
  const out = Object.assign({}, row);
  if (!out.url && out.fileId) {
    out.url = String(out.fileId).startsWith('http') ? out.fileId : ('https://drive.google.com/file/d/' + out.fileId + '/view');
  }
  if (!out.thumbUrl) {
    if (out.fileId && !String(out.fileId).startsWith('http')) {
      out.thumbUrl = 'https://drive.google.com/thumbnail?id=' + out.fileId + '&sz=w800';
    } else if (out.url && String(out.url).startsWith('http')) {
      out.thumbUrl = out.url;
    } else if (out.fileId && String(out.fileId).startsWith('http')) {
      out.thumbUrl = out.fileId;
    }
  }
  return out;
}

function computeStats_(projects, submissions, observations, daily, user, scaffold, ppe, audits) {
  scaffold = scaffold || [];
  ppe = ppe || [];
  audits = audits || [];
  const byProject = {};
  const regions = {};

  projects.forEach((p) => {
    const subs = submissions.filter((s) => s.projectId === p.id);
    const obs = observations.filter((s) => s.projectId === p.id);
    const logs = daily.filter((s) => s.projectId === p.id);
    const scfList = scaffold.filter((s) => s.projectId === p.id);
    const ppeList = ppe.filter((s) => s.projectId === p.id);
    const adtList = audits.filter((s) => s.projectId === p.id);

    const approved = subs.filter((s) => s.status === STATUS.APPROVED).length;
    const pending = subs.filter(
      (s) => s.status === STATUS.SUBMITTED || s.status === STATUS.RESUBMITTED,
    ).length;
    const rejected = subs.filter((s) => s.status === STATUS.REJECTED).length;
    const openObs = obs.filter(
      (s) => s.status === STATUS.OPEN || s.status === STATUS.FALLBACK,
    ).length;
    const closedObs = obs.filter((s) => s.status === STATUS.CLOSED).length;

    // Latest log & sums
    const latestLog = logs.length ? logs[logs.length - 1] : {};
    const totalManpower = Number(latestLog.totalManpower || latestLog.workers || 0);
    const staff = Number(latestLog.staff || 0);
    const workers = Number(latestLog.workers || 0);
    const safeManHours = logs.reduce((a, r) => a + Number(r.safeManHours || r.totalManHours || 0), 0);
    const manHours = logs.reduce((a, r) => a + Number(r.totalManHours || 0), 0);
    const cumSafeManHours = Number(latestLog.cumSafeManHours || safeManHours || 0);

    // Inductions & TBT
    const inductions = logs.reduce((a, r) => a + Number(r.inductions || 0), 0);
    const tbtCount = logs.reduce((a, r) => a + Number(r.tbtCount || 0), 0);
    const tbtPersons = logs.reduce((a, r) => a + Number(r.tbtPersons || 0), 0);

    // Work Permits
    const permitHot = logs.reduce((a, r) => a + Number(r.permitHot || 0), 0);
    const permitElectrical = logs.reduce((a, r) => a + Number(r.permitElectrical || 0), 0);
    const permitCold = logs.reduce((a, r) => a + Number(r.permitCold || 0), 0);
    const permitGeneral = logs.reduce((a, r) => a + Number(r.permitGeneral || 0), 0);
    const permitOthers = logs.reduce((a, r) => a + Number(r.permitOthers || 0), 0);
    const permitTotal = permitHot + permitElectrical + permitCold + permitGeneral + permitOthers;

    // Safety & Accident Statistics
    const firstAid = logs.reduce((a, r) => a + Number(r.firstAid || 0), 0);
    const nearMiss = logs.reduce((a, r) => a + Number(r.nearMiss || 0), 0);
    const ltiCount = logs.reduce((a, r) => a + Number(r.ltiCount || 0), 0);

    // Scaffolds & Ladders
    const latestScf = scfList.length ? scfList[scfList.length - 1] : {};
    const totalScaffold = Number(latestScf.totalScaffold || 0);
    const totalLadder = Number(latestScf.totalLadder || 0);
    const returnedScaffold = Number(latestScf.returnedScaffold || 0);
    const returnedLadder = Number(latestScf.returnedLadder || 0);

    // Audits
    const latestAdt = adtList.length ? adtList[adtList.length - 1] : {};
    const auditScore = Number(latestAdt.totalScore || 0);
    const auditMax = Number(latestAdt.maxScore || 500);
    const auditPercent = Number(latestAdt.percent || (auditMax ? Math.round((auditScore / auditMax) * 100) : 0));

    // Compliance
    const compliance = subs.length
      ? Math.round((approved / subs.length) * 100)
      : (auditPercent || 95);

    const reg = p.region || "Bangalore";
    if (!regions[reg]) {
      regions[reg] = { region: reg, projectCount: 0, manpower: 0, safeHours: 0, openObs: 0, complianceSum: 0, scaffolds: 0, ladders: 0 };
    }
    regions[reg].projectCount++;
    regions[reg].manpower += totalManpower;
    regions[reg].safeHours += cumSafeManHours;
    regions[reg].openObs += openObs;
    regions[reg].complianceSum += compliance;
    regions[reg].scaffolds += totalScaffold;
    regions[reg].ladders += totalLadder;

    byProject[p.id] = {
      projectId: p.id,
      code: p.code,
      name: p.name,
      region: reg,
      client: p.client || "",
      pmc: p.pmc || "",
      inCharge: p.inCharge || "",
      manager: p.manager || "",
      areaSqft: p.areaSqft || "",
      poNo: p.poNo || "",
      projectDuration: p.projectDuration || "",
      status: p.status,
      approved: approved,
      pending: pending,
      rejected: rejected,
      openObs: openObs,
      closedObs: closedObs,
      manHours: manHours,
      safeManHours: safeManHours,
      cumSafeManHours: cumSafeManHours,
      totalManpower: totalManpower,
      staff: staff,
      workers: workers,
      inductions: inductions,
      tbtCount: tbtCount,
      tbtPersons: tbtPersons,
      permitHot: permitHot,
      permitElectrical: permitElectrical,
      permitCold: permitCold,
      permitGeneral: permitGeneral,
      permitOthers: permitOthers,
      permitTotal: permitTotal,
      firstAid: firstAid,
      nearMiss: nearMiss,
      ltiCount: ltiCount,
      totalScaffold: totalScaffold,
      totalLadder: totalLadder,
      returnedScaffold: returnedScaffold,
      returnedLadder: returnedLadder,
      auditScore: auditScore,
      auditMax: auditMax,
      auditPercent: auditPercent,
      compliance: compliance,
      totalSubs: subs.length,
    };
  });

  const values = Object.keys(byProject).map((k) => byProject[k]);
  Object.keys(regions).forEach((r) => {
    regions[r].avgCompliance = regions[r].projectCount ? Math.round(regions[r].complianceSum / regions[r].projectCount) : 0;
  });

  return {
    role: user.role,
    projectCount: projects.length,
    pendingApprovals: values.reduce((a, x) => a + x.pending, 0),
    openObservations: values.reduce((a, x) => a + x.openObs, 0),
    closedObservations: values.reduce((a, x) => a + x.closedObs, 0),
    approved: values.reduce((a, x) => a + x.approved, 0),
    rejected: values.reduce((a, x) => a + x.rejected, 0),
    manHours: values.reduce((a, x) => a + x.manHours, 0),
    totalSafeManHours: values.reduce((a, x) => a + x.safeManHours, 0),
    totalCumSafeManHours: values.reduce((a, x) => a + x.cumSafeManHours, 0),
    totalManpower: values.reduce((a, x) => a + x.totalManpower, 0),
    totalInductions: values.reduce((a, x) => a + x.inductions, 0),
    totalTbtCount: values.reduce((a, x) => a + x.tbtCount, 0),
    totalTbtPersons: values.reduce((a, x) => a + x.tbtPersons, 0),
    totalPermits: values.reduce((a, x) => a + x.permitTotal, 0),
    totalFirstAid: values.reduce((a, x) => a + x.firstAid, 0),
    totalNearMiss: values.reduce((a, x) => a + x.nearMiss, 0),
    totalLti: values.reduce((a, x) => a + x.ltiCount, 0),
    totalScaffolds: values.reduce((a, x) => a + x.totalScaffold, 0),
    totalLadders: values.reduce((a, x) => a + x.totalLadder, 0),
    avgCompliance: values.length ? Math.round(values.reduce((a, x) => a + x.compliance, 0) / values.length) : 0,
    byProject: byProject,
    byRegion: regions,
  };
}

function apiGetFormMeta(token, formCode) {
  requireUser_(token);
  const def = FORM_DEFS.find((f) => f.formCode === formCode);
  if (!def) throw new Error("Unknown form");
  return { ok: true, def: def, fields: getFormFields_(formCode) };
}

function apiSubmitForm(token, payload) {
  const user = requireUser_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const projectId = payload.projectId;
    const formCode = payload.formCode;
    assertProjectAccess_(user, projectId);
    if (
      user.role === ROLES.LEAD &&
      scopedProjectIds_(user).indexOf(projectId) < 0
    ) {
      throw new Error("Lead can only submit for mapped site.");
    }
    const def = FORM_DEFS.find((f) => f.formCode === formCode);
    if (!def) throw new Error("Unknown form");
    if (!canSubmitForm_(user.role, def.entryType)) {
      throw new Error("This upload is managed by the project management team.");
    }
    const requestId = String(payload.requestId || '').trim();
    if (requestId) {
      const duplicate = findOne_(SHEETS.SUBMISSIONS, 'requestId', requestId);
      if (duplicate) {
        return { ok: true, duplicate: true, submission: decorateSubmission_(duplicate), pdfUrl: duplicate.pdfFileId ? apiFileUrl(token, duplicate.pdfFileId).url : '' };
      }
    }
    const project = findOne_(SHEETS.PROJECTS, "id", projectId);
    const existingId = payload.submissionId;
    let version = 1;
    let id = uid_("SUB");
    let prev = null;
    if (existingId) {
      prev = findOne_(SHEETS.SUBMISSIONS, "id", existingId);
      if (!prev) throw new Error("Submission not found");
      const canManage = user.role === ROLES.ASST || canMutateAll_(user.role);
      if (!canManage && prev.status !== STATUS.REJECTED && prev.status !== STATUS.DRAFT) {
        throw new Error(
          "Only rejected / draft items can be updated by resubmit.",
        );
      }
      version = Number(prev.version || 1) + 1;
      id = existingId;
    }
    const data = payload.fields || {};
    const pdf = generateSubmissionPdf_(project, def, data, user, version);
    const row = {
      id: id,
      requestId: requestId,
      projectId: projectId,
      formCode: formCode,
      version: version,
      status: existingId ? STATUS.RESUBMITTED : STATUS.SUBMITTED,
      submittedBy: user.employeeId,
      submittedAt: nowIso_(),
      reviewedBy: "",
      reviewedAt: "",
      pdfFileId: pdf.fileId,
      docFileId: pdf.docId,
      payloadJson: JSON.stringify(data),
    };
    if (existingId) {
      updateRowById_(SHEETS.SUBMISSIONS, id, row);
    } else {
      appendRow_(SHEETS.SUBMISSIONS, row);
    }
    const templateKey = templateKeyForForm_(formCode);
    const templateRecord = {
      id: uid_('TPL'),
      submissionId: id,
      projectId: projectId,
      formCode: formCode,
      templateKey: templateKey,
      version: version,
      status: row.status,
      submittedBy: user.employeeId,
      submittedAt: row.submittedAt,
      reviewedBy: '',
      reviewedAt: '',
      payloadJson: JSON.stringify(data)
    };
    appendRow_(SHEETS.TEMPLATE_RECORDS, templateRecord);
    if (formCode === "OBS_DAILY") {
      createObservationFromForm_(projectId, data, user);
    }
    notifyApprovers_(user, project, def, row);
    writeAudit_(
      user.employeeId,
      existingId ? "RESUBMIT" : "SUBMIT",
      "Submission",
      id,
      formCode,
    );
    return { ok: true, submission: decorateSubmission_(row), pdfUrl: pdf.url };
  } finally {
    lock.releaseLock();
  }
}

function apiReview(token, submissionId, action, comment) {
  const user = requireUser_(token);
  if (!canApprove_(user.role))
    throw new Error("No approval rights for this role.");
  const sub = findOne_(SHEETS.SUBMISSIONS, "id", submissionId);
  if (!sub) throw new Error("Not found");
  assertProjectAccess_(user, sub.projectId);
  const act = String(action).toUpperCase();
  if (act !== "APPROVE" && act !== "REJECT") throw new Error("Invalid action");
  if (act === "REJECT" && !String(comment || "").trim())
    throw new Error("Comment is required when rejecting.");
  const status = act === "APPROVE" ? STATUS.APPROVED : STATUS.REJECTED;
  updateRowById_(SHEETS.SUBMISSIONS, submissionId, {
    status: status,
    reviewedBy: user.employeeId,
    reviewedAt: nowIso_(),
  });
  updateTemplateRecordsForSubmission_(submissionId, status, user.employeeId);
  if (comment) {
    appendRow_(SHEETS.COMMENTS, {
      id: uid_("CMT"),
      entityType: "Submission",
      entityId: submissionId,
      projectId: sub.projectId,
      byEmployeeId: user.employeeId,
      byName: user.name,
      role: user.role,
      text: comment,
      createdAt: nowIso_(),
      broadcast:
        user.role === ROLES.DIRECTOR || user.role === ROLES.MANAGER
          ? "TRUE"
          : "FALSE",
    });
  }
  const submitter = findOne_(SHEETS.USERS, "employeeId", sub.submittedBy);
  pushNotify_(
    submitter ? submitter.employeeId : sub.submittedBy,
    sub.projectId,
    user,
    act === "APPROVE" ? "Submission approved" : "Submission returned",
    sub.formCode + " was " + status + (comment ? ": " + comment : ""),
    "WORKFLOW",
  );
  writeAudit_(user.employeeId, act, "Submission", submissionId, comment || "");
  return { ok: true };
}

function apiAddComment(token, payload) {
  const user = requireUser_(token);
  if (!payload || !payload.projectId) {
    throw new Error("Project ID is required.");
  }
  assertProjectAccess_(user, payload.projectId);
  const text = String(payload.text || "").trim();
  if (!text) {
    throw new Error("Comment cannot be empty.");
  }

  const broadcast =
    (payload.broadcast === true || payload.broadcast === "TRUE") &&
    (user.role === ROLES.MANAGER || user.role === ROLES.DIRECTOR)
      ? "TRUE"
      : "FALSE";

  const commentRow = {
    id: uid_("CMT"),
    entityType: payload.entityType || "Submission",
    entityId: payload.entityId || payload.projectId,
    projectId: payload.projectId,
    byEmployeeId: user.employeeId,
    byName: user.name,
    role: user.role,
    text: text,
    createdAt: nowIso_(),
    broadcast: broadcast,
  };

  appendRow_(SHEETS.COMMENTS, commentRow);

  // Smart context-aware notification routing
  try {
    if (broadcast === "TRUE") {
      scopedUsersForProject_(payload.projectId).forEach((u) => {
        if (u.employeeId !== user.employeeId) {
          pushNotify_(
            u.employeeId,
            payload.projectId,
            user,
            "Notice from " + user.name + " (" + user.role + ")",
            text,
            "COMMENT",
          );
        }
      });
    } else if (commentRow.entityType === "Submission") {
      const sub = findOne_(SHEETS.SUBMISSIONS, "id", commentRow.entityId);
      if (sub) {
        if (user.employeeId !== sub.submittedBy) {
          pushNotify_(
            sub.submittedBy,
            sub.projectId,
            user,
            "New review note on " + (sub.formCode || sub.title || "submission"),
            user.name + " (" + user.role + "): " + text,
            "COMMENT",
          );
        } else {
          scopedUsersForProject_(sub.projectId).forEach((u) => {
            if (canApprove_(u.role) && u.employeeId !== user.employeeId) {
              pushNotify_(
                u.employeeId,
                sub.projectId,
                user,
                "Lead reply on " + (sub.formCode || sub.title || "submission"),
                user.name + ": " + text,
                "COMMENT",
              );
            }
          });
        }
      }
    } else if (commentRow.entityType === "Observation") {
      const obs = findOne_(SHEETS.OBSERVATIONS, "id", commentRow.entityId);
      if (obs && obs.ownerEmployeeId && obs.ownerEmployeeId !== user.employeeId) {
        pushNotify_(
          obs.ownerEmployeeId,
          obs.projectId,
          user,
          "Observation discussion: " + (obs.reportNo || obs.id),
          user.name + ": " + text,
          "COMMENT",
        );
      }
    }
  } catch (err) {
    Logger.log("apiAddComment notify warning: " + err);
  }

  return { ok: true, comment: commentRow };
}

function apiDeleteComment(token, commentId) {
  const user = requireUser_(token);
  const c = findOne_(SHEETS.COMMENTS, "id", commentId);
  if (!c) return { ok: true };
  assertProjectAccess_(user, c.projectId);
  if (
    c.byEmployeeId !== user.employeeId &&
    user.role !== ROLES.DIRECTOR &&
    user.role !== ROLES.MANAGER
  ) {
    throw new Error("You can only delete your own comments.");
  }
  const sh = sheet_(SHEETS.COMMENTS);
  sh.deleteRow(c._row);
  try {
    writeAudit_(user.employeeId, "DELETE_COMMENT", "Comment", commentId, c.text);
  } catch (e) {}
  return { ok: true };
}

function apiDeleteSubmission(token, submissionId) {
  const user = requireUser_(token);
  const sub = findOne_(SHEETS.SUBMISSIONS, "id", submissionId);
  if (!sub) throw new Error("Not found");
  assertDelete_(user, sub.projectId);
  const sh = sheet_(SHEETS.SUBMISSIONS);
  sh.deleteRow(sub._row);
  writeAudit_(
    user.employeeId,
    "DELETE",
    "Submission",
    submissionId,
    sub.formCode,
  );
  return { ok: true };
}

function apiSaveDailyLog(token, row) {
  const user = requireUser_(token);
  assertEdit_(user, row.projectId);
  const staff = Number(row.staff || 0);
  const workers = Number(row.workers || 0);
  const hours = Number(row.workingHours || 8);
  row.id = row.id || uid_("DL");
  row.totalManpower = staff + workers;
  row.totalManHours = (staff + workers) * hours;
  row.safeManHours = Number(row.safeManHours) || row.totalManHours;
  row.cumSafeManHours = Number(row.cumSafeManHours) || row.safeManHours;
  row.enteredBy = user.employeeId;
  row.date = row.date || todayIso_();
  appendRow_(SHEETS.DAILY_LOG, row);
  writeAudit_(user.employeeId, "DAILY_LOG", "DailyLog", row.id, row.date);
  return { ok: true, id: row.id };
}

function apiSavePpe(token, row) {
  const user = requireUser_(token);
  assertEdit_(user, row.projectId);
  row.id = row.id || uid_("PPE");
  row.enteredBy = user.employeeId;
  row.date = row.date || todayIso_();
  const received = Number(row.totalReceived || 0);
  const issued = Number(row.issuedQty || 0);
  if (received > 0 && !row.balanceStock) {
    row.balanceStock = Math.max(0, received - issued);
  }
  appendRow_(SHEETS.PPE, row);
  writeAudit_(user.employeeId, "SAVE_PPE", "PpeIssue", row.id, row.contractor || "");
  return { ok: true };
}

function apiSaveScaffold(token, row) {
  const user = requireUser_(token);
  assertEdit_(user, row.projectId);
  row.id = row.id || uid_("SCF");
  row.enteredBy = user.employeeId;
  row.date = row.date || todayIso_();
  const nums = ["sharavFab", "sesipl", "vinayaka", "hbs", "other"];
  const returnedScf = ["returnedSesipl", "returnedHbs", "returnedVinayaka", "returnedSharav"];
  const totalScf = nums.reduce((a, k) => a + Number(row[k] || 0), 0);
  row.totalScaffold = row.totalScaffold ? Number(row.totalScaffold) : totalScf;
  const ladders = [
    "ladderSharav",
    "ladderSesipl",
    "ladderVinayaka",
    "ladderHbs",
    "ladderRental",
    "airportLadder",
    "workstationLadder",
    "frpMsafe",
    "frpYoungman",
  ];
  const totalLadders = ladders.reduce((a, k) => a + Number(row[k] || 0), 0);
  row.totalLadder = row.totalLadder ? Number(row.totalLadder) : totalLadders;
  appendRow_(SHEETS.SCAFFOLD, row);
  writeAudit_(user.employeeId, "SAVE_SCAFFOLD", "ScaffoldTracker", row.id, row.projectId);
  return { ok: true };
}

function apiSaveProjectOrgMember(token, row) {
  const user = requireUser_(token);
  if (!canMutateAll_(user.role) && user.role !== ROLES.ASST) {
    throw new Error("Only managerial roles can update project organization chart.");
  }
  if (!row.projectId) throw new Error("Missing projectId");
  if (row.id) {
    updateRowById_(SHEETS.PROJECT_ORG, row.id, row);
  } else {
    row.id = uid_("ORG");
    appendRow_(SHEETS.PROJECT_ORG, row);
  }
  writeAudit_(user.employeeId, "SAVE_ORG", "ProjectOrgChart", row.id, row.roleTitle || "");
  return { ok: true, id: row.id };
}

function apiSaveMovement(token, row) {
  const user = requireUser_(token);
  if (!canMutateAll_(user.role) && user.role !== ROLES.ASST)
    throw new Error("Movement tracker is managerial.");
  row.id = uid_("MOV");
  row.enteredBy = user.employeeId;
  row.date = row.date || todayIso_();
  appendRow_(SHEETS.MOVEMENT, row);
  return { ok: true };
}

function apiSaveObservation(token, row) {
  const user = requireUser_(token);
  assertProjectAccess_(user, row.projectId);
  row.id = uid_("OBS");
  row.status = STATUS.OPEN;
  row.submittedBy = user.employeeId;
  row.date = row.date || todayIso_();
  const due = new Date();
  due.setHours(due.getHours() + ESCALATE_HOURS);
  row.dueAt = Utilities.formatDate(
    due,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss",
  );
  row.escalateAt = row.dueAt;
  appendRow_(SHEETS.OBSERVATIONS, row);
  const owner = row.ownerEmployeeId || leadForProject_(row.projectId);
  if (owner) {
    pushNotify_(
      owner,
      row.projectId,
      user,
      "New observation assigned",
      row.observation || "",
      "OBSERVATION",
    );
    mailUser_(
      owner,
      "EHS Observation assigned — action in 24 hours",
      "An observation was logged for your site. Close it or write a fallback review within 24 hours.\n\n" +
        (row.observation || ""),
    );
  }
  return { ok: true, id: row.id };
}

function apiCloseObservation(token, id, fallbackNote) {
  const user = requireUser_(token);
  const obs = findOne_(SHEETS.OBSERVATIONS, "id", id);
  if (!obs) throw new Error("Not found");
  assertProjectAccess_(user, obs.projectId);
  if (fallbackNote) {
    updateRowById_(SHEETS.OBSERVATIONS, id, {
      status: STATUS.FALLBACK,
      fallbackNote: fallbackNote,
      fallbackAt: nowIso_(),
    });
    notifyHigherManager_(
      obs,
      "Fallback review on open observation",
      fallbackNote,
      user,
    );
    return { ok: true, status: STATUS.FALLBACK };
  }
  updateRowById_(SHEETS.OBSERVATIONS, id, { status: STATUS.CLOSED });
  return { ok: true, status: STATUS.CLOSED };
}

function apiSaveTraining(token, row) {
  const user = requireUser_(token);
  if (user.role !== ROLES.MANAGER && user.role !== ROLES.DIRECTOR && user.role !== ROLES.ASST) {
    throw new Error("Only Manager, Director, or Assistant Manager can manage the training calendar.");
  }

  if (row.id) {
    const existing = findOne_(SHEETS.TRAINING, "id", row.id);
    if (!existing) throw new Error("Training session not found: " + row.id);
    updateRowById_(SHEETS.TRAINING, row.id, row);
    try {
      writeAudit_(user.employeeId, "UPDATE_TRAINING", "Training", row.id, row.topic || "");
    } catch (_) {}
    return { ok: true, id: row.id };
  }

  row.id = uid_("TRN");
  row.createdBy = user.employeeId;
  row.status = row.status || "SCHEDULED";
  appendRow_(SHEETS.TRAINING, row);

  try {
    writeAudit_(user.employeeId, "CREATE_TRAINING", "Training", row.id, row.topic || "");
  } catch (_) {}

  rowsToObjects_(SHEETS.USERS).forEach((u) => {
    if (String(u.active).toUpperCase() === "TRUE") {
      pushNotify_(
        u.employeeId,
        row.projectId || "",
        user,
        "Training calendar updated",
        (row.date || "") + " — " + (row.topic || ""),
        "CALENDAR",
      );
    }
  });
  return { ok: true, id: row.id };
}

function apiDeleteTraining(token, trainingId) {
  const user = requireUser_(token);
  if (user.role !== ROLES.MANAGER && user.role !== ROLES.DIRECTOR && user.role !== ROLES.ASST) {
    throw new Error("Only Manager, Director, or Assistant Manager can delete training sessions.");
  }
  const id = String(trainingId || "").trim();
  if (!id) throw new Error("Training ID is required.");
  const item = findOne_(SHEETS.TRAINING, "id", id);
  if (!item) return { ok: true };
  const sh = sheet_(SHEETS.TRAINING);
  sh.deleteRow(item._row);
  try {
    writeAudit_(user.employeeId, "DELETE_TRAINING", "Training", id, item.topic || "");
  } catch (_) {}
  return { ok: true, deletedId: id };
}

function apiUploadMeta(token, payload) {
  const user = requireUser_(token);
  if (!canUploadManagedFile_(user.role)) {
    throw new Error(
      "Only Asst EHS Manager, EHS Manager or Director can upload managed files.",
    );
  }
  assertEdit_(user, payload.projectId);
  const project = findOne_(SHEETS.PROJECTS, "id", payload.projectId);
  const folderName =
    payload.kind === "GALLERY" ? "Gallery" : payload.moduleTitle || "eLibrary";
  const folder = getNamedSubfolder_(project, folderName);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(payload.base64),
    payload.mimeType,
    payload.fileName,
  );
  const file = folder.createFile(blob);
  if (payload.kind === "GALLERY") {
    const cat =
      GALLERY_CATEGORIES.indexOf(payload.category) >= 0
        ? payload.category
        : "EVENT";
    appendRow_(SHEETS.GALLERY, {
      id: uid_("GAL"),
      projectId: payload.projectId,
      category: cat,
      title: payload.title || payload.fileName,
      fileId: file.getId(),
      mimeType: payload.mimeType,
      uploadedBy: user.employeeId,
      uploadedAt: nowIso_(),
    });
    broadcastGallery_(user, project, payload.title || payload.fileName, cat);
  } else {
    appendRow_(SHEETS.LIBRARY, {
      id: uid_("LIB"),
      projectId: payload.projectId,
      module: payload.module || "EHS_DOCS",
      title: payload.title || payload.fileName,
      fileId: file.getId(),
      uploadedBy: user.employeeId,
      uploadedAt: nowIso_(),
      tags: payload.tags || "",
    });
  }
  writeAudit_(
    user.employeeId,
    "UPLOAD",
    payload.kind || "FILE",
    file.getId(),
    payload.fileName,
  );
  return { ok: true, fileId: file.getId(), url: file.getUrl() };
}

function apiDeleteUploadedFile(token, fileId) {
  const user = requireUser_(token);
  const libraryFile = findOne_(SHEETS.LIBRARY, "fileId", fileId);
  const galleryFile = findOne_(SHEETS.GALLERY, "fileId", fileId);
  const row = libraryFile || galleryFile;
  if (!row) throw new Error("Uploaded file not found.");
  assertDelete_(user, row.projectId);

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {}
  const sheetName = libraryFile ? SHEETS.LIBRARY : SHEETS.GALLERY;
  sheet_(sheetName).deleteRow(row._row);
  writeAudit_(user.employeeId, "DELETE", sheetName, fileId, row.title || "");
  return { ok: true };
}

function apiExportProject(token, projectId, format) {
  const user = requireUser_(token);
  assertProjectAccess_(user, projectId);
  const project = findOne_(SHEETS.PROJECTS, "id", projectId);
  const subs = rowsToObjects_(SHEETS.SUBMISSIONS).filter(
    (s) => s.projectId === projectId,
  );
  if (format === "xlsx") {
    const book = SpreadsheetApp.create(project.code + ' EHS Export ' + todayIso_());
    writeExportSheet_(book.getSheets()[0], "Submissions",
      ["id", "formCode", "status", "version", "submittedBy", "submittedAt", "reviewedBy", "payload"],
      subs.map((s) => [s.id, s.formCode, s.status, s.version, s.submittedBy, s.submittedAt, s.reviewedBy, s.payloadJson]));
    const storedRecords = rowsToObjects_(SHEETS.TEMPLATE_RECORDS).filter(r => r.projectId === projectId);
    const knownSubmissionIds = storedRecords.map(r => r.submissionId);
    const records = storedRecords.concat(subs.filter(s => knownSubmissionIds.indexOf(s.id) < 0).map(s => ({
      id: 'LEGACY_' + s.id,
      submissionId: s.id,
      projectId: s.projectId,
      formCode: s.formCode,
      templateKey: templateKeyForForm_(s.formCode),
      version: s.version,
      status: s.status,
      submittedBy: s.submittedBy,
      submittedAt: s.submittedAt,
      payloadJson: s.payloadJson
    })));
    writeTemplateExportSheets_(book, records, project);
    const file = DriveApp.getFileById(book.getId());
    getNamedSubfolder_(project, "Generated PDFs").addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    return { ok: true, url: book.getUrl(), name: book.getName() };
  }
  const html = buildReportHtml_(project, subs, user);
  const pdf = htmlToPdfFile_(
    html,
    project.code + "-EHS-Report-" + todayIso_(),
    getNamedSubfolder_(project, "Generated PDFs"),
  );
  return { ok: true, url: pdf.getUrl(), name: pdf.getName() };
}

function templateKeyForForm_(formCode) {
  if (formCode.indexOf('CL_') === 0 || formCode.indexOf('TAG_') === 0) return 'CHECKLIST_FORMATS';
  if (formCode.indexOf('WP_') === 0) return 'WORK_PERMIT';
  if (formCode.indexOf('WR_') === 0) return 'WEEKLY_MONTHLY_REPORT';
  if (formCode === 'OBS_DAILY') return 'SAFETY_OBSERVATION';
  if (formCode.indexOf('_UPLOAD') >= 0 || formCode === 'TC_UPLOAD') return 'DOCUMENT_UPLOAD';
  return 'EHS_FORM';
}

function writeExportSheet_(bookSheet, name, headers, rows) {
  bookSheet.setName(name.substring(0, 99));
  bookSheet.clear();
  bookSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  bookSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  if (rows.length) bookSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  bookSheet.setFrozenRows(1);
  bookSheet.autoResizeColumns(1, headers.length);
}

function writeTemplateExportSheets_(book, records, project) {
  const groups = {};
  records.forEach(r => {
    if (r.templateKey === 'SAFETY_OBSERVATION') return;
    const key = r.templateKey || 'EHS_FORM';
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });
  const operational = [
    ['DAILY_LOG', SHEETS.DAILY_LOG],
    ['PPE_REGISTER', SHEETS.PPE],
    ['SCAFFOLD_LADDER', SHEETS.SCAFFOLD],
    ['SAFETY_OBSERVATION', SHEETS.OBSERVATIONS],
    ['AUDIT_CHECKLIST', SHEETS.AUDITS]
  ];
  operational.forEach(pair => {
    const key = pair[0];
    groups[key] = rowsToObjects_(pair[1])
      .filter(r => r.projectId === project.id)
      .map(r => ({ id: r.id, formCode: key, version: '', status: r.status || '', submittedBy: r.enteredBy || r.submittedBy || '', submittedAt: r.date || '', payloadJson: JSON.stringify(r) }));
  });
  groups.SCAFFOLD_MOVEMENT = rowsToObjects_(SHEETS.MOVEMENT)
    .filter(r => sameProjectValue_(r.fromProject, project) || sameProjectValue_(r.toProject, project))
    .map(r => ({ id: r.id, formCode: 'SCAFFOLD_MOVEMENT', version: '', status: '', submittedBy: r.enteredBy || '', submittedAt: r.date || '', payloadJson: JSON.stringify(r) }));
  const layouts = {
    DAILY_LOG: ['SL.NO', 'Date', 'Staff', 'Workers', 'Total manpower', 'Working hours', 'Total man hours', 'Safe man hours', 'Cumulative safe man hours', 'Safety inductions', 'TBT count', 'TBT persons', 'Training topic', 'Hot permits', 'General permits', 'First aid', 'Near miss', 'Remarks'],
    PPE_REGISTER: ['S no.', 'Contractor', 'Received by', 'Helmet', 'Jacket', 'Cotton hand gloves', 'Face shield', 'Mask', 'Apron', 'Leather hand gloves', 'Goggle', 'Shoulder pad', 'Ear muff', 'Remarks'],
    SAFETY_OBSERVATION: ['Sl.No', 'Date', 'Observation', 'Unsafe photographs', 'Contractor', 'Location', 'Preventive measures', 'Rectified photos', 'Communicated through & to whom', 'Open / Closed'],
    SCAFFOLD_LADDER: ['Sno', 'Sharav fab', 'SESIPL', 'Vinayaka', 'HBS', 'Other', 'Total', 'Sharav fab ladder', 'SESIPL ladder', 'Vinayaka ladder', 'HBS ladder', 'Airport ladder', 'Workstation ladder', 'FRP Msafe', 'FRP Youngman', 'Total ladder'],
    SCAFFOLD_MOVEMENT: ['S no', 'Date', 'From project', 'From quantity', 'To project', 'To quantity', 'Remarks'],
    AUDIT_CHECKLIST: ['Record ID', 'Audit date', 'Auditor', 'Location', 'Total score', 'Maximum score', 'Percent', 'Grade', 'Status'],
    WEEKLY_MONTHLY_REPORT: ['Record ID', 'Form', 'Version', 'Status', 'Submitted by', 'Submitted at', 'Fields'],
    WORK_PERMIT: ['Record ID', 'Permit form', 'Version', 'Status', 'Submitted by', 'Submitted at', 'Fields'],
    CHECKLIST_FORMATS: ['Record ID', 'Checklist / tag', 'Version', 'Status', 'Submitted by', 'Submitted at', 'Fields'],
    DOCUMENT_UPLOAD: ['Record ID', 'Document form', 'Version', 'Status', 'Submitted by', 'Submitted at', 'Fields'],
    EHS_FORM: ['Record ID', 'Form', 'Version', 'Status', 'Submitted by', 'Submitted at', 'Fields']
  };
  const first = book.getSheets()[0];
  Object.keys(layouts).forEach((key, index) => {
    const sheet = index === 0 && first.getName() === 'Submissions' ? book.insertSheet() : (book.getSheetByName(key) || book.insertSheet());
    const rows = groups[key] || [];
    const headers = layouts[key];
    const values = rows.map((r, i) => templateRow_(key, r, i + 1));
    writeExportSheet_(sheet, key, headers, values);
  });
}

function templateRow_(key, record, sequence) {
  const data = jsonSafe_(record.payloadJson) || {};
  if (key === 'DAILY_LOG') return [sequence, data.date || '', data.staff || '', data.workers || '', data.totalManpower || '', data.workingHours || '', data.totalManHours || '', data.safeManHours || '', data.cumSafeManHours || '', data.inductions || '', data.tbtCount || '', data.tbtPersons || '', data.trainingTopic || '', data.permitHot || '', data.permitGeneral || '', data.firstAid || '', data.nearMiss || '', data.remarks || ''];
  if (key === 'PPE_REGISTER') return [sequence, data.contractor || '', data.receivedBy || '', data.helmet || '', data.jacket || '', data.cottonGloves || '', data.faceShield || '', data.mask || '', data.apron || '', data.leatherGloves || '', data.goggle || '', data.shoulderPad || '', data.earMuff || '', data.remarks || ''];
  if (key === 'SAFETY_OBSERVATION') return [sequence, data.date || '', data.observation || '', data.unsafeFileId || '', data.contractor || '', data.location || '', data.preventive || '', data.rectifiedFileId || '', data.communicatedTo || '', data.status || record.status || ''];
  if (key === 'SCAFFOLD_LADDER') return [sequence, data.sharavFab || '', data.sesipl || '', data.vinayaka || '', data.hbs || '', data.other || '', data.totalScaffold || '', data.ladderSharav || '', data.ladderSesipl || '', data.ladderVinayaka || '', data.ladderHbs || '', data.airportLadder || '', data.workstationLadder || '', data.frpMsafe || '', data.frpYoungman || '', data.totalLadder || ''];
  if (key === 'SCAFFOLD_MOVEMENT') return [sequence, data.date || '', data.fromProject || '', data.fromQty || '', data.toProject || '', data.toQty || '', data.remarks || ''];
  if (key === 'AUDIT_CHECKLIST') return [record.id, data.auditDate || '', data.auditor || '', data.location || '', data.totalScore || '', data.maxScore || '', data.percent || '', data.grade || '', data.status || record.status || ''];
  return [record.id, record.formCode, record.version, record.status, record.submittedBy, record.submittedAt, JSON.stringify(data)];
}

function updateTemplateRecordsForSubmission_(submissionId, status, reviewedBy) {
  rowsToObjects_(SHEETS.TEMPLATE_RECORDS)
    .filter(r => r.submissionId === submissionId)
    .forEach(r => updateRowById_(SHEETS.TEMPLATE_RECORDS, r.id, {
      status: status,
      reviewedBy: reviewedBy,
      reviewedAt: nowIso_()
    }));
}

function sameProjectValue_(value, project) {
  const v = String(value || '').trim().toUpperCase();
  return v === String(project.id || '').toUpperCase() ||
    v === String(project.code || '').trim().toUpperCase() ||
    v === String(project.name || '').trim().toUpperCase();
}

function apiSaveUser(token, row) {
  const user = requireUser_(token);
  if (!canMutateAll_(user.role))
    throw new Error("Only Manager / Director can manage users.");
  const existing = findOne_(SHEETS.USERS, "employeeId", row.employeeId);
  if (existing) {
    const sh = sheet_(SHEETS.USERS);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    Object.keys(row).forEach((k) => {
      const c = headers.indexOf(k);
      if (c >= 0) sh.getRange(existing._row, c + 1).setValue(row[k]);
    });
  } else {
    row.active = row.active || "TRUE";
    appendRow_(SHEETS.USERS, row);
  }
  return { ok: true };
}

function apiCreateProject(token, payload) {
  const user = requireUser_(token);
  if (!canManageProjects_(user.role)) {
    throw new Error(
      "Access Denied: Only Director, EHS Manager, or Assistant EHS Manager can add new projects.",
    );
  }

  const code = String(payload.code || "").trim().toUpperCase();
  const name = String(payload.name || "").trim();
  if (!code || !name) {
    throw new Error("Project Code and Project Name are required.");
  }

  const existing = findOne_(SHEETS.PROJECTS, "code", code);
  if (existing) {
    throw new Error("A project with code '" + code + "' already exists.");
  }

  const cleanCode = code.replace(/[^A-Z0-9_-]/g, "");
  const projId = payload.id || ("PRJ-" + cleanCode);

  const newProj = {
    id: projId,
    code: code,
    name: name,
    client: String(payload.client || "").trim(),
    pmc: String(payload.pmc || "").trim(),
    inCharge: String(payload.inCharge || user.name || "").trim(),
    manager: String(payload.manager || "").trim(),
    scope: String(payload.scope || "Electrical & Safety Infrastructure").trim(),
    startDate: payload.startDate || nowIso_().split(" ")[0],
    endDate: payload.endDate || "",
    areaSqft: String(payload.areaSqft || "").trim(),
    poNo: String(payload.poNo || "").trim(),
    status: String(payload.status || "RUNNING").toUpperCase(),
    region: String(payload.region || "Bangalore").trim(),
    projectDuration: String(payload.projectDuration || "12 Months").trim(),
  };

  appendRow_(SHEETS.PROJECTS, newProj);

  // Map the creator and optional lead so they have access
  try {
    appendRow_(SHEETS.PROJECT_USERS, {
      employeeId: user.employeeId,
      projectId: projId,
      role: user.role,
    });
    if (payload.leadEmployeeId && payload.leadEmployeeId !== user.employeeId) {
      appendRow_(SHEETS.PROJECT_USERS, {
        employeeId: payload.leadEmployeeId,
        projectId: projId,
        role: ROLES.LEAD,
      });
    }
  } catch (me) {
    Logger.log("Project user mapping notice: " + me);
  }

  // Create Drive project folder structure
  try {
    ensureProjectFolder_(newProj);
  } catch (fe) {
    Logger.log("ensureProjectFolder_ notice: " + fe);
  }

  writeAudit_(
    user.employeeId,
    "CREATE_PROJECT",
    "Projects",
    projId,
    newProj.name,
  );

  // Push notification to management team
  try {
    const notifyBody =
      user.name +
      " (" +
      user.role +
      ") added new project: " +
      newProj.code +
      " - " +
      newProj.name +
      " (" +
      newProj.region +
      ").";
    const users = rowsToObjects_(SHEETS.USERS).filter(
      (u) => u.active !== "FALSE",
    );
    users.forEach((u) => {
      if (canManageProjects_(u.role) || u.employeeId === payload.leadEmployeeId) {
        pushNotify_(
          u.employeeId,
          projId,
          user,
          "New Project Added: " + newProj.code,
          notifyBody,
          "INFO",
        );
      }
    });
  } catch (ne) {
    Logger.log("Project notification notice: " + ne);
  }

  return { ok: true, project: newProj };
}

function apiUpdateProject(token, payload) {
  const user = requireUser_(token);
  if (!canManageProjects_(user.role)) {
    throw new Error(
      "Access Denied: Only Director, EHS Manager, or Assistant EHS Manager can edit projects.",
    );
  }

  const projId = String(payload.id || "").trim();
  if (!projId) {
    throw new Error("Project ID is required.");
  }

  const proj = findOne_(SHEETS.PROJECTS, "id", projId);
  if (!proj) {
    throw new Error("Project not found: " + projId);
  }

  const name = String(payload.name || "").trim();
  if (!name) {
    throw new Error("Project Name is required.");
  }

  const patch = {
    name: name,
    client: String(payload.client || "").trim(),
    pmc: String(payload.pmc || "").trim(),
    inCharge: String(payload.inCharge || "").trim(),
    manager: String(payload.manager || "").trim(),
    scope: String(payload.scope || "").trim(),
    startDate: payload.startDate || proj.startDate || "",
    endDate: payload.endDate || "",
    areaSqft: String(payload.areaSqft || "").trim(),
    poNo: String(payload.poNo || "").trim(),
    status: String(payload.status || "RUNNING").toUpperCase(),
    region: String(payload.region || "Bangalore").trim(),
    projectDuration: String(payload.projectDuration || "").trim(),
  };

  // If code is changed, verify uniqueness
  const newCode = String(payload.code || "").trim().toUpperCase();
  if (newCode && newCode !== String(proj.code || "").toUpperCase()) {
    const conflict = findOne_(SHEETS.PROJECTS, "code", newCode);
    if (conflict && String(conflict.id) !== projId) {
      throw new Error("Project Code '" + newCode + "' is already in use by another project.");
    }
    patch.code = newCode;
  }

  updateRowById_(SHEETS.PROJECTS, projId, patch);

  // If a lead was assigned, map them in PROJECT_USERS if not present
  if (payload.leadEmployeeId) {
    try {
      const mappings = rowsToObjects_(SHEETS.PROJECT_USERS);
      const exists = mappings.some(
        (m) => String(m.projectId) === projId && String(m.employeeId) === String(payload.leadEmployeeId),
      );
      if (!exists) {
        appendRow_(SHEETS.PROJECT_USERS, {
          employeeId: payload.leadEmployeeId,
          projectId: projId,
          role: ROLES.LEAD,
        });
      }
    } catch (me) {
      Logger.log("Update project lead mapping notice: " + me);
    }
  }

  writeAudit_(
    user.employeeId,
    "UPDATE_PROJECT",
    "Projects",
    projId,
    "Updated " + (patch.code || proj.code) + " - " + patch.name,
  );

  return { ok: true, project: Object.assign({}, proj, patch) };
}

function apiDeleteProject(token, projectId) {
  const user = requireUser_(token);
  if (!canManageProjects_(user.role)) {
    throw new Error(
      "Access Denied: Only Director, EHS Manager, or Assistant EHS Manager can delete projects.",
    );
  }

  const projId = String(projectId || "").trim();
  if (!projId) {
    throw new Error("Project ID is required.");
  }

  const proj = findOne_(SHEETS.PROJECTS, "id", projId);
  if (!proj) {
    throw new Error("Project not found: " + projId);
  }

  // Delete project from SHEETS.PROJECTS
  const shPrj = sheet_(SHEETS.PROJECTS);
  shPrj.deleteRow(proj._row);

  // Clean up SHEETS.PROJECT_USERS mappings
  try {
    const shUsers = sheet_(SHEETS.PROJECT_USERS);
    const userMaps = rowsToObjects_(SHEETS.PROJECT_USERS);
    for (let i = userMaps.length - 1; i >= 0; i--) {
      if (String(userMaps[i].projectId) === projId) {
        shUsers.deleteRow(userMaps[i]._row);
      }
    }
  } catch (err) {
    Logger.log("apiDeleteProject user mappings cleanup notice: " + err);
  }

  writeAudit_(
    user.employeeId,
    "DELETE_PROJECT",
    "Projects",
    projId,
    "Deleted " + (proj.code || "") + " - " + (proj.name || ""),
  );

  return { ok: true, deletedId: projId };
}

function apiMarkRead(token, notificationId) {
  const user = requireUser_(token);
  const n = findOne_(SHEETS.NOTIFICATIONS, "id", notificationId);
  if (!n) return { ok: true };
  updateRowById_(SHEETS.NOTIFICATIONS, notificationId, { read: "TRUE" });
  return { ok: true };
}

function apiDismissNotification(token, notificationId) {
  const user = requireUser_(token);
  const n = findOne_(SHEETS.NOTIFICATIONS, "id", notificationId);
  if (!n) return { ok: true };
  updateRowById_(SHEETS.NOTIFICATIONS, notificationId, { read: "TRUE", dismissed: "TRUE" });
  return { ok: true };
}

function apiClearAllNotifications(token) {
  const user = requireUser_(token);
  const sh = sheet_(SHEETS.NOTIFICATIONS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true };
  const headers = values[0];
  const toCol = headers.findIndex(h => normalizeKey_(h) === 'toemployeeid');
  const readCol = headers.findIndex(h => normalizeKey_(h) === 'read');
  let disCol = headers.findIndex(h => normalizeKey_(h) === 'dismissed');
  if (disCol < 0) {
    const nextCol = headers.length + 1;
    sh.getRange(1, nextCol).setValue('DISMISSED');
    disCol = headers.length;
  }
  for (let i = 1; i < values.length; i++) {
    const to = String(values[i][toCol] || '').trim();
    if (to === user.employeeId || to === '*' || to === 'ROLE:' + user.role) {
      if (readCol >= 0) sh.getRange(i + 1, readCol + 1).setValue('TRUE');
      if (disCol >= 0) sh.getRange(i + 1, disCol + 1).setValue('TRUE');
    }
  }
  return { ok: true };
}

function apiFileUrl(token, fileId) {
  requireUser_(token);
  return { ok: true, url: DriveApp.getFileById(fileId).getUrl() };
}

function scopedUsersForProject_(projectId) {
  const map = rowsToObjects_(SHEETS.PROJECT_USERS)
    .filter((r) => r.projectId === projectId)
    .map((r) => r.employeeId);
  return rowsToObjects_(SHEETS.USERS).filter((u) => {
    if (u.role === ROLES.MANAGER || u.role === ROLES.DIRECTOR) return true;
    if (map.indexOf(u.employeeId) >= 0) return true;
    return (
      String(u.mappedProjects || "")
        .split(",")
        .map((s) => s.trim())
        .indexOf(projectId) >= 0
    );
  });
}

function leadForProject_(projectId) {
  const row = rowsToObjects_(SHEETS.PROJECT_USERS).find(
    (r) => r.projectId === projectId && r.role === ROLES.LEAD,
  );
  if (row) return row.employeeId;
  const u = rowsToObjects_(SHEETS.USERS).find(
    (x) =>
      x.role === ROLES.LEAD && String(x.mappedProjects).indexOf(projectId) >= 0,
  );
  return u ? u.employeeId : "";
}

function createObservationFromForm_(projectId, data, user) {
  apiSaveObservationFromInternal_(projectId, data, user);
  const owner = data.ownerEmployeeId || leadForProject_(projectId);
  if (owner) {
    pushNotify_(
      owner,
      projectId,
      user,
      "New observation assigned",
      data.observation || "",
      "OBSERVATION",
    );
    mailUser_(
      owner,
      "EHS Observation assigned — action in 24 hours",
      "An observation was logged for your site. Close it or write a fallback review within 24 hours.\n\n" +
        (data.observation || ""),
    );
  }
}

function apiSaveObservationFromInternal_(projectId, data, user) {
  const due = new Date();
  due.setHours(due.getHours() + ESCALATE_HOURS);
  appendRow_(SHEETS.OBSERVATIONS, {
    id: uid_("OBS"),
    projectId: projectId,
    reportNo: data.reportNo || "",
    date: data.date || todayIso_(),
    vendor: data.vendor || "",
    auditedBy: data.auditedBy || user.name,
    location: data.location || "",
    contractor: data.contractor || "",
    observation: data.observation || "",
    preventive: data.preventive || "",
    status: STATUS.OPEN,
    ownerEmployeeId: data.ownerEmployeeId || leadForProject_(projectId),
    dueAt: Utilities.formatDate(
      due,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss",
    ),
    escalateAt: Utilities.formatDate(
      due,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss",
    ),
    fallbackNote: "",
    fallbackAt: "",
    unsafeFileId: "",
    rectifiedFileId: "",
    submittedBy: user.employeeId,
  });
}

function ensureDemoGallery_(projects) {
  const p1 = projects[0] ? projects[0].id : "PRJ-2026-001";
  const p2 = projects[1] ? projects[1].id : p1;
  const p3 = projects[2] ? projects[2].id : p1;
  const seeds = [
    { id: uid_('GAL'), projectId: p1, category: 'TRAINING', title: 'Work at Height & Safety Harness Induction', fileId: 'https://images.unsplash.com/photo-1541888946425-d0fbb186156a?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP001', uploadedAt: '2026-09-17 10:30' },
    { id: uid_('GAL'), projectId: p1, category: 'COMPANY', title: 'Scaffold Cuplock & Sole Board Audit Inspection', fileId: 'https://images.unsplash.com/photo-1504307651254-35680f356dfd?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP002', uploadedAt: '2026-09-17 09:15' },
    { id: uid_('GAL'), projectId: p2, category: 'MEETING', title: 'Daily Tool Box Talk (TBT) - Electrical Safety & LOTO', fileId: 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP003', uploadedAt: '2026-09-16 08:30' },
    { id: uid_('GAL'), projectId: p1, category: 'EVENT', title: 'Celebration of 370,000 Safe Man-Hours Zero LTI Milestone', fileId: 'https://images.unsplash.com/photo-1531482615713-2afd69097998?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP001', uploadedAt: '2026-09-16 16:00' },
    { id: uid_('GAL'), projectId: p2, category: 'TRAINING', title: 'Live Fire Extinguisher Drill & Emergency Evacuation', fileId: 'https://images.unsplash.com/photo-1574680096145-d05b474e2155?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP002', uploadedAt: '2026-09-16 14:15' },
    { id: uid_('GAL'), projectId: p3, category: 'COMPANY', title: 'Substation Control Panel Earthing & PPE Verification', fileId: 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP005', uploadedAt: '2026-09-15 15:20' },
    { id: uid_('GAL'), projectId: p3, category: 'MEETING', title: 'Joint PMC & Contractor Weekly Safety Committee Walk', fileId: 'https://images.unsplash.com/photo-1578575437130-527eed3abbec?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP005', uploadedAt: '2026-09-15 09:15' },
    { id: uid_('GAL'), projectId: p2, category: 'EVENT', title: 'National Safety Month Best EHS Lead Recognition Award', fileId: 'https://images.unsplash.com/photo-1513694203232-719a280e022f?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP001', uploadedAt: '2026-09-15 17:30' },
    { id: uid_('GAL'), projectId: p1, category: 'COMPANY', title: 'Deep Excavation Shoring & Edge Protection Verification', fileId: 'https://images.unsplash.com/photo-1590069261209-f8e9b8642343?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP002', uploadedAt: '2026-09-14 11:00' },
    { id: uid_('GAL'), projectId: p2, category: 'TRAINING', title: 'Confined Space Entry & Multi-Gas Detector Calibration', fileId: 'https://images.unsplash.com/photo-1517048676732-d65bc937f952?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP003', uploadedAt: '2026-09-14 14:00' },
    { id: uid_('GAL'), projectId: p3, category: 'MEETING', title: 'Pre-Shift Hazard Identification & Lifting Plan Briefing', fileId: 'https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP005', uploadedAt: '2026-09-14 08:00' },
    { id: uid_('GAL'), projectId: p1, category: 'EVENT', title: 'SESIPL Annual HSE Leadership Summit & Trophy Handover', fileId: 'https://images.unsplash.com/photo-1492684223066-81342ee5ff30?w=800&auto=format&fit=crop&q=80', mimeType: 'image/jpeg', uploadedBy: 'EMP001', uploadedAt: '2026-09-13 18:30' }
  ];
  try {
    const sh = sheet_(SHEETS.GALLERY);
    batchWriteObjects_(sh, HEADERS.Gallery, seeds);
  } catch (e) {
    Logger.log("ensureDemoGallery_ warning: " + e);
  }
  return seeds;
}

function apiGetEhsAuditData(token, projectId) {
  const user = requireUser_(token);
  const project = (projectId ? findOne_(SHEETS.PROJECTS, "id", projectId) : null) || (rowsToObjects_(SHEETS.PROJECTS)[0]) || null;
  const pid = project ? project.id : (projectId || 'PRJ001');

  let auditRow = null;
  try {
    auditRow = rowsToObjects_(SHEETS.AUDITS).find(a => a.projectId === pid);
  } catch (err) {
    Logger.log("Notice: sheet Audits lookup: " + err);
  }

  let auditData = null;

  if (auditRow) {
    let scores = [];
    try {
      scores = rowsToObjects_(SHEETS.AUDIT_SCORES).filter(s => s.auditId === auditRow.id);
    } catch (err) {
      Logger.log("Notice: sheet AuditScores lookup: " + err);
    }
    const sectionScores = {};
    scores.forEach(s => {
      const maxVal = (AUDIT_SECTIONS.find(x => x.id === s.section) || {}).max || 0;
      const actualVal = Number(s.score || 0);
      const pct = maxVal > 0 ? Math.round((actualVal / maxVal) * 100) : null;
      sectionScores[s.section] = {
        id: s.section,
        name: (AUDIT_SECTIONS.find(x => x.id === s.section) || {}).name || s.section,
        max: maxVal,
        actual: actualVal,
        percent: pct,
        percentText: pct !== null ? pct + '%' : '#DIV/0!'
      };
    });
    auditData = {
      id: auditRow.id,
      projectId: pid,
      projectName: project ? project.name : 'SESIPL Site',
      projectLocation: project ? (project.name + ', ' + (project.areaSqft || 'Bengaluru')) : 'Bengaluru',
      auditDate: auditRow.auditDate || '2026-09-10',
      auditor: auditRow.auditor || 'CBRE Lead Auditor',
      auditorEmail: 'info@shankarelectricals.com',
      auditorWebsite: 'www.shankarelectricals.com',
      totalScore: Number(auditRow.totalScore || 363),
      maxScore: Number(auditRow.maxScore || 525),
      percent: Number(auditRow.percent || 69),
      grade: auditRow.grade || 'Silver',
      status: auditRow.status || 'APPROVED',
      sectionScores: sectionScores
    };
  } else {
    auditData = getDefaultAuditSeedData_(project);
  }

  return {
    ok: true,
    audit: {
      id: auditData.id,
      projectId: auditData.projectId,
      projectName: auditData.projectName,
      projectLocation: auditData.projectLocation,
      auditDate: auditData.auditDate,
      auditor: auditData.auditor,
      auditorEmail: auditData.auditorEmail,
      auditorWebsite: auditData.auditorWebsite,
      totalScore: auditData.totalScore,
      maxScore: auditData.maxScore,
      percent: auditData.percent,
      grade: auditData.grade,
      status: auditData.status,
      sectionScores: auditData.sectionScores
    },
    project: project ? { id: project.id, code: project.code, name: project.name, areaSqft: project.areaSqft } : null,
    performanceBands: AUDIT_PERFORMANCE_BANDS
  };
}

function apiExportEhsAudit(token, projectId) {
  const user = requireUser_(token);
  const project = (projectId ? findOne_(SHEETS.PROJECTS, "id", projectId) : null) || (rowsToObjects_(SHEETS.PROJECTS)[0]) || null;
  const auditRes = apiGetEhsAuditData(token, projectId);
  const auditData = auditRes.audit;
  auditData.schema = AUDIT_CHECKLIST_SCHEMA;

  const html = buildEhsAuditPdfHtml_(auditData, project, user);
  const name = (project ? project.code : 'SESIPL') + '-EHS-Audit-Checklist-' + todayIso_();
  let pdfRes = { url: '', fileId: '' };

  try {
    const folder = getNamedSubfolder_(project, 'Audits');
    pdfRes = htmlToPdfFile_(html, name, folder);
  } catch (err) {
    Logger.log('apiExportEhsAudit PDF creation notice: ' + err);
  }

  return {
    ok: true,
    url: pdfRes.url || '',
    name: name + '.pdf',
    html: html
  };
}

function apiExportEhsAuditExcel(token, projectId) {
  const user = requireUser_(token);
  const project = (projectId ? findOne_(SHEETS.PROJECTS, "id", projectId) : null) || (rowsToObjects_(SHEETS.PROJECTS)[0]) || null;
  const auditRes = apiGetEhsAuditData(token, projectId);
  const auditData = auditRes.audit;
  auditData.schema = AUDIT_CHECKLIST_SCHEMA;
  const name = (project ? project.code : 'SESIPL') + '-EHS-Audit-Checklist-' + todayIso_();

  const book = SpreadsheetApp.create(name);
  const bookId = book.getId();

  try {
    const sheet = book.getSheets()[0];
    sheet.setName("EHS Audit Checklist");
    buildAuditSingleSheetExcel_(sheet, auditData, project);

    const file = DriveApp.getFileById(bookId);
    try {
      const folder = getNamedSubfolder_(project, 'Audits');
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (folderErr) {
      Logger.log("Notice moving audit excel: " + folderErr);
    }

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("Notice setting audit excel sharing: " + shareErr);
    }
  } catch (err) {
    Logger.log("Error building audit excel: " + err);
  }

  const downloadUrl = 'https://docs.google.com/spreadsheets/d/' + bookId + '/export?format=xlsx';

  return {
    ok: true,
    id: bookId,
    url: book.getUrl(),
    downloadUrl: downloadUrl,
    name: name + '.xlsx',
    format: 'xlsx'
  };
}

function buildAuditSingleSheetExcel_(sheet, data, project) {
  sheet.clear();

  const pName = data ? data.projectName : (project ? project.name : 'SESIPL Site');
  const pLoc = data ? data.projectLocation : (project ? (project.name + ', ' + (project.areaSqft || '')) : 'Bengaluru');
  const aDate = data ? data.auditDate : todayIso_();
  const auditor = data ? data.auditor : 'CBRE Lead Auditor';
  const email = data ? data.auditorEmail : 'info@shankarelectricals.com';
  const website = data ? data.auditorWebsite : 'www.shankarelectricals.com';
  const schema = (data && data.schema) || AUDIT_CHECKLIST_SCHEMA;
  const secScores = (data && data.sectionScores) || {};

  const rows = [];

  // Row 1: Company Logo / Name Header
  rows.push(['SESIPL', 'Shankar Electricals Services (I) Pvt. Ltd.', '', '', '', '', '', '', '"Sree Devi Arcade"', '']);
  // Row 2: Sub-Banner / Address
  rows.push(['', 'ENVIRONMENT, HEALTH & SAFETY MANAGEMENT SYSTEM — AUDIT CHECKLIST', '', '', '', '', '', '', '668/A, 2nd Floor, 17th C Main, 6th Block, Koramangala, Bengaluru - 560095', '']);
  // Row 3: Spacing
  rows.push(['', '', '', '', '', '', '', '', '', '']);
  // Row 4: Project Info Row 1
  rows.push(['Project Name:', pName, '', '', '', '', 'Audit Date:', '', aDate, '']);
  // Row 5: Project Info Row 2
  rows.push(['Location:', pLoc, '', '', '', '', 'Auditor:', '', auditor, '']);
  // Row 6: Contact Info
  rows.push(['Contact:', email + ' | ' + website, '', '', '', '', '', '', '', '']);
  // Row 7: Spacing before checklist
  rows.push(['', '', '', '', '', '', '', '', '', '']);

  // Rows 8 to 163: 18 Sections A to R (120 checkpoints + 18 headers + 18 subtotals = 156 rows)
  const sectionHeaderRows = [];
  const sectionTotalRows = [];
  const dotChar = '·';

  schema.forEach(sec => {
    // Section Header row (Col A: ID, Col B: Name, Cols C-I: 0 1 2 3 4 5 NA)
    const secHeaderRowIndex = rows.length + 1;
    sectionHeaderRows.push(secHeaderRowIndex);
    rows.push([sec.id, sec.name, '0', '1', '2', '3', '4', '5', 'NA', '']);

    let secActual = 0;
    sec.items.forEach((item, idx) => {
      const scoreVal = (typeof item.defaultScore !== 'undefined') ? item.defaultScore : 3;
      if (typeof scoreVal === 'number') {
        secActual += scoreVal;
      }

      const itemRow = [
        item.sn || (idx + 1),
        item.text,
        scoreVal === 0 ? dotChar : '',
        scoreVal === 1 ? dotChar : '',
        scoreVal === 2 ? dotChar : '',
        scoreVal === 3 ? dotChar : '',
        scoreVal === 4 ? dotChar : '',
        scoreVal === 5 ? dotChar : '',
        (scoreVal === 'NA' || scoreVal === 'N/A') ? dotChar : '',
        ''
      ];
      rows.push(itemRow);
    });

    // Section Total row
    const secTotalRowIndex = rows.length + 1;
    sectionTotalRows.push({ row: secTotalRowIndex, actual: secActual, max: sec.max });
    rows.push(['', 'Section Total ' + sec.max, '', '', '', secActual, '', '', '', '']);
  });

  // Row 164: Legend line
  const legendRowIndex = rows.length + 1;
  rows.push(['0 - Major NC; 1 - Minor NC; 2 - Partial Compliance; 3 - Full Compliance; NA - Not Applicable', '', '', '', '', '', '', '', '', '']);

  // Row 165: Summary Scorecard header
  const scHeaderRowIndex = rows.length + 1;
  rows.push(['SN', 'Item', '', '', '', '', 'Max', 'Actual', '%', 'Performance']);

  // Rows 166 to 183: 18 Section Scorecard Rows (A through R)
  const scSectionStartRow = rows.length + 1;
  schema.forEach(sec => {
    const info = secScores[sec.id] || {};
    let actual = typeof info.actual === 'number' ? info.actual : 0;
    if (typeof info.actual === 'undefined') {
      actual = sec.items.reduce((acc, it) => typeof it.defaultScore === 'number' ? acc + it.defaultScore : acc, 0);
    }
    let pctText = '0%';
    if (sec.max > 0) {
      pctText = Math.round((actual / sec.max) * 100) + '%';
    } else {
      pctText = '####';
    }

    rows.push([sec.id, sec.name, '', '', '', '', sec.max, actual, pctText, '']);
  });

  // Row 184: Total Score row
  const totMax = (data && data.maxScore) || 525;
  const totActual = (data && data.totalScore) || 363;
  const totPct = (data && data.percent) || 69;
  const totalRowIndex = rows.length + 1;
  rows.push(['Total Score', '', '', '', '', '', totMax, totActual, totPct + '%', '']);

  // Write all rows in a single batch
  sheet.getRange(1, 1, rows.length, 10).setValues(rows);

  // Column widths
  sheet.setColumnWidth(1, 45);   // A: SN / Sec
  sheet.setColumnWidth(2, 480);  // B: Description
  sheet.setColumnWidth(3, 30);   // C: 0
  sheet.setColumnWidth(4, 30);   // D: 1
  sheet.setColumnWidth(5, 30);   // E: 2
  sheet.setColumnWidth(6, 30);   // F: 3
  sheet.setColumnWidth(7, 30);   // G: 4
  sheet.setColumnWidth(8, 30);   // H: 5
  sheet.setColumnWidth(9, 38);   // I: NA
  sheet.setColumnWidth(10, 140); // J: Performance Tier

  // Format Header Rows 1 to 7
  sheet.getRange("A1").setFontWeight('bold').setFontSize(13).setFontColor('#0369a1').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange("B1:H1").merge().setFontWeight('bold').setFontSize(14).setBackground('#c6efce').setFontColor('#000000').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange("I1:J1").merge().setFontSize(8.5).setFontStyle('italic').setHorizontalAlignment('right').setVerticalAlignment('middle');

  sheet.getRange("B2:H2").merge().setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange("I2:J2").merge().setFontSize(8).setFontColor('#475569').setHorizontalAlignment('right').setVerticalAlignment('middle');

  sheet.getRange("A4:A6").setFontWeight('bold').setFontColor('#334155');
  sheet.getRange("B4:F4").merge();
  sheet.getRange("B5:F5").merge();
  sheet.getRange("B6:F6").merge();
  sheet.getRange("G4:H4").merge().setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange("G5:H5").merge().setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange("I4:J4").merge();
  sheet.getRange("I5:J5").merge();

  // Section Headers formatting (Rows 8, 16, 28, etc.)
  sectionHeaderRows.forEach(r => {
    sheet.getRange(r, 1, 1, 9).setFontWeight('bold').setBackground('#d9d9d9').setFontColor('#000000');
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 2).setHorizontalAlignment('left');
    sheet.getRange(r, 3, 1, 7).setHorizontalAlignment('center');
  });

  // Section Totals formatting (orange badge for score)
  sectionTotalRows.forEach(st => {
    sheet.getRange(st.row, 2).setFontWeight('bold').setHorizontalAlignment('right');
    sheet.getRange(st.row, 6, 1, 4).merge().setFontWeight('bold').setBackground('#ed7d31').setFontColor('#000000').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // Checklist items range formatting (alignment & wrap text)
  sheet.getRange(8, 1, 156, 1).setHorizontalAlignment('center');
  sheet.getRange(8, 2, 156, 1).setWrap(true);
  sheet.getRange(8, 3, 156, 7).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(8, 1, 156, 9).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Row 164: Legend line formatting
  sheet.getRange(legendRowIndex, 1, 1, 10).merge()
    .setFontWeight('bold')
    .setFontColor('#ff0000')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#ffffff');
  sheet.setRowHeight(legendRowIndex, 24);

  // Row 165: Scorecard header formatting
  sheet.getRange(scHeaderRowIndex, 2, 1, 5).merge();
  sheet.getRange(scHeaderRowIndex, 1, 1, 10)
    .setBackground('#ffc000')
    .setFontWeight('bold')
    .setFontColor('#000000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(scHeaderRowIndex, 24);

  // Rows 166 to 183: Scorecard Section rows
  for (let r = scSectionStartRow; r < scSectionStartRow + 18; r++) {
    sheet.getRange(r, 2, 1, 5).merge();
    sheet.getRange(r, 1).setHorizontalAlignment('center').setFontWeight('bold');
    sheet.getRange(r, 2).setHorizontalAlignment('left').setFontWeight('bold');
    sheet.getRange(r, 7, 1, 3).setHorizontalAlignment('center');
    sheet.getRange(r, 8).setFontWeight('bold');
  }

  // Merged Performance blocks in Col J:
  // J166:J170 (A to E) -> Platinum
  sheet.getRange(166, 10, 5, 1).merge()
    .setValue("Platinum\n85 -100 %")
    .setFontWeight('bold')
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // J171:J175 (F to J) -> Gold
  sheet.getRange(171, 10, 5, 1).merge()
    .setValue("Gold\n71 - 84 %")
    .setFontWeight('bold')
    .setBackground('#ffc000')
    .setFontColor('#000000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // J176:J179 (K to N) -> Silver
  sheet.getRange(176, 10, 4, 1).merge()
    .setValue("Silver\n55 - 70 %")
    .setFontWeight('bold')
    .setBackground('#c6efce')
    .setFontColor('#000000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // J180:J183 (O to R) -> Blue
  sheet.getRange(180, 10, 4, 1).merge()
    .setValue("Blue\n< 54 %")
    .setFontWeight('bold')
    .setBackground('#0070c0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Row 184: Total Score row formatting
  sheet.getRange(totalRowIndex, 1, 1, 6).merge().setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange(totalRowIndex, 7).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(totalRowIndex, 8).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(totalRowIndex, 9, 1, 2).merge().setFontWeight('bold').setFontSize(12).setBackground('#c6efce').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(totalRowIndex, 26);

  // Scorecard grid borders
  sheet.getRange(scHeaderRowIndex, 1, 20, 10).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Scorecard outer thick blue border (around rows 164 to 184)
  sheet.getRange(legendRowIndex, 1, 21, 10).setBorder(true, true, true, true, null, null, '#002060', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(legendRowIndex, 1, 1, 10).setBorder(true, true, true, true, null, null, '#002060', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Blue vertical line along right edge of Col J
  sheet.getRange(1, 10, totalRowIndex, 1).setBorder(null, null, null, true, null, null, '#002060', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

