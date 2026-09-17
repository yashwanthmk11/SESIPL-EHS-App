function ensureDemoUsers_() {
  const userRows = [
    { employeeId: 'EMP001', uan: 'UAN001', name: 'Site Lead (Intuit)', role: ROLES.LEAD, email: 'harish.ehs@sesipl.com', phone: '+91 98450 44004', active: 'TRUE', mappedProjects: 'PRJ_INTUIT' },
    { employeeId: 'EMP002', uan: 'UAN002', name: 'Asst EHS Manager', role: ROLES.ASST, email: 'asst.mgr@sesipl.com', phone: '+91 98450 22005', active: 'TRUE', mappedProjects: 'PRJ_INTUIT,PRJ_SIEMENS' },
    { employeeId: 'EMP003', uan: 'UAN003', name: 'EHS Manager', role: ROLES.MANAGER, email: 'manager.ehs@sesipl.com', phone: '+91 98450 33003', active: 'TRUE', mappedProjects: '' },
    { employeeId: 'EMP004', uan: 'UAN004', name: 'Director', role: ROLES.DIRECTOR, email: 'director@sesipl.com', phone: '+91 98450 11001', active: 'TRUE', mappedProjects: '' },
    { employeeId: 'EMP005', uan: 'UAN005', name: 'Site Lead (Qualcomm)', role: ROLES.LEAD, email: 'murugan.ehs@sesipl.com', phone: '+91 98450 44009', active: 'TRUE', mappedProjects: 'PRJ_QUALCOMM' }
  ];
  const sh = sheet_(SHEETS.USERS);
  batchWriteObjects_(sh, HEADERS.Users, userRows);
  return rowsToObjects_(SHEETS.USERS);
}

function ensureDemoProjects_() {
  const projects = [
    { id: 'PRJ_INTUIT', code: 'INTUIT', name: 'Intuit', client: 'Intuit', pmc: 'CBRE', inCharge: 'Mr.Harish', manager: 'HR Ravikiran', scope: 'Internal Electrical work (Fit Out)', startDate: '2025-01-01', endDate: '2026-08-31', areaSqft: '389175', poNo: 'C 47344', status: 'RUNNING', region: 'Bangalore', projectDuration: '8 Months' },
    { id: 'PRJ_SIEMENS', code: 'SIEMENS', name: 'Siemens', client: 'Siemens', pmc: 'Cushman & Wakefield', inCharge: 'S. Rajesh', manager: 'K. Sharma', scope: 'Electrical Fit Out & Commissioning', startDate: '2025-06-01', endDate: '2026-05-31', areaSqft: '245000', poNo: 'C 48120', status: 'RUNNING', region: 'Bangalore', projectDuration: '12 Months' },
    { id: 'PRJ_QUALCOMM', code: 'QUALCOMM', name: 'Qualcomm-CH', client: 'Qualcomm', pmc: 'JLL', inCharge: 'V. Murugan', manager: 'HR Ravikiran', scope: 'HV & LV Electrical Installation', startDate: '2025-03-01', endDate: '2026-01-31', areaSqft: '410000', poNo: 'C 49055', status: 'RUNNING', region: 'Chennai', projectDuration: '10 Months' },
    { id: 'PRJ_INFOSYS', code: 'INFOSYS', name: 'Infosys', client: 'Infosys', pmc: 'Turner & Townsend', inCharge: 'A. Reddy', manager: 'K. Sharma', scope: 'Internal Electrical & Substation', startDate: '2025-02-01', endDate: '2026-04-30', areaSqft: '520000', poNo: 'C 50210', status: 'RUNNING', region: 'Hyderabad', projectDuration: '14 Months' }
  ];
  const sh = sheet_(SHEETS.PROJECTS);
  batchWriteObjects_(sh, HEADERS.Projects, projects);
  return rowsToObjects_(SHEETS.PROJECTS);
}

function ensureDemoProjectUsers_() {
  const projectUserRows = [
    { employeeId: 'EMP001', projectId: 'PRJ_INTUIT', role: ROLES.LEAD },
    { employeeId: 'EMP005', projectId: 'PRJ_QUALCOMM', role: ROLES.LEAD },
    { employeeId: 'EMP002', projectId: 'PRJ_INTUIT', role: ROLES.ASST },
    { employeeId: 'EMP002', projectId: 'PRJ_SIEMENS', role: ROLES.ASST }
  ];
  const sh = sheet_(SHEETS.PROJECT_USERS);
  batchWriteObjects_(sh, HEADERS.ProjectUsers, projectUserRows);
  return rowsToObjects_(SHEETS.PROJECT_USERS);
}

function login(employeeId, uan) {
  try {
    const id = String(employeeId || "").trim();
    const pin = String(uan || "").trim();
    if (!id || !pin) return { ok: false, error: "Enter Employee ID and UAN." };

    let users = rowsToObjects_(SHEETS.USERS);
    if (!users || !users.length) {
      try {
        users = ensureDemoUsers_();
      } catch (e) {
        Logger.log('Could not auto-seed demo users: ' + e);
      }
    }
    if (!users || !users.length) {
      return {
        ok: false,
        error: "User database is empty. Please run initializeSystem() in Apps Script editor."
      };
    }

    const user = users.find(
      (u) => String(u.employeeId || "").trim().toUpperCase() === id.toUpperCase(),
    );
    if (!user) return { ok: false, error: "Invalid Employee ID or UAN." };
    if (
      String(user.active || "").toUpperCase() !== "TRUE" &&
      String(user.active || "") !== "1" &&
      String(user.active || "").toUpperCase() !== "YES"
    ) {
      return { ok: false, error: "Account is inactive. Contact EHS Manager." };
    }
    if (String(user.uan || "").trim().toUpperCase() !== pin.toUpperCase()) {
      return { ok: false, error: "Invalid Employee ID or UAN." };
    }

    const token = makeToken_(user);
    const sessionUser = publicUser_(user);
    const sessionStr = JSON.stringify(sessionUser);
    const cacheKey = "sess_" + token.slice(0, 100);
    try {
      CacheService.getScriptCache().put(cacheKey, sessionStr, 14400);
    } catch (e) {}
    try {
      writeAudit_(user.employeeId, "LOGIN", "User", user.employeeId, "");
    } catch (e) {}

    return {
      ok: true,
      token: token,
      user: sessionUser
    };
  } catch (err) {
    Logger.log("login error: " + err);
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function makeToken_(user) {
  const payload = {
    u: String(user.employeeId || "").trim(),
    r: String(user.role || "").trim(),
    t: Date.now()
  };
  return Utilities.base64EncodeWebSafe(JSON.stringify(payload));
}

function parseToken_(token) {
  if (!token) return null;
  try {
    const raw = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
    const payload = JSON.parse(raw);
    if (!payload || !payload.u || !payload.t) return null;
    // Expire after 8 hours (8 * 3600 * 1000 = 28800000 ms)
    if (Date.now() - payload.t > 28800000) return null;
    return payload;
  } catch (e) {
    return null;
  }
}

function logout(token) {
  if (token) {
    try { CacheService.getScriptCache().remove("sess_" + token.slice(0, 100)); } catch (e) {}
  }
  return { ok: true };
}

function requireUser_(token) {
  if (!token) throw new Error("Session expired. Please login again.");

  // 1. Fast in-memory cache check
  const cacheKey = "sess_" + token.slice(0, 100);
  let raw = null;
  try {
    raw = CacheService.getScriptCache().get(cacheKey);
  } catch (e) {}
  if (raw) {
    try { return JSON.parse(raw); } catch (e) {}
  }

  // 2. Decode the self-validating token payload
  const payload = parseToken_(token);
  if (!payload) {
    throw new Error("Session expired. Please login again.");
  }

  // 3. Resolve user from database or fallback demo seed
  let users = rowsToObjects_(SHEETS.USERS);
  if (!users || !users.length) {
    users = ensureDemoUsers_();
  }
  const user = users.find(
    (u) => String(u.employeeId || "").trim().toUpperCase() === payload.u.toUpperCase()
  );
  if (!user) {
    throw new Error("User account not found. Please login again.");
  }

  const sessionUser = publicUser_(user);
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(sessionUser), 14400);
  } catch (e) {}

  return sessionUser;
}

function publicUser_(user) {
  return {
    employeeId: user.employeeId,
    name: user.name,
    role: user.role,
    email: user.email || "",
    phone: user.phone || "",
    mappedProjects: String(user.mappedProjects || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean),
  };
}

function canSeeAllProjects_(role) {
  return role === ROLES.MANAGER || role === ROLES.DIRECTOR;
}

function canApprove_(role) {
  return (
    role === ROLES.ASST || role === ROLES.MANAGER || role === ROLES.DIRECTOR
  );
}

function canMutateAll_(role) {
  return role === ROLES.MANAGER || role === ROLES.DIRECTOR;
}

function canSubmitForm_(role, entryType) {
  if (entryType === "UPLOAD") return canManageOneTimeUploads_(role);
  return role === ROLES.LEAD || role === ROLES.ASST || canMutateAll_(role);
}

function canManageOneTimeUploads_(role) {
  return (
    role === ROLES.ASST || role === ROLES.MANAGER || role === ROLES.DIRECTOR
  );
}

function canUploadManagedFile_(role) {
  return role === ROLES.ASST || canMutateAll_(role);
}

function canManageProjects_(role) {
  return (
    role === ROLES.ASST || role === ROLES.MANAGER || role === ROLES.DIRECTOR
  );
}

function scopedProjectIds_(user) {
  if (canSeeAllProjects_(user.role)) {
    let prjs = rowsToObjects_(SHEETS.PROJECTS);
    if (!prjs.length) prjs = ensureDemoProjects_();
    return prjs.map((p) => p.id);
  }
  const fromUser = user.mappedProjects || [];
  let mapRows = rowsToObjects_(SHEETS.PROJECT_USERS);
  if (!mapRows.length && !fromUser.length) {
    mapRows = ensureDemoProjectUsers_();
  }
  const fromMap = mapRows
    .filter((r) => String(r.employeeId) === String(user.employeeId))
    .map((r) => r.projectId);
  return Array.from(new Set(fromUser.concat(fromMap)));
}

function assertProjectAccess_(user, projectId) {
  const allowed = scopedProjectIds_(user);
  if (allowed.indexOf(projectId) < 0)
    throw new Error("No access to this project.");
}

function assertEdit_(user, projectId) {
  assertProjectAccess_(user, projectId);
  if (user.role === ROLES.LEAD) return;
  if (user.role === ROLES.ASST || canMutateAll_(user.role)) return;
  throw new Error("You do not have edit access.");
}

function assertDelete_(user, projectId) {
  assertProjectAccess_(user, projectId);
  if (user.role === ROLES.ASST || canMutateAll_(user.role)) return;
  throw new Error("You do not have delete access.");
}
