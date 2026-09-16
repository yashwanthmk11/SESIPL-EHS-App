function login(employeeId, uan) {
  const id = String(employeeId || "").trim();
  const pin = String(uan || "").trim();
  if (!id || !pin) return { ok: false, error: "Enter Employee ID and UAN." };

  let users = rowsToObjects_(SHEETS.USERS);
  if (!users.length) {
    try {
      initializeSystem();
    } catch (e) {
      return {
        ok: false,
        error: "User database is not initialized or cannot be opened. Share the configured spreadsheet with the Apps Script owner, run initializeSystem(), and try again."
      };
    }
    users = rowsToObjects_(SHEETS.USERS);
  }
  const user = users.find(
    (u) => String(u.employeeId).trim().toUpperCase() === id.toUpperCase(),
  );
  if (!user) return { ok: false, error: "Invalid Employee ID or UAN." };
  if (
    String(user.active).toUpperCase() !== "TRUE" &&
    String(user.active) !== "1" &&
    String(user.active).toUpperCase() !== "YES"
  ) {
    return { ok: false, error: "Account is inactive. Contact EHS Manager." };
  }
  if (String(user.uan).trim().toUpperCase() !== pin.toUpperCase()) {
    return { ok: false, error: "Invalid Employee ID or UAN." };
  }

  const token = Utilities.getUuid();
  const sessionUser = publicUser_(user);
  const sessionStr = JSON.stringify(sessionUser);
  try {
    CacheService.getScriptCache().put("sess_" + token, sessionStr, SESSION_TTL_SEC);
  } catch (e) {}
  try {
    PropertiesService.getUserProperties().setProperty("sess_" + token, sessionStr);
  } catch (e) {}
  writeAudit_(user.employeeId, "LOGIN", "User", user.employeeId, "");
  return {
    ok: true,
    token: token,
    user: sessionUser,
    catalog: FORM_DEFS,
    modules: MODULES,
  };
}

function logout(token) {
  if (token) {
    try { CacheService.getScriptCache().remove("sess_" + token); } catch (e) {}
    try { PropertiesService.getUserProperties().deleteProperty("sess_" + token); } catch (e) {}
  }
  return { ok: true };
}

function requireUser_(token) {
  if (!token) throw new Error("Session expired. Please login again.");
  let raw = null;
  try {
    raw = CacheService.getScriptCache().get("sess_" + token);
  } catch (e) {}
  if (!raw) {
    try {
      raw = PropertiesService.getUserProperties().getProperty("sess_" + token);
    } catch (e) {}
    if (raw) {
      try {
        CacheService.getScriptCache().put("sess_" + token, raw, SESSION_TTL_SEC);
      } catch (e) {}
    }
  }
  if (!raw) throw new Error("Session expired. Please login again.");
  const user = JSON.parse(raw);
  return user;
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

function scopedProjectIds_(user) {
  if (canSeeAllProjects_(user.role)) {
    return rowsToObjects_(SHEETS.PROJECTS).map((p) => p.id);
  }
  const fromUser = user.mappedProjects || [];
  const fromMap = rowsToObjects_(SHEETS.PROJECT_USERS)
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
