/**
 * Hourly trigger created by initializeSystem().
 * Hour 0: mail responsible lead.
 * Hour 25: if still OPEN and no fallback, escalate to Manager / Director.
 */
function checkEscalations() {
  const now = new Date();
  const observations = rowsToObjects_(SHEETS.OBSERVATIONS);
  observations.forEach((obs) => {
    if (obs.status !== STATUS.OPEN && obs.status !== STATUS.FALLBACK) return;
    const due = obs.dueAt ? new Date(obs.dueAt.replace(" ", "T")) : null;
    if (!due || isNaN(due.getTime())) return;

    const already = rowsToObjects_(SHEETS.ESCALATIONS).filter(
      (e) => e.entityId === obs.id,
    );
    const hasL1 = already.some((e) => String(e.level) === "1");
    const hasL2 = already.some((e) => String(e.level) === "2");

    if (!hasL1) {
      const owner = obs.ownerEmployeeId || leadForProject_(obs.projectId);
      mailUser_(
        owner,
        "Action required within 24 hours",
        "Open EHS activity / observation needs closure.\n\n" +
          (obs.observation || "") +
          "\nDue: " +
          obs.dueAt,
      );
      pushNotify_(
        owner,
        obs.projectId,
        { employeeId: "SYSTEM" },
        "24h action required",
        obs.observation || "",
        "ESCALATION",
      );
      appendRow_(SHEETS.ESCALATIONS, {
        id: uid_("ESC"),
        entityType: "Observation",
        entityId: obs.id,
        level: 1,
        toEmployeeId: owner || "",
        sentAt: nowIso_(),
        reason: "INITIAL_24H",
      });
    }

    const escalateAfter = new Date(due.getTime() + 60 * 60 * 1000);
    if (now >= escalateAfter && obs.status === STATUS.OPEN && !hasL2) {
      notifyHigherManager_(
        obs,
        "Escalation: unresolved after 24 hours",
        "Lead did not close or write a fallback review. Auto-escalated at hour 25.",
        { employeeId: "SYSTEM" },
      );
      appendRow_(SHEETS.ESCALATIONS, {
        id: uid_("ESC"),
        entityType: "Observation",
        entityId: obs.id,
        level: 2,
        toEmployeeId: "MANAGERS",
        sentAt: nowIso_(),
        reason: "HOUR_25",
      });
    }
  });

  const pending = rowsToObjects_(SHEETS.SUBMISSIONS).filter(
    (s) => s.status === STATUS.SUBMITTED || s.status === STATUS.RESUBMITTED,
  );
  pending.forEach((s) => {
    const def = FORM_DEFS.find((f) => f.formCode === s.formCode);
    if (!def || !def.slaHours) return;
    const submitted = s.submittedAt
      ? new Date(s.submittedAt.replace(" ", "T"))
      : null;
    if (!submitted) return;
    const ageH = (now - submitted) / 36e5;
    if (ageH < Number(def.slaHours) + 1) return;
    const already = rowsToObjects_(SHEETS.ESCALATIONS).filter(
      (e) => e.entityId === s.id && String(e.level) === "2",
    );
    if (already.length) return;
    const fake = {
      projectId: s.projectId,
      observation: "Pending approval for " + s.formCode,
      id: s.id,
    };
    notifyHigherManager_(
      fake,
      "Pending approval overdue",
      s.formCode + " waiting more than SLA.",
      { employeeId: "SYSTEM" },
    );
    appendRow_(SHEETS.ESCALATIONS, {
      id: uid_("ESC"),
      entityType: "Submission",
      entityId: s.id,
      level: 2,
      toEmployeeId: "MANAGERS",
      sentAt: nowIso_(),
      reason: "APPROVAL_SLA",
    });
  });
}
