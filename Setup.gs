/**
 * Run once from the Apps Script editor: initializeSystem()
 * Creates the database spreadsheet, Drive folders, seed users and demo project.
 */
function initializeSystem() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const configuredId = String(DATABASE_SPREADSHEET_ID || "").trim();
    let ssId = configuredId || props.getProperty("SPREADSHEET_ID");
    let ss;
    if (ssId) {
      try {
        ss = SpreadsheetApp.openById(ssId);
      } catch (e) {
        if (configuredId) {
          throw new Error(
            "Cannot open configured database spreadsheet " +
              configuredId +
              ". Share it with the Apps Script owner and run initializeSystem() again. Details: " +
              e,
          );
        }
        Logger.log(
          "Stored database spreadsheet could not be opened; creating a replacement. " +
            e,
        );
        ssId = "";
      }
    }
    if (!ss) {
      ss = SpreadsheetApp.create(APP_NAME + " Database");
      ssId = ss.getId();
      props.setProperty("SPREADSHEET_ID", ssId);
    }
    Logger.log("Database spreadsheet: " + ss.getUrl());

    // 1. CLEAN PURGE & REBUILD FROM SCRATCH
    // Ensure temporary scratch sheet exists so the spreadsheet is never empty during wipe
    let tempSheet = ss.getSheetByName("__TEMP_SCRATCH__");
    if (!tempSheet) {
      tempSheet = ss.insertSheet("__TEMP_SCRATCH__");
    }

    // Delete all other sheets safely
    const sheetsToPurge = ss.getSheets();
    for (let i = 0; i < sheetsToPurge.length; i++) {
      const s = sheetsToPurge[i];
      let sName = "";
      try {
        sName = s.getName();
      } catch (err) {
        continue;
      }
      if (sName && sName !== "__TEMP_SCRATCH__") {
        try {
          ss.deleteSheet(s);
          Logger.log("Purged sheet: " + sName);
        } catch (e) {
          Logger.log("Could not purge sheet " + sName + ": " + e);
        }
      }
    }

    // 2. Recreate all canonical sheets fresh from HEADERS definition
    const sheetMap = {};
    Object.keys(HEADERS).forEach((name) => {
      let sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
      } else {
        sh.clear();
      }
      sheetMap[name] = sh;
      const cols = HEADERS[name];
      const headerLabels = cols.map((h) => formatHeaderLabel_(h));
      sh.getRange(1, 1, 1, cols.length).setValues([headerLabels]);
    });

    // 3. Ensure Root Drive Folder exists
    let rootId = props.getProperty("ROOT_FOLDER_ID");
    let root;
    if (rootId) {
      try {
        root = DriveApp.getFolderById(rootId);
      } catch (e) {
        Logger.log(
          "Stored root folder could not be opened; creating a replacement. " +
            e,
        );
        rootId = "";
      }
    }
    if (!root) {
      root = DriveApp.createFolder(APP_NAME + " Files");
      rootId = root.getId();
      props.setProperty("ROOT_FOLDER_ID", rootId);
    }
    Logger.log("Root folder: " + root.getUrl());

    // 4. Canonical Demo Projects & Users
    const p1 = "PRJ_INTUIT";
    const p2 = "PRJ_SIEMENS";
    const p3 = "PRJ_QUALCOMM";
    const p4 = "PRJ_INFOSYS";

    const userRows = [
      {
        employeeId: "EMP001",
        uan: "UAN001",
        name: "Site Lead (Intuit)",
        role: ROLES.LEAD,
        email: "harish.ehs@sesipl.com",
        phone: "+91 98450 44004",
        active: "TRUE",
        mappedProjects: p1,
      },
      {
        employeeId: "EMP002",
        uan: "UAN002",
        name: "Asst EHS Manager",
        role: ROLES.ASST,
        email: "asst.mgr@sesipl.com",
        phone: "+91 98450 22005",
        active: "TRUE",
        mappedProjects: p1 + "," + p2,
      },
      {
        employeeId: "EMP003",
        uan: "UAN003",
        name: "EHS Manager",
        role: ROLES.MANAGER,
        email: "manager.ehs@sesipl.com",
        phone: "+91 98450 33003",
        active: "TRUE",
        mappedProjects: "",
      },
      {
        employeeId: "EMP004",
        uan: "UAN004",
        name: "Director",
        role: ROLES.DIRECTOR,
        email: "director@sesipl.com",
        phone: "+91 98450 11001",
        active: "TRUE",
        mappedProjects: "",
      },
      {
        employeeId: "EMP005",
        uan: "UAN005",
        name: "Site Lead (Qualcomm)",
        role: ROLES.LEAD,
        email: "murugan.ehs@sesipl.com",
        phone: "+91 98450 44009",
        active: "TRUE",
        mappedProjects: p3,
      },
    ];

    const projects = [
      {
        id: p1,
        code: "INTUIT",
        name: "Intuit",
        client: "Intuit",
        pmc: "CBRE",
        inCharge: "Mr.Harish",
        manager: "HR Ravikiran",
        scope: "Internal Electrical work (Fit Out)",
        startDate: "2025-01-01",
        endDate: "2026-08-31",
        areaSqft: "389175",
        poNo: "C 47344",
        status: "RUNNING",
        region: "Bangalore",
        projectDuration: "8 Months",
      },
      {
        id: p2,
        code: "SIEMENS",
        name: "Siemens",
        client: "Siemens",
        pmc: "Cushman & Wakefield",
        inCharge: "S. Rajesh",
        manager: "K. Sharma",
        scope: "Electrical Fit Out & Commissioning",
        startDate: "2025-06-01",
        endDate: "2026-05-31",
        areaSqft: "245000",
        poNo: "C 48120",
        status: "RUNNING",
        region: "Bangalore",
        projectDuration: "12 Months",
      },
      {
        id: p3,
        code: "QUALCOMM",
        name: "Qualcomm-CH",
        client: "Qualcomm",
        pmc: "JLL",
        inCharge: "V. Murugan",
        manager: "HR Ravikiran",
        scope: "HV & LV Electrical Installation",
        startDate: "2025-03-01",
        endDate: "2026-01-31",
        areaSqft: "410000",
        poNo: "C 49055",
        status: "RUNNING",
        region: "Chennai",
        projectDuration: "10 Months",
      },
      {
        id: p4,
        code: "INFOSYS",
        name: "Infosys",
        client: "Infosys",
        pmc: "Turner & Townsend",
        inCharge: "A. Reddy",
        manager: "K. Sharma",
        scope: "Internal Electrical & Substation",
        startDate: "2025-02-01",
        endDate: "2026-04-30",
        areaSqft: "520000",
        poNo: "C 50210",
        status: "RUNNING",
        region: "Hyderabad",
        projectDuration: "14 Months",
      },
    ];

    const projectUserRows = [
      { employeeId: "EMP001", projectId: p1, role: ROLES.LEAD },
      { employeeId: "EMP005", projectId: p3, role: ROLES.LEAD },
      { employeeId: "EMP002", projectId: p1, role: ROLES.ASST },
      { employeeId: "EMP002", projectId: p2, role: ROLES.ASST },
    ];

    // Slide 5 Project Org Hierarchy
    const orgSeeds = [
      {
        id: uid_("ORG"),
        projectId: p1,
        level: "1",
        roleTitle: "Director",
        name: "Mr. Shankar",
        phone: "+91 98450 11001",
        email: "director@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p1,
        level: "2",
        roleTitle: "General Manager",
        name: "Mr. C. Shekhar",
        phone: "+91 98450 22002",
        email: "gm@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p1,
        level: "3",
        roleTitle: "Project Manager",
        name: "HR Ravikiran",
        phone: "+91 98450 33003",
        email: "ravikiran@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p1,
        level: "4",
        roleTitle: "HSE Officer / In charge",
        name: "Mr. Harish",
        phone: "+91 98450 44004",
        email: "harish.ehs@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p1,
        level: "5",
        roleTitle: "Project Engineer",
        name: "P. Naveen",
        phone: "+91 98450 55005",
        email: "naveen.pe@sesipl.com",
      },

      {
        id: uid_("ORG"),
        projectId: p2,
        level: "1",
        roleTitle: "Director",
        name: "Mr. Shankar",
        phone: "+91 98450 11001",
        email: "director@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p2,
        level: "2",
        roleTitle: "General Manager",
        name: "Mr. C. Shekhar",
        phone: "+91 98450 22002",
        email: "gm@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p2,
        level: "3",
        roleTitle: "Project Manager",
        name: "K. Sharma",
        phone: "+91 98450 33006",
        email: "sharma@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p2,
        level: "4",
        roleTitle: "HSE Officer / In charge",
        name: "S. Rajesh",
        phone: "+91 98450 44007",
        email: "rajesh.ehs@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p2,
        level: "5",
        roleTitle: "Project Engineer",
        name: "M. Karthik",
        phone: "+91 98450 55008",
        email: "karthik.pe@sesipl.com",
      },

      {
        id: uid_("ORG"),
        projectId: p3,
        level: "1",
        roleTitle: "Director",
        name: "Mr. Shankar",
        phone: "+91 98450 11001",
        email: "director@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p3,
        level: "2",
        roleTitle: "General Manager",
        name: "Mr. C. Shekhar",
        phone: "+91 98450 22002",
        email: "gm@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p3,
        level: "3",
        roleTitle: "Project Manager",
        name: "HR Ravikiran",
        phone: "+91 98450 33003",
        email: "ravikiran@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p3,
        level: "4",
        roleTitle: "HSE Officer / In charge",
        name: "V. Murugan",
        phone: "+91 98450 44009",
        email: "murugan.ehs@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p3,
        level: "5",
        roleTitle: "Project Engineer",
        name: "S. Vignesh",
        phone: "+91 98450 55010",
        email: "vignesh.pe@sesipl.com",
      },

      {
        id: uid_("ORG"),
        projectId: p4,
        level: "1",
        roleTitle: "Director",
        name: "Mr. Shankar",
        phone: "+91 98450 11001",
        email: "director@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p4,
        level: "2",
        roleTitle: "General Manager",
        name: "Mr. C. Shekhar",
        phone: "+91 98450 22002",
        email: "gm@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p4,
        level: "3",
        roleTitle: "Project Manager",
        name: "K. Sharma",
        phone: "+91 98450 33006",
        email: "sharma@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p4,
        level: "4",
        roleTitle: "HSE Officer / In charge",
        name: "A. Reddy",
        phone: "+91 98450 44011",
        email: "reddy.ehs@sesipl.com",
      },
      {
        id: uid_("ORG"),
        projectId: p4,
        level: "5",
        roleTitle: "Project Engineer",
        name: "T. Venkat",
        phone: "+91 98450 55012",
        email: "venkat.pe@sesipl.com",
      },
    ];

    // Slide 8 Standard Daily Logs
    const logs = [
      {
        id: uid_("DL"),
        projectId: p1,
        date: "2026-09-15",
        staff: "6",
        workers: "48",
        totalManpower: "54",
        workingHours: "8",
        totalManHours: "432",
        safeManHours: "432",
        cumSafeManHours: "370840",
        inductions: "5",
        indStaff: "1",
        indWorkers: "4",
        tbtCount: "3",
        tbtPersons: "48",
        trainingTopic: "Earth pit excavation & isolation",
        trainingPersons: "18",
        permitHot: "2",
        permitElectrical: "4",
        permitCold: "1",
        permitGeneral: "3",
        permitOthers: "0",
        firstAid: "0",
        nearMiss: "0",
        ltiCount: "0",
        accidentDetails: "Nil - Safe day",
        remarks: "Good site housekeeping maintained",
        enteredBy: "EMP001",
        verifiedBy: "Ravikiran",
        approvedBy: "Mr. Shankar",
      },
      {
        id: uid_("DL"),
        projectId: p2,
        date: "2026-09-15",
        staff: "4",
        workers: "38",
        totalManpower: "42",
        workingHours: "8",
        totalManHours: "336",
        safeManHours: "336",
        cumSafeManHours: "215420",
        inductions: "3",
        indStaff: "0",
        indWorkers: "3",
        tbtCount: "2",
        tbtPersons: "38",
        trainingTopic: "Cable laying and tray installation",
        trainingPersons: "14",
        permitHot: "1",
        permitElectrical: "3",
        permitCold: "0",
        permitGeneral: "2",
        permitOthers: "0",
        firstAid: "0",
        nearMiss: "0",
        ltiCount: "0",
        accidentDetails: "Nil - Safe day",
        remarks: "Cable puller PPE inspected",
        enteredBy: "EMP002",
        verifiedBy: "K. Sharma",
        approvedBy: "Mr. Shankar",
      },
      {
        id: uid_("DL"),
        projectId: p3,
        date: "2026-09-15",
        staff: "8",
        workers: "60",
        totalManpower: "68",
        workingHours: "8",
        totalManHours: "544",
        safeManHours: "544",
        cumSafeManHours: "410290",
        inductions: "6",
        indStaff: "2",
        indWorkers: "4",
        tbtCount: "4",
        tbtPersons: "60",
        trainingTopic: "Working at height & full-body harness",
        trainingPersons: "24",
        permitHot: "3",
        permitElectrical: "5",
        permitCold: "2",
        permitGeneral: "4",
        permitOthers: "1",
        firstAid: "1",
        nearMiss: "1",
        ltiCount: "0",
        accidentDetails: "Minor finger scrape; treated with first aid kit",
        remarks: "First aid kit replenished",
        enteredBy: "EMP005",
        verifiedBy: "Ravikiran",
        approvedBy: "Mr. Shankar",
      },
      {
        id: uid_("DL"),
        projectId: p4,
        date: "2026-09-15",
        staff: "10",
        workers: "70",
        totalManpower: "80",
        workingHours: "8",
        totalManHours: "640",
        safeManHours: "640",
        cumSafeManHours: "280150",
        inductions: "8",
        indStaff: "1",
        indWorkers: "7",
        tbtCount: "4",
        tbtPersons: "70",
        trainingTopic: "Substation panel commissioning & LOTO",
        trainingPersons: "28",
        permitHot: "2",
        permitElectrical: "6",
        permitCold: "1",
        permitGeneral: "5",
        permitOthers: "0",
        firstAid: "0",
        nearMiss: "0",
        ltiCount: "0",
        accidentDetails: "Nil - Safe day",
        remarks: "Substation room entry controlled",
        enteredBy: "EMP003",
        verifiedBy: "K. Sharma",
        approvedBy: "Mr. Shankar",
      },
    ];

    // Slide 11 Master Tracker of Scaffolding & Ladders
    const scf = [
      {
        id: uid_("SCF"),
        projectId: p1,
        date: "2026-09-15",
        region: "Bangalore",
        sharavFab: "10",
        sesipl: "4",
        vinayaka: "2",
        hbs: "3",
        other: "0",
        totalScaffold: "14",
        ladderSharav: "12",
        ladderSesipl: "0",
        ladderVinayaka: "0",
        ladderHbs: "0",
        ladderRental: "0",
        airportLadder: "2",
        workstationLadder: "2",
        frpMsafe: "3",
        frpYoungman: "2",
        totalLadder: "9",
        returnedScaffold: "5",
        returnedLadder: "3",
        remarks: "14 Scaffolds in site; 9 Ladders in site",
        enteredBy: "EMP001",
      },
      {
        id: uid_("SCF"),
        projectId: p2,
        date: "2026-09-15",
        region: "Bangalore",
        sharavFab: "6",
        sesipl: "3",
        vinayaka: "1",
        hbs: "2",
        other: "0",
        totalScaffold: "12",
        ladderSharav: "6",
        ladderSesipl: "2",
        ladderVinayaka: "0",
        ladderHbs: "0",
        ladderRental: "2",
        airportLadder: "1",
        workstationLadder: "1",
        frpMsafe: "2",
        frpYoungman: "2",
        totalLadder: "8",
        returnedScaffold: "2",
        returnedLadder: "1",
        remarks: "Periodic inspection completed",
        enteredBy: "EMP002",
      },
      {
        id: uid_("SCF"),
        projectId: p3,
        date: "2026-09-15",
        region: "Chennai",
        sharavFab: "8",
        sesipl: "4",
        vinayaka: "1",
        hbs: "2",
        other: "0",
        totalScaffold: "15",
        ladderSharav: "0",
        ladderSesipl: "0",
        ladderVinayaka: "0",
        ladderHbs: "0",
        ladderRental: "0",
        airportLadder: "0",
        workstationLadder: "0",
        frpMsafe: "0",
        frpYoungman: "0",
        totalLadder: "0",
        returnedScaffold: "0",
        returnedLadder: "0",
        remarks: "15 Scaffolds in site",
        enteredBy: "EMP005",
      },
      {
        id: uid_("SCF"),
        projectId: p4,
        date: "2026-09-15",
        region: "Hyderabad",
        sharavFab: "5",
        sesipl: "2",
        vinayaka: "3",
        hbs: "0",
        other: "0",
        totalScaffold: "10",
        ladderSharav: "6",
        ladderSesipl: "0",
        ladderVinayaka: "0",
        ladderHbs: "0",
        ladderRental: "8",
        airportLadder: "2",
        workstationLadder: "2",
        frpMsafe: "4",
        frpYoungman: "4",
        totalLadder: "14",
        returnedScaffold: "0",
        returnedLadder: "0",
        remarks: "10 Scaffolds in site; 14 Ladders in site",
        enteredBy: "EMP003",
      },
    ];

    // Slide 10 PPE Stock Register Data
    const ppeSeeds = [
      {
        id: uid_("PPE"),
        projectId: p1,
        date: "2026-09-15",
        contractor: "Vinayaka Electricals",
        receivedBy: "R. Prakash",
        helmetWhite: "3",
        helmetGreen: "2",
        helmetBlue: "3",
        helmetRed: "2",
        jacketGreen: "3",
        jacketOrange: "4",
        cottonGloves: "20",
        leatherGloves: "8",
        goggle: "15",
        faceShield: "4",
        mask: "30",
        apron: "2",
        shoulderPad: "2",
        earMuff: "4",
        dcNo: "274",
        totalReceived: "10",
        returnable: "Yes",
        balanceStock: "8",
        remarks: "Store In-charge: Mr. Harish",
        enteredBy: "EMP001",
      },
      {
        id: uid_("PPE"),
        projectId: p2,
        date: "2026-09-15",
        contractor: "Apex Infra",
        receivedBy: "S. Kumar",
        helmetWhite: "4",
        helmetGreen: "2",
        helmetBlue: "2",
        helmetRed: "0",
        jacketGreen: "2",
        jacketOrange: "4",
        cottonGloves: "15",
        leatherGloves: "6",
        goggle: "10",
        faceShield: "2",
        mask: "25",
        apron: "0",
        shoulderPad: "0",
        earMuff: "2",
        dcNo: "281",
        totalReceived: "12",
        returnable: "Yes",
        balanceStock: "9",
        remarks: "Store In-charge: S. Rajesh",
        enteredBy: "EMP002",
      },
      {
        id: uid_("PPE"),
        projectId: p3,
        date: "2026-09-15",
        contractor: "Sharav Power",
        receivedBy: "T. Murthy",
        helmetWhite: "5",
        helmetGreen: "3",
        helmetBlue: "4",
        helmetRed: "2",
        jacketGreen: "4",
        jacketOrange: "6",
        cottonGloves: "30",
        leatherGloves: "12",
        goggle: "20",
        faceShield: "6",
        mask: "40",
        apron: "4",
        shoulderPad: "4",
        earMuff: "6",
        dcNo: "295",
        totalReceived: "18",
        returnable: "Yes",
        balanceStock: "14",
        remarks: "Store In-charge: V. Murugan",
        enteredBy: "EMP005",
      },
      {
        id: uid_("PPE"),
        projectId: p4,
        date: "2026-09-15",
        contractor: "Southern Tech",
        receivedBy: "G. Naidu",
        helmetWhite: "4",
        helmetGreen: "4",
        helmetBlue: "3",
        helmetRed: "1",
        jacketGreen: "5",
        jacketOrange: "5",
        cottonGloves: "25",
        leatherGloves: "10",
        goggle: "18",
        faceShield: "4",
        mask: "35",
        apron: "2",
        shoulderPad: "2",
        earMuff: "5",
        dcNo: "304",
        totalReceived: "15",
        returnable: "Yes",
        balanceStock: "11",
        remarks: "Store In-charge: A. Reddy",
        enteredBy: "EMP003",
      },
    ];

    // Slide 15 Audit Scores
    const auditSeeds = [
      {
        id: uid_("ADT"),
        projectId: p1,
        auditDate: "2026-09-10",
        auditor: "CBRE Lead Auditor",
        location: "Intuit Bellandur",
        totalScore: "485",
        maxScore: "500",
        percent: "97",
        grade: "A+",
        status: "APPROVED",
        submittedBy: "EMP003",
      },
      {
        id: uid_("ADT"),
        projectId: p2,
        auditDate: "2026-09-08",
        auditor: "Cushman Auditor",
        location: "Siemens E-City",
        totalScore: "460",
        maxScore: "500",
        percent: "92",
        grade: "A",
        status: "APPROVED",
        submittedBy: "EMP003",
      },
      {
        id: uid_("ADT"),
        projectId: p3,
        auditDate: "2026-09-12",
        auditor: "JLL EHS Auditor",
        location: "Qualcomm Chennai",
        totalScore: "475",
        maxScore: "500",
        percent: "95",
        grade: "A",
        status: "APPROVED",
        submittedBy: "EMP003",
      },
      {
        id: uid_("ADT"),
        projectId: p4,
        auditDate: "2026-09-05",
        auditor: "Turner Auditor",
        location: "Infosys Hyderabad",
        totalScore: "450",
        maxScore: "500",
        percent: "90",
        grade: "A",
        status: "APPROVED",
        submittedBy: "EMP003",
      },
    ];

    // Audit Section Scores Breakdown
    const auditScoreSeeds = [
      {
        id: uid_("AS"),
        auditId: auditSeeds[0].id,
        section: "A",
        sn: "1",
        particulars: "EHS Policy Displayed at Site Entrance",
        score: "5",
        remarks: "Bilingual display in English and Kannada",
      },
      {
        id: uid_("AS"),
        auditId: auditSeeds[0].id,
        section: "A",
        sn: "2",
        particulars: "EHS Organization Chart with Emergency Contacts",
        score: "5",
        remarks: "Verified contact numbers active",
      },
      {
        id: uid_("AS"),
        auditId: auditSeeds[0].id,
        section: "H",
        sn: "1",
        particulars: "Drinking Water & Rest Area Facilities",
        score: "5",
        remarks: "RO water dispenser clean with test report",
      },
      {
        id: uid_("AS"),
        auditId: auditSeeds[0].id,
        section: "P",
        sn: "1",
        particulars: "Hot Work Permit & Fire Watcher Deployed",
        score: "5",
        remarks: "Fire extinguisher inspected with green tag",
      },
      {
        id: uid_("AS"),
        auditId: auditSeeds[1].id,
        section: "A",
        sn: "1",
        particulars: "EHS Policy Displayed at Site Entrance",
        score: "5",
        remarks: "Displayed at safety induction room",
      },
      {
        id: uid_("AS"),
        auditId: auditSeeds[2].id,
        section: "N",
        sn: "1",
        particulars: "Scaffold Inspection & Tagging System",
        score: "5",
        remarks: "All green tagged and signed",
      },
    ];

    // Seed Realistic Observations
    const obsSeeds = [
      {
        id: uid_("OBS"),
        projectId: p1,
        reportNo: "OBS-001",
        date: "2026-09-15",
        vendor: "Vinayaka Electricals",
        auditedBy: "Harish (Lead)",
        location: "Floor 3 - DB Room",
        contractor: "Vinayaka Electricals",
        observation: "Work at height without double lanyard safety harness",
        preventive:
          "Issued warning and provided standard EN-361 full body harness with shock absorber",
        status: "OPEN",
        ownerEmployeeId: "EMP001",
        dueAt: "2026-09-17 18:00",
        escalateAt: "2026-09-18 18:00",
        fallbackNote: "",
        fallbackAt: "",
        unsafeFileId: "",
        rectifiedFileId: "",
        submittedBy: "EMP001",
      },
      {
        id: uid_("OBS"),
        projectId: p1,
        reportNo: "OBS-002",
        date: "2026-09-14",
        vendor: "Apex Infra",
        auditedBy: "Harish (Lead)",
        location: "Basement Substation",
        contractor: "Apex Infra",
        observation:
          "Temporary cables laid across pedestrian walkway without rubber ramp protection",
        preventive:
          "Rerouted overhead and installed heavy duty cable ramp protector",
        status: "APPROVED",
        ownerEmployeeId: "EMP001",
        dueAt: "2026-09-15 18:00",
        escalateAt: "2026-09-16 18:00",
        fallbackNote: "",
        fallbackAt: "",
        unsafeFileId: "",
        rectifiedFileId: "",
        submittedBy: "EMP001",
      },
      {
        id: uid_("OBS"),
        projectId: p2,
        reportNo: "OBS-003",
        date: "2026-09-15",
        vendor: "Apex Infra",
        auditedBy: "S. Rajesh (Lead)",
        location: "Block B Ground Floor",
        contractor: "Apex Infra",
        observation:
          "Portable angle grinder guard removed during conduit cutting",
        preventive:
          "Work stopped immediately; wheel guard and deadman switch re-installed",
        status: "SUBMITTED",
        ownerEmployeeId: "EMP002",
        dueAt: "2026-09-16 18:00",
        escalateAt: "2026-09-17 18:00",
        fallbackNote: "",
        fallbackAt: "",
        unsafeFileId: "",
        rectifiedFileId: "",
        submittedBy: "EMP002",
      },
      {
        id: uid_("OBS"),
        projectId: p3,
        reportNo: "OBS-004",
        date: "2026-09-13",
        vendor: "Sharav Power",
        auditedBy: "V. Murugan (Lead)",
        location: "Terrace Cable Tray",
        contractor: "Sharav Power",
        observation: "Static lifeline slackness noticed across 12m span",
        preventive:
          "Tensioned lifeline with turnbuckle and verified anchor points",
        status: "CLOSED",
        ownerEmployeeId: "EMP005",
        dueAt: "2026-09-14 18:00",
        escalateAt: "2026-09-15 18:00",
        fallbackNote: "",
        fallbackAt: "",
        unsafeFileId: "",
        rectifiedFileId: "",
        submittedBy: "EMP005",
      },
    ];

    // Seed Interactive Notifications for Mobile-Style Notification Tray
    const notifSeeds = [
      {
        id: uid_("NTF"),
        toEmployeeId: "EMP001",
        projectId: p1,
        title: "Safety Audit Score: 97% (Grade A+)",
        body: "Intuit Bellandur scored 97% on client monthly audit.",
        type: "AUDIT",
        read: "FALSE",
        createdAt: nowIso_(),
        createdBy: "EMP003",
      },
      {
        id: uid_("NTF"),
        toEmployeeId: "EMP001",
        projectId: p1,
        title: "Observation OBS-001 Open",
        body: "Floor 3 DB Room safety observation pending corrective action verification.",
        type: "OBSERVATION",
        read: "FALSE",
        createdAt: nowIso_(),
        createdBy: "EMP001",
      },
      {
        id: uid_("NTF"),
        toEmployeeId: "EMP002",
        projectId: p2,
        title: "Weekly EHS Report Review",
        body: "Siemens E-City weekly report ready for Asst Manager review.",
        type: "REPORT",
        read: "FALSE",
        createdAt: nowIso_(),
        createdBy: "EMP003",
      },
      {
        id: uid_("NTF"),
        toEmployeeId: "EMP003",
        projectId: p3,
        title: "Scaffold Inspection Passed",
        body: "Qualcomm Chennai 15 scaffolds certified green tagged.",
        type: "SCAFFOLD",
        read: "FALSE",
        createdAt: nowIso_(),
        createdBy: "EMP005",
      },
      {
        id: uid_("NTF"),
        toEmployeeId: "EMP004",
        projectId: p1,
        title: "Zero LTI Safety Milestone",
        body: "Cumulative safe man-hours reached 370,840 hrs with zero lost time injuries.",
        type: "MILESTONE",
        read: "FALSE",
        createdAt: nowIso_(),
        createdBy: "SYSTEM",
      },
    ];

    // Seed Training Calendar
    const trainingSeeds = [
      {
        id: uid_("TRN"),
        projectId: p1,
        date: "2026-09-18",
        topic: "LOTO (Lockout / Tagout) Procedures",
        owner: "Harish",
        dept: "Electrical Fit-Out",
        status: "SCHEDULED",
        notes: "Mandatory for all panel technicians",
        createdBy: "EMP001",
      },
      {
        id: uid_("TRN"),
        projectId: p2,
        date: "2026-09-20",
        topic: "Emergency Evacuation & Fire Drill",
        owner: "S. Rajesh",
        dept: "Site Wide",
        status: "SCHEDULED",
        notes: "Assembly point B coordination",
        createdBy: "EMP002",
      },
      {
        id: uid_("TRN"),
        projectId: p3,
        date: "2026-09-22",
        topic: "Height Safety & Fall Protection",
        owner: "V. Murugan",
        dept: "Cable Tray Team",
        status: "SCHEDULED",
        notes: "Inspection of lifelines and harnesses",
        createdBy: "EMP005",
      },
    ];

    // 5. BATCH SEED DATA WRITING (Lightning Fast Single-Call setValues)
    batchWriteObjects_(sheetMap[SHEETS.USERS], HEADERS.Users, userRows);
    batchWriteObjects_(sheetMap[SHEETS.PROJECTS], HEADERS.Projects, projects);
    batchWriteObjects_(
      sheetMap[SHEETS.PROJECT_USERS],
      HEADERS.ProjectUsers,
      projectUserRows,
    );
    batchWriteObjects_(
      sheetMap[SHEETS.PROJECT_ORG],
      HEADERS.ProjectOrgChart,
      orgSeeds,
    );
    batchWriteObjects_(
      sheetMap[SHEETS.FORM_CATALOG],
      HEADERS.FormCatalog,
      FORM_DEFS,
    );
    batchWriteObjects_(sheetMap[SHEETS.DAILY_LOG], HEADERS.DailyLog, logs);
    batchWriteObjects_(sheetMap[SHEETS.PPE], HEADERS.PpeIssue, ppeSeeds);
    batchWriteObjects_(sheetMap[SHEETS.SCAFFOLD], HEADERS.ScaffoldTracker, scf);
    batchWriteObjects_(sheetMap[SHEETS.AUDITS], HEADERS.Audits, auditSeeds);
    batchWriteObjects_(
      sheetMap[SHEETS.AUDIT_SCORES],
      HEADERS.AuditScores,
      auditScoreSeeds,
    );
    batchWriteObjects_(
      sheetMap[SHEETS.OBSERVATIONS],
      HEADERS.Observations,
      obsSeeds,
    );
    batchWriteObjects_(
      sheetMap[SHEETS.NOTIFICATIONS],
      HEADERS.Notifications,
      notifSeeds,
    );
    batchWriteObjects_(
      sheetMap[SHEETS.TRAINING],
      HEADERS.TrainingCalendar,
      trainingSeeds,
    );
    batchWriteObjects_(
      sheetMap[SHEETS.GALLERY],
      HEADERS.Gallery,
      ensureDemoGallery_(projects),
    );

    // 6. FORMAT ALL CANONICAL SHEETS (Highlighted Teal Headers, Plain Body, Creative Special Cases)
    Object.keys(HEADERS).forEach((name) => {
      formatSheetProfessionally_(sheetMap[name], name, HEADERS[name]);
    });

    // 7. REMOVE SCRATCH SHEET (Spreadsheet is now 100% pristine and canonical)
    const finalTemp = ss.getSheetByName("__TEMP_SCRATCH__");
    if (finalTemp && ss.getSheets().length > 1) {
      try {
        ss.deleteSheet(finalTemp);
      } catch (e) {
        Logger.log("Could not remove temp sheet: " + e);
      }
    }

    // 8. Ensure Drive Project Folders
    projects.forEach((p) => ensureProjectFolder_(p));

    // 9. Ensure Escalation Trigger
    if (
      !ScriptApp.getProjectTriggers().some(
        (t) => t.getHandlerFunction() === "checkEscalations",
      )
    ) {
      ScriptApp.newTrigger("checkEscalations")
        .timeBased()
        .everyHours(1)
        .create();
    }

    ss.getRange("A1");
    const result = {
      ok: true,
      spreadsheetUrl: ss.getUrl(),
      spreadsheetId: ssId,
      folderUrl: root.getUrl(),
      sheetCount: ss.getSheets().length,
      userCount: rowsToObjects_(SHEETS.USERS).length,
      projectCount: rowsToObjects_(SHEETS.PROJECTS).length,
      formCount: rowsToObjects_(SHEETS.FORM_CATALOG).length,
      triggerCount: ScriptApp.getProjectTriggers().filter(
        (t) => t.getHandlerFunction() === "checkEscalations",
      ).length,
      demoLogins: userRows.map((u) => ({
        employeeId: u.employeeId,
        uan: u.uan,
        role: u.role,
        name: u.name,
      })),
    };
    Logger.log(JSON.stringify(result));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function setupStatus() {
  const props = PropertiesService.getScriptProperties();
  const ssId =
    String(DATABASE_SPREADSHEET_ID || "").trim() ||
    props.getProperty("SPREADSHEET_ID");
  if (ssId) props.setProperty("SPREADSHEET_ID", ssId);
  const rootId = props.getProperty("ROOT_FOLDER_ID");
  const status = {
    spreadsheetId: ssId || "",
    spreadsheetUrl: "",
    rootFolderId: rootId || "",
    rootFolderUrl: "",
    sheets: [],
    users: 0,
    projects: 0,
    forms: 0,
    escalationTriggers: 0,
  };
  try {
    if (ssId) {
      const ss = SpreadsheetApp.openById(ssId);
      status.spreadsheetUrl = ss.getUrl();
      status.sheets = ss.getSheets().map((s) => s.getName());
      status.users = rowsToObjects_(SHEETS.USERS).length;
      status.projects = rowsToObjects_(SHEETS.PROJECTS).length;
      status.forms = rowsToObjects_(SHEETS.FORM_CATALOG).length;
    }
    if (rootId) status.rootFolderUrl = DriveApp.getFolderById(rootId).getUrl();
    status.escalationTriggers = ScriptApp.getProjectTriggers().filter(
      (t) => t.getHandlerFunction() === "checkEscalations",
    ).length;
    status.ok = Boolean(
      status.spreadsheetId &&
      status.rootFolderId &&
      status.users &&
      status.projects &&
      status.forms,
    );
    status.message = status.ok
      ? "System is initialized."
      : "System is incomplete. Run initializeSystem() again.";
  } catch (e) {
    status.ok = false;
    status.message = String(e && e.message ? e.message : e);
  }
  Logger.log(JSON.stringify(status));
  return status;
}

function ensureSheetSchema_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  cleanAndAlignSheet_(sh, name, headers);
  return sh;
}

function ensureProjectFolder_(project) {
  const root = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("ROOT_FOLDER_ID"),
  );
  const name = project.code + " - " + project.name;
  const folders = root.getFoldersByName(name);
  const folder = folders.hasNext() ? folders.next() : root.createFolder(name);
  MODULES.forEach((m) => {
    const kids = folder.getFoldersByName(m.title);
    if (!kids.hasNext()) folder.createFolder(m.title);
  });
  ["Gallery", "eLibrary", "Generated PDFs"].forEach((n) => {
    const kids = folder.getFoldersByName(n);
    if (!kids.hasNext()) folder.createFolder(n);
  });
  return folder.getId();
}

function getProjectFolder_(project) {
  const root = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("ROOT_FOLDER_ID"),
  );
  const name = project.code + " - " + project.name;
  const folders = root.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.getFolderById(ensureProjectFolder_(project));
}

function getNamedSubfolder_(project, name) {
  const pf = getProjectFolder_(project);
  const it = pf.getFoldersByName(name);
  return it.hasNext() ? it.next() : pf.createFolder(name);
}
