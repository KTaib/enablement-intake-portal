// ============================================================
// ACQUIA SALES ENABLEMENT INTAKE — PHASE 1 FINAL
// Mapped to YOUR exact form (taeeb.khan@acquia.com form)
// Column mapping verified from form screenshots June 2026
// ============================================================

// ─── CONFIG — fill these in before deploying ────────────────
const CONFIG = {
  SLACK_WEBHOOK:   "https://hooks.slack.com/services/YOUR/WEBHOOK/URL",
  SLACK_BOT_TOKEN: "xoxb-YOUR-BOT-TOKEN",
  SLACK_CHANNEL:   "C0ANRQXDBDY",                   // #enablement-intake
  ASANA_TOKEN:     "YOUR_ASANA_PERSONAL_ACCESS_TOKEN",
  ASANA_PROJECT:   "YOUR_INTAKE_TRACKER_PROJECT_GID",
  TEAM_EMAIL:      "taeeb.khan@acquia.com",
  FORM_BASE_URL:   "https://docs.google.com/forms/d/YOUR_FORM_ID/viewform",

  ASANA_SECTIONS: {
    "Unscheduled": "YOUR_UNSCHEDULED_SECTION_GID",
    "Scheduled":   "YOUR_SCHEDULED_SECTION_GID",
    "In Review":   "YOUR_IN_REVIEW_SECTION_GID",
    "In Progress": "YOUR_IN_PROGRESS_SECTION_GID",
    "Done":        "YOUR_DONE_SECTION_GID"
  },

  // Slack user IDs for direct DM outreach per team
  // Right-click user in Slack > View profile > More (...) > Copy member ID
  TEAM_SLACK_IDS: {
    "Sales":     ["U01XXXXXXX"],   // replace with real IDs
    "CS":        ["U01AAAAAAA"],
    "Marketing": ["U01CCCCCCC"]
  }
};

// ─── COLUMN INDEX MAP (0-based, A=0) ────────────────────────
// Verified against your form screenshot — do not change unless
// you reorder questions in the form.
const COL = {
  TIMESTAMP:   0,   // A — auto
  EMAIL:       1,   // B — Google account email
  TYPE:        2,   // C — Recorded / Live / Task branch selector

  // Recorded branch
  REC_TITLE:   3,   // D
  REC_CRITERIA:4,   // E — productivity criteria (checkboxes, comma-sep)
  REC_FOCUS:   5,   // F — focus area (checkboxes)
  REC_WHY:     6,   // G — why important
  REC_SUMMARY: 7,   // H — training summary / event description
  REC_NEWREF:  8,   // I — new or refresher
  REC_AUDIENCE:9,   // J — who enrolled (checkboxes)
  REC_URGENCY: 10,  // K — how soon
  REC_PEOPLE:  11,  // L — individuals in recording/deck

  // Live branch
  LIV_TITLE:   12,  // M
  LIV_CRITERIA:13,  // N
  LIV_FOCUS:   14,  // O
  LIV_SUMMARY: 15,  // P
  LIV_NEWREF:  16,  // Q
  LIV_AUDIENCE:17,  // R — who to invite
  LIV_URGENCY: 18,  // S
  LIV_PRESENTERS:19,// T — presenters / deck creators
  LIV_OTHERS:  20,  // U — who else must be present

  // Task branch
  TSK_AUDIENCE:21,  // V — teams to reach (checkboxes)
  TSK_DESC:    22,  // W — describe the request
  TSK_PREWORK: 23,  // X — pre-work needed
  TSK_DATES:   24,  // Y — important dates / key takeaways
  TSK_DOCS:    25   // Z — supporting documents link
};

// ─── URGENCY NORMALISER ──────────────────────────────────────
// Your form uses long labels — map them to short canonical values
function normaliseUrgency(raw) {
  const r = (raw || "").toLowerCase();
  if (r.includes("yesterday"))  return "Yesterday";
  if (r.includes("right now"))  return "Right now";
  if (r.includes("0 - 1") || r.includes("0-1")) return "0-1 month";
  if (r.includes("1 month"))    return "1 month+";
  return raw || "Unknown";
}

// ─── TYPE NORMALISER ─────────────────────────────────────────
function normaliseType(raw) {
  const r = (raw || "").toLowerCase();
  if (r.includes("recorded")) return "Recorded";
  if (r.includes("live"))     return "Live";
  if (r.includes("task"))     return "Task";
  return "Task";
}

// ─── MAIN TRIGGER ────────────────────────────────────────────
function onFormSubmit(e) {
  try {
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row    = sheet.getLastRow();
    const data   = sheet.getRange(row, 1, 1, 26).getValues()[0];

    const intake = parseSubmission(data, row);

    // Write Intake ID back into a spare column (column 27 = AA)
    // If you want it in a specific column, change the 27 below
    sheet.getRange(row, 27).setValue(intake.id);

    createAsanaTask(intake);
    postSlackNotification(intake);
    sendConfirmationEmail(intake);

    Logger.log("Processed: " + intake.id + " | " + intake.type + " | " + intake.title);
  } catch (err) {
    Logger.log("ERROR: " + err.message + "\n" + err.stack);
    MailApp.sendEmail(CONFIG.TEAM_EMAIL, "Intake script error — " + new Date(), err.message);
  }
}

// ─── PARSE SUBMISSION ────────────────────────────────────────
function parseSubmission(data, row) {
  const timestamp = data[COL.TIMESTAMP];
  const email     = data[COL.EMAIL] || "";
  const typeRaw   = data[COL.TYPE]  || "";
  const type      = normaliseType(typeRaw);

  let title, audience, urgency, summary, people, extra;

  if (type === "Recorded") {
    title    = data[COL.REC_TITLE]    || "";
    audience = data[COL.REC_AUDIENCE] || "";
    urgency  = normaliseUrgency(data[COL.REC_URGENCY]);
    summary  = data[COL.REC_SUMMARY]  || "";
    people   = data[COL.REC_PEOPLE]   || "";
    extra = {
      criteria: data[COL.REC_CRITERIA] || "",
      focus:    data[COL.REC_FOCUS]    || "",
      why:      data[COL.REC_WHY]      || "",
      newRef:   data[COL.REC_NEWREF]   || ""
    };

  } else if (type === "Live") {
    title    = data[COL.LIV_TITLE]      || "";
    audience = data[COL.LIV_AUDIENCE]   || "";
    urgency  = normaliseUrgency(data[COL.LIV_URGENCY]);
    summary  = data[COL.LIV_SUMMARY]    || "";
    people   = data[COL.LIV_PRESENTERS] || "";
    extra = {
      criteria: data[COL.LIV_CRITERIA] || "",
      focus:    data[COL.LIV_FOCUS]    || "",
      newRef:   data[COL.LIV_NEWREF]   || "",
      others:   data[COL.LIV_OTHERS]   || ""
    };

  } else {
    // Task
    title    = "Task Request";
    audience = data[COL.TSK_AUDIENCE] || "";
    urgency  = "0-1 month";
    summary  = data[COL.TSK_DESC]     || "";
    people   = "";
    extra = {
      prework: data[COL.TSK_PREWORK] || "",
      dates:   data[COL.TSK_DATES]   || "",
      docs:    data[COL.TSK_DOCS]    || ""
    };
  }

  // Generate Intake ID: SE-YYYYMMDD-NNN
  const d    = new Date(timestamp);
  const pad  = n => String(n).padStart(2, "0");
  const date = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
  const seq  = String(row - 1).padStart(3, "0");
  const id   = `SE-${date}-${seq}`;

  const team = detectTeam(email, audience);

  return { id, type, title, email, audience, urgency, summary, people, extra, team, timestamp };
}

function detectTeam(email, audience) {
  const s = (email + " " + audience).toLowerCase();
  if (s.includes("sales") || s.includes("ae") || s.includes("bdr")) return "Sales";
  if (s.includes("csm") || s.includes("customer success") || s.includes("tam")) return "CS";
  if (s.includes("marketing")) return "Marketing";
  if (s.includes("presales")) return "Presales";
  return "General";
}

// ─── ASANA TASK ───────────────────────────────────────────────
function createAsanaTask(intake) {
  const urgencyDays = { "Yesterday": 1, "Right now": 2, "0-1 month": 21, "1 month+": 45 };
  const due = new Date();
  due.setDate(due.getDate() + (urgencyDays[intake.urgency] || 30));

  const notes = buildAsanaNotes(intake);

  const payload = {
    data: {
      name:     "[" + intake.id + "] " + intake.title,
      notes:    notes,
      projects: [CONFIG.ASANA_PROJECT],
      due_on:   due.toISOString().split("T")[0]
    }
  };

  const resp = UrlFetchApp.fetch("https://app.asana.com/api/1.0/tasks", {
    method:      "post",
    contentType: "application/json",
    headers:     { "Authorization": "Bearer " + CONFIG.ASANA_TOKEN },
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(resp.getContentText());

  if (result.data && result.data.gid) {
    const sectionGid = CONFIG.ASANA_SECTIONS["Unscheduled"];
    if (sectionGid) {
      UrlFetchApp.fetch("https://app.asana.com/api/1.0/sections/" + sectionGid + "/addTask", {
        method:      "post",
        contentType: "application/json",
        headers:     { "Authorization": "Bearer " + CONFIG.ASANA_TOKEN },
        payload:     JSON.stringify({ data: { task: result.data.gid } }),
        muteHttpExceptions: true
      });
    }
    Logger.log("Asana task created: " + result.data.gid);
  } else {
    Logger.log("Asana error: " + resp.getContentText());
  }
}

function buildAsanaNotes(intake) {
  const lines = [
    "Intake ID:  " + intake.id,
    "Type:       " + intake.type,
    "Requestor:  " + intake.email,
    "Team:       " + intake.team,
    "Audience:   " + intake.audience,
    "Urgency:    " + intake.urgency,
    ""
  ];

  if (intake.type === "Recorded") {
    lines.push("Criteria:   " + (intake.extra.criteria || "—"));
    lines.push("Focus area: " + (intake.extra.focus    || "—"));
    lines.push("New/Refresh:" + (intake.extra.newRef   || "—"));
    lines.push("Why needed: " + (intake.extra.why      || "—"));
    lines.push("Presenters: " + (intake.people         || "—"));
  } else if (intake.type === "Live") {
    lines.push("Criteria:   " + (intake.extra.criteria || "—"));
    lines.push("Focus area: " + (intake.extra.focus    || "—"));
    lines.push("New/Refresh:" + (intake.extra.newRef   || "—"));
    lines.push("Presenters: " + (intake.people         || "—"));
    lines.push("Others req: " + (intake.extra.others   || "—"));
  } else {
    lines.push("Pre-work:   " + (intake.extra.prework  || "—"));
    lines.push("Key dates:  " + (intake.extra.dates    || "—"));
    lines.push("Docs link:  " + (intake.extra.docs     || "—"));
  }

  lines.push("");
  lines.push("Summary:");
  lines.push(intake.summary);

  return lines.join("\n");
}

// ─── SLACK NOTIFICATION ──────────────────────────────────────
function postSlackNotification(intake) {
  const urgencyEmoji = {
    "Yesterday": ":rotating_light:",
    "Right now": ":red_circle:",
    "0-1 month": ":yellow_circle:",
    "1 month+":  ":white_circle:"
  };
  const emoji = urgencyEmoji[intake.urgency] || ":white_circle:";

  const typeLabel = {
    "Recorded": "Recorded session",
    "Live":     "Live session",
    "Task":     "Task request"
  }[intake.type] || intake.type;

  const text =
    emoji + " *New intake: " + intake.id + "*\n" +
    "*" + typeLabel + "* — " + intake.title + "\n" +
    "Requested by " + intake.email + " for *" + (intake.audience || "unspecified audience") + "*. " +
    "Urgency: " + intake.urgency + ".";

  UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK, {
    method:      "post",
    contentType: "application/json",
    payload:     JSON.stringify({
      channel:      CONFIG.SLACK_CHANNEL,
      text:         text,
      unfurl_links: false
    }),
    muteHttpExceptions: true
  });
}

// ─── CONFIRMATION EMAIL ──────────────────────────────────────
function sendConfirmationEmail(intake) {
  const typeLabels = {
    "Recorded": "Recorded Enablement Session",
    "Live":     "Live Enablement Session",
    "Task":     "Task Request"
  };

  const subject = "Your enablement request has been received — " + intake.id;

  const body =
    "Hi,\n\n" +
    "Thanks for submitting your Sales Enablement request. Here is a summary of what we received:\n\n" +
    "  Intake ID:  " + intake.id + "\n" +
    "  Type:       " + (typeLabels[intake.type] || intake.type) + "\n" +
    (intake.title !== "Task Request" ? "  Title:      " + intake.title + "\n" : "") +
    "  Audience:   " + (intake.audience || "—") + "\n" +
    "  Urgency:    " + intake.urgency + "\n\n" +
    "Our team will review this and follow up within 2 business days. " +
    "You can track status updates in the #enablement-intake Slack channel.\n\n" +
    "If anything needs clarifying, just reply to this email.\n\n" +
    "— Acquia Sales Enablement Team";

  MailApp.sendEmail({
    to:      intake.email,
    cc:      CONFIG.TEAM_EMAIL,
    subject: subject,
    body:    body
  });
}

// ─── TEAM OUTREACH — send DMs with team-prefilled form links ─
// Call manually or schedule: sends Slack DMs to each team with
// a pre-filled form link that sets their branch + defaults.
// IMPORTANT: Get your entry IDs first — see DEPLOYMENT.md Step 3.
function sendTeamOutreach(teamName) {
  const ids = CONFIG.TEAM_SLACK_IDS[teamName] || [];

  // Branch value as it appears exactly in your form radio button
  const branchValues = {
    "Sales":     "Recorded Enablement Session - Ideal for product announcements or courses with a knowledge check. Offers on-demand, self-paced learning.",
    "CS":        "Live Enablement Session - Best for sales skills, competitor updates, value selling, and live Q&A. Offers interactive, instructor-led learning.",
    "Marketing": "Task request - For specific enablement needs such as sales assets, process documentation, or resource updates. Offers support through tailored deliverables."
  };

  // Replace ENTRY_XXXXXXXXXX with your actual entry IDs from the form prefill URL
  // See DEPLOYMENT.md — Step 3: Get prefill entry IDs
  const entryIds = {
    TYPE: "entry.REPLACE_WITH_YOUR_TYPE_ENTRY_ID"
  };

  const branchVal = branchValues[teamName] || "";
  const formUrl   = CONFIG.FORM_BASE_URL + "?" +
                    entryIds.TYPE + "=" + encodeURIComponent(branchVal);

  const messages = {
    "Sales":
      "Hi — quick ask from the Acquia Enablement team. If your team has any upcoming product training, " +
      "competitive enablement, or onboarding needs, please submit a request here. " +
      "Pre-filled for Sales so it takes under 2 minutes: " + formUrl,
    "CS":
      "Hi — Enablement planning is open for the next cycle. If your team needs live sessions, " +
      "onboarding walkthroughs, or refresher content, submit a request here (pre-filled for CS): " + formUrl,
    "Marketing":
      "Hi — Enablement here. If Marketing needs content support, training assets, or co-created material " +
      "this quarter, drop a request here (pre-filled for Marketing): " + formUrl
  };

  const messageText = messages[teamName] ||
    "Please submit any enablement requests here: " + CONFIG.FORM_BASE_URL;

  ids.forEach(function(userId) {
    UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
      method:  "post",
      headers: {
        "Authorization": "Bearer " + CONFIG.SLACK_BOT_TOKEN,
        "Content-Type":  "application/json"
      },
      payload: JSON.stringify({
        channel:      userId,
        text:         messageText,
        unfurl_links: false
      }),
      muteHttpExceptions: true
    });
  });

  Logger.log("DMs sent to " + ids.length + " members of " + teamName);
}

function sendAllTeamOutreach() {
  ["Sales", "CS", "Marketing"].forEach(sendTeamOutreach);
}

// ─── TEST FUNCTION — run this first to verify everything ─────
// Creates a fake intake and runs the full pipeline without a real
// form submission. Check Logger + your email + Asana + Slack.
function testPipeline() {
  const fakeIntake = {
    id:        "SE-TEST-001",
    type:      "Recorded",
    title:     "Test: Q3 Product Enablement Module",
    email:     CONFIG.TEAM_EMAIL,
    audience:  "AEs, BDRs",
    urgency:   "0-1 month",
    summary:   "This is a test submission from the Apps Script test function.",
    people:    "Taeeb Khan",
    team:      "Sales",
    timestamp: new Date(),
    extra: {
      criteria: "Win Rate, Deal Value",
      focus:    "Product Knowledge",
      why:      "Reps need to understand the new feature set before Q3 launch.",
      newRef:   "New"
    }
  };

  Logger.log("--- TEST INTAKE ---");
  Logger.log(JSON.stringify(fakeIntake, null, 2));
  Logger.log("Asana notes preview:\n" + buildAsanaNotes(fakeIntake));

  // Uncomment these one at a time to test each integration:
  // createAsanaTask(fakeIntake);
  // postSlackNotification(fakeIntake);
  // sendConfirmationEmail(fakeIntake);
}
