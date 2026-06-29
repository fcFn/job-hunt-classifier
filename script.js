function filterJobRejectionEmails() {
  // ==========================================
  // ⚙️ CONFIGURATION ZONE
  // ==========================================
  // 1. Paste your OpenAI API Key here
  const OPENAI_API_KEY = "sk-YOUR-API-KEY-HERE";
  // 2. The date to start looking from (YYYY/MM/DD)
  const START_DATE = "2026/06/20";
  // 3. Where emails go if the AI thinks they are NOT
  //    rejections/next steps (Human review)
  const RETAINED_FOLDER_NAME = "Job Hunt - Review";
  // 4. Where emails go if they are scheduling requests or
  //    next steps
  const NEXT_STEP_FOLDER_NAME = "Job Hunt - Next Steps";
  // 5. Words that trigger the filter to look at an email
  const JOB_KEYWORDS = [
    "application",
    "role",
    "candidate",
    "interview",
    "position",
    "job", // had one that had none of the above so added this >:-(
  ];
  // 5. Excluded senders, e.g. newsletters, job boards; you might want
  //    to add more discriminators if you find unnecessary
  //    classifications.
  const EXCLUDED_SENDERS = [
    "me",
    "LinkedIn Job Alerts",
    "LinkedIn Job Recommendations",
  ];
  // 6. Status labels used for classification and reporting
  const STATUS_UNKNOWN = "Unknown";
  const STATUS_REJECTED = "Rejected";
  const STATUS_NEXT_STEP = "Next Step / Invitation";
  const STATUS_NEEDS_REVIEW = "Needs Human Review";

  // ==========================================
  // 🔍 STAGE 1: GMAIL PRE-FILTERING
  // ==========================================
  const keywordQuery = `(${JOB_KEYWORDS.join(" OR ")})`;
  const excludeQuery = EXCLUDED_SENDERS
    .map((sub) => `-from:"${sub}"`)
    .join(" ");

  // Exclude both target labels to avoid processing the same email twice
  const searchQuery =
    `${keywordQuery} in:inbox is:unread ` +
    `-label:"${RETAINED_FOLDER_NAME}" ` +
    `-label:"${NEXT_STEP_FOLDER_NAME}" ` +
    `${excludeQuery} after:${START_DATE}`;

  const threads = GmailApp.search(searchQuery);
  if (threads.length === 0) {
    Logger.log("No new emails matching the job search criteria.");
    return;
  }

  // Set up Review Label
  let reviewFolder = GmailApp.getUserLabelByName(RETAINED_FOLDER_NAME);
  if (!reviewFolder) {
    reviewFolder = GmailApp.createLabel(RETAINED_FOLDER_NAME);
  }

  // Set up Next Steps Label
  let nextStepFolder = GmailApp.getUserLabelByName(
    NEXT_STEP_FOLDER_NAME
  );
  if (!nextStepFolder) {
    nextStepFolder = GmailApp.createLabel(NEXT_STEP_FOLDER_NAME);
  }

  // Statistics & Logs
  let emailsSentToLLM = 0;
  let emailsRetained = 0;
  let nextStepsFound = 0;
  let errors = 0;
  const processedEmailsLog = [];

  // ==========================================
  // 🤖 STAGE 2: LLM CLASSIFICATION
  // ==========================================
  for (let i = 0; i < threads.length; i++) {
    const latestMessage = threads[i].getMessages().slice(-1)[0];
    const emailBody = latestMessage
      .getPlainBody()
      .substring(0, 2000);

    const payload = {
      model: "gpt-4o-mini",
      messages: [
        {
          role: "system",
          content:
            "You are a precise email classifier for a job hunter. " +
            "Read the provided email text and categorize it into " +
            "exactly one of three tags:\n\n" +
            "1. 'REJECTION' - Use this ONLY if it is a definitive, " +
            "unambiguous rejection notice (e.g., 'we decided to " +
            "move forward with other candidates', 'not selecting " +
            "you at this time').\n" +
            "2. 'NEXT_STEP' - Use this ONLY if it is an invitation " +
            "to schedule an interview, a coding challenge, a " +
            "phone screen, a job offer, or a request to provide " +
            "availability for the next step.\n" +
            "3. 'UNKNOWN' - Use this if the email is generic, a " +
            "confirmation of application submission, requires " +
            "updates, or if you are at all uncertain or find it " +
            "ambiguous. Lean heavily on involving a human in " +
            "the loop.\n\n" +
            "Your output must be exactly one of these words: " +
            "REJECTION, NEXT_STEP, or UNKNOWN. Do not include " +
            "any other text or punctuation."
        },
        {
          role: "user",
          content: emailBody,
        },
      ],
      temperature: 0,
    };

    const options = {
      method: "post",
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    try {
      const response = UrlFetchApp.fetch(
        "https://api.openai.com/v1/chat/completions",
        options
      );
      emailsSentToLLM++;

      const json = JSON.parse(response.getContentText());
      const aiDecision = json.choices[0].message.content
        .trim()
        .toUpperCase();

      const emailMetadata = {
        subject: latestMessage.getSubject(),
        from: latestMessage.getFrom(),
        link: threads[i].getPermalink(),
        status: STATUS_UNKNOWN
      };

      if (aiDecision === "REJECTION") {
        threads[i].markRead();
        threads[i].moveToArchive();
        emailMetadata.status = STATUS_REJECTED;
        processedEmailsLog.push(emailMetadata);
      } else if (aiDecision === "NEXT_STEP") {
        threads[i].addLabel(nextStepFolder);
        threads[i].moveToArchive();
        nextStepsFound++;
        emailMetadata.status = STATUS_NEXT_STEP;
        processedEmailsLog.push(emailMetadata);
      } else {
        // Fallback / UNKNOWN -> Needs manual human review
        threads[i].addLabel(reviewFolder);
        threads[i].moveToArchive();
        emailsRetained++;
        emailMetadata.status = STATUS_NEEDS_REVIEW;
        processedEmailsLog.push(emailMetadata);
      }
    } catch (e) {
      Logger.log(`API Error: ${e.toString()}`);
      errors++;
    }
    // An extremely generous 1-second pause to prevent hitting API
    // rate limits
    Utilities.sleep(1000);
  }

  // ==========================================
  // 📊 STAGE 3: REPORTING & LOGGING (HTML UTF-8)
  // ==========================================
  if (emailsSentToLLM > 0) {
    const userEmailAddress = Session.getActiveUser().getEmail();
    const date = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    // Separate rejections for the CSV file
    const rejectedEmails = processedEmailsLog.filter(
      (e) => e.status === STATUS_REJECTED
    );

    // Filter out actionable emails to build our inline dashboard
    // table
    const actionableEmails = processedEmailsLog.filter(
      (e) => e.status === STATUS_NEXT_STEP
    );

    // 1. Build the Rejection-Only CSV
    let csvContent = "Date,From,Subject,Link,AI_Classification\n";
    rejectedEmails.forEach((email) => {
      const safeFrom = email.from.replace(/"/g, '""');
      const safeSubject = email.subject.replace(/"/g, '""');
      csvContent +=
        `${date},"${safeFrom}","${safeSubject}",` +
        `"${email.link}",${email.status}\n`;
    });
    const csvBlob =
      rejectedEmails.length > 0
        ? Utilities.newBlob(
            csvContent,
            "text/csv",
            `rejections_${date}.csv`
          )
        : null;

    const emailSubject = "Automated Job Filter: Execution Summary";

    // 2. Build the HTML Document
    let htmlBody =
      `<!DOCTYPE html>` +
      `<html>` +
      `<head>` +
      `  <meta http-equiv="Content-Type" ` +
      `content="text/html; charset=UTF-8">` +
      `  <style>` +
      `    body { ` +
      `      font-family: -apple-system, BlinkMacSystemFont, ` +
      `"Segoe UI", Roboto, sans-serif; ` +
      `      line-height: 1.5; ` +
      `      color: #333; ` +
      `    }` +
      `    ul { padding-left: 20px; margin-bottom: 25px; }` +
      `    h3 { margin-top: 25px; color: #1a1a1a; }` +
      `    .error { color: #dc3545; font-weight: bold; }` +
      `    table { ` +
      `      width: 100%; ` +
      `      border-collapse: collapse; ` +
      `      margin-top: 15px; ` +
      `      box-shadow: 0 1px 3px rgba(0,0,0,0.1); ` +
      `    }` +
      `    th, td { ` +
      `      text-align: left; ` +
      `      padding: 10px 12px; ` +
      `      border: 1px solid #e0e0e0; ` +
      `    }` +
      `    th { ` +
      `      background-color: #f7f9fa; ` +
      `      font-weight: 600; ` +
      `      color: #4a5568; ` +
      `    }` +
      `    tr:nth-child(even) { background-color: #fcfcfd; }` +
      `    .badge { ` +
      `      padding: 4px 8px; ` +
      `      border-radius: 4px; ` +
      `      font-size: 12px; ` +
      `      font-weight: bold; ` +
      `      text-transform: uppercase; ` +
      `    }` +
      `    .badge-next { ` +
      `      background-color: #d4edda; ` +
      `      color: #155724; ` +
      `    }` +
      `    .badge-review { ` +
      `      background-color: #fff3cd; ` +
      `      color: #856404; ` +
      `    }` +
      `    .link-btn { ` +
      `      text-decoration: none; ` +
      `      color: #007bff; ` +
      `      font-weight: bold; ` +
      `    }` +
      `  </style>` +
      `</head>` +
      `<body>` +
      `  <p>The automated job filter has finished its run.</p>` +
      `  <ul>` +
      `    <li>Total scanned by AI: ` +
      `<strong>${emailsSentToLLM}</strong></li>` +
      `    <li>&#9989; Moved to '${NEXT_STEP_FOLDER_NAME}': ` +
      `<strong>${nextStepsFound}</strong></li>` +
      `    <li>&#128269; Moved to '${RETAINED_FOLDER_NAME}' ` +
      `(Human Review): <strong>${emailsRetained}</strong></li>` +
      `    <li>&#128371; Archived as Rejections: ` +
      `<strong>${rejectedEmails.length}</strong></li>` +
      `  </ul>`;

    // 3. Dynamically Generate Table if Actionable Items Exist
    if (actionableEmails.length > 0) {
      htmlBody += `<h3>&#128197; Action Items & Focus Areas</h3>`;
      htmlBody += `<table>`;
      htmlBody += `  <thead>`;
      htmlBody += `    <tr>`;
      htmlBody += `      <th>Category</th>`;
      htmlBody += `      <th>From</th>`;
      htmlBody += `      <th>Subject</th>`;
      htmlBody += `      <th>Action</th>`;
      htmlBody += `    </tr>`;
      htmlBody += `  </thead>`;
      htmlBody += `  <tbody>`;

      actionableEmails.forEach((email) => {
        const isNextStep = email.status === STATUS_NEXT_STEP;
        const badgeClass = isNextStep ? "badge-next" : "badge-review";
        const statusLabel = isNextStep ? "Next Step" : "Review";

        const cleanFromCell = email.from
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        const cleanSubjectCell = email.subject
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");

        htmlBody += `    <tr>`;
        htmlBody +=
          `      <td><span class="badge ${badgeClass}">` +
          `${statusLabel}</span></td>`;
        htmlBody += `      <td>${cleanFromCell}</td>`;
        htmlBody += `      <td>${cleanSubjectCell}</td>`;
        htmlBody +=
          `      <td><a class="link-btn" ` +
          `href="${email.link}" target="_blank">` +
          `Open &#8594;</a></td>`;
        htmlBody += `    </tr>`;
      });

      htmlBody += `  </tbody>`;
      htmlBody += `</table>`;
    } else {
      htmlBody +=
        `<p><em>No positive leads or review items ` +
        `detected in this run.</em></p>`;
    }

    // 4. Handle Footer & Attachments
    if (csvBlob) {
      htmlBody +=
        `<p style="margin-top:25px; font-size: 13px; color: #666;">` +
        `Attached is the CSV log of archived rejections. ` +
        `Please review it periodically ` +
        `to ensure no false positives slipped through.</p>`;
    } else {
      htmlBody += `No rejections found this round!`;
    }
    if (errors > 0) {
      htmlBody +=
        `<p class="error">\n&#9888; There were ${errors} ` +
        `errors when requesting the OpenAI API. ` +
        `Please review the execution log.</p>`;
    }

    htmlBody += `</body></html>`;

    // 5. Final Dispatch
    const options = {
      htmlBody: htmlBody
    };

    if (csvBlob) {
      options.attachments = [csvBlob];
    }

    GmailApp.sendEmail(userEmailAddress, emailSubject, "", options);
  }
}
