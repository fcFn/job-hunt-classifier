function filterJobRejectionEmails() {
  // ==========================================
  // ⚙️ CONFIGURATION ZONE
  // ==========================================
  // 1. Paste your OpenAI API Key here
  const OPENAI_API_KEY = "sk-YOUR-API-KEY-HERE";
  // 2. The date to start looking from (YYYY/MM/DD)
  const START_DATE = "2026/06/20";
  // 3. Where emails go if the AI thinks they are NOT rejections
  const RETAINED_FOLDER_NAME = "Job Hunt - Review";
  // 4. Words that trigger the filter to look at an email
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
  // ==========================================
  // 🔍 STAGE 1: GMAIL PRE-FILTERING
  // ==========================================
  const keywordQuery = `(${JOB_KEYWORDS.join(" OR ")})`;
  const excludeQuery = EXCLUDED_SENDERS.map((sub) => `-from:"${sub}"`).join(
    " ",
  );
  // the inbox and unread labels make sure we don't process the same
  // email twice
  const searchQuery =
    `${keywordQuery} in:inbox is:unread ` +
    `-label:"${RETAINED_FOLDER_NAME}" ${excludeQuery} after:${START_DATE}`;
  const threads = GmailApp.search(searchQuery);
  if (threads.length === 0) {
    Logger.log("No new emails matching the job search criteria.");
    return;
  }
  let targetFolder = GmailApp.getUserLabelByName(RETAINED_FOLDER_NAME);
  if (!targetFolder) {
    targetFolder = GmailApp.createLabel(RETAINED_FOLDER_NAME);
  }
  // A few very basic statistics
  let emailsSentToLLM = 0;
  let emailsRetained = 0;
  let errors = 0;
  const rejectedEmails = [];

  // ==========================================
  // 🤖 STAGE 2: LLM CLASSIFICATION
  // ==========================================
  for (let i = 0; i < threads.length; i++) {
    const latestMessage = threads[i].getMessages().slice(-1)[0];
    const emailBody = latestMessage.getPlainBody().substring(0, 2000);
    const payload = {
      model: "gpt-4o-mini",
      messages: [
        {
          role: "system",
          content:
            "You are a precise email classifier. " +
            "Read the provided email text " +
            "and determine if it is a job rejection notice. " +
            "Respond with exactly 'YES' if it is a definitive rejection, " +
            "and 'NO' if it is not. " +
            "If the email is an interview invitation, a job offer, " +
            "a request for more information, or if you are unsure, " +
            "you must output 'NO'.",
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
        options,
      );
      emailsSentToLLM++;
      const json = JSON.parse(response.getContentText());
      const aiDecision = json.choices[0].message.content.trim().toUpperCase();
      if (aiDecision === "YES") {
        threads[i].markRead();
        threads[i].moveToArchive();
        rejectedEmails.push({
          subject: latestMessage.getSubject(),
          from: latestMessage.getFrom(),
          link: threads[i].getPermalink(),
        });
      } else {
        threads[i].addLabel(targetFolder);
        threads[i].moveToArchive();
        emailsRetained++;
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
  // 📊 STAGE 3: REPORTING & LOGGING
  // ==========================================
  // To reduce noise, we only send the email if we had any qualified
  // emails to process. Adjust it if you want, for example, only if
  // you have anything to review.
  if (emailsSentToLLM > 0) {
    const userEmailAddress = Session.getActiveUser().getEmail();
    let csvContent = "Date,From,Subject,Link,Status\n";
    const date = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    rejectedEmails.forEach((email) => {
      const safeFrom = email.from.replace(/"/g, '""');
      const safeSubject = email.subject.replace(/"/g, '""');
      csvContent +=
        `${date},"${safeFrom}","${safeSubject}"` +
        `,"${email.link}",Rejected\n`;
    });
    const csvBlob =
      rejectedEmails.length > 0
        ? Utilities.newBlob(csvContent, "text/csv", `rejections_${date}.csv`)
        : null;
    const emailSubject = "Automated Job Filter: Execution Summary";
    let emailBody =
      `The automated job filter has finished its run.\n\n` +
      `Total scanned by AI: ${emailsSentToLLM}\n` +
      `Moved to '${RETAINED_FOLDER_NAME}': ${emailsRetained}\n` +
      `Banished to Purgatory: ${rejectedEmails.length}\n\n`;
    const options = {};
    if (csvBlob) {
      options.attachments = [csvBlob];
      emailBody +=
        `Attached is the CSV log of archived rejections. ` +
        `Please review it periodically ` +
        `to ensure no false positives slipped through.`;
    } else {
      emailBody += `No rejections found this round!`;
    }
    if (errors > 0) {
      emailBody +=
        `\n⚠️ There were ${errors} errors when requesting the OpenAI API. ` +
        `Please review the execution log.`;
    }
    GmailApp.sendEmail(userEmailAddress, emailSubject, emailBody, options);
  }
}
