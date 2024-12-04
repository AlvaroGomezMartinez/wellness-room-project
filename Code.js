/**
 * Triggered on form submission. Calculates the duration between "Checking In" and "Checking Out"
 * actions for the same ID on the same day and records the duration in Column J (10th column).
 * Sends an automated email to the counselor (Dr. Anderson) if specific responses (they indicated
 * that they needed additional support) are detected in columns G, H, or I.
 * 
 * @author Alvaro Gomez, Academic Technology Coach
 * Office: 210-397-9408
 *   Cell: 210-363-1577
 * 
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e - The event object for the form submission trigger.
 */
function onFormSubmit(e) {
  const sheet = e.source.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  // Get data from the new form submission (last row)
  const newRowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const newTimestamp = new Date(newRowData[0]); // Column A (Timestamp of the form submission)
  const idNumber = newRowData[1]; // Column B (ID Number)
  const action = newRowData[2]; // Column C (Action: "Checking In" or "Checking Out")

  // Only process if the new action is "Checking Out"
  if (action !== "Checking Out") return;

  let durationCalculated = false; // Flag to ensure only one duration is calculated

  // Get data for all previous rows to search for a matching "Checking In"
  const data = sheet.getRange(1, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Iterate through previous rows starting from the most recent
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const timestamp = new Date(row[0]); // Column A (Timestamp)
    const previousId = row[1]; // Column B (ID Number)
    const previousAction = row[2]; // Column C (Action)

    // Check for matching ID, "Checking In" action, and same date
    if (
      previousId === idNumber &&
      previousAction === "Checking In" &&
      timestamp.toDateString() === newTimestamp.toDateString()
    ) {
      // Calculate the time duration in milliseconds
      const durationMs = newTimestamp - timestamp;

      // Convert duration to hours and minutes
      const hours = Math.floor(durationMs / (1000 * 60 * 60));
      const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));

      // Write the calculated duration in Column J (10th column) of the "Checking Out" row
      sheet.getRange(lastRow, 10).setValue(`${hours}h ${minutes}m`);

      durationCalculated = true;
      break; // Exit the loop after finding the first valid match
    }
  }

  checkAndSendEmail(newRowData, idNumber);
}

/**
 * Checks specific columns (G, H, I) in the new row data for predefined values
 * and sends an email if any of the conditions are met.
 * 
 * @param {Array} rowData - The row data from the form submission.
 * @param {string|number} idNumber - The ID number of the student from the form submission.
 */
function checkAndSendEmail(rowData, idNumber) {
  const valueG = rowData[6];
  const valueH = rowData[7];
  const valueI = rowData[8];

  if (
    valueG === "5. I'm struggling, I need outside support." ||
    valueH === "5. I still need support." ||
    valueI === "Yes"
  ) {

    const recipient = "denecia-1.anderson@nisd.net"
    // const recipient = "alvaro.gomez@nisd.net"; // For testing
    const subject = "🚩 Wellness Room Follow-up";
    const message = `Dr. Anderson,<br>
    
This is an automated message regarding the Wellness Room check-out form submitted by the student with ID: <b>${idNumber}</b>, who indicated that they need additional support in one of the questions.<br><br>

<table border="1" style="border-collapse: collapse; width: 50%;">
  <tr>
    <th>Question</th>
    <th>Student Response</th>
  </tr>
  <tr>
    <td>On a scale of 1-5, how were you feeling coming into the Wellness Room?</td>
    <td>${rowData[6]}</td>
  </tr>
  <tr>
    <td>One a scale of 1-5, how do you feel leaving the Wellness Room?</td>
    <td>${rowData[7]}</td>
  </tr>
  <tr>
    <td>I would like outside support. (Meeting with the counselor, or supporting resources)</td>
    <td>${rowData[8]}</td>
  </tr>
</table><br><br>


<a href="https://docs.google.com/spreadsheets/d/1ILPCsLeJccz4_YN21znSFZRTyy6vbYPkvtVvMSB903k/edit?resourcekey=&gid=350830063#gid=350830063">Wellness Room Responses Sheet</a>`;

    // Use MailApp to send the email
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: message,
    });
  }
}
