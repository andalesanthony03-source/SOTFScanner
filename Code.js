const CFG = {

  EVENT: "Song of the Fireflies",

  DATE: "October 20, 2026",

  TIME: "8:00 PM",

  VENUE: "Samsung Performing Arts Theater",
WEB_APP_URL: "https://script.google.com/macros/s/AKfycby8OLZDDvJDM89swlMRTmCuc25IMQVtp1yu2U3M5kmjldDERMlptMq_XpE2t38SEw/exec"

};



// RUN ONCE to create the ticket sheets.

function setup() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();



  makeSheet_(ss, "Reservations", [

    "Reservation ID", "Name", "Email", "Quantity",

    "Payment", "Email Status", "Created"

  ]);



  makeSheet_(ss, "Tickets", [

    "Ticket ID", "Reservation ID", "Ticket No.",

    "Name", "Email", "Status", "Check-in Time"

  ]);

}



function makeSheet_(ss, name, headers) {

  let s = ss.getSheetByName(name);

  if (!s) s = ss.insertSheet(name);

  if (s.getLastRow() === 0) {

    s.appendRow(headers);

    s.setFrozenRows(1);

  }

}



// INSTALL a spreadsheet "On form submit" trigger.

function onFormSubmit(e) {

  const source = e.range.getSheet();

  const row = e.range.getRow();

  const headers = source.getRange(

    1, 1, 1, source.getLastColumn()

  ).getDisplayValues()[0].map(h => h.toLowerCase().trim());



  const answers = source.getRange(

    row, 1, 1, source.getLastColumn()

  ).getDisplayValues()[0];



  const find = fn => headers.findIndex(fn);

  const ni = find(h => h.includes("full name") || h === "name");

  const ei = find(h => h.includes("email"));

  const qi = find(h =>

    h.includes("number of tickets") ||

    h.includes("ticket quantity") ||

    h.includes("ticket count") ||

    h.includes("tickets")

  );



  if (ni < 0 || ei < 0 || qi < 0)

    throw Error("Check your Full Name, Email and Number of Tickets questions.");



  const name = answers[ni].trim();

  const email = answers[ei].trim();

  const qtyText = answers[qi].trim();

  const qty = Number(qtyText);



  if (!name || !/^[^\s@]+@[^\s@]+\\.[^\s@]+$/.test(email) ||

      !/^\d+$/.test(qtyText) || qty < 1 || qty > 10)

    throw Error("Invalid name, email or ticket quantity (1–10).");



  const lock = LockService.getScriptLock();

  lock.waitLock(10000);



  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const rs = ss.getSheetByName("Reservations");

    const ts = ss.getSheetByName("Tickets");



    const id = "SOTF-" +

      Utilities.getUuid().split("-")[0].toUpperCase();



    rs.appendRow([id, name, email, qty, "Pending", "Not Sent", new Date()]);





    const tickets = [];



    for (let i = 1; i <= qty; i++) {

      tickets.push([

        id + "-T" + String(i).padStart(2, "0"), // Ticket ID

        id,                                     // Reservation ID

        name,                                   // Guest Name

        email,                                  // Buyer Email

        "Unused",                               // Status

        ""                                      // Check-in Time

      ]);

    }



    ts.getRange(ts.getLastRow() + 1, 1, tickets.length, 6)

      .setValues(tickets);

  } finally {

    lock.releaseLock();

  }





  // No email is sent automatically.

}



// APPROVE payment: select a Reservations row, then run this.

function approvePayment() {

  const s = SpreadsheetApp.getActiveSheet();

  const r = s.getActiveRange().getRow();



  if (s.getName() !== "Reservations" || r < 2)

    throw Error("Select a reservation row in Reservations.");



  const ui = SpreadsheetApp.getUi();

  if (ui.alert("Confirm you verified the payment?",

      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;



  s.getRange(r, 5).setValue("Approved");

}



// MANUALLY SEND tickets: select the reservation row, then run this.



function sendTickets() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const s = ss.getActiveSheet();

  const r = s.getActiveRange().getRow();



  if (s.getName() !== "Reservations" || r < 2)

    throw Error("Select a reservation row in Reservations.");



  const v = s.getRange(r, 1, 1, 7).getValues()[0];

  const [id, name, email, qty, payment, status] = v;



  if (payment !== "Approved")

    throw Error("Approve payment before sending tickets.");



  if (status === "Sent")

    throw Error("Tickets were already sent.");



  if (!email || !Number.isInteger(Number(qty)) ||

      qty < 1 || qty > 10)

    throw Error("Invalid reservation details.");



 const base = CFG.WEB_APP_URL;



  // Replace with the Google Drive ID of your poster.

  const POSTER_FILE_ID = "1b1GX8joYpTt-OLAWTQl3YpBPcMD7HWIj";



  const poster = DriveApp.getFileById(POSTER_FILE_ID)

    .getBlob()

    .setName("song-of-the-fireflies-poster");



  const ts = ss.getSheetByName("Tickets");

  const all = ts.getDataRange().getValues();

  const tickets = all.slice(1).filter(x => x[1] === id);



  if (tickets.length !== Number(qty))

    throw Error("Ticket records do not match the reservation.");



  const images = { poster: poster };

  const cards = [];



  tickets.forEach((t, i) => {

    const tid = String(t[0]);

    const ticketName = String(t[2] || name);

    const url = base + "?ticket=" + encodeURIComponent(tid);



    const qrUrl =

      "https://api.qrserver.com/v1/create-qr-code/" +

      "?size=400x400&margin=12&data=" +

      encodeURIComponent(url);



    const res = UrlFetchApp.fetch(qrUrl);



    if (res.getResponseCode() !== 200)

      throw Error("QR generation failed for " + tid);



    const key = "qr" + i;

    images[key] = res.getBlob().setName(tid + ".png");





cards.push(

  "<div style='width:100%;max-width:420px;margin:20px auto;" +

  "border:1px solid #d9c56b;border-radius:10px;" +

  "overflow:hidden;background:#00232d;color:white;" +

  "font-family:Arial,sans-serif;text-align:center'>" +



    // Compact poster section: QR code only on artwork.

    "<div style='background:#003c49;" +

    "background-image:url(cid:poster);" +

    "background-position:center;background-size:cover;" +

    "padding:18px'>" +



      "<div style='background:white;border:2px solid #f4e878;" +

      "border-radius:8px;padding:6px;width:120px;" +

      "max-width:65%;margin:auto'>" +

        "<img src='cid:" + key + "' width='120' " +

        "alt='Ticket QR code' " +

        "style='display:block;width:100%;height:auto'>" +

      "</div>" +



    "</div>" +



    // Compact details panel.

    "<div style='padding:14px 16px;background:#00232d'>" +



      "<div style='font-size:10px;letter-spacing:2px;" +

      "color:#f4e878'>SPECIAL SCREENING</div>" +



      "<div style='font-size:21px;font-weight:bold;" +

      "line-height:1.2;color:#f4e878;margin-top:5px'>" +

      "SONG OF THE FIREFLIES</div>" +



      "<div style='font-size:10px;letter-spacing:1.5px;" +

      "margin-top:5px'>OFFICIAL SCREENING TICKET</div>" +



      "<div style='font-size:11px;color:#f4e878;" +

      "margin:12px 0'>" +

      "TICKET " + (i + 1) + " OF " + tickets.length +

      "</div>" +



      "<div style='font-size:13px;line-height:1.6'>" +

      esc_(CFG.DATE) + " · " + esc_(CFG.TIME) +

      "<br>" + esc_(CFG.VENUE) +

      "</div>" +



      "<div style='border-top:1px solid #547078;" +

      "margin-top:10px;padding-top:9px'>" +



        "<div style='font-size:14px;font-weight:bold;" +

        "overflow-wrap:anywhere'>" +

        esc_(ticketName) + "</div>" +



        "<div style='font-size:11px;color:#d9c56b;" +

        "margin-top:4px;overflow-wrap:anywhere'>" +

        esc_(tid) + "</div>" +



      "</div>" +

    "</div>" +

  "</div>"

);









  });





const body =

  "<div style='font-family:Arial,sans-serif;color:#20343a;" +

  "line-height:1.6;max-width:640px;margin:auto'>" +



    "<div style='background:#00232d;padding:24px 20px;" +

    "text-align:center;border-radius:10px 10px 0 0'>" +

      "<div style='font-size:11px;letter-spacing:3px;" +

      "color:#f4e878'>OFFICIAL SCREENING TICKETS</div>" +

      "<h2 style='color:#f4e878;margin:10px 0 4px;" +

      "font-size:25px'>Song of the Fireflies</h2>" +

      "<p style='color:#ffffff;margin:0;font-size:13px'>" +

      "Your tickets are here! 🎟️</p>" +

    "</div>" +



    "<div style='padding:22px 18px;background:#ffffff;" +

    "border:1px solid #e0e6e8;border-top:0'>" +



      "<p>Hello " + esc_(name) + "!</p>" +



     "<p>We're pleased to confirm that your reservation for " +

     "<strong>Song of the Fireflies</strong> has been confirmed! " +

     "Thank you for being part of this special screening.</p>" +



      "<p>Your official screening tickets are included below. " +

      "Please keep this email handy and have your QR code ready " +

      "when you arrive at the venue.</p>" +



      cards.join("") +



      "<div style='background:#f5f7f7;border-left:4px solid #d9c56b;" +

      "padding:14px 16px;margin:24px 0'>" +

        "<h3 style='margin:0 0 8px;color:#00232d;font-size:15px'>" +

        "SCREENING DETAILS</h3>" +

        "<p style='margin:4px 0'><strong>Date:</strong> " +

        esc_(CFG.DATE) + "</p>" +

        "<p style='margin:4px 0'><strong>Time:</strong> " +

        esc_(CFG.TIME) + "</p>" +

        "<p style='margin:4px 0'><strong>Venue:</strong> " +

        esc_(CFG.VENUE) + "</p>" +

      "</div>" +



      "<h3 style='color:#00232d;font-size:15px'>" +

      "IMPORTANT REMINDERS</h3>" +

      "<ul style='padding-left:20px'>" +

        "<li>Each QR code represents one ticket and admits one person.</li>" +

        "<li>Present your QR code at the entrance for verification.</li>" +

        "<li>Each ticket is valid for one check-in only.</li>" +

        "<li>Please keep your QR codes secure to prevent unauthorized use.</li>" +

      "</ul>" +



      "<p style='margin-top:24px'><strong>Reservation ID:</strong><br>" +

      "<span style='color:#52676d;overflow-wrap:anywhere'>" +

      esc_(id) + "</span></p>" +



      "<p style='margin-top:24px'>Thank you for being part of this " +

      "special screening. We look forward to welcoming you!</p>" +



      "<p style='font-weight:bold;color:#00232d'>" +

      "See you at the screening! 🎬✨</p>" +



    "</div>" +



    "<div style='text-align:center;padding:14px;" +

    "font-size:11px;color:#6b7b80'>" +

      "Song of the Fireflies · Official Screening" +

    "</div>" +

  "</div>";





  MailApp.sendEmail({

    to: email,

    subject: "Your " + CFG.EVENT + " tickets",

    body: "Your payment is approved. Reservation: " + id +

      ". Screening: " + CFG.DATE + ", " + CFG.TIME +

      ", " + CFG.VENUE +

      ". Each ticket has its own QR code.",

    htmlBody: body,

    inlineImages: images

  });



  s.getRange(r, 6).setValue("Sent");

}





// QR verification page.



function doGet(e) {
  const id = String(
    (e && e.parameter && e.parameter.ticket) || ""
  ).trim();

  if (!id) {
    return page_("INVALID TICKET", false, "");
  }

  const ts = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Tickets");

  if (!ts) {
    return page_("CHECK-IN ERROR: Tickets sheet was not found.", false, id);
  }

  const rows = ts.getDataRange().getDisplayValues();
  const i = rows.findIndex(
    (row, index) => index > 0 && String(row[0]).trim() === id
  );

  if (i < 0) {
    return page_("INVALID TICKET", false, "");
  }

  // Six-column layout: column E (index 4) is Status.
  if (String(rows[i][4]).trim() !== "Unused") {
    return page_("ALREADY USED — ENTRY DENIED", false, id);
  }

  return page_("VALID TICKET", true, id);
}


// The owner scans the QR code and taps CHECK IN. No PIN is required.
function doPost(e) {
  const id = String(
    (e && e.parameter && e.parameter.ticket) || ""
  ).trim();

  if (!id) {
    return page_("INVALID TICKET", false, "");
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ts = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Tickets");

    if (!ts) {
      throw new Error("Tickets sheet was not found.");
    }

    const rows = ts.getDataRange().getDisplayValues();
    const i = rows.findIndex(
      (row, index) => index > 0 && String(row[0]).trim() === id
    );

    if (i < 0) {
      return page_("INVALID TICKET", false, "");
    }

    // Column E (index 4) is Status.
    if (String(rows[i][4]).trim() !== "Unused") {
      return page_("ALREADY USED — ENTRY DENIED", false, id);
    }

    // Column E = Status; column F = Check-in Time.
    ts.getRange(i + 1, 5).setValue("Used");
    ts.getRange(i + 1, 6).setValue(new Date());
    SpreadsheetApp.flush();

    return page_("CHECK-IN SUCCESSFUL!", false, id);

  } catch (err) {
    return page_(
      "CHECK-IN ERROR: " + err.message,
      false,
      id
    );
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


function page_(message, showButton, id) {

const base = CFG.WEB_APP_URL;



  const success = message === "CHECK-IN SUCCESSFUL!";

  const invalid = message === "INVALID TICKET";
  const valid = message === "VALID TICKET";

  const color = (success || valid) ? "#15803d" : "#b91c1c";



  const heading = (success || valid)
    ? "✓"
    : "⚠";



  const button = showButton

    ? "<form method='post' action='" + esc_(base) + "'>" +

      "<input type='hidden' name='ticket' value='" +

      esc_(id) + "'>" +

      "<button type='submit' style='" +

      "background:#b91c1c;color:white;border:none;" +

      "padding:16px 30px;border-radius:8px;" +

      "font-size:18px;font-weight:bold;" +

      "cursor:pointer;margin-top:20px'>" +

      "CHECK IN</button></form>"

    : "";



  const ticketInfo = id

    ? "<p style='font-size:16px;color:#475569'>" +

      "Ticket Number</p>" +

      "<p style='font-size:18px;font-weight:bold;" +

      "overflow-wrap:anywhere'>" + esc_(id) + "</p>"

    : "";



  const detail = success
    ? "This ticket has been checked in successfully."
    : valid
      ? "Ticket verified. Tap CHECK IN to mark this ticket as used."
      : invalid
        ? "This ticket could not be found. Please verify the QR code."
        : message === "ALREADY USED — ENTRY DENIED"
          ? "This ticket has already been checked in. Entry is denied."
          : "Please contact the registration team for assistance.";



  return HtmlService.createHtmlOutput(

    "<!DOCTYPE html>" +

    "<html><head>" +

    "<meta name='viewport' content='width=device-width, initial-scale=1'>" +

    "<meta charset='UTF-8'>" +

    "<title>SOTF Check-in</title>" +

    "</head>" +



    "<body style='margin:0;background:#f1f5f9;" +

    "font-family:Arial,sans-serif;text-align:center;" +

    "padding:35px 15px;color:#0f172a'>" +



      "<div style='max-width:420px;margin:30px auto;" +

      "background:white;padding:30px 20px;" +

      "border-radius:16px;box-shadow:0 4px 18px " +

      "rgba(0,0,0,.10)'>" +



        "<div style='font-size:60px;color:" + color + "'>" +

        heading + "</div>" +



        "<h1 style='font-size:25px;color:" + color + ";" +

        "line-height:1.3'>" + esc_(message) + "</h1>" +



        "<p style='font-size:16px;line-height:1.6;" +

        "color:#475569'>" + detail + "</p>" +



        ticketInfo +



        button +



        "<p style='font-size:12px;color:#64748b;" +

        "margin-top:25px'>SONG OF THE FIREFLIES</p>" +



      "</div>" +

    "</body></html>"

  ).setTitle("SOTF Check-in");

}









function esc_(v) {

  return String(v).replace(/[&<>"']/g, c => ({

    "&": "&amp;", "<": "&lt;", ">": "&gt;",

    '"': "&quot;", "'": "&#39;"

  })[c]);

}






// Run once before using the updated six-column Tickets layout.
// Makes a backup tab, then converts old or partially converted rows safely.
function migrateTicketsToSixColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = ss.getSheetByName("Tickets");

  if (!ts) throw new Error("Tickets sheet was not found.");

  const expected = [
    "Ticket ID", "Reservation ID", "Guest Name",
    "Buyer Email", "Status", "Check-in Time"
  ];

  const currentHeaders = ts.getRange(1, 1, 1, Math.max(ts.getLastColumn(), 1))
    .getDisplayValues()[0]
    .map(value => String(value).trim());

  if (
    currentHeaders.length === expected.length &&
    expected.every((header, index) => currentHeaders[index] === header)
  ) {
    SpreadsheetApp.getUi().alert("Tickets already uses the six-column layout. No migration was needed.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    "Migrate Tickets sheet?",
    "A backup copy of the current Tickets sheet will be created before the six-column layout is applied. Continue?",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const backupName = "Tickets Backup " +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd_HHmmss");
  ts.copyTo(ss).setName(backupName);

  const lastRow = ts.getLastRow();
  const lastCol = ts.getLastColumn();
  const oldRows = lastRow > 1
    ? ts.getRange(2, 1, lastRow - 1, lastCol).getValues()
    : [];

  const converted = oldRows
    .filter(row => row[0] !== "" && row[0] != null)
    .map(row => {
      // Already generated using six values but still has old headers:
      // A=Ticket ID, B=Reservation ID, C=Guest Name, D=Buyer Email, E=Status, F=Check-in Time.
      if (
        String(row[4] || "").trim() === "Unused" ||
        String(row[4] || "").trim() === "Used"
      ) {
        return [row[0], row[1], row[2], row[3], row[4], row[5] || ""];
      }

      // Original seven-column layout:
      // A=Ticket ID, B=Reservation ID, C=Ticket No., D=Name, E=Email, F=Status, G=Check-in Time.
      return [row[0], row[1], row[3], row[4], row[5] || "Unused", row[6] || ""];
    });

  ts.clearContents();
  ts.getRange(1, 1, 1, expected.length).setValues([expected]);
  if (converted.length) {
    ts.getRange(2, 1, converted.length, expected.length).setValues(converted);
  }
  ts.setFrozenRows(1);
  ts.autoResizeColumns(1, expected.length);

  ui.alert(
    "Migration complete",
    "The Tickets sheet now uses six columns. Backup created: " + backupName,
    ui.ButtonSet.OK
  );
}

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("SOTF Ticketing")

    .addItem("Approve Selected Payment", "approvePayment")

    .addItem("Send Selected Tickets", "sendTickets")

    .addToUi();

}