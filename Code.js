const FP_LOG_SHEET = "FP Log";
const VIP_SHEET = "VIP_Responses";

const HEADERS = [
  "Timestamp",
  "Presenter Date",
  "Presenter Name",
  "Business Type",
  "Presenter Contact",
  "Member Name",
  "Role",
  "Prospect Name",
  "Industry",
  "Location",
  "Connection Type",
  "Comfort Level",
  "Support Type",
  "Follow-up Owner",
  "Notes"
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("BNI VIP Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUpcomingWednesday() {
  const now = new Date();
  const today = new Date(now);
  today.setHours(0,0,0,0);

  let day = today.getDay();
  let daysUntilWed = (3 - day + 7) % 7;

  if (daysUntilWed === 0 && now.getHours() >= 7) {
    daysUntilWed = 7;
  }

  const upcomingWed = new Date(today);
  upcomingWed.setDate(today.getDate() + daysUntilWed);
  return upcomingWed;
}

function getUpcomingPresenters() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(FP_LOG_SHEET);
    
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    const upcomingWed = getUpcomingWednesday();
    upcomingWed.setHours(0,0,0,0);

    const presenters = [];

    // Skip header row (i=1)
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][1]); // Column B - Date of Meeting
      rowDate.setHours(0,0,0,0);

      // Filter for upcoming Wednesday only
      if (rowDate.getTime() === upcomingWed.getTime()) {
        presenters.push({
          name: data[i][3] || "Unknown",          // Column D - Presenter Name
          business: data[i][4] || "N/A",          // Column E - Business Type
          contact: data[i][5] || "",              // Column F - Contact Info
          date: Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd")
        });
      }
    }

    return presenters;
    
  } catch (error) {
    Logger.log("Error in getUpcomingPresenters: " + error.toString());
    return [];
  }
}

function submitVIP(formData) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(VIP_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(VIP_SHEET);
    }

    const lastRow = sheet.getLastRow();
    
    if (lastRow === 0) {
      sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1,1,1,HEADERS.length).setFontWeight("bold");
    }

    sheet.appendRow([
      new Date(),
      formData.presenter_date,
      formData.presenter_name,
      formData.presenter_business,
      formData.presenter_contact,
      formData.member_name || "",
      formData.role || "",
      formData.prospect_name || "",
      formData.industry || "",
      formData.location || "",
      formData.connection_type || "",
      formData.comfort_level || "",
      formData.support_type || "",
      formData.followup_owner || "",
      formData.notes || ""
    ]);

    return {status: "success", message: "VIP recorded successfully!"};
    
  } catch (error) {
    Logger.log("Error in submitVIP: " + error.toString());
    return {status: "error", message: "Failed to submit. Please try again."};
  }
}
