function exportGroupedCategories() {
  const sheetName = "Summary by Major Category";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create("Calendar Category Report");
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  const calendarId = "primary";
  const timeMin = new Date("2025-05-01T00:00:00Z").toISOString();
  const timeMax = new Date("2025-06-13T23:59:59Z").toISOString();

  const response = Calendar.Events.list(calendarId, {
    timeMin,
    timeMax,
    singleEvents: true,
    maxResults: 2500,
    orderBy: "startTime"
  });

  const colorIdMap = {
    "2": "Impedance",
    "3": "Meetings",
    "4": "Onboarding",
    "5": "Research/Scientific Support",
    "6": "Appointments/Sick time",
    "7": "Professional Development",
    "9": "Technical/Operational Support",
    "11": "Administrative"
  };

  const majorGroups = {
    "Administrative": ["Onboarding", "Administrative"],
    "Research/Science": ["Research/Scientific Support", "Impedance"],
    "Technical/Operational Support": ["Technical/Operational Support", "Professional Development", "Meetings"]
  };

  const individualTotals = {};
  Object.values(colorIdMap).forEach(cat => {
    individualTotals[cat] = { hours: 0, count: 0 };
  });

  const events = response.items || [];
  if (events.length === 0) {
    SpreadsheetApp.getUi().alert("No events found in the specified date range.");
    return;
  }

  for (const event of events) {
    const colorId = event.colorId || "0";
    const category = colorIdMap[colorId];

    const start = new Date(event.start.dateTime || event.start.date);
    const end = new Date(event.end.dateTime || event.end.date);
    const durationHours = (end - start) / (1000 * 60 * 60);

    Logger.log(`(All Events Log) Event: ${event.summary || "(No title)"} | Color ID: ${colorId} | Category: ${category || "Unmapped"} | Start: ${start} | End: ${end} | Duration: ${durationHours.toFixed(2)} hrs`);

    if (!category) continue;

    if (!individualTotals[category]) {
      individualTotals[category] = { hours: 0, count: 0 };
    }

    individualTotals[category].hours += durationHours;
    individualTotals[category].count += 1;
  }

  const majorTotals = {
    "Administrative": { hours: 0, count: 0 },
    "Research/Science": { hours: 0, count: 0 },
    "Technical/Operational Support": { hours: 0, count: 0 }
  };

  for (const [group, subcategories] of Object.entries(majorGroups)) {
    subcategories.forEach(sub => {
      if (individualTotals[sub]) {
        majorTotals[group].hours += individualTotals[sub].hours;
        majorTotals[group].count += individualTotals[sub].count;
      }
    });
  }

  const totalHours = Object.values(majorTotals).reduce((sum, obj) => sum + obj.hours, 0);

  sheet.getRange(1, 1, 1, 4).setValues([["Major Category", "Total Hours", "Event Count", "Percent of Time"]]);

  const rows = Object.entries(majorTotals).map(([group, { hours, count }]) => {
    const percent = totalHours > 0 ? ((hours / totalHours) * 100).toFixed(1) + "%" : "0%";
    return [group, hours.toFixed(2), count, percent];
  });

  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.autoResizeColumns(1, 4);

// create and populate ColorID Summary Tab
  const colorSummarySheetName = "ColorID Summary";
  let colorSummarySheet = spreadsheet.getSheetByName(colorSummarySheetName);
  if (!colorSummarySheet) {
    colorSummarySheet = spreadsheet.insertSheet(colorSummarySheetName);
  } else {
    colorSummarySheet.clearContents();
  }

  const totalIndividualHours = Object.values(individualTotals).reduce((sum, obj) => sum + obj.hours, 0);

  colorSummarySheet.getRange(1, 1, 1, 4).setValues([["Color Category", "Total Hours", "Event Count", "Percent of Time"]]);

  const colorRows = Object.entries(individualTotals).map(([category, { hours, count }]) => {
    const percent = totalIndividualHours > 0 ? ((hours / totalIndividualHours) * 100).toFixed(1) + "%" : "0%";
    return [category, hours.toFixed(2), count, percent];
  });

  colorSummarySheet.getRange(2, 1, colorRows.length, 4).setValues(colorRows);
  colorSummarySheet.autoResizeColumns(1, 4);
}

