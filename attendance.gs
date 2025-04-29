// This function runs when the spreadsheet is opened and adds a custom menu.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance Tools')
    .addItem('Match and Highlight Attendance', 'matchAttendanceAndHighlightByAgent')
    .addToUi();
}

// Main function to match and highlight attendance by agent and add agent names to the AttendanceList.
function matchAttendanceAndHighlightByAgent() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the sheets
  var agentSheet = ss.getSheetByName('AgentList');
  var attendanceSheet = ss.getSheetByName('AttendanceList');
  
  if (!agentSheet || !attendanceSheet) {
    SpreadsheetApp.getUi().alert('Error: "AgentList" or "AttendanceList" sheet not found. Please check the sheet names.');
    return;
  }
  
  // Get data from AgentList: Worker ID in column B (index 1), Agent Name in column D (index 3)
  var agentData = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues(); 
  Logger.log('Agent Data: ' + JSON.stringify(agentData));
  
  // Get data from AttendanceList: Worker ID is now in columns C, D, and E
  var attendanceData = attendanceSheet.getRange(2, 3, attendanceSheet.getLastRow() - 1, 3).getValues();  // Columns C, D, E
  Logger.log('Attendance Data: ' + JSON.stringify(attendanceData));
  
  if (agentData.length === 0 || attendanceData.length === 0) {
    SpreadsheetApp.getUi().alert('Error: No data found in "AgentList" or "AttendanceList".');
    return;
  }

  // Define softer colors for different agents
  var colors = ['#92d050', '#4a86e8', '#e06666', '#ffd966', '#76a5af', '#ffb6c1', '#a64d79', '#674ea7', '#ff7f50', '#b7b7b7', '#c1fff4'];
  var agentColorMap = {};  // To store agent names with assigned colors
  var currentColorIndex = 0;
  
  // Create or clear a Results sheet to display the legend
  var resultSheet = ss.getSheetByName('Results');
  if (!resultSheet) {
    resultSheet = ss.insertSheet('Results');
  } else {
    resultSheet.clear();  // Clear old data
  }
  
  // Add header for the color legend
  resultSheet.getRange(1, 1).setValue("Agent Name");
  resultSheet.getRange(1, 2).setValue("Assigned Color");
  
  // Loop through AgentList and AttendanceList to match and highlight
  for (var i = 0; i < agentData.length; i++) {
    var agentWorkerID = agentData[i][1];  // Worker ID from AgentList (column B, index 1)
    var agentName = agentData[i][3];  // Agent Name from AgentList (column D, index 3)
    
    // Skip if Worker ID is empty (null or blank)
    if (!agentWorkerID) {
      continue;
    }
    
    // Assign a color for each agent if not already assigned
    if (!(agentName in agentColorMap)) {
      if (currentColorIndex < colors.length) {
        agentColorMap[agentName] = colors[currentColorIndex];
        currentColorIndex++;
      } else {
        agentColorMap[agentName] = colors[colors.length - 1]; // Assign last color if we run out
      }
    }
    
    // Add the agent name and its assigned color to the Results sheet for the legend
    resultSheet.getRange(i + 2, 1).setValue(agentName);
    resultSheet.getRange(i + 2, 2).setBackground(agentColorMap[agentName]);
    
    for (var j = 0; j < attendanceData.length; j++) {
      // Concatenate the Worker ID from columns C, D, and E in AttendanceList
      var attendanceWorkerID = attendanceData[j][0] + attendanceData[j][1] + attendanceData[j][2];
      
      // Skip if Worker ID in AttendanceList is empty (null or blank)
      if (!attendanceWorkerID) {
        continue;
      }

      if (agentWorkerID == attendanceWorkerID) {
        Logger.log('Match found: Agent Name: ' + agentName + ', Worker ID: ' + agentWorkerID);
        
        // Highlight the row in AttendanceList for the matched worker
        attendanceSheet.getRange(j + 2, 3, 1, attendanceSheet.getLastColumn() - 2).setBackground(agentColorMap[agentName]);

        // Add the agent name to column AP, starting from row 7 to row 86 in AttendanceList
    if (j + 2 >= 7 && j + 2 <= 86) { // Only update rows between 7 and 86
      attendanceSheet.getRange(j + 2, 42).setValue(agentName);  // Column AP is 42
        }
      }
    }
  }
  
  SpreadsheetApp.getUi().alert('Matching, highlighting, and legend creation completed successfully!');
}
// initial commit
