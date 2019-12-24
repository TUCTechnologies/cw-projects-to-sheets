// Set API properties
var companyId = '';
var publicKey = '';
var privateKey = '';
var clientId = '';
var authRaw = companyId + '+' + publicKey + ':' + privateKey;
var auth = 'Basic ' + Utilities.base64Encode(authRaw);

function ProjectReport() {
  
  // Sheet details
  var sheetName = "Connectwise Integration";
  var cells = "A2:F50";
  
  // Clear the existing content
  //var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(cells).clearContent();
  
  // Display projects
  var row = 0;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  //var newrange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(cells);
  //var values = range.getValues();
  
  //var projectParams = { "conditions": 'project/status/name="New" or project/status/name="In Progress"' };
  //var projects = search("'project/projects",projectParams);
  
  // Display header
  var row = 1;
  sheet.getRange(row,1).setValue("Company");
  sheet.getRange(row,2).setValue("Project");
  sheet.getRange(row,3).setValue("Status");
  sheet.getRange(row,4).setValue("Comments");
  sheet.getRange(row,5).setValue("Next Ticket is...");
  sheet.getRange(row,5).setValue("Ticket Due Date");
  row++;
  
  // Display project info
  var projects = fetch("/project/projects?pagesize=100");
  for each(var project in projects) {
    var column = 1;
    if(project.status.name == "New" || project.status.name == "In Progress" || project.status.name == "On-Hold") {
      
      sheet.getRange(row,column).setValue(project.company.name); column++;
      sheet.getRange(row,column).setValue(project.name); column++;
      sheet.getRange(row,column).setValue(project.status.name); column++;
      
      // Display the latest note about project
      var notes = fetch("/project/projects/" + project.id + "/notes");
      for each(var note in notes) {
        sheet.getRange(row,column).setValue(note.text);
      }
      column++;
      
      // Display the next step in a project
      var params = { "conditions": "project/id=" + project.id};
      var ticketsInProject = search("/project/tickets", params);
      for each(var ticket in ticketsInProject) {
        if(ticket.status.name == "Open" || ticket.status.name == "In Progress" || ticket.status.name == "Scheduled") { 
          
          // Display the project ticket summary with hyperlink
          sheet.getRange(row,column).setValue('=HYPERLINK("https://api-na.myconnectwise.net/v4_6_release/services/system_io/router/openrecord.rails?locale=en_US&companyName='+companyId+'&recordType=ServiceFV&recid='+ticket.id+'","' + ticket.summary + '")');
          column++;
          
          if(ticket.requiredDate != undefined) {
            sheet.getRange(row,column).setValue(ticket.requiredDate);
          }
          break;
        }
      }
      
      // Go to the next row
      row++;
    }
  }
  
  // Sort and display the projects
  sheet.getRange(2,1,row,column).sort([1,2]);
  
}

/* 
  Returns data from Connectwise API
*/
function fetch(path)
{
  var url = "https://api-na.myconnectwise.net/v4_6_release/apis/3.0" + path;
  var options = {
    method : 'get',
    contentType: "application/json",
    headers: {
      Authorization: auth,
      clientId: clientId
    },
    muteHttpExceptions: false
  };
  
  // Return data
  return JSON.parse(UrlFetchApp.fetch(url, options));
}

/* 
  Returns data from Connectwise API search
*/
function search(path, params)
{
  var url = "https://api-na.myconnectwise.net/v4_6_release/apis/3.0" + path + "/search";
  var options = {
    method : 'post',
    contentType: "application/json",
    payload: JSON.stringify(params),
    headers: {
      Authorization: auth,
      clientId: clientId
    }
  };
  
  // Return data
  return JSON.parse(UrlFetchApp.fetch(url, options));
}

