// Set API properties
var companyId = "";
var publicKey = '';
var privateKey = '';
var authRaw = companyId + '+' + publicKey + ':' + privateKey;
var auth = 'Basic ' + Utilities.base64Encode(authRaw);

function ProjectReport() {
  
  // Sheet details
  var sheetName = "Connectwise Integration";
  var cells = "A2:F50";
  
  // Clear the existing content
  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(cells).clearContent();
  
  // Display projects
  var row = 0;
  var newrange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(cells);
  var values = range.getValues();
  var projects = fetch("/project/projects?pagesize=100");
  for(var projectIndex in projects) {
    var project = projects[projectIndex];
    
    if(project.status.name == "New" || project.status.name == "In Progress") {
      
      // Display project attributes
      values[row][0] = project.company.name;
      values[row][1] = project.name;
      values[row][2] = project.status.name;
      
      // Display the latest note about project
      var notes = fetch("/project/projects/" + project.id + "/notes");
      for (var notesIndex in notes) {
        var note = notes[notesIndex];
        values[row][3] = note.text;
      }
      
      // Display the next step in a project
      var params = { "conditions": "project/id=" + project.id };
      var ticketsInProject = search("/service/tickets", params);
      for(var ticketIndex in ticketsInProject) {
        var ticket = ticketsInProject[ticketIndex];
        Logger.log(ticket);
        if(ticket.status.name == "Open" || ticket.status.name == "In Progress" || ticket.status.name == "Scheduled") { 
          
          // Display the project ticket summary with hyperlink
          values[row][4] = '=HYPERLINK("https://api-na.myconnectwise.net/v4_6_release/services/system_io/router/openrecord.rails?locale=en_US&companyName='+companyId+'&recordType=ServiceFV&recid='+ticket.id+'","' + ticket.summary + '")';
          
          if(ticket.requiredDate != undefined) {
            values[row][5] = ticket.requiredDate;
          }
          break;
        }
      }
      
      // Go to the next row
      row++;
    }
  }
  
  // Sort and display the projects
  range.setValues(values);
  range.sort([1,2]);
  
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
      Authorization: auth
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
      Authorization: auth
    }
  };
  
  Logger.log(url);
  Logger.log(options);
  
  // Return data
  return JSON.parse(UrlFetchApp.fetch(url, options));
}

