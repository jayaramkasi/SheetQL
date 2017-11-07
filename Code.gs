/**
 * @OnlyCurrentDoc
 *
 */

/*
   This addon helps analysts create queries and post data to a spreadsheet on the fly (I have tested for mysql DB). 
   
   Developed for Rednile Innovations Private Limited. Released under MIT License.
*/

//   Add the option on Addons menu when the spreadsheet is open
function onOpen(e) {
  SpreadsheetApp.getActiveSpreadsheet().toast("Initializing addon menu","Initializing",1);
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Query', 'showSidebar')
  .addItem('Connection','connectionDetails')
  .addToUi();
  
}
// Add option to addon menu on installing the addon
function onInstall(e) {
  SpreadsheetApp.getActiveSpreadsheet().toast("Running onopen script","Initializing",1);
  onOpen(e);
  SpreadsheetApp.getActiveSpreadsheet().toast("Initializing user properties for the script","Initializing",1);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("serverIp", "")
    .setProperty("sqlPort", "")
    .setProperty("sqlUser", "")
    .setProperty("sqlPassword", "")
    .setProperty("sqlDB", "");

}

//Show contents of sidebar.html
function showSidebar(){
  SpreadsheetApp.getActiveSpreadsheet().toast("Initializing Sidebar","Initializing",1);
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Run Query');
  SpreadsheetApp.getUi().showSidebar(ui);
}

// This bit of code is not ready yet. Right now I have set UserProperties in File > Project Properties

function connectionDetails(){
  SpreadsheetApp.getActiveSpreadsheet().toast("Initializing Connection details","Initializing",1);
  var ui = HtmlService.createHtmlOutputFromFile('Connection')
      .setTitle('Set Mysql Connection details');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function getUserProperties(){
  SpreadsheetApp.getActiveSpreadsheet().toast("Retrieving user properties","Initializing",1);
  var userProperties = PropertiesService.getUserProperties();
  var connectionPrefs = {
    "serverIp": userProperties.getProperty("serverIp"),
    "sqlPort": userProperties.getProperty("sqlPort"),
    "sqlUser": userProperties.getProperty("sqlUser"),
    "sqlPassword": userProperties.getProperty("sqlPassword"),
    "sqlDB": userProperties.getProperty("sqlDB")
  }
  return connectionPrefs;
}
function setUserProperties(connectionProperties){
    SpreadsheetApp.getActiveSpreadsheet().toast("Updating user properties","Updating",1);
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperties(connectionProperties);
}

function runQuery(query){
  SpreadsheetApp.getActiveSpreadsheet().toast("Initializing connection details","Running",1);
  var userProperties = PropertiesService.getUserProperties();
  var ConnectString = 'jdbc:mysql://' + userProperties.getProperty("serverIp")+':'
                       +userProperties.getProperty("sqlPort")+'/'+userProperties.getProperty("sqlDB");

  var conn = Jdbc.getConnection(ConnectString, userProperties.getProperty("sqlUser"), userProperties.getProperty("sqlPassword"));
  var execStmt = conn.createStatement();
  SpreadsheetApp.getActiveSpreadsheet().toast("Executing query","Running",1);
  var mysqlQuery = execStmt.executeQuery(query);
  var resultsArray = [[]];
  var columns = [];
  // Get each column name and push to the result
  var numCols = mysqlQuery.getMetaData().getColumnCount();
  SpreadsheetApp.getActiveSpreadsheet().toast("Retrieving column names from metadata","Running",1);
  for(i=0;i<numCols;i++)
    columns.push({
      "name":mysqlQuery.getMetaData().getColumnName(i+1),
      "type":mysqlQuery.getMetaData().getColumnTypeName(i+1)
                     });
  var row = 0;
  SpreadsheetApp.getActiveSpreadsheet().toast("Retrieving results array","Running",1);
  while(mysqlQuery.next()){
    var rowArray = [];
    for(i=0;i<numCols;i++){
      rowArray[i] = columns[i].type == 'INT'?mysqlQuery.getInt(i+1) : mysqlQuery.getString(i+1);
    }
    resultsArray[row++] = rowArray;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Closing JDBC connection","Running",1);
  mysqlQuery.close();
  execStmt.close();
  conn.close();
  updateSheet( { results: resultsArray, columns: columns });
}

function updateSheet(data){
  SpreadsheetApp.getActiveSpreadsheet().toast("Updating sheet","Running",1);
  SpreadsheetApp.getActiveSheet().clear();
  SpreadsheetApp.getActiveSpreadsheet().toast("setting columns","Running",1);
  for(i=0;i<data.columns.length;i++){
    SpreadsheetApp.getActiveSheet().getRange(1,i+1).setValue(data.columns[i].name).setBackground("#439fd9").setFontColor("#FFFFFF");
  }
  SpreadsheetApp.getActiveSheet().getRange(2,1,data.results.length,data.columns.length).setValues(data.results);
  SpreadsheetApp.getActiveSpreadsheet().toast("Done. Refreshing sheet","Running",1);
}
