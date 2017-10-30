/**
 * @OnlyCurrentDoc
 *
 */

/*
   This addon helps analysts create queries and post data to a spreadsheet on the fly (I have tested for mysql DB). 
   
   Connection details are to be given as User preferences. I have not worked out how to add that yet. Can someone help me with that?
*/

//   Add the option on Addons menu when the spreadsheet is open
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Query', 'showSidebar')
  .addItem('Connection','connectionDetails')
  .addToUi();
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("serverIp", "");
  userProperties.setProperty("sqlPort", "");
  userProperties.setProperty("sqlUser", "");
  userProperties.setProperty("sqlPassword", "");
  userProperties.setProperty("sqlDB", "");
}
// Add option to addon menu on installing the addon
function onInstall(e) {
  onOpen(e);
}

//Show contents of sidebar.html
function showSidebar(){
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Run Query');
  SpreadsheetApp.getUi().showSidebar(ui);
}

// This bit of code is not ready yet. Right now I have set UserProperties in File > Project Properties

function connectionDetails(){
  var ui = HtmlService.createHtmlOutputFromFile('Connection')
      .setTitle('Set Mysql Connection details');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function getUserProperties(){
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
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperties(connectionProperties);
}

function runQuery(query){
  var userProperties = PropertiesService.getUserProperties();
  var ConnectString = 'jdbc:mysql://' + userProperties.getProperty("serverIp")+':'
                       +userProperties.getProperty("sqlPort")+'/'+userProperties.getProperty("sqlDB");

  var conn = Jdbc.getConnection(ConnectString, userProperties.getProperty("sqlUser"), userProperties.getProperty("sqlPassword"));
  var execStmt = conn.createStatement();
  
  var mysqlQuery = execStmt.executeQuery(query);
  var resultsArray = [];
  var columns = [];
  // Get each column name and push to the result
  var numCols = mysqlQuery.getMetaData().getColumnCount();
  for(i=0;i<numCols;i++)
    columns.push({
      "name":mysqlQuery.getMetaData().getColumnName(i+1),
      "type":mysqlQuery.getMetaData().getColumnTypeName(i+1)
                     });
  
  while(mysqlQuery.next()){
    var rowObject = {};
    for(i=0;i<numCols;i++){
      rowObject[columns[i].name] = columns[i].type == 'INT'?mysqlQuery.getInt(i+1): mysqlQuery.getString(i+1);
    }
    resultsArray.push(rowObject);
  }
  mysqlQuery.close();
  execStmt.close();
  conn.close();
  updateSheet( { results: resultsArray, columns: columns });
}

function updateSheet(data){
  SpreadsheetApp.getActiveSheet().clear(); // Clear sheet first
  for(i=0;i<data.columns.length;i++){
    SpreadsheetApp.getActiveSheet().getRange(1,i+1).setValue(data.columns[i].name);
  }
  for(i=0;i<data.results.length;i++){
    for(j=0;j<data.columns.length;j++){
      SpreadsheetApp.getActiveSheet().getRange(i+2,j+1).setValue(data.results[i][data.columns[j].name]);
      SpreadsheetApp.getActiveSheet().getRange(i+2,j+1).set
    }
  }
}
