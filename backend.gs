function doGet (e) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log').appendRow([JSON.stringify(e)])
  let action = e.parameter.action
  switch (action) {
      case 'login':
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName('login record').appendRow([decodeURI(e.queryString)])
          return login();
      case 'addTodo':
          return addTodo(e); // 另一個處理函數
      case 'todoList':
          return todoList(e);
      case 'register':
          return register(e);
      default:
          result = "無效的動作";
  }
}

function login(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('login record')
  const rowNum = sheet.getLastRow()
  const loginValue = sheet.getRange(rowNum, 4).getValue();
    
    if (loginValue === true) {
        const userID = sheet.getRange(rowNum,2).getValue();
        return ContentService.createTextOutput("true")
    } else {
        return ContentService.createTextOutput("false")
    }
  }

function register(e){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log').appendRow([JSON.stringify(e)])
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('user data')
  let userID = e.parameter.userID 
  let res = sheet.getDataRange().getValues()
    .filter(row => row[1] === userID)
  if (hasValue(res) === false) {
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName('user data').appendRow([e.queryString])
    return ContentService.createTextOutput("true")
  }
    else {
    return ContentService.createTextOutput("false")
    }
  }

function addTodo(e) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('task').appendRow([e.queryString])
  return ContentService.createTextOutput("true")
  }
  
function todoList(e){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log').appendRow([JSON.stringify(e)])
  let userID = e.parameter.userID  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('task')
  let res = sheet.getDataRange().getValues()
    .filter(row => row[1] === userID).map(row => row[0]);
  const jsonOutput = JSON.stringify(res)
  return ContentService.createTextOutput(jsonOutput)
  }

function hasValue(arr) {
    // 檢查是否為數組且不為空
    if (!Array.isArray(arr)) {
        return false; // 如果不是數組，返回 false
    }
    
    return arr.some(value => value !== undefined && value !== null && value !== '');
}
