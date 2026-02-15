var SHEET_NAME = 'comments'; // 確保你的工作表名稱叫 comments

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('北海道討論區')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 取得所有留言
function getComments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  // 將 Date 物件轉為字串，避免 JSON 序列化問題
  var result = data.map(function(row) {
    return [
      row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
      String(row[1]),
      String(row[2])
    ];
  });

  result.reverse(); // 最新的在上面
  return result;
}

// 儲存留言
function addComment(nickname, content) {
  if (!nickname || !content) return false;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return false;
  sheet.appendRow([new Date(), nickname, content]);
  return true;
}