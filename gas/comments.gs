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

// ===== 投票功能 (Voting) =====
var VOTE_SHEET_NAME = 'votes';

// 投票
function submitVote(nickname, option) {
  if (!nickname || !option) return { success: false, message: '請填寫暱稱並選擇方案！' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VOTE_SHEET_NAME);
  
  // 如果沒有這個 sheet，自動建立
  if (!sheet) {
    sheet = ss.insertSheet(VOTE_SHEET_NAME);
    sheet.appendRow(['時間', '暱稱', '選項']); // header
  }
  
  // 檢查是否投過票 (簡單檢查：同暱稱不能重複投)
  // 若要更嚴謹可用 cookie 或 IP，但 GAS 抓 IP 不易，這裡先用暱稱與簡單 cookie (前端控制)
  var data = sheet.getDataRange().getValues(); // [[Time, Nick, Opt], ...]
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(nickname).trim()) {
      return { success: false, message: '這個暱稱已經投過票囉！' };
    }
  }

  sheet.appendRow([new Date(), nickname, option]);
  return { success: true, message: '投票成功！' };
}

// 取得投票結果
function getVoteResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VOTE_SHEET_NAME);
  
  var results = { A: 0, B: 0, C: 0 };
  
  if (!sheet) return results;
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return results; // 只有標題或空的
  
  var data = sheet.getRange(2, 3, lastRow - 1, 1).getValues(); // 取第3欄 (選項)
  
  // 計算票數
  data.forEach(function(row) {
    var opt = String(row[0]).trim().toUpperCase(); // 轉大寫避免大小寫問題
    // 如果選項是 A, B, C 其中之一 (或包含 "方案A" 之類的關鍵字，這裡先精確比對 A, B, C)
    if (results.hasOwnProperty(opt)) {
      results[opt]++;
    } else if (opt.includes('A')) { results.A++; }
    else if (opt.includes('B')) { results.B++; }
    else if (opt.includes('C')) { results.C++; }
  });
  
  return results;
}