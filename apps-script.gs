// === Google Apps Script — MBTI 数据后端 ===
// 部署步骤（5分钟搞定）：
// 1. 打开 sheets.google.com，新建空白表格
// 2. 菜单：扩展程序 → Apps Script
// 3. 把这段代码全部粘贴进去
// 4. 点击"部署"→"新部署"→类型选"web 应用"
//    执行身份：我，访问权限：任何人（Anyone）
// 5. 复制部署后得到的 URL
// 6. 把 URL 填回 index.html 和 admin.html 的 SHEETS_URL 变量中

// ====== 配置区 ======
var SHEET_NAME = 'Sheet1';
var ADMIN_PASSWORD = 'mbti2026'; // 管理后台密码

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // 如果表头为空，先写表头
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['昵称', '类型', 'E/I分', 'S/N分', 'T/F分', 'J/P分', 'E/I%', 'S/N%', 'T/F%', 'J/P%', '提交时间']);
      sheet.setFrozenRows(1);
      // 自动加粗表头
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    }

    sheet.appendRow([
      data.nickname || '匿名',
      data.type,
      data.ei_score,
      data.sn_score,
      data.tf_score,
      data.jp_score,
      data.ei_pct,
      data.sn_pct,
      data.tf_pct,
      data.jp_pct,
      data.time || new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' })
    ]);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.action !== 'read') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'forbidden' }));
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (sheet.getLastRow() <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    var data = values.map(function(row) {
      return {
        nickname: row[0],
        type: row[1],
        ei_score: row[2],
        sn_score: row[3],
        tf_score: row[4],
        jp_score: row[5],
        ei_pct: row[6],
        sn_pct: row[7],
        tf_pct: row[8],
        jp_pct: row[9],
        time: row[10]
      };
    });

    return ContentService.createTextOutput(JSON.stringify({ data: data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
