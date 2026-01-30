/**
 * ウェブアプリの初期表示
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('学内連絡・更新システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * メール送信処理
 * targetGroup: '全員', '1年生', '2年生', '3年生', '4年生', '○○ゼミ' など
 */
function sendEmailToGroup(subject, body, targetGroup) {
  const spreadsheetId = '1Ij8TXyQEaj34yaKTUjjGIK0Due0612UBuIUM3FsTYm0';
  const sheetName = 'Sheet1'; 
  
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('データが登録されていません。');

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    let sentCount = 0;

    data.forEach(row => {
      const email = row[0].toString().trim();
      const name = row[1].toString().trim();
      const group = row[2].toString().trim();

      if (targetGroup === '全員' || group === targetGroup) {
        if (email && email.match(/.+@.+\..+/)) {
          const personalizedBody = body.replace(/\{\{NAME\}\}/g, name + 'さん');
          MailApp.sendEmail({
            to: email,
            subject: subject,
            body: personalizedBody
          });
          sentCount++;
        }
      }
    });

    return { success: true, count: sentCount };
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * メールアドレス更新処理
 */
function updateEmailAddress(targetName, newEmail) {
  const spreadsheetId = '1Ij8TXyQEaj34yaKTUjjGIK0Due0612UBuIUM3FsTYm0';
  const sheetName = 'Sheet1';
  
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B列（Name）を取得
    
    let foundRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString().trim() === targetName) {
        foundRow = i + 2; // 実際の行番号
        break;
      }
    }
    
    if (foundRow === -1) throw new Error('その名前のユーザーは見つかりません。正確に入力してください。');
    
    // A列（Email）を更新
    sheet.getRange(foundRow, 1).setValue(newEmail);
    return { success: true };
  } catch (e) {
    throw new Error(e.message);
  }
}