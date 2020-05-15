/**
* 送信URLが記載されたメールを一斉送信
*
* @param None
* @return None
*/
function mailBroadcast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1');
  const dataRange = sheet.getDataRange();
  
  // 名前、会社名、E-mail、送信URLの列数を取得（列数固定だと、列数が変わったときに機能しなくなるため）
  let nameIndex = 0;
  let companyNameIndex = 0;
  let emailIndex = 0;
  let sendUrlIndex = 0;
  for (var i = 1; i <= dataRange.getNumColumns(); i++) {
    switch (sheet.getRange(row=1, column=i).getValue()) {
      case '名前':
        nameIndex = i -1;  // 二次元配列のループは0から始まるため、列数を -1 する（以下同様）
      case '会社名':
        companyNameIndex = i -1; 
      case 'E-mail':
        emailIndex = i -1;
      case '送信URL':
        sendUrlIndex = i -1;
      default:
        ;
    }
  }
    
  // E-mail送信
  const subject = 'アンケートご協力のお願い';
  const data = dataRange.getValues();
  for (var i = 1; i < dataRange.getNumRows(); i++) {
    let name = data[i][nameIndex];  // 0行目はヘッダー情報のためスキップ（以下同様）
    let company = data[i][companyNameIndex];
    let email = data[i][emailIndex];
    let sendUrl = data[i][sendUrlIndex];
    
    // メール本文
    let body = `
    ${company}
    ${name}様
    
    お世話になっております。
    ●●でございます。
    
    この度、${name}様の率直なお気持ち、ご意見を頂戴したく、
    アンケートにご協力頂きたく存じます。
    
    【アンケートURL】
    ${sendUrl}
    
    何卒宜しくお願い申し上げます。
    
    ●●
    `
    
    try {
      GmailApp.sendEmail(email, subject, body, options={from: 'abc@gmail.com', bcc: 'xyz@gmail.com'});
      console.log('送信OK: ${name} ${company} ${email} ${sendUrl}');
    } catch (e) {
      console.log('送信NG: ${name} ${e.name} ${e.message}');
    }
  }
}
