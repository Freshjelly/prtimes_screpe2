function fetchPRTIMES() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const keyword = sheet.getRange('A2').getValue();
  
  if (!keyword) {
    SpreadsheetApp.getUi().alert('A2セルに検索キーワードを入力してください。');
    return;
  }
  
  // ヘッダー行を設定
  sheet.getRange('B1').setValue('記事URL');
  sheet.getRange('C1').setValue('会社名');
  sheet.getRange('D1').setValue('担当者名');
  sheet.getRange('E1').setValue('メールアドレス');
  sheet.getRange('F1').setValue('電話番号');
  
  // 既存のデータをクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 5).clearContent();
  }
  
  try {
    // 記事URLを収集
    const articleUrls = collectArticleUrls(keyword);
    
    // 各記事から情報を抽出してスプレッドシートに出力
    extractAndOutputData(sheet, articleUrls);
    
    SpreadsheetApp.getUi().alert(`処理完了: ${articleUrls.length}件の記事を処理しました。`);
  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

function collectArticleUrls(keyword) {
  const baseUrl = 'https://prtimes.jp/main/html/searchrlp/search_result.html';
  const articleUrls = new Set();
  const maxPages = 3;
  const maxUrls = 50;
  
  for (let page = 1; page <= maxPages && articleUrls.size < maxUrls; page++) {
    try {
      const searchUrl = `${baseUrl}?search_word=${encodeURIComponent(keyword)}&page=${page}`;
      const response = UrlFetchApp.fetch(searchUrl);
      const html = response.getContentText();
      
      // 記事URLを抽出（/main/html/rd/p/xxxxx.html形式）
      const urlPattern = /\/main\/html\/rd\/p\/\d+\.html/g;
      const matches = html.match(urlPattern);
      
      if (matches) {
        matches.forEach(url => {
          if (articleUrls.size < maxUrls) {
            articleUrls.add('https://prtimes.jp' + url);
          }
        });
      }
      
      // レート制限対策
      Utilities.sleep(1000);
    } catch (error) {
      console.error(`ページ ${page} の取得でエラー:`, error);
    }
  }
  
  return Array.from(articleUrls);
}

function extractAndOutputData(sheet, articleUrls) {
  let row = 2;
  
  articleUrls.forEach((url, index) => {
    try {
      console.log(`処理中: ${index + 1}/${articleUrls.length} - ${url}`);
      
      const response = UrlFetchApp.fetch(url);
      const html = response.getContentText();
      
      // メディア関係者情報を抽出
      const mediaInfo = extractMediaInfo(html);
      
      // スプレッドシートに出力
      sheet.getRange(row, 2).setValue(url); // B列: 記事URL
      sheet.getRange(row, 3).setValue(mediaInfo.company); // C列: 会社名
      sheet.getRange(row, 4).setValue(mediaInfo.contact); // D列: 担当者名
      sheet.getRange(row, 5).setValue(mediaInfo.email); // E列: メールアドレス
      sheet.getRange(row, 6).setValue(mediaInfo.phone); // F列: 電話番号
      
      row++;
      
      // レート制限対策
      Utilities.sleep(1500);
      
    } catch (error) {
      console.error(`記事の処理でエラー (${url}):`, error);
      // エラーの場合はスキップして次へ
    }
  });
}

function extractMediaInfo(html) {
  const info = {
    company: '',
    contact: '',
    email: '',
    phone: ''
  };
  
  try {
    // HTMLから不要なタグを除去してテキストのみにする
    const text = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                    .replace(/<[^>]+>/g, ' ')
                    .replace(/\s+/g, ' ');
    
    // 会社名を抽出（株式会社、有限会社、合同会社など）
    const companyPatterns = [
      /(?:株式会社|有限会社|合同会社|合資会社|特定非営利活動法人|一般社団法人|一般財団法人|公益社団法人|公益財団法人)[^\s\n\r、。]+/g,
      /[^\s\n\r、。]+(?:株式会社|有限会社|合同会社|合資会社|Inc\.|LLC|Co\.,Ltd\.|Corporation)/g
    ];
    
    for (const pattern of companyPatterns) {
      const matches = text.match(pattern);
      if (matches && matches.length > 0) {
        info.company = matches[0].trim();
        break;
      }
    }
    
    // 担当者名を抽出
    const contactPatterns = [
      /担当[：:]\s*([^\s\n\r、。]+)/g,
      /連絡先[：:]\s*([^\s\n\r、。]+)/g,
      /問い合わせ先[：:]\s*([^\s\n\r、。]+)/g,
      /お問い合わせ[：:]\s*([^\s\n\r、。]+)/g
    ];
    
    for (const pattern of contactPatterns) {
      const match = pattern.exec(text);
      if (match && match[1]) {
        info.contact = match[1].trim();
        break;
      }
    }
    
    // メールアドレスを抽出
    const emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
    const emailMatches = text.match(emailPattern);
    if (emailMatches && emailMatches.length > 0) {
      // 一般的でないドメインを除外
      const validEmails = emailMatches.filter(email => 
        !email.includes('example.com') && 
        !email.includes('test.com') &&
        !email.includes('dummy.com')
      );
      if (validEmails.length > 0) {
        info.email = validEmails[0];
      }
    }
    
    // 電話番号を抽出
    const phonePatterns = [
      /(?:TEL|電話|Tel|tel)[：:\s]*([0-9\-\(\)]+)/g,
      /(\d{2,4}[-\s]?\d{2,4}[-\s]?\d{3,4})/g,
      /(0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4})/g
    ];
    
    for (const pattern of phonePatterns) {
      const match = pattern.exec(text);
      if (match && match[1]) {
        const phone = match[1].replace(/[^\d\-]/g, '');
        if (phone.length >= 10) {
          info.phone = phone;
          break;
        }
      }
    }
    
    // より具体的な電話番号パターン
    if (!info.phone) {
      const specificPhonePattern = /0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}/g;
      const phoneMatch = text.match(specificPhonePattern);
      if (phoneMatch && phoneMatch.length > 0) {
        info.phone = phoneMatch[0].replace(/\s/g, '');
      }
    }
    
  } catch (error) {
    console.error('情報抽出でエラー:', error);
  }
  
  return info;
}

// メニューに追加する関数
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PR TIMES取得')
    .addItem('記事情報を取得', 'fetchPRTIMES')
    .addToUi();
}