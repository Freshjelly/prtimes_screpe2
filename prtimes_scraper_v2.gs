function fetchPRTIMES() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const keyword = sheet.getRange('A2').getValue();
  
  if (!keyword) {
    SpreadsheetApp.getUi().alert('A2セルに検索キーワードを入力してください。');
    return;
  }
  
  // ヘッダー行を設定
  const headers = ['記事URL', '会社名', '担当者名', 'メールアドレス', '電話番号'];
  sheet.getRange(1, 2, 1, headers.length).setValues([headers]);
  
  // 既存のデータをクリア（2行目以降のB列〜F列）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 5).clearContent();
  }
  
  try {
    // 記事URLを収集
    const articleUrls = collectArticleUrls(keyword);
    console.log(`${articleUrls.length}件の記事URLを収集しました`);
    
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
      console.log(`検索ページ取得中: ${searchUrl}`);
      
      const response = UrlFetchApp.fetch(searchUrl, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        console.error(`ページ取得エラー: ${response.getResponseCode()}`);
        continue;
      }
      
      const html = response.getContentText();
      
      // 記事URLを抽出（複数のパターンで試行）
      const urlPatterns = [
        /href="(\/main\/html\/rd\/p\/\d+\.html)"/g,
        /href='(\/main\/html\/rd\/p\/\d+\.html)'/g,
        /\/main\/html\/rd\/p\/\d+\.html/g
      ];
      
      for (const pattern of urlPatterns) {
        const matches = [...html.matchAll(pattern)];
        if (matches.length > 0) {
          matches.forEach(match => {
            if (articleUrls.size < maxUrls) {
              const url = match[1] || match[0];
              const fullUrl = url.startsWith('http') ? url : 'https://prtimes.jp' + url;
              articleUrls.add(fullUrl);
            }
          });
          break;
        }
      }
      
      console.log(`ページ ${page}: ${articleUrls.size}件の記事URLを収集済み`);
      
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
  const batchData = [];
  
  articleUrls.forEach((url, index) => {
    try {
      console.log(`処理中: ${index + 1}/${articleUrls.length} - ${url}`);
      
      const response = UrlFetchApp.fetch(url, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        console.error(`記事取得エラー: ${response.getResponseCode()}`);
        return;
      }
      
      const html = response.getContentText();
      
      // メディア関係者情報を抽出
      const mediaInfo = extractMediaInfo(html);
      
      // データを配列に追加
      batchData.push([
        url,
        mediaInfo.company,
        mediaInfo.contact,
        mediaInfo.email,
        mediaInfo.phone
      ]);
      
      // 10件ごとにスプレッドシートに書き込み（API制限対策）
      if (batchData.length >= 10 || index === articleUrls.length - 1) {
        sheet.getRange(row, 2, batchData.length, 5).setValues(batchData);
        row += batchData.length;
        batchData.length = 0;
      }
      
      // レート制限対策
      Utilities.sleep(1500);
      
    } catch (error) {
      console.error(`記事の処理でエラー (${url}):`, error);
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
    // HTMLをクリーンアップ
    const cleanHtml = html
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/<!--[\s\S]*?-->/g, '');
    
    // メディア関係者情報セクションを探す
    const mediaPatterns = [
      /メディア関係者.*?お問い合わせ先[\s\S]*?<\/[^>]+>/gi,
      /報道関係者.*?お問い合わせ[\s\S]*?<\/[^>]+>/gi,
      /プレスリリース.*?お問い合わせ[\s\S]*?<\/[^>]+>/gi,
      /お問い合わせ先[\s\S]{0,500}?(?=<\/div|<\/section|<hr|$)/gi
    ];
    
    let targetSection = '';
    for (const pattern of mediaPatterns) {
      const match = cleanHtml.match(pattern);
      if (match) {
        targetSection = match[0];
        break;
      }
    }
    
    // セクションが見つからない場合は全体から検索
    const searchText = targetSection || cleanHtml;
    const text = searchText.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
    
    // 会社名を抽出
    const companyPatterns = [
      /(?:会社名|企業名|社名)[：:：\s]*([^\s\n\r、。；;]+(?:株式会社|有限会社|合同会社|Inc\.|LLC|Corporation))/,
      /(株式会社[^\s\n\r、。；;]{1,30})/,
      /([^\s\n\r、。；;]{1,30}株式会社)/,
      /((?:株式会社|有限会社|合同会社)[^\s\n\r、。；;]{1,30})/
    ];
    
    for (const pattern of companyPatterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        info.company = match[1].trim()
          .replace(/\s+/g, '')
          .replace(/[【】\[\]]/g, '');
        break;
      }
    }
    
    // 担当者名を抽出
    const contactPatterns = [
      /(?:担当者?|連絡先|広報|PR)[：:：\s]*([^\s\n\r、。；;]{2,10}(?:様|さん)?)/,
      /(?:お問い合わせ先?|問合せ先?)[：:：\s]*([^\s\n\r、。；;]{2,10})/,
      /(?:氏名|お名前)[：:：\s]*([^\s\n\r、。；;]{2,10})/
    ];
    
    for (const pattern of contactPatterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        const contact = match[1].trim();
        // 会社名や部署名を除外
        if (!contact.includes('会社') && !contact.includes('部') && 
            !contact.includes('課') && !contact.includes('室')) {
          info.contact = contact.replace(/[様さん]$/, '');
          break;
        }
      }
    }
    
    // メールアドレスを抽出
    const emailPattern = /[a-zA-Z0-9][a-zA-Z0-9._%+-]*@[a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,}/g;
    const emailMatches = text.match(emailPattern);
    if (emailMatches && emailMatches.length > 0) {
      // PR関連のメールアドレスを優先
      const prEmails = emailMatches.filter(email => 
        email.toLowerCase().includes('pr') || 
        email.toLowerCase().includes('press') ||
        email.toLowerCase().includes('media') ||
        email.toLowerCase().includes('info')
      );
      
      info.email = prEmails.length > 0 ? prEmails[0] : emailMatches[0];
    }
    
    // 電話番号を抽出
    const phonePatterns = [
      /(?:TEL|Tel|tel|電話|℡)[：:：\s]*([0-9０-９]{2,4}[-－ー\s]?[0-9０-９]{2,4}[-－ー\s]?[0-9０-９]{3,4})/,
      /(?:TEL|Tel|tel|電話|℡)[：:：\s]*([0-9\-\(\)]+)/,
      /(0[0-9]{1,4}[-－ー\s]?[0-9]{1,4}[-－ー\s]?[0-9]{3,4})/
    ];
    
    for (const pattern of phonePatterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        // 全角数字を半角に変換
        let phone = match[1]
          .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
          .replace(/[－ー]/g, '-')
          .replace(/[^\d\-]/g, '');
        
        if (phone.match(/^\d{2,4}-?\d{2,4}-?\d{3,4}$/)) {
          info.phone = phone;
          break;
        }
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