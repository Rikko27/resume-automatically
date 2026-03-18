/**
 * 職務経歴書自動生成ツール (Resume Generation Tool)
 * 
 * このスクリプトは、Googleカレンダーの面談イベントから文字起こし内容を取得し、
 * Gemini APIを使用してプロフェッショナルな職務経歴書を生成します。
 */

// ==========================================
// 設定項目 (Settings)
// ==========================================
const CONFIG = {
  // ★ APIキーはスクリプトプロパティ「GEMINI_API_KEY」から取得します
  // GitHubに公開しても安全な形式になりました
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  
  // ★ 使用するGeminiのモデル名
  // ご自身のリストに存在し、かつエラーが出にくい「gemini-flash-latest」を指定します
  GEMINI_MODEL: 'gemini-flash-latest',
  
  // 検索対象のキーワード
  SEARCH_KEYWORD: '面談',
  
  // 取得を開始する特定の日付（例: '2026-03-01'）
  // 指定しない（null）場合は、DAYS_BACKが適用されます
  FETCH_START_DATE: null,
  
  // 何日前までのイベントを取得するか（FETCH_START_DATEがnullの場合のみ使用）
  DAYS_BACK: 21,
  
  // 生成された職務経歴書を保存するフォルダ名
  OUTPUT_FOLDER_NAME: '生成済み職務経歴書',
  
  // スプレッドシートのシート名
  SHEET_NAME: '面談一覧'
};

// ==========================================
// メイン関数 (Main Functions)
// ==========================================

/**
 * スプレッドシートを開いた時に実行される
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('職務経歴書ツール')
    .addItem('カレンダーから面談を取得', 'fetchRecentInterviews')
    .addSeparator()
    .addItem('選択した行の職務経歴書を生成', 'generateResumeFromSelectedRow')
    .addSeparator()
    .addItem('APIモデル一覧の確認（デバッグ用）', 'debugListModels')
    .addToUi();
}

/**
 * カレンダーから最近の面談イベントを取得し、スプレッドシートに書き出す
 */
function fetchRecentInterviews() {
  const sheet = getOrCreateSheet();
  const now = new Date();
  
  let startTime;
  if (CONFIG.FETCH_START_DATE) {
    startTime = new Date(CONFIG.FETCH_START_DATE);
    startTime.setHours(0, 0, 0, 0);
  } else {
    startTime = new Date(now.getTime() - (CONFIG.DAYS_BACK * 24 * 60 * 60 * 1000));
  }
  
  // 「野山」という名前のカレンダーを探す
  let targetCalendar = CalendarApp.getCalendarsByName(CONFIG.SEARCH_KEYWORD)[0];
  if (!targetCalendar) {
    targetCalendar = CalendarApp.getDefaultCalendar();
    Logger.log('「' + CONFIG.SEARCH_KEYWORD + '」という名前のカレンダーが見つからなかったため、デフォルトのカレンダーを使用します。');
  } else {
    Logger.log('「' + CONFIG.SEARCH_KEYWORD + '」カレンダーから読み込みます。');
  }

  const events = targetCalendar.getEvents(startTime, now);
  
  // フィルタリング：タイトルまたは説明にキーワードが含まれる、または「氏名：」がある
  const interviewEvents = events.filter(e => {
    const title = e.getTitle();
    const desc = e.getDescription();
    return title.includes(CONFIG.SEARCH_KEYWORD) || desc.includes(CONFIG.SEARCH_KEYWORD) || desc.includes('氏名：');
  });
  
  if (interviewEvents.length === 0) {
    SpreadsheetApp.getUi().alert('対象のイベント（' + CONFIG.SEARCH_KEYWORD + '）が見つかりませんでした。期間：' + Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd') + ' ～ ' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd'));
    return;
  }
  
  // 既存のデータを取得（重複チェック用）
  const existingTitlesRange = sheet.getRange(2, 3, Math.max(sheet.getLastRow() - 1, 1), 1);
  const existingTitles = existingTitlesRange.getValues().flat();
  
  interviewEvents.forEach(event => {
    const title = event.getTitle();
    if (!existingTitles.includes(title)) {
      const date = Utilities.formatDate(event.getStartTime(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const description = event.getDescription();
      const candidateName = extractCandidateName(description) || '不明';
      
      sheet.appendRow([
        false, // 選択用チェックボックス
        date,
        title,
        candidateName,
        '未作成',
        '', // リンク
        description // 隠し列または詳細用
      ]);
    }
  });
  
  // チェックボックスを全行に再適用（appendRowでの値上書き対策）
  const maxRows = sheet.getMaxRows();
  if (maxRows >= 2) {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 1, maxRows - 1, 1).setDataValidation(rule);
  }
  
  SpreadsheetApp.getUi().alert('面談イベントを取得しました。');
}

/**
 * 選択された行のデータを使って職務経歴書を生成する
 */
function generateResumeFromSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「' + CONFIG.SHEET_NAME + '」シートが見つかりません。先に「カレンダーから面談を取得」を実行してください。');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('シートにデータがありません。先に「カレンダーから面談を取得」を実行してください。');
    return;
  }
  
  let processedCount = 0;
  let selectedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // チェックボックスの値（Booleanとして判定、または文字列の"TRUE"を考慮）
    const isSelected = (row[0] === true || String(row[0]).toUpperCase() === 'TRUE');
    const status = row[4];
    
    if (isSelected) {
      selectedCount++;
      
      if (status === '完了') {
        continue;
      }
      
      const dateStr = row[1];
      const title = row[2];
      const candidateName = row[3];
      const memo = row[6];
      
      try {
        const resumeData = callGeminiAPI(memo);
        const docUrl = createResumeDocument(candidateName, String(dateStr), resumeData);
        
        // シートを更新
        sheet.getRange(i + 1, 5).setValue('完了');
        sheet.getRange(i + 1, 6).setValue(docUrl);
        sheet.getRange(i + 1, 1).setValue(false); // チェックボックスを外す
        
        processedCount++;
      } catch (e) {
        Logger.log('エラーが発生しました: ' + e.toString());
        sheet.getRange(i + 1, 5).setValue('エラー');
        SpreadsheetApp.getUi().alert('「' + candidateName + '」様の生成中にエラーが発生しました：\n' + e.toString());
      }
    }
  }
  
  if (processedCount > 0) {
    SpreadsheetApp.getUi().alert(processedCount + ' 件の職務経歴書を生成しました。');
  } else if (selectedCount === 0) {
    SpreadsheetApp.getUi().alert('左端の「選択」欄にチェックが入っている行がありません。\n作成したい人の列にあるチェックボックスをクリックして、✔︎を入れてから実行してください。');
  } else {
    SpreadsheetApp.getUi().alert('選択された行はすべて「完了」になっています。新しく作成する場合は、ステータスを消すか、別の行を選択してください。');
  }
}

// ==========================================
// ユーティリティ関数 (Utility Functions)
// ==========================================

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['選択', '日程', 'イベント名', '候補者名', 'ステータス', 'URL', 'メモ(非表示)']);
    sheet.setFrozenRows(1);
    // メモ列を非表示に
    sheet.hideColumns(7);
  }
  
  // 常にA2以降のデータがある範囲（シートの最大行まで）をチェックボックス形式にする
  const lastRow = sheet.getMaxRows();
  if (lastRow >= 2) {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 1, lastRow - 1, 1).setDataValidation(rule);
  }
  
  return sheet;
}

function extractCandidateName(description) {
  const match = description.match(/氏名：\s*(.+)/);
  return match ? match[1].trim() : null;
}

function callGeminiAPI(memo) {
  const apiKey = CONFIG.GEMINI_API_KEY;
  if (apiKey === 'YOUR_API_KEY_HERE' || apiKey === 'YOUR_GEMINI_API_KEY') {
    throw new Error('GeminiのAPIキーが設定されていません。13行目の設定を確認してください。');
  }

  // 互換性の高い v1beta エンドポイントを使用
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const prompt = `あなたはプロのキャリアアドバイザー(CA)です。以下の構成案に従い、提供された面談メモからプロフェッショナルな職務経歴書を作成してください。
出力は必ず以下のJSONフォーマットに厳密に従い、JSONそのものだけを出力してください。Markdownのバッククォートなどは含めないでください。

【構成案 JSONフォーマット】
{
  "career_summary": "キャリアサマリの内容（3-5文）",
  "job_history": [
    {
      "company_name": "会社名",
      "period": "20XX年X月～現在",
      "employment_type": "正社員",
      "position": "役職",
      "business_content": "事業内容の説明",
      "sales_amount": "売上高",
      "employee_count": "従業員数",
      "department": "担当部署",
      "details": [
        {
          "duration": "20XX年X月～現在",
          "content": "【営業スタイル】新規●%、既存●%\\n【担当地域】●●\\n【取引顧客】●●業界\\n【取引商品】●●\\n【担当業務】\\n・業務内容1\\n・業務内容2\\n\\n■実績\\n・20XX年度：目標達成率●%\\n\\n■ポイント\\n・エピソードタイトル\\n課題→仮説→取り組み→結果の流れで具体的エピソードを記載してください。"
        }
      ]
    }
  ],
  "skills": [
    {
      "title": "スキルタイトル",
      "description": "具体的な説明"
    }
  ]
}

【面談メモ】
${memo}

※情報が不足している部分は[要確認]としてください。
※丁寧で説得力のある日本語で作成してください。`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);

  if (json.candidates && json.candidates[0] && json.candidates[0].content) {
    let resultText = json.candidates[0].content.parts[0].text;
    // JSON部分のみを抽出（Markdownタグを削除するための正規表現）
    const jsonMatch = resultText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
    throw new Error('Geminiの出力からJSONを解析できませんでした。');
  } else {
    throw new Error('Gemini APIからのエラー応答を受け取りました: ' + responseText);
  }
}

/**
 * デバッグ用：現在使用可能なモデル一覧をログとアラートに出力する
 */
function debugListModels() {
  const apiKey = CONFIG.GEMINI_API_KEY;
  if (apiKey === 'YOUR_API_KEY_HERE' || apiKey === 'YOUR_GEMINI_API_KEY') {
    SpreadsheetApp.getUi().alert('APIキーが設定されていません。');
    return;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    
    if (json.models) {
      const modelNames = json.models.map(m => m.name.replace('models/', '')).join('\n');
      Logger.log('使用可能なモデル一覧:\n' + modelNames);
      SpreadsheetApp.getUi().alert('使用可能なモデル一覧:\n\n' + modelNames + '\n\nこの中にある名前をCONFIG.GEMINI_MODELに設定してください。');
    } else {
      SpreadsheetApp.getUi().alert('モデル一覧を取得できませんでした。応答: ' + response.getContentText());
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + e.toString());
  }
}

function createResumeDocument(name, date, data) {
  const folder = getOrCreateFolder(CONFIG.OUTPUT_FOLDER_NAME);
  const docName = `職務経歴書_${name}_${date.replace(/[/:\s]/g, '')}`;
  const doc = DocumentApp.create(docName);
  const docId = doc.getId();
  const file = DriveApp.getFileById(docId);
  
  // 指定のフォルダに移動
  folder.addFile(file);
  const parentFolders = file.getParents();
  while (parentFolders.hasNext()) {
    const parent = parentFolders.next();
    if (parent.getId() !== folder.getId()) {
      parent.removeFile(file);
    }
  }
  
  const body = doc.getBody();
  body.clear();
  
  // タイトル
  const titlePara = body.appendParagraph("職務経歴書");
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph(Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年MM月" + "×" + "日現在")).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph("氏名： " + name).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  // キャリアサマリ
  addSectionHeader(body, "キャリアサマリ");
  body.appendParagraph(data.career_summary);
  body.appendParagraph("");

  // 職務経歴
  addSectionHeader(body, "職務経歴");
  
  data.job_history.forEach((job, index) => {
    body.appendParagraph(`■職務経歴${index + 1}`).setBold(true);
    const jobLine = body.appendParagraph(`${job.period}  ${job.company_name}（${job.employment_type}）  役職：${job.position}`);
    jobLine.setBold(true).setForegroundColor('#0000FF').setUnderline(true);
    
    body.appendParagraph("■事業内容").setBold(true);
    body.appendParagraph(job.business_content);
    body.appendParagraph(`【売上高】 ${job.sales_amount}`);
    body.appendParagraph(`【従業員数】 ${job.employee_count}`);
    body.appendParagraph(`【担当部署】 ${job.department}`);
    body.appendParagraph("");
    
    // テーブル作成
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell("期間").setBackgroundColor('#F3F3F3').setWidth(100);
    headerRow.appendTableCell("主な職務内容").setBackgroundColor('#F3F3F3');
    
    job.details.forEach(detail => {
      const row = table.appendTableRow();
      row.appendTableCell(detail.duration);
      const contentCell = row.appendTableCell();
      detail.content.split('\n').forEach(line => {
        const p = contentCell.appendParagraph(line);
        if (line.startsWith('■') || line.startsWith('【')) {
            p.setBold(true);
        }
      });
    });
    body.appendParagraph("");
  });

  // 活かせる経験・知識・スキル
  addSectionHeader(body, "活かせる経験・知識・スキル");
  data.skills.forEach(skill => {
    body.appendParagraph(`【${skill.title}】`).setBold(true);
    body.appendParagraph(skill.description);
    body.appendParagraph("");
  });

  body.appendParagraph("以上").setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  doc.saveAndClose();
  return doc.getUrl();
}

function addSectionHeader(body, text) {
  const p = body.appendParagraph(text);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p.setBold(true);
  p.setAttributes({
    [DocumentApp.Attribute.SPACING_BEFORE]: 12,
    [DocumentApp.Attribute.SPACING_AFTER]: 6
  });
  body.appendHorizontalRule();
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}
