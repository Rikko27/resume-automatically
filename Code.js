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
  // ★ APIキーはスクリプトプロパティ「GEMINI_API_KEY」から動的に取得します
  get GEMINI_API_KEY() {
    return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  },
  
  // ★ 使用するGeminiのモデル名
  // ご自身のリストに存在し、かつエラーが出にくい「gemini-flash-latest」を指定します
  GEMINI_MODEL: 'gemini-flash-latest',
  
  // 検索対象のキーワード
  SEARCH_KEYWORD: '初回面談',
  
  // 取得を開始する特定の日付（例: '2026-03-01'）
  // 指定しない（null）場合は、DAYS_BACKが適用されます
  FETCH_START_DATE: null,
  
  // 何日前までのイベントを取得するか（FETCH_START_DATEがnullの場合のみ使用）
  DAYS_BACK: 21,
  
  // 生成された職務経歴書を保存するフォルダ名
  OUTPUT_FOLDER_NAME: '生成済み職務経歴書',
  
  // スプレッドシートのシート名
  SHEET_NAME: '面談履歴'
};

// ==========================================
// メイン関数 (Main Functions)
// ==========================================

/**
 * スプレッドシートを開いた時に実行される
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📑 職務経歴書ツール')
      .addItem('1. カレンダーから面談を取得', 'fetchRecentInterviews')
      .addItem('2. 選択した行の職務経歴書を生成', 'generateResumeFromSelectedRow')
      .addSeparator()
      .addItem('🔑 APIキーの初期設定/再設定', 'setupApiKey')
      .addItem('🔍 診断：カレンダー取得の確認', 'diagnoseCalendar')
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
  
  // カレンダーの取得
  let targetCalendar = CalendarApp.getCalendarsByName(CONFIG.SEARCH_KEYWORD)[0];
  if (!targetCalendar) {
    targetCalendar = CalendarApp.getDefaultCalendar();
    Logger.log('「' + CONFIG.SEARCH_KEYWORD + '」という名前のカレンダーは見つからなかったため、メインのカレンダーを使用します。');
  } else {
    Logger.log('「' + CONFIG.SEARCH_KEYWORD + '」という名前のカレンダーから読み込みます。');
  }

  const events = targetCalendar.getEvents(startTime, now);
  Logger.log('期間内に ' + events.length + ' 件のイベントが見つかりました。');
  
  // フィルタリング：タイトルまたは説明にキーワードが含まれる
  const interviewEvents = events.filter(e => {
    const title = e.getTitle();
    const desc = e.getDescription();
    return title.includes(CONFIG.SEARCH_KEYWORD) || desc.includes(CONFIG.SEARCH_KEYWORD);
  });
  
  if (interviewEvents.length === 0) {
    SpreadsheetApp.getUi().alert('対象のイベント（' + CONFIG.SEARCH_KEYWORD + '）が見つかりませんでした。期間：' + Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd') + ' ～ ' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd'));
    return;
  }
  
  // 既存のデータを取得（重複チェック用）
  const lastRow = sheet.getLastRow();
  const existingTitles = lastRow > 1 
    ? sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat() 
    : [];
  
  // 実際にデータが入っている最後の行を探す（チェックボックス列を除外して判定）
  let actualLastRow = 1;
  if (lastRow > 1) {
    const dataRange = sheet.getRange(1, 2, lastRow, 1).getValues();
    for (let i = dataRange.length - 1; i >= 0; i--) {
      if (dataRange[i][0] !== "") {
        actualLastRow = i + 1;
        break;
      }
    }
  }

  interviewEvents.forEach(event => {
    const title = event.getTitle();
    if (!existingTitles.includes(title)) {
      const date = Utilities.formatDate(event.getStartTime(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const description = event.getDescription();
      const candidateName = extractCandidateName(description) || '不明';
      
      actualLastRow++;
      sheet.getRange(actualLastRow, 1, 1, 7).setValues([[
        false, // 選択用
        date,
        title,
        candidateName,
        '未作成',
        '', // リンク
        description
      ]]);
    }
  });
  
  // チェックボックスをデータがある行だけに適用
  if (actualLastRow >= 2) {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 1, actualLastRow - 1, 1).setDataValidation(rule);
    
    // データがない下の行のチェックボックス（もしあれば）を削除
    const totalMaxRows = sheet.getMaxRows();
    if (totalMaxRows > actualLastRow) {
      sheet.getRange(actualLastRow + 1, 1, totalMaxRows - actualLastRow, 1).clearDataValidations().clearContent();
    }
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
        // メモおよびリンク先ドキュメントから全内容を取得
        const fullContent = getFullContentFromMemo(memo);
        Logger.log('抽出されたコンテキストの内容: ' + fullContent.substring(0, 500) + '...');
        
        const resumeData = callGeminiAPI(fullContent);
        
        // AIが特定した名前があれば、それを使う（「不明」だった場合の補完）
        let finalName = candidateName;
        if ((!candidateName || candidateName === '不明') && resumeData.candidate_name) {
          finalName = resumeData.candidate_name;
          sheet.getRange(i + 1, 4).setValue(finalName); // シートの名前列を更新
        }
        
        const docUrl = createResumeDocument(finalName, String(dateStr), resumeData);
        Logger.log('生成されたドキュメントURL: ' + docUrl);
        
        // 列インデックスを動的に取得して確実に書き込む
        const headers = data[0];
        const statusCol = headers.indexOf('ステータス') + 1 || 5;
        const urlCol = headers.indexOf('URL') + 1 || 6;
        const checkboxCol = headers.indexOf('選択') + 1 || 1;
        
        Logger.log(`書き込み先: ステータス=${statusCol}, URL=${urlCol}`);
        
        // シートを更新
        sheet.getRange(i + 1, statusCol).setValue('完了');
        // ハイパーリンク形式で確実に表示させる
        sheet.getRange(i + 1, urlCol).setFormula(`=HYPERLINK("${docUrl}", "開く")`);
        sheet.getRange(i + 1, checkboxCol).setValue(false);
        
        SpreadsheetApp.flush(); // 即座に画面へ反映
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
  // 複数のパターンに対応 (氏名, 名前, 候補者名 など)
  const regex = /(?:氏名|名前|候補者名)\s*[:：]\s*([^\n\r]+)/i;
  const match = description.match(regex);
  return match ? match[1].trim() : null;
}

/**
 * メモ欄から情報を抽出する。Googleドキュメントのリンクがあれば、その中身も取得する。
 */
function getFullContentFromMemo(memo) {
  if (!memo) return '';
  
  let fullContent = '【面談メモ】\n' + memo + '\n\n';
  
  // GoogleドキュメントのURLを抽出 (docs.google.com/document/d/...)
  const docUrlMatch = memo.match(/https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
  
  if (docUrlMatch) {
    const docId = docUrlMatch[1];
    try {
      const doc = DocumentApp.openById(docId);
      const docText = doc.getBody().getText();
      fullContent += '【提供されたドキュメント（文字起こし等）の内容】\n' + docText;
      Logger.log('ドキュメントから内容を取得しました: ' + doc.getName());
    } catch (e) {
      Logger.log('ドキュメントの取得に失敗しました: ' + e.toString());
      fullContent += '\n※注意: リンク先のドキュメントを読み取れませんでした。権限を確認してください。';
    }
  }
  
  return fullContent;
}

function callGeminiAPI(memo) {
  // 直接スクリプトプロパティから取得（Getterの不具合を回避）
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey || apiKey === 'YOUR_API_KEY_HERE' || apiKey === 'YOUR_GEMINI_API_KEY' || apiKey.trim() === '') {
    throw new Error('GeminiのAPIキーが正しく設定されていません。現在の設定内容：' + (apiKey ? '設定済み(一部：' + apiKey.substring(0, 4) + '...)' : '未設定(null)') + 
      '\n\nGoogle Apps Scriptエディタの左メニュー「設定（歯車マーク）」＞「スクリプトプロパティ」に、\nプロパティ名: GEMINI_API_KEY\n値: (あなたのAPIキー)\nとして登録されているか、スペルミスがないか再度ご確認ください。');
  }

  // 互換性の高い v1beta エンドポイントを使用
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const prompt = `あなたは「通過率が上がる職務経歴書」を作成する超一流のキャリアアドバイザーです。
以下の【厳守ルール】に従い、提供された情報から最高品質の職務経歴書を作成してください。

【厳守ルール】
1. 構成は以下の3項目のみとし、他は一切追加しないでください：
   - キャリアサマリ
   - 職務経歴
   - 活かせる経験・知識・スキル

2. 各項目の記述ルール：
   - キャリアサマリ：3-5文で簡潔かつ強力に。
   - 職務経歴：各社ごとに「企業名」「期間」「事業内容」「職務内容」「実績（数値化）」「ポイント」を記述。
     ※「ポイント」は文章で「なぜ成果が出たか」をエピソード付きで具体的に。
   - 活かせる経験・知識・スキル：箇条書きは絶対に禁止。文章でストーリーとして記述。事実に基づきつつも、プロフェッショナルな表現で候補者の魅力を最大限に引き出してください。

3. スタンス：
   - 単なる情報の整理ではなく、採用担当者が「会いたい」と思う魅力的なアピールに仕上げること。
   - 日付は自動的に本日（${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd")}）として扱ってください。

4. 正確性と表現力の高度なバランス：
   - 「数値、期間、社名」といった具体的事実は、絶対に捏造しないでください。不明な場合は「[要確認]」とします。
   - 一方で、具体的なエピソードや職務内容については、提供された事実の断片から、プロのキャリアアドバイザーとしてふさわしい「専門的な言い回し」や「深掘りした記述」に膨らませてください。
   - 採用担当者が一目で「レベルが高い」と感じるような、洗練されたビジネス日本語を使用してください。

出力は必ず以下のJSONフォーマットに厳密に従い、JSONそのものだけを出力してください。Markdownのバッククォートなどは含めないでください。

【JSONフォーマット】
{
  "candidate_name": "候補者のフルネーム（見つからない場合はnull）",
  "career_summary": "キャリアサマリの内容",
  "job_history": [
    {
      "company_name": "会社名",
      "period": "期間（例：2020年4月～2023年3月）",
      "employment_type": "雇用形態（例：正社員）",
      "position": "役職",
      "overview": "【業務内容】にあたる概要（ミッションや役割を2-3行で）",
      "tasks": ["具体的なタスク（箇条書き用）"],
      "achievements": ["定量的な成果や表彰歴（箇条書き用）"],
      "points": "工夫した点や戦略（2-3行の文章）"
    }
  ],
  "skills_story": "活かせる経験・知識・スキルの内容（ストーリー形式）"
}

【面談メモ/文字起こし】
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
  const docName = `職務経歴書_${name}_${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd")}`;
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
  
  body.appendParagraph(Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年MM月dd日現在")).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph("氏名： " + name + " 様").setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  // 1. キャリアサマリ
  addSectionHeader(body, "■ キャリアサマリ");
  body.appendParagraph(data.career_summary || '');
  body.appendParagraph("");

  // 2. 職務経歴
  addSectionHeader(body, '■ 職務経歴');
  (data.job_history || []).forEach((job, index) => {
    // 職務経歴ヘッダー情報（太字なし）
    const headerText = `${job.period}  ${job.company_name}（${job.employment_type || '正社員'}）  役職：${job.position || '－'}`;
    body.appendParagraph(headerText).setBold(false).setFontSize(11);
    
    // 2列のテーブルを作成
    const table = body.appendTable();
    
    // 表のヘッダー（太字なし）
    const headRow = table.appendTableRow();
    headRow.appendTableCell('期間').setBold(false).setBackgroundColor('#F3F3F3').setWidth(100);
    headRow.appendTableCell('主な職務内容').setBold(false).setBackgroundColor('#F3F3F3');
    
    const contentRow = table.appendTableRow();
    
    // 左列：期間
    const leftCell = contentRow.appendTableCell(job.period || '');
    leftCell.setWidth(100);
    leftCell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
    
    // 右列：主な職務内容
    const rightCell = contentRow.appendTableCell();
    
    // 1. 【業務内容】
    rightCell.appendParagraph('【業務内容】').setBold(false);
    rightCell.appendParagraph(job.overview || '').setBold(false);
    rightCell.appendParagraph(''); // 空行
    
    // 2. 【担当業務】
    rightCell.appendParagraph('【担当業務】').setBold(false);
    (job.tasks || []).forEach(task => {
      const li = rightCell.appendListItem('');
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
      appendFormattedText(li, task);
    });
    rightCell.appendParagraph(''); // 空行
    
    // 3. ■実績
    rightCell.appendParagraph('■実績').setBold(false);
    (job.achievements || []).forEach(ach => {
      const li = rightCell.appendListItem('');
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
      appendFormattedText(li, ach);
    });
    rightCell.appendParagraph(''); // 空行
    
    // 4. ■ポイント
    rightCell.appendParagraph('■ポイント').setBold(false);
    const pointPara = rightCell.appendParagraph('');
    appendFormattedText(pointPara, job.points || '');
    
    body.appendParagraph(""); // スペース
  });

  // 3. 活かせる経験・知識・スキル
  addSectionHeader(body, "■ 活かせる経験・知識・スキル");
  // ストーリー形式（箇条書き禁止）
  body.appendParagraph(data.skills_story || '');

  body.appendParagraph("以上").setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  doc.saveAndClose();
  return doc.getUrl();
}

function addSectionHeader(body, text) {
  const p = body.appendParagraph(text);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p.setBold(true); // 大見出しのみ太字にする
  p.setAttributes({
    [DocumentApp.Attribute.SPACING_BEFORE]: 12,
    [DocumentApp.Attribute.SPACING_AFTER]: 6
  });
}

/**
 * 太字マーカー（**）を処理してパラグラフに追加するヘルパー
 */
function appendFormattedText(paragraph, text) {
  if (!text) return;
  const parts = text.split(/(\*\*.*?\*\*)/);
  parts.forEach(part => {
    if (!part) return; // 空の文字列はスキップ（Googleドキュメントのエラー回避）
    if (part.startsWith('**') && part.endsWith('**')) {
      paragraph.appendText(part.slice(2, -2)).setBold(true);
    } else {
      paragraph.appendText(part).setBold(false);
    }
  });
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * 画面上のプロンプトからAPIキーを安全にスクリプトプロパティに保存する
 */
function debugSetApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Gemini APIキーの登録', 
    'ご自身のGemini APIキー（AIza...で始まる文字列）を貼り付けてOKを押してください。\n' +
    '※この操作により、GitHubに公開したコードを変更せずに安全にキーを保存できます。', 
    ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const key = response.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
      ui.alert('APIキーを正常に保存しました！これで職務経歴書の生成が可能になります。');
    } else {
      ui.alert('キーが入力されなかったため、保存を中止しました。');
    }
  }
}

/**
 * デバッグ用：現在使用可能なカレンダーを一覧表示する
 */
function diagnoseCalendar() {
  const calendars = CalendarApp.getAllCalendars();
  let msg = "【カレンダー取得診断結果】\n\n";
  msg += "あなたの環境で利用可能なカレンダー:\n";
  calendars.forEach(cal => {
    msg += `- ${cal.getName()} (ID: ${cal.getId()})${cal.isPrimary() ? ' [メイン]' : ''}\n`;
  });
  
  msg += `\n現在、タイトルまたは説明に「${CONFIG.SEARCH_KEYWORD}」が含まれる予定を検索しています。`;
  msg += "\n※もし特定の名前のカレンダー（例：共有カレンダーなど）から取得したい場合は、そのカレンダー名を『初回面談』にするか、プログラムの設定を変更する必要があります。";
  
  SpreadsheetApp.getUi().alert(msg);
}
