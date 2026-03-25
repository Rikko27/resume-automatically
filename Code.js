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
  
  // 検索対象のカレンダーの絞り込みキーワード (例: '佐木川')
  // 空（""）にすると全カレンダー（マイカレンダー、共有カレンダー等）を検索します
  CALENDAR_NAME_FILTER: '',
  
  // 取得を開始する特定の日付（例: '2026-03-01'）
  // 指定しない（null）場合は、今日を起点にします
  FETCH_START_DATE: null,
  
  // 今日から何日前までのイベントを取得するか
  DAYS_BACK: 0,
  
  // 今日から何日後までのイベントを取得するか
  DAYS_FORWARD: 14,
  
  // 抽出対象のイベントを特定するキーワード (タイトルまたは説明文に含まれるもの)
  SEARCH_KEYWORD: '初回面談',
  
  // 生成された職務経歴書を保存するフォルダ名
  OUTPUT_FOLDER_NAME: '生成済み職務経歴書',
  
  // スプレッドシートのシート名
  SHEET_NAME: '面談履歴'
};

// ==========================================
// メイン関数 (Main Functions)
// ==========================================

/**
 * エラーを回避しつつアラートを表示する（スプレッドシート用）
 */
function safeAlert(msg) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) ui.alert(msg);
  } catch (e) {
    Logger.log('Alert (skip UI): ' + msg);
  }
}

/**
 * スプレッドシートを開いた時に実行される
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📑 職務経歴書ツール')
      .addItem('1. カレンダーから面談を取得', 'fetchRecentInterviews')
      .addItem('2. 選択した行の職務経歴書を生成', 'generateResumeFromSelectedRow')
      .addSeparator()
      .addSubMenu(ui.createMenu('⏰ 自動実行設定')
        .addItem('▶︎ 毎朝10時の自動同期をONにする', 'setupDailyTrigger')
        .addItem('■ 自動同期をOFFにする', 'removeDailyTrigger')
        .addItem('🔍 設定状況を確認する', 'checkTriggerStatus'))
      .addSeparator()
      .addItem('🔑 APIキーの初期設定/再設定', 'setupApiKey')
      .addItem('🔍 診断：カレンダー取得の確認', 'diagnoseCalendar')
      .addItem('APIモデル一覧の確認（デバッグ用）', 'debugListModels')
      .addToUi();
  } catch (e) {
    // スプレッドシートUIがない場合はスキップ
  }
}

/**
 * カレンダーから最近の面談イベントを取得し、スプレッドシートに書き出す
 * @return {Object} 取得結果（count: 取得数, message: メッセージ）
 */
function fetchRecentInterviews() {
  try {
    const sheet = getOrCreateSheet();
    const now = new Date();
    
    let startTime;
    if (CONFIG.FETCH_START_DATE) {
      startTime = new Date(CONFIG.FETCH_START_DATE);
      startTime.setHours(0, 0, 0, 0);
    } else {
      startTime = new Date(now.getTime() - (CONFIG.DAYS_BACK * 24 * 60 * 60 * 1000));
    }
    
    // 取得終了時間を現在（または未来）に設定
    const endTime = new Date(now.getTime() + (CONFIG.DAYS_FORWARD * 24 * 60 * 60 * 1000));
    
    // ★ 実行者のプライマリカレンダーのみを取得対象にする
    const calendars = [CalendarApp.getDefaultCalendar()];
    let interviewEvents = [];
    
    calendars.forEach(cal => {
      const name = cal.getName();
      
      try {
        const events = cal.getEvents(startTime, endTime);
        const filtered = events.filter(e => {
          const title = e.getTitle();
          const desc = e.getDescription();
          return title.includes(CONFIG.SEARCH_KEYWORD) || desc.includes(CONFIG.SEARCH_KEYWORD);
        });
        
        if (filtered.length > 0) {
          Logger.log('カレンダー「' + name + '」から ' + filtered.length + ' 件見つかりました。');
          interviewEvents = interviewEvents.concat(filtered);
        }
      } catch (e) {
        Logger.log('スキップ: ' + name);
      }
    });

    // 重複を排除 (複数のカレンダーにまたがる場合)
    const seenIds = new Set();
    const targetEvents = [];
    interviewEvents.forEach(e => {
      const id = e.getId();
      if (!seenIds.has(id)) {
        seenIds.add(id);
        targetEvents.push(e);
      }
    });

    if (targetEvents.length === 0) {
      const msg = '対象のイベントが見つかりませんでした。すべてのカレンダーをスキャン済です。';
      safeAlert(msg);
      return { count: 0, message: msg };
    }
    
    // ★ 一括書き込み用にデータを収集する
    const rowsToAdd = [];
    const lastRow = sheet.getLastRow();
    
    // 重複チェック用の既存タイトル取得 (C列: イベント内容)
    const existingTitles = lastRow > 1 ? sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat() : [];
    
    // 実際にデータがある最後の行を探す（B列：日程 を基準にする）
    const bValues = lastRow > 0 ? sheet.getRange(1, 2, lastRow, 1).getValues().flat() : [];
    let actualLastRow = 1;
    for (let i = bValues.length - 1; i >= 0; i--) {
      if (bValues[i] !== '') {
        actualLastRow = i + 1;
        break;
      }
    }

    targetEvents.forEach(event => {
      const title = event.getTitle();
      const startTime = event.getStartTime();
      // タイトルと開始時間の組み合わせでより厳密に重複チェック
      const dateKey = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      
      if (!existingTitles.includes(title)) {
        // HTMLタグの除去と改行の整理（見栄えの改善とAIへの入力適正化）
        const rawDesc = event.getDescription() || '';
        let cleanDesc = rawDesc
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<\/p>/gi, '\n')
          .replace(/<[^>]+>/g, '')         // HTMLタグを除去
          .replace(/&nbsp;/gi, ' ')        // 特殊文字をスペースに
          .trim();
        
        // DriveリンクがHTMLタグ(href内)に隠れていた場合も考慮し、生テキストからIDを抽出して再付与する
        let driveIds = extractDriveFileIds(rawDesc);
        
        // Advanced Calendar Serviceを用いて、UIから直接添付されたファイル(クリップマーク)のIDも取得する
        try {
          const apiEventId = event.getId().replace('@google.com', '');
          // 対象カレンダーは実行者のプライマリカレンダーとしているため 'primary' を指定
          const advancedEvent = Calendar.Events.get('primary', apiEventId);
          if (advancedEvent && advancedEvent.attachments) {
            advancedEvent.attachments.forEach(att => {
              if (att.fileId && !driveIds.includes(att.fileId)) {
                driveIds.push(att.fileId);
              }
            });
          }
        } catch (err) {
          safeAlert('【エラー発覚①】カレンダーのクリップマーク取得で失敗しました：\n' + err.toString());
        }

        if (driveIds.length > 0) {
          cleanDesc += '\n\n【添付ファイルリンク】\n' + driveIds.map(id => `https://drive.google.com/file/d/${id}/view`).join('\n');
        }
        
        const candidateName = extractCandidateName(title, cleanDesc) || '不明';
        
        rowsToAdd.push([
          false, // 選択用
          dateKey,
          title,
          candidateName,
          '未作成',
          '', // リンク
          cleanDesc
        ]);
      }
    });
    
    if (rowsToAdd.length > 0) {
      // まとめて書き込み (高速)
      sheet.getRange(actualLastRow + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
      
      // チェックボックスの適用
      const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      sheet.getRange(actualLastRow + 1, 1, rowsToAdd.length, 1).setDataValidation(rule);
      
      const successMsg = rowsToAdd.length + ' 件の新しい面談を取得・保存しました。';
      Logger.log('【成功】' + successMsg);
      safeAlert(successMsg);
      return { count: rowsToAdd.length, message: successMsg };
    } else {
      const noNewMsg = '新しい面談は見つかりませんでした（既に保存済みか、対象外です）。';
      Logger.log('【終了】' + noNewMsg);
      safeAlert(noNewMsg);
      return { count: 0, message: noNewMsg };
    }
  } catch (e) {
    const errorMsg = '重大なエラー: ' + e.toString();
    Logger.log(errorMsg);
    safeAlert(errorMsg);
    return { count: 0, message: errorMsg };
  }
}

/**
 * 選択された行のデータを使って職務経歴書を生成する
 */
function generateResumeFromSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    safeAlert('「' + CONFIG.SHEET_NAME + '」シートが見つかりません。先に「カレンダーから面談を取得」を実行してください。');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    safeAlert('シートにデータがありません。先に「カレンダーから面談を取得」を実行してください。');
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
        safeAlert('「' + candidateName + '」様の生成中にエラーが発生しました：\n' + e.toString());
      }
    }
  }
  
  if (processedCount > 0) {
    safeAlert(processedCount + ' 件の職務経歴書を生成しました。');
  } else if (selectedCount === 0) {
    safeAlert('左端の「選択」欄にチェックが入っている行がありません。\n作成したい人の列にあるチェックボックスをクリックして、✔︎を入れてから実行してください。');
  } else {
    safeAlert('選択された行はすべて「完了」になっています。新しく作成する場合は、ステータスを消すか、別の行を選択してください。');
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
  }
  
  // シートが作成・取得されるたびに、確実にG列(メモ列)を非表示にする
  sheet.hideColumns(7);
  
  // 常にA2以降のデータがある範囲（シートの最大行まで）をチェックボックス形式にする
  const lastRow = sheet.getMaxRows();
  if (lastRow >= 2) {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 1, lastRow - 1, 1).setDataValidation(rule);
  }
  
  return sheet;
}

/**
 * 説明文またはタイトルから候補者名を抽出する
 */
function extractCandidateName(title, description) {
  // 1. タイトルから探す (最初に出てきた名前を優先)
  if (title) {
    // 最初に出てくる、空白、x、:、：、までを名前として抽出
    const firstMatch = title.trim().match(/^([^\s：:x]{2,10})/i);
    if (firstMatch) {
      const name = firstMatch[1].trim();
      // もし抽出されたのが「初回面談」などのキーワードなら次へ
      if (!name.includes('初回面談') && !name.includes('面談')) {
        return name;
      }
    }
    
    // 他のパターン (念のため)
    const bracketMatch = title.match(/】([^／\s：:]+)/);
    if (bracketMatch) return bracketMatch[1];
  }

  // 2. 説明文から探す
  if (!description) return null;
  const regex = /(?:氏名|名前|候補者名)\s*[:：]\s*([^\n\r]+)/i;
  const match = description.match(regex);
  return match ? match[1].trim() : null;
}

/**
 * Google DriveのファイルIDを抽出する
 */
function extractDriveFileIds(text) {
  if (!text) return [];
  const ids = new Set();
  const regex1 = /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/g;
  const regex2 = /drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/g;
  
  let match;
  while ((match = regex1.exec(text)) !== null) { ids.add(match[1]); }
  while ((match = regex2.exec(text)) !== null) { ids.add(match[1]); }
  
  return Array.from(ids);
}

/**
 * メモ欄から情報を抽出する。
 * カレンダーのメモ（説明文）の内容だけを直接AIへ渡します。
 */
function getFullContentFromMemo(memo) {
  if (!memo) return '';
  
  return '【面談メモ】\n' + memo + '\n\n';
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

  const prompt = `あなたは「営業職の職務経歴書」を採用担当が一目でレベルが高いと評価する水準に整理・生成するプロのキャリアライターです。
候補者情報をもとに、①職務経歴書（職務要約＋各社ブロック）と、②自己PR欄に転用できる再現性エピソード（skills_listへ格納）を作成してください。

【最重要方針（品質担保）】
1. 捏造の絶対禁止と【※要確認】の徹底：提供された情報（メモや添付PDF等）に書かれている事実・経歴・数値のみを使用してください。そこに書かれていない実績やエピソードを勝手に創作・追加することは一切禁じます。文章をプロらしく肉付け・整理するのは構いませんが、事実関係が読み取れない・不足している部分は勝手に推測せず、必ず「【※要確認】」としてください。
2. 会社概要の補完：会社名が提示されている場合は、「company_name」にそのまま出力し、絶対に「要確認」で省略・上書きしないこと。その上で事業内容/規模等を一般的なビジネス知識で補完し、自信がない補完部分にのみ【※要確認】とする。
3. 職務要約：全社歴をひとつに統合して作成し、各社ブロックには繰り返さないこと。
4. 自己PR（skills_listへの格納）：提供情報から「課題の発見 → 仮説立て → 実行 → 結果」の順で構成した自然文（150〜250文字）を作成すること。箇条書き・感情語・自己評価は禁止。冒頭に必ず【経営層アプローチスキル】【課題発見・仮説設計スキル】などのスキル名を付けること。

出力は必ず以下のJSONフォーマットに厳密に従い、JSONそのものだけを出力してください。Markdownのバッククォートなどは含めないでください。

【JSONフォーマット】
{
  "candidate_name": "候補者のフルネーム（「様」などの敬称やカッコは除外）",
  "career_summary": "キャリアサマリの内容（全社歴を統合。情報不足は【※要確認】）",
  "job_history": [
    {
      "company_name": "会社名（提示された社名をそのまま出力。「要確認」として消さないこと）",
      "period": "期間（例：2020年4月～2023年3月 / 不明は【※要確認】）",
      "employment_type": "雇用形態",
      "position": "役職【※要確認】",
      "capital": "資本金【※要確認】",
      "employees": "従業員数【※要確認】",
      "department": "担当部署名【※要確認】",
      "business_content": "事業内容の簡潔な説明（補完不可なら【※要確認】）",
      "overview": "【業務内容】にあたる概要（誰の/何の課題を、何を用いて、どう解決したか）",
      "tasks": ["具体的なタスク（箇条書き。リード獲得/アプローチ/提案/商談/クロージング等）"],
      "achievements": ["定量的な成果、KPI、成功事例（箇条書き。事実のみ、不明は【※要確認】）"],
      "points": "工夫した点（業務改善やチーム貢献など、2-3行の文章）"
    }
  ],
  "skills_list": [
    "【（スキル名）】（課題→仮説→実行→結果のエピソード自然文）",
    "【（スキル名）】（課題→仮説→実行→結果のエピソード自然文）"
  ]
}

【最重要ルール】
- どの項目（キャリアサマリ、職務内容、スキル等）においても、太字（**）や装飾マーカーは一切使用しないでください。ドキュメントの装飾はシステム側で行うため、プレーンなテキストのみを出力してください。

【面談メモ/文字起こし】
${memo}`;

  const parts = [];
  const driveIds = extractDriveFileIds(memo);
  
  if (driveIds.length > 0) {
    driveIds.forEach(id => {
      try {
        const file = DriveApp.getFileById(id);
        const mimeType = file.getMimeType();
        
        if (mimeType === 'application/pdf') {
          const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
          parts.push({
            inline_data: { mime_type: 'application/pdf', data: base64Data }
          });
        } else if (mimeType.startsWith('image/')) {
          const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
          parts.push({
            inline_data: { mime_type: mimeType, data: base64Data }
          });
        } else if (mimeType === MimeType.GOOGLE_DOCS) {
          const text = DocumentApp.openById(id).getBody().getText();
          parts.push({ text: `【添付資料: ${file.getName()} の内容】\n${text}\n\n` });
        } else if (mimeType === MimeType.PLAIN_TEXT || mimeType === 'text/plain' || mimeType === 'text/csv') {
          const text = file.getBlob().getDataAsString('UTF-8');
          parts.push({ text: `【添付資料: ${file.getName()} の内容】\n${text}\n\n` });
        }
      } catch (e) {
        safeAlert('【エラー発覚②】PDFの中身を読み込もうとして失敗しました：\n' + e.toString());
      }
    });
  }
  
  parts.push({ text: prompt });

  const payload = {
    contents: [{
      parts: parts
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
    safeAlert('APIキーが設定されていません。');
    return;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    
    if (json.models) {
      const modelNames = json.models.map(m => m.name.replace('models/', '')).join('\n');
      Logger.log('使用可能なモデル一覧:\n' + modelNames);
      safeAlert('使用可能なモデル一覧:\n\n' + modelNames + '\n\nこの中にある名前をCONFIG.GEMINI_MODELに設定してください。');
    } else {
      safeAlert('モデル一覧を取得できませんでした。応答: ' + response.getContentText());
    }
  } catch (e) {
    safeAlert('エラーが発生しました: ' + e.toString());
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
  const cleanName = name.replace(/[「」]/g, '').replace(/(?:さま|様|さん|氏)$/, '').trim();
  body.appendParagraph("氏名： " + cleanName + " 様").setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  // 1. キャリアサマリ
  addSectionHeader(body, "■ キャリアサマリ");
  body.appendParagraph(data.career_summary || '').setBold(false);
  body.appendParagraph("");

  // 2. 職務経歴
  addSectionHeader(body, '■ 職務経歴');
  const numbers = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
  
  (data.job_history || []).forEach((job, index) => {
    const num = numbers[index] || (index + 1);
    
    // 1. ■職務経歴① 会社名 (太字)
    body.appendParagraph(`■職務経歴${num} ${job.company_name}`).setBold(true).setFontSize(11);
    
    // 2. 期間・会社名・雇用形態・役職 (太字)
    const headerLine = `${job.period}（${job.employment_type || '正社員'}）${job.company_name} 役職：${job.position || '－'}`;
    body.appendParagraph(headerLine).setBold(true).setFontSize(11);
    
    // 3. ■事業内容 【資本金】... (標準)
    const bizInfo = `■事業内容 【資本金】${job.capital || '－'} 【従業員数】${job.employees || '－'}`;
    body.appendParagraph(bizInfo).setBold(false).setFontSize(11);
    
    // 4. 【担当部署】... (標準)
    body.appendParagraph(`【担当部署】${job.department || '－'}`).setBold(false).setFontSize(11);
    
    // 2列のテーブルを作成 (標準)
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
      const para = rightCell.appendParagraph('・');
      appendFormattedText(para, task);
    });
    rightCell.appendParagraph(''); // 空行
    
    // 3. ■実績
    rightCell.appendParagraph('■実績').setBold(false);
    (job.achievements || []).forEach(ach => {
      const para = rightCell.appendParagraph('・');
      appendFormattedText(para, ach);
    });
    rightCell.appendParagraph(''); // 空行
    
    // 4. ■ポイント
    rightCell.appendParagraph('■ポイント').setBold(false);
    const pointPara = rightCell.appendParagraph('');
    appendFormattedText(pointPara, job.points || '');
    pointPara.setBold(false); // 項目名も含め太字にしない
    
    body.appendParagraph(""); // スペース
  });

  // 3. 活かせる経験・知識・スキル
  addSectionHeader(body, '■ 活かせる経験・知識・スキル');
  (data.skills_list || []).forEach(skill => {
    const para = body.appendParagraph('・');
    appendFormattedText(para, skill);
    para.setBold(false); // 項目名も含め太字にしない
  });

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
function setupApiKey() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Gemini APIキーの登録', 
      'ご自身のGemini APIキー（AIza...で始まる文字列）を貼り付けてOKを押してください。\n' +
      '※この操作により、GitHubに公開したコードを変更せずに安全にキーを保存できます。', 
      ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
      const key = response.getResponseText().trim();
      if (key) {
        PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
        safeAlert('APIキーを正常に保存しました！これで職務経歴書の生成が可能になります。');
      } else {
        safeAlert('キーが入力されなかったため、保存を中止しました。');
      }
    }
  } catch (e) {
    Logger.log('UI not available for setupApiKey');
  }
}

/**
 * デバッグ用：現在使用可能なカレンダーを一覧表示する
 */
function diagnoseCalendar() {
  const cal = CalendarApp.getDefaultCalendar();
  let msg = "【カレンダー取得診断結果】\n\n";
  msg += "現在、取得対象となっているカレンダー:\n";
  msg += `- ${cal.getName()} (ID: ${cal.getId()}) [メイン]\n`;
  
  msg += `\n現在、タイトルまたは説明に「${CONFIG.SEARCH_KEYWORD}」が含まれる予定を検索しています。`;
  msg += "\n※現在は実行者のプライマリカレンダー（メイン）のみを検索する設定になっています。";
  
  safeAlert(msg);
}

// ==========================================
// 自動化設定 (Automation Settings)
// ==========================================

/**
 * 毎朝10時にカレンダー同期を実行するトリガーを設定する
 */
function setupDailyTrigger() {
  // 既存の同一関数のトリガーを削除（重複防止）
  removeDailyTrigger(true);
  
  // 毎日午前10時〜11時の間に実行するように設定
  ScriptApp.newTrigger('fetchRecentInterviews')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
    
  safeAlert('【設定完了】\n毎日午前10時頃に自動でカレンダーを確認し、新しい「初回面談」をシートに追加するようにしました。');
}

/**
 * 自動同期のトリガーを解除する
 */
function removeDailyTrigger(isSilent) {
  const triggers = ScriptApp.getProjectTriggers();
  let found = false;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'fetchRecentInterviews') {
      ScriptApp.deleteTrigger(triggers[i]);
      found = true;
    }
  }
  
  if (!isSilent) {
    if (found) {
      safeAlert('【設定解除】\n自動同期をOFFにしました。');
    } else {
      safeAlert('現在、自動同期は設定されていません。');
    }
  }
}

/**
 * 現在のオートメーション設定状況を確認する
 */
function checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  let isActive = false;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'fetchRecentInterviews') {
      isActive = true;
      break;
    }
  }
  
  if (isActive) {
    safeAlert('【現在の設定】\n✅ 自動同期：ON (毎日10時頃)\n\n毎朝自動的にカレンダーの「初回面談」を取得してスプレッドシートを更新しています。');
  } else {
    safeAlert('【現在の設定】\n× 自動同期：OFF\n\n現在は手動での同期のみ可能です。');
  }
}
