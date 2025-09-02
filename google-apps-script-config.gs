// 管理用スプレッドシートのGoogle Apps Script
// 設定の読み込み・保存機能とフォームデータの送信機能

// 日時を日本時間のISO形式でフォーマット（Zなし）（2025-08-09T03:03:38.259）
function formatDateTimeForJapan(date) {
  // 日本時間に変換（UTC+9）
  const japanTime = new Date(date.getTime() + (9 * 60 * 60 * 1000));
  
  const year = japanTime.getUTCFullYear();
  const month = String(japanTime.getUTCMonth() + 1).padStart(2, '0');
  const day = String(japanTime.getUTCDate()).padStart(2, '0');
  const hours = String(japanTime.getUTCHours()).padStart(2, '0');
  const minutes = String(japanTime.getUTCMinutes()).padStart(2, '0');
  const seconds = String(japanTime.getUTCSeconds()).padStart(2, '0');
  const milliseconds = String(japanTime.getUTCMilliseconds()).padStart(3, '0');
  
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}.${milliseconds}`;
}

// 奨励会負担金の金額を設定から取得
function getEncouragementFeeAmount() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const burdenSheet = sheet.getSheetByName('負担金金額設定');
    
    if (!burdenSheet) {
      console.warn('負担金金額設定シートが見つかりません。デフォルト値1000を使用します。');
      return 1000;
    }
    
    const burdenData = burdenSheet.getDataRange().getValues();
    
    // encouragement_fee の行を探す
    for (let i = 1; i < burdenData.length; i++) {
      const row = burdenData[i];
      const id = row[0];
      const amount = parseInt(row[2]);
      const isActive = row[4];
      
      if (id === 'encouragement_fee' && isActive && !isNaN(amount)) {
        console.log('奨励会負担金を設定から取得:', amount);
        return amount;
      }
    }
    
    console.warn('奨励会負担金の設定が見つかりません。デフォルト値1000を使用します。');
    return 1000;
  } catch (error) {
    console.error('奨励会負担金取得エラー:', error);
    console.warn('デフォルト値1000を使用します。');
    return 1000;
  }
}

// 負担金の単価を設定から取得（id = 'burden_fee' の金額（C列））
function getBurdenFeeUnitAmount() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const burdenSheet = sheet.getSheetByName('負担金金額設定');
    
    if (!burdenSheet) {
      console.warn('負担金金額設定シートが見つかりません。デフォルト値40000を使用します。');
      return 40000;
    }
    
    const burdenData = burdenSheet.getDataRange().getValues();
    for (let i = 1; i < burdenData.length; i++) {
      const row = burdenData[i];
      const id = row[0];
      const raw = row[2];
      let amount = 0;
      if (typeof raw === 'number') {
        amount = raw;
      } else if (raw !== null && raw !== undefined) {
        amount = Number(String(raw).replace(/[^0-9.-]/g, '')) || 0;
      }
      const isActive = row[4];
      if (id === 'burden_fee' && isActive && !isNaN(amount)) {
        console.log('負担金単価を設定から取得:', amount);
        return amount;
      }
    }
    console.warn('負担金単価の設定が見つかりません。デフォルト値40000を使用します。');
    return 40000;
  } catch (error) {
    console.error('負担金単価取得エラー:', error);
    console.warn('デフォルト値40000を使用します。');
    return 40000;
  }
}

function doGet(e) {
  // CORS設定を追加
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  // パラメータが存在しない場合の処理
  if (!e || !e.parameter) {
    return output.setContent(JSON.stringify({
      success: false,
      error: 'No parameters provided'
    }));
  }
  
  const action = e.parameter.action;
  
  if (action === 'getSettings') {
    return getSettings();
  } else if (action === 'getFormSettings') {
    return getFormSettings();
  }
  
  return output.setContent(JSON.stringify({
    success: false,
    error: 'Invalid action'
  }));
}

function doPost(e) {
  // CORS設定を追加
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  // パラメータが存在しない場合の処理
  if (!e || !e.postData) {
    return output.setContent(JSON.stringify({
      success: false,
      error: 'No post data provided'
    }));
  }
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === 'saveSettings') {
      return saveSettings(data.settings);
    } else if (action === 'submitFormData') {
      return submitFormData(data.formData);
    }
    
    return output.setContent(JSON.stringify({
      success: false,
      error: 'Invalid action'
    }));
  } catch (error) {
    return output.setContent(JSON.stringify({
      success: false,
      error: 'Invalid JSON data: ' + error.toString()
    }));
  }
}

// フォーム設定からキーの値を取得（C列がTRUEの行のみ有効）
function getFormSettingValue(key) {
  try {
    console.log(`getFormSettingValue開始: キー="${key}"`);
    
    const active = SpreadsheetApp.getActiveSpreadsheet();
    console.log('ActiveSpreadsheet名:', active.getName());
    
    const formSettingsSheet = active.getSheetByName('フォーム設定');
    if (!formSettingsSheet) {
      console.log('フォーム設定シートが見つかりません');
      return '';
    }
    
    console.log('フォーム設定シートを発見');
    const values = formSettingsSheet.getDataRange().getValues();
    console.log('フォーム設定シートの行数:', values.length);
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const k = row[0];
      const v = row[1];
      const enabled = row[2];
      
      console.log(`行${i + 1}: キー="${k}", 値="${v}", 有効=${enabled}`);
      
      if (enabled && String(k).trim() === key && v) {
        const result = String(v).trim();
        console.log(`キー"${key}"の値を発見: "${result}"`);
        return result;
      }
    }
    
    console.log(`キー"${key}"の値が見つかりませんでした`);
    return '';
    
  } catch (error) {
    console.error(`getFormSettingValueエラー (キー: ${key}):`, error);
    return '';
  }
}

// URLまたはIDからスプレッドシートIDを抽出
function extractSpreadsheetId(idOrUrl) {
  if (!idOrUrl) return '';
  const raw = String(idOrUrl).trim();
  const m = raw.match(/\/d\/([a-zA-Z0-9-_]+)\//);
  if (m && m[1]) return m[1];
  return raw; // 既にIDならそのまま
}

// 集計用ブックを取得（設定未登録時はActiveSpreadsheetをフォールバック）
function getAggregationBook() {
  try {
    console.log('=== getAggregationBook開始 ===');
    
    // フォーム設定から集計用ブックのID/URLを取得
    const aggregationBookId = getFormSettingValue('aggregationBookId');
    const aggregationBookUrl = getFormSettingValue('aggregationBookUrl');
    
    console.log('フォーム設定から取得:');
    console.log('- aggregationBookId:', aggregationBookId);
    console.log('- aggregationBookUrl:', aggregationBookUrl);
    
    const idOrUrl = aggregationBookId || aggregationBookUrl;
    
    if (idOrUrl) {
      console.log('集計用ブックのID/URLを発見:', idOrUrl);
      const id = extractSpreadsheetId(idOrUrl);
      console.log('抽出されたID:', id);
      
      const book = SpreadsheetApp.openById(id);
      console.log('集計用ブックを開きました:', book.getName());
      console.log('=== getAggregationBook完了 ===');
      return book;
    } else {
      console.log('フォーム設定に集計用ブックのID/URLが設定されていません');
    }
  } catch (e) {
    console.error('集計用ブック取得エラー:', e);
    console.error('エラースタック:', e.stack);
  }
  
  console.warn('aggregationBookId/url が未設定のため、ActiveSpreadsheetを使用します');
  const fallbackBook = SpreadsheetApp.getActiveSpreadsheet();
  console.log('フォールバック: ActiveSpreadsheetを使用:', fallbackBook.getName());
  console.log('=== getAggregationBook完了（フォールバック） ===');
  return fallbackBook;
}

// フォーム設定を取得
function getFormSettings() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('フォーム設定取得 - スプレッドシート名:', sheet.getName());
    
    const formSettingsSheet = sheet.getSheetByName('フォーム設定');
    
    if (!formSettingsSheet) {
      console.log('フォーム設定シートが見つかりません。デフォルト設定を返します。');
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        formSettings: {
          formTitle: '奉加帳提出報告フォーム',
          formSubtitle: '〇〇大会振込報告フォーム'
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
    }
    
    // フォーム設定を読み込み
    const formData = formSettingsSheet.getDataRange().getValues();
    const formSettings = {};
    
    // ヘッダー行をスキップして2行目から処理
    for (let i = 1; i < formData.length; i++) {
      const row = formData[i];
      const settingKey = row[0];
      const settingValue = row[1];
      const isActive = row[2];
      
      if (isActive && settingKey && settingValue) {
        formSettings[settingKey] = settingValue;
      }
    }
    
    // デフォルト値を設定（設定が見つからない場合）
    if (!formSettings.formTitle) {
      formSettings.formTitle = '奉加帳提出報告フォーム';
    }
    if (!formSettings.formSubtitle) {
      formSettings.formSubtitle = '〇〇大会振込報告フォーム';
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      formSettings: formSettings
    }))
    .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('フォーム設定取得エラー:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 設定を取得
function getSettings() {
  try {
    // スプレッドシートIDを直接指定する場合（必要に応じて変更）
    // const sheet = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID');
    
    // 現在のスプレッドシートを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('スプレッドシート名:', sheet.getName());
    console.log('スプレッドシートID:', sheet.getId());
    
    // デバッグ用：利用可能なシート名をログに出力
    const sheetNames = sheet.getSheets().map(s => s.getName());
    console.log('利用可能なシート:', sheetNames);
    
    const giftSheet = sheet.getSheetByName('返礼品設定');
    const amountSheet = sheet.getSheetByName('寄付金金額設定');
    const burdenSheet = sheet.getSheetByName('負担金金額設定');
    
    // 各シートの存在確認とログ出力
    console.log('返礼品設定シート:', giftSheet ? '存在' : '不存在');
    console.log('寄付金金額設定シート:', amountSheet ? '存在' : '不存在');
    console.log('負担金金額設定シート:', burdenSheet ? '存在' : '不存在');
    
    if (!giftSheet || !amountSheet || !burdenSheet) {
      const missingSheets = [];
      if (!giftSheet) missingSheets.push('返礼品設定');
      if (!amountSheet) missingSheets.push('寄付金金額設定');
      if (!burdenSheet) missingSheets.push('負担金金額設定');
      
      throw new Error(`必要なシートが見つかりません: ${missingSheets.join(', ')}`);
    }
    
    // 返礼品設定を読み込み
    const giftData = giftSheet.getDataRange().getValues();
    const giftRules = {};
    
    // 新スキーマ（minAmount 列削除）: A:ID, B:名前, C:説明, D:有効
    for (let i = 1; i < giftData.length; i++) {
      const row = giftData[i];
      const id = row[0];
      const name = row[1];
      const description = row[2];
      const isActive = row[3];
      
      if (isActive && id && name) {
        giftRules[id] = {
          name: name,
          description: description
        };
      }
    }
    
    // 寄付金金額設定を読み込み
    const amountData = amountSheet.getDataRange().getValues();
    const amountOptions = {};
    
    for (let i = 1; i < amountData.length; i++) {
      const row = amountData[i];
      const id = row[0];
      const amount = parseInt(row[1]);
      const displayName = row[2];
      const isActive = row[3];
      
      if (isActive && id && amount) {
        amountOptions[id] = {
          amount: amount,
          displayName: displayName
        };
      }
    }
    
    // 負担金金額設定を読み込み
    const burdenData = burdenSheet.getDataRange().getValues();
    const burdenSettings = {};
    
    for (let i = 1; i < burdenData.length; i++) {
      const row = burdenData[i];
      const id = row[0];
      const name = row[1];
      const amount = parseInt(row[2]);
      const description = row[3];
      const isActive = row[4];
      
      if (isActive && id && name) {
        burdenSettings[id] = {
          name: name,
          amount: amount,
          description: description
        };
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      giftRules: giftRules,
      amountOptions: amountOptions,
      burdenSettings: burdenSettings
    }))
    .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 設定を保存
function saveSettings(settings) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 返礼品設定を保存
    if (settings.giftRules) {
      const giftSheet = sheet.getSheetByName('返礼品設定');
      if (giftSheet) {
        saveGiftRules(giftSheet, settings.giftRules);
      }
    }
    
    // 寄付金金額設定を保存
    if (settings.amountOptions) {
      const amountSheet = sheet.getSheetByName('寄付金金額設定');
      if (amountSheet) {
        saveAmountOptions(amountSheet, settings.amountOptions);
      }
    }
    
    // 負担金金額設定を保存
    if (settings.burdenSettings) {
      const burdenSheet = sheet.getSheetByName('負担金金額設定');
      if (burdenSheet) {
        saveBurdenSettings(burdenSheet, settings.burdenSettings);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: '設定を保存しました'
    }))
    .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}

// 返礼品設定を保存
function saveGiftRules(sheet, giftRules) {
  // 既存データをクリア（ヘッダー行は保持）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  
  // 新しいデータを書き込み（新スキーマ: A:ID, B:名前, C:説明, D:有効）
  const newData = [];
  Object.entries(giftRules).forEach(([id, rule]) => {
    newData.push([
      id,
      rule.name,
      rule.description || '',
      true // 有効
    ]);
  });
  
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, 4).setValues(newData);
  }
}

// 寄付金金額設定を保存
function saveAmountOptions(sheet, amountOptions) {
  // 既存データをクリア（ヘッダー行は保持）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  
  // 新しいデータを書き込み
  const newData = [];
  Object.entries(amountOptions).forEach(([id, option]) => {
    newData.push([
      id,
      option.amount,
      option.displayName,
      true // 有効
    ]);
  });
  
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, 4).setValues(newData);
  }
}

// 負担金金額設定を保存
function saveBurdenSettings(sheet, burdenSettings) {
  // 既存データをクリア（ヘッダー行は保持）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
  
  // 新しいデータを書き込み
  const newData = [];
  Object.entries(burdenSettings).forEach(([id, setting]) => {
    newData.push([
      id,
      setting.name,
      setting.amount,
      setting.description || '',
      true // 有効
    ]);
  });
  
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, 5).setValues(newData);
  }
}

// フォームデータを送信
function submitFormData(formData) {
  try {
    // デバッグ用：受信したデータの構造をログに出力
    console.log('受信したformData:', JSON.stringify(formData, null, 2));
    
    // 設定用ブック（Active）と集計用ブック（年次）を分離
    const aggregationBook = getAggregationBook();
    console.log('集計用ブック名:', aggregationBook.getName());
    console.log('集計用ブックID:', aggregationBook.getId());
    
    // 1ページ目データ整理シートにデータを追加（集計用ブック）
    const page1DataSheet = aggregationBook.getSheetByName('1ページ目データ整理');
    console.log('1ページ目データ整理シート(集計用):', page1DataSheet ? '存在' : '不存在');
    if (page1DataSheet) {
      addPage1Data(page1DataSheet, formData);
    }
    
    // 2ページ目データ整理シートにデータを追加（集計用ブック）
    const page2DataSheet = aggregationBook.getSheetByName('2ページ目データ整理');
    console.log('2ページ目データ整理シート(集計用):', page2DataSheet ? '存在' : '不存在');
    if (page2DataSheet) {
      addPage2Data(page2DataSheet, formData);
    }
    
    // 寄付金一覧シートにデータを追加（集計用ブック）
    const donationSheet = aggregationBook.getSheetByName('寄付金一覧');
    console.log('寄付金一覧シート(集計用):', donationSheet ? '存在' : '不存在');
    if (donationSheet) {
      addDonationData(donationSheet, formData);
      // 手動実行時と同じ処理を実行（ヘッダー初期化 + 集計）
      console.log('フォーム送信時の寄付金一覧集計更新を開始');
      try {
        // 寄付金一覧シートの初期化（ヘッダー設定）
        initializeDonationSummarySheet(donationSheet);
        console.log('フォーム送信時のヘッダー初期化完了');
        
        // 集計を実行
        updateDonationSummary(donationSheet);
        console.log('フォーム送信時の寄付金一覧集計更新完了');
      } catch (error) {
        console.error('フォーム送信時の寄付金一覧集計更新エラー:', error);
        // エラーが発生した場合は従来のupdateDonationSummaryのみを使用
        console.log('エラーが発生したため、従来の集計方法のみを使用');
        try {
          updateDonationSummary(donationSheet);
          console.log('従来の集計方法での更新完了');
        } catch (fallbackError) {
          console.error('従来の集計方法でもエラーが発生:', fallbackError);
        }
      }
    }
    
    // 返礼品一覧シートを更新（集計用ブック）
    const giftSheet = aggregationBook.getSheetByName('返礼品一覧');
    console.log('返礼品一覧シート(集計用):', giftSheet ? '存在' : '不存在');
    if (giftSheet) {
      // 手動実行時と同じ処理を実行（ヘッダー初期化 + 集計）
      console.log('フォーム送信時の返礼品一覧集計更新を開始');
      try {
        // 設定用ブックから必要なシートを取得
        const page1DataSheet = aggregationBook.getSheetByName('1ページ目データ整理');
        const page2DataSheet = aggregationBook.getSheetByName('2ページ目データ整理');
        const giftSettingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('返礼品設定');
        
        if (page1DataSheet && page2DataSheet && giftSettingsSheet) {
          // まず既存の集計行をクリア
          clearGiftSummaryRows(giftSheet);
          console.log('既存の集計行をクリア完了');
          
          // 新しい集計ロジックで更新
          updateGiftSummaryWithNewLogic(giftSheet, page1DataSheet, page2DataSheet, giftSettingsSheet);
          console.log('フォーム送信時の返礼品一覧集計更新完了');
          
          // テンプレートの貼り付けも実行（返礼品設定の同期）
          try {
            syncGiftSettingsToSheet(giftSheet);
            console.log('フォーム送信時のテンプレート貼り付け完了');
          } catch (templateError) {
            console.error('フォーム送信時のテンプレート貼り付けエラー:', templateError);
          }
        } else {
          console.log('必要なシートが見つからないため、返礼品一覧集計の更新をスキップ');
          // 代替手段として従来のupdateGiftSummaryを使用
          updateGiftSummary(giftSheet);
        }
      } catch (error) {
        console.error('フォーム送信時の返礼品一覧集計更新エラー:', error);
        // エラーが発生した場合は従来のupdateGiftSummaryを使用
        console.log('エラーが発生したため、従来の集計方法を使用');
        try {
          updateGiftSummary(giftSheet);
          console.log('従来の集計方法での更新完了');
        } catch (fallbackError) {
          console.error('従来の集計方法でもエラーが発生:', fallbackError);
        }
      }
    }
    
    // 個票を自動生成
    let individualReportResult = { success: true, message: '' };
    try {
      console.log('=== 個票自動生成開始 ===');
      console.log('formData構造:', JSON.stringify(formData, null, 2));
      
      // データ構造を正しく取得（ネストされた構造に対応）
      const basicInfo = formData.basicInfo.basicInfo || formData.basicInfo;
      const studentName = basicInfo.name;
      console.log('取得した生徒名:', studentName);
      
      if (studentName && studentName.trim() !== '') {
        console.log('個票を生成中:', studentName);
        
        // スプレッドシートへのデータ保存完了を待つ
        console.log('データ保存完了を待機中...');
        Utilities.sleep(2000); // 2秒待機
        
        generateIndividualReport(studentName);
        console.log('個票生成完了:', studentName);
        individualReportResult.message = '個票を正常に生成しました。';
      } else {
        console.log('生徒名が無効です:', studentName);
        individualReportResult.success = false;
        individualReportResult.message = '生徒名が無効なため個票を生成できませんでした。';
      }
      console.log('=== 個票自動生成終了 ===');
    } catch (error) {
      console.error('個票生成エラー:', error);
      console.error('エラースタック:', error.stack);
      individualReportResult.success = false;
      individualReportResult.message = `個票生成エラー: ${error.message}`;
      // 個票生成エラーがあってもフォーム送信は成功とする
    }
    
    // レスポンスメッセージを作成
    let responseMessage = 'データを正常に送信しました。';
    if (individualReportResult.success) {
      responseMessage += ' ' + individualReportResult.message;
    } else {
      responseMessage += ' ただし、' + individualReportResult.message;
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: responseMessage,
      individualReport: individualReportResult
    }))
    .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('submitFormDataエラー:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}

// 1ページ目データ整理シートにデータを追加
function addPage1Data(sheet, formData) {
  console.log('addPage1Data - formData:', JSON.stringify(formData, null, 2));
  
  // ヘッダー行を設定（初回または必要に応じて）
  setupPage1DataHeaders(sheet);
  
  // データ構造を正しく取得（ネストされた構造に対応）
  const basicInfo = formData.basicInfo.basicInfo || formData.basicInfo;
  const breakdownDetails = formData.basicInfo.breakdownDetails || formData.breakdownDetails;
  
  console.log('basicInfo:', JSON.stringify(basicInfo, null, 2));
  console.log('breakdownDetails:', JSON.stringify(breakdownDetails, null, 2));
  
      // 振込金額の取得（修正：totalTransferAmountではなくtransferAmountのみ）
    const transferAmount = parseInt(breakdownDetails.transferAmount) || 0;
    console.log('transferAmount (振込金額（直接振込を除く）):', transferAmount);
    
    // 直接振込金額の取得
    const directTransferAmount = parseInt(breakdownDetails.directTransferAmount) || 0;
    console.log('directTransferAmount:', directTransferAmount);
  
  // 負担金の計算（新しい兄弟機能では直接金額が送信される）
  let burdenAmount = 0;
  console.log('burdenFee value:', breakdownDetails.burdenFee, 'type:', typeof breakdownDetails.burdenFee);
  if (typeof breakdownDetails.burdenFee === 'number') {
    burdenAmount = breakdownDetails.burdenFee;
  } else if (typeof breakdownDetails.burdenFee === 'string') {
    burdenAmount = parseInt(breakdownDetails.burdenFee) || 0;
  }
  console.log('burdenAmount:', burdenAmount);
  
  // 奨励会負担金の計算（設定から動的に取得）
  let encouragementAmount = 0;
  console.log('hasEncouragementFee value:', breakdownDetails.hasEncouragementFee, 'type:', typeof breakdownDetails.hasEncouragementFee);
  if (breakdownDetails.hasEncouragementFee === true || breakdownDetails.hasEncouragementFee === 'true') {
    // 設定から奨励会負担金の金額を取得
    encouragementAmount = getEncouragementFeeAmount();
  }
  console.log('encouragementAmount:', encouragementAmount);
  
  // 基本の行データを準備
  const newRow = [
    formatDateTimeForJapan(new Date()), // A列：送信日時
    basicInfo.name,                    // B列：氏名
    basicInfo.grade,                   // C列：学年
    basicInfo.transferDate,            // D列：振込日
    transferAmount,                    // E列：振込金額
    directTransferAmount,              // F列：直接振込金額
    breakdownDetails.donationAmount,    // G列：寄付金
    burdenAmount,                      // H列：負担金
    breakdownDetails.tshirtAmount || 0, // I列：Tシャツ代
    encouragementAmount,               // J列：奨励会負担金
    breakdownDetails.otherFee || 0     // K列：その他の費用
  ];
  
  console.log('兄弟情報の詳細:', {
    hasSibling: basicInfo.hasSibling,
    siblingCount: basicInfo.siblingCount,
    sibling1Grade: basicInfo.sibling1Grade,
    sibling1Name: basicInfo.sibling1Name,
    sibling2Grade: basicInfo.sibling2Grade,
    sibling2Name: basicInfo.sibling2Name,
    sibling3Grade: basicInfo.sibling3Grade,
    sibling3Name: basicInfo.sibling3Name,
    sibling4Grade: basicInfo.sibling4Grade,
    sibling4Name: basicInfo.sibling4Name,
    sibling5Grade: basicInfo.sibling5Grade,
    sibling5Name: basicInfo.sibling5Name
  });
  
  // 兄弟の学年と名前を交互に追加（最大5人分）
  for (let i = 1; i <= 5; i++) {
    const siblingGrade = basicInfo[`sibling${i}Grade`] || '';
    const siblingName = basicInfo[`sibling${i}Name`] || '';
    
    // 兄弟iの学年を追加
    newRow.push(siblingGrade);
    console.log(`兄弟${i}の学年:`, siblingGrade, '列位置:', newRow.length, `(${String.fromCharCode(64 + newRow.length)}列)`);
    
    // 兄弟iの名前を追加
    newRow.push(siblingName);
    console.log(`兄弟${i}の名前:`, siblingName, '列位置:', newRow.length, `(${String.fromCharCode(64 + newRow.length)}列)`);
  }
  
  // V列: 兄弟の人数（フォームで選択された人数）
  let siblingCount = 0;
  
  if (basicInfo.hasSibling) {
    // 兄弟がいる場合のみ人数を取得
    if (basicInfo.siblingCount && basicInfo.siblingCount !== '') {
      siblingCount = parseInt(basicInfo.siblingCount) || 0;
    }
  }
  
  // 確実に数値として設定
  siblingCount = Number(siblingCount);
  console.log('=== 兄弟の人数デバッグ ===');
  console.log('basicInfo全体:', JSON.stringify(basicInfo, null, 2));
  console.log('basicInfo.siblingCount:', basicInfo.siblingCount);
  console.log('basicInfo.hasSibling:', basicInfo.hasSibling);
  console.log('basicInfo.siblingCount type:', typeof basicInfo.siblingCount);
  console.log('basicInfo.siblingCount === 0:', basicInfo.siblingCount === 0);
  console.log('basicInfo.siblingCount === "0":', basicInfo.siblingCount === "0");
  console.log('basicInfo.siblingCount == 0:', basicInfo.siblingCount == 0);
  console.log('計算された兄弟の人数:', siblingCount);
  console.log('計算された兄弟の人数 type:', typeof siblingCount);
  console.log('V列の位置:', newRow.length + 1, `(${String.fromCharCode(64 + newRow.length + 1)}列)`);
  newRow.push(siblingCount);
  console.log('兄弟の人数:', siblingCount, '列位置:', newRow.length, `(${String.fromCharCode(64 + newRow.length)}列)`);
  console.log('兄弟の人数の値:', newRow[newRow.length - 1]);
  console.log('兄弟の人数の値 type:', typeof newRow[newRow.length - 1]);
  
  newRow.push(''); // W列：備考列
  console.log('備考列の位置:', newRow.length, `(${String.fromCharCode(64 + newRow.length)}列)`);
  
  console.log('最終的なnewRow:', newRow);
  console.log('newRowの長さ:', newRow.length);
  console.log('期待される列数: 23列 (A列〜W列)');
  console.log('実際の列数:', newRow.length);
  
  // 列数チェック
  if (newRow.length !== 23) {
    console.warn(`警告: 期待される列数(23)と実際の列数(${newRow.length})が一致しません`);
  }
  
  sheet.appendRow(newRow);
}

// 2ページ目データ整理シートにデータを追加
function addPage2Data(sheet, formData) {
  console.log('addPage2Data - formData:', JSON.stringify(formData, null, 2));
  
  // データ構造を正しく取得（ネストされた構造に対応）
  const basicInfo = formData.basicInfo.basicInfo || formData.basicInfo;
  const donors = formData.donors;
  
  console.log('basicInfo:', JSON.stringify(basicInfo, null, 2));
  console.log('donors:', JSON.stringify(donors, null, 2));
  
  // 返礼品設定を取得してIDから名前へのマッピングを作成
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const giftSettingsSheet = spreadsheet.getSheetByName('返礼品設定');
  const giftIdToName = {};
  
  if (giftSettingsSheet) {
    const giftData = giftSettingsSheet.getDataRange().getValues();
    // ヘッダー行をスキップして2行目から処理
    for (let i = 1; i < giftData.length; i++) {
      const row = giftData[i];
      const id = row[0];
      const name = row[1];
      const isActive = row[3]; // 有効/無効の列はD列（4列目）
      
      if (isActive && id && name) {
        giftIdToName[id] = name;
      }
    }
  }
  
  console.log('giftIdToName mapping:', giftIdToName);
  
  // ヘッダー名で列位置を動的に特定
  const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndexMap = {};
  headerValues.forEach((h, idx) => { colIndexMap[String(h).trim()] = idx + 1; }); // 1-based

  // 必須列のインデックスを取得（見出し名で追従）
  const COL = {
    sentAt: colIndexMap['送信日時'] || 1,
    page1Name: colIndexMap['1ページ目氏名'] || 2,
    donorName: colIndexMap['寄付者名'] || 3,
    amount: colIndexMap['寄付金額'] || 4,
    gift: colIndexMap['返礼品'] || 5,
    clearfile: colIndexMap['クリアファイル枚数'] || 6,
    directTransfer: colIndexMap['直接振込'] || 7, // ご指定: G列
    note: colIndexMap['その他特筆事項'] || colIndexMap['備考'] || 8
  };

  donors.forEach((donor, index) => {
    console.log(`donor ${index}:`, JSON.stringify(donor, null, 2));
    console.log(`donor ${index} 直接振込状態:`, {
      isDirectTransfer: donor.isDirectTransfer,
      type: typeof donor.isDirectTransfer,
      value: donor.isDirectTransfer ? '✓' : ''
    });
    
    // 返礼品の処理（IDまたは名前の両方に対応）
    let giftName = '';
    if (donor.gift) {
      // まず返礼品設定シートからID→名前のマッピングを試行
      if (giftIdToName[donor.gift]) {
        giftName = giftIdToName[donor.gift];
      } else if (donor.gift === 'クリアファイル' || donor.gift === 'タオル' || donor.gift === 'お菓子大' || donor.gift === 'キーホルダー' || donor.gift === 'お菓子小') {
        // フォールバック: 既存の名前ベースの処理
        giftName = donor.gift;
      } else {
        // その他の場合
        giftName = donor.gift;
      }
    }
    
    console.log(`donor ${index} gift conversion:`, donor.gift, '->', giftName);
    
    // 行データ配列（既存ヘッダー長に合わせて初期化）
    const rowArray = new Array(headerValues.length).fill('');
    rowArray[COL.sentAt - 1] = formatDateTimeForJapan(new Date());
    rowArray[COL.page1Name - 1] = basicInfo.name;
    rowArray[COL.donorName - 1] = donor.name;
    rowArray[COL.amount - 1] = donor.amount;
    rowArray[COL.gift - 1] = giftName;
    if (COL.clearfile) rowArray[COL.clearfile - 1] = donor.sheets || 0;
    if (COL.directTransfer) {
      const directTransferValue = donor.isDirectTransfer ? '✓' : '';
      rowArray[COL.directTransfer - 1] = directTransferValue;
      console.log(`donor ${index} 直接振込出力:`, directTransferValue, '列位置:', COL.directTransfer);
    }
    if (COL.note) rowArray[COL.note - 1] = donor.note || '';
    
    console.log(`donor ${index} rowArray:`, rowArray);
    sheet.appendRow(rowArray);
  });
}

// 寄付金一覧シートにデータを追加（新しいレイアウト対応）
function addDonationData(sheet, formData) {
  // データ構造を正しく取得（ネストされた構造に対応）
  const basicInfo = formData.basicInfo.basicInfo || formData.basicInfo;
  const breakdownDetails = formData.basicInfo.breakdownDetails || formData.breakdownDetails;
  
  // 振込金額の取得
  const transferAmount = parseInt(breakdownDetails.transferAmount) || 0;
  
  // 負担金の計算（新しい兄弟機能では直接金額が送信される）
  let burdenAmount = 0;
  if (typeof breakdownDetails.burdenFee === 'number') {
    burdenAmount = breakdownDetails.burdenFee;
  } else if (typeof breakdownDetails.burdenFee === 'string') {
    burdenAmount = parseInt(breakdownDetails.burdenFee) || 0;
  }
  
  // 奨励会負担金の計算
  let encouragementAmount = 0;
  if (breakdownDetails.hasEncouragementFee === true || breakdownDetails.hasEncouragementFee === 'true') {
    encouragementAmount = 1000;
  }
  
  // 寄付・負担金計の計算
  const donationBurdenTotal = breakdownDetails.donationAmount + burdenAmount;
  
  // 兄弟の名前を取得（M、O、Q、S、U列から）
  const siblingNames = [];
  for (let i = 12; i <= 20; i += 2) { // M(13)、O(15)、Q(17)、S(19)、U(21)列
    const siblingName = basicInfo[`sibling${Math.floor((i-12)/2)+1}Name`];
    if (siblingName && siblingName.trim() !== '') {
      siblingNames.push(siblingName.trim());
    }
  }
  const siblingNote = siblingNames.length > 0 ? siblingNames.join('、') : '';
  
  const newRow = [
    basicInfo.name,                    // A列：氏名
    basicInfo.grade,                   // B列：学年
    basicInfo.transferDate,            // C列：振込日
    transferAmount,                    // D列：振込金額
    breakdownDetails.donationAmount,    // E列：寄付金（直接振込を除く）
    breakdownDetails.directTransferAmount || 0, // F列：直接振込金額
    burdenAmount,                      // G列：負担金
    breakdownDetails.tshirtAmount || 0, // H列：Tシャツ代
    encouragementAmount,               // I列：激励会負担金
    breakdownDetails.otherFee || 0,    // J列：その他の費用
    donationBurdenTotal,               // K列：寄付・負担金計
    siblingNote,                       // L列：備考（兄弟の名前）
    ''                                 // M列：チェック（空白）
  ];
  
  // 2行目に挿入（既存データを下にずらす）
  try {
    sheet.insertRowBefore(2);
    sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
  } catch (error) {
    console.error('データ行の挿入エラー:', error);
    throw error;
  }
}

// 寄付金一覧シートを学年別に集計・整理
function updateDonationSummary(sheet) {
  try {
    console.log('=== 寄付金一覧シートの学年別集計を開始 ===');
    console.log('対象シート名:', sheet.getName());
    console.log('対象シートID:', sheet.getSheetId());
    
    // ヘッダー行を設定
    const headers = [
      '氏名', '学年', '振込日', '振込金額', '寄付金\n*直接振込を除く', 
      '直接振込金額', '負担金', 'Tシャツ代', '激励会負担金', 'その他の費用', 
      '寄付・負担金計', '備考', 'チェック'
    ];
    
    console.log('ヘッダー設定開始:', headers);
    
    // ヘッダーを設定
    try {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      console.log('ヘッダー範囲設定:', `行1, 列1, 行数1, 列数${headers.length}`);
      headerRange.setValues([headers]);
      
      // ヘッダー行のテキスト折り返しを有効にする
      headerRange.setWrap(true);
      
      console.log('ヘッダー設定完了');
    } catch (error) {
      console.error('ヘッダー行の設定エラー:', error);
      console.error('エラーの詳細:', error.stack);
      throw error;
    }
    
    // 1ページ目データ整理シートからデータを取得
    console.log('集計用ブックの取得開始');
    const spreadsheet = getAggregationBook();
    console.log('集計用ブック名:', spreadsheet.getName());
    console.log('集計用ブックID:', spreadsheet.getId());
    
    const page1Sheet = spreadsheet.getSheetByName('1ページ目データ整理');
    
    if (!page1Sheet) {
      console.error('1ページ目データ整理シートが見つかりません');
      throw new Error('1ページ目データ整理シートが見つかりません');
    }
    
    console.log('1ページ目データ整理シート取得成功');
    
    const page1Data = page1Sheet.getDataRange().getValues();
    console.log('1ページ目データ行数:', page1Data.length);
    console.log('1ページ目データ列数:', page1Data[0] ? page1Data[0].length : 0);
    
    if (page1Data.length <= 1) {
      console.log('1ページ目データ整理シートにデータがありません');
      return;
    }
    
    // ヘッダー行をスキップして2行目から処理
    const dataRows = page1Data.slice(1);
    console.log('処理対象データ行数:', dataRows.length);
    
    // 学年別にデータを整理
    const gradeData = {};
    const gradeOrder = ['3年', '2年', '1年']; // 3年が一番上
    
    console.log('データ収集開始');
    
    // 各学年のデータを収集
    dataRows.forEach((row, index) => {
      try {
        console.log(`行${index + 2}の処理開始:`, row);
        
        const name = row[1];        // B列：氏名
        const grade = row[2];       // C列：学年
        const transferDate = row[3]; // D列：振込日
        const transferAmount = row[4]; // E列：振込金額
        const directTransferAmount = row[5]; // F列：直接振込金額
        const donationAmount = row[6]; // G列：寄付金
        const burdenAmount = row[7]; // H列：負担金
        const tshirtAmount = row[8]; // I列：Tシャツ代
        const encouragementAmount = row[9]; // J列：激励会負担金
        const otherFee = row[10];   // K列：その他の費用
        
        console.log(`行${index + 2}の基本データ:`, {
          name, grade, transferDate, transferAmount, donationAmount, burdenAmount
        });
        
        // データの妥当性チェック
        if (!name || !grade || !transferDate) {
          console.log(`行${index + 2}: 不完全なデータをスキップ:`, { name, grade, transferDate });
          return;
        }
        
        // 兄弟の名前を取得（L〜U列から）
        const siblingNames = [];
        console.log(`行${index + 2}の兄弟データ列確認:`, {
          'L列(12)': row[11], 'M列(13)': row[12],
          'N列(14)': row[13], 'O列(15)': row[14],
          'P列(16)': row[15], 'Q列(17)': row[16],
          'R列(18)': row[17], 'S列(19)': row[18],
          'T列(20)': row[19], 'U列(21)': row[20]
        });
        
        // V列から兄弟の人数を取得（フォームで選択された人数）
        const siblingCount = parseInt(row[21]) || 0; // V列（22列目）= 21 (0-based)
        console.log(`行${index + 2}の兄弟の人数:`, siblingCount, '列位置: V列(22列目)');
        
        // 兄弟の名前を取得（M、O、Q、S、U列から）
        const nameColumns = [12, 14, 16, 18, 20];
        console.log(`兄弟名の列をチェック中:`, nameColumns.map(col => `${col + 1}列目`));
        
        for (const col of nameColumns) {
          const siblingName = row[col];
          console.log(`${col + 1}列目の値: "${siblingName}" (型: ${typeof siblingName})`);
          if (siblingName && String(siblingName).trim() !== '') {
            console.log(`兄弟の名前を発見: ${siblingName} (${col + 1}列目)`);
            
            // 兄弟の名前をsiblingNames配列にも追加（備考用）
            const siblingGrade = row[col - 1]; // 学年の列（名前の列の1つ前）
            if (siblingGrade && String(siblingGrade).trim() !== '') {
              // 学年がある場合は「名前（学年）」の形式
              siblingNames.push(`${siblingName.toString().trim()}（${siblingGrade.toString().trim()}）`);
            } else {
              // 学年がない場合は名前のみ
              siblingNames.push(siblingName.toString().trim());
            }
          }
        }
        const siblingNote = siblingNames.length > 0 ? siblingNames.join('、') : '';
        
        console.log(`行${index + 2}の兄弟情報:`, siblingNames);
        
        // 寄付・負担金計を計算
        const donationBurdenTotal = (parseInt(donationAmount) || 0) + (parseInt(burdenAmount) || 0);
        
        if (!gradeData[grade]) {
          gradeData[grade] = [];
        }
        
        const processedData = {
          name: name,
          grade: grade,
          transferDate: transferDate,
          transferAmount: transferAmount,
          donationAmount: donationAmount,
          directTransferAmount: directTransferAmount,
          burdenAmount: burdenAmount,
          tshirtAmount: tshirtAmount,
          encouragementAmount: encouragementAmount,
          otherFee: otherFee,
          donationBurdenTotal: donationBurdenTotal,
          siblingNote: siblingNote
        };
        
        gradeData[grade].push(processedData);
        console.log(`行${index + 2}の処理完了:`, processedData);
        
      } catch (error) {
        console.error(`行${index + 2}の処理中にエラー:`, error);
        console.error('エラーの詳細:', error.stack);
        // エラーが発生しても処理を続行
      }
    });
    
    console.log('学年別データ収集結果:', Object.keys(gradeData).map(grade => `${grade}: ${gradeData[grade].length}件`));
    
    // 既存データをクリア（ヘッダー行は保持）
    console.log('既存データのクリア開始');
    const lastRow = sheet.getLastRow();
    console.log('最終行:', lastRow);
    
    if (lastRow > 1) {
      try {
        const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
        console.log('クリア対象範囲:', `行2, 列1, 行数${lastRow - 1}, 列数${lastColumn}`);
        const clearRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
        clearRange.clearContent();
        // 書式も完全にクリア
        clearRange.clearFormat();
        console.log('既存データと書式のクリア完了');
      } catch (error) {
        console.error('既存データクリアエラー:', error);
        console.error('エラーの詳細:', error.stack);
        // エラーが発生しても処理を続行
      }
    }
    
    let currentRow = 2; // ヘッダー行の次から開始
    let grandTotal = {
      transferAmount: 0,
      donationAmount: 0,
      directTransferAmount: 0,
      burdenAmount: 0,
      tshirtAmount: 0,
      encouragementAmount: 0,
      otherFee: 0,
      donationBurdenTotal: 0
    };
    
    console.log('学年別データ出力開始');
    
    // 学年順にデータを出力
    gradeOrder.forEach(grade => {
      try {
        if (gradeData[grade] && gradeData[grade].length > 0) {
          console.log(`${grade}のデータを処理中: ${gradeData[grade].length}件`);
          
          let gradeTotal = {
            transferAmount: 0,
            donationAmount: 0,
            directTransferAmount: 0,
            burdenAmount: 0,
            tshirtAmount: 0,
            encouragementAmount: 0,
            otherFee: 0,
            donationBurdenTotal: 0
          };
          
          // その学年のデータを出力
          gradeData[grade].forEach((data, dataIndex) => {
            try {
              console.log(`${grade}のデータ${dataIndex + 1}件目を処理:`, data.name);
              
              const row = [
                data.name,                    // A列：氏名
                data.grade,                   // B列：学年
                data.transferDate,            // C列：振込日
                data.transferAmount,          // D列：振込金額
                data.donationAmount,          // E列：寄付金（直接振込を除く）
                data.directTransferAmount,    // F列：直接振込金額
                data.burdenAmount,            // G列：負担金
                data.tshirtAmount,            // H列：Tシャツ代
                data.encouragementAmount,     // I列：激励会負担金
                data.otherFee,                // J列：その他の費用
                data.donationBurdenTotal,     // K列：寄付・負担金計
                data.siblingNote,             // L列：備考
                ''                            // M列：チェック
              ];
              
              console.log(`行${currentRow}のデータ設定:`, row);
              
              try {
                const dataRange = sheet.getRange(currentRow, 1, 1, row.length);
                console.log(`データ範囲設定: 行${currentRow}, 列1, 行数1, 列数${row.length}`);
                dataRange.setValues([row]);
                console.log(`行${currentRow}のデータ設定完了`);
                currentRow++;
              } catch (error) {
                console.error(`行${currentRow}のデータ設定エラー:`, error);
                console.error('エラーの詳細:', error.stack);
                throw error;
              }
              
              // 学年小計を計算
              gradeTotal.transferAmount += parseInt(data.transferAmount) || 0;
              gradeTotal.donationAmount += parseInt(data.donationAmount) || 0;
              gradeTotal.directTransferAmount += parseInt(data.directTransferAmount) || 0;
              gradeTotal.burdenAmount += parseInt(data.burdenAmount) || 0;
              gradeTotal.tshirtAmount += parseInt(data.tshirtAmount) || 0;
              gradeTotal.encouragementAmount += parseInt(data.encouragementAmount) || 0;
              gradeTotal.otherFee += parseInt(data.otherFee) || 0;
              gradeTotal.donationBurdenTotal += parseInt(data.donationBurdenTotal) || 0;
              
            } catch (error) {
              console.error(`${grade}のデータ${dataIndex + 1}件目の処理エラー:`, error);
              console.error('エラーの詳細:', error.stack);
              // エラーが発生しても処理を続行
            }
          });
          
          console.log(`${grade}の小計計算完了:`, gradeTotal);
          
          // 学年小計行を挿入
          const gradeSubtotalRow = [
            `${grade}小計`,              // A列：氏名
            '',                          // B列：学年
            '',                          // C列：振込日
            gradeTotal.transferAmount,   // D列：振込金額
            gradeTotal.donationAmount,   // E列：寄付金（直接振込を除く）
            gradeTotal.directTransferAmount, // F列：直接振込金額
            gradeTotal.burdenAmount,     // G列：負担金
            gradeTotal.tshirtAmount,     // H列：Tシャツ代
            gradeTotal.encouragementAmount, // I列：激励会負担金
            gradeTotal.otherFee,         // J列：その他の費用
            gradeTotal.donationBurdenTotal, // K列：寄付・負担金計
            '',                          // L列：備考
            ''                           // M列：チェック
          ];
          
          console.log(`行${currentRow}の${grade}小計行設定:`, gradeSubtotalRow);
          
          try {
            const subtotalRange = sheet.getRange(currentRow, 1, 1, gradeSubtotalRow.length);
            console.log(`小計行範囲設定: 行${currentRow}, 列1, 行数1, 列数${gradeSubtotalRow.length}`);
            subtotalRange.setValues([gradeSubtotalRow]);
            
            // 小計行の背景色を設定
            subtotalRange.setBackground('#E8F5E8');
            sheet.getRange(currentRow, 1, 1, 1).setFontWeight('bold'); // 氏名列を太字
            console.log(`${grade}小計行の設定完了`);
          } catch (error) {
            console.error(`${grade}小計行の設定エラー:`, error);
            console.error('エラーの詳細:', error.stack);
            throw error;
          }
          
          currentRow++;
          
          // 全体合計に加算
          grandTotal.transferAmount += gradeTotal.transferAmount;
          grandTotal.donationAmount += gradeTotal.donationAmount;
          grandTotal.directTransferAmount += gradeTotal.directTransferAmount;
          grandTotal.burdenAmount += gradeTotal.burdenAmount;
          grandTotal.tshirtAmount += gradeTotal.tshirtAmount;
          grandTotal.encouragementAmount += gradeTotal.encouragementAmount;
          grandTotal.otherFee += gradeTotal.otherFee;
          grandTotal.donationBurdenTotal += gradeTotal.donationBurdenTotal;
          
        } else {
          console.log(`${grade}のデータはありません`);
        }
        
      } catch (error) {
        console.error(`${grade}の処理中にエラー:`, error);
        console.error('エラーの詳細:', error.stack);
        // エラーが発生しても処理を続行
      }
    });
    
    console.log('全体合計計算完了:', grandTotal);
    
    // 全体合計行を挿入
    const grandTotalRow = [
      '全体合計',                    // A列：氏名
      '',                           // B列：学年
      '',                           // C列：振込日
      grandTotal.transferAmount,    // D列：振込金額
      grandTotal.donationAmount,    // E列：寄付金（直接振込を除く）
      grandTotal.directTransferAmount, // F列：直接振込金額
      grandTotal.burdenAmount,      // G列：負担金
      grandTotal.tshirtAmount,      // H列：Tシャツ代
      grandTotal.encouragementAmount, // I列：激励会負担金
      grandTotal.otherFee,          // J列：その他の費用
      grandTotal.donationBurdenTotal, // K列：寄付・負担金計
      '',                           // L列：備考
      ''                            // M列：チェック
    ];
    
    console.log(`行${currentRow}の全体合計行設定:`, grandTotalRow);
    
    try {
      const grandTotalRange = sheet.getRange(currentRow, 1, 1, grandTotalRow.length);
      console.log(`全体合計行範囲設定: 行${currentRow}, 列1, 行数1, 列数${grandTotalRow.length}`);
      grandTotalRange.setValues([grandTotalRow]);
      
      // 全体合計行の背景色を設定
      grandTotalRange.setBackground('#FFE8E8');
      sheet.getRange(currentRow, 1, 1, 1).setFontWeight('bold'); // 氏名列を太字
      console.log('全体合計行の設定完了');
    } catch (error) {
      console.error('全体合計行の設定エラー:', error);
      console.error('エラーの詳細:', error.stack);
      throw error;
    }
    
    // 列幅を適切に設定（集計後も列幅を維持）
    console.log('列幅設定開始');
    const columnWidths = [
      120,  // A列：氏名
      50,   // B列：学年（最小化）
      100,  // C列：振込日
      100,  // D列：振込金額
      120,  // E列：寄付金*直接振込を除く（縮小）
      120,  // F列：直接振込金額
      80,   // G列：負担金
      80,   // H列：Tシャツ代
      120,  // I列：激励会負担金
      100,  // J列：その他の費用
      120,  // K列：寄付・負担金計
      150,  // L列：備考
      80    // M列：チェック
    ];
    
    // 各列の幅を設定
    columnWidths.forEach((width, index) => {
      try {
        sheet.setColumnWidth(index + 1, width);
        console.log(`列${index + 1}の幅設定完了: ${width}px`);
      } catch (error) {
        console.error(`列${index + 1}の幅設定エラー:`, error);
        console.error('エラーの詳細:', error.stack);
      }
    });
    
    // 条件付き書式を適用
    console.log('条件付き書式の適用開始');
    applyConditionalFormatting(sheet);
    
    console.log('=== 寄付金一覧シートの学年別集計が完了しました ===');
    
  } catch (error) {
    console.error('寄付金一覧シートの学年別集計エラー:', error);
    console.error('エラーの詳細:', error.stack);
    throw error;
  }
}

// 返礼品一覧シートを動的に更新
function updateGiftSummary(sheet) {
  try {
    console.log('=== 返礼品一覧シートの更新を開始 ===');
    
    const aggregationBook = getAggregationBook();
    
    // 設定用ブックから返礼品設定を取得
    const settingsBook = SpreadsheetApp.getActiveSpreadsheet();
    const giftSettingsSheet = settingsBook.getSheetByName('返礼品設定');
    if (!giftSettingsSheet) {
      console.log('返礼品設定シートが見つかりません');
      return;
    }
    
    // 2ページ目データ整理シートを取得
    const page2DataSheet = aggregationBook.getSheetByName('2ページ目データ整理');
    if (!page2DataSheet) {
      console.log('2ページ目データ整理シートが見つかりません');
      return;
    }
    
    // 1ページ目データ整理シートを取得
    const page1DataSheet = aggregationBook.getSheetByName('1ページ目データ整理');
    if (!page1DataSheet) {
      console.log('1ページ目データ整理シートが見つかりません');
      return;
    }
    
    // 返礼品設定を返礼品一覧シートの2行目に同期
    syncGiftSettingsToSheet(sheet);
    
    // 新しい集計ロジックでデータを更新
    updateGiftSummaryWithNewLogic(sheet, page1DataSheet, page2DataSheet, giftSettingsSheet);
    
    console.log('=== 返礼品一覧シートの更新が完了しました ===');
    
  } catch (error) {
    console.error('返礼品一覧シートの更新エラー:', error);
    throw error;
  }
}

// 返礼品設定を取得
function getGiftSettings(sheet) {
  const data = sheet.getDataRange().getValues();
  const settings = [];
  
  // ヘッダー行をスキップして2行目から処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    const name = row[1];
    const isActive = row[3]; // 新スキーマ: D列
    
    if (isActive && id && name) {
      settings.push({ id: id, name: name });
    }
  }
  
  return settings;
}

// 新しい集計ロジックで返礼品一覧シートを更新
function updateGiftSummaryWithNewLogic(sheet, page1DataSheet, page2DataSheet, giftSettingsSheet) {
  try {
    console.log('新しい集計ロジックで返礼品一覧を更新開始');
    
    // 1ページ目データから氏名と学年を取得
    const page1Data = page1DataSheet.getDataRange().getValues();
    const nameGradeMap = {};
    
    // ヘッダー行をスキップして2行目から処理
    for (let i = 1; i < page1Data.length; i++) {
      const row = page1Data[i];
      const name = row[1]; // B列：氏名
      const grade = row[2]; // C列：学年
      
      if (name && name.toString().trim() !== '') {
        nameGradeMap[name.toString().trim()] = grade;
      }
    }
    
    console.log('1ページ目データから取得した氏名・学年マップ:', nameGradeMap);
    
    // 2ページ目データから返礼品を集計
    const page2Data = page2DataSheet.getDataRange().getValues();
    const giftSummary = {};
    
    // ヘッダー行をスキップして2行目から処理
    for (let i = 1; i < page2Data.length; i++) {
      const row = page2Data[i];
      const page1Name = row[1]; // B列：1ページ目氏名
      const gift = row[4]; // E列：返礼品（名前）
      const sheets = parseInt(row[5]) || 0; // F列：クリアファイル枚数
      
      if (page1Name && page1Name.toString().trim() !== '' && nameGradeMap[page1Name.toString().trim()]) {
        const trimmedName = page1Name.toString().trim();
        
        if (!giftSummary[trimmedName]) {
          giftSummary[trimmedName] = {
            grade: nameGradeMap[trimmedName],
            gifts: {},
            clearfileTotal: 0,
            donorCount: 0
          };
        }
        
        // 寄付者数をカウント
        giftSummary[trimmedName].donorCount++;
        
        // 返礼品のカウント
        if (gift && gift.toString().trim() !== '') {
          const giftName = gift.toString().trim();
          giftSummary[trimmedName].gifts[giftName] = (giftSummary[trimmedName].gifts[giftName] || 0) + 1;
          
          // クリアファイルの返礼品がある場合のみ、その枚数を加算
          if (giftName === '③ クリアファイル') {
            giftSummary[trimmedName].clearfileTotal += sheets;
          }
        }
      }
    }
    
    console.log('2ページ目データから集計した返礼品サマリー:', giftSummary);
    
    // 返礼品一覧シートの2行目から設定値を取得
    const giftSheetData = sheet.getDataRange().getValues();
    if (giftSheetData.length < 2) {
      console.error('返礼品一覧シートに2行目の設定行がありません');
      return;
    }
    
    const configRow = giftSheetData[1]; // 2行目（インデックス1）
    console.log('返礼品一覧シート2行目の設定値:', configRow);
    
            // 設定列の返礼品名を取得
        const configGifts = {
          D2: configRow[3], // D列（インデックス3）
          F2: configRow[5], // F列（インデックス5）
          H2: configRow[7], // H列（インデックス7）
          I2: configRow[8], // I列（インデックス8）
          J2: configRow[9], // J列（インデックス9）
          L2: configRow[11] // L列（インデックス11）
        };
    
    console.log('設定列の返礼品名:', configGifts);
    
    // 既存データをクリア（ヘッダー行とテンプレート行は保持）
    const lastRow = sheet.getLastRow();
    if (lastRow > 5) {
      try {
        const clearRange = sheet.getRange(6, 1, lastRow - 5, sheet.getLastColumn());
        clearRange.clearContent();
        clearRange.clearFormat();
        console.log('既存データ行（6行目以降）をクリアしました');
      } catch (error) {
        console.error('既存データクリアエラー:', error);
      }
    }
    
    // 新しいデータ行を追加
    let currentRow = 6; // 6行目から開始
    let rowNumber = 1;
    
    // 学年別の集計用データ
    const gradeSummary = {
      '1年': { gifts: {}, clearfileTotal: 0, donorCount: 0 },
      '2年': { gifts: {}, clearfileTotal: 0, donorCount: 0 },
      '3年': { gifts: {}, clearfileTotal: 0, donorCount: 0 }
    };
    
    Object.keys(giftSummary).forEach(name => {
      try {
        const summary = giftSummary[name];
        console.log(`生徒「${name}」のデータを処理中:`, summary);
        
        // 4-5行目のテンプレートをコピー（書式・行幅を含めて完全コピー）
        const templateRange = sheet.getRange(4, 1, 2, sheet.getLastColumn());
        const newRowRange = sheet.getRange(currentRow, 1, 2, sheet.getLastColumn());
        
        // 書式・行幅を含めて完全にコピー
        templateRange.copyTo(newRowRange, {contentsOnly: false});
        
        // 行幅をコピー（copyToでは行幅が反映されないため個別に設定）
        const templateRow4 = sheet.getRange(4, 1, 1, sheet.getLastColumn());
        const templateRow5 = sheet.getRange(5, 1, 1, sheet.getLastColumn());
        const newRow1 = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn());
        const newRow2 = sheet.getRange(currentRow + 1, 1, 1, sheet.getLastColumn());
        
        // 4行目の書式を6行目に、5行目の書式を7行目にコピー
        templateRow4.copyTo(newRow1, {contentsOnly: false});
        templateRow5.copyTo(newRow2, {contentsOnly: false});
        
        // テンプレート行の行高さを取得して、データ行に適用
        const templateRow4Height = sheet.getRowHeight(4);
        const templateRow5Height = sheet.getRowHeight(5);
        
        // データ行の行高さをテンプレート行と同じに設定
        sheet.setRowHeight(currentRow, templateRow4Height);
        sheet.setRowHeight(currentRow + 1, templateRow5Height);
        
        console.log(`行高さを設定: 行${currentRow}=${templateRow4Height}, 行${currentRow + 1}=${templateRow5Height}`);
        
        // 基本情報を設定
        sheet.getRange(currentRow, 1).setValue(rowNumber); // A列：連番（1から開始）
        sheet.getRange(currentRow, 2).setValue(summary.grade); // B列：学年
        sheet.getRange(currentRow, 3).setValue(name); // C列：1ページ目氏名
        
        // 返礼品の数を設定
        const giftCounts = {
          D2: configGifts.D2 ? (summary.gifts[configGifts.D2] || 0) : 0,
          F2: configGifts.F2 ? (summary.gifts[configGifts.F2] || 0) : 0,
          H2: configGifts.H2 ? (summary.gifts[configGifts.H2] || 0) : 0,
          I2: configGifts.I2 ? (summary.gifts[configGifts.I2] || 0) : 0,
          J2: configGifts.J2 ? (summary.gifts[configGifts.J2] || 0) : 0,
          L2: configGifts.L2 ? (summary.gifts[configGifts.L2] || 0) : 0
        };
        
        // ファイル数の計算
        // M列 = クリアファイル枚数の合計 + 返礼品の数の合計 - L2の返礼品数
        const totalGiftCount = Object.values(giftCounts).reduce((sum, count) => sum + count, 0);
        const fileCount = summary.clearfileTotal + totalGiftCount - (giftCounts['L'] || 0);
        
        // お礼状の計算
        // O列 = 寄付者数 + 1
        const thankYouCount = summary.donorCount + 1;
        
        // 各列に返礼品の数を設定
        sheet.getRange(currentRow, 5).setValue(giftCounts.D2);  // E列：D2の返礼品数
        sheet.getRange(currentRow, 7).setValue(giftCounts.F2);  // G列：F2の返礼品数
        sheet.getRange(currentRow, 9).setValue(giftCounts.H2);  // I列：H2の返礼品数
        sheet.getRange(currentRow, 11).setValue(giftCounts.I2); // K列：I2の返礼品数
        sheet.getRange(currentRow, 13).setValue(fileCount);     // M列：ファイル数
        sheet.getRange(currentRow, 15).setValue(thankYouCount); // O列：お礼状数
        
        // 学年別集計に加算
        const grade = summary.grade;
        if (gradeSummary[grade]) {
          gradeSummary[grade].donorCount += summary.donorCount;
          gradeSummary[grade].clearfileTotal += summary.clearfileTotal;
          
          // 各返礼品の数を学年別集計に加算
          if (giftCounts.D2) {
            if (!gradeSummary[grade].gifts.D2) gradeSummary[grade].gifts.D2 = 0;
            gradeSummary[grade].gifts.D2 += giftCounts.D2;
          }
          if (giftCounts.F2) {
            if (!gradeSummary[grade].gifts.F2) gradeSummary[grade].gifts.F2 = 0;
            gradeSummary[grade].gifts.F2 += giftCounts.F2;
          }
          if (giftCounts.H2) {
            if (!gradeSummary[grade].gifts.H2) gradeSummary[grade].gifts.H2 = 0;
            gradeSummary[grade].gifts.H2 += giftCounts.H2;
          }
          if (giftCounts.I2) {
            if (!gradeSummary[grade].gifts.I2) gradeSummary[grade].gifts.I2 = 0;
            gradeSummary[grade].gifts.I2 += giftCounts.I2;
          }
          if (giftCounts.J2) {
            if (!gradeSummary[grade].gifts.J2) gradeSummary[grade].gifts.J2 = 0;
            gradeSummary[grade].gifts.J2 += giftCounts.J2;
          }
          if (giftCounts.L2) {
            if (!gradeSummary[grade].gifts.L2) gradeSummary[grade].gifts.L2 = 0;
            gradeSummary[grade].gifts.L2 += giftCounts.L2;
          }
          
          // ファイル数とお礼状数も学年別集計に加算
          if (!gradeSummary[grade].fileTotal) gradeSummary[grade].fileTotal = 0;
          if (!gradeSummary[grade].thankYouTotal) gradeSummary[grade].thankYouTotal = 0;
          gradeSummary[grade].fileTotal += fileCount;
          gradeSummary[grade].thankYouTotal += thankYouCount;
        }
        
        console.log(`行${currentRow}の設定完了:`, {
          name, grade: summary.grade, giftCounts, fileCount, thankYouCount
        });
        
        currentRow += 2; // 2行分のテンプレートをコピーしたので2行進める
        rowNumber++;
        
      } catch (error) {
        console.error(`生徒「${name}」のデータ処理エラー:`, error);
        currentRow += 2; // エラーが発生しても行を進める
        rowNumber++;
      }
    });
    
    // 学年別小計行を追加
    const grades = ['1年', '2年', '3年'];
    grades.forEach(grade => {
      if (gradeSummary[grade].donorCount > 0) {
        try {
          console.log(`${grade}小計行を追加中:`, gradeSummary[grade]);
          
          // 小計行の背景色を設定
          const subtotalRow = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn());
          subtotalRow.setBackground('#E8F5E8'); // 薄い緑色
          subtotalRow.setFontWeight('bold');
          
          // 小計行の内容を設定
          sheet.getRange(currentRow, 1).setValue(''); // A列：空白
          sheet.getRange(currentRow, 2).setValue(grade); // B列：学年
          sheet.getRange(currentRow, 3).setValue(`${grade}小計`); // C列：学年小計
          
          // 各返礼品の小計を設定
          sheet.getRange(currentRow, 5).setValue(gradeSummary[grade].gifts.D2 || 0);  // E列：D2の返礼品数小計
          sheet.getRange(currentRow, 7).setValue(gradeSummary[grade].gifts.F2 || 0);  // G列：F2の返礼品数小計
          sheet.getRange(currentRow, 9).setValue(gradeSummary[grade].gifts.H2 || 0);  // I列：H2の返礼品数小計
          sheet.getRange(currentRow, 11).setValue(gradeSummary[grade].gifts.I2 || 0); // K列：I2の返礼品数小計
          
          // D,F,H,J,L,N,P列に返礼品名を設定（4行目の値と同じ）
          sheet.getRange(currentRow, 4).setValue(sheet.getRange(4, 4).getValue() || '');  // D列：D4の値（タオル）
          sheet.getRange(currentRow, 6).setValue(sheet.getRange(4, 6).getValue() || '');  // F列：F4の値（お菓子大）
          sheet.getRange(currentRow, 8).setValue(sheet.getRange(4, 8).getValue() || '');  // H列：H4の値（キーホルダー）
          sheet.getRange(currentRow, 10).setValue(sheet.getRange(4, 10).getValue() || ''); // J列：J4の値（お菓子小）
          sheet.getRange(currentRow, 12).setValue(sheet.getRange(4, 12).getValue() || ''); // L列：L4の値（クリアファイル）
          sheet.getRange(currentRow, 14).setValue(sheet.getRange(4, 14).getValue() || ''); // N列：N4の値
          sheet.getRange(currentRow, 16).setValue(sheet.getRange(4, 16).getValue() || ''); // P列：P4の値
          
          // ファイル数とお礼状数の小計（集計済みの値を使用）
          const fileCount = gradeSummary[grade].fileTotal || 0;
          const thankYouCount = gradeSummary[grade].thankYouTotal || 0;
          
          sheet.getRange(currentRow, 13).setValue(fileCount);     // M列：ファイル数小計
          sheet.getRange(currentRow, 15).setValue(thankYouCount); // O列：お礼状数小計
          
          // Q列の集計（該当学年のQ6~の和を計算）
          let qColumnTotal = 0;
          for (let row = 6; row < currentRow; row += 2) { // 6行目から現在の行まで2行ずつ
            const gradeInRow = sheet.getRange(row, 2).getValue(); // B列：学年
            if (gradeInRow === grade) { // 該当学年のみ集計
              const qValue = sheet.getRange(row, 17).getValue(); // Q列（17列目）
              if (qValue && !isNaN(qValue)) {
                qColumnTotal += Number(qValue);
              }
            }
          }
          sheet.getRange(currentRow, 17).setValue(qColumnTotal); // Q列：小計
          
          currentRow++;
          console.log(`${grade}小計行の追加完了`);
          
        } catch (error) {
          console.error(`${grade}小計行の追加エラー:`, error);
          currentRow++;
        }
      }
    });
    
        console.log('学年別小計行追加完了。現在の行番号:', currentRow);
    
    // 全体合計行を追加
    try {
      console.log('全体合計行を追加中。現在の行番号:', currentRow);
      
      // currentRowが正しく設定されているか確認
      if (currentRow <= 0) {
        console.error('currentRowが正しく設定されていません:', currentRow);
        currentRow = sheet.getLastRow() + 1;
        console.log('currentRowを修正しました:', currentRow);
      }
      
      // 全体合計の計算
      const grandTotal = {
        gifts: {},
        clearfileTotal: 0,
        donorCount: 0
      };
      
      // 各学年の小計を合計
      Object.values(gradeSummary).forEach(gradeData => {
        grandTotal.donorCount += gradeData.donorCount;
        grandTotal.clearfileTotal += gradeData.clearfileTotal;
        
        Object.keys(gradeData.gifts).forEach(colLetter => {
          if (!grandTotal.gifts[colLetter]) {
            grandTotal.gifts[colLetter] = 0;
          }
          grandTotal.gifts[colLetter] += gradeData.gifts[colLetter];
        });
      });
      
      console.log('全体合計計算完了:', grandTotal);
      
      // 全体合計行の背景色を設定
      const grandTotalRow = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn());
      grandTotalRow.setBackground('#FFE8E8'); // 薄い赤色
      grandTotalRow.setFontWeight('bold');
      
      // 全体合計行の内容を設定
      sheet.getRange(currentRow, 1).setValue(''); // A列：空白
      sheet.getRange(currentRow, 2).setValue(''); // B列：空白
      sheet.getRange(currentRow, 3).setValue('全体合計'); // C列：全体合計
      
                // 各返礼品の全体合計を設定
          sheet.getRange(currentRow, 5).setValue(grandTotal.gifts.D2 || 0);  // E列：D2の返礼品数全体合計
          sheet.getRange(currentRow, 7).setValue(grandTotal.gifts.F2 || 0);  // G列：F2の返礼品数全体合計
          sheet.getRange(currentRow, 9).setValue(grandTotal.gifts.H2 || 0);  // I列：H2の返礼品数全体合計
          sheet.getRange(currentRow, 11).setValue(grandTotal.gifts.I2 || 0); // K列：I2の返礼品数全体合計
          
          // D,F,H,J,L,N,P列に返礼品名を設定（4行目の値と同じ）
          sheet.getRange(currentRow, 4).setValue(sheet.getRange(4, 4).getValue() || '');  // D列：D4の値（タオル）
          sheet.getRange(currentRow, 6).setValue(sheet.getRange(4, 6).getValue() || '');  // F列：F4の値（お菓子大）
          sheet.getRange(currentRow, 8).setValue(sheet.getRange(4, 8).getValue() || '');  // H列：H4の値（キーホルダー）
          sheet.getRange(currentRow, 10).setValue(sheet.getRange(4, 10).getValue() || ''); // J列：J4の値（お菓子小）
          sheet.getRange(currentRow, 12).setValue(sheet.getRange(4, 12).getValue() || ''); // L列：L4の値（クリアファイル）
          sheet.getRange(currentRow, 14).setValue(sheet.getRange(4, 14).getValue() || ''); // N列：N4の値
          sheet.getRange(currentRow, 16).setValue(sheet.getRange(4, 16).getValue() || ''); // P列：P4の値
      
      // ファイル数とお礼状数の全体合計（各学年の小計を合計）
      let fileCount = 0;
      let thankYouCount = 0;
      
      Object.values(gradeSummary).forEach(gradeData => {
        fileCount += (gradeData.fileTotal || 0);
        thankYouCount += (gradeData.thankYouTotal || 0);
      });
      
      sheet.getRange(currentRow, 13).setValue(fileCount);     // M列：ファイル数全体合計
      sheet.getRange(currentRow, 15).setValue(thankYouCount); // O列：お礼状数全体合計
      
      // Q列の全体合計（各学年の小計行のQ列値を合計）
      let qColumnGrandTotal = 0;
      
      // 各学年の小計行のQ列値を取得して合計
      const grades = ['1年', '2年', '3年'];
      grades.forEach(grade => {
        // 該当学年の小計行を探してQ列の値を取得
        for (let row = 6; row < currentRow; row++) {
          const gradeValue = sheet.getRange(row, 2).getValue(); // B列：学年
          const cValue = sheet.getRange(row, 3).getValue();     // C列：内容
          
          if (gradeValue === grade && cValue === `${grade}小計`) {
            const qValue = sheet.getRange(row, 17).getValue(); // Q列（17列目）
            if (qValue && !isNaN(qValue)) {
              qColumnGrandTotal += Number(qValue);
              console.log(`${grade}小計行のQ列値:`, qValue);
            }
            break; // 該当学年の小計行を見つけたら次の学年へ
          }
        }
      });
      
      sheet.getRange(currentRow, 17).setValue(qColumnGrandTotal); // Q列：全体合計
      
      console.log('Q列全体合計計算完了（各学年小計の合計）:', qColumnGrandTotal);
      console.log('Q列全体合計を設定した行番号:', currentRow);
      
      console.log('全体合計行の追加完了。行番号:', currentRow, '内容:', grandTotal);
      
    } catch (error) {
      console.error('全体合計行の追加エラー:', error);
      console.error('エラースタック:', error.stack);
    }
    
    console.log('返礼品一覧シートの更新が完了しました');
    
  } catch (error) {
    console.error('新しい集計ロジックでの更新エラー:', error);
    throw error;
  }
}

// 既存の集計行をクリア
function clearGiftSummaryRows(sheet) {
  try {
    console.log('既存の集計行のクリア開始');
    
    // シートの最終行を取得
    const lastRow = sheet.getLastRow();
    if (lastRow < 6) {
      console.log('データ行が6行未満のため、クリアする集計行がありません');
      return;
    }
    
    // 6行目以降から集計行を探して削除
    let rowToCheck = 6;
    let deletedRows = 0;
    
    while (rowToCheck <= lastRow) {
      const cValue = sheet.getRange(rowToCheck, 3).getValue(); // C列の値
      
      // 集計行かどうかを判定（「小計」や「全体合計」が含まれている行）
      if (cValue && typeof cValue === 'string' && 
          (cValue.includes('小計') || cValue.includes('全体合計'))) {
        
        console.log(`集計行を削除: 行${rowToCheck}, 内容: ${cValue}`);
        sheet.deleteRow(rowToCheck);
        deletedRows++;
        
        // 行を削除したので、lastRowを更新
        const newLastRow = sheet.getLastRow();
        if (newLastRow < lastRow) {
          lastRow = newLastRow;
        }
      } else {
        rowToCheck++;
      }
    }
    
    console.log(`集計行のクリア完了: ${deletedRows}行を削除`);
    
  } catch (error) {
    console.error('集計行クリアエラー:', error);
    throw error;
  }
}

// 返礼品一覧シートのヘッダーを更新
function updateGiftSheetHeaders(sheet, giftSettings) {
  // 既存のヘッダーをクリア
  sheet.clear();
  
  // 新しいヘッダーを作成
  const headers = ['No.', '学年', '生徒氏名'];
  
  // 返礼品の列を追加
  giftSettings.forEach(gift => {
    headers.push(gift.name);
  });
  
  // クリアファイル列を追加
  headers.push('クリアファイル');
  
  // ヘッダーを設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

// 設定用ブックの返礼品設定を返礼品一覧シートの2行目に反映
function syncGiftSettingsToSheet(sheet) {
  try {
    console.log('返礼品設定を返礼品一覧シートに同期開始');
    
    // 設定用ブックから返礼品設定を取得
    const settingsBook = SpreadsheetApp.getActiveSpreadsheet();
    const giftSettingsSheet = settingsBook.getSheetByName('返礼品設定');
    
    if (!giftSettingsSheet) {
      console.log('返礼品設定シートが見つかりません');
      return;
    }
    
    // 返礼品設定から有効な返礼品を取得
    const giftData = giftSettingsSheet.getDataRange().getValues();
    const activeGifts = [];
    
    // ヘッダー行をスキップして2行目から処理
    for (let i = 1; i < giftData.length; i++) {
      const row = giftData[i];
      const id = row[0];
      const name = row[1];
      const isActive = row[3]; // D列：有効/無効
      
      if (isActive && id && name) {
        activeGifts.push({ id: id, name: name });
      }
    }
    
    console.log('有効な返礼品設定:', activeGifts);
    
    // 返礼品一覧シートの2行目に設定値を反映
    if (activeGifts.length > 0) {
      // 2行目が存在しない場合は作成
      if (sheet.getLastRow() < 2) {
        sheet.insertRowAfter(1);
      }
      
      // 設定列に返礼品名を設定
      const configRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 設定列の位置を特定（D, F, H, J, L, N列）- 一つおきに設定
      const configColumns = [3, 5, 7, 9, 11, 13]; // 0-based index (D, F, H, J, L, N列)
      
      // 各設定列に返礼品名を設定
      configColumns.forEach((colIndex, giftIndex) => {
        if (giftIndex < activeGifts.length) {
          sheet.getRange(2, colIndex + 1).setValue(activeGifts[giftIndex].name);
          console.log(`列${colIndex + 1}（${String.fromCharCode(65 + colIndex)}列）に「${activeGifts[giftIndex].name}」を設定`);
        }
      });
      
      console.log('返礼品設定の同期が完了しました');
    }
    
  } catch (error) {
    console.error('返礼品設定の同期エラー:', error);
    throw error;
  }
}

// 返礼品一覧シートのデータを更新
function updateGiftSheetData(sheet, page1DataSheet, page2DataSheet, giftSettings) {
  // 1ページ目データから氏名と学年を取得
  const page1Data = page1DataSheet.getDataRange().getValues();
  const nameGradeMap = {};
  
  // ヘッダー行をスキップして2行目から処理
  for (let i = 1; i < page1Data.length; i++) {
    const row = page1Data[i];
    const name = row[1]; // B列：氏名
    const grade = row[2]; // C列：学年
    
    if (name) {
      nameGradeMap[name] = grade;
    }
  }
  
  // 2ページ目データから返礼品を集計
  const page2Data = page2DataSheet.getDataRange().getValues();
  const giftSummary = {};
  
  // ヘッダー行をスキップして2行目から処理
  for (let i = 1; i < page2Data.length; i++) {
    const row = page2Data[i];
    const page1Name = row[1]; // B列：1ページ目氏名
    const gift = row[4]; // E列：返礼品（名前）
    const sheets = parseInt(row[5]) || 0; // F列：クリアファイル枚数
    
    if (page1Name && nameGradeMap[page1Name]) {
      if (!giftSummary[page1Name]) {
        giftSummary[page1Name] = {
          grade: nameGradeMap[page1Name],
          gifts: {},
          clearfileTotal: 0
        };
      }
      
      // 返礼品のカウント（名前で集計）
      if (gift && gift !== 'クリアファイル') {
        giftSummary[page1Name].gifts[gift] = (giftSummary[page1Name].gifts[gift] || 0) + 1;
        // クリアファイルは1枚つく
        giftSummary[page1Name].clearfileTotal += 1;
      } else if (gift === 'クリアファイル') {
        // クリアファイルを選択した場合は選択した枚数
        giftSummary[page1Name].clearfileTotal += sheets;
      }
    }
  }
  
  // シートにデータを書き込み
  const newData = [];
  let rowNumber = 1;
  
  Object.keys(giftSummary).forEach(name => {
    const summary = giftSummary[name];
    const row = [rowNumber, summary.grade, name];
    
    // 各返礼品の数を追加（名前でマッチング）
    giftSettings.forEach(gift => {
      row.push(summary.gifts[gift.name] || 0);
    });
    
    // クリアファイルの合計を追加
    row.push(summary.clearfileTotal);
    
    newData.push(row);
    rowNumber++;
  });
  
  // データを書き込み（ヘッダー行の下から）
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
  }
}



// OPTIONSリクエストに対応（CORSプリフライトリクエスト用）
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// テスト用関数（Google Apps Scriptエディタで実行可能）
function testGetSettings() {
  try {
    const result = getSettings();
    console.log('テスト結果:', result.getContent());
    return result;
  } catch (error) {
    console.error('テストエラー:', error);
    throw error;
  }
}

// テスト用関数（スプレッドシート情報の確認）
function testSpreadsheetInfo() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('スプレッドシート名:', sheet.getName());
    console.log('スプレッドシートID:', sheet.getId());
    
    const sheetNames = sheet.getSheets().map(s => s.getName());
    console.log('利用可能なシート:', sheetNames);
    
    return {
      name: sheet.getName(),
      id: sheet.getId(),
      sheets: sheetNames
    };
  } catch (error) {
    console.error('テストエラー:', error);
    throw error;
  }
}

// テスト用関数（フォームデータの構造を確認）
function testFormDataStructure() {
  // サンプルデータを作成
  const sampleFormData = {
    basicInfo: {
      basicInfo: {
        name: "テスト太郎",
        grade: "3年",
        transferDate: "2024-01-15"
      },
      breakdownDetails: {
        transferAmount: "50000",
        donationAmount: 30000,
        burdenFee: "1",
        hasEncouragementFee: true,
        tshirtAmount: 2000
      }
    },
    donors: [
      {
        name: "寄付者1",
        amount: 10000,
        gift: "towel",
        sheets: null,
        note: "テスト"
      },
      {
        name: "寄付者2", 
        amount: 5000,
        gift: "clearfile",
        sheets: 3,
        note: ""
      }
    ]
  };
  
  console.log('テスト用フォームデータ:', JSON.stringify(sampleFormData, null, 2));
  
  // 各関数をテスト
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1ページ目データ整理シートをテスト
    const page1Sheet = sheet.getSheetByName('1ページ目データ整理');
    if (page1Sheet) {
      console.log('1ページ目データ整理シートをテスト中...');
      addPage1Data(page1Sheet, sampleFormData);
      console.log('1ページ目データ整理シートのテスト完了');
    }
    
    // 2ページ目データ整理シートをテスト
    const page2Sheet = sheet.getSheetByName('2ページ目データ整理');
    if (page2Sheet) {
      console.log('2ページ目データ整理シートをテスト中...');
      addPage2Data(page2Sheet, sampleFormData);
      console.log('2ページ目データ整理シートのテスト完了');
    }
    
    console.log('すべてのテスト完了');
    
  } catch (error) {
    console.error('テストエラー:', error);
    throw error;
  }
}

// テスト用関数（実際のフォーム送信をシミュレート）
function testActualFormSubmission() {
  // 実際のフォーム送信をシミュレート
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        action: 'submitFormData',
        formData: {
          basicInfo: {
            basicInfo: {
              name: "実際の太郎",
              grade: "2年",
              transferDate: "2024-08-05"
            },
            breakdownDetails: {
              transferAmount: "75000",
              donationAmount: 45000,
              burdenFee: "2",
              hasEncouragementFee: true,
              tshirtAmount: 3000
            }
          },
          donors: [
            {
              name: "実際の寄付者1",
              amount: 15000,
              gift: "sweets_large",
              sheets: null,
              note: "実際のテスト"
            },
            {
              name: "実際の寄付者2",
              amount: 8000,
              gift: "clearfile",
              sheets: 5,
              note: ""
            }
          ]
        }
      })
    }
  };
  
  console.log('実際のフォーム送信をシミュレート中...');
  console.log('mockEvent:', JSON.stringify(mockEvent, null, 2));
  
  try {
    const result = doPost(mockEvent);
    console.log('doPost実行結果:', result.getContent());
  } catch (error) {
    console.error('doPost実行エラー:', error);
  }
}

// 個票自動生成機能
// 個票ブックの取得（固定ID版）
function createIndividualReportBook() {
  try {
    // 「フォーム設定」シートから individualBookId/url を取得（未設定時は既存固定IDにフォールバック）
    const active = SpreadsheetApp.getActiveSpreadsheet();
    let configuredBookId = '';
    try {
      const formSettingsSheet = active.getSheetByName('フォーム設定');
      if (formSettingsSheet) {
        const values = formSettingsSheet.getDataRange().getValues();
        for (let i = 1; i < values.length; i++) {
          const key = values[i][0];
          const val = values[i][1];
          const enabled = values[i][2];
          const k = String(key).trim();
          if (enabled && (k === 'individualBookId' || k === 'individualBookUrl') && val) {
            const idOrUrl = String(val).trim();
            configuredBookId = extractSpreadsheetId(idOrUrl);
            break;
          }
        }
      }
    } catch (e) {
      console.warn('フォーム設定の読込時警告:', e);
    }

    const fallbackBookId = '1ycm993MubqVBt4WuB3gO5kcW4lOddenQvvfSlRrsNpY';
    const bookIdToOpen = configuredBookId || fallbackBookId;
    const individualBook = SpreadsheetApp.openById(bookIdToOpen);
    console.log((configuredBookId ? '設定の個票ブックIDを使用:' : 'フォールバック個票ブックIDを使用:'), individualBook.getName());
    return individualBook;
  } catch (error) {
    console.error('個票ブック取得エラー:', error);
    throw error;
  }
}

// テンプレートから個票シートを作成（同一ブック内の個票_テンプレートをコピー）
function createIndividualSheetFromTemplate(targetBook, sheetName) {
  try {
    // まず対象ブック内の「個票_テンプレート」を探す
    let templateSheet = targetBook.getSheetByName('個票_テンプレート');
    
    // 見つからない場合は、テンプレート用スプレッドシート（固定ID）からコピー
    if (!templateSheet) {
      const fallbackTemplateId = '1IZeLmRAYArRUyvenKJJANLwf56KceTsDeJAcM-OhJ7A';
      const templateBook = SpreadsheetApp.openById(fallbackTemplateId);
      templateSheet = templateBook.getSheets()[0];
      if (!templateSheet) {
        throw new Error('テンプレートが見つかりません（個票_テンプレート もしくは テンプレート用スプレッドシート）');
      }
    }
    const copiedSheet = templateSheet.copyTo(targetBook);
    copiedSheet.setName(sheetName);
    // コピー直後に末尾へ移動して見つけやすくする
    targetBook.setActiveSheet(copiedSheet);
    targetBook.moveActiveSheet(targetBook.getNumSheets());
    console.log('個票シート作成完了:', sheetName);
    return copiedSheet;
  } catch (error) {
    console.error('テンプレートコピーエラー:', error);
    throw error;
  }
}

// 個票シートにデータを設定（新テンプレート対応）
function setIndividualReportData(sheet, studentData, donorsData) {
  try {
    console.log('個票データ設定開始:', studentData.name);
    console.log('生徒データ:', studentData);
    console.log('寄付者データ件数:', donorsData.length);
    console.log('寄付者データ詳細:', donorsData);
    
    // 基本情報を設定
    sheet.getRange('D1').setValue(studentData.grade);        // 学年
    sheet.getRange('D2').setValue(studentData.name);         // 生徒名
    sheet.getRange('H1').setValue(studentData.transferDate); // 振込日
    sheet.getRange('H2').setValue('¥' + studentData.transferAmount.toLocaleString()); // 振込金額（¥マーク付き）
    
    // 寄付者データを設定（B7から開始）
    let currentRow = 7;
    let totalDonation = 0;
    
    for (let i = 0; i < donorsData.length && currentRow <= 21; i++) { // 最大15行（7-21行）
      const donor = donorsData[i];
      
      sheet.getRange(`B${currentRow}`).setValue(donor.name);   // 御芳名
      sheet.getRange(`C${currentRow}`).setValue('¥' + donor.amount.toLocaleString()); // 金額（¥マーク付き）
      
              // 返礼品の設定（ID→名前の変換、クリアファイルの場合は枚数も表示）
        let giftDisplay = '';
        
        // 返礼品IDから名前への変換
        if (donor.gift) {
          // 設定用ブックから返礼品設定を取得
          let giftName = donor.gift; // デフォルトはIDのまま
          
          try {
            // 設定用ブック（フォーム設定で指定されたブック）から返礼品設定を取得
            const configBook = getConfigBook();
            if (configBook) {
              const giftSettingsSheet = configBook.getSheetByName('返礼品設定');
              if (giftSettingsSheet) {
                const giftData = giftSettingsSheet.getDataRange().getValues();
                console.log(`返礼品設定シート取得: ${giftData.length}行`);
                
                // ヘッダー行をスキップして2行目から処理
                for (let i = 1; i < giftData.length; i++) {
                  const row = giftData[i];
                  const id = row[0];
                  const name = row[1];
                  const isActive = row[4]; // 有効/無効の列はE列（5列目）
                  
                  console.log(`返礼品設定行${i + 1}: ID="${id}", 名前="${name}", 有効="${isActive}"`);
                  
                  if (isActive && id === donor.gift && name) {
                    giftName = name;
                    console.log(`返礼品名変換: "${donor.gift}" → "${giftName}"`);
                    break;
                  }
                }
              } else {
                console.log('返礼品設定シートが見つかりません');
              }
            } else {
              console.log('設定用ブックが取得できません');
            }
          } catch (e) {
            console.log('返礼品設定取得時のエラー:', e);
          }
          
          // クリアファイルの場合は枚数も追加（返礼品名に「クリアファイル」が含まれている場合）
          if (donor.gift === 'clearfile' || (giftName && giftName.includes('クリアファイル'))) {
            console.log(`クリアファイル処理開始: 返礼品ID="${donor.gift}", 変換後名前="${giftName}", 枚数="${donor.clearFileCount}"`);
            
            // 枚数が0の場合も含めて、常に枚数を表示
            const clearFileCount = donor.clearFileCount || 0;
            giftDisplay = `${giftName}(${clearFileCount}枚)`;
            console.log(`クリアファイル枚数付き表示: ${giftDisplay}`);
          } else {
            giftDisplay = giftName;
            console.log(`通常返礼品表示: ${giftDisplay}`);
          }
        }
      
      sheet.getRange(`D${currentRow}`).setValue(giftDisplay);  // 返礼品
      
      // デバッグログ
      console.log(`個票行${currentRow}: 返礼品ID="${donor.gift}", 表示名="${giftDisplay}", クリアファイル枚数=${donor.clearFileCount}`);
      
      totalDonation += donor.amount;
      currentRow++;
    }
    
    // 集計データを設定（¥マーク付き）
    sheet.getRange('I6').setValue('¥' + totalDonation.toLocaleString());                    // 寄付金合計
    sheet.getRange('I7').setValue('¥' + (totalDonation - studentData.directTransferAmount).toLocaleString()); // 金額（寄付金合計 - 直接振込金額）
    // その他の費用（数値に正規化してから出力）
    const otherFeeValue = Number(studentData.otherFee) || 0;
    console.log('個票: その他の費用（数値化後）:', otherFeeValue);
    sheet.getRange('I12').setValue('¥' + otherFeeValue.toLocaleString());      // その他の費用
    sheet.getRange('I8').setValue('¥' + studentData.directTransferAmount.toLocaleString()); // 直接振込金額
    
    // 負担金人数と金額を設定（人数 × 設定した負担金単価）
    const burdenFeeCount = studentData.burdenFeeCount;
    const unitAmount = getBurdenFeeUnitAmount();
    const burdenTotal = burdenFeeCount > 0 ? burdenFeeCount * unitAmount : 0;
    sheet.getRange('H9').setValue(`${burdenFeeCount > 0 ? burdenFeeCount : 0}名`);
    sheet.getRange('I9').setValue('¥' + burdenTotal.toLocaleString());
    
    sheet.getRange('I10').setValue('¥' + (studentData.tshirtAmount || 0).toLocaleString());        // Tシャツ代（¥マーク付き）
    sheet.getRange('I11').setValue('¥' + (studentData.encouragementFee || 0).toLocaleString());   // 激励会負担金（¥マーク付き）
    
    console.log('個票データ設定完了:', studentData.name);
    
  } catch (error) {
    console.error('個票データ設定エラー:', error);
    throw error;
  }
}



// 設定用ブックを取得する関数
function getConfigBook() {
  try {
    // 年次の集計用ブックからフォーム設定を取得
    const aggregationBook = getAggregationBook();
    const formSettingsSheet = aggregationBook.getSheetByName('フォーム設定');
    
    if (formSettingsSheet) {
      const values = formSettingsSheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        const key = values[i][0];
        const val = values[i][1];
        const enabled = values[i][2];
        
        if (enabled && key === 'configBookId' && val) {
          const configBookId = String(val).trim();
          console.log('設定用ブックIDを取得:', configBookId);
          return SpreadsheetApp.openById(configBookId);
        }
      }
    }
    
    // フォールバック: デフォルトの設定用ブックID
    const fallbackConfigId = '1ycm993MubqVBt4WuB3gO5kcW4lOddenQvvfSlRrsNpY';
    console.log('フォールバック設定用ブックIDを使用:', fallbackConfigId);
    return SpreadsheetApp.openById(fallbackConfigId);
    
  } catch (error) {
    console.error('設定用ブック取得エラー:', error);
    return null;
  }
}

// 個票シートを生成
function generateIndividualReport(studentName) {
  try {
    // 年次の集計用ブックからデータを取得
    const aggregationBook = getAggregationBook();
    console.log('集計用ブック名:', aggregationBook.getName());
    console.log('集計用ブックID:', aggregationBook.getId());
    
    const page1Sheet = aggregationBook.getSheetByName('1ページ目データ整理');
    const page2Sheet = aggregationBook.getSheetByName('2ページ目データ整理');
    
    console.log('1ページ目データ整理シート:', page1Sheet ? '存在' : '不存在');
    console.log('2ページ目データ整理シート:', page2Sheet ? '存在' : '不存在');
    
    if (!page1Sheet || !page2Sheet) {
      console.log('必要なシートが見つかりません');
      console.log('利用可能なシート:', aggregationBook.getSheets().map(s => s.getName()));
      return;
    }
    
    // 1ページ目データから該当する生徒の情報を取得
    const page1Data = page1Sheet.getDataRange().getValues();
    let studentInfo = null;
    
    for (let i = 1; i < page1Data.length; i++) {
      if (page1Data[i][1] === studentName) { // 氏名列（B列）
        // V列から兄弟の人数を取得（フォームで選択された人数）
        const siblingCount = parseInt(page1Data[i][21]) || 0; // V列（22列目）= 21 (0-based)
        console.log(`行${i + 2}の兄弟の人数:`, siblingCount, '列位置: V列(22列目)');
        
        // 負担金の人数を計算: 1 + 兄弟の数（本人を除く）
        const burdenFeeCount = 1 + siblingCount;
        console.log(`最終的な負担金人数: ${burdenFeeCount}`);
        
        studentInfo = {
          name: studentName,
          grade: page1Data[i][2], // C列：学年
          transferDate: page1Data[i][3], // D列：振込日
          transferAmount: page1Data[i][4], // E列：振込金額
          directTransferAmount: page1Data[i][5], // F列：直接振込金額
          tshirtAmount: page1Data[i][8], // I列：Tシャツ代
          encouragementFee: page1Data[i][9], // J列：奨励会負担金
          otherFee: page1Data[i][10] || 0, // K列：その他の費用
          burdenFeeCount: burdenFeeCount
        };
        break;
      }
    }
    
    if (!studentInfo) {
      console.log('生徒情報が見つかりません:', studentName);
      return;
    }
    
    // 2ページ目データから該当する生徒の寄付者情報を取得
    const page2Data = page2Sheet.getDataRange().getValues();
    const donors = [];
    
    console.log(`2ページ目データ検索開始: 生徒名="${studentName}"`);
    console.log(`2ページ目データ総行数: ${page2Data.length}`);
    
    // ヘッダー行の情報を出力（列位置の確認用）
    if (page2Data.length > 0) {
      const headers = page2Data[0];
      console.log('2ページ目ヘッダー行:', {
        'A列(0)': headers[0],
        'B列(1)': headers[1],
        'C列(2)': headers[2],
        'D列(3)': headers[3],
        'E列(4)': headers[4],
        'F列(5)': headers[5],
        'G列(6)': headers[6]
      });
    }
    
    for (let i = 1; i < page2Data.length; i++) {
      const rowName = page2Data[i][1]; // B列：1ページ目氏名
      if (rowName === studentName) {
        // クリアファイル枚数を確実に数値として取得
        let clearFileCount = 0;
        const rawClearFileCount = page2Data[i][5]; // F列：クリアファイル枚数
        
        if (rawClearFileCount !== null && rawClearFileCount !== undefined && rawClearFileCount !== '') {
          if (typeof rawClearFileCount === 'number') {
            clearFileCount = rawClearFileCount;
          } else {
            const parsed = Number(String(rawClearFileCount).replace(/[^0-9]/g, ''));
            clearFileCount = isNaN(parsed) ? 0 : parsed;
          }
        }
        
        const donorData = {
          name: page2Data[i][2], // C列：御芳名
          amount: page2Data[i][3], // D列：寄付金額
          gift: page2Data[i][4], // E列：返礼品
          clearFileCount: clearFileCount,
          isDirectTransfer: page2Data[i][6] === '✓'
        };
        
        console.log(`寄付者データ行${i + 1}:`, {
          御芳名: donorData.name,
          寄付金額: donorData.amount,
          返礼品: donorData.gift,
          クリアファイル枚数: donorData.clearFileCount,
          直接振込: donorData.isDirectTransfer
        });
        
        donors.push(donorData);
      }
    }
    
    if (donors.length === 0) {
      console.log('寄付者データが見つかりません:', studentName);
      return;
    }
    
    console.log(`寄付者データ取得完了: ${donors.length}件`);
    console.log('取得した寄付者データ:', donors);
    
    // 「1ページ目データ整理」シートから負担金（ヘッダー名: 負担金）を取得（氏名=B列で特定）
    let burdenAmountFromSummary = 0;
    try {
      if (page1Sheet) {
        const lastCol = page1Sheet.getLastColumn();
        const lastRow = page1Sheet.getLastRow();
        if (lastCol > 0 && lastRow >= 2) {
          // 固定列で取得: B列=氏名(2), H列=負担金(8)
          const COL_NAME = 2;
          const COL_BURDEN = 8;
          const values = page1Sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
          // 最新（最下行）から検索して直近の負担金を採用
          for (let i = values.length - 1; i >= 0; i--) {
            const rowName = values[i][COL_NAME - 1];
            if (rowName && String(rowName).trim() === String(studentName).trim()) {
              const burdenRaw = values[i][COL_BURDEN - 1];
              let parsed = 0;
              if (typeof burdenRaw === 'number') {
                parsed = burdenRaw;
              } else if (burdenRaw) {
                parsed = Number(String(burdenRaw).replace(/[^0-9.-]/g, '')) || 0;
              }
              burdenAmountFromSummary = parsed;
              break;
            }
          }
        }
      }
    } catch (e) {
      console.log('負担金取得時の例外（1ページ目データ整理参照）:', e);
    }

    // 個票ブックを取得または作成
    const individualBook = createIndividualReportBook();
    
    // シート名を決定（既存の同名シートがある場合は日付を追加）
    let sheetName = studentName;
    const existingSheets = individualBook.getSheets();
    const existingSheetNames = existingSheets.map(s => s.getName());
    
    if (existingSheetNames.includes(sheetName)) {
      const today = new Date();
      const dateStr = today.getFullYear().toString() + 
                     (today.getMonth() + 1).toString().padStart(2, '0') + 
                     today.getDate().toString().padStart(2, '0');
      sheetName = `${studentName}_${dateStr}`;
    }
    
    // テンプレートから新しいシートを作成
    const newSheet = createIndividualSheetFromTemplate(individualBook, sheetName);
    
    // 個票にデータを設定（寄付金一覧の負担金を付与）
    studentInfo.burdenAmountFromSummary = burdenAmountFromSummary;
    setIndividualReportData(newSheet, studentInfo, donors);
    
    console.log('個票生成完了:', studentName);
    
  } catch (error) {
    console.error('個票生成エラー:', error);
    throw error;
  }
}

// 2ページ目データ整理シートの変更を監視して個票を自動生成
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // 2ページ目データ整理シートの変更のみを監視
    if (sheetName !== '2ページ目データ整理') {
      return;
    }
    
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    // 1ページ目氏名列（B列）の変更のみを監視
    if (col !== 2 || row === 1) { // ヘッダー行は除外
      return;
    }
    
    // 変更されたセルの値を取得
    const studentName = range.getValue();
    
    if (studentName && studentName.trim() !== '') {
      console.log('2ページ目データ整理シートに新しいデータが追加されました:', studentName);
      
      // 少し待ってから個票を生成（データの完全な保存を待つ）
      Utilities.sleep(1000);
      
      // 個票を生成
      generateIndividualReport(studentName);
    }
    
  } catch (error) {
    console.error('onEditエラー:', error);
  }
}

// メニューから「一番下に1行追加」を実行できるようにする
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('管理ツール')
      .addItem('一番下に1行追加', 'addOneRowToBottom')
      .addSeparator()
      .addItem('寄付金一覧を初期化', 'initializeDonationSummary')
      .addItem('寄付金一覧を集計', 'runDonationSummary')
      .addSeparator()
      .addItem('返礼品一覧を初期化', 'initializeGiftSummary')
      .addItem('返礼品一覧を集計', 'runGiftSummary')
      .addToUi();
  } catch (error) {
    console.error('onOpenメニュー設定エラー:', error);
  }
}

// アクティブシートの一番下に1行だけ追加（データ検証をテンプレート行から継承）
function addOneRowToBottom() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // 最終行の直下に1行挿入
    sheet.insertRowAfter(lastRow);
    const newRowIndex = lastRow + 1;

    // テンプレート行（基本: ヘッダー行の次=2行目）の検証を新規行にコピー
    if (lastRow >= 2 && lastCol > 0) {
      const templateRowIndex = 2;
      const templateRange = sheet.getRange(templateRowIndex, 1, 1, lastCol);
      const templateValidations = templateRange.getDataValidations();
      const newRange = sheet.getRange(newRowIndex, 1, 1, lastCol);
      newRange.setDataValidations(templateValidations);
    }

    // 新規行の内容をクリア（万一の初期値を除去）
    sheet.getRange(newRowIndex, 1, 1, lastCol).clearContent();

    console.log(`シート「${sheet.getName()}」に1行追加: 行${newRowIndex}`);
  } catch (error) {
    console.error('一番下に1行追加エラー:', error);
    throw error;
  }
}

// 手動で個票を生成する関数（テスト用）
function generateIndividualReportManually(studentName) {
  try {
    console.log('手動で個票を生成中:', studentName);
    generateIndividualReport(studentName);
  } catch (error) {
    console.error('手動個票生成エラー:', error);
    throw error;
  }
}

// フォーム送信時の個票生成プロセスをデバッグ
function debugFormSubmissionIndividualReport(studentName) {
  // 引数が指定されていない場合のデフォルト値
  if (!studentName) {
    studentName = "日本太郎";
  }
  
  try {
    console.log('=== フォーム送信個票生成デバッグ開始 ===');
    console.log('対象生徒名:', studentName);
    
    // 1. メインスプレッドシートの確認
    console.log('1. メインスプレッドシート確認');
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('メインスプレッドシート名:', mainSheet.getName());
    console.log('メインスプレッドシートID:', mainSheet.getId());
    
    // 2. 1ページ目データの確認
    console.log('2. 1ページ目データ確認');
    const page1Sheet = mainSheet.getSheetByName('1ページ目データ整理');
    if (page1Sheet) {
      const page1Data = page1Sheet.getDataRange().getValues();
      console.log('1ページ目データ行数:', page1Data.length);
      
      // 該当する生徒のデータを探す
      let studentFound = false;
      for (let i = 1; i < page1Data.length; i++) {
        if (page1Data[i][1] === studentName) { // B列：氏名
          console.log(`生徒データ発見 (行${i + 1}):`, page1Data[i]);
          studentFound = true;
          break;
        }
      }
      if (!studentFound) {
        console.log('該当する生徒のデータが1ページ目に見つかりません');
      }
    } else {
      console.log('1ページ目データ整理シートが見つかりません');
    }
    
    // 3. 2ページ目データの確認
    console.log('3. 2ページ目データ確認');
    const page2Sheet = mainSheet.getSheetByName('2ページ目データ整理');
    if (page2Sheet) {
      const page2Data = page2Sheet.getDataRange().getValues();
      console.log('2ページ目データ行数:', page2Data.length);
      
      // 該当する生徒の寄付者データを探す
      const studentDonors = [];
      for (let i = 1; i < page2Data.length; i++) {
        if (page2Data[i][1] === studentName) { // B列：1ページ目氏名
          studentDonors.push(page2Data[i]);
        }
      }
      console.log(`生徒の寄付者データ数: ${studentDonors.length}`);
      studentDonors.forEach((donor, index) => {
        console.log(`寄付者${index + 1}:`, donor);
      });
      
      if (studentDonors.length === 0) {
        console.log('該当する生徒の寄付者データが2ページ目に見つかりません');
      }
    } else {
      console.log('2ページ目データ整理シートが見つかりません');
    }
    
    // 4. 個票ブックへのアクセス確認
    console.log('4. 個票ブックアクセス確認');
    try {
      const individualBook = createIndividualReportBook();
      console.log('個票ブック取得成功:', individualBook.getName());
      console.log('個票ブック書き込み権限確認...');
      
      // テスト用セルに書き込みテスト
      const testSheet = individualBook.getSheetByName('個票_テンプレート');
      if (testSheet) {
        const originalValue = testSheet.getRange('A1').getValue();
        testSheet.getRange('A1').setValue('テスト書き込み');
        console.log('書き込みテスト成功');
        testSheet.getRange('A1').setValue(originalValue); // 元に戻す
      }
    } catch (error) {
      console.error('個票ブックアクセスエラー:', error);
    }
    
    // 5. 実際の個票生成テスト
    console.log('5. 個票生成実行テスト');
    try {
      generateIndividualReport(studentName);
      console.log('個票生成テスト成功');
    } catch (error) {
      console.error('個票生成テストエラー:', error);
      console.error('エラースタック:', error.stack);
    }
    
    console.log('=== フォーム送信個票生成デバッグ終了 ===');
    
  } catch (error) {
    console.error('デバッグ全体エラー:', error);
    throw error;
  }
}

// デバッグ用：個票生成の詳細ログ出力
function debugIndividualReportGeneration() {
  try {
    console.log('=== 個票生成デバッグ開始 ===');
    
    // 1. 個票ブックの取得テスト
    console.log('1. 個票ブック取得テスト');
    const individualBook = createIndividualReportBook();
    console.log('個票ブック名:', individualBook.getName());
    console.log('個票ブックID:', individualBook.getId());
    
    // 2. シート一覧の確認
    console.log('2. 個票ブック内のシート一覧:');
    const sheets = individualBook.getSheets();
    sheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. ${sheet.getName()}`);
    });
    
    // 3. テンプレートシートの確認
    console.log('3. テンプレートシート確認');
    const templateSheet = individualBook.getSheetByName('個票_テンプレート');
    if (templateSheet) {
      console.log('個票_テンプレート シートが見つかりました');
    } else {
      console.log('個票_テンプレート シートが見つかりません');
    }
    
    // 4. メインスプレッドシートのデータ確認
    console.log('4. メインスプレッドシートのデータ確認');
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    const page1Sheet = mainSheet.getSheetByName('1ページ目データ整理');
    const page2Sheet = mainSheet.getSheetByName('2ページ目データ整理');
    
    if (page1Sheet) {
      const page1Data = page1Sheet.getDataRange().getValues();
      console.log('1ページ目データ行数:', page1Data.length);
      if (page1Data.length > 1) {
        console.log('1ページ目サンプルデータ:', page1Data[1]);
      }
    } else {
      console.log('1ページ目データ整理シートが見つかりません');
    }
    
    if (page2Sheet) {
      const page2Data = page2Sheet.getDataRange().getValues();
      console.log('2ページ目データ行数:', page2Data.length);
      if (page2Data.length > 1) {
        console.log('2ページ目サンプルデータ:', page2Data[1]);
      }
    } else {
      console.log('2ページ目データ整理シートが見つかりません');
    }
    
    console.log('=== 個票生成デバッグ終了 ===');
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    throw error;
  }
}

// 全生徒の個票を一括生成する関数
function generateAllIndividualReports() {
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    const page1Sheet = mainSheet.getSheetByName('1ページ目データ整理');
    
    if (!page1Sheet) {
      console.log('1ページ目データ整理シートが見つかりません');
      return;
    }
    
    const page1Data = page1Sheet.getDataRange().getValues();
    const students = new Set();
    
    // 1ページ目データから生徒名を収集
    for (let i = 1; i < page1Data.length; i++) {
      const studentName = page1Data[i][1]; // 氏名列
      if (studentName && studentName.trim() !== '') {
        students.add(studentName);
      }
    }
    
    console.log('個票生成対象生徒数:', students.size);
    
    // 各生徒の個票を生成
    for (const studentName of students) {
      try {
        generateIndividualReport(studentName);
        Utilities.sleep(500); // 少し待機
      } catch (error) {
        console.error(`個票生成エラー (${studentName}):`, error);
      }
    }
    
    console.log('全個票生成完了');
    
  } catch (error) {
    console.error('全個票生成エラー:', error);
    throw error;
  }
}

// 寄付金一覧シートの集計を実行
function runDonationSummary() {
  try {
    console.log('寄付金一覧集計を開始');
    
    // 共通の集計処理を使用
    executeDonationSummary();
    
    // 成功メッセージを表示
    SpreadsheetApp.getUi().alert('完了', '寄付金一覧の集計が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('寄付金一覧集計エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `集計中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 返礼品一覧シートの集計を実行
function runGiftSummary() {
  try {
    console.log('返礼品一覧集計を開始');
    
    // 共通の集計処理を使用
    executeGiftSummary();
    
    // 成功メッセージを表示
    SpreadsheetApp.getUi().alert('完了', '返礼品一覧の集計（学年別・全体集計含む）が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('返礼品一覧集計エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `集計中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 返礼品一覧シートの初期化
function initializeGiftSummary() {
  try {
    console.log('返礼品一覧初期化を開始');
    
    // 集計用ブックを取得
    const aggregationBook = getAggregationBook();
    const giftSheet = aggregationBook.getSheetByName('返礼品一覧');
    
    if (!giftSheet) {
      console.error('返礼品一覧シートが見つかりません');
      SpreadsheetApp.getUi().alert('エラー', '返礼品一覧シートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 返礼品一覧シートの初期化
    initializeGiftSummarySheet(giftSheet);
    
    // 成功メッセージを表示
    SpreadsheetApp.getUi().alert('完了', '返礼品一覧シートの初期化が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('返礼品一覧初期化エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `初期化中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 返礼品一覧シートの初期化（6行目以降のデータ行のみ削除）
function initializeGiftSummarySheet(sheet) {
  try {
    console.log('返礼品一覧シートの初期化を開始（6行目以降のデータ行のみ削除）');
    
    // 6行目以降のデータ行のみを削除（1-5行目のテンプレートは保持）
    const lastRow = sheet.getLastRow();
    if (lastRow > 5) {
      try {
        const clearRange = sheet.getRange(6, 1, lastRow - 5, sheet.getLastColumn());
        clearRange.clearContent();
        clearRange.clearFormat();
        console.log(`6行目以降の${lastRow - 5}行をクリアしました`);
      } catch (error) {
        console.error('データ行クリアエラー:', error);
        throw error;
      }
    } else {
      console.log('6行目以降のデータ行がないため、クリア処理をスキップしました');
    }
    
    // 返礼品設定を返礼品一覧シートの2行目に同期（設定が消えている場合の復旧）
    try {
      syncGiftSettingsToSheet(sheet);
      console.log('返礼品設定の同期を実行しました');
    } catch (error) {
      console.warn('返礼品設定の同期でエラーが発生しましたが、処理を続行します:', error);
    }
    
    console.log('返礼品一覧シートの初期化が完了しました（テンプレートは保持）');
    
  } catch (error) {
    console.error('返礼品一覧シートの初期化エラー:', error);
    throw error;
  }
}

// 寄付金一覧シートの初期化（ヘッダー設定）
function initializeDonationSummarySheet(sheet) {
  try {
    console.log('寄付金一覧シートの初期化を開始');
    
    // ヘッダー行を設定
    const headers = [
      '氏名', '学年', '振込日', '振込金額', '寄付金\n*直接振込を除く', 
      '直接振込金額', '負担金', 'Tシャツ代', '激励会負担金', 'その他の費用', 
      '寄付・負担金計', '備考', 'チェック'
    ];
    
    // ヘッダーを設定
    try {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } catch (error) {
      console.error('ヘッダー行の設定エラー:', error);
      throw error;
    }
    
    // ヘッダー行のスタイルを設定
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // 適切な列幅を設定（ピクセル単位）
    const columnWidths = [
      120,  // A列：氏名
      50,   // B列：学年（最小化）
      100,  // C列：振込日
      100,  // D列：振込金額
      120,  // E列：寄付金*直接振込を除く（縮小）
      120,  // F列：直接振込金額
      80,   // G列：負担金
      80,   // H列：Tシャツ代
      120,  // I列：激励会負担金
      100,  // J列：その他の費用
      120,  // K列：寄付・負担金計
      150,  // L列：備考
      80    // M列：チェック
    ];
    
    // 各列の幅を設定
    columnWidths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    
    console.log('寄付金一覧シートの初期化が完了しました');
    
  } catch (error) {
    console.error('寄付金一覧シートの初期化エラー:', error);
    throw error;
  }
}

// 寄付金一覧シートの初期化を実行
function initializeDonationSummary() {
  try {
    console.log('寄付金一覧初期化を開始');
    
    // 集計用ブックを取得
    const aggregationBook = getAggregationBook();
    const donationSheet = aggregationBook.getSheetByName('寄付金一覧');
    
    if (!donationSheet) {
      console.error('寄付金一覧シートが見つかりません');
      SpreadsheetApp.getUi().alert('エラー', '寄付金一覧シートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 初期化を実行
    initializeDonationSummarySheet(donationSheet);
    
    // 成功メッセージを表示
    SpreadsheetApp.getUi().alert('完了', '寄付金一覧シートの初期化が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('寄付金一覧初期化エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `初期化中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 条件付き書式を適用する関数
function applyConditionalFormatting(sheet) {
  try {
    console.log('条件付き書式の適用を開始');
    
    // データ行の範囲を取得（ヘッダー行を除く）
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('データ行がないため条件付き書式を適用しません');
      return;
    }
    
    console.log(`条件付き書式適用対象: 最終行${lastRow}`);
    
    // 既存の条件付き書式をクリア
    try {
      sheet.clearConditionalFormatRules();
      console.log('既存の条件付き書式をクリアしました');
    } catch (error) {
      console.error('既存条件付き書式クリアエラー:', error);
    }
    
    // データ行の範囲（ヘッダー行を除く）
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    console.log(`条件付き書式適用範囲: 行2, 列1, 行数${lastRow - 1}, 列数${sheet.getLastColumn()}`);
    
    // 条件付き書式ルールを作成
    const rules = [];
    
    // ルール1: E列とF列の合計が50,000未満の場合（黄色背景）
    try {
      const rule1 = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($E2<>"",$F2<>"",$E2+$F2<50000)')
        .setBackground('#FFFF99') // 薄い黄色
        .setRanges([sheet.getRange(2, 5, lastRow - 1, 2)]) // E列とF列のみ
        .build();
      
      rules.push(rule1);
      console.log('ルール1（金額条件）を作成しました');
    } catch (error) {
      console.error('ルール1作成エラー:', error);
    }
    
    // ルール2: 小計行や全体合計行を除外（A列に「小計」や「合計」が含まれていない行のみ）
    try {
      const rule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($E2<>"",$F2<>"",$E2+$F2<50000,NOT(ISNUMBER(SEARCH("小計",$A2)),NOT(ISNUMBER(SEARCH("合計",$A2))))')
        .setBackground('#FFFF99') // 薄い黄色
        .setRanges([sheet.getRange(2, 5, lastRow - 1, 2)]) // E列とF列のみ
        .build();
      
      rules.push(rule2);
      console.log('ルール2（除外条件付き）を作成しました');
    } catch (error) {
      console.error('ルール2作成エラー:', error);
    }
    
    // 条件付き書式を適用
    if (rules.length > 0) {
      try {
        sheet.setConditionalFormatRules(rules);
        console.log(`${rules.length}個の条件付き書式ルールを適用しました`);
      } catch (error) {
        console.error('条件付き書式適用エラー:', error);
        throw error;
      }
    } else {
      console.log('適用する条件付き書式ルールがありません');
    }
    
    // 手動で条件付き書式を確認
    console.log('条件付き書式の確認開始');
    const appliedRules = sheet.getConditionalFormatRules();
    console.log(`適用された条件付き書式ルール数: ${appliedRules.length}`);
    
    appliedRules.forEach((rule, index) => {
      try {
        const ranges = rule.getRanges();
        console.log(`ルール${index + 1}: ${ranges.length}個の範囲に適用`);
        ranges.forEach((range, rangeIndex) => {
          console.log(`  範囲${rangeIndex + 1}: 行${range.getRow()}, 列${range.getColumn()}, 行数${range.getNumRows()}, 列数${range.getNumColumns()}`);
        });
      } catch (error) {
        console.error(`ルール${index + 1}の詳細確認エラー:`, error);
      }
    });
    
    console.log('条件付き書式の適用が完了しました');
    
    // 条件付き書式が適用されない場合の代替手段：手動で背景色を設定
    console.log('手動背景色設定の確認開始');
    try {
      let manualHighlightCount = 0;
      
      // データ行を1行ずつチェック
      for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
        try {
          const nameCell = sheet.getRange(rowIndex, 1); // A列（氏名）
          const nameValue = nameCell.getValue();
          
          // 小計行や全体合計行は除外
          if (nameValue && typeof nameValue === 'string' && 
              !nameValue.includes('小計') && !nameValue.includes('合計')) {
            
            const eValue = sheet.getRange(rowIndex, 5).getValue(); // E列
            const fValue = sheet.getRange(rowIndex, 6).getValue(); // F列
            
            // 数値に変換
            const eNum = parseFloat(eValue) || 0;
            const fNum = parseFloat(fValue) || 0;
            
            // E列+F列 < 50000の場合、黄色背景を設定
            if (eNum + fNum < 50000 && (eNum > 0 || fNum > 0)) {
              const efRange = sheet.getRange(rowIndex, 5, 1, 2); // E列とF列
              efRange.setBackground('#FFFF99'); // 薄い黄色
              manualHighlightCount++;
              console.log(`行${rowIndex}を手動で黄色背景に設定: E=${eNum}, F=${fNum}, 合計=${eNum + fNum}`);
            }
          }
        } catch (error) {
          console.error(`行${rowIndex}の手動背景色設定エラー:`, error);
        }
      }
      
      console.log(`手動背景色設定完了: ${manualHighlightCount}行を黄色背景に設定`);
      
    } catch (error) {
      console.error('手動背景色設定エラー:', error);
    }
    
  } catch (error) {
    console.error('条件付き書式の適用エラー:', error);
    console.error('エラーの詳細:', error.stack);
    // エラーが発生しても処理を続行
  }
}

// 1ページ目データ整理シートのヘッダー行を設定
function setupPage1DataHeaders(sheet) {
  console.log('1ページ目データ整理シートのヘッダー行を設定中...');
  
  const headers = [
    '送信日時',           // A列
    '氏名',              // B列
    '学年',              // C列
    '振込日',            // D列
    '振込金額',          // E列
    '直接振込金額',      // F列
    '寄付金',            // G列
    '負担金',            // H列
    'Tシャツ代',         // I列
    '奨励会負担金',      // J列
    'その他の費用',      // K列
    '兄弟1の学年',       // L列
    '兄弟1の名前',       // M列
    '兄弟2の学年',       // N列
    '兄弟2の名前',       // O列
    '兄弟3の学年',       // P列
    '兄弟3の名前',       // Q列
    '兄弟4の学年',       // R列
    '兄弟4の名前',       // S列
    '兄弟5の学年',       // T列
    '兄弟5の名前',       // U列
    '兄弟の人数',        // V列
    '備考'               // W列
  ];
  
  console.log('設定するヘッダー:', headers);
  console.log('ヘッダー数:', headers.length);
  
  // ヘッダー行を設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行の書式を設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  console.log('ヘッダー行の設定完了');
  
  // 列幅を自動調整
  sheet.autoResizeColumns(1, headers.length);
  
  return headers;
}

// 兄弟の人数の出力をテスト
function testSiblingCountOutput() {
  console.log('=== 兄弟の人数出力テスト開始 ===');
  
  try {
    // 設定された集計用ブックのIDを取得
    const aggregationBookId = getAggregationBookId();
    
    if (!aggregationBookId) {
      console.error('エラー: 集計用ブックのIDが設定されていません。');
      console.error('以下の手順で設定してください:');
      console.error('1. setAggregationBookId("YOUR_SPREADSHEET_ID") を実行');
      console.error('2. スプレッドシートIDはURLの /d/ と /edit の間の文字列です');
      return;
    }
    
    // 集計用ブックを開く
    const aggregationBook = SpreadsheetApp.openById(aggregationBookId);
    if (!aggregationBook) {
      console.error('集計用ブックを開けませんでした。IDを確認してください。');
      return;
    }
    
    console.log('集計用ブック名:', aggregationBook.getName());
    
    const page1Sheet = aggregationBook.getSheetByName('1ページ目データ整理');
    
    if (!page1Sheet) {
      console.error('集計用ブック内に「1ページ目データ整理」シートが見つかりません');
      console.log('利用可能なシート:', aggregationBook.getSheets().map(s => s.getName()));
      return;
    }
    
    console.log('1ページ目データ整理シートを発見:', page1Sheet.getName());
    
    // ヘッダー行を設定
    const headers = setupPage1DataHeaders(page1Sheet);
    console.log('ヘッダー設定完了:', headers);
    
    // テストデータを作成
    const testData = {
      basicInfo: {
        name: 'テスト太郎',
        grade: '3年',
        transferDate: '2025-01-15',
        hasSibling: true,
        siblingCount: 2,
        sibling1Grade: '1年',
        sibling1Name: 'テスト次郎',
        sibling2Grade: '2年',
        sibling2Name: 'テスト三郎'
      },
      breakdownDetails: {
        transferAmount: '50000',
        directTransferAmount: '10000',
        donationAmount: 20000,
        burdenFee: 120000,
        hasEncouragementFee: true,
        tshirtAmount: 3000,
        otherFee: 0
      }
    };
    
    console.log('テストデータ:', JSON.stringify(testData, null, 2));
    
    // テストデータを追加
    addPage1Data(page1Sheet, testData);
    
    console.log('テスト完了: 兄弟の人数がV列に正しく出力されているか確認してください');
    
  } catch (error) {
    console.error('テスト実行エラー:', error);
    console.error('エラーの詳細:', error.stack);
  }
}

// 集計用ブックのIDを設定
function setAggregationBookId(bookId) {
  try {
    // プロパティサービスを使用してIDを保存
    PropertiesService.getScriptProperties().setProperty('AGGREGATION_BOOK_ID', bookId);
    console.log('集計用ブックIDを設定しました:', bookId);
    
    // 設定したIDでブックを開けるかテスト
    const testBook = SpreadsheetApp.openById(bookId);
    if (testBook) {
      console.log('集計用ブックの接続テスト成功:', testBook.getName());
      console.log('利用可能なシート:', testBook.getSheets().map(s => s.getName()));
    } else {
      console.error('集計用ブックの接続テスト失敗');
    }
    
  } catch (error) {
    console.error('集計用ブックID設定エラー:', error);
  }
}

// 集計用ブックのIDを取得
function getAggregationBookId() {
  try {
    const bookId = PropertiesService.getScriptProperties().getProperty('AGGREGATION_BOOK_ID');
    if (bookId) {
      console.log('設定済みの集計用ブックID:', bookId);
      return bookId;
    } else {
      console.log('集計用ブックIDが設定されていません');
      return null;
    }
  } catch (error) {
    console.error('集計用ブックID取得エラー:', error);
    return null;
  }
}

// ========================================
// 共通の集計処理関数（手動・自動両方で使用）
// ========================================
// 
// 【重要】集計処理を修正する際は、以下の関数を修正してください。
// 手動実行時と自動実行時で同じ処理が実行されるようになります。
// 
// 修正が必要な場合：
// 1. updateDonationSummary() 関数を修正
// 2. updateGiftSummaryWithNewLogic() 関数を修正
// 3. 必要に応じて initializeDonationSummarySheet() や clearGiftSummaryRows() も修正
// 
// これらの関数は手動・自動両方から呼び出されるため、
// 一箇所を修正するだけで両方に反映されます。

// 共通の寄付金一覧集計処理（手動・自動両方で使用）
function executeDonationSummary() {
  try {
    console.log('=== 共通の寄付金一覧集計処理開始 ===');
    
    // 集計用ブックを取得
    const aggregationBook = getAggregationBook();
    const donationSheet = aggregationBook.getSheetByName('寄付金一覧');
    
    if (!donationSheet) {
      console.error('寄付金一覧シートが見つかりません');
      throw new Error('寄付金一覧シートが見つかりません');
    }
    
    // ヘッダー初期化（手動実行時と同じ処理）
    initializeDonationSummarySheet(donationSheet);
    
    // 集計実行
    updateDonationSummary(donationSheet);
    
    console.log('=== 共通の寄付金一覧集計処理完了 ===');
    
  } catch (error) {
    console.error('共通の寄付金一覧集計処理エラー:', error);
    throw error;
  }
}

// 共通の返礼品一覧集計処理（手動・自動両方で使用）
function executeGiftSummary() {
  try {
    console.log('=== 共通の返礼品一覧集計処理開始 ===');
    
    // 集計用ブックを取得
    const aggregationBook = getAggregationBook();
    const giftSheet = aggregationBook.getSheetByName('返礼品一覧');
    
    if (!giftSheet) {
      console.error('返礼品一覧シートが見つかりません');
      throw new Error('返礼品一覧シートが見つかりません');
    }
    
    // 設定用ブックから必要なシートを取得
    const settingsBook = SpreadsheetApp.getActiveSpreadsheet();
    const page1DataSheet = aggregationBook.getSheetByName('1ページ目データ整理');
    const page2DataSheet = aggregationBook.getSheetByName('2ページ目データ整理');
    const giftSettingsSheet = settingsBook.getSheetByName('返礼品設定');
    
    if (!page1DataSheet || !page2DataSheet || !giftSettingsSheet) {
      console.error('必要なシートが見つかりません');
      throw new Error('必要なシートが見つかりません。1ページ目データ整理、2ページ目データ整理、返礼品設定シートが必要です。');
    }
    
    // 既存の集計行をクリア
    clearGiftSummaryRows(giftSheet);
    
    // 新しい集計ロジックで更新
    updateGiftSummaryWithNewLogic(giftSheet, page1DataSheet, page2DataSheet, giftSettingsSheet);
    
    // テンプレートの貼り付けも実行
    syncGiftSettingsToSheet(giftSheet);
    
    console.log('=== 共通の返礼品一覧集計処理完了 ===');
    
  } catch (error) {
    console.error('共通の返礼品一覧集計処理エラー:', error);
    throw error;
  }
}