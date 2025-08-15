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
  const active = SpreadsheetApp.getActiveSpreadsheet();
  const formSettingsSheet = active.getSheetByName('フォーム設定');
  if (!formSettingsSheet) return '';
  const values = formSettingsSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const k = values[i][0];
    const v = values[i][1];
    const enabled = values[i][2];
    if (enabled && String(k).trim() === key && v) {
      return String(v).trim();
    }
  }
  return '';
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
    const idOrUrl = getFormSettingValue('aggregationBookId') || getFormSettingValue('aggregationBookUrl');
    if (idOrUrl) {
      const id = extractSpreadsheetId(idOrUrl);
      const book = SpreadsheetApp.openById(id);
      return book;
    }
  } catch (e) {
    console.warn('集計用ブック取得の警告:', e);
  }
  console.warn('aggregationBookId/url が未設定のため、ActiveSpreadsheetを使用します');
  return SpreadsheetApp.getActiveSpreadsheet();
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
    }
    
    // 返礼品一覧シートを更新（集計用ブック）
    const giftSheet = aggregationBook.getSheetByName('返礼品一覧');
    console.log('返礼品一覧シート(集計用):', giftSheet ? '存在' : '不存在');
    if (giftSheet) {
      updateGiftSummary(giftSheet);
    }
    
    // 個票を自動生成
    let individualReportResult = { success: true, message: '' };
    try {
      console.log('=== 個票自動生成開始 ===');
      console.log('formData構造:', JSON.stringify(formData, null, 2));
      
      const studentName = formData.basicInfo.basicInfo.name;
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
  
  const basicInfo = formData.basicInfo.basicInfo;
  const breakdownDetails = formData.basicInfo.breakdownDetails;
  
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
  
  // 兄弟の名前を追加（最大5人分）
  for (let i = 1; i <= 5; i++) {
    const siblingName = basicInfo[`sibling${i}Name`] || '';
    newRow.push(siblingName); // L列～P列：兄弟1の名前～兄弟5の名前
  }
  
  newRow.push(''); // Q列：備考
  
  console.log('newRow:', newRow);
  sheet.appendRow(newRow);
}

// 2ページ目データ整理シートにデータを追加
function addPage2Data(sheet, formData) {
  console.log('addPage2Data - formData:', JSON.stringify(formData, null, 2));
  
  const basicInfo = formData.basicInfo.basicInfo;
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
      const isActive = row[3]; // 新スキーマ: D列
      
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
    directDate: colIndexMap['直接振込日'] || 7, // ご指定: G列
    note: colIndexMap['その他特筆事項'] || colIndexMap['備考'] || 8
  };

  donors.forEach((donor, index) => {
    console.log(`donor ${index}:`, JSON.stringify(donor, null, 2));
    
    // 返礼品の処理（IDまたは名前の両方に対応）
    let giftName = '';
    if (donor.gift) {
      if (donor.gift === 'クリアファイル' || donor.gift === 'タオル' || donor.gift === 'お菓子大' || donor.gift === 'キーホルダー' || donor.gift === 'お菓子小') {
        giftName = donor.gift;
      } else if (donor.gift === 'clearfile') {
        giftName = 'クリアファイル';
      } else if (giftIdToName[donor.gift]) {
        giftName = giftIdToName[donor.gift];
      } else {
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
    if (COL.directDate) rowArray[COL.directDate - 1] = donor.directTransferDate || '';
    if (COL.note) rowArray[COL.note - 1] = donor.note || '';
    
    console.log(`donor ${index} rowArray:`, rowArray);
    sheet.appendRow(rowArray);
  });
}

// 寄付金一覧シートにデータを追加
function addDonationData(sheet, formData) {
  const basicInfo = formData.basicInfo.basicInfo;
  const breakdownDetails = formData.basicInfo.breakdownDetails;
  
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
  
  const newRow = [
    basicInfo.name,                    // A列：氏名
    '',                                // B列：備考
    breakdownDetails.donationAmount,    // C列：寄付金
    burdenAmount,                      // D列：負担金
    breakdownDetails.tshirtAmount || 0, // E列：Tシャツ代
    encouragementAmount,               // F列：奨励会負担金
    breakdownDetails.otherFee || 0,    // G列：その他の費用
    transferAmount,                    // H列：振込合計金額
    donationBurdenTotal               // I列：寄付・負担金計
  ];
  
  // 2行目に挿入（既存データを下にずらす）
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
}

// 返礼品一覧シートを動的に更新
function updateGiftSummary(sheet) {
  try {
    const spreadsheet = getAggregationBook();
    
    // 返礼品設定を取得
    // 返礼品設定は「設定用ブック（Active）」から取得
    const settingsBook = SpreadsheetApp.getActiveSpreadsheet();
    const giftSettingsSheet = settingsBook.getSheetByName('返礼品設定');
    if (!giftSettingsSheet) {
      console.log('返礼品設定シートが見つかりません');
      return;
    }
    
    // 2ページ目データ整理シートを取得
    const page2DataSheet = spreadsheet.getSheetByName('2ページ目データ整理');
    if (!page2DataSheet) {
      console.log('2ページ目データ整理シートが見つかりません');
      return;
    }
    
    // 1ページ目データ整理シートを取得
    const page1DataSheet = spreadsheet.getSheetByName('1ページ目データ整理');
    if (!page1DataSheet) {
      console.log('1ページ目データ整理シートが見つかりません');
      return;
    }
    
    // 返礼品設定から動的に列を生成
    const giftSettings = getGiftSettings(giftSettingsSheet);
    
    // シートのヘッダーを更新
    updateGiftSheetHeaders(sheet, giftSettings);
    
    // データを集計して更新
    updateGiftSheetData(sheet, page1DataSheet, page2DataSheet, giftSettings);
    
  } catch (error) {
    console.error('返礼品一覧シートの更新エラー:', error);
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
      
      // 返礼品の設定（クリアファイルの場合は枚数も表示）
      let giftDisplay = donor.gift;
      if (donor.gift === 'クリアファイル' && donor.clearFileCount > 0) {
        giftDisplay = `クリアファイル(${donor.clearFileCount}枚)`;
      }
      sheet.getRange(`D${currentRow}`).setValue(giftDisplay);  // 返礼品
      
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



// 個票シートを生成
function generateIndividualReport(studentName) {
  try {
    // 年次の集計用ブックからデータを取得
    const aggregationBook = getAggregationBook();
    const page1Sheet = aggregationBook.getSheetByName('1ページ目データ整理');
    const page2Sheet = aggregationBook.getSheetByName('2ページ目データ整理');
    
    if (!page1Sheet || !page2Sheet) {
      console.log('必要なシートが見つかりません');
      return;
    }
    
    // 1ページ目データから該当する生徒の情報を取得
    const page1Data = page1Sheet.getDataRange().getValues();
    let studentInfo = null;
    
    for (let i = 1; i < page1Data.length; i++) {
      if (page1Data[i][1] === studentName) { // 氏名列（B列）
        // 負担金の人数を計算: 1 + 兄弟の数（本人を除く）
        // 兄弟名は L〜P 列（最大5人分）
        let siblingCount = 0;
        for (let col = 11; col <= 15; col++) { // L(12)-P(16) だが0-basedで+1ずれるため11..15
          const name = page1Data[i][col];
          if (name && String(name).trim() !== '') siblingCount++;
        }
        const burdenFeeCount = 1 + siblingCount;
        
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
    
    for (let i = 1; i < page2Data.length; i++) {
      if (page2Data[i][1] === studentName) { // B列：1ページ目氏名
        donors.push({
          name: page2Data[i][2], // C列：御芳名
          amount: page2Data[i][3], // D列：寄付金額
          gift: page2Data[i][4], // E列：返礼品
          clearFileCount: page2Data[i][5] || 0 // F列：クリアファイル枚数
        });
      }
    }
    
    if (donors.length === 0) {
      console.log('寄付者データが見つかりません:', studentName);
      return;
    }
    
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