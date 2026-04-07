// =============================================================================
// スパイクテスト（負荷テスト）用ツール
// =============================================================================

/**
 * 準備1：テスト用のダミー名簿を自動生成する（500人分）
 * 名簿シートに「test1@example.com」のようなアドレスを一気に書き込みます。
 */
function createDummyUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName('名簿とトークン');
  var lastRow = rosterSheet.getLastRow();
  
  var dummyData = [];
  var TEST_COUNT = 500; // テストする人数
  
  for (var i = 1; i <= TEST_COUNT; i++) {
    dummyData.push(['test' + i + '@example.com', '', '', '']);
  }
  
  rosterSheet.getRange(lastRow + 1, 1, TEST_COUNT, 4).setValues(dummyData);
  SpreadsheetApp.getUi().alert(TEST_COUNT + '件のダミーユーザーを追加しました。\nメニューから「② トークン生成のみ」を実行してトークンを発行してください。');
}

/**
 * テスト実行：未投票のトークンを集めて、一斉にWebアプリにPOST送信する
 */
/**
 * リトライ付きHTTPリクエストを実行する
 * @param {string} url - リクエスト先URL
 * @param {object} options - リクエストオプション
 * @param {number} maxRetries - 最大リトライ回数（デフォルト: 3）
 * @param {number} delay - 初回リトライ間隔（ミリ秒, デフォルト: 1000）
 * @return {GoogleAppsScript.URL_Fetch.HTTPResponse} レスポンス
 */
function fetchWithRetry(url, options, maxRetries = 3, delay = 1000) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return UrlFetchApp.fetch(url, options);
    } catch (error) {
      if (i === maxRetries - 1) throw error;
      Logger.log(`リトライ中... (${i + 1}/${maxRetries})`);
      Utilities.sleep(delay * (i + 1)); // 指数バックオフ
    }
  }
}

/**
 * スパイクテスト実行（リトライロジック付き）
 */
function runSpikeTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName('名簿とトークン');
  var settingsSheet = ss.getSheetByName('設定');
  
  // WebアプリのURLを取得
  var appUrl = String(settingsSheet.getRange('A4').getValue()).trim();
  if (!appUrl || !appUrl.startsWith('https://script.google.com/')) {
    SpreadsheetApp.getUi().alert('エラー: 設定シートのA4に正しいWebアプリURLを入力してください。');
    return;
  }

  // 選択肢を取得
  var validChoice = String(settingsSheet.getRange('B2').getValue());
  var validChoicesArray = [validChoice];

  // 未投票のトークンを収集
  var data = rosterSheet.getDataRange().getValues();
  var tokens = [];
  for (var i = 1; i < data.length; i++) {
    var token = data[i][1];
    var voted = data[i][3];
    if (token && voted !== true) {
      tokens.push(token);
    }
    if (tokens.length >= 500) break;
  }

  if (tokens.length === 0) {
    SpreadsheetApp.getUi().alert('テスト用のトークンがありません。ダミーユーザーを作成し、トークンを発行してください。');
    return;
  }

  Logger.log('【スパイクテスト開始】 ' + tokens.length + ' 件の同時リクエストを送信します...');

  // リクエストの配列を構築
  var requests = [];
  for (var j = 0; j < tokens.length; j++) {
    var payload = {
      action: 'submitVote',
      token: tokens[j],
      choices: validChoicesArray
    };
    requests.push({
      url: appUrl,
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  }

  // リトライ付きで一斉送信
  var startTime = new Date().getTime();
  var responses = [];
  for (var req of requests) {
    try {
      var res = fetchWithRetry(req.url, req, 3, 1000);
      responses.push(res);
    } catch (error) {
      Logger.log(`リクエスト失敗: ${error.message}`);
      responses.push({ getResponseCode: () => 500, getContentText: () => '{}' });
    }
  }
  var endTime = new Date().getTime();

  // 結果の集計
  var successCount = 0;
  var errorCount = 0;
  var errorMessages = {};
  for (var k = 0; k < responses.length; k++) {
    var res = responses[k];
    if (res.getResponseCode() === 200) {
      var resBody = JSON.parse(res.getContentText());
      if (resBody.success) {
        successCount++;
      } else {
        errorCount++;
        errorMessages[resBody.message] = (errorMessages[resBody.message] || 0) + 1;
      }
    } else {
      errorCount++;
      var status = res.getResponseCode();
      errorMessages['HTTP Error ' + status] = (errorMessages['HTTP Error ' + status] || 0) + 1;
    }
  }

  var timeTaken = (endTime - startTime) / 1000;
  Logger.log('====================================');
  Logger.log('テスト完了！');
  Logger.log('所要時間: ' + timeTaken + ' 秒');
  Logger.log('成功: ' + successCount + ' 件');
  Logger.log('失敗: ' + errorCount + ' 件');
  if (errorCount > 0) {
    Logger.log('【エラー内訳】');
    for (var msg in errorMessages) {
      Logger.log(' - ' + msg + ' : ' + errorMessages[msg] + '件');
    }
  }
  Logger.log('====================================');
}