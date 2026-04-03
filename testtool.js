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
function runSpikeTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName('名簿とトークン');
  var settingsSheet = ss.getSheetByName('設定');
  
  // WebアプリのURLを取得
  var appUrl = String(settingsSheet.getRange('A4').getValue()).trim();
  if (!appUrl || !appUrl.startsWith('https://script.google.com/')) {
    Logger.log('エラー: 設定シートのA4に正しいWebアプリURLを入力してください。');
    return;
  }

  // 設定から有効な「選択肢」を1つ自動取得する（エラーにならないようにするため）
  var validChoice = String(settingsSheet.getRange('B2').getValue()); 
  var validChoicesArray = [validChoice]; // 複数投票がある場合は [validChoice, validChoice2...] となりますが、今回は1つ目の投票のみテストとして代入

  // 1. 名簿から未投票のトークンを収集
  var data = rosterSheet.getDataRange().getValues();
  var tokens = [];
  for (var i = 1; i < data.length; i++) {
    var token = data[i][1]; // B列（トークン）
    var voted = data[i][3]; // D列（投票済みフラグ）
    if (token && voted !== true) {
      tokens.push(token);
    }
    // GASの1回の一斉送信上限(1000)を超えないよう制限
    if (tokens.length >= 500) break; 
  }

  if (tokens.length === 0) {
    Logger.log('テスト用のトークンがありません。ダミーユーザーを作成し、トークンを発行してください。');
    return;
  }

  Logger.log('【スパイクテスト開始】 ' + tokens.length + ' 件の同時リクエストを送信します...');

  // 2. リクエストの配列を構築
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
      muteHttpExceptions: true // エラーが返ってきてもスクリプトを止めない
    });
  }

  // 3. 💥一斉送信（ここで数百件のリクエストが完全に並列で発火します）💥
  var startTime = new Date().getTime();
  var responses = UrlFetchApp.fetchAll(requests);
  var endTime = new Date().getTime();

  // 4. 結果の集計
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
  
  // 5. ログ出力
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