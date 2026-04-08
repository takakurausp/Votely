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
 * テスト実行：未投票のトークンを集めて、一斉にWebアプリにPOST送信する。
 *
 * 重要な設計ポイント：
 *  - UrlFetchApp.fetchAll() を使い「数百件を完全に並列で発火」させる。
 *    1件ずつ UrlFetchApp.fetch() のループにしてしまうと直列実行になり、
 *    スパイクテストとしての意味が失われるので注意。
 *  - リトライも 5xx/429 の失敗分だけ集めて再度 fetchAll() する 2 ラウンド方式。
 *    こうすると並列性を保ったまま一過性エラーを吸収できる。
 *  - Content-Type は text/plain。本番フロント (docs/index.html) と完全に揃える
 *    ことで、テスト経路でだけ挙動が違ってしまう事故を防ぐ。
 *  - 投票項目が複数ある場合、各投票の選択肢を「設定シート 2 行目」から
 *    まとめて読み取り、_recordVoteDirect の空チェックで弾かれないようにする。
 */
function runSpikeTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName('名簿とトークン');
  var settingsSheet = ss.getSheetByName('設定');

  // WebアプリのURLを取得（GASに直接POSTするので A4 = GAS URL を使う）
  var appUrl = String(settingsSheet.getRange('A4').getValue()).trim();
  if (!appUrl || !appUrl.startsWith('https://script.google.com/')) {
    SpreadsheetApp.getUi().alert('エラー: 設定シートのA4に正しいWebアプリURLを入力してください。');
    return;
  }

  // 投票数ぶんの選択肢を組み立てる
  // - 行1の B 列以降にタイトルがある列を「投票」として認識
  // - 各列の行2を「その投票の先頭の選択肢」として採用
  var validChoicesArray = _collectFirstChoicesPerVote(settingsSheet);
  if (validChoicesArray.length === 0) {
    SpreadsheetApp.getUi().alert('エラー: 設定シートに投票項目（B1セル以降のタイトル）が見つかりません。');
    return;
  }

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
  Logger.log('投票数: ' + validChoicesArray.length + '（各投票で送る選択肢: ' + JSON.stringify(validChoicesArray) + '）');

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
      method: 'post',
      // GAS の doPost は Content-Type を問わず e.postData.contents を JSON.parse するので、
      // 本番フロント (docs/index.html) と同じ text/plain に揃えて挙動差をなくす。
      contentType: 'text/plain',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  }

  // 💥 一斉送信ラウンド1（完全並列）
  var startTime = new Date().getTime();
  var responses = UrlFetchApp.fetchAll(requests);

  // ラウンド2: 5xx / 429 のものだけ集めてもう一度並列で投げる
  var MAX_RETRY_ROUNDS = 2;
  for (var round = 1; round <= MAX_RETRY_ROUNDS; round++) {
    var retryRequests = [];
    var retryIndices = [];
    for (var k = 0; k < responses.length; k++) {
      var code = responses[k].getResponseCode();
      if (code >= 500 || code === 429) {
        retryRequests.push(requests[k]);
        retryIndices.push(k);
      }
    }
    if (retryRequests.length === 0) break;

    Logger.log('リトライ ラウンド ' + round + ': ' + retryRequests.length + ' 件を再送信');
    Utilities.sleep(1000 * round); // 1秒, 2秒... の線形バックオフ
    var retryResponses = UrlFetchApp.fetchAll(retryRequests);
    for (var r = 0; r < retryResponses.length; r++) {
      responses[retryIndices[r]] = retryResponses[r];
    }
  }

  var endTime = new Date().getTime();

  // 結果の集計
  var successCount = 0;
  var errorCount = 0;
  var errorMessages = {};
  for (var m = 0; m < responses.length; m++) {
    var res = responses[m];
    if (res.getResponseCode() === 200) {
      var resBody;
      try {
        resBody = JSON.parse(res.getContentText());
      } catch (e) {
        errorCount++;
        errorMessages['JSON parse error'] = (errorMessages['JSON parse error'] || 0) + 1;
        continue;
      }
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

/**
 * 設定シートから「各投票の先頭の選択肢」を 1 度の getValues 呼び出しで集める。
 * 構造前提:
 *   行1: B1, C1, D1, ... に投票タイトル
 *   行2: B2, C2, D2, ... に各投票の最初の選択肢
 * タイトルが空の列は投票として扱わない。
 */
function _collectFirstChoicesPerVote(settingsSheet) {
  var lastCol = settingsSheet.getLastColumn();
  if (lastCol < 2) return [];

  var headerRow = settingsSheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  var firstOptionRow = settingsSheet.getRange(2, 2, 1, lastCol - 1).getValues()[0];

  var choices = [];
  for (var i = 0; i < headerRow.length; i++) {
    var title = headerRow[i];
    if (title === '' || title === null || title === undefined) continue;
    var opt = firstOptionRow[i];
    choices.push(String(opt == null ? '' : opt));
  }
  return choices;
}
