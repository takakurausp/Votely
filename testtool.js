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

  // WebアプリのURLを取得（GASに直接POSTするので B4 = GAS URL を使う）
  var appUrl = String(settingsSheet.getRange('B4').getValue()).trim();
  if (!appUrl || !appUrl.startsWith('https://script.google.com/')) {
    SpreadsheetApp.getUi().alert('エラー: 設定シートのB4に正しいWebアプリURLを入力してください。');
    return;
  }

  // 投票数ぶんの選択肢を組み立てる
  // - 行1の C 列以降にタイトルがある列を「投票」として認識
  // - 各列の行2を「その投票の先頭の選択肢」として採用
  var validChoicesArray = _collectFirstChoicesPerVote(settingsSheet);
  if (validChoicesArray.length === 0) {
    SpreadsheetApp.getUi().alert('エラー: 設定シートに投票項目（C1セル以降のタイトル）が見つかりません。');
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

  // ── 時間制限 ───────────────────────────────────────────────────────────────
  // GAS の実行時間上限は 6 分（360 秒）。
  // fetchAll の応答待ち（ロックタイムアウト最大 30 秒 × バッチ数）が積み重なると
  // 容易に超過するため、残り 60 秒を切ったら送信を打ち切って集計に移行する。
  var EXEC_LIMIT_MS   = 5 * 60 * 1000;   // 5 分で打ち切り（1 分のバッファ）
  var startTime       = new Date().getTime();

  function _timeRemaining() {
    return EXEC_LIMIT_MS - (new Date().getTime() - startTime);
  }

  // 1バッチあたりのリクエスト数。
  // GAS の urlfetch は短時間大量呼び出しで "Service invoked too many times" が出るため、
  // BATCH_SIZE 件ずつ fetchAll → 1秒 sleep を繰り返してレート制限を回避する。
  var BATCH_SIZE = 50;

  Logger.log('【スパイクテスト開始】 ' + tokens.length + ' 件を ' + BATCH_SIZE + ' 件/バッチで送信します...');
  Logger.log('投票数: ' + validChoicesArray.length + '（各投票で送る選択肢: ' + JSON.stringify(validChoicesArray) + '）');
  Logger.log('実行時間上限: ' + (EXEC_LIMIT_MS / 1000) + ' 秒（超過時は途中集計して終了）');

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

  // 💥 バッチ並列送信（BATCH_SIZE 件ずつ fetchAll、バッチ間は 1 秒 sleep）
  var responses = _fetchAllBatched(requests, BATCH_SIZE, _timeRemaining);

  // ラウンド2: 5xx / 429 のものだけ集めてもう一度バッチ並列で投げる
  // 時間が残っている場合のみ実施する
  var MAX_RETRY_ROUNDS = 2;
  for (var round = 1; round <= MAX_RETRY_ROUNDS; round++) {
    if (_timeRemaining() < 30000) {
      Logger.log('⏱ 残り時間が少ないためリトライをスキップします（残 ' + Math.round(_timeRemaining()/1000) + ' 秒）');
      break;
    }
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
    Utilities.sleep(1000 * round);
    var retryResponses = _fetchAllBatched(retryRequests, BATCH_SIZE, _timeRemaining);
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
 * リクエスト配列を batchSize 件ずつに分割して fetchAll し、全レスポンスを結合して返す。
 * バッチ間は 1 秒 sleep することで GAS の urlfetch レート制限（短時間大量呼び出し禁止）を回避する。
 *
 * @param {Object[]} allRequests    - UrlFetchApp.fetchAll に渡せる形式のリクエスト配列
 * @param {number}   batchSize      - 1 バッチあたりの最大リクエスト数（推奨: 50）
 * @param {Function} [timeRemaining] - 残り時間(ms)を返す関数。30秒以下になったら打ち切る
 * @returns {HTTPResponse[]} allRequests と同じ順序・長さのレスポンス配列
 *          時間切れで打ち切った分は {getResponseCode:()=>0, getContentText:()=>'{}'} の
 *          ダミーオブジェクトで埋める。
 */
function _fetchAllBatched(allRequests, batchSize, timeRemaining) {
  var allResponses = [];

  // 未送信分をダミーレスポンスで初期化しておく（打ち切り時に長さを合わせるため）
  var _dummy = {
    getResponseCode: function() { return 0; },
    getContentText:  function() { return '{}'; }
  };
  for (var init = 0; init < allRequests.length; init++) {
    allResponses.push(_dummy);
  }

  for (var start = 0; start < allRequests.length; start += batchSize) {
    // 時間切れチェック：30 秒以下になったら送信を打ち切る
    if (typeof timeRemaining === 'function' && timeRemaining() < 30000) {
      var remaining = allRequests.length - start;
      Logger.log('⏱ 時間切れのため残り ' + remaining + ' 件の送信をスキップします。');
      break;
    }

    var batch = allRequests.slice(start, start + batchSize);
    var batchNum = Math.floor(start / batchSize) + 1;
    var totalBatches = Math.ceil(allRequests.length / batchSize);
    Logger.log('バッチ ' + batchNum + '/' + totalBatches + ': ' + batch.length + ' 件を送信中...');

    var batchResponses = UrlFetchApp.fetchAll(batch);
    for (var i = 0; i < batchResponses.length; i++) {
      allResponses[start + i] = batchResponses[i];
    }

    // 最後のバッチ以外は sleep でレート制限を回避
    if (start + batchSize < allRequests.length) {
      Utilities.sleep(1000);
    }
  }
  return allResponses;
}

/**
 * 設定シートから「各投票の先頭の選択肢」を 1 度の getValues 呼び出しで集める。
 * 構造前提:
 *   A 列 = 項目名ラベル、B 列 = 設定値、C 列以降 = 投票項目
 *   行1: C1, D1, E1, ... に投票タイトル
 *   行2: C2, D2, E2, ... に各投票の最初の選択肢
 * タイトルが空の列は投票として扱わない。
 */
function _collectFirstChoicesPerVote(settingsSheet) {
  var lastCol = settingsSheet.getLastColumn();
  if (lastCol < 3) return [];

  // C 列（列番号3）以降を読む
  var headerRow      = settingsSheet.getRange(1, 3, 1, lastCol - 2).getValues()[0];
  var firstOptionRow = settingsSheet.getRange(2, 3, 1, lastCol - 2).getValues()[0];

  var choices = [];
  for (var i = 0; i < headerRow.length; i++) {
    var title = headerRow[i];
    if (title === '' || title === null || title === undefined) continue;
    var opt = firstOptionRow[i];
    choices.push(String(opt == null ? '' : opt));
  }
  return choices;
}
