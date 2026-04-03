// =============================================================================
// Votely - GASを用いたオンライン匿名投票システム
// コード.gs - サーバーサイドロジック全体
// =============================================================================

// ---- シート名定数 ----
var SHEET_SETTINGS   = '設定';
var SHEET_ROSTER     = '名簿とトークン';
var SHEET_BALLOT_BOX = '投票箱';
var SHEET_SYSLOG     = 'システム管理';

// ---- 名簿とトークンシートの列インデックス（1始まり） ----
var COL_EMAIL = 1; // A列: メールアドレス
var COL_TOKEN = 2; // B列: トークン
var COL_URL   = 3; // C列: 投票用URL
var COL_VOTED = 4; // D列: 投票済みフラグ

// =============================================================================
// メニュー登録
// =============================================================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Votely 管理')
    .addItem('⓪ 【初回のみ】データベース自動構築', 'initializeSystem')
    .addSeparator()
    .addItem('① トークン生成 ＋ 案内メール送信（GAS直接）', 'generateTokensAndSendEmails')
    .addItem('② トークン生成のみ（URLリスト作成・外部メーラー用）', 'generateTokensOnly')
    .addSeparator()
    .addItem('③ ▶ 投票開始（※高速モード用: バッチ処理開始）', 'startVotingTrigger')
    .addItem('④ ⏹ 投票終了（※高速モード用: バッチ処理停止）', 'stopVotingTrigger')
    .addSeparator()
    .addItem('⑤ 集計・結果通知（手動実行）', 'tallySendResults')
    .addItem('⑥ 締め切りトリガーをセットアップ', 'setupTrigger')
    .addToUi();
}

// =============================================================================
// ⓪ システム初期化（データベース構築）
// =============================================================================

/**
 * 必要なシート群を自動作成し、初期レイアウトを整える。
 * 新規スプレッドシートで初回のみ実行する。
 */
function initializeSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  function getOrCreateSheet(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    return sheet;
  }

  // 1. 設定シート
  var settingsSheet = getOrCreateSheet(SHEET_SETTINGS);
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.getRange('A1').setNote('【必須】主催者メールアドレス（結果通知先。複数ある場合はカンマ区切り）');
    settingsSheet.getRange('A2').setNote('【任意】締め切り日時（例: 2026/04/10 12:00）');
    settingsSheet.getRange('A3').setNote('【必須】事務局パスワード（当日受付用）');
    settingsSheet.getRange('A4').setNote('【必須】WebアプリURL（GASをデプロイしたURL）');
    settingsSheet.getRange('A5').setNote('【必須】フロントエンドのURL（GitHub Pagesなど）');
    settingsSheet.getRange('A6').setNote('【必須】動作モード（通常モード or 高速モード）');
    
    settingsSheet.getRange('A6').setValue('通常モード');
    
    settingsSheet.getRange('B1').setValue('投票タイトル（例: 懇親会の場所）');
    settingsSheet.getRange('B2').setValue('選択肢A');
    settingsSheet.getRange('B3').setValue('選択肢B');
    settingsSheet.setColumnWidth(1, 300); 
  }

  // 2. 名簿とトークンシート
  var rosterSheet = getOrCreateSheet(SHEET_ROSTER);
  if (rosterSheet.getLastRow() === 0) {
    rosterSheet.appendRow(['メールアドレス', 'トークン', '投票用URL', '投票済みフラグ']);
    rosterSheet.getRange('A1:D1').setBackground('#d9ead3').setFontWeight('bold');
    rosterSheet.setFrozenRows(1);
  }

  // 3. 投票箱シート
  var ballotSheet = getOrCreateSheet(SHEET_BALLOT_BOX);
  if (ballotSheet.getLastRow() === 0) {
    ballotSheet.appendRow(['タイムスタンプ', '投票1の選択肢', '投票2の選択肢']);
    ballotSheet.getRange('A1:D1').setBackground('#cfe2f3').setFontWeight('bold');
    ballotSheet.setFrozenRows(1);
  }

  // 4. システム管理シート
  var sysSheet = getOrCreateSheet(SHEET_SYSLOG);
  if (sysSheet.getLastRow() === 0) {
    sysSheet.getRange('A1').setValue(false);
    sysSheet.getRange('B1').setValue('← 集計完了フラグ（手動で変更しないでください）');
    sysSheet.getRange('A1:B1').setBackground('#fce5cd');
  }

  // 5. デフォルトシート削除
  var defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  ui.alert('Votelyのデータベース構築が完了しました！\n\nA6セルで「通常モード」と「高速モード」を切り替えられます。');
}

// =============================================================================
// エントリポイント: doGet / doPost
// =============================================================================

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', app: 'Votely' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || '';

    if (action === 'getVoteFormData') {
      result = getVoteFormData(body.token);
    } else if (action === 'submitVote') {
      result = submitVote(body.token, body.choices);
    } else if (action === 'verifyPassword') {
      result = { success: verifyAdminPassword(body.password) };
    } else if (action === 'registerAttendee') {
      result = registerAttendee(body.email, body.password);
    } else {
      result = { success: false, message: '不明なアクション: ' + action };
    }
  } catch (err) {
    result = { success: false, message: 'サーバーエラー: ' + err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// 機能A: 事前準備・案内メール送信
// =============================================================================

/**
 * トークンとURLを生成し、GASから直接メール送信する（送信数制限の警告あり）
 */
function generateTokensAndSendEmails() {
  var ui = SpreadsheetApp.getUi();
  
  var alertMessage = '【注意】GASによるメール一斉送信について\n\n' +
                     'Googleアカウントには1日あたりのメール送信数に厳しい上限があります。\n' +
                     '・無料のGoogleアカウント：1日 100件まで\n' +
                     '・Google Workspaceアカウント：1日 1,500件まで\n\n' +
                     '上限を超えると途中でエラーとなり、システムが停止します。\n' +
                     '参加者が上限を超える場合は「いいえ」を押し、メニューの「② トークン生成のみ」を使って外部の配信ツールをご利用ください。\n\n' +
                     'このままメール送信を開始してもよろしいですか？';
                     
  var response = ui.alert('送信制限の確認', alertMessage, ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  var settings    = _getSettings();
  var pagesUrl    = settings.pagesUrl || settings.appUrl;
  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();
  var sentCount   = 0;

  for (var i = 2; i <= lastRow; i++) {
    var email         = rosterSheet.getRange(i, COL_EMAIL).getValue();
    var existingToken = rosterSheet.getRange(i, COL_TOKEN).getValue();

    if (email && !existingToken) {
      var voteUrl = _generateTokenForRow(i, rosterSheet, pagesUrl);
      _sendInvitationEmail(email, voteUrl, settings);
      sentCount++;
    }
  }
  ui.alert(sentCount + ' 件のアドレスにトークンを生成し、案内メールを送信しました。');
}

/**
 * トークンとURLの生成のみ行う（外部メーラー用）
 */
function generateTokensOnly() {
  var settings    = _getSettings();
  var pagesUrl    = settings.pagesUrl || settings.appUrl;
  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();
  var genCount    = 0;

  for (var i = 2; i <= lastRow; i++) {
    var email         = rosterSheet.getRange(i, COL_EMAIL).getValue();
    var existingToken = rosterSheet.getRange(i, COL_TOKEN).getValue();

    if (email && !existingToken) {
      _generateTokenForRow(i, rosterSheet, pagesUrl);
      genCount++;
    }
  }
  SpreadsheetApp.getUi().alert(
    genCount + ' 件のトークンとURLを生成しました。\n「名簿とトークン」シートをご確認ください。'
  );
}

// =============================================================================
// 機能B: 事務局ポータル
// =============================================================================

function verifyAdminPassword(password) {
  var settings = _getSettings();
  return String(settings.adminPassword).trim() === String(password).trim();
}

function registerAttendee(email, password) {
  if (!verifyAdminPassword(password)) {
    return { success: false, url: '', message: '認証エラー：パスワードが正しくありません。' };
  }
  if (!email || !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
    return { success: false, url: '', message: '有効なメールアドレスを入力してください。' };
  }

  var settings = _getSettings();
  var appUrl   = settings.appUrl;
  var pagesUrl = settings.pagesUrl || appUrl;

  if (settings.deadline && new Date() > new Date(settings.deadline)) {
    return { success: false, url: '', message: '投票の締め切りを過ぎています。登録できません。' };
  }

  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    if (rosterSheet.getRange(i, COL_EMAIL).getValue() === email) {
      var existingUrl = rosterSheet.getRange(i, COL_URL).getValue();
      return { success: true, url: existingUrl, message: '登録済みです。既存のURLを返します。' };
    }
  }

  var newRow  = lastRow + 1;
  rosterSheet.getRange(newRow, COL_EMAIL).setValue(email);
  var voteUrl = _generateTokenForRow(newRow, rosterSheet, pagesUrl);

  try {
    _sendInvitationEmail(email, voteUrl, settings);
  } catch (mailErr) {
    return { success: true, url: voteUrl, message: '登録完了。メール送信失敗のためURLを直接お伝えください。' };
  }
  return { success: true, url: voteUrl, message: '登録完了。案内メールを送信しました。' };
}

// =============================================================================
// 機能C: 投票フォームデータ取得・投票記録
// =============================================================================

function getVoteFormData(token) {
  var validation = _validateToken(token);
  if (!validation.valid) {
    return { valid: false, message: validation.message, settings: null };
  }

  var settings = _getSettings();
  return {
    valid:   true,
    message: 'OK',
    settings: {
      votes:    settings.votes,
      deadline: settings.deadline ? Utilities.formatDate(new Date(settings.deadline), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : ''
    }
  };
}

/**
 * 投票を記録する（モード分岐）
 */
function submitVote(token, choices) {
  try {
    var settings = _getSettings();
    
    // 設定シートA6の値によって処理を切り替え
    if (settings.mode === '高速モード') {
      _recordVoteCache(token, choices);
    } else {
      _recordVoteDirect(token, choices);
    }
    
    return { success: true, message: '投票が完了しました。ご参加ありがとうございました。' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 【通常モード】スプレッドシートに直接書き込む。小・中規模向け。
 */
function _recordVoteDirect(token, choices) {
  var settings = _getSettings();

  for (var v = 0; v < settings.votes.length; v++) {
    if (!choices[v] || choices[v] === '') {
      throw new Error('「' + settings.votes[v].title + '」の選択肢を選んでください。');
    }
  }

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  var ballotSheet = ss.getSheetByName(SHEET_BALLOT_BOX);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var validation = _validateToken(token);
    if (!validation.valid) throw new Error(validation.message);

    var alreadyVoted = rosterSheet.getRange(validation.row, COL_VOTED).getValue();
    if (alreadyVoted === true) throw new Error('このトークンは既に使用されています。');

    var row = [new Date()];
    for (var i = 0; i < settings.votes.length; i++) {
      row.push(choices[i] || '');
    }
    ballotSheet.appendRow(row);
    rosterSheet.getRange(validation.row, COL_VOTED).setValue(true);

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

/**
 * 【高速モード】CacheServiceに一時保存する。大規模一斉投票向け。
 */
function _recordVoteCache(token, choices) {
  var settings = _getSettings();

  for (var v = 0; v < settings.votes.length; v++) {
    if (!choices[v] || choices[v] === '') {
      throw new Error('「' + settings.votes[v].title + '」の選択肢を選んでください。');
    }
  }

  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  lock.waitLock(3000);
  
  try {
    var cacheVotedKey = 'VOTED_' + token;
    if (cache.get(cacheVotedKey)) {
      throw new Error('このトークンは既に使用されています。（処理中）');
    }

    var validation = _validateToken(token);
    if (!validation.valid) throw new Error(validation.message);

    cache.put(cacheVotedKey, 'true', 21600);

    var bufferKey = 'BUFFER_' + new Date().getTime() + '_' + token.substring(0, 8);
    var payload = { token: token, choices: choices, date: new Date().toISOString() };
    cache.put(bufferKey, JSON.stringify(payload), 21600);

    var queueStr = cache.get('MASTER_QUEUE') || '[]';
    var queue = JSON.parse(queueStr);
    queue.push(bufferKey);
    cache.put('MASTER_QUEUE', JSON.stringify(queue), 21600);

  } finally {
    lock.releaseLock();
  }
}

function _validateToken(token) {
  if (!token) return { valid: false, message: '投票URLが正しくありません。', row: null };

  var settings = _getSettings();
  if (settings.deadline && new Date() > new Date(settings.deadline)) {
    return { valid: false, message: '投票の受付は終了しました。', row: null };
  }

  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var rowToken = rosterSheet.getRange(i, COL_TOKEN).getValue();
    if (rowToken === token) {
      var voted = rosterSheet.getRange(i, COL_VOTED).getValue();
      if (voted === true) return { valid: false, message: 'このURLはすでに使用済みです。', row: i };
      return { valid: true, message: 'OK', row: i };
    }
  }
  return { valid: false, message: '有効なトークンが見つかりません。', row: null };
}

// =============================================================================
// 機能D: 自動集計・結果通知
// =============================================================================

/**
 * 投票箱を集計し、主催者メールアドレスに結果を通知する。
 */
function tallySendResults() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sysSheet = ss.getSheetByName(SHEET_SYSLOG);

  var alreadyDone = sysSheet.getRange('A1').getValue();
  if (alreadyDone === true) {
    Logger.log('集計は既に完了済みです。');
    try { SpreadsheetApp.getUi().alert('集計は既に完了しています。'); } catch (e) {}
    return;
  }

  var settings    = _getSettings();
  var ballotSheet = ss.getSheetByName(SHEET_BALLOT_BOX);
  var lastRow     = ballotSheet.getLastRow();

  if (lastRow < 1) {
    Logger.log('集計: 投票データがありません。');
    try { SpreadsheetApp.getUi().alert('投票データがまだありません。'); } catch (e) {}
    return;
  }

  var numVotes = settings.votes.length;
  var tallies = settings.votes.map(function(vote) {
    var t = {};
    vote.options.forEach(function(opt) { t[opt] = 0; });
    return t;
  });

  var totalVotes = 0;
  var allData = ballotSheet.getRange(1, 1, lastRow, numVotes + 1).getValues();

  allData.forEach(function(rowData) {
    if (!rowData[0]) return; 
    totalVotes++;
    for (var v = 0; v < numVotes; v++) {
      var choice = rowData[v + 1]; 
      if (choice && tallies[v].hasOwnProperty(choice)) {
        tallies[v][choice]++;
      }
    }
  });

  var now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var body = '【Votely 投票結果通知】\n\n';
  body += '集計日時: ' + now + '\n';
  body += '総投票数: ' + totalVotes + ' 票\n\n';

  settings.votes.forEach(function(vote, idx) {
    body += '■ ' + vote.title + '\n';
    Object.keys(tallies[idx])
      .sort(function(a, b) { return tallies[idx][b] - tallies[idx][a]; })
      .forEach(function(opt) {
        body += '  ' + opt + ': ' + tallies[idx][opt] + ' 票\n';
      });
    body += '\n';
  });

  if (settings.organizerEmails.length > 0) {
    var mailOptions = {
      to:      settings.organizerEmails[0],
      subject: '【Votely】投票結果のお知らせ',
      body:    body
    };
    if (settings.organizerEmails.length > 1) {
      mailOptions.bcc = settings.organizerEmails.slice(1).join(',');
    }
    MailApp.sendEmail(mailOptions);
  }

  sysSheet.getRange('A1').setValue(true);
  SpreadsheetApp.flush();

  // バッチ処理トリガーを自動解除
  try { stopVotingTrigger(); } catch (e) {}

  Logger.log('集計・通知が完了しました。\n' + body);
  try {
    SpreadsheetApp.getUi().alert('集計・結果通知が完了しました。バッチ処理も停止しました。\n\n' + body);
  } catch (e) {}
}

// =============================================================================
// トリガーセットアップ関連
// =============================================================================

function setupTrigger() {
  var settings = _getSettings();
  if (!settings.deadline) {
    SpreadsheetApp.getUi().alert('設定シートのA2に締め切り日時が設定されていません。');
    return;
  }

  var deadline = new Date(settings.deadline);
  if (isNaN(deadline.getTime())) {
    SpreadsheetApp.getUi().alert('締め切り日時の形式が正しくありません。');
    return;
  }

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tallySendResults') ScriptApp.deleteTrigger(t);
  });

  var triggerTime = new Date(deadline.getTime() + 5 * 60 * 1000);
  ScriptApp.newTrigger('tallySendResults').timeBased().at(triggerTime).create();

  var triggerTimeStr = Utilities.formatDate(triggerTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  SpreadsheetApp.getUi().alert('集計トリガーをセットしました。\n実行予定日時: ' + triggerTimeStr);
}

function startVotingTrigger() {
  var functionName = 'processVoteBuffer';
  _deleteTriggersByName(functionName);
  ScriptApp.newTrigger(functionName).timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert('【投票開始】\n1分おきのバッチ処理トリガーをセットしました。\n終了後は「⏹ 投票終了」を実行してください。');
}

function stopVotingTrigger() {
  var functionName = 'processVoteBuffer';
  var deletedCount = _deleteTriggersByName(functionName);
  try {
    if (deletedCount > 0) SpreadsheetApp.getUi().alert('【投票終了】\nバッチ処理トリガーを停止しました。');
    else SpreadsheetApp.getUi().alert('稼働中のバッチ処理トリガーは見つかりませんでした。');
  } catch (e) {} // トリガーからの自動解除時はUIを出さない
}

function _deleteTriggersByName(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  return count;
}

// =============================================================================
// 高速モード バッチ処理
// =============================================================================

/**
 * キャッシュに溜まったデータをスプレッドシートに一括書き込みする。
 */
function processVoteBuffer() {
  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  
  if (!lock.tryLock(10000)) return; 

  try {
    var queueStr = cache.get('MASTER_QUEUE');
    if (!queueStr || queueStr === '[]') return;
    
    var queue = JSON.parse(queueStr);
    cache.put('MASTER_QUEUE', '[]', 21600);

    var cachedData = cache.getAll(queue);
    var ballotRows = [];
    var votedTokens = [];

    for (var i = 0; i < queue.length; i++) {
      var key = queue[i];
      if (cachedData[key]) {
        var data = JSON.parse(cachedData[key]);
        var row = [new Date(data.date)].concat(data.choices);
        ballotRows.push(row);
        votedTokens.push(data.token);
        cache.remove(key); 
      }
    }

    if (ballotRows.length === 0) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ballotSheet = ss.getSheetByName(SHEET_BALLOT_BOX);
    var startRow = ballotSheet.getLastRow() + 1;
    ballotSheet.getRange(startRow, 1, ballotRows.length, ballotRows[0].length).setValues(ballotRows);

    _batchUpdateRosterVotedFlags(ss, votedTokens);

  } catch(e) {
    Logger.log('バッチ処理エラー: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function _batchUpdateRosterVotedFlags(ss, votedTokens) {
  var rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  var lastRow = rosterSheet.getLastRow();
  if (lastRow < 2) return;

  var tokensRange = rosterSheet.getRange(2, COL_TOKEN, lastRow - 1, 1);
  var votedRange = rosterSheet.getRange(2, COL_VOTED, lastRow - 1, 1);
  var tokenValues = tokensRange.getValues();
  var votedValues = votedRange.getValues();

  var tokenSet = {};
  votedTokens.forEach(function(t) { tokenSet[t] = true; });

  var isUpdated = false;
  for (var i = 0; i < tokenValues.length; i++) {
    if (tokenSet[tokenValues[i][0]]) {
      votedValues[i][0] = true;
      isUpdated = true;
    }
  }

  if (isUpdated) {
    votedRange.setValues(votedValues);
  }
}

// =============================================================================
// ユーティリティ（内部関数）
// =============================================================================

function _getSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  var organizerEmails = _parseEmails(sheet.getRange('A1').getValue());
  var deadline        = sheet.getRange('A2').getValue();
  var adminPassword   = sheet.getRange('A3').getValue();
  var appUrl          = _normalizeAppUrl(String(sheet.getRange('A4').getValue()).trim());
  var pagesUrl        = String(sheet.getRange('A5').getValue()).trim();
  var mode            = String(sheet.getRange('A6').getValue()).trim() || '通常モード';

  var maxRow = sheet.getLastRow();
  var votes   = [];
  var col     = 2;

  while (true) {
    var title = sheet.getRange(1, col).getValue();
    if (title === '' || title === null || title === undefined) break;

    var options = [];
    for (var r = 2; r <= maxRow; r++) {
      var val = sheet.getRange(r, col).getValue();
      if (val !== '' && val !== null && val !== undefined) {
        options.push(String(val));
      }
    }
    if (options.length > 0) votes.push({ title: String(title), options: options });
    col++;
  }

  return {
    organizerEmails: organizerEmails,
    deadline:        deadline,
    adminPassword:   adminPassword,
    appUrl:          appUrl,
    pagesUrl:        pagesUrl,
    mode:            mode,
    votes:           votes
  };
}

function _generateUUID() {
  return Utilities.getUuid();
}

function _generateTokenForRow(row, sheet, pagesUrl) {
  var token   = _generateUUID();
  var base    = String(pagesUrl).replace(/\/$/, '');
  var voteUrl = base + '/index.html?token=' + token;

  sheet.getRange(row, COL_TOKEN).setValue(token);
  sheet.getRange(row, COL_URL).setValue(voteUrl);
  sheet.getRange(row, COL_VOTED).setValue(false);

  SpreadsheetApp.flush();
  return voteUrl;
}

function _sendInvitationEmail(email, voteUrl, settings) {
  var deadlineStr = settings.deadline ? Utilities.formatDate(new Date(settings.deadline), 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm') : '（未設定）';
  var subject = '【Votely】投票のご案内';
  var body = [
    'このたびは投票へのご参加ありがとうございます。',
    '',
    '以下の専用URLからご投票ください。',
    'このURLはあなた専用です。他の方と共有しないでください。',
    '',
    '▼ 投票URL',
    voteUrl,
    '',
    '締め切り: ' + deadlineStr,
    '',
    '※ このURLは一度しか使用できません。',
    '※ 投票は匿名で処理されます（誰が何に投票したかは記録されません）。'
  ].join('\n');
  MailApp.sendEmail({ to: email, subject: subject, body: body });
}

function _normalizeAppUrl(url) {
  if (!url) return '';
  url = url.replace(/\/a\/[^\/]+\/macros\//, '/macros/');
  url = url.replace(/\/macros\/u\/\d+\/s\//, '/macros/s/');
  return url;
}

function _parseEmails(raw) {
  if (!raw) return [];
  return String(raw).split(',')
    .map(function(addr) { return addr.replace(/[\s\u3000]+/g, ''); })
    .filter(function(addr) { return addr.length > 0; });
}
