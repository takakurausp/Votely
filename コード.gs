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
    .addSeparator()
    .addItem('⑦ 🎫 当日参加者用チケット一括発行（PDF印刷）', 'generateGuestTickets')
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

  // A6セルにプルダウン（データ検証）を設定（初回・再実行どちらでも適用）
  var modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['通常モード', '高速モード'], true)
    .setAllowInvalid(false)
    .setHelpText('通常モード（小〜中規模）または 高速モード（大規模一斉投票）を選択してください。')
    .build();
  settingsSheet.getRange('A6').setDataValidation(modeRule);

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

/**
 * 事務局パスワードを検証する。
 * 失敗回数を CacheService で数え、5 回失敗で 5 分間ロックアウトする。
 * GAS は IP を取れないため、ロックアウトはスクリプト全体に対するグローバルなもの。
 * 正しいパスワードでログイン中の事務局担当は失敗カウンタを増やさないので、
 * 攻撃者の総当たりが続いている間も registerAttendee は通り続ける（ロック中を除く）。
 */
function verifyAdminPassword(password) {
  var cache   = CacheService.getScriptCache();
  var FAIL_KEY   = 'ADMIN_AUTH_FAILS';
  var LOCK_KEY   = 'ADMIN_AUTH_LOCKED';
  var MAX_FAILS  = 5;
  var FAIL_TTL   = 600;  // 失敗カウンタの保持秒数（10 分）
  var LOCK_TTL   = 300;  // ロックアウト時間（5 分）

  if (cache.get(LOCK_KEY)) {
    return false;
  }

  var settings = _getSettings();
  var ok = String(settings.adminPassword).trim() === String(password).trim();

  if (ok) {
    cache.remove(FAIL_KEY);
    return true;
  }

  var fails = Number(cache.get(FAIL_KEY) || '0') + 1;
  if (fails >= MAX_FAILS) {
    cache.put(LOCK_KEY, '1', LOCK_TTL);
    cache.remove(FAIL_KEY);
  } else {
    cache.put(FAIL_KEY, String(fails), FAIL_TTL);
  }
  return false;
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
  var settings = _getSettings();
  var validation = _validateToken(token, settings);
  if (!validation.valid) {
    return { valid: false, message: validation.message, settings: null };
  }

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
      _recordVoteCache(token, choices, settings);
    } else {
      _recordVoteDirect(token, choices, settings);
    }

    return { success: true, message: '投票が完了しました。ご参加ありがとうございました。' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 【通常モード】スプレッドシートに直接書き込む。小・中規模向け。
 */
function _recordVoteDirect(token, choices, settings) {
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
    var validation = _validateToken(token, settings);
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
 *
 * 【設計メモ：Option C（カウンタ＋個別キー方式）】
 * 旧実装は MASTER_QUEUE に配列をシリアライズして毎回 read-modify-write していたため、
 * 件数が増えるとロック内の JSON parse/stringify がスループットの上限を決めていた。
 * GAS の ScriptLock はグローバルで粒度を下げられないので、ロック内の作業量を
 * 「カウンタ get → BUFFER_n put → カウンタ put」の O(1) に固定し、
 * バッチ側 (processVoteBuffer) で 0..n-1 を一括取得する形に変更している。
 */
function _recordVoteCache(token, choices, settings) {
  for (var v = 0; v < settings.votes.length; v++) {
    if (!choices[v] || choices[v] === '') {
      throw new Error('「' + settings.votes[v].title + '」の選択肢を選んでください。');
    }
  }

  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var cacheVotedKey = 'VOTED_' + token;
    if (cache.get(cacheVotedKey)) {
      throw new Error('このトークンは既に使用されています。（処理中）');
    }

    var validation = _validateToken(token, settings);
    if (!validation.valid) throw new Error(validation.message);

    cache.put(cacheVotedKey, 'true', 21600);

    // Option C: カウンタを進めて個別キーに書き込むだけ。配列を持ち回さない。
    var n = Number(cache.get('QUEUE_LEN') || '0');
    var payload = { token: token, choices: choices, date: new Date().toISOString() };
    cache.put('BUFFER_' + n, JSON.stringify(payload), 21600);
    cache.put('QUEUE_LEN', String(n + 1), 21600);

  } finally {
    lock.releaseLock();
  }
}

/**
 * トークンの妥当性を検証する。settings は呼び出し側で 1 度だけ取得して渡す前提。
 * 名簿は (token, url, voted) の 3 列を 1 度のレンジ取得で読み出し、
 * セル単位の getValue ループを避ける。
 */
function _validateToken(token, settings) {
  if (!token) return { valid: false, message: '投票URLが正しくありません。', row: null };

  if (settings.deadline && new Date() > new Date(settings.deadline)) {
    return { valid: false, message: '投票の受付は終了しました。', row: null };
  }

  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();
  if (lastRow < 2) return { valid: false, message: '有効なトークンが見つかりません。', row: null };

  var numCols = COL_VOTED - COL_TOKEN + 1;
  var rows    = rosterSheet.getRange(2, COL_TOKEN, lastRow - 1, numCols).getValues();
  var votedIdx = COL_VOTED - COL_TOKEN;

  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === token) {
      var voted = rows[i][votedIdx];
      if (voted === true) return { valid: false, message: 'このURLはすでに使用済みです。', row: i + 2 };
      return { valid: true, message: 'OK', row: i + 2 };
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
  var dataRows = lastRow - 1; // ヘッダー行を除く
  if (dataRows < 1) {
    try { SpreadsheetApp.getUi().alert('投票データがまだありません。'); } catch (e) {}
    return;
  }
  var allData = ballotSheet.getRange(2, 1, dataRows, numVotes + 1).getValues();

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

  // 高速モードかつトリガーが存在する場合のみバッチ処理トリガーを自動解除
  if (settings.mode === '高速モード' && _triggerExists('processVoteBuffer')) {
    _deleteTriggersByName('processVoteBuffer');
  }

  Logger.log('集計・通知が完了しました。\n' + body);
  try {
    var doneMsg = settings.mode === '高速モード'
      ? '集計・結果通知が完了しました。バッチ処理も停止しました。\n\n' + body
      : '集計・結果通知が完了しました。\n\n' + body;
    SpreadsheetApp.getUi().alert(doneMsg);
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
  var exists = _triggerExists('processVoteBuffer');
  if (!exists) {
    try { SpreadsheetApp.getUi().alert('稼働中のバッチ処理トリガーは見つかりませんでした。'); } catch (e) {}
    return;
  }
  _deleteTriggersByName('processVoteBuffer');
  try { SpreadsheetApp.getUi().alert('【投票終了】\nバッチ処理トリガーを停止しました。'); } catch (e) {}
}

/**
 * 指定した関数名のトリガーが存在するか確認する。
 */
function _triggerExists(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) return true;
  }
  return false;
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

  // ScriptLock はグローバルなので、このバッチが走っている間は新規 submit は待たされる。
  // 1 分おきトリガー × 数秒で済む処理量という前提なので、ロックは全工程を通じて保持する。
  if (!lock.tryLock(10000)) return;

  try {
    var n = Number(cache.get('QUEUE_LEN') || '0');
    if (n === 0) return;

    var keys = [];
    for (var i = 0; i < n; i++) keys.push('BUFFER_' + i);

    var cachedData = cache.getAll(keys) || {};
    var ballotRows = [];
    var votedTokens = [];

    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      if (cachedData[key]) {
        var data = JSON.parse(cachedData[key]);
        var row = [new Date(data.date)].concat(data.choices);
        ballotRows.push(row);
        votedTokens.push(data.token);
      }
    }

    // 取り出し済みなのでカウンタを 0 に戻し、個別キーも掃除する。
    // ロックを保持したまま行うため、この間に新規 submit が入り込んで
    // BUFFER_0 を上書きする心配はない。
    cache.removeAll(keys);
    cache.put('QUEUE_LEN', '0', 21600);

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
// 機能E: 当日参加者用チケット一括発行（PDF印刷）
// =============================================================================

/**
 * 当日参加者用のトークンを N 件まとめて発行し、QRコード付きのチケットを
 * モーダルダイアログに表示する。ユーザーはブラウザの「PDFとして保存／印刷」で
 * A4 用紙に 10 枚（5 行 × 2 列）レイアウトで出力できる。
 *
 * 名簿シートには「guest_001」「guest_002」のような連番ラベルで追記される。
 * 既存の guest_NNN を走査して、最大値+1 から続きの番号を採番する。
 */
function generateGuestTickets() {
  var html = HtmlService.createHtmlOutput(_buildGuestTicketsDialogHtml())
    .setWidth(640)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '🎫 当日参加者用チケット一括発行');
}

/**
 * モーダル側の JavaScript から google.script.run 経由で呼ばれるサーバーハンドラ。
 * トークンを発行し、名簿に一括追記し、印刷用データを返す。
 */
function createGuestTickets(params) {
  var num = parseInt(params && params.num, 10);
  if (!num || num <= 0) throw new Error('発行枚数は1以上の整数を指定してください。');
  if (num > 200) throw new Error('一度に生成できるのは200枚までです。');

  var title = String(params.title || 'Votely 投票チケット');
  var desc  = String(params.desc  || '');

  var settings = _getSettings();
  var pagesUrl = settings.pagesUrl || settings.appUrl;
  if (!pagesUrl) {
    throw new Error('設定シートの A5（または A4）に投票画面URLを設定してください。');
  }

  if (settings.deadline && new Date() > new Date(settings.deadline)) {
    throw new Error('投票の締め切りを過ぎています。チケットを発行できません。');
  }

  var deadlineStr = settings.deadline
    ? Utilities.formatDate(new Date(settings.deadline), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
    : '';

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName(SHEET_ROSTER);

  // 既存の guest_NNN を走査して、続きの番号から採番する
  var startNum = _getNextGuestNumber(rosterSheet);
  var base     = String(pagesUrl).replace(/\/$/, '');

  var rows    = [];
  var tickets = [];
  for (var i = 0; i < num; i++) {
    var n       = startNum + i;
    var label   = 'guest_' + _padNum(n, 3);
    var token   = Utilities.getUuid();
    var voteUrl = base + '/index.html?token=' + token;
    rows.push([label, token, voteUrl, false]);
    tickets.push({ label: label, token: token, url: voteUrl });
  }

  // 名簿に一括追記（appendRow ループより速く、ロック競合も少ない）
  var startRow = rosterSheet.getLastRow() + 1;
  rosterSheet.getRange(startRow, 1, rows.length, 4).setValues(rows);
  SpreadsheetApp.flush();

  return {
    tickets:  tickets,
    title:    title,
    desc:     desc,
    deadline: deadlineStr
  };
}

/**
 * 名簿シートの A 列を走査して「guest_NNN」形式の最大番号 + 1 を返す。
 * 1 件もなければ 1 を返す。
 */
function _getNextGuestNumber(rosterSheet) {
  var lastRow = rosterSheet.getLastRow();
  if (lastRow < 2) return 1;
  var values = rosterSheet.getRange(2, COL_EMAIL, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < values.length; i++) {
    var m = String(values[i][0]).match(/^guest_(\d+)$/);
    if (m) {
      var n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
  }
  return max + 1;
}

function _padNum(n, w) {
  var s = String(n);
  while (s.length < w) s = '0' + s;
  return s;
}

/**
 * チケット発行モーダルの HTML を組み立てる。
 * - 上部: 入力フォーム（タイトル / 説明 / 枚数）と「発行」ボタン
 * - 下部: 発行後にチケットを 2 列グリッドで描画（QR は api.qrserver.com を使用）
 * - @media print で印刷時はフォーム部分を隠して A4 レイアウトに整える
 */
function _buildGuestTicketsDialogHtml() {
  return [
    '<!doctype html><html lang="ja"><head><meta charset="utf-8"><style>',
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;margin:0;padding:14px;font-size:13px;color:#212529;}',
    'h2{margin:0 0 10px;font-size:15px;color:#0d6efd;}',
    'label{display:block;margin:8px 0 3px;font-weight:600;}',
    'input[type=text],input[type=number]{width:100%;box-sizing:border-box;padding:6px 8px;border:1px solid #ced4da;border-radius:4px;font-size:13px;}',
    'button{margin-top:12px;padding:8px 18px;background:#0d6efd;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:600;}',
    'button.print{background:#28a745;}',
    'button:disabled{background:#999;cursor:wait;}',
    '#status{color:#0d6efd;font-weight:600;margin-top:10px;min-height:1em;}',
    '#err{color:#dc3545;margin-top:6px;min-height:1em;}',
    '#result{margin-top:14px;display:none;}',
    '.tickets{display:grid;grid-template-columns:1fr 1fr;gap:6mm;margin-top:10px;}',
    '.ticket{border:1px solid #999;padding:5mm;display:flex;justify-content:space-between;align-items:flex-start;break-inside:avoid;background:#fff;}',
    '.ticket .info{flex:1;padding-right:3mm;min-width:0;}',
    '.ticket .t-title{font-size:11pt;font-weight:700;margin:0 0 2mm;line-height:1.25;}',
    '.ticket .t-desc{font-size:8pt;margin:0 0 2mm;color:#333;line-height:1.3;}',
    '.ticket .t-deadline{font-size:8pt;color:#c0392b;margin:0 0 2mm;}',
    '.ticket .t-no{font-size:7pt;color:#555;margin:2mm 0 0;}',
    '.ticket .t-token{font-size:5.5pt;color:#888;word-break:break-all;margin:0;}',
    '.ticket img{width:30mm;height:30mm;flex-shrink:0;}',
    '@media print {',
    '  body{padding:0;}',
    '  .form-area{display:none !important;}',
    '  #result{display:block !important;margin:0;}',
    '  .tickets{gap:4mm;}',
    '  .ticket{page-break-inside:avoid;}',
    '  button.print{display:none;}',
    '  @page{size:A4;margin:8mm;}',
    '}',
    '</style></head><body>',
    '<div class="form-area">',
    '<h2>🎫 当日参加者チケット一括発行</h2>',
    '<label>チケットタイトル</label>',
    '<input type="text" id="title" value="Votely 投票チケット">',
    '<label>説明文</label>',
    '<input type="text" id="desc" value="QRコードを読み取って投票してください">',
    '<label>発行枚数（1〜200）</label>',
    '<input type="number" id="num" value="10" min="1" max="200">',
    '<button id="genBtn" onclick="onGenerate()">トークン発行＆チケット表示</button>',
    '<p id="status"></p>',
    '<p id="err"></p>',
    '</div>',
    '<div id="result">',
    '<button class="print" onclick="window.print()">🖨 PDFとして保存／印刷</button>',
    '<div class="tickets" id="ticketsBox"></div>',
    '</div>',
    '<script>',
    'function escapeHtml(s){',
    '  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/\'/g,"&#039;");',
    '}',
    'function onGenerate(){',
    '  var btn=document.getElementById("genBtn");',
    '  var status=document.getElementById("status");',
    '  var err=document.getElementById("err");',
    '  err.textContent="";',
    '  btn.disabled=true;status.textContent="生成中...";',
    '  var params={',
    '    num:document.getElementById("num").value,',
    '    title:document.getElementById("title").value,',
    '    desc:document.getElementById("desc").value',
    '  };',
    '  google.script.run',
    '    .withSuccessHandler(function(data){',
    '      btn.disabled=false;status.textContent="";',
    '      render(data);',
    '    })',
    '    .withFailureHandler(function(e){',
    '      btn.disabled=false;status.textContent="";',
    '      err.textContent="エラー: "+(e && e.message ? e.message : e);',
    '    })',
    '    .createGuestTickets(params);',
    '}',
    'function render(data){',
    '  var box=document.getElementById("ticketsBox");',
    '  box.innerHTML="";',
    '  data.tickets.forEach(function(t){',
    '    var qr="https://api.qrserver.com/v1/create-qr-code/?size=200x200&data="+encodeURIComponent(t.url);',
    '    var html=',
    '      \'<div class="ticket">\'+',
    '        \'<div class="info">\'+',
    '          \'<p class="t-title">\'+escapeHtml(data.title)+\'</p>\'+',
    '          (data.desc?\'<p class="t-desc">\'+escapeHtml(data.desc)+\'</p>\':"")+',
    '          (data.deadline?\'<p class="t-deadline">締切: \'+escapeHtml(data.deadline)+\'</p>\':"")+',
    '          \'<p class="t-no">No: \'+escapeHtml(t.label)+\'</p>\'+',
    '          \'<p class="t-token">Token: \'+escapeHtml(t.token)+\'</p>\'+',
    '        \'</div>\'+',
    '        \'<img src="\'+qr+\'" alt="QR">\'+',
    '      \'</div>\';',
    '    box.insertAdjacentHTML("beforeend",html);',
    '  });',
    '  document.getElementById("result").style.display="block";',
    '}',
    '</script>',
    '</body></html>'
  ].join('\n');
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
