/**
 * 問い合わせフォーム 自動通知・日次集計ツール
 *
 * - フォームに回答が届いたら、内容をすぐにメールで通知する
 * - 毎朝、過去24時間の回答を種別ごとに集計し、シートに記録してメールで送る
 *
 * フォームの回答先スプレッドシートに紐づけて使う（拡張機能 > Apps Script）。
 */

// ===== 設定（ここだけ書き換えれば使える） =====
const CONFIG = {
  // 通知先。空なら、スクリプトを実行したアカウントのアドレスに送る
  NOTIFY_EMAIL: '',
  // フォームの回答が入るシート名
  RESPONSE_SHEET: 'フォームの回答 1',
  // 集計結果を書き込むシート名（なければ自動で作る）
  SUMMARY_SHEET: '日次集計',
  // 種別ごとに集計したい質問のタイトル
  CATEGORY_QUESTION: 'お問い合わせ種別',
  // 日次集計を送る時刻（0〜23時）
  DAILY_HOUR: 8,
};

// ===== 初回に1回だけ実行する =====

/**
 * トリガー（自動実行の設定）を登録する。
 * 何度実行しても重複しないよう、このツールのトリガーを消してから作り直す。
 */
function setupTriggers() {
  const handlers = ['notifyNewResponse', 'sendDailySummary'];
  ScriptApp.getProjectTriggers()
    .filter((t) => handlers.includes(t.getHandlerFunction()))
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('notifyNewResponse')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.DAILY_HOUR)
    .create();

  Logger.log('トリガーを登録しました');
}

// ===== トリガーから呼ばれる処理 =====

/**
 * 新しい回答が届いたときに、内容をメールで通知する。
 * @param {Object} e フォーム送信イベント。e.namedValues は { 質問タイトル: [回答] } の形
 */
function notifyNewResponse(e) {
  const answers = e.namedValues;
  const category = firstValue(answers[CONFIG.CATEGORY_QUESTION]) || '未分類';

  const body = Object.keys(answers)
    .map((question) => `■ ${question}\n${answers[question].join(', ')}`)
    .join('\n\n');

  MailApp.sendEmail({
    to: getNotifyEmail(),
    subject: `【新着】お問い合わせ（${category}）`,
    body: `${body}\n\n---\n回答一覧: ${SpreadsheetApp.getActive().getUrl()}`,
  });
}

/**
 * 過去24時間の回答を種別ごとに数え、集計シートに記録してメールで送る。
 */
function sendDailySummary() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.RESPONSE_SHEET);
  if (!sheet) {
    throw new Error(`シート「${CONFIG.RESPONSE_SHEET}」が見つかりません。CONFIG を確認してください`);
  }

  const [header, ...rows] = sheet.getDataRange().getValues();
  const categoryIndex = header.indexOf(CONFIG.CATEGORY_QUESTION);
  const since = new Date(Date.now() - 24 * 60 * 60 * 1000);

  // 1列目はフォームが自動で入れるタイムスタンプ
  const recentRows = rows.filter((row) => row[0] instanceof Date && row[0] >= since);

  const counts = {};
  recentRows.forEach((row) => {
    const category = (categoryIndex >= 0 && row[categoryIndex]) || '未分類';
    counts[category] = (counts[category] || 0) + 1;
  });

  const breakdown = Object.entries(counts)
    .map(([category, count]) => `${category}: ${count}件`)
    .join(' / ');

  appendSummaryRow([new Date(), recentRows.length, breakdown || 'なし']);

  MailApp.sendEmail({
    to: getNotifyEmail(),
    subject: `【日次集計】過去24時間のお問い合わせ: ${recentRows.length}件`,
    body: [
      `過去24時間のお問い合わせは ${recentRows.length} 件でした。`,
      '',
      breakdown ? breakdown.split(' / ').map((line) => `・${line}`).join('\n') : '・なし',
      '',
      '---',
      `回答一覧: ${SpreadsheetApp.getActive().getUrl()}`,
    ].join('\n'),
  });
}

// ===== 動作確認用 =====

/**
 * フォームを送信しなくても通知メールを試せるよう、ダミーの回答で notifyNewResponse を呼ぶ。
 */
function testNotify() {
  notifyNewResponse({
    namedValues: {
      'タイムスタンプ': [new Date().toLocaleString('ja-JP')],
      'お名前': ['テスト 太郎'],
      'メールアドレス': ['test@example.com'],
      'お問い合わせ種別': ['予約について'],
      'お問い合わせ内容': ['来週土曜の18時に2名で予約できますか？'],
    },
  });
}

// ===== 内部で使う小さな関数 =====

function getNotifyEmail() {
  return CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
}

function firstValue(values) {
  return Array.isArray(values) ? values[0] : '';
}

function appendSummaryRow(row) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET) || ss.insertSheet(CONFIG.SUMMARY_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['集計日時', '件数', '内訳']);
  }
  sheet.appendRow(row);
}
