/**
 * AIライフコーチ — バックエンド (Google Apps Script)
 *
 * 構成:
 *  - Users シート:    id | name | createdAt
 *  - Messages シート: id | userId | role | content | timestamp
 *
 * 必須の設定 (スクリプト プロパティ):
 *  - GEMINI_API_KEY : Google AI Studio で取得した Gemini API キー (必須)
 *  - GEMINI_MODEL   : 使用するモデル名 (省略可。未設定なら gemini-2.5-flash)
 *
 * 設定方法: Apps Script エディタ → 左メニューの「プロジェクトの設定」(歯車アイコン)
 *          → 「スクリプト プロパティ」→ プロパティを追加
 *
 * デプロイ: デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *          → 実行ユーザー「自分」、アクセスできるユーザー「全員」
 */

const USERS_SHEET = 'Users';
const MESSAGES_SHEET = 'Messages';
const MAX_HISTORY_FOR_AI = 40; // Geminiに渡す直近メッセージ数の上限

const SYSTEM_PROMPT = `あなたは「人生を変えるための対話」に付き添うAIライフコーチです。次の姿勢で会話してください。

- まず相手の話を丁寧に受け止め、気持ちや状況に共感する。すぐに評価したり否定したりしない。
- アドバイスや「正解」を急いで提示せず、オープンな質問を通して相手自身の中にある考えや気づきを引き出す。
- 会話が深まってきたら、相手のペースに合わせて、自然な流れで以下のステップに進む(順番は固定せず、必要に応じて戻ってもよい):
  1. 現状の整理: 今どんな状況で、何を感じているか
  2. 理想・目標の明確化: 本当はどうなりたいか、何を大切にしたいか
  3. 行動計画: 次の一歩として、現実的に何ができそうか
  4. 振り返り・フォローアップ: 次回以降、進捗や気持ちの変化を確認し、必要なら計画を見直す
- 説教的・批判的にならず、常に相手を尊重し、励ます。
- 返答は長すぎないようにする(おおよそ1〜3段落)。対話のキャッチボールを大切にする。
- 自傷・自殺念慮など深刻なメンタルヘルスの問題が見られた場合は、専門の相談窓口や専門家に相談することを優しく勧める。
- 日本語で会話する。`;

/**
 * すべてのリクエストはこの doGet で処理する (JSONP)。
 * パラメータ:
 *   callback : JSONPのコールバック関数名
 *   action   : 'login' | 'getHistory' | 'sendMessage'
 *   payload  : JSON文字列 (URLエンコード済み)
 */
function doGet(e) {
  let result;
  try {
    const action = e.parameter.action;
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};

    switch (action) {
      case 'login':
        result = login(payload);
        break;
      case 'getHistory':
        result = getHistory(payload);
        break;
      case 'sendMessage':
        result = sendMessage(payload);
        break;
      default:
        result = { error: '不明なactionです: ' + action };
    }
  } catch (err) {
    result = { error: 'サーバーエラー: ' + err.message };
  }

  const callback = e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('シートが見つかりません: ' + name);
  return sheet;
}

/** 名前でログイン。なければ新規ユーザーを作成。 */
function login(payload) {
  const name = (payload.name || '').toString().trim();
  if (!name) return { error: 'お名前を入力してください。' };

  const sheet = getSheet_(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = Number(data[i][0]) || 0;
    if (id > maxId) maxId = id;
    if (String(data[i][1]).trim() === name) {
      return { userId: id, name: String(data[i][1]) };
    }
  }
  const newId = maxId + 1;
  sheet.appendRow([newId, name, new Date()]);
  return { userId: newId, name: name };
}

/** 指定ユーザーの会話履歴を取得 */
function getHistory(payload) {
  const userId = Number(payload.userId);
  if (!userId) return { error: 'userIdが必要です。' };

  const sheet = getSheet_(MESSAGES_SHEET);
  const data = sheet.getDataRange().getValues();
  const messages = [];
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][1]) === userId) {
      messages.push({
        role: String(data[i][2]),
        content: String(data[i][3]),
        timestamp: data[i][4] instanceof Date ? data[i][4].toISOString() : String(data[i][4])
      });
    }
  }
  return { messages: messages };
}

/** ユーザーのメッセージを保存し、Geminiに問い合わせて返信を保存・返却 */
function sendMessage(payload) {
  const userId = Number(payload.userId);
  const userMessage = (payload.message || '').toString().trim();
  if (!userId) return { error: 'userIdが必要です。' };
  if (!userMessage) return { error: 'メッセージが空です。' };

  const sheet = getSheet_(MESSAGES_SHEET);
  const data = sheet.getDataRange().getValues();

  // 既存履歴 (Gemini用、直近 MAX_HISTORY_FOR_AI 件)
  const history = [];
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = Number(data[i][0]) || 0;
    if (id > maxId) maxId = id;
    if (Number(data[i][1]) === userId) {
      history.push({ role: String(data[i][2]), content: String(data[i][3]) });
    }
  }
  const recentHistory = history.slice(-MAX_HISTORY_FOR_AI);

  // ユーザーメッセージを保存
  const now = new Date();
  sheet.appendRow([maxId + 1, userId, 'user', userMessage, now]);

  // Gemini API キーの確認
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { error: 'Gemini APIキーが設定されていません。Apps Scriptの「プロジェクトの設定」→「スクリプト プロパティ」で GEMINI_API_KEY を設定してください。' };
  }
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';

  // Gemini API 用のcontents組み立て
  const contents = recentHistory.map(h => ({
    role: (h.role === 'model') ? 'model' : 'user',
    parts: [{ text: h.content }]
  }));
  contents.push({ role: 'user', parts: [{ text: userMessage }] });

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;
  const requestBody = {
    contents: contents,
    systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    return { error: 'AI APIエラー (' + code + '): ' + text.substring(0, 300) };
  }

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    return { error: 'AI応答のJSON解析に失敗しました。' };
  }

  let aiReply;
  try {
    aiReply = json.candidates[0].content.parts[0].text;
  } catch (e) {
    const finishReason = (json.candidates && json.candidates[0] && json.candidates[0].finishReason) || '不明';
    return { error: 'AIからの応答を取得できませんでした (finishReason: ' + finishReason + ')' };
  }

  // AI返信を保存
  const aiTime = new Date();
  sheet.appendRow([maxId + 2, userId, 'model', aiReply, aiTime]);

  return { reply: aiReply, timestamp: aiTime.toISOString() };
}

/**
 * 初回セットアップ用: Users / Messages シートをヘッダー付きで作成する。
 * Apps Scriptエディタでこの関数を一度だけ実行してください。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let usersSheet = ss.getSheetByName(USERS_SHEET);
  if (!usersSheet) usersSheet = ss.insertSheet(USERS_SHEET);
  usersSheet.clear();
  usersSheet.getRange(1, 1, 1, 3).setValues([['id', 'name', 'createdAt']]);

  let messagesSheet = ss.getSheetByName(MESSAGES_SHEET);
  if (!messagesSheet) messagesSheet = ss.insertSheet(MESSAGES_SHEET);
  messagesSheet.clear();
  messagesSheet.getRange(1, 1, 1, 5).setValues([['id', 'userId', 'role', 'content', 'timestamp']]);

  // デフォルトの「シート1」が残っていれば削除
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('セットアップ完了: Users / Messages シートを作成しました。');
}
