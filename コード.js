/**
 * ToDo メモ Webアプリ（system-f）v2 — 編集できる入れ子アウトライン
 * ============================================================================
 * 大タブ（生活 / ファイナンス / 音楽 / カメラ / お仕事）で分け、各タブの中を
 * 自由に入れ子（連絡／家／デスクワーク＞検索・手続き…／場所 など）で整理できる
 * チェックリスト。タブも中身も、追加・名前変更・削除・並び替え・階層変更が自由。
 * 「今日のチェック」タブで、毎日くり返すタスクをその日ごとにチェック（翌日リセット）。
 *
 * データは Google スプレッドシートに保存：
 *   ・「タブ」     … ID / 名前 / 種類 / 並び順
 *   ・「ノード」   … ID / タブID / 親ID / 内容 / 並び順 / 完了 / 折りたたみ / 毎日 / メモ
 *   ・「今日のチェック」… 日付 / ノードID
 *
 * デプロイ：スプレッドシート →［拡張機能］→［Apps Script］に本コードと
 *           「画面.html」を貼り付け →［デプロイ］→ ウェブアプリ。
 */

// ===== 設定 =================================================================
const CONFIG = {
  // 保存先。ID/URL を入れると固定。空''なら バウンドのアクティブブック →
  // 無ければ「ToDoメモ」を自動作成して記憶。
  SS_ID: '',
  TAB_SHEET: 'タブ',
  NODE_SHEET: 'ノード',
  CHECK_SHEET: '今日のチェック',
};

// 初回に用意する大タブと中身は buildSeed_()（このファイル下部）で定義。
// 既に流し込んだあと、最初から作り直したいときは reseedMyData() を実行。

// 新しいタブを作ったとき自動で入る「ひな形」（すべて後から編集可）
const SCAFFOLD_NORMAL = [
  { t: '連絡' },
  { t: '家' },
  { t: 'デスクワーク', c: [{ t: '検索' }, { t: '手続き' }, { t: '購入' }, { t: '検討' }, { t: '定期' }] },
  { t: '場所（お出かけリスト）' },
];
const SCAFFOLD_WORK = [
  { t: '人', c: [] },
  { t: 'デスクワーク', c: [] },
];

// ===== Webアプリ入口 ========================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('画面')
    .setTitle('ToDo メモ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🗒 ToDoメモ')
      .addItem('アプリURLを表示', 'showWebAppUrl')
      .addToUi();
  } catch (e) { /* スタンドアロンでは無視 */ }
}
function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    url
      ? '<p style="font-family:sans-serif">ToDoメモ アプリのURL（ブックマーク推奨）：</p>' +
        '<p><a href="' + url + '" target="_blank">' + url + '</a></p>'
      : '<p style="font-family:sans-serif">まだ公開されていません。［デプロイ］→［新しいデプロイ］→「ウェブアプリ」で公開してください。</p>')
    .setWidth(480).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, 'ToDoメモ');
}

// ===== 画面の初期データ ======================================================
function getData(dateStr) {
  const today = dateStr || todayStr_();
  ensureSeed_();
  ensureSubscriptionTab_();   // 「定期購入」大タブが無ければ一度だけ用意（既にあれば何もしない）
  ensurePrioritySheet_();     // 「優先順位」シートが無ければ一度だけ用意（AI連携の土台）
  return {
    today: today,
    tabs: readTabs_(),
    nodes: readNodes_(),
    todayChecks: readChecksForDate_(today),
  };
}

// ===== タブ操作 =============================================================
// 新タブ作成（idはクライアント生成）。種類に応じてひな形ノードも作る。
function addTab(id, name, type) {
  id = String(id || '').trim() || newId_();
  name = String(name || '').trim() || '新しいタブ';
  type = (type === 'work') ? 'work' : 'normal';
  return withLock_(function () {
    const sh = tabSheet_();
    const order = nextTabOrder_(sh);
    sh.appendRow([id, name, type, order]);
    insertScaffold_(id, (type === 'work' ? SCAFFOLD_WORK : SCAFFOLD_NORMAL), '');
    return { tabs: readTabs_(), nodes: readNodes_() };
  });
}

function renameTab(id, name) {
  name = String(name || '').trim();
  if (!name) throw new Error('タブ名が空です。');
  return withLock_(function () {
    const r = findRow_(tabSheet_(), id);
    if (r) tabSheet_().getRange(r, 2).setValue(name);
    return { id: id, name: name };
  });
}

function deleteTab(id) {
  return withLock_(function () {
    const r = findRow_(tabSheet_(), id);
    if (r) tabSheet_().deleteRow(r);
    // そのタブのノードと、そのチェックを全部消す
    const nsh = nodeSheet_();
    const v = nsh.getDataRange().getValues();
    const delIds = {};
    for (let i = v.length - 1; i >= 1; i--) {
      if (String(v[i][1]) === String(id)) { delIds[String(v[i][0])] = true; nsh.deleteRow(i + 1); }
    }
    removeChecksForNodes_(delIds);
    return { tabs: readTabs_() };
  });
}

function reorderTabs(idsInOrder) {
  return withLock_(function () {
    const sh = tabSheet_();
    const v = sh.getDataRange().getValues();
    const rowById = {};
    for (let i = 1; i < v.length; i++) rowById[String(v[i][0])] = i + 1;
    (idsInOrder || []).forEach(function (id, idx) {
      const r = rowById[String(id)];
      if (r) sh.getRange(r, 4).setValue(idx);
    });
    return { tabs: readTabs_() };
  });
}

// ===== ノード操作 ===========================================================
function addNode(id, tabId, parentId, order, text) {
  id = String(id || '').trim() || newId_();
  return withLock_(function () {
    nodeSheet_().appendRow([id, String(tabId || ''), String(parentId || ''),
      String(text || ''), Number(order) || 0, false, false, false, '', false]);
    return { id: id };
  });
}

// field: text / done / collapsed / daily / note / parentId / order
function updateNode(id, field, value) {
  const COL = { text: 4, order: 5, done: 6, collapsed: 7, daily: 8, note: 9, today: 10, parentId: 3, tabId: 2 };
  const c = COL[field];
  if (!c) throw new Error('不明なフィールド: ' + field);
  return withLock_(function () {
    const r = findRow_(nodeSheet_(), id);
    if (r) nodeSheet_().getRange(r, c).setValue(value);
    return { id: id, field: field };
  });
}

// 1ノード削除（子孫もまとめて削除）
function deleteNode(id) {
  return withLock_(function () {
    const sh = nodeSheet_();
    const v = sh.getDataRange().getValues();
    // 親子関係を作り、対象の子孫を集める
    const children = {};
    for (let i = 1; i < v.length; i++) {
      const p = String(v[i][2]);
      (children[p] = children[p] || []).push(String(v[i][0]));
    }
    const kill = {};
    (function collect(x) { kill[x] = true; (children[x] || []).forEach(collect); })(String(id));
    for (let i = v.length - 1; i >= 1; i--) {
      if (kill[String(v[i][0])]) sh.deleteRow(i + 1);
    }
    removeChecksForNodes_(kill);
    return { deleted: Object.keys(kill) };
  });
}

// 並び替え・階層変更をまとめて適用：updates = [{id,parentId,order}]
function applyMoves(updates) {
  return withLock_(function () {
    const sh = nodeSheet_();
    const v = sh.getDataRange().getValues();
    const rowById = {};
    for (let i = 1; i < v.length; i++) rowById[String(v[i][0])] = i + 1;
    (updates || []).forEach(function (u) {
      const r = rowById[String(u.id)];
      if (!r) return;
      if (u.parentId !== undefined) sh.getRange(r, 3).setValue(String(u.parentId || ''));
      if (u.order !== undefined) sh.getRange(r, 5).setValue(Number(u.order) || 0);
    });
    return { ok: true };
  });
}

// 指定ノード群のマークを一括設定（タブ一括選択用）。kind: 'daily'|'today'|'none'
// 列まとめ読み→まとめ書きで、件数が多くても速い。
function bulkSetMark(ids, kind) {
  const set = {};
  (ids || []).forEach(function (id) { set[String(id)] = true; });
  const daily = (kind === 'daily'), today = (kind === 'today');
  return withLock_(function () {
    const sh = nodeSheet_();
    const last = sh.getLastRow();
    if (last < 2) return { ok: true, count: 0 };
    const idCol = sh.getRange(2, 1, last - 1, 1).getValues();   // ID
    const dCol = sh.getRange(2, 8, last - 1, 1).getValues();    // 毎日
    const tCol = sh.getRange(2, 10, last - 1, 1).getValues();   // その日
    let n = 0;
    for (let i = 0; i < idCol.length; i++) {
      if (set[String(idCol[i][0])]) { dCol[i][0] = daily; tCol[i][0] = today; n++; }
    }
    sh.getRange(2, 8, last - 1, 1).setValues(dCol);
    sh.getRange(2, 10, last - 1, 1).setValues(tCol);
    return { ok: true, count: n };
  });
}

// 「今日のチェック」その日付の達成チェックを付け外し
function setDailyCheck(id, dateStr, done) {
  const date = dateStr || todayStr_();
  return withLock_(function () {
    const sh = checkSheet_();
    const v = sh.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]) === date && String(v[i][1]) === String(id)) { foundRow = i + 1; break; }
    }
    if (done && foundRow < 0) sh.appendRow([date, id]);
    else if (!done && foundRow > 0) sh.deleteRow(foundRow);
    return { id: id, date: date, done: !!done };
  });
}

// ===== 保存（エクスポート）==================================================
// scope: 'all'（全タブ）/ 'tab'（指定タブ）。tabId は scope==='tab' のとき必須。
// 戻り値 {url, name}。フロントでリンク表示する。

// Googleドキュメントに保存（おすすめ：チェックリストとして読みやすい・印刷も綺麗）
function exportToDoc(scope, tabId) {
  const data = collectExport_(scope, tabId);
  const doc = DocumentApp.create('memo チェックリスト ' + data.stamp);
  const body = doc.getBody();
  body.appendParagraph('memo  チェックリスト').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(data.stamp).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  data.tabs.forEach(function (t) {
    body.appendParagraph(t.name).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    t.items.forEach(function (it) {
      const glyph = it.done ? '☑ ' : '☐ ';
      const mark = it.daily ? '  〔定期〕' : it.today ? '  〔その日〕' : '';
      const p = body.appendParagraph(glyph + it.text + mark);
      p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      p.setIndentStart(it.depth * 18).setIndentFirstLine(it.depth * 18).setSpacingAfter(2);
    });
  });
  doc.saveAndClose();
  return { url: doc.getUrl(), name: doc.getName() };
}

// スプレッドシートに保存（表で扱いたいとき向け）
function exportToSheet(scope, tabId) {
  const data = collectExport_(scope, tabId);
  const ss = SpreadsheetApp.create('memo チェックリスト ' + data.stamp);
  const sh = ss.getSheets()[0].setName('チェックリスト');
  const rows = [['タブ', '項目', 'マーク', '完了']];
  data.tabs.forEach(function (t) {
    t.items.forEach(function (it) {
      const indent = new Array(it.depth + 1).join('　'); // 全角スペースで字下げ
      rows.push([t.name, indent + it.text, it.daily ? '定期' : it.today ? 'その日' : '', it.done ? '✓' : '']);
    });
  });
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#efece6');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 4);
  return { url: ss.getUrl(), name: ss.getName() };
}

// 【一度だけ実行】保存機能（ドキュメント/スプレッドシート作成）の権限を承認する。
// GASエディタで関数 authorizeExports を選んで ▶実行 → 権限ダイアログで許可。
// 確認用のテストファイルは作成後すぐゴミ箱へ入れる。
function authorizeExports() {
  const d = DocumentApp.create('memo 権限確認（消してOK）'); const did = d.getId(); d.saveAndClose();
  const s = SpreadsheetApp.create('memo 権限確認（消してOK）'); const sid = s.getId();
  try { DriveApp.getFileById(did).setTrashed(true); DriveApp.getFileById(sid).setTrashed(true); } catch (e) {}
  return '保存機能の権限を承認しました（テストファイルはゴミ箱へ）。アプリの「保存」が使えます。';
}

// scope に応じてタブと、その入れ子項目を深さ付き・並び順で集める
function collectExport_(scope, tabId) {
  const tabs = readTabs_();
  const nodes = readNodes_();
  const chosen = (scope === 'tab') ? tabs.filter(function (t) { return t.id === String(tabId); }) : tabs;
  const byParent = {};
  nodes.forEach(function (n) { const k = n.tab + '|' + n.parentId; (byParent[k] = byParent[k] || []).push(n); });
  Object.keys(byParent).forEach(function (k) { byParent[k].sort(function (a, b) { return a.order - b.order; }); });
  const out = chosen.map(function (t) {
    const items = [];
    (function walk(parentId, depth) {
      (byParent[t.id + '|' + parentId] || []).forEach(function (n) {
        items.push({ text: n.text, depth: depth, done: n.done, daily: n.daily, today: n.today });
        walk(n.id, depth + 1);
      });
    })('', 0);
    return { name: t.name, items: items };
  });
  return { tabs: out, stamp: todayStr_() };
}

// ===== 読み込み =============================================================
function readTabs_() {
  const sh = tabSheet_();
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const id = String(v[i][0] || '').trim();
    if (!id) continue;
    out.push({ id: id, name: String(v[i][1] || ''), type: String(v[i][2] || 'normal'), order: Number(v[i][3]) || i });
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

function readNodes_() {
  const sh = nodeSheet_();
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const id = String(v[i][0] || '').trim();
    if (!id) continue;
    out.push({
      id: id,
      tab: String(v[i][1] || ''),
      parentId: String(v[i][2] || ''),
      text: String(v[i][3] || ''),
      order: Number(v[i][4]) || 0,
      done: truthy_(v[i][5]),
      collapsed: truthy_(v[i][6]),
      daily: truthy_(v[i][7]),
      note: String(v[i][8] || ''),
      today: truthy_(v[i][9]),
    });
  }
  return out;
}

function readChecksForDate_(date) {
  const sh = checkSheet_();
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === date) out.push(String(v[i][1]));
  return out;
}

// ===== シート確保・初期シード ===============================================
function getSS_() {
  if (CONFIG.SS_ID) return SpreadsheetApp.openById(idFromUrl_(CONFIG.SS_ID));
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* 消えていたら作り直す */ } }
  const ss = SpreadsheetApp.create('ToDoメモ');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function ensureSheet_(name, headers) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#efece6');
    sh.setFrozenRows(1);
    const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
    if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch (e) {} }
  }
  return sh;
}
function tabSheet_() { return ensureSheet_(CONFIG.TAB_SHEET, ['ID', '名前', '種類', '並び順']); }
function nodeSheet_() { return ensureSheet_(CONFIG.NODE_SHEET, ['ID', 'タブID', '親ID', '内容', '並び順', '完了', '折りたたみ', '毎日', 'メモ', 'その日']); }
function checkSheet_() { return ensureSheet_(CONFIG.CHECK_SHEET, ['日付', 'ノードID']); }

// 初回だけ、既定タブ＋中身（buildSeed_）を流し込む
function ensureSeed_() {
  const sh = tabSheet_();
  if (sh.getLastRow() >= 2) return; // すでに何かある
  withLock_(function () {
    if (sh.getLastRow() >= 2) return;
    applySeed_();
  });
}

// 中身を一気に書き込む（タブ＋入れ子ノード）
function applySeed_() {
  const tsh = tabSheet_(), nsh = nodeSheet_();
  const seed = buildSeed_();
  const rows = [];
  seed.forEach(function (tab, ti) {
    const tabId = newId_();
    tsh.appendRow([tabId, tab.name, tab.type, ti]);
    collectSeedRows_(rows, tabId, '', tab.nodes);
  });
  if (rows.length) nsh.getRange(nsh.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
}

// シード表記をノード行へ。表記：'文字'＝葉、['親','子'...]＝入れ子、{t,done,daily,c}＝詳細
function collectSeedRows_(rows, tabId, parentId, list) {
  (list || []).forEach(function (node, i) {
    let text, children = [], done = false, daily = false, today = false;
    if (typeof node === 'string') { text = node; }
    else if (Object.prototype.toString.call(node) === '[object Array]') { text = node[0]; children = node.slice(1); }
    else { text = node.t; children = node.c || []; done = !!node.done; daily = !!node.daily; today = !!node.today; }
    const id = newId_();
    rows.push([id, tabId, parentId, text, i, done, false, daily, '', today]);
    if (children.length) collectSeedRows_(rows, tabId, id, children);
  });
}

// 【メンテ用】全部消して、最初から中身を作り直す（GASエディタで実行）。
function reseedMyData() {
  return withLock_(function () {
    clearBody_(tabSheet_());
    clearBody_(nodeSheet_());
    clearBody_(checkSheet_());
    applySeed_();
    return debugInfo();
  });
}
function clearBody_(sh) { if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); }

// ===== 定期購入タブの中身（buildSeed_ と setupSubscriptionTab で共用）=========
function subscriptionNodes_() {
  return [
    ['◆スーパー',
      ['食糧', 'ニラ', { t: '白菜', today: true }, { t: '玉ねぎ', today: true }, 'もやし×3', { t: 'きのこ', today: true },
        { t: '納豆×3', today: true }, { t: 'バナナ', today: true }, 'キウイ', '魚', '肉', { t: '水500ml×3', today: true },
        '卵 半分', 'うどん', 'ラーメン', { t: 'カットトマト缶', today: true }, { t: 'ヨーグルト', today: true }, 'パン', 'ハム',
        { t: '牛乳', today: true }, '米', '豆腐', 'そうめん', 'コーンフレーク', 'R1', '抹茶', 'BB', 'パスタ',
        { t: '冷凍シーフード', today: true }, 'チーズ'],
      ['調味料', 'カレールー', 'ポン酢', '醤油', 'マヨネーズ', 'みりん', 'ケチャップ', '焼肉用たれ', { t: '唐辛子', today: true },
        { t: '塩', today: true }, '砂糖', '七味', 'ブラックペッパー', 'オリーブオイル', { t: 'コンソメ', today: true },
        { t: 'ほんだし', today: true }, { t: '鶏がらスープ', today: true }, '蜂蜜（ハンガリー）', 'ごま油', 'にんにく', '生姜', 'ローレル', '酒', '味噌']],
    ['◆ドラッグストア【ポイント使う】', '白潤プレミアム化粧水', 'KOSE make keep mist', 'フライパン', { t: '米唐辛子', today: true },
      ['化粧品・理容', { t: 'アイブロウペンシル', today: true }, 'アイブロウブラシ', 'マスカラリムーバー', 'ポイントメイク落とし', 'マスカラ',
        'アイライン', 'リップクリーム', 'シェイバー', '歯ブラシ', '舌ブラシ', { t: '歯磨き粉', today: true }, '歯間クロス',
        '生理用品 大', '生理用品 小', 'タンポン', 'inclear', 'コットン', '鼻セレブ', '汗ふきシート 小2個'],
      ['キッチン用品', 'ウェットティッシュ', 'キッチンペーパー', 'フキン', '食器洗剤 magica除菌＋', '排水口ぬめりとりスプレー',
        '排水口ぬめりとりジェル', '排水溝ネット', 'げきおちくん', 'スポンジ：オレンジ', 'スポンジ：青', 'サランラップ', '魚用アルミホイル'],
      ['洗濯用品', '洗剤', '柔軟剤', 'エマール', '漂白剤'],
      ['トイレ用品', 'トイレットペーパー', 'ハンコ', 'トイレクイックル', 'トイレマジックリン', 'トイレ スクラビングバブル'],
      ['掃除用品', 'コバエスプレー', ['ウェットティッシュ', '部屋用×2', 'キッチン用', 'レンジ用'], '手袋', 'クイックルワイパー',
        'マジックリン（バス青・ガラス）', '洗面台用スポンジ', 'お風呂スポンジ', { t: 'ゴミ袋45L', today: true }, 'ゴミ袋 小20L',
        { t: '冷蔵庫消臭剤', today: true }, { t: '冷凍庫消臭剤', today: true },
        '△ファブリーズ（靴・キッチン・トイレ・部屋・クローゼット）', 'クローゼット用防虫剤・タンス用防虫剤', '消臭スプレー×2',
        { t: 'バド靴用消臭剤', today: true }]],
    ['◆お風呂', ['資生堂 サブリミック ルミノフォース', 'トリートメント 450g 4620円', 'シャンプー 450g 3520円', 'トリートメント 250g'],
      'クレンジングオイル（ビックカメラ）'],
    ['◆100均一', 'スプレー 小'],
    ['◆インターネット', { t: 'wellaヘアオイルスプレー', today: true }, { t: 'ハンドソープ', today: true }, 'コピー機インク'],
    ['◆皮膚科', 'ステロイド（肩甲骨）', '軟膏クリーム 50×2・25×2', 'ローション 50×1・25×2', 'ワセリン 100×1', '※ローション50mgを1個増やしてもらう'],
    ['◆ファンケル', '化粧水', '乳液', '美容液', '日焼け止め（顔）×1', '日焼け止め（体）', '下地', 'ファンデーション', 'パウダー', 'ボディクリーム'],
    ['◆Suqqu', 'チーク'],
    ['◆Aesop', 'ハンドクリーム'],
    ['◆無印', 'ステンレスラック 30×10×5.5 より小さいもの']
  ];
}

// アプリ読み込み時に「定期購入」大タブを “一度だけ” 自動で用意する。
// すでに同名タブがある／一度用意済み（フラグあり）なら何もしない＝重複も上書きもしない。
// （あとで自分で消したときに勝手に復活しないよう、作成済みフラグで一回限りにしている）
function ensureSubscriptionTab_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SUBS_TAB_DONE')) return;
  if (readTabs_().some(function (t) { return String(t.name).trim() === '定期購入'; })) {
    props.setProperty('SUBS_TAB_DONE', '1'); return; // 既にあるので作らず、以後チェックも省略
  }
  withLock_(function () {
    if (readTabs_().some(function (t) { return String(t.name).trim() === '定期購入'; })) return;
    const tsh = tabSheet_(), nsh = nodeSheet_();
    const tabId = newId_();
    tsh.appendRow([tabId, '定期購入', 'normal', nextTabOrder_(tsh)]);
    const rows = [];
    collectSeedRows_(rows, tabId, '', subscriptionNodes_());
    if (rows.length) nsh.getRange(nsh.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
  });
  props.setProperty('SUBS_TAB_DONE', '1');
}

// ===== 優先順位シート（AI連携の土台）=======================================
// memo のデータが入っているのと同じスプレッドシートに「優先順位」シートを足す。
// 大タブ／中タブ（各タブ直下の項目）を最初から埋めておき、優先度・期限・背景などは
// あとから手入力してもらう想定。AI（Claude Code等）はこの表を読んで段取りを決める。
const PRIORITY_HEADERS = ['大タブ', '中タブ', '小タブ', '優先度', '期限', '目的・背景', '進め方の希望', 'AIに任せたいこと'];

function ensurePrioritySheet_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('PRIORITY_SHEET_DONE_V3')) return;
  if (getSS_().getSheetByName('優先順位')) { props.setProperty('PRIORITY_SHEET_DONE_V3', '1'); return; }
  withLock_(function () { if (!getSS_().getSheetByName('優先順位')) createPrioritySheet_(); });
  props.setProperty('PRIORITY_SHEET_DONE_V3', '1');
}

function createPrioritySheet_() {
  const ss = getSS_();
  const sh = ss.insertSheet('優先順位');
  sh.getRange(1, 1, 1, PRIORITY_HEADERS.length).setValues([PRIORITY_HEADERS]).setFontWeight('bold').setBackground('#efece6');
  sh.setFrozenRows(1);
  const tabs = readTabs_(), nodes = readNodes_();
  const hasKids = function (id) { return nodes.some(function (n) { return n.parentId === id; }); };
  const childrenOf = function (tabId, pid) {
    return nodes.filter(function (n) { return n.tab === tabId && n.parentId === pid; })
      .sort(function (a, b) { return a.order - b.order; });
  };
  const rows = [];
  tabs.forEach(function (t) {
    const mids = childrenOf(t.id, '');  // 中タブ＝各タブ直下の項目
    if (!mids.length) { rows.push([t.name, '', '', '', '', '', '', '']); return; }
    mids.forEach(function (m) {
      rows.push([t.name, m.text, '', '', '', '', '', '']);                 // 中タブ行（小タブ空欄）
      childrenOf(t.id, m.id).forEach(function (c) {                        // その中のサブカテゴリ＝小タブ
        if (hasKids(c.id)) rows.push([t.name, m.text, c.text, '', '', '', '', '']);
      });
    });
  });
  if (rows.length) sh.getRange(2, 1, rows.length, PRIORITY_HEADERS.length).setValues(rows);
  sh.autoResizeColumns(1, PRIORITY_HEADERS.length);
  return rows.length;
}

// 【既存シートに後付け】「優先順位」シートの中タブの右に「小タブ」列を追加する。
// 既に入力済みのデータは消さずに、列だけ差し込む。すでに小タブ列があれば何もしない。
function addKotabColumn() {
  return withLock_(function () {
    const sh = getSS_().getSheetByName('優先順位');
    if (!sh) return '「優先順位」シートがありません。先に setupPrioritySheet を実行してください。';
    const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    if (head.indexOf('小タブ') >= 0) return '「小タブ」列はすでにあります。';
    sh.insertColumnAfter(2);  // 中タブ（2列目）の右に1列挿入
    sh.getRange(1, 3).setValue('小タブ').setFontWeight('bold').setBackground('#efece6');
    return '「小タブ」列を中タブの右に追加しました。';
  });
}

// 【手動でも実行可】優先順位シートを用意する。既にあれば中身は触らない。
function setupPrioritySheet() {
  return withLock_(function () {
    if (getSS_().getSheetByName('優先順位')) return '「優先順位」シートは既にあります（中身は変更していません）。';
    const n = createPrioritySheet_();
    return '「優先順位」シートを作成しました（' + n + '行）。スプレッドシートの下のタブで開けます。';
  });
}

// 【メンテ用・これを実行】「定期購入」を独立した“大タブ”として確実に用意する。
// 既存の「定期購入」タブや「定期購入（消耗品）」のかたまりがあれば一旦消して、
// 最新の中身で作り直す（他のタブ＝生活・ファイナンス等には一切触りません）。
// 何度実行してもOK。GASエディタで関数 setupSubscriptionTab を選んで ▶実行。
function setupSubscriptionTab() {
  return withLock_(function () {
    const tsh = tabSheet_(), nsh = nodeSheet_();

    // (1) 既存の「定期購入」大タブを削除（あれば）。消したタブのIDを覚えておく。
    const removedTabIds = {};
    const tv = tsh.getDataRange().getValues();
    for (let i = tv.length - 1; i >= 1; i--) {
      if (String(tv[i][1] || '').trim() === '定期購入') { removedTabIds[String(tv[i][0])] = true; tsh.deleteRow(i + 1); }
    }

    // (2) ノード側の掃除：上で消したタブのノード＋「定期購入（消耗品）/定期購入」のかたまりを削除。
    const nv = nsh.getDataRange().getValues();
    const children = {};
    for (let i = 1; i < nv.length; i++) { const p = String(nv[i][2]); (children[p] = children[p] || []).push(String(nv[i][0])); }
    const kill = {};
    for (let i = 1; i < nv.length; i++) {
      if (removedTabIds[String(nv[i][1])]) kill[String(nv[i][0])] = true;            // 消したタブの中身
      const t = String(nv[i][3] || '').trim();
      if (t === '定期購入（消耗品）' || t === '定期購入') {                            // 入れ物ごと子孫を回収
        (function collect(x) { kill[x] = true; (children[x] || []).forEach(collect); })(String(nv[i][0]));
      }
    }
    for (let i = nv.length - 1; i >= 1; i--) if (kill[String(nv[i][0])]) nsh.deleteRow(i + 1);
    removeChecksForNodes_(kill);

    // (3) 新しい「定期購入」大タブを作って、最新の中身を流し込む。
    const tabId = newId_();
    tsh.appendRow([tabId, '定期購入', 'normal', nextTabOrder_(tsh)]);
    const rows = [];
    collectSeedRows_(rows, tabId, '', subscriptionNodes_());
    if (rows.length) nsh.getRange(nsh.getLastRow() + 1, 1, rows.length, 10).setValues(rows);

    return '「定期購入」タブを用意しました（' + rows.length + ' 行）。アプリを再読み込みしてください。';
  });
}

// ===== 初期データ（あなたの実際の内容）=======================================
// 表記ルール：
//   '文字'                … 葉（チェック項目）
//   ['親', '子', ['孫', ...]] … 入れ子（先頭が親、以降が子）
//   {t:'文字', done:true}  … 完了済みで入れる（毎日に入れたいときは daily:true）
function buildSeed_() {
  return [
    { name: '生活', type: 'normal', nodes: [
      ['連絡', '聴診器', '斡旋', '健康診断', 'あしださん', 'スポーツ振興課・スポーツ協会', '（酒党派・28バド・荒川・村上）',
        ['11日', '警察', 'スポーツ振興課・スポーツ協会', '（酒党派・28バド・荒川・村上）']],
      ['家', 'クローゼット防虫交換', '化粧品捨てる', 'パウダーパフ・ブラシ洗う',
        ['定期', '掃除・洗濯', 'お布団洗う・交換', 'パウダーパフ・ブラシ洗う']],
      ['デスクワーク',
        ['検索', 'バドミントン靴下', 'ユニオン（団体交渉）目星つけとく'],
        ['手続き', 'ヒートウェーブ手続き', '健診予約', 'ルミネクレカ暗証番号', '看護師保険', '火災保険', '弁護士保険',
          'バドミントンシューズ交換', 'https://www.yonex.co.jp/news/2025/3095.html', '６月計画（運転免許証-7/16）',
          '仕事（クラウドワークス・ランサーズ・ココナラ）', 'メルカリ（ラケット・携帯・パソコン・本・ローランド）',
          'みほちゃん写真', ['到着次第', 'NL弁護士保険'], '郵送待ち NL弁護士保険'],
        ['デスクワーク'],
        ['購入', 'マスク', '水筒', '新社会人用（伊達メガネ・カバン・靴）', 'Amazon（単四充電器・変換プラグ）', 'Fancl（フェイスパウダー）',
          ['ビックカメラ（充電器・ワイヤレスイヤホン）',
            ['充電器', 'SMARTCOBY SLIMⅡ Wireless2.0 SS5K', 'SMARTCOBY SLIMⅡ Wireless2.2 Pro SS10K', 'SMARTCOBY SLIMII Wireless2.2 8K']],
          '資生堂：下地・ファンデなくなったら資生堂ポイント使う',
          ['お金できたら', 'キーケース', 'ヘアマスカラ', 'ヘアスプレー（トリエ１０）', 'シルク枕（CocoSILK）']],
        ['検討', 'KINJYO 目白インターナショナル英会話', '水筒', '６月の計画（バドミントン）', '弁護士（WOOD相談）', '聴診器', '美容院',
          '婦人科（ピル・検診・卵子凍結）・歯科・皮膚科・ローズ受診', 'バドミントン（MIX練習日）',
          'chocozap（筋トレ・ホワイトニング・脱毛）', '掃除', '洗濯', '聴診器連絡', 'WOOD・ドラム支払い方法（NLクレカ）'],
        ['定期', ['毎日少しずつ消す', '携帯容量整理', 'PC容量整理', 'Gmail容量メール消す']]],
      ['場所（お出かけリスト）',
        ['練馬',
          ['キレイモ', '免許証写真準備'],
          ['100均一', '名誉毀損：ファイル・インデックス', '茶葉入れ', '靴消臭', 'バド靴入れ', '唐辛子・砂糖・塩・コンソメ入れ',
            'クリアファイル（書類用・コスメ用）', 'ティッシュ入れ', 'クレジットカードケース', '食品クリップ'],
          ['ウェルシア', ['20日の下見（駅前・豊玉）', 'トリエ10', '靴消臭', '使い捨て手袋', '日焼け止め（アリィ）', 'Fanclボディークリーム', '漂白剤']]],
        ['池袋',
          ['３コインズ', 'メガネ入れ'],
          ['アンファミエ', 'ナースカーディガン', 'ナースシューズ', 'ナースズボン']],
        ['らもーれ', 'スタバ取りに行く'],
        ['忠医会病院', '聴診器取りに行く', 'カード返却']],
      ['ゴミ捨て', '日曜：可燃', '火曜：プラ・古紙', '水曜：可燃', '木曜：不燃', '金曜：ビン・カン'],
      ['映画', 'ユナイテッド・シネマ', 'https://epotoku.eposcard.co.jp/detail/index.html?ssfid=18025', '会員証更新', 'エポスカード→ローソン・ミニストップ']
    ]},

    { name: '定期購入', type: 'normal', nodes: subscriptionNodes_() },

    { name: 'ファイナンス', type: 'normal', nodes: [
      ['連絡', '看護師派遣登録', '福祉事務所（10万円止まっているか）',
        '市役所（保険料支払い終わっているか・お米券確認・補助金確認）', '健康保険', '弁護士相談'],
      ['家'],
      ['デスクワーク',
        ['検索', '看護師・カメラマン・音楽 単発アルバイト', '市役所：補助金確認'],
        ['手続き', '楽天保険 申請できなかったもの確認', '楽天証券 引き出し',
          ['５月20日', 'Jacst・V ⇨ ウェルシアポイント'],
          ['５月25日', 'ファンケル支払い'],
          ['６月以降', '部屋更新費用・手続き', '楽天保険（カメラ請求）', '10万円返済', '賃貸連絡（IH・キッチンの戸棚・お風呂ホース・エアコン）'],
          ['りそな銀行ポイント確認', 'OKりそな開設 1000円', 'OKりそなデビットポイント 1000円', 'りそなクレジットポイント ５万円以上で5000円',
            'https://www.resonabank.co.jp/kojin/hiraku/cam/2026_spring/', 'MyJCBアプリ ３ヶ月後にgifteeBOXあり', 'りそなクレカ 2027年3月解約']],
        ['デスクワーク', ['書類作成', '美容院裁判', '美容院裁判 ⇨ 労働斡旋', 'スポーツ', 'WOOD', '名誉毀損']],
        ['購入'],
        ['検討', '部屋更新代金・ボーナス払い代金 計画'],
        ['定期',
          ['毎月1日', 'auじぶん銀行に預金を集める（じぶん0.46%／SBI0.41%）',
            'au Payチャージ（じぶんプラス：20日判定・1円以上） https://www.jibunbank.co.jp/customer_stage/',
            'paypayへ10万円チャージ（paypay→paypay銀行出金→auじぶん銀行へ入金）'],
          'SBI証券：東海カーボン売却（3000円超えでS株 成行売り）',
          'R2：1Gで通知設定 https://dash.cloudflare.com/93ca1682c10b8c1c72169e0fd96e81de/billing']],
      ['場所（お出かけリスト）',
        ['練馬', '生活 3000円', '活動：10000円分（100円×30＝3000／500円×6＝3000／1000円×4＝4000）',
          '三菱UFJ：現金全額入金 ⇨ きっちりにして引き出し',
          ['⇨ コンビニ', 'paypay 60000円引き出し', '活動用・生活用 現金のこす', 'au入金']]]
    ]},

    { name: '音楽', type: 'normal', nodes: [
      ['連絡'],
      ['家'],
      ['デスクワーク',
        ['検索', 'Avid 更新確認', 'YouTube 勝又 消す'],
        ['手続き'],
        ['デスクワーク', 'CD印刷', 'Over', 'IBUKI片目', '看板作り', 'グッツ一覧表', 'オケ作り直し',
          'コード譜・歌詞カード ⇨ 暗譜', 'チケット作成', '名刺作成・変更', 'アルバム作成'],
        ['購入', '変換プラグ', 'CD・CDケース', '島村ポイント使用（弦・ピック）', '音楽用ノート',
          'マイクスタンドにつく携帯スタンド', 'チェキフィルム', '小物BOX', 'ライブ配信機材', '路上ライブ機材', 'ライト', 'usbC-イヤホン変換プラグ'],
        ['検討', '音楽活動（CDたて・CD入れ・チケットたて・CDフィルム・写真フィルム）',
          '予約計画（活動予定・路上ライブ・WOOD予約・ドラム予約）', 'SNS計画（SNS用写真用意）', 'ayaka-music.com どうするか'],
        ['定期']],
      ['場所（お出かけリスト）',
        ['練馬',
          ['ノジマ', 'チェキフィルム'],
          ['100円均一', 'お金ポーチ']]]
    ]},

    { name: 'カメラ', type: 'normal', nodes: [
      ['連絡'],
      ['家'],
      ['デスクワーク',
        ['検索', 'deltaphoto', 'https://www.geeq.co.jp/'],
        ['手続き'], ['デスクワーク'], ['購入'], ['検討'], ['定期']],
      ['場所（お出かけリスト）']
    ]},

    { name: 'お仕事', type: 'work', nodes: [
      ['人',
        ['渡辺さん', '木下さんに連絡したい',
          ['webアプリ', 'キャンセルどうするか', 'いつ予約は決定するか・予約ページを消すタイミング'],
          '記録改善方向性（応急処置：通所＝業務日誌／訪問＝フェイスシート）', 'お金お預かり表', 'マニュアル化', '記録改善', 'ご利用者ページ', '電子化サイン'],
        ['茂木さん', 'MCSアカウント', '土曜出勤'],
        ['堀越さん', { t: 'スケッターアプリでURL飛べない', done: true }],
        ['大原さん', { t: '通信：エクセルにする', done: true }],
        ['佐久間さん', { t: 'シフト', done: true }, 'ふうとさん 土日の訪問時間', '通所追加時の対応（食事）', '担当者'],
        ['林さん・猪越さん', '１Day業務', 'フェイスシート作成']],
      ['デスクワーク', '録画みる／書類提出', 'フェイスシート', '健康診断結果探す', 'バインダー探す',
        ['事務', 'マウスパッド・パソコンセットアップ', 'メモアプリ入れる'],
        ['購入', 'ナースパンツ２着', 'ナースシューズ', '携帯保護画面', 'iPadペンシル', '携帯スタンド']]
    ]}
  ];
}

// ひな形（入れ子）をノードシートへ挿入
function insertScaffold_(tabId, items, parentId) {
  const sh = nodeSheet_();
  const rows = [];
  (function walk(list, parent) {
    (list || []).forEach(function (it, i) {
      const id = newId_();
      rows.push([id, tabId, parent, it.t, i, false, false, false, '', false]);
      if (it.c && it.c.length) walk(it.c, id);
    });
  })(items, parentId);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
}

// ===== ヘルパー =============================================================
function findRow_(sh, id) {
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 1;
  return -1;
}
function nextTabOrder_(sh) {
  const v = sh.getDataRange().getValues();
  let mx = -1;
  for (let i = 1; i < v.length; i++) mx = Math.max(mx, Number(v[i][3]) || 0);
  return mx + 1;
}
function removeChecksForNodes_(idMap) {
  const sh = checkSheet_();
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (idMap[String(v[i][1])]) sh.deleteRow(i + 1);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { /* 取れなくても続行 */ }
  try { return fn(); } finally { try { lock.releaseLock(); } catch (e) {} }
}
function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
}
function newId_() { return 'N' + new Date().getTime().toString(36) + Math.floor(Math.random() * 1e6).toString(36); }
function truthy_(v) {
  if (v === true) return true;
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === '✓' || s === '◯' || s === '○' || s === 'はい';
}
function idFromUrl_(s) {
  const str = String(s).trim();
  const m = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : str;
}

// 【動作確認用】GASエディタで実行：保存先ブックURLと件数をログに出す。
function debugInfo() {
  const ss = getSS_();
  ensureSeed_();
  const msg = [
    '■ 保存先ブック：' + ss.getName(),
    '  URL：' + ss.getUrl(),
    '  タブ：' + readTabs_().map(function (t) { return t.name + '(' + t.type + ')'; }).join(' / '),
    '  ノード数：' + readNodes_().length + '件',
    '  今日(' + todayStr_() + ')のチェック：' + readChecksForDate_(todayStr_()).length + '件',
  ].join('\n');
  Logger.log(msg);
  return msg;
}