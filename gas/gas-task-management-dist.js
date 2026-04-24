// ===================================================================
// 現場タスク管理 v2 - GAS バックエンド
// 使い方: script.google.com で新規プロジェクト → このコードを貼付 → setup() 実行 → デプロイ
// ===================================================================

const SHEET_PROP  = '物件';
const SHEET_PHASE = '工程';
const SHEET_TASK  = 'タスク';
const SHEET_TMPL  = 'テンプレート';
const SHEET_CONF  = '設定';
const KEY_DEFAULT = '1111';

const TEMPLATE_PHASES = [
  '現場FB','図面FB','地鎮祭','地縄立会い','生産移管',
  '引継ぎ会','施工計画説明','遣り方検査',
  '配筋検査','型枠検査','基礎検査','建て方',
  '構造検査','構造立会い','雨仕舞い検査',
  '木完検査','木完立会い','竣工検査','竣工立会い','引渡し'
];

// ── セットアップ（初回1回だけ実行） ─────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function mk(name, headers) {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    return sh;
  }

  mk(SHEET_PROP,  ['name','city','start','frame','completion','handover']);
  mk(SHEET_PHASE, ['id','prop','name','date','time','order','alert_prep','alert_post']);
  mk(SHEET_TASK,  ['id','prop','phase_id','col','name','done','order','priority']);
  mk(SHEET_TMPL,  ['id','phase_name','col','name','order','priority']);
  const c = mk(SHEET_CONF, ['key','value']);
  if (c.getLastRow() < 2) c.appendRow(['accessKey', KEY_DEFAULT]);

  SpreadsheetApp.getUi().alert(
    'セットアップ完了！\n' +
    'デフォルトアクセスキー: ' + KEY_DEFAULT + '\n\n' +
    '次: [デプロイ] → [新しいデプロイ] → ウェブアプリ → 全員アクセス可'
  );
}

// ── エントリーポイント ───────────────────────────────────────────
function doPost(e) {
  try {
    const d  = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (d.mode !== 'login' && String(d.key) !== String(conf(ss, 'accessKey')))
      return out({ ok: false, error: '認証エラー' });

    const fn = {
      login:          () => login(ss, d),
      changeKey:      () => changeKey(ss, d),
      getAll:         () => getAll(ss, d),
      addProperty:    () => addProperty(ss, d),
      updateProperty: () => updateProperty(ss, d),
      deleteProperty: () => deleteProperty(ss, d),
      addPhase:       () => addPhase(ss, d),
      updatePhase:    () => updatePhase(ss, d),
      deletePhase:    () => deletePhase(ss, d),
      reorderPhases:  () => reorderPhases(ss, d),
      addTask:        () => addTask(ss, d),
      updateTask:     () => updateTask(ss, d),
      moveTask:       () => moveTask(ss, d),
      deleteTask:     () => deleteTask(ss, d),
      getTemplates:    () => getTemplates(ss, d),
      saveTmplPhases:  () => saveTmplPhases(ss, d),
      seedTemplates:   () => seedTemplates(ss),
      addTemplate:     () => addTemplate(ss, d),
      updateTemplate:  () => updateTemplate(ss, d),
      deleteTemplate:  () => deleteTemplate(ss, d),
    };

    if (!fn[d.mode]) return out({ ok: false, error: 'unknown: ' + d.mode });
    return out(fn[d.mode]());
  } catch (err) {
    return out({ ok: false, error: err.message });
  }
}

function out(d) {
  return ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ユーティリティ ───────────────────────────────────────────────
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function rows(ss, sheet) {
  const sh = ss.getSheetByName(sheet);
  if (!sh || sh.getLastRow() < 2) return [];
  const [h, ...data] = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  return data.map(r => Object.fromEntries(h.map((k, i) => {
    let v = r[i];
    if (v instanceof Date) {
      // 1900年以前 = 時刻のみの値（Sheetsの時刻型）
      v = v.getFullYear() <= 1900
        ? Utilities.formatDate(v, tz, 'HH:mm')
        : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    }
    return [k, v];
  })));
}

function conf(ss, key) {
  const r = rows(ss, SHEET_CONF).find(r => r.key === key);
  return r ? r.value : null;
}

function setConf(ss, key, val) {
  const sh = ss.getSheetByName(SHEET_CONF);
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === key) { sh.getRange(i + 1, 2).setValue(val); return; }
  }
  sh.appendRow([key, val]);
}

function upd(ss, sheet, idCol, idVal, updates) {
  const sh = ss.getSheetByName(sheet);
  const d = sh.getDataRange().getValues();
  const h = d[0];
  const ci = h.indexOf(idCol);
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][ci]) === String(idVal)) {
      Object.keys(updates).forEach(k => {
        const j = h.indexOf(k);
        if (j >= 0) sh.getRange(i + 1, j + 1).setValue(updates[k]);
      });
      return true;
    }
  }
  return false;
}

function del(ss, sheet, col, val) {
  const sh = ss.getSheetByName(sheet);
  if (!sh || sh.getLastRow() < 2) return;
  const d = sh.getDataRange().getValues();
  const ci = d[0].indexOf(col);
  for (let i = d.length - 1; i >= 1; i--)
    if (String(d[i][ci]) === String(val)) sh.deleteRow(i + 1);
}

// ── 認証 ────────────────────────────────────────────────────────
function login(ss, d) {
  if (String(d.key) === String(conf(ss, 'accessKey')))
    return { ok: true, firstLogin: String(d.key) === KEY_DEFAULT };
  return { ok: false, error: 'アクセスキーが違います' };
}

function changeKey(ss, d) {
  setConf(ss, 'accessKey', d.newKey);
  return { ok: true };
}

// ── 全データ取得 ─────────────────────────────────────────────────
function getAll(ss, d) {
  const props  = rows(ss, SHEET_PROP);
  let phases   = rows(ss, SHEET_PHASE);
  let tasks    = rows(ss, SHEET_TASK);
  if (d.prop) {
    phases = phases.filter(r => r.prop === d.prop);
    tasks  = tasks.filter(r  => r.prop === d.prop);
  }
  phases.sort((a, b) => Number(a.order) - Number(b.order));
  tasks.sort((a, b)  => Number(a.order) - Number(b.order));
  return { ok: true, props, phases, tasks };
}

// ── 物件 ────────────────────────────────────────────────────────
function addProperty(ss, d) {
  // 物件を追加
  ss.getSheetByName(SHEET_PROP)
    .appendRow([d.name, d.city||'', d.start||'', d.frame||'', d.completion||'', d.handover||'']);

  // 工程リスト：設定シートのカスタム順 → なければデフォルト
  const phasesJson = conf(ss, 'tmplPhases');
  const phaseNames = (phasesJson ? JSON.parse(phasesJson) : null) || TEMPLATE_PHASES;

  // テンプレートタスクを一括読み込み
  const tmpls = rows(ss, SHEET_TMPL);

  // 工程行・タスク行をまとめて構築
  const phaseRows = [];
  const taskRows  = [];

  phaseNames.forEach((name, i) => {
    const phId = uid();
    phaseRows.push([phId, d.name, name, '', '', i + 1, 7, 3]);
    tmpls
      .filter(t => t.phase_name === name)
      .sort((a, b) => Number(a.order) - Number(b.order))
      .forEach(t => taskRows.push([uid(), d.name, phId, t.col, t.name, false, t.order, t.priority || false]));
  });

  // 工程を一括書き込み
  const ph = ss.getSheetByName(SHEET_PHASE);
  if (phaseRows.length) {
    const phStart = ph.getLastRow() + 1;
    ph.getRange(phStart, 1, phaseRows.length, 8).setValues(phaseRows);
  }

  // タスクを一括書き込み
  const tk = ss.getSheetByName(SHEET_TASK);
  if (taskRows.length) {
    const tkStart = tk.getLastRow() + 1;
    tk.getRange(tkStart, 1, taskRows.length, 8).setValues(taskRows);
  }

  return { ok: true };
}

function updateProperty(ss, d) {
  const u = {};
  ['city','start','frame','completion','handover'].forEach(f => {
    if (d[f] !== undefined) u[f] = d[f];
  });
  upd(ss, SHEET_PROP, 'name', d.name, u);
  return { ok: true };
}

function deleteProperty(ss, d) {
  del(ss, SHEET_TASK,  'prop', d.name);
  del(ss, SHEET_PHASE, 'prop', d.name);
  del(ss, SHEET_PROP,  'name', d.name);
  return { ok: true };
}

// ── 工程 ────────────────────────────────────────────────────────
function addPhase(ss, d) {
  const existing = rows(ss, SHEET_PHASE).filter(r => r.prop === d.prop);
  const maxOrd   = existing.length ? Math.max(...existing.map(p => Number(p.order) || 0)) : 0;
  const id = uid();
  ss.getSheetByName(SHEET_PHASE)
    .appendRow([id, d.prop, d.name, d.date||'', d.time||'', maxOrd + 1, d.alert_prep||7, d.alert_post||3]);
  return { ok: true, id };
}

function updatePhase(ss, d) {
  const u = {};
  ['name','date','time','alert_prep','alert_post'].forEach(f => {
    if (d[f] !== undefined) u[f] = d[f];
  });
  upd(ss, SHEET_PHASE, 'id', d.id, u);
  return { ok: true };
}

function deletePhase(ss, d) {
  del(ss, SHEET_TASK,  'phase_id', d.id);
  del(ss, SHEET_PHASE, 'id',       d.id);
  return { ok: true };
}

function reorderPhases(ss, d) {
  // d.order = [{id, order}, ...]
  d.order.forEach(item => upd(ss, SHEET_PHASE, 'id', item.id, { order: item.order }));
  return { ok: true };
}

// ── タスク ───────────────────────────────────────────────────────
function addTask(ss, d) {
  const existing = rows(ss, SHEET_TASK)
    .filter(r => r.phase_id === d.phase_id && r.col === d.col);
  const maxOrd = existing.length ? Math.max(...existing.map(t => Number(t.order) || 0)) : 0;
  const id = uid();
  ss.getSheetByName(SHEET_TASK)
    .appendRow([id, d.prop, d.phase_id, d.col, d.name, false, maxOrd + 1, false]);
  return { ok: true, id };
}

function updateTask(ss, d) {
  const u = {};
  ['name','done','priority'].forEach(f => { if (d[f] !== undefined) u[f] = d[f]; });
  upd(ss, SHEET_TASK, 'id', d.id, u);
  return { ok: true };
}

function moveTask(ss, d) {
  // 別工程または別列へ移動
  const u = {};
  ['phase_id','col','prop'].forEach(f => { if (d[f] !== undefined) u[f] = d[f]; });
  // 移動先の末尾に追加
  const dest = rows(ss, SHEET_TASK)
    .filter(r => r.phase_id === d.phase_id && r.col === d.col);
  u.order = dest.length ? Math.max(...dest.map(t => Number(t.order) || 0)) + 1 : 1;
  upd(ss, SHEET_TASK, 'id', d.id, u);
  return { ok: true };
}

function deleteTask(ss, d) {
  del(ss, SHEET_TASK, 'id', d.id);
  return { ok: true };
}

// ── テンプレート ─────────────────────────────────────────────────
function getTemplates(ss, d) {
  const templates = rows(ss, SHEET_TMPL);
  templates.sort((a, b) => Number(a.order) - Number(b.order));
  const phasesJson = conf(ss, 'tmplPhases');
  const phases = phasesJson ? JSON.parse(phasesJson) : null;
  return { ok: true, templates, phases };
}

function saveTmplPhases(ss, d) {
  setConf(ss, 'tmplPhases', JSON.stringify(d.phases));
  return { ok: true };
}

function addTemplate(ss, d) {
  const existing = rows(ss, SHEET_TMPL)
    .filter(r => r.phase_name === d.phase_name && r.col === d.col);
  const maxOrd = existing.length ? Math.max(...existing.map(t => Number(t.order) || 0)) : 0;
  const id = uid();
  ss.getSheetByName(SHEET_TMPL)
    .appendRow([id, d.phase_name, d.col, d.name, maxOrd + 1, false]);
  return { ok: true, id };
}

function updateTemplate(ss, d) {
  const u = {};
  ['name','priority','phase_name','col','order'].forEach(f => { if (d[f] !== undefined) u[f] = d[f]; });
  upd(ss, SHEET_TMPL, 'id', d.id, u);
  return { ok: true };
}

function deleteTemplate(ss, d) {
  del(ss, SHEET_TMPL, 'id', d.id);
  return { ok: true };
}

// ── 旧アプリのマスタータスクをテンプレートに一括登録（GASエディタから実行） ──
// 使い方: GASエディタで seedTemplates() を選択して実行
// ※既存のテンプレートタスクは上書きされます
// GASエディタから実行する場合は引数なし。アプリAPIから呼ぶ場合は ss を渡す
function seedTemplates(ss) {
  const _ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sh = _ss.getSheetByName(SHEET_TMPL);

  // シート全体をクリアしてヘッダーを再設定
  sh.clearContents();
  sh.getRange(1, 1, 1, 6).setValues([['id','phase_name','col','name','order','priority']]);

  // [phase_name, col, name]
  // col: 'prep'=準備 / 'day'=当日 / 'post'=事後
  const tasks = [
    // ── 現場FB ──
    ['現場FB','day', '配置基準の境界確認：丁張りが可能か'],
    ['現場FB','day', '引込位置の確認：電柱・弱電位置の確認'],
    ['現場FB','day', '電線保護カバー・敷鉄板・道路使用の有無の確認'],
    ['現場FB','post','仮設計画図・生産補正シートの作成'],
    ['現場FB','post','資料提出'],
    // ── 図面FB ──
    ['図面FB','day', '図面チェック※Documents使用'],
    ['図面FB','post','図面チェック完了をLINEで報告'],
    // ── 地鎮祭 ──
    ['地鎮祭','prep','テント設営の依頼'],
    ['地鎮祭','prep','奉献酒の用意'],
    ['地鎮祭','day', '鎮め物を業者へ渡す'],
    ['地鎮祭','prep','仮設計画図の作成・手配'],
    // ── 地縄立会い ──
    ['地縄立会い','prep','地縄張りの依頼'],
    ['地縄立会い','prep','近隣挨拶分の作成'],
    ['地縄立会い','prep','近隣挨拶タオルの用意'],
    ['地縄立会い','prep','地縄張り用配置図を設計担当へLINE確認'],
    ['地縄立会い','prep','長期優良住宅の申請日の確認※地盤改良工事前の申請が必須'],
    ['地縄立会い','prep','立会いノートの作成'],
    ['地縄立会い','prep','施工計画説明の日時決定'],
    ['地縄立会い','prep','ショールームの予約'],
    ['地縄立会い','prep','引継ぎ会日時の依頼'],
    ['地縄立会い','post','立会い完了の報告、施工計画説明日時のカレンダー共有'],
    ['地縄立会い','post','実績入力：地縄立会い'],
    ['地縄立会い','post','立会いノート・遣り方検査日時依頼の伝言板アップ'],
    // ── 生産移管 ──
    ['生産移管','prep','工程表の作成'],
    ['生産移管','prep','引継書の印刷'],
    ['生産移管','prep','電子発注明細の印刷'],
    ['生産移管','prep','仮設計画図兼作業指示書の作成'],
    ['生産移管','prep','NACCS業者登録・メンテ'],
    ['生産移管','prep','杭ナビデータ有：業者登録※栃井建設のみ'],
    ['生産移管','prep','電子先行発注：プレカット・軒天'],
    ['生産移管','prep','長期優良住宅認可予定日の確認'],
    ['生産移管','prep','生産補正シート作成・予算入力依頼'],
    ['生産移管','prep','図面マーキング'],
    ['生産移管','prep','引継ぎ前チェック：引継ぎ書・提案工事内容・エプコ配管経路図・スミテン図'],
    ['生産移管','prep','設計指示価格の見積もり期日確認'],
    ['生産移管','day', '引継ぎ会資料の提出'],
    ['生産移管','post','実績入力：引継ぎ'],
    // ── 施工計画説明 ──
    ['施工計画説明','prep','立会いノートの作成'],
    ['施工計画説明','prep','お客様への質疑まとめ'],
    ['施工計画説明','prep','着工合意書の内容確認'],
    ['施工計画説明','prep','お客様配布物：工程表・製本図面・工事案内ファイル'],
    ['施工計画説明','prep','構造立会い日の決定'],
    ['施工計画説明','prep','手形式の有無の確認'],
    ['施工計画説明','prep','棟札ご持参の案内'],
    ['施工計画説明','day', 'TV・ネットの早期申込みの説明'],
    ['施工計画説明','post','実績入力：施工計画説明'],
    ['施工計画説明','post','立会いノートを伝言板にアップ：契約電気容量・メーター名義名を共有'],
    ['施工計画説明','post','朱書き図面を伝言板にアップ'],
    // ── 遣り方検査 ──
    ['遣り方検査','prep','基礎工事計画書の提出'],
    ['遣り方検査','day', '遣り方シールの確認・記載'],
    ['遣り方検査','day', '引継ぎ会資料の持参：工務店担当にサインもらう'],
    ['遣り方検査','day', '安全日誌の記載'],
    ['遣り方検査','post','NACCSへ着工写真をアップ'],
    ['遣り方検査','post','引継ぎ会資料の提出'],
    ['遣り方検査','post','実績入力：本体着工'],
    ['遣り方検査','post','実績入力：着工前ミーティング'],
    ['遣り方検査','post','建性①配筋検査の申込'],
    // ── 配筋検査 ──
    ['配筋検査','prep','CON打設の近隣挨拶分の作成・持参'],
    ['配筋検査','prep','長期優良住宅が認可済かの確認※ベースCON打設までが必須'],
    ['配筋検査','day', '島基礎の計測'],
    ['配筋検査','day', 'スリーブ位置の計測・記録'],
    ['配筋検査','day', 'スペーサーブロック・シート重ね長さの計測・記録'],
    ['配筋検査','day', 'コーナー・隅角部補強金・主筋継手位置の記録'],
    ['配筋検査','day', '性能評価シールを確認看板に貼る'],
    ['配筋検査','day', '近隣挨拶：CON打設'],
    ['配筋検査','day', '安全日誌の記載'],
    ['配筋検査','day', 'iPadで性能評価書類の記載・報告※生産補助へ'],
    ['配筋検査','post','図面の提出'],
    ['配筋検査','post','実績入力：配筋検査'],
    // ── 型枠検査 ──
    ['型枠検査','day', '島基礎の計測・記録'],
    ['型枠検査','day', 'ボルト類の位置・レベルの計測・記録'],
    ['型枠検査','day', 'スラブ厚の計測・記録'],
    ['型枠検査','day', '被り厚検査棒による被り厚の確認'],
    ['型枠検査','day', '安全日誌の記載'],
    ['型枠検査','post','図面提出'],
    // ── 基礎検査 ──
    ['基礎検査','prep','建て方の近隣挨拶分の作成・持参'],
    ['基礎検査','prep','近隣挨拶：建て方'],
    ['基礎検査','day', '防蟻パイプの本数計測・記録'],
    ['基礎検査','day', '安全日誌の記載'],
    ['基礎検査','post','図面提出'],
    ['基礎検査','post','実績入力：基礎検査'],
    ['基礎検査','post','伝言板で構造検査・構造立会い日時の共有'],
    ['基礎検査','post','基礎精算の依頼'],
    ['基礎検査','post','建て方の予定の社内報告・カレンダー登録'],
    // ── 建て方 ──
    ['建て方','prep','建て方前の入金確認※支払い条件による'],
    ['建て方','prep','建て方人数の確認、施主報告'],
    ['建て方','prep','建て方計画書の確認・承認'],
    ['建て方','day', '全景写真撮影、次長へLINE報告'],
    ['建て方','day', '安全日誌の確認・記載'],
    ['建て方','day', '足場点検実施→是正があればLINEグループで指示'],
    ['建て方','day', '実績入力：建て方開始'],
    ['建て方','post','実績入力：野地板完了'],
    ['建て方','post','建て方完了の報告→施主・社内'],
    ['建て方','post','建性②：構造検査日時の確認※申込みは生産事務が行う'],
    // ── 構造検査 ──
    ['構造検査','day', '階高の計測・記録'],
    ['構造検査','day', '構造材種の確認'],
    ['構造検査','day', 'センサーライトを玄関先に取付'],
    ['構造検査','day', '防蟻剤容器に危険物シールが貼ってあるか確認'],
    ['構造検査','day', '各階消火器設置の確認'],
    ['構造検査','day', '安全日誌の記載'],
    ['構造検査','post','図面提出'],
    ['構造検査','post','基礎精算・発注'],
    // ── 構造立会い ──
    ['構造立会い','prep','お客様に立会い日程・棟札・手形の確認'],
    ['構造立会い','prep','立会いノートの作成'],
    ['構造立会い','prep','木完立会い日時の決定'],
    ['構造立会い','prep','支給品を木完立会い持参の案内'],
    ['構造立会い','post','社内：立会い完了の報告'],
    ['構造立会い','post','伝言板で木完検査・木完立会い日時の共有'],
    ['構造立会い','post','「木完検査」NACCS送信ボタンのクリック'],
    // ── 雨仕舞い検査 ──
    ['雨仕舞い検査','prep','吹付の近隣挨拶分の作成'],
    ['雨仕舞い検査','prep','実績入力：構造検査※生産事務代理'],
    ['雨仕舞い検査','day', '電気配線・BOXの確認'],
    ['雨仕舞い検査','day', '防火区画テープ貼りの確認・記録※平屋は該当無し'],
    ['雨仕舞い検査','day', '近隣挨拶：吹付作業'],
    ['雨仕舞い検査','day', '安全日誌の記入'],
    ['雨仕舞い検査','post','図面提出'],
    ['雨仕舞い検査','post','実績入力：構造雨仕舞い'],
    ['雨仕舞い検査','post','建性③：断熱検査日時の確認※申込みは生産事務が行う'],
    // ── 木完検査 ──
    ['木完検査','prep','実績入力：断熱検査※生産事務代理'],
    ['木完検査','prep','実績入力：左官防水検査'],
    ['木完検査','prep','実績入力：足場解体'],
    ['木完検査','prep','社内：追加変更の覚書の有無の確認'],
    ['木完検査','day', 'BCのPBビスピッチ確認・記録'],
    ['木完検査','day', '天井高さの測定'],
    ['木完検査','day', 'クロスサンプルの持参・貼付け'],
    ['木完検査','day', '安全日誌の記載'],
    ['木完検査','post','図面提出'],
    ['木完検査','post','実績入力：木完検査'],
    ['木完検査','post','建性④：完成検査の申込み'],
    ['木完検査','post','完了検査の申込み→申込み用紙・省令準耐火チェックシートを生産事務へ提出'],
    ['木完検査','post','伝言板で竣工検査日時の共有'],
    ['木完検査','post','仮設撤去の依頼'],
    // ── 木完立会い ──
    ['木完立会い','prep','お客様に立会い日時の確認'],
    ['木完立会い','prep','立会いノートの作成'],
    ['木完立会い','day', '施主支給品の受取り'],
    ['木完立会い','day', 'ライフライン名義変更の案内'],
    ['木完立会い','day', '竣工立会い・引渡しスケジュールの確定'],
    ['木完立会い','day', '最終金額確認書についての案内'],
    ['木完立会い','post','社内：立会い完了の報告・引渡し日のカレンダー登録'],
    ['木完立会い','post','引渡し申請をLINEで依頼'],
    ['木完立会い','post','竣工立会い・引渡しの工程メンテ'],
    ['木完立会い','post','「引渡し」NACCS送信ボタンのクリック'],
    ['木完立会い','post','伝言板で引渡しまでのスケジュール共有'],
    // ── 竣工検査 ──
    ['竣工検査','prep','取説ファイリング'],
    ['竣工検査','day', '玄関土間の段差の測定・記録'],
    ['竣工検査','day', 'センサー照明の設定番号確認'],
    ['竣工検査','day', '防犯カメラの登録確認・センサー消音'],
    ['竣工検査','day', '給気フィルターが入っているか、内部清掃'],
    ['竣工検査','day', 'UB立上りスリーブの断熱施工・点検口テープ貼りの確認'],
    ['竣工検査','day', '熱源機の2重ナット固定・アース接続の確認'],
    ['竣工検査','day', 'エアコンスリーブ断熱材が有るか、内部清掃状況の確認'],
    ['竣工検査','day', '給水・給湯ヘッダーの2箇所以上の固定確認'],
    ['竣工検査','day', '小屋裏の断熱施工確認'],
    ['竣工検査','day', '小屋裏の太陽光配線確認、LINEグループへ報告'],
    ['竣工検査','day', '安全日誌の記載'],
    ['竣工検査','post','手直し資料のまとめ、伝言板アップ'],
    ['竣工検査','post','図面の提出'],
    ['竣工検査','post','実績入力：竣工検査・4回目検査'],
    ['竣工検査','post','社内：完了検査・性能評価検査の完了報告'],
    // ── 竣工立会い ──
    ['竣工立会い','prep','お客様に立会い日時の確認'],
    ['竣工立会い','prep','立会いノートの作成'],
    ['竣工立会い','prep','引渡しの実印持参の案内'],
    ['竣工立会い','day', '災害用スリーブの説明'],
    ['竣工立会い','day', '汚水桝（地域による）・雨水桝の清掃の説明'],
    ['竣工立会い','day', '三協立山アルミの電池錠：取説を送る、ユーザー登録の依頼'],
    ['竣工立会い','post','手直し資料のまとめ、伝言板アップ'],
    ['竣工立会い','post','実績入力：竣工立会い'],
    ['竣工立会い','post','立会い完了の社内報告'],
    // ── 引渡し ──
    ['引渡し','prep','お客様に引渡し日時の確認'],
    ['引渡し','prep','立会いノートの作成'],
    ['引渡し','prep','引渡し書類・記念時計品・鍵の持参'],
    ['引渡し','day', '電気錠・電池錠の設定・登録'],
    ['引渡し','day', '入居後訪問（LINE聞き取り）の案内'],
    ['引渡し','post','引渡受書・外観写真をLINEで送る'],
    ['引渡し','post','完成写真をLINEグループに送る'],
    ['引渡し','post','引渡し書類の提出（受書・AM引継書・長期優良認定書・図面）'],
    ['引渡し','post','実績入力：入居予定日'],
  ];

  // col ごとの order カウンター → 一括でシートに書き込む
  const orderMap = {};
  const rows = tasks.map(([phase, col, name]) => {
    const key = phase + '|' + col;
    orderMap[key] = (orderMap[key] || 0) + 1;
    return [uid(), phase, col, name, orderMap[key], false];
  });

  sh.getRange(2, 1, rows.length, 6).setValues(rows);

  // エディタから直接実行した場合のみ alert を表示
  if (!ss) {
    SpreadsheetApp.getUi().alert(
      rows.length + '件のテンプレートタスクを登録しました\n※引継ぎ会は空欄です'
    );
  }
  return { ok: true, count: rows.length };
}

// ── テンプレート一括インポート（GASエディタから手動実行） ─────────
// 使い方: GASエディタで importTasksAsTemplates() を実行する
// 　→ 既存の物件タスクをテンプレートシートに取り込む
function importTasksAsTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 物件一覧を取得してUIで選択
  const props = rows(ss, SHEET_PROP);
  if (props.length === 0) {
    SpreadsheetApp.getUi().alert('物件が登録されていません');
    return;
  }
  const propList = props.map((p, i) => `${i+1}. ${p.name}`).join('\n');
  const res = SpreadsheetApp.getUi().inputBox(
    '物件を選択',
    '番号を入力してください:\n' + propList,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton().toString() !== 'OK') return;
  const idx = parseInt(res.getResponseText().trim()) - 1;
  if (isNaN(idx) || idx < 0 || idx >= props.length) {
    SpreadsheetApp.getUi().alert('番号が正しくありません');
    return;
  }
  const propName = props[idx].name;

  const taskRows  = rows(ss, SHEET_TASK).filter(r => r.prop === propName);
  const phaseRows = rows(ss, SHEET_PHASE).filter(r => r.prop === propName);
  const existing  = rows(ss, SHEET_TMPL);
  const sh = ss.getSheetByName(SHEET_TMPL);

  let added = 0, skipped = 0;
  taskRows.forEach(task => {
    const phase = phaseRows.find(p => p.id === task.phase_id);
    if (!phase) return;
    // 同じ工程・列・タスク名が既存なら重複スキップ
    const dup = existing.find(t =>
      t.phase_name === phase.name && t.col === task.col && t.name === task.name
    );
    if (dup) { skipped++; return; }
    sh.appendRow([uid(), phase.name, task.col, task.name, task.order, task.priority || false]);
    existing.push({ phase_name: phase.name, col: task.col, name: task.name }); // 重複チェック用
    added++;
  });

  SpreadsheetApp.getUi().alert(
    `「${propName}」からインポート完了\n追加: ${added}件　重複スキップ: ${skipped}件`
  );
}
