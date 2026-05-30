/**
 * TBBS-JIG 管理ダッシュボード — スプレッドシート セットアップスクリプト
 *
 * 使い方
 *   1. 対象のスプレッドシートを開く
 *   2. メニュー：拡張機能 → Apps Script
 *   3. このファイルの中身を全てコピー＆貼り付け（既存コードは置き換え）
 *   4. 保存（💾）→ 上部の関数選択で `setupJIG` を選び ▶ 実行
 *   5. 初回のみ権限承認が必要（このスクリプトを自分のアカウントで実行する許可）
 *   6. 完了。スプレッドシートに戻ると「JIG」メニューが出る（次回はメニューから 1 クリック）
 *
 * 編集 API（インライン編集を有効化したい場合）
 *   1. 上記セットアップ後、エディタ右上の「デプロイ」→「新しいデプロイ」
 *   2. 種類：ウェブアプリ
 *   3. 説明：JIG Mission 編集 API
 *   4. 実行ユーザー：自分
 *   5. アクセスできるユーザー：全員
 *   6. 「デプロイ」→ 権限承認 → Web App URL をコピー
 *   7. index.html の SHEET_WRITE_URL に貼り付け
 *   ※ WRITE_TOKEN（下記）は index.html 側と同じ値にする
 *
 * 再実行しても安全（idempotent）：
 *   - 既存の列・タブ・行は壊さない
 *   - 不足している列のみ追加、既にある列は触らない
 *   - 条件付き書式は重複追加されうるので、必要なら「もう一度すべて整える」を実行
 */

// ===== 設定 =====
const STATUS_VALUES = ['未開始', '策劃中', '需確認', '進行中', '結案'];
const STATUS_COLORS = {
  '未開始': '#EEEAE0',
  '策劃中': '#D1F2F7',
  '需確認': '#FEE7BB',
  '進行中': '#DCEBFB',
  '結案':   '#D4F2DD'
};
const ISSUE_NEW_COLS = ['Confluence URL', '狀態', '事務局備註', '更新日'];
const MISSION_HEADERS = [
  '編號',          // A
  'Mission',       // B
  '親編號',        // C
  '戰略負責人',    // D
  '狀態',          // E
  '事務局備註',    // F
  '更新日',        // G
  'Confluence URL' // H
];
const MISSION_COL_WIDTHS = { 1: 100, 2: 320, 3: 90, 4: 90, 5: 80, 6: 220, 7: 90, 8: 240 };

// ===== 編集 API 設定 =====
// 編集 API のトークン（index.html の WRITE_TOKEN と同じ値にする）
const WRITE_TOKEN = 'JIG-WRITE-TBBS-2026';
// 編集対象シート
const MISSION_SHEET_FOR_WRITE = 'Mission一覽';
// ダッシュボード側のキー → Sheet 列名のマッピング
const WRITE_FIELDS = {
  '狀態':   '狀態',
  '備註':   '事務局備註',
  '更新日': '更新日'
};
const ALLOWED_STATUS = new Set(STATUS_VALUES);

// ===== メニュー登録 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('JIG')
    .addItem('セットアップ（不足分のみ追加）', 'setupJIG')
    .addItem('もう一度すべて整える（条件付き書式を入れ直す）', 'resetAndSetup')
    .addToUi();
}

// ===== 編集 API（Web App エンドポイント）=====
// ダッシュボードから fetch POST で呼ばれる。
// 本文（Content-Type: text/plain）に JSON：
//   { token, mission: '戰略-1-M1', 狀態: '進行中', 備註: '...', 更新日: '2026-05-30' }
// 狀態 / 備註 のみが指定されていれば、更新日は自動で today にセット。
// 更新日が明示されていればそちらを優先（手動補正可）。
function doPost(e) {
  try {
    console.log('=== doPost start ===');
    console.log('e.postData:', e && e.postData ? JSON.stringify({ type: e.postData.type, length: (e.postData.contents||'').length }) : 'null');
    if (!e || !e.postData || !e.postData.contents) {
      console.log('  → no_payload');
      return _writeJson({ ok: false, error: 'no_payload' });
    }
    let data;
    try { data = JSON.parse(e.postData.contents); }
    catch (_) {
      console.log('  → invalid_json. raw =', e.postData.contents.slice(0, 200));
      return _writeJson({ ok: false, error: 'invalid_json' });
    }
    console.log('  parsed payload:', JSON.stringify({
      mission: data.mission, hasStatus: '狀態' in data, hasRemark: '備註' in data,
      hasDate: '更新日' in data, tokenMatch: data.token === WRITE_TOKEN
    }));

    if (!data || data.token !== WRITE_TOKEN) {
      console.log('  → forbidden (token mismatch)');
      return _writeJson({ ok: false, error: 'forbidden' });
    }

    const missionId = String(data.mission || '').trim();
    if (!missionId) {
      console.log('  → mission_required');
      return _writeJson({ ok: false, error: 'mission_required' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.log('  → no_active_spreadsheet (script may be standalone, not Sheet-bound)');
      return _writeJson({ ok: false, error: 'no_active_spreadsheet' });
    }
    console.log('  spreadsheet:', ss.getName(), ' id:', ss.getId());
    console.log('  available sheets:', ss.getSheets().map(s => s.getName()).join(' / '));
    const sheet = ss.getSheetByName(MISSION_SHEET_FOR_WRITE);
    if (!sheet) {
      console.log('  → mission_sheet_not_found (looking for', MISSION_SHEET_FOR_WRITE, ')');
      return _writeJson({ ok: false, error: 'mission_sheet_not_found', sheetName: MISSION_SHEET_FOR_WRITE });
    }

    // 編號（A 列）で行を特定
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return _writeJson({ ok: false, error: 'no_data_rows' });
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowIdx = -1;
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === missionId) { rowIdx = i + 2; break; }
    }
    if (rowIdx < 0) {
      console.log('  → mission_not_found:', missionId);
      console.log('  sample IDs in sheet:', ids.slice(0, 5).map(r => JSON.stringify(r[0])).join(', '));
      return _writeJson({ ok: false, error: 'mission_not_found', mission: missionId });
    }
    console.log('  row index:', rowIdx);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    console.log('  headers:', JSON.stringify(headers));
    const changes = {};
    let touchedContent = false;

    // 狀態
    if (Object.prototype.hasOwnProperty.call(data, '狀態')) {
      const v = String(data['狀態']).trim();
      if (v && !ALLOWED_STATUS.has(v)) return _writeJson({ ok: false, error: 'invalid_status', value: v });
      const col = headers.indexOf(WRITE_FIELDS['狀態']) + 1;
      console.log('  狀態 col index:', col, '(WRITE_FIELDS["狀態"]=', WRITE_FIELDS['狀態'], ')');
      if (col > 0) {
        sheet.getRange(rowIdx, col).setValue(v);
        changes['狀態'] = v;
        touchedContent = true;
      }
    }

    // 備註
    if (Object.prototype.hasOwnProperty.call(data, '備註')) {
      const col = headers.indexOf(WRITE_FIELDS['備註']) + 1;
      console.log('  備註 col index:', col, '(WRITE_FIELDS["備註"]=', WRITE_FIELDS['備註'], ')');
      if (col > 0) {
        sheet.getRange(rowIdx, col).setValue(String(data['備註']));
        changes['備註'] = String(data['備註']);
        touchedContent = true;
      }
    }

    // 更新日：明示指定があればそれを、なければ狀態/備註変更時に today を自動セット
    let updValue = null;
    if (Object.prototype.hasOwnProperty.call(data, '更新日')) {
      updValue = _parseDateLooseGS(String(data['更新日']));
      if (!updValue && String(data['更新日']).trim()) {
        return _writeJson({ ok: false, error: 'invalid_date', value: data['更新日'] });
      }
    }
    if (!updValue && touchedContent) {
      updValue = new Date();
    }
    if (updValue) {
      const col = headers.indexOf(WRITE_FIELDS['更新日']) + 1;
      console.log('  更新日 col index:', col);
      if (col > 0) {
        sheet.getRange(rowIdx, col).setValue(updValue);
        changes['更新日'] = Utilities.formatDate(updValue, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd');
      }
    }

    SpreadsheetApp.flush(); // 即時反映を保証
    console.log('  → ok, changes:', JSON.stringify(changes));
    return _writeJson({ ok: true, mission: missionId, row: rowIdx, changes });
  } catch (err) {
    console.log('  → exception:', String(err && err.stack || err));
    return _writeJson({ ok: false, error: 'exception', message: String(err && err.message || err) });
  }
}

// 動作確認用：ブラウザで URL を直接開くと簡易な OK 応答を返す
function doGet() {
  return _writeJson({ ok: true, service: 'JIG Mission write API', version: 1 });
}

function _writeJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _parseDateLooseGS(v) {
  if (!v) return null;
  v = String(v).trim();
  let m;
  m = v.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  m = v.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m) {
    const a = +m[1], b = +m[2], y = +m[3];
    return (a > 12) ? new Date(y, b-1, a) : new Date(y, a-1, b);
  }
  const p = Date.parse(v);
  return isNaN(p) ? null : new Date(p);
}

// ===== メインのセットアップ =====
function setupJIG() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // --- 1. Issue主檔 の列補完 ---
  const issueSheet = ss.getSheetByName('Issue主檔');
  if (!issueSheet) {
    SpreadsheetApp.getUi().alert('「Issue主檔」シートが見つかりません。シート名を確認してください。');
    return;
  }
  let headers = readHeaders(issueSheet);
  ISSUE_NEW_COLS.forEach(name => {
    if (!headers.includes(name)) {
      const col = issueSheet.getLastColumn() + 1;
      issueSheet.getRange(1, col).setValue(name)
        .setFontWeight('bold').setBackground('#F7F4EC');
      log.push(`✓ Issue主檔 に列「${name}」を追加`);
    }
  });
  headers = readHeaders(issueSheet);

  // Issue主檔 の 狀態・更新日 にプルダウンと色付け
  const issueRows = Math.max(issueSheet.getMaxRows() - 1, 100);
  const issueStatusCol = headers.indexOf('狀態') + 1;
  const issueUpdCol    = headers.indexOf('更新日') + 1;
  if (issueStatusCol > 0) {
    applyStatusValidation(issueSheet.getRange(2, issueStatusCol, issueRows, 1));
  }
  const issueRules = issueSheet.getConditionalFormatRules();
  let issueRulesAdded = 0;
  if (issueStatusCol > 0) {
    issueRulesAdded += addStatusColorRulesIfMissing(issueRules,
      issueSheet.getRange(2, issueStatusCol, issueRows, 1));
  }
  if (issueUpdCol > 0) {
    issueRulesAdded += addFreshnessRulesIfMissing(issueRules,
      issueSheet.getRange(2, issueUpdCol, issueRows, 1), issueUpdCol);
  }
  if (issueRulesAdded > 0) {
    issueSheet.setConditionalFormatRules(issueRules);
    log.push(`✓ Issue主檔 に条件付き書式 ${issueRulesAdded} 件を追加`);
  }

  // --- 2. Mission一覽 タブ ---
  let taskSheet = ss.getSheetByName('Mission一覽');
  let taskCreated = false;
  if (!taskSheet) {
    taskSheet = ss.insertSheet('Mission一覽');
    taskCreated = true;
    log.push('✓ 「Mission一覽」タブを作成');
  }
  // ヘッダ
  taskSheet.getRange(1, 1, 1, MISSION_HEADERS.length).setValues([MISSION_HEADERS])
    .setFontWeight('bold').setBackground('#F7F4EC');

  // 列幅
  Object.keys(MISSION_COL_WIDTHS).forEach(k => taskSheet.setColumnWidth(Number(k), MISSION_COL_WIDTHS[k]));

  // 行 freeze
  taskSheet.setFrozenRows(1);

  // プルダウン（狀態：E列）
  const taskRows = Math.max(taskSheet.getMaxRows() - 1, 1000);
  applyStatusValidation(taskSheet.getRange(2, 5, taskRows, 1));

  // 条件付き書式
  const taskRules = taskSheet.getConditionalFormatRules();
  const added1 = addStatusColorRulesIfMissing(taskRules, taskSheet.getRange(2, 5, taskRows, 1));
  const added2 = addFreshnessRulesIfMissing(taskRules, taskSheet.getRange(2, 7, taskRows, 1), 7);
  if (added1 + added2 > 0) {
    taskSheet.setConditionalFormatRules(taskRules);
    log.push(`✓ Mission一覽 に条件付き書式 ${added1 + added2} 件を追加`);
  }
  if (taskCreated) {
    log.push('  　└ 列幅・行 freeze・プルダウンも設定済み');
  }

  SpreadsheetApp.getUi().alert(
    log.length ? '✅ セットアップ完了\n\n' + log.join('\n')
               : 'ℹ️ 既に整っています（追加・変更なし）'
  );
}

// ===== 条件付き書式をクリアして入れ直す（重複が気になったとき用）=====
function resetAndSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Issue主檔', 'Mission一覽'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.setConditionalFormatRules([]);  // 全クリア（注意：手動で入れたルールも消える）
  });
  setupJIG();
}

// ===== ヘルパー =====
function readHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
}

function applyStatusValidation(range) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALUES, true)
    .setAllowInvalid(false)
    .setHelpText('未開始 / 策劃中 / 需確認 / 進行中 / 結案 のいずれかを選択')
    .build();
  range.setDataValidation(rule);
}

// 同じレンジ・同じテキストのルールが既にあれば追加しない
function addStatusColorRulesIfMissing(rules, range) {
  let added = 0;
  STATUS_VALUES.forEach(v => {
    const exists = rules.some(r => {
      try {
        const conds = r.getBooleanCondition && r.getBooleanCondition();
        if (!conds) return false;
        const txt = conds.getCriteriaValues && conds.getCriteriaValues();
        const onSameRange = (r.getRanges() || []).some(rg => rg.getA1Notation() === range.getA1Notation());
        return onSameRange && txt && String(txt[0]) === v;
      } catch (e) { return false; }
    });
    if (!exists) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(v)
        .setBackground(STATUS_COLORS[v])
        .setRanges([range])
        .build());
      added++;
    }
  });
  return added;
}

function addFreshnessRulesIfMissing(rules, range, colNum) {
  const colLetter = colToLetter(colNum);
  const exprWarn  = `=AND($${colLetter}2<>"", TODAY()-$${colLetter}2>7, TODAY()-$${colLetter}2<=14)`;
  const exprStale = `=AND($${colLetter}2<>"", TODAY()-$${colLetter}2>14)`;
  const hasExpr = (expr) => rules.some(r => {
    try {
      const c = r.getBooleanCondition && r.getBooleanCondition();
      if (!c) return false;
      const vals = c.getCriteriaValues && c.getCriteriaValues();
      const onSameRange = (r.getRanges() || []).some(rg => rg.getA1Notation() === range.getA1Notation());
      return onSameRange && vals && String(vals[0]) === expr;
    } catch (e) { return false; }
  });
  let added = 0;
  if (!hasExpr(exprWarn)) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(exprWarn)
      .setBackground('#FFF3CC').setFontColor('#7A5A00')
      .setRanges([range]).build());
    added++;
  }
  if (!hasExpr(exprStale)) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(exprStale)
      .setBackground('#FCD7D7').setFontColor('#8A0000')
      .setRanges([range]).build());
    added++;
  }
  return added;
}

function colToLetter(c) {
  let s = '';
  let n = c;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
