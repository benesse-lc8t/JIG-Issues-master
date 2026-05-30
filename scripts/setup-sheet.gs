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
 * 再実行しても安全（idempotent）：
 *   - 既存の列・タブ・行は壊さない
 *   - 不足している列のみ追加、既にある列は触らない
 *   - 条件付き書式は重複追加されうるので、必要なら「もう一度すべて整える」を実行
 */

// ===== 設定 =====
const STATUS_VALUES = ['未開始', '進行中', '完成'];
const STATUS_COLORS = { '未開始': '#EEEAE0', '進行中': '#DCEBFB', '完成': '#D4F2DD' };
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

// ===== メニュー登録 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('JIG')
    .addItem('セットアップ（不足分のみ追加）', 'setupJIG')
    .addItem('もう一度すべて整える（条件付き書式を入れ直す）', 'resetAndSetup')
    .addToUi();
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
    .setHelpText('未開始 / 進行中 / 完成 のいずれかを選択')
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
