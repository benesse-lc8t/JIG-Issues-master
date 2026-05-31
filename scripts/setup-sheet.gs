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
 * ★ スタンドアロン型でデプロイする場合（Sheet に紐づけない場合）
 *   - 下記 SHEET_ID にスプレッドシートの ID を設定すること
 *   - スプレッドシート URL の /d/XXXX/ の XXXX 部分が ID
 *   - Sheet 紐づけ型（Bound Script）なら SHEET_ID は空欄のままでよい
 *
 * 再実行しても安全（idempotent）：
 *   - 既存の列・タブ・行は壊さない
 *   - 不足している列のみ追加、既にある列は触らない
 *   - 条件付き書式は重複追加されうるので、必要なら「もう一度すべて整える」を実行
 */

// ===== 設定 =====
// スプレッドシートの ID（URL の /d/XXXX/ 部分）
// スタンドアロン型 Apps Script として Web App デプロイする場合は必須。
// Sheet 紐づけ型（Bound Script）なら空欄のままでも動く。
const SHEET_ID = '1C1dVsZ_7vfWO3fFUH9pHk1MCCNglAQxAaF5cHwjYo_4';
// 再公開が反映されたか確認するための目印。doGet が返す。変更のたびに上げる。
const CODE_VERSION = 'gs-2026-05-31-edit2';

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
  'Mission進度',   // F
  '更新日',        // G
  'Confluence URL' // H
];
const MISSION_COL_WIDTHS = { 1: 100, 2: 320, 3: 90, 4: 90, 5: 80, 6: 220, 7: 90, 8: 240 };

// ===== Mission進度ログ（追記専用ログ）=====
// 本人が週1で書く進度を 1 記入＝1 行で溜める。上書きしない。
// 回収・AI コンテキスト供給・停滞検知の燃料（CLAUDE.md §1.5）。
const LOG_SHEET_NAME = 'Mission進度ログ';
const LOG_HEADERS     = ['編號', '日時', '擔當', '進度'];
const LOG_COL_WIDTHS  = { 1: 110, 2: 140, 3: 90, 4: 480 };

// ===== 人員マスタ（名簿の一次ソース）=====
// 記入者チップの名前・色（處ベース）の出どころ。管理者がここに名簿を貼る。
const PERSON_SHEET_NAME = '人員';
const PERSON_HEADERS     = ['姓名', '處', '組', '顯示順'];
const PERSON_COL_WIDTHS  = { 1: 120, 2: 130, 3: 150, 4: 80 };

// ===== 編集 API 設定 =====
// 編集 API のトークン（index.html の WRITE_TOKEN と同じ値にする）
const WRITE_TOKEN = 'JIG-WRITE-TBBS-2026';
// 編集対象シート
const MISSION_SHEET_FOR_WRITE = 'Mission一覽';
const ISSUE_SHEET_FOR_WRITE   = 'Issue主檔'; // 新規 Issue 追加の対象
// ダッシュボード側のキー → Sheet 列名のマッピング
const WRITE_FIELDS = {
  '狀態':   '狀態',
  '備註':   'Mission進度',
  '更新日': '更新日',
  'Confluence URL': 'Confluence URL'   // 複数 URL は改行区切りで保持
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
      hasDate: '更新日' in data, hasConf: 'Confluence URL' in data, tokenMatch: data.token === WRITE_TOKEN
    }));
    console.log('  data keys:', JSON.stringify(Object.keys(data)));
    console.log('  備註 value:', JSON.stringify(data['備註']));
    console.log('  hasOwnProperty 備註:', Object.prototype.hasOwnProperty.call(data, '備註'));

    if (!data || data.token !== WRITE_TOKEN) {
      console.log('  → forbidden (token mismatch)');
      return _writeJson({ ok: false, error: 'forbidden' });
    }

    // 追加／編集アクション。action 無しは従来どおりの更新（狀態/進度ログ等）。
    const action = String(data.action || '').trim();
    if (action === 'addMission' || action === 'addIssue' || action === 'updateMission' || action === 'updateIssue') {
      const ssA = _getSpreadsheet();
      if (!ssA) return _writeJson({ ok: false, error: 'no_spreadsheet', hint: 'Set SHEET_ID in the script.' });
      if (action === 'addMission')    return _handleAddMission(ssA, data);
      if (action === 'addIssue')      return _handleAddIssue(ssA, data);
      if (action === 'updateMission') return _handleUpdate(ssA, data, MISSION_SHEET_FOR_WRITE);
      return _handleUpdate(ssA, data, ISSUE_SHEET_FOR_WRITE);
    }

    const missionId = String(data.mission || '').trim();
    if (!missionId) {
      console.log('  → mission_required');
      return _writeJson({ ok: false, error: 'mission_required' });
    }

    const ss = _getSpreadsheet();
    if (!ss) {
      console.log('  → no_spreadsheet. SHEET_ID=', SHEET_ID || '(empty)');
      return _writeJson({ ok: false, error: 'no_spreadsheet', hint: 'Set SHEET_ID in the script.' });
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

    // 進度ログ（追記）：本人の週次一行。`Mission進度ログ` へ 1 行 append し、
    // Mission一覽 の現況(Mission進度 H列)へ最新行をミラーする。上書きはしない。
    if (Object.prototype.hasOwnProperty.call(data, '進度')) {
      const lineText = String(data['進度']);
      if (lineText.trim()) {
        const author = String(data['擔當'] || '').trim();
        const when = new Date();
        const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
        const whenStr = Utilities.formatDate(when, tz, 'yyyy-MM-dd HH:mm'); // 文字列で保存（gviz 文字化け回避）
        let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
        if (!logSheet) {
          // 未作成なら自動生成（setupJIG を流し忘れても動くように）
          logSheet = ss.insertSheet(LOG_SHEET_NAME);
          logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
            .setFontWeight('bold').setBackground('#F7F4EC');
          logSheet.setFrozenRows(1);
          Object.keys(LOG_COL_WIDTHS).forEach(k => logSheet.setColumnWidth(Number(k), LOG_COL_WIDTHS[k]));
        }
        const mcol = headers.indexOf(WRITE_FIELDS['備註']) + 1;

        // 初回ログ時の移行：この Mission のログがまだ無く、既存の現況(H)が非空なら、
        // 既存内容を「過去分（（既存））」として 1 行 seed してから新規を追記する（消えないように）。
        let hasPriorLog = false;
        const logLast = logSheet.getLastRow();
        if (logLast >= 2) {
          const logIds = logSheet.getRange(2, 1, logLast - 1, 1).getValues();
          hasPriorLog = logIds.some(r => String(r[0]).trim() === missionId);
        }
        const existingCurrent = mcol > 0 ? String(sheet.getRange(rowIdx, mcol).getValue()).trim() : '';
        if (!hasPriorLog && existingCurrent && existingCurrent !== lineText) {
          let seedDate = '';
          const updColH = headers.indexOf(WRITE_FIELDS['更新日']) + 1;
          if (updColH > 0) {
            const uv = sheet.getRange(rowIdx, updColH).getValue();
            if (uv instanceof Date) seedDate = Utilities.formatDate(uv, tz, 'yyyy-MM-dd HH:mm');
            else if (uv) seedDate = String(uv);
          }
          logSheet.appendRow([missionId, seedDate, '（既存）', existingCurrent]);
          console.log('  進度ログ seed（既存現況を移行）:', existingCurrent.slice(0, 40));
        }

        // 新規行を追記
        logSheet.appendRow([missionId, whenStr, author, lineText]);
        // 現況(H=Mission進度)へミラー
        if (mcol > 0) {
          sheet.getRange(rowIdx, mcol).setValue(lineText);
          changes['備註'] = lineText;
        }
        changes['logEntry'] = { '日時': whenStr, '擔當': author, '進度': lineText };
        console.log('  進度ログ append:', JSON.stringify(changes['logEntry']));
        touchedContent = true;
      }
    }

    // Confluence URL（複数 URL は改行区切り）。行の更新なので 更新日 も自動更新する。
    if (Object.prototype.hasOwnProperty.call(data, 'Confluence URL')) {
      const col = headers.indexOf(WRITE_FIELDS['Confluence URL']) + 1;
      console.log('  Confluence URL col index:', col);
      if (col > 0) {
        sheet.getRange(rowIdx, col).setValue(String(data['Confluence URL']));
        changes['Confluence URL'] = String(data['Confluence URL']);
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
        changes['更新日'] = Utilities.formatDate(updValue, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
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

// ===== 行 append 系ヘルパー（新規 Issue / Mission 追加）=====

// ヘッダ名で列を特定して 1 行を追加する。
// 重要：`親編號` や `Confluence URL` などが**数式（ARRAYFORMULA 等）で自動算出**される
// 設計のシートがある。数式列に値を書くと #REF! でその列全体が壊れるため、
// (1) まずアンカー列（編號）だけ書いて数式に算出させ、
// (2) 数式で埋まった列／既に値がある列はスキップし、空の手入力列にだけ書く。
function _appendByHeaders(sheet, valueMap) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const colOf = name => headers.indexOf(name); // 0-based, -1 = 無し
  const targetRow = sheet.getLastRow() + 1;

  // (1) アンカー（編號）を先に書く → 数式列（親編號等）が自動で埋まる
  const idCol = colOf('編號');
  if (idCol >= 0 && valueMap['編號'] != null) {
    sheet.getRange(targetRow, idCol + 1).setValue(valueMap['編號']);
  }
  SpreadsheetApp.flush();

  // (2) 残りの値は「数式由来／自動算出済み」の列を避けて書く
  const skipped = [];   // 数式列（自動算出に任せた）
  const rejected = [];  // データ入力規則などで弾かれた列
  Object.keys(valueMap).forEach(name => {
    if (name === '編號') return;
    const ci = colOf(name);
    if (ci < 0) return;
    const cell = sheet.getRange(targetRow, ci + 1);
    const hasFormula = !!cell.getFormula();                 // 数式アンカー
    const autoFilled = String(cell.getValue()).trim() !== ''; // 数式スピルで既に埋まった
    if (hasFormula || autoFilled) { skipped.push(name); return; }
    try {
      cell.setValue(valueMap[name]);
      SpreadsheetApp.flush();   // 入力規則違反は flush で発火するため、ここで発火させ try 内で捕捉する
    } catch (e) {
      // データ入力規則（無効値を拒否）等で弾かれた → その列をクリアして行作成は続行
      rejected.push(name);
      try { cell.clearContent(); SpreadsheetApp.flush(); } catch (_) {}
      console.log('  _appendByHeaders rejected on', name, ':', String(e && e.message || e));
    }
  });
  SpreadsheetApp.flush();
  if (skipped.length)  console.log('  _appendByHeaders skipped formula/auto cols:', skipped.join(','));
  if (rejected.length) console.log('  _appendByHeaders rejected (validation) cols:', rejected.join(','));
  return { row: targetRow, headers: headers, skipped: skipped, rejected: rejected };
}

// 親編號配下の次の Mission 編號（親編號-M{n}）を採番する。
function _nextMissionId(sheet, parentId) {
  const last = sheet.getLastRow();
  let max = 0;
  if (last >= 2) {
    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    const esc = parentId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp('^' + esc + '-M(\\d+)$');
    ids.forEach(r => { const m = String(r[0]).trim().match(re); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  }
  return parentId + '-M' + (max + 1);
}

function _handleAddMission(ss, data) {
  const parent = String(data['親編號'] || '').trim();
  const name   = String(data['Mission'] || '').trim();
  if (!parent) return _writeJson({ ok: false, error: 'parent_required' });
  if (!name)   return _writeJson({ ok: false, error: 'mission_name_required' });
  const status = String(data['狀態'] || '').trim();
  if (status && !ALLOWED_STATUS.has(status)) return _writeJson({ ok: false, error: 'invalid_status', value: status });

  // 親 Issue の存在チェック（Issue主檔 の編號）
  const issueSheet = ss.getSheetByName(ISSUE_SHEET_FOR_WRITE);
  if (issueSheet) {
    const il = issueSheet.getLastRow();
    if (il >= 2) {
      const ihead = issueSheet.getRange(1, 1, 1, issueSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const inum = ihead.indexOf('編號');
      if (inum >= 0) {
        const ivals = issueSheet.getRange(2, inum + 1, il - 1, 1).getValues();
        if (!ivals.some(r => String(r[0]).trim() === parent)) {
          return _writeJson({ ok: false, error: 'parent_not_found', parent: parent });
        }
      }
    }
  }

  const sheet = ss.getSheetByName(MISSION_SHEET_FOR_WRITE);
  if (!sheet) return _writeJson({ ok: false, error: 'mission_sheet_not_found', sheetName: MISSION_SHEET_FOR_WRITE });
  const newId = _nextMissionId(sheet, parent);
  const vmap = {
    '編號': newId,
    'Mission': name,
    '親編號': parent,
    '戰略負責人': String(data['戰略負責人'] || ''),
    '擔當': String(data['擔當'] || ''),
    '狀態': status,
    '更新日': new Date(),
    'Confluence URL': String(data['Confluence URL'] || '')
  };
  const r = _appendByHeaders(sheet, vmap);
  SpreadsheetApp.flush();
  console.log('  → addMission ok:', newId, 'row', r.row);
  return _writeJson({ ok: true, action: 'addMission', mission: newId, '編號': newId, row: r.row, parent: parent, warnings: r.rejected });
}

// 既存行（編號で特定）の指定フィールドを更新。数式列は触らず、入力規則違反はスキップ。
function _updateRowByHeaders(sheet, rowIdx, valueMap) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const colOf = name => headers.indexOf(name);
  const changed = [], skipped = [], rejected = [];
  Object.keys(valueMap).forEach(name => {
    const ci = colOf(name);
    if (ci < 0) return;
    const cell = sheet.getRange(rowIdx, ci + 1);
    if (cell.getFormula()) { skipped.push(name); return; } // 数式列は上書きしない
    try {
      cell.setValue(valueMap[name]);
      SpreadsheetApp.flush(); // 入力規則違反はここで発火 → catch
      changed.push(name);
    } catch (e) {
      rejected.push(name); // 既存値はそのまま（クリアしない）
      console.log('  _updateRow rejected on', name, ':', String(e && e.message || e));
    }
  });
  return { changed: changed, skipped: skipped, rejected: rejected };
}

// updateMission / updateIssue 共通。編號で行を特定し、許可フィールドのみ反映。
// 編號・親編號・進度ログは対象外（編號は不変、進度は専用ログ経路）。
function _handleUpdate(ss, data, sheetName) {
  const id = String(data['編號'] || data.mission || '').trim();
  if (!id) return _writeJson({ ok: false, error: 'id_required' });
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return _writeJson({ ok: false, error: 'sheet_not_found', sheetName: sheetName });
  const last = sheet.getLastRow();
  if (last < 2) return _writeJson({ ok: false, error: 'no_data_rows' });
  // 編號の列はシートにより異なる（Mission一覽=A / Issue主檔=C）。ヘッダ名で特定する。
  const head0 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const numC = head0.indexOf('編號');
  if (numC < 0) return _writeJson({ ok: false, error: 'no_id_column', sheetName: sheetName });
  const ids = sheet.getRange(2, numC + 1, last - 1, 1).getValues();
  let rowIdx = -1;
  for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]).trim() === id) { rowIdx = i + 2; break; } }
  if (rowIdx < 0) return _writeJson({ ok: false, error: 'not_found', id: id });

  const status = String(data['狀態'] || '').trim();
  if (status && !ALLOWED_STATUS.has(status)) return _writeJson({ ok: false, error: 'invalid_status', value: status });

  const vmap = {};
  ['Issue', 'Mission', '擔當', '戰略負責人', '處', '組', '狀態', 'Confluence URL'].forEach(k => {
    if (Object.prototype.hasOwnProperty.call(data, k)) vmap[k] = String(data[k]);
  });
  vmap['更新日'] = new Date();
  const r = _updateRowByHeaders(sheet, rowIdx, vmap);
  SpreadsheetApp.flush();
  console.log('  → ' + (data.action || 'update') + ' ok:', id, 'row', rowIdx, 'changed', r.changed.join(','));
  return _writeJson({ ok: true, action: data.action, '編號': id, mission: id, row: rowIdx, changed: r.changed, warnings: r.rejected });
}

function _handleAddIssue(ss, data) {
  const id   = String(data['編號'] || '').trim();
  const name = String(data['Issue'] || '').trim();
  if (!id)   return _writeJson({ ok: false, error: 'issue_id_required' });
  if (!name) return _writeJson({ ok: false, error: 'issue_name_required' });
  const status = String(data['狀態'] || '').trim();
  if (status && !ALLOWED_STATUS.has(status)) return _writeJson({ ok: false, error: 'invalid_status', value: status });

  const sheet = ss.getSheetByName(ISSUE_SHEET_FOR_WRITE);
  if (!sheet) return _writeJson({ ok: false, error: 'issue_sheet_not_found', sheetName: ISSUE_SHEET_FOR_WRITE });

  // 編號の一意チェック
  const last = sheet.getLastRow();
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const numC = head.indexOf('編號');
  if (numC >= 0 && last >= 2) {
    const vals = sheet.getRange(2, numC + 1, last - 1, 1).getValues();
    if (vals.some(r => String(r[0]).trim() === id)) return _writeJson({ ok: false, error: 'duplicate_id', '編號': id });
  }

  const vmap = {
    '編號': id,
    'Issue': name,
    '處': String(data['處'] || ''),
    '組': String(data['組'] || ''),
    '戰略負責人': String(data['戰略負責人'] || ''),
    '狀態': status,
    '更新日': new Date(),
    'Confluence URL': String(data['Confluence URL'] || '')
  };
  const r = _appendByHeaders(sheet, vmap);
  SpreadsheetApp.flush();
  console.log('  → addIssue ok:', id, 'row', r.row);
  return _writeJson({ ok: true, action: 'addIssue', '編號': id, mission: id, row: r.row, warnings: r.rejected });
}

// 診断用：ブラウザで Web App URL を開くと、デプロイ済みコードが実際に見ている
// WRITE_FIELDS・シートのヘッダ・列インデックスを返す。
function doGet() {
  try {
    const ss = _getSpreadsheet();
    const sheet = ss ? ss.getSheetByName(MISSION_SHEET_FOR_WRITE) : null;
    const headers = sheet
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim())
      : null;
    const remarkTarget = WRITE_FIELDS['備註'];
    const remarkCol    = headers ? headers.indexOf(remarkTarget) : -99;
    const logSheet = ss ? ss.getSheetByName(LOG_SHEET_NAME) : null;
    return _writeJson({
      ok: true,
      codeVersion: CODE_VERSION,   // 再公開が反映されたか確認用（最新値が出れば反映済み）
      WRITE_FIELDS,
      sheetName:    MISSION_SHEET_FOR_WRITE,
      headers,
      remarkTarget,
      remarkColIndex: remarkCol,   // -1 なら列が見つかっていない
      logSheetName:   LOG_SHEET_NAME,
      logSheetExists: !!logSheet,  // false ならログタブ未作成（setupJIG 要実行）
      personSheetName:   PERSON_SHEET_NAME,
      personSheetExists: !!(ss && ss.getSheetByName(PERSON_SHEET_NAME)),
    });
  } catch (err) {
    return _writeJson({ ok: false, error: String(err.message) });
  }
}

function _getSpreadsheet() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _writeJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 診断用：Mission一覽 の全列ヘッダと最初の 3 行のデータを実行ログに出力する。
 * setupJIG でヘッダがズレた疑いがあるときに実行して確認する。
 */
function diagnoseMissionSheet() {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName('Mission一覽');
  if (!sheet) { console.log('Mission一覽 が見つかりません'); return; }
  const lastCol = sheet.getLastColumn();
  const rows = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 4), lastCol).getValues();
  const headers = rows[0];
  console.log('=== Mission一覽 列構成 ===');
  headers.forEach((h, i) => {
    const samples = rows.slice(1).map(r => String(r[i] || '').slice(0, 20)).join(' | ');
    console.log(`  Col ${i + 1} (${String.fromCharCode(65 + i)}): ヘッダ="${h}"  サンプル: ${samples}`);
  });
}

/**
 * ヘッダ修復用：diagnoseMissionSheet の結果に基づいた正しいヘッダに書き直す。
 * 診断結果：
 *   A=編號, B=Mission, C=親編號,
 *   D=管理頁面(データ: BCL登録用戸擴大 等), E=戦略負責人(人名),
 *   F=擔當(人名), G=狀態(進行中等), H=Mission進度(空),
 *   I=更新日(空), J=Confluence URL(空)
 */
function fixMissionHeaders() {
  const CORRECT_HEADERS = [
    '編號',           // A
    'Mission',        // B
    '親編號',         // C
    '管理頁面',       // D  ← BCL登録用戸擴大 等の継承ページ名
    '戰略負責人',     // E  ← 育菱 等の人名
    '擔當',           // F  ← 少琪 等の人名
    '狀態',           // G  ← 進行中 等のステータス
    'Mission進度',    // H  ← 空（書き込み対象）
    '更新日',         // I  ← 空（書き込み対象）
    'Confluence URL', // J  ← 空
  ];
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName('Mission一覽');
  if (!sheet) { console.log('Mission一覽 が見つかりません'); return; }
  sheet.getRange(1, 1, 1, CORRECT_HEADERS.length).setValues([CORRECT_HEADERS])
    .setFontWeight('bold').setBackground('#F7F4EC');
  console.log('[fixMissionHeaders] ヘッダ修復完了：', CORRECT_HEADERS.join(' / '));
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
  const ss = _getSpreadsheet();
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
    // 新規作成時のみ MISSION_HEADERS で 8 列を設定
    taskSheet = ss.insertSheet('Mission一覽');
    taskCreated = true;
    taskSheet.getRange(1, 1, 1, MISSION_HEADERS.length).setValues([MISSION_HEADERS])
      .setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(MISSION_COL_WIDTHS).forEach(k => taskSheet.setColumnWidth(Number(k), MISSION_COL_WIDTHS[k]));
    log.push('✓ 「Mission一覽」タブを作成（8 列）');
  } else {
    // 既存シートはヘッダを上書きせず、列名のリネームのみ行う（データ保護）
    const mHeaders = readHeaders(taskSheet);
    const RENAME_MAP = { '事務局備註': 'Mission進度', '備註': 'Mission進度', 'Task': 'Mission' };
    mHeaders.forEach((h, i) => {
      const newName = RENAME_MAP[String(h).trim()];
      if (newName) {
        taskSheet.getRange(1, i + 1).setValue(newName).setFontWeight('bold').setBackground('#F7F4EC');
        log.push(`✓ Mission一覽 列「${h}」→「${newName}」に変更`);
      }
    });
  }

  // 行 freeze
  taskSheet.setFrozenRows(1);

  // 狀態列のプルダウン・条件付き書式（列名で位置を特定）
  const mHeaders2  = readHeaders(taskSheet);
  const mStatusCol = mHeaders2.indexOf('狀態') + 1;
  const mUpdCol    = mHeaders2.indexOf('更新日') + 1;
  const taskRows   = Math.max(taskSheet.getMaxRows() - 1, 1000);
  if (mStatusCol > 0) applyStatusValidation(taskSheet.getRange(2, mStatusCol, taskRows, 1));

  const taskRules = taskSheet.getConditionalFormatRules();
  let added1 = 0, added2 = 0;
  if (mStatusCol > 0) added1 = addStatusColorRulesIfMissing(taskRules, taskSheet.getRange(2, mStatusCol, taskRows, 1));
  if (mUpdCol    > 0) added2 = addFreshnessRulesIfMissing(taskRules, taskSheet.getRange(2, mUpdCol, taskRows, 1), mUpdCol);
  if (added1 + added2 > 0) {
    taskSheet.setConditionalFormatRules(taskRules);
    log.push(`✓ Mission一覽 に条件付き書式 ${added1 + added2} 件を追加`);
  }
  if (taskCreated) log.push('  　└ 列幅・行 freeze・プルダウンも設定済み');

  // --- 3. Mission進度ログ タブ（追記専用）---
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
      .setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(LOG_COL_WIDTHS).forEach(k => logSheet.setColumnWidth(Number(k), LOG_COL_WIDTHS[k]));
    logSheet.setFrozenRows(1);
    log.push('✓ 「Mission進度ログ」タブを作成（追記専用・4列）');
  }

  // --- 4. 人員 タブ（名簿の一次ソース）---
  let personSheet = ss.getSheetByName(PERSON_SHEET_NAME);
  if (!personSheet) {
    personSheet = ss.insertSheet(PERSON_SHEET_NAME);
    personSheet.getRange(1, 1, 1, PERSON_HEADERS.length).setValues([PERSON_HEADERS])
      .setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(PERSON_COL_WIDTHS).forEach(k => personSheet.setColumnWidth(Number(k), PERSON_COL_WIDTHS[k]));
    personSheet.setFrozenRows(1);
    log.push('✓ 「人員」タブを作成（姓名 / 處 / 組 / 顯示順）。ここに名簿を貼ると記入者チップに色が付きます');
  }

  const msg = log.length ? '✅ セットアップ完了\n\n' + log.join('\n')
                         : 'ℹ️ 既に整っています（追加・変更なし）';
  console.log('[setupJIG] ' + msg);
}

// ===== 条件付き書式をクリアして入れ直す（重複が気になったとき用）=====
function resetAndSetup() {
  const ss = _getSpreadsheet();
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
