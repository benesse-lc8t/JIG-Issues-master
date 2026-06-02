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
const CODE_VERSION = 'gs-2026-06-02-taskdef1';

// ===== 状態（2026-06-01 リデザイン：Mission/Task は7状態）=====
// REDESIGN-PLAN.md D3。Issue は7状態を付けない（D2）。
// 旧5状態（未開始/需確認）は表示側で読み替え（未開始→構想中・需確認→待審核）。
const STATUS_VALUES = ['構想中', '策劃中', '待審核', '進行中', '結案', '凍結', '中止'];
const STATUS_COLORS = {
  '構想中': '#EEEAE0',
  '策劃中': '#D1F2F7',
  '待審核': '#FEE7BB',
  '進行中': '#DCEBFB',
  '結案':   '#D4F2DD',
  '凍結':   '#E2E3E5',
  '中止':   '#F3D9D9'
};
// 旧状態（移行期の互換。ドロップダウンには出さないが、書込は許容して弾かない）
const LEGACY_STATUS = ['未開始', '需確認'];
const ISSUE_NEW_COLS = ['Confluence URL', '狀態', '事務局備註', '更新日', 'Issue定義', '協作'];
const MISSION_HEADERS = [
  '編號',          // A
  'Mission',       // B
  '親編號',        // C
  '戰略負責人',    // D
  '狀態',          // E
  'Mission進度',   // F
  '更新日',        // G
  'Confluence URL',// H
  'Mission定義'    // I（2026-06-01 追加：カスケード健全性の言語化＝D8）
];
const MISSION_COL_WIDTHS = { 1: 100, 2: 320, 3: 90, 4: 90, 5: 80, 6: 220, 7: 90, 8: 240, 9: 320 };

// ===== Task一覽（2026-06-01 リデザイン：3層目の実体）=====
// REDESIGN-PLAN.md §3.2。親 Mission 編號の配下に -K{n} で採番。
// 協作は列内マルチ値（区切り , / 、）。連結は「名前|URL」を複数（改行/カンマ区切り）。
const TASK_SHEET_NAME = 'Task一覽';
const TASK_HEADERS = [
  '編號',     // A
  'Task',     // B
  '親編號',   // C（Mission一覽 の編號）
  'DRI',      // D（組員。空なら親 Mission 継承）
  '協作',     // E（複数可・列内マルチ値）
  '狀態',     // F（7状態）
  '進度',     // G（一行・短文）
  '更新日',   // H
  '連結',     // I（名前|URL を複数）
  'Task定義'  // J（2026-06-02 追加：カスケード健全性＝Mission定義の下位整合）
];
const TASK_COL_WIDTHS = { 1: 130, 2: 300, 3: 110, 4: 90, 5: 120, 6: 80, 7: 240, 8: 100, 9: 240, 10: 320 };

// ===== Mission進度ログ（追記専用ログ）=====
// 本人が週1で書く進度を 1 記入＝1 行で溜める。上書きしない。
// 回収・AI コンテキスト供給・停滞検知の燃料（CLAUDE.md §1.5）。
const LOG_SHEET_NAME = 'Mission進度ログ';
const LOG_HEADERS     = ['編號', '日時', '擔當', '進度'];
const LOG_COL_WIDTHS  = { 1: 110, 2: 140, 3: 90, 4: 480 };

// ===== 個人備註（事務局メンバーの私的メモ）=====
// 詩雅／育菱が自分の Mission に書く私的メモ。1人1ミッション1行（編號+姓名で upsert）。
// UI 上は本人のタブにのみ表示（厳密非公開ではない＝§7 の制約は受容）。
const MEMO_SHEET_NAME = '個人備註';
const MEMO_HEADERS     = ['編號', '姓名', '備註', '留言', '重要', '更新日'];
const MEMO_COL_WIDTHS  = { 1: 110, 2: 90, 3: 340, 4: 340, 5: 60, 6: 140 };

// ===== 公告（トップのお知らせボード）=====
// 管理者が編集、全員に読み取り表示。1 件のみ（2 行目）。
const ANNOUNCE_SHEET_NAME = '公告';
const ANNOUNCE_HEADERS    = ['訊息', '更新日'];
const ANNOUNCE_COL_WIDTHS = { 1: 600, 2: 140 };

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
// 書込許容：7状態＋旧2状態（移行期の互換）。ドロップダウンは STATUS_VALUES のみ。
const ALLOWED_STATUS = new Set([...STATUS_VALUES, ...LEGACY_STATUS]);

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
    if (action === 'addMission' || action === 'addIssue' || action === 'addTask' || action === 'updateMission' || action === 'updateIssue' || action === 'updateTask' || action === 'saveMemo' || action === 'saveAnnounce') {
      const ssA = _getSpreadsheet();
      if (!ssA) return _writeJson({ ok: false, error: 'no_spreadsheet', hint: 'Set SHEET_ID in the script.' });
      if (action === 'addMission')    return _handleAddMission(ssA, data);
      if (action === 'addIssue')      return _handleAddIssue(ssA, data);
      if (action === 'addTask')       return _handleAddTask(ssA, data);
      if (action === 'updateMission') return _handleUpdate(ssA, data, MISSION_SHEET_FOR_WRITE);
      if (action === 'updateIssue')   return _handleUpdate(ssA, data, ISSUE_SHEET_FOR_WRITE);
      if (action === 'updateTask')    return _handleUpdate(ssA, data, TASK_SHEET_NAME);
      if (action === 'saveAnnounce')  return _handleSaveAnnounce(ssA, data);
      return _handleSaveMemo(ssA, data);
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
    'Confluence URL': String(data['Confluence URL'] || ''),
    'Mission定義': String(data['Mission定義'] || '')
  };
  const r = _appendByHeaders(sheet, vmap);
  SpreadsheetApp.flush();
  console.log('  → addMission ok:', newId, 'row', r.row);
  return _writeJson({ ok: true, action: 'addMission', mission: newId, '編號': newId, row: r.row, parent: parent, warnings: r.rejected });
}

// 親 Mission 編號配下の次の Task 編號（親編號-K{n}）を採番する。
function _nextTaskId(sheet, parentId) {
  const last = sheet.getLastRow();
  let max = 0;
  if (last >= 2) {
    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    const esc = parentId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp('^' + esc + '-T(\\d+)$');
    ids.forEach(r => { const m = String(r[0]).trim().match(re); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  }
  return parentId + '-T' + (max + 1); // -K から -T に統一（表示/移行と揃える）
}

// 残骸 Task を掃除：編號が旧 -K 採番、または親が Mission一覽 に存在しない（孤児）行を削除。
function cleanupOrphanTasks() {
  const ss = _getSpreadsheet();
  if (!ss) return;
  const tSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!tSheet || tSheet.getLastRow() < 2) { console.log('[cleanup] Task一覽 なし'); return; }
  const mSheet = ss.getSheetByName(MISSION_SHEET_FOR_WRITE);
  const colOf = (sh, n) => sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim()).indexOf(n);
  const missionIds = new Set();
  if (mSheet && mSheet.getLastRow() >= 2) {
    const mc = colOf(mSheet, '編號');
    if (mc >= 0) mSheet.getRange(2, mc + 1, mSheet.getLastRow() - 1, 1).getValues().forEach(r => { const v = String(r[0]).trim(); if (v) missionIds.add(v); });
  }
  const tNum = colOf(tSheet, '編號'), tPar = colOf(tSheet, '親編號'), tLast = tSheet.getLastRow();
  const nums = tSheet.getRange(2, tNum + 1, tLast - 1, 1).getValues();
  const pars = tPar >= 0 ? tSheet.getRange(2, tPar + 1, tLast - 1, 1).getValues() : null;
  let removed = 0;
  for (let i = nums.length - 1; i >= 0; i--) {
    const id = String(nums[i][0]).trim();
    const par = pars ? String(pars[i][0]).trim() : '';
    const isOldK = /-K\d+$/.test(id);
    const isOrphan = par && missionIds.size && !missionIds.has(par);
    if (!id || isOldK || isOrphan) { tSheet.deleteRow(i + 2); removed++; }
  }
  SpreadsheetApp.flush();
  console.log('[cleanup] 削除 ' + removed + ' 行（旧-K採番・孤児・空）');
}

// Task の追加（3層目）。親は Mission一覽 の編號。-K{n} で自動採番。
function _handleAddTask(ss, data) {
  const parent = String(data['親編號'] || '').trim();
  const name   = String(data['Task'] || '').trim();
  if (!parent) return _writeJson({ ok: false, error: 'parent_required' });
  if (!name)   return _writeJson({ ok: false, error: 'task_name_required' });
  const status = String(data['狀態'] || '').trim();
  if (status && !ALLOWED_STATUS.has(status)) return _writeJson({ ok: false, error: 'invalid_status', value: status });

  // Task一覽 が無ければ作成（setupJIG 未実行でも初回 add で器を用意）
  let sheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TASK_SHEET_NAME);
    sheet.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS])
      .setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(TASK_COL_WIDTHS).forEach(k => sheet.setColumnWidth(Number(k), TASK_COL_WIDTHS[k]));
    sheet.setFrozenRows(1);
  }

  // 親 Mission の存在チェック（Mission一覽 の編號）
  const mSheet = ss.getSheetByName(MISSION_SHEET_FOR_WRITE);
  if (mSheet) {
    const ml = mSheet.getLastRow();
    if (ml >= 2) {
      const mhead = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const mnum = mhead.indexOf('編號');
      if (mnum >= 0) {
        const mvals = mSheet.getRange(2, mnum + 1, ml - 1, 1).getValues();
        if (!mvals.some(r => String(r[0]).trim() === parent)) {
          return _writeJson({ ok: false, error: 'parent_not_found', parent: parent });
        }
      }
    }
  }

  const newId = _nextTaskId(sheet, parent);
  const vmap = {
    '編號': newId,
    'Task': name,
    '親編號': parent,
    'DRI': String(data['DRI'] || ''),
    '協作': String(data['協作'] || ''),
    '狀態': status,
    '進度': String(data['進度'] || ''),
    '更新日': new Date(),
    '連結': String(data['連結'] || data['Confluence URL'] || ''),
    'Task定義': String(data['Task定義'] || '')
  };
  const r = _appendByHeaders(sheet, vmap);
  SpreadsheetApp.flush();
  console.log('  → addTask ok:', newId, 'row', r.row);
  return _writeJson({ ok: true, action: 'addTask', task: newId, '編號': newId, row: r.row, parent: parent, warnings: r.rejected });
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
  // 編號は先頭ゼロ（"039"）と数値化（39）が混在し得るので正規化して照合
  const _normId = s => { s = String(s == null ? '' : s).trim(); return /^\d+$/.test(s) ? String(parseInt(s, 10)) : s; };
  const idN = _normId(id);
  let rowIdx = -1;
  for (let i = 0; i < ids.length; i++) { const c = String(ids[i][0]).trim(); if (c === id || _normId(c) === idN) { rowIdx = i + 2; break; } }
  if (rowIdx < 0) return _writeJson({ ok: false, error: 'not_found', id: id });

  const status = String(data['狀態'] || '').trim();
  if (status && !ALLOWED_STATUS.has(status)) return _writeJson({ ok: false, error: 'invalid_status', value: status });

  const vmap = {};
  // Issue/Mission/Task で使い得るフィールドを一括許容（対象シートに無い列は colOf=-1 でスキップ）
  ['Issue', 'Mission', 'Task', '擔當', '戰略負責人', 'DRI', '協作', '處', '組',
   '狀態', 'Confluence URL', '連結', '進度', 'Issue定義', 'Mission定義', 'Task定義', '事務局備註'].forEach(k => {
    if (Object.prototype.hasOwnProperty.call(data, k)) vmap[k] = String(data[k]);
  });
  vmap['更新日'] = new Date();
  const r = _updateRowByHeaders(sheet, rowIdx, vmap);
  SpreadsheetApp.flush();
  // logProgress フラグ付きで 進度 が来たら進度ログへ追記（モーダルの Progress 記入）
  if (data['logProgress'] && String(data['進度'] || '').trim()) {
    _appendLog(ss, id, String(data['author'] || data['擔當'] || ''), String(data['進度']).trim());
  }
  console.log('  → ' + (data.action || 'update') + ' ok:', id, 'row', rowIdx, 'changed', r.changed.join(','));
  return _writeJson({ ok: true, action: data.action, '編號': id, mission: id, row: rowIdx, changed: r.changed, warnings: r.rejected });
}

// 個人備註の保存（編號＋姓名で upsert。上書き可＝進度ログとは別）。
function _handleSaveMemo(ss, data) {
  const id  = String(data['編號'] || '').trim();
  const who = String(data['姓名'] || '').trim();
  if (!id || !who) return _writeJson({ ok: false, error: 'id_or_name_required' });

  let sheet = ss.getSheetByName(MEMO_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MEMO_SHEET_NAME);
    sheet.getRange(1, 1, 1, MEMO_HEADERS.length).setValues([MEMO_HEADERS]).setFontWeight('bold').setBackground('#F7F4EC');
    sheet.setFrozenRows(1);
    Object.keys(MEMO_COL_WIDTHS).forEach(k => sheet.setColumnWidth(Number(k), MEMO_COL_WIDTHS[k]));
  }
  let lastCol = sheet.getLastColumn();
  let head = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  // 不足列（留言 / 重要 等）を必要に応じて追記
  MEMO_HEADERS.forEach(name => {
    if (head.indexOf(name) < 0) { sheet.getRange(1, sheet.getLastColumn() + 1).setValue(name).setFontWeight('bold'); }
  });
  lastCol = sheet.getLastColumn();
  head = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const ci = { id: head.indexOf('編號'), who: head.indexOf('姓名'), upd: head.indexOf('更新日') };
  if (ci.id < 0 || ci.who < 0) return _writeJson({ ok: false, error: 'memo_headers_invalid', headers: head });
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  // 指定されたフィールドだけ書く（備註 / 留言 / 重要 のうち届いたもの）
  const writes = {};
  ['備註', '留言', '重要'].forEach(k => { if (Object.prototype.hasOwnProperty.call(data, k)) writes[k] = String(data[k]); });

  const last = sheet.getLastRow();
  let rowIdx = -1;
  if (last >= 2) {
    const vals = sheet.getRange(2, 1, last - 1, lastCol).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][ci.id]).trim() === id && String(vals[i][ci.who]).trim() === who) { rowIdx = i + 2; break; }
    }
  }
  if (rowIdx < 0) {
    const row = new Array(lastCol).fill('');
    row[ci.id] = id; row[ci.who] = who;
    Object.keys(writes).forEach(k => { const c = head.indexOf(k); if (c >= 0) row[c] = writes[k]; });
    if (ci.upd >= 0) row[ci.upd] = now;
    sheet.appendRow(row);
    rowIdx = sheet.getLastRow();
  } else {
    Object.keys(writes).forEach(k => { const c = head.indexOf(k); if (c >= 0) sheet.getRange(rowIdx, c + 1).setValue(writes[k]); });
    if (ci.upd >= 0) sheet.getRange(rowIdx, ci.upd + 1).setValue(now);
  }
  SpreadsheetApp.flush();
  console.log('  → saveMemo ok:', id, who, 'row', rowIdx, 'fields', Object.keys(writes).join(','));
  return _writeJson({ ok: true, action: 'saveMemo', '編號': id, '姓名': who, mission: id, row: rowIdx, '更新': now });
}

// 公告（お知らせ）の保存。1 件のみ（2 行目を upsert）。
function _handleSaveAnnounce(ss, data) {
  const msg = String(data['訊息'] != null ? data['訊息'] : '');
  let sheet = ss.getSheetByName(ANNOUNCE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANNOUNCE_SHEET_NAME);
    sheet.getRange(1, 1, 1, ANNOUNCE_HEADERS.length).setValues([ANNOUNCE_HEADERS]).setFontWeight('bold').setBackground('#F7F4EC');
    sheet.setFrozenRows(1);
    Object.keys(ANNOUNCE_COL_WIDTHS).forEach(k => sheet.setColumnWidth(Number(k), ANNOUNCE_COL_WIDTHS[k]));
  }
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const ci = { msg: head.indexOf('訊息'), upd: head.indexOf('更新日') };
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  if (ci.msg >= 0) sheet.getRange(2, ci.msg + 1).setValue(msg);
  if (ci.upd >= 0) sheet.getRange(2, ci.upd + 1).setValue(now);
  SpreadsheetApp.flush();
  console.log('  → saveAnnounce ok, len', msg.length);
  return _writeJson({ ok: true, action: 'saveAnnounce', '更新': now });
}

function _handleAddIssue(ss, data) {
  const name = String(data['Issue'] || '').trim();
  if (!name) return _writeJson({ ok: false, error: 'issue_name_required' });
  const status = String(data['狀態'] || '').trim();
  if (status && !ALLOWED_STATUS.has(status)) return _writeJson({ ok: false, error: 'invalid_status', value: status });

  const sheet = ss.getSheetByName(ISSUE_SHEET_FOR_WRITE);
  if (!sheet) return _writeJson({ ok: false, error: 'issue_sheet_not_found', sheetName: ISSUE_SHEET_FOR_WRITE });

  // 自動採番（指定があれば尊重、無ければ連番 001…）
  let id = String(data['編號'] || '').trim();
  if (!id) id = _nextIssueNo(sheet);

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
    '擔當': String(data['擔當'] || ''),
    '戰略負責人': String(data['戰略負責人'] || ''),
    '協作': String(data['協作'] || ''),
    '狀態': status,
    '更新日': new Date(),
    'Confluence URL': String(data['Confluence URL'] || ''),
    'Issue定義': String(data['Issue定義'] || '')
  };
  const r = _appendByHeaders(sheet, vmap);
  SpreadsheetApp.flush();
  // 編號の先頭ゼロが数値化で消えないよう、書き込んだセルをテキスト書式で再設定
  if (numC >= 0 && r.row) sheet.getRange(r.row, numC + 1).setNumberFormat('@').setValue(id);
  // 初回 Progress があればログへ追記
  const prog = String(data['進度'] || data['備註'] || '').trim();
  if (prog) _appendLog(ss, id, String(data['author'] || data['擔當'] || ''), prog);
  console.log('  → addIssue ok:', id, 'row', r.row);
  return _writeJson({ ok: true, action: 'addIssue', '編號': id, mission: id, row: r.row, warnings: r.rejected });
}

// 次の Issue 連番（数字ゼロ詰め 3 桁）。数字のみの編號から最大＋1。
// 次の Issue 編號（I + 数字ゼロ詰め。例 I001）。数値化を避けるため接頭辞 I を付ける。
function _nextIssueNo(sheet) {
  const last = sheet.getLastRow();
  let max = 0;
  if (last >= 2) {
    const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const c = head.indexOf('編號');
    if (c >= 0) {
      sheet.getRange(2, c + 1, last - 1, 1).getValues().forEach(r => {
        const m = String(r[0]).trim().match(/(\d+)$/); // I001 / 001 / 戰略-1 いずれも末尾数字
        if (m) max = Math.max(max, parseInt(m[1], 10));
      });
    }
  }
  return 'I' + String(max + 1).padStart(3, '0');
}

// Mission進度ログへ1行追記（Issue/Mission/Task 共通。編號でぶら下げる）。
function _appendLog(ss, id, who, text) {
  text = String(text || '').trim();
  if (!text) return;
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold').setBackground('#F7F4EC');
    sh.setFrozenRows(1);
  }
  const tz = (typeof Session !== 'undefined' && Session.getScriptTimeZone()) ? Session.getScriptTimeZone() : 'Asia/Taipei';
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  sh.appendRow([String(id), now, String(who || ''), text]); // LOG_HEADERS: 編號/日時/擔當/進度
}

// ===== 連番への一括移行（一度だけ実行）=====
// Issue を 001,002… に振り直し、Mission(親-M{n})・Task(親-M{n}-T{k}) の編號/親編號、
// 進度ログ・個人備註の編號参照も連動更新する。全ダミー前提・非可逆なので注意。
function migrateNumbers() {
  const ss = _getSpreadsheet();
  if (!ss) { console.log('[migrate] no spreadsheet'); return; }
  const iSheet = ss.getSheetByName(ISSUE_SHEET_FOR_WRITE);
  const mSheet = ss.getSheetByName(MISSION_SHEET_FOR_WRITE);
  const tSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!iSheet) { console.log('[migrate] no Issue主檔'); return; }
  const colOf = (sh, name) => sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim()).indexOf(name);

  // --- Issues → I001.. ---
  const issueMap = {}, issueMapByNum = {};
  // 末尾数字で正規化（"001"/"1"/"039" のズレを吸収）
  const _numKey = s => { const m = String(s).trim().match(/(\d+)$/); return m ? String(parseInt(m[1], 10)) : ''; };
  const iNumC = colOf(iSheet, '編號'), iLast = iSheet.getLastRow();
  if (iNumC >= 0 && iLast >= 2) {
    const col = iSheet.getRange(2, iNumC + 1, iLast - 1, 1).getValues();
    let seq = 0;
    const out = col.map(r => {
      const old = String(r[0]).trim();
      if (!old) return [''];
      seq++; const nw = 'I' + String(seq).padStart(3, '0');
      issueMap[old] = nw; const nk = _numKey(old); if (nk) issueMapByNum[nk] = nw; // 接頭辞 I で数値化を防ぐ
      return [nw];
    });
    iSheet.getRange(2, iNumC + 1, out.length, 1).setNumberFormat("@").setValues(out);
  }
  const _issueNew = old => issueMap[old] || issueMapByNum[_numKey(old)] || old; // 親編號の正規化照合

  // --- Missions ---
  const missionMap = {};
  if (mSheet && mSheet.getLastRow() >= 2) {
    const mNumC = colOf(mSheet, '編號'), mParC = colOf(mSheet, '親編號'), mLast = mSheet.getLastRow();
    const parAF = mParC >= 0 && /^=\s*ARRAYFORMULA/i.test(String(mSheet.getRange(2, mParC + 1).getFormula() || ''));
    const nums = mSheet.getRange(2, mNumC + 1, mLast - 1, 1).getValues();
    const pars = mParC >= 0 ? mSheet.getRange(2, mParC + 1, mLast - 1, 1).getValues() : null;
    const outN = [], outP = [];
    for (let i = 0; i < nums.length; i++) {
      const oldM = String(nums[i][0]).trim();
      const oldP = pars ? String(pars[i][0]).trim() : '';
      const mm = oldM.match(/^(.*)-M(\d+)$/);
      let newM = oldM, newP = oldP;
      if (mm) { const nb = _issueNew(mm[1]); newM = nb + '-M' + mm[2]; newP = _issueNew(oldP) !== oldP ? _issueNew(oldP) : nb; }
      else { newP = _issueNew(oldP); }
      if (oldM) missionMap[oldM] = newM;
      outN.push([newM]); outP.push([newP]);
    }
    mSheet.getRange(2, mNumC + 1, outN.length, 1).setNumberFormat("@").setValues(outN);
    if (mParC >= 0 && !parAF) mSheet.getRange(2, mParC + 1, outP.length, 1).setNumberFormat("@").setValues(outP); // ARRAYFORMULA 列は編號から自動再計算に任せる
  }

  // --- Tasks ---
  const taskMap = {};
  if (tSheet && tSheet.getLastRow() >= 2) {
    const tNumC = colOf(tSheet, '編號'), tParC = colOf(tSheet, '親編號'), tLast = tSheet.getLastRow();
    const nums = tSheet.getRange(2, tNumC + 1, tLast - 1, 1).getValues();
    const pars = tParC >= 0 ? tSheet.getRange(2, tParC + 1, tLast - 1, 1).getValues() : null;
    const outN = [], outP = [];
    for (let i = 0; i < nums.length; i++) {
      const oldT = String(nums[i][0]).trim();
      const oldPM = pars ? String(pars[i][0]).trim() : '';
      const tm = oldT.match(/^(.*)-T(\d+)$/);
      let newT = oldT, newPM = oldPM;
      if (tm) { const nbm = missionMap[tm[1]] || tm[1]; newT = nbm + '-T' + tm[2]; newPM = missionMap[oldPM] || nbm; }
      else { newPM = missionMap[oldPM] || oldPM; }
      if (oldT) taskMap[oldT] = newT;
      outN.push([newT]); outP.push([newPM]);
    }
    tSheet.getRange(2, tNumC + 1, outN.length, 1).setNumberFormat("@").setValues(outN);
    if (tParC >= 0) tSheet.getRange(2, tParC + 1, outP.length, 1).setNumberFormat("@").setValues(outP);
  }

  // --- ログ・個人備註の編號参照 ---
  const all = Object.assign({}, issueMap, missionMap, taskMap);
  [LOG_SHEET_NAME, MEMO_SHEET_NAME].forEach(nm => {
    const sh = ss.getSheetByName(nm);
    if (!sh || sh.getLastRow() < 2) return;
    const c = colOf(sh, '編號'); if (c < 0) return;
    const col = sh.getRange(2, c + 1, sh.getLastRow() - 1, 1).getValues();
    const out = col.map(r => { const o = String(r[0]).trim(); return [all[o] || o]; });
    sh.getRange(2, c + 1, out.length, 1).setNumberFormat("@").setValues(out);
  });
  SpreadsheetApp.flush();
  console.log('[migrate] 完了：Issue ' + Object.keys(issueMap).length + ' / Mission ' + Object.keys(missionMap).length + ' / Task ' + Object.keys(taskMap).length);
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
      taskSheetName:   TASK_SHEET_NAME,
      taskSheetExists: !!(ss && ss.getSheetByName(TASK_SHEET_NAME)),  // false なら Task一覽 未作成
      statusValues:    STATUS_VALUES,  // 7状態（再公開確認用）
      personSheetName:   PERSON_SHEET_NAME,
      personSheetExists: !!(ss && ss.getSheetByName(PERSON_SHEET_NAME)),
      memoSheetName:     MEMO_SHEET_NAME,
      memoSheetExists:   !!(ss && ss.getSheetByName(MEMO_SHEET_NAME)),
      announceSheetExists: !!(ss && ss.getSheetByName(ANNOUNCE_SHEET_NAME)),
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

  // 既存 Mission一覽 に不足列（Mission定義＝カスケード健全性の言語化）を追記（非破壊）
  {
    const mh = readHeaders(taskSheet);
    ['Mission定義'].forEach(name => {
      if (!mh.includes(name)) {
        const col = taskSheet.getLastColumn() + 1;
        taskSheet.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#F7F4EC');
        log.push(`✓ Mission一覽 に列「${name}」を追加`);
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

  // --- 2.5 Task一覽 タブ（3層目の実体／2026-06-01 リデザイン）---
  let tkSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!tkSheet) {
    tkSheet = ss.insertSheet(TASK_SHEET_NAME);
    tkSheet.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS])
      .setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(TASK_COL_WIDTHS).forEach(k => tkSheet.setColumnWidth(Number(k), TASK_COL_WIDTHS[k]));
    tkSheet.setFrozenRows(1);
    log.push('✓ 「Task一覽」タブを作成（' + TASK_HEADERS.join(' / ') + '）');
  } else {
    // 既存タブに不足列を追記（非破壊）
    const th = readHeaders(tkSheet);
    TASK_HEADERS.forEach(name => {
      if (!th.includes(name)) {
        const col = tkSheet.getLastColumn() + 1;
        tkSheet.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#F7F4EC');
        log.push('✓ Task一覽 に列「' + name + '」を追加');
      }
    });
  }
  // Task一覽 の 狀態 プルダウン・色、更新日 の鮮度色（列名で特定）
  {
    const thd = readHeaders(tkSheet);
    const tkStatusCol = thd.indexOf('狀態') + 1;
    const tkUpdCol    = thd.indexOf('更新日') + 1;
    const tkRows      = Math.max(tkSheet.getMaxRows() - 1, 1000);
    if (tkStatusCol > 0) applyStatusValidation(tkSheet.getRange(2, tkStatusCol, tkRows, 1));
    const tkRules = tkSheet.getConditionalFormatRules();
    let tA = 0, tB = 0;
    if (tkStatusCol > 0) tA = addStatusColorRulesIfMissing(tkRules, tkSheet.getRange(2, tkStatusCol, tkRows, 1));
    if (tkUpdCol    > 0) tB = addFreshnessRulesIfMissing(tkRules, tkSheet.getRange(2, tkUpdCol, tkRows, 1), tkUpdCol);
    if (tA + tB > 0) {
      tkSheet.setConditionalFormatRules(tkRules);
      log.push(`✓ Task一覽 に条件付き書式 ${tA + tB} 件を追加`);
    }
  }

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

  // --- 5. 個人備註 タブ（事務局の私的メモ＋小沼の留言＋重要フラグ）---
  let memoSheet = ss.getSheetByName(MEMO_SHEET_NAME);
  if (!memoSheet) {
    memoSheet = ss.insertSheet(MEMO_SHEET_NAME);
    memoSheet.getRange(1, 1, 1, MEMO_HEADERS.length).setValues([MEMO_HEADERS])
      .setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(MEMO_COL_WIDTHS).forEach(k => memoSheet.setColumnWidth(Number(k), MEMO_COL_WIDTHS[k]));
    memoSheet.setFrozenRows(1);
    log.push('✓ 「個人備註」タブを作成（' + MEMO_HEADERS.join(' / ') + '）');
  } else {
    // 既存タブに不足列（留言 / 重要 等）を追記
    const mh = readHeaders(memoSheet);
    MEMO_HEADERS.forEach(name => {
      if (!mh.includes(name)) {
        const col = memoSheet.getLastColumn() + 1;
        memoSheet.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#F7F4EC');
        log.push('✓ 個人備註 に列「' + name + '」を追加');
      }
    });
  }

  // --- 6. 公告 タブ（お知らせボード）---
  let announceSheet = ss.getSheetByName(ANNOUNCE_SHEET_NAME);
  if (!announceSheet) {
    announceSheet = ss.insertSheet(ANNOUNCE_SHEET_NAME);
    announceSheet.getRange(1, 1, 1, ANNOUNCE_HEADERS.length).setValues([ANNOUNCE_HEADERS])
      .setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(ANNOUNCE_COL_WIDTHS).forEach(k => announceSheet.setColumnWidth(Number(k), ANNOUNCE_COL_WIDTHS[k]));
    announceSheet.setFrozenRows(1);
    log.push('✓ 「公告」タブを作成（訊息 / 更新日）。トップのお知らせボード');
  }

  const msg = log.length ? '✅ セットアップ完了\n\n' + log.join('\n')
                         : 'ℹ️ 既に整っています（追加・変更なし）';
  console.log('[setupJIG] ' + msg);
}

// ===== 条件付き書式をクリアして入れ直す（重複が気になったとき用）=====
function resetAndSetup() {
  const ss = _getSpreadsheet();
  ['Issue主檔', 'Mission一覽', TASK_SHEET_NAME].forEach(name => {
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
    .setAllowInvalid(true)  // 旧状態(未開始/需確認)が残る行を弾かない（移行期の互換）
    .setHelpText('構想中 / 策劃中 / 待審核 / 進行中 / 結案 / 凍結 / 中止 のいずれかを選択')
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

// =====================================================================
// ダミーデータ投入（UI 確認用・2026-06-01）
//   使い方：Apps Script エディタでこのファイルを保存 →（未実施なら）setupJIG を実行
//           → seedJIGDummy() を 1 回実行。Web App の再公開は不要（直接シートに書く）。
//           ダッシュボードは gviz で読むので、リロードすれば「我的」テーブルに出る。
//   取り消し：clearJIGDummy() を実行（ダミーの Mission/Task 行を名前一致で削除）。
//   親 Issue は Issue主檔 に実在する編號（戰略-1 等）に紐づける。
//   DRI/協作/戰略組擔當 は AIC4 組織ページの実在の名前を使用。
// =====================================================================
const DUMMY_MARK = '〔範例〕'; // Mission進度 / 進度 の先頭に付与（clear の目印）

// Issue が無い部門用の範例 Issue（無ければ seedJIGDummy が作成。事務局備註に DUMMY_MARK）
const DUMMY_ISSUES_FALLBACK = {
  '財務-1':   { name: '財務數據可視化與內控強化', dept: '財務',   group: '9.財務',     def: '統一前提的損益/預算/現金流數據架構，支援經營決策。' },
  '總管理-1': { name: '營運基盤與跨部門協調強化', dept: '總管理', group: '10.統籌整合', def: '建構共同節奏與文化定著，整合電商×CS×倉庫×配送端對端節奏。' },
};

// 1 組 = 1 Mission（lead=處長＝戰略組擔當・擔當=組代表）、組員 = 各 1 Task の DRI。全 51 名を網羅。
// n=姓名 / t=Task / s=狀態 / a=更新日(何日前) / p=進度 / c=協作
const DUMMY_GROUPS = [
  { issue: '戰略-1', mission: '經營戰略組：BCL/LC8T 推進', lead: '小沼和幸', status: '進行中', age: 1,
    def: '彙整全公司 KPI，推動 BCL 新客招募與 LC8T 行銷自動化，並協調跨部門策略。',
    members: [
      { n:'小沼和幸', t:'TBBS 事業結構改革推進', s:'進行中', a:1, p:'年度方針與資源分配' },
      { n:'王詩雅', t:'事業 KPI 儀表板與異常檢測', s:'進行中', a:3, p:'財務儀表板建置中', c:'傅貞甄' },
      { n:'張育菱', t:'BCL 低成本獲取×LC8T 自動化', s:'進行中', a:1, p:'名單品質提升中' },
      { n:'陳少琪', t:'LBCL 活動與 LINE OA 維運', s:'策劃中', a:4, p:'每週接觸節奏設計' },
      { n:'阿比留華', t:'LC8T 數據洞察與 LTV 假設', s:'構想中', a:8, p:'資料導出顧客需求' },
      { n:'竹下友梨', t:'戰略議題轉事業企劃', s:'待審核', a:6, p:'論點×假設草案待審' },
    ] },
  { issue: '學習-1', mission: '學習商品開發組：商品開發與品質', lead: '江美齡', status: '進行中', age: 4,
    def: '學習套組開發與品質管理，經營數位學習網與 YouTube，推進 IP 異業合作。',
    members: [
      { n:'陳瓊芳', t:'新數位平台×實體最小範疇定義', s:'策劃中', a:5, p:'連動核心體驗規格' },
      { n:'于安平', t:'教材企劃與教具選品判準', s:'進行中', a:2, p:'跨媒材製作流程管控' },
      { n:'陳勝朋', t:'AI 影音流程與 2027 新站準備', s:'進行中', a:3, p:'自動字幕剪輯測試', c:'曹舒涵' },
      { n:'陳乃菁', t:'套組多媒體設計對齊教育價值', s:'待審核', a:8, p:'初稿待處長確認' },
    ] },
  { issue: '授權-1', mission: '品牌授權策略組：IP 價值最大化', lead: '江美齡', status: '進行中', age: 6,
    def: '維護品牌形象並極大化 IP 價值，拓展商品/空間/活動授權業務。',
    members: [
      { n:'謝惠琪', t:'新規開發清單與提案管線', s:'策劃中', a:6, p:'優先順位方法建構' },
      { n:'劉玉珊', t:'夢想樂園第二店舗選址談判', s:'進行中', a:2, p:'合作條件評估中' },
      { n:'陳伊柔', t:'招商×企劃價值說明框架', s:'構想中', a:12, p:'可複用提案套件' },
      { n:'陳怡如', t:'二次使用授權流程標準化', s:'進行中', a:4, p:'平台成果追蹤節奏' },
    ] },
  { issue: '舞台劇-1', mission: '表演活動組：舞台劇事業成長', lead: '江美齡', status: '構想中', age: 16,
    def: '統籌舞台劇製作與行銷，建立會員回流機制與現場商品營運。',
    members: [
      { n:'沈美君', t:'舞台劇會員制度設計', s:'構想中', a:16, p:'LINE 分眾年 2〜3 場' },
      { n:'藍靜儀', t:'現場商品開發與選品打法', s:'進行中', a:3, p:'小批量也成立' },
      { n:'林欣亭', t:'廣宣發稿節奏優化', s:'凍結', a:30, p:'暫緩至下季' },
    ] },
  { issue: '整合-2', mission: '數位行銷組：AWRT 與廣告優化', lead: '劉靜芸', status: '進行中', age: 2,
    def: '數位通路全方位經營，整合社群/KOL，運用 AI 與數據優化投放與轉換。',
    members: [
      { n:'陳筱昀', t:'全通路×預算配分計畫', s:'進行中', a:2, p:'官網/蝦皮/SNS 整合' },
      { n:'鍾明雯', t:'GA4 分析與 KOL 素材流程', s:'進行中', a:3, p:'學習價值敘事架構', c:'白如雪' },
      { n:'陳思嘉', t:'廣告素材 A/B 與蝦皮合作', s:'策劃中', a:5, p:'多切角素材開發', c:'戴詠' },
      { n:'曹舒涵', t:'LINE 貼圖變現×Threads 經營', s:'構想中', a:10, p:'導流變現流程' },
    ] },
  { issue: '整合-5', mission: '通路整合組：官方通路 LTV 經營', lead: '劉靜芸', status: '進行中', age: 3,
    def: '官方自營通路營運與內容策劃，優化體驗路徑提升回購與忠誠。',
    members: [
      { n:'周宜柔', t:'會員轉換 LTV 引擎×Mirafeel', s:'進行中', a:2, p:'年度目標 166 萬' },
      { n:'郭怡梅', t:'Super8 自動旅程×會員制度', s:'策劃中', a:5, p:'DB 成長與 LTV' },
      { n:'張菀庭', t:'Mirafeel 尿布行銷與成效追蹤', s:'進行中', a:3, p:'會員分級回購旅程' },
      { n:'鍾佳臻', t:'顧客成長模型與支付評估', s:'構想中', a:11, p:'回購/升級/跨店' },
      { n:'李佳霖', t:'跨平台導流×跨境商品', s:'進行中', a:4, p:'社群×APP×LINE 串聯' },
    ] },
  { issue: 'CX-1', mission: 'LTV 戰略組：ACTS 自動化', lead: '鮑慧芬', status: '進行中', age: 1,
    def: '制定 LTV 策略，運用 CRM 標籤與 ibo 自動化精準分眾並產出高潛名單。',
    members: [
      { n:'賴純美', t:'ibo 自動化×高潛名單產出', s:'進行中', a:2, p:'目標去電≥300/月', c:'郭姿伶' },
      { n:'郭姿伶', t:'0-7 歲分齡推播策略', s:'待審核', a:7, p:'分齡內容推播設計' },
    ] },
  { issue: 'CX-2', mission: 'CS 戰略組：電話銷售與顧客關係', lead: '鮑慧芬', status: '進行中', age: 2,
    def: '第一線顧客溝通與高潛名單外撥銷售，提升成交率並回饋名單品質。',
    members: [
      { n:'鄭心怡', t:'可預測銷售管理系統', s:'進行中', a:2, p:'目標→每日行動→成交' },
      { n:'張秀美', t:'電話銷售與帳款回收', s:'進行中', a:3, p:'訂單處理中' },
      { n:'傅暄',   t:'訂單處理與行政作業', s:'策劃中', a:5, p:'流程整理中' },
      { n:'黃秋如', t:'售後諮詢與教材推薦', s:'進行中', a:4, p:'主動挖掘需求' },
      { n:'嚴雅怡', t:'客訴處理與情緒安撫', s:'進行中', a:3, p:'疑難排解' },
      { n:'莊雅涵', t:'Shopee 聊聊客服×售後', s:'構想中', a:9, p:'配送/資料修改' },
      { n:'張金葵', t:'客戶服務與商品推廣', s:'進行中', a:6, p:'' },
      { n:'黃鈺茹', t:'電銷與售後滿意度', s:'進行中', a:2, p:'抱怨處理' },
      { n:'許世芬', t:'Email 諮詢回覆×推廣', s:'策劃中', a:8, p:'信箱諮詢' },
    ] },
  { issue: '財務-1', mission: '財務組：可視化與內控', lead: '中智玄', status: '進行中', age: 2,
    def: '營運資金管理、預算差異分析與帳務合規，編製管理報表。',
    members: [
      { n:'傅貞甄', t:'月結→預算差異→儀表板一條龍', s:'進行中', a:2, p:'損益即時可視化' },
      { n:'許珮珊', t:'多平台銷售報表標準化', s:'進行中', a:4, p:'收益調整表流程' },
    ] },
  { issue: '資訊-1', mission: '資訊組：系統整合與資安', lead: '白如雪', status: '策劃中', age: 3,
    def: '系統管理維護與資安監控，端對端整合與數據流通能力建構。',
    members: [
      { n:'吳家慶', t:'SQL 效能優化與備份', s:'進行中', a:3, p:'索引調校中' },
      { n:'戴詠',   t:'端對端整合 API/Webhook 標準', s:'進行中', a:2, p:'連接規格草擬', c:'陳思嘉' },
      { n:'吳昌儒', t:'IT Help Desk 與資安', s:'構想中', a:11, p:'新人上手' },
    ] },
  { issue: '總管理-1', mission: '統籌整合組：營運基盤', lead: '范巧惠', status: '進行中', age: 3,
    def: '人事/總務/法務/物流日常運作，導入 AI 提升行政效率。',
    members: [
      { n:'鄧宏毅', t:'一頁談判摘要×合約標準化', s:'進行中', a:3, p:'爭點/風險格式化' },
      { n:'蔡品媛', t:'AIC 學習路徑×福委會互動', s:'策劃中', a:6, p:'角色情境分流' },
      { n:'林昀萱', t:'入庫數據防呆×倉儲協調', s:'進行中', a:4, p:'核檢機制建立' },
    ] },
];

function _dummyDate(age) {
  const d = new Date(Date.now() - (age || 0) * 86400000);
  const p = x => String(x).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

// 高速一括追加：行（{ヘッダ名:値} の配列）を列ごとに 1 回ずつ setValues で書く。
// ARRAYFORMULA 列（親編號 等の自動算出）はスキップ。flush は最後に 1 回だけ（タイムアウト回避）。
function _bulkAppend(sheet, rows) {
  if (!rows || !rows.length) return;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const start = sheet.getLastRow() + 1;
  const n = rows.length;
  // ARRAYFORMULA 列の検出（2 行目の数式を見る）。これらは編號から自動算出されるので書かない。
  const isArrayF = {};
  if (sheet.getLastRow() >= 2) {
    const f = sheet.getRange(2, 1, 1, lastCol).getFormulas()[0];
    for (let c = 0; c < lastCol; c++) isArrayF[c] = /^=\s*ARRAYFORMULA/i.test(String(f[c] || ''));
  }
  headers.forEach((h, c) => {
    if (!h || isArrayF[c]) return;
    let any = false;
    const colVals = rows.map(r => {
      const v = Object.prototype.hasOwnProperty.call(r, h) ? r[h] : '';
      if (v !== '' && v != null) any = true;
      return [v];
    });
    if (!any && h !== '編號') return; // 全空列はスキップ（編號は必ず書く）
    sheet.getRange(start, c + 1, n, 1).setValues(colVals);
  });
  try { SpreadsheetApp.flush(); } catch (e) { console.log('[bulkAppend] flush warn:', String(e && e.message || e)); }
}

function seedJIGDummy() {
  const ss = _getSpreadsheet();
  if (!ss) { console.log('[seed] SHEET_ID 未設定'); return; }
  const mSheet = ss.getSheetByName(MISSION_SHEET_FOR_WRITE);
  if (!mSheet) { console.log('[seed] Mission一覽 が無い。先に setupJIG を実行'); return; }
  const issueSheet = ss.getSheetByName(ISSUE_SHEET_FOR_WRITE);
  if (!issueSheet) { console.log('[seed] Issue主檔 が無い'); return; }
  let tSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!tSheet) {
    tSheet = ss.insertSheet(TASK_SHEET_NAME);
    tSheet.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS]).setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(TASK_COL_WIDTHS).forEach(k => tSheet.setColumnWidth(Number(k), TASK_COL_WIDTHS[k]));
    tSheet.setFrozenRows(1);
  }
  // 既存 Issue 編號
  const issueIds = new Set();
  if (issueSheet.getLastRow() >= 2) {
    const ih = issueSheet.getRange(1, 1, 1, issueSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const ic = ih.indexOf('編號');
    if (ic >= 0) issueSheet.getRange(2, ic + 1, issueSheet.getLastRow() - 1, 1).getValues().forEach(r => issueIds.add(String(r[0]).trim()));
  }
  // Issue が無い部門は範例 Issue を作成
  const issueRows = [];
  Object.keys(DUMMY_ISSUES_FALLBACK).forEach(id => {
    if (issueIds.has(id)) return;
    const fb = DUMMY_ISSUES_FALLBACK[id];
    issueRows.push({ '編號': id, 'Issue': fb.name, '處': fb.dept, '組': fb.group, '狀態': '進行中',
      '事務局備註': DUMMY_MARK + '範例 Issue', '更新日': _dummyDate(3), 'Issue定義': fb.def || '' });
    issueIds.add(id);
  });
  if (issueRows.length) _bulkAppend(issueSheet, issueRows);

  // 既存 Mission から parent ごとの次 -M 番号（メモリ内で採番）
  const next = {};
  if (mSheet.getLastRow() >= 2) {
    mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 1).getValues().forEach(r => {
      const m = String(r[0]).trim().match(/^(.*)-M(\d+)$/);
      if (m) next[m[1]] = Math.max(next[m[1]] || 0, +m[2]);
    });
  }
  const missionRows = [], taskRows = [], skipped = [];
  DUMMY_GROUPS.forEach(g => {
    if (!issueIds.has(g.issue)) { skipped.push(g.issue); return; }
    next[g.issue] = (next[g.issue] || 0) + 1;
    const missionId = g.issue + '-M' + next[g.issue];
    const owner = (g.members[0] && g.members[0].n) || '';
    missionRows.push({ '編號': missionId, 'Mission': g.mission, '親編號': g.issue,
      '戰略負責人': g.lead || '', '擔當': owner, '狀態': g.status || '進行中',
      'Mission進度': DUMMY_MARK + (g.def ? g.def.slice(0, 30) : ''), '更新日': _dummyDate(g.age),
      'Confluence URL': '', 'Mission定義': g.def || '' });
    (g.members || []).forEach((mem, k) => {
      taskRows.push({ '編號': missionId + '-T' + (k + 1), 'Task': mem.t || (mem.n + ' 的任務'), '親編號': missionId,
        'DRI': mem.n || '', '協作': mem.c || '', '狀態': mem.s || '構想中',
        '進度': DUMMY_MARK + (mem.p || ''), '更新日': _dummyDate(mem.a), '連結': '' });
    });
  });
  _bulkAppend(mSheet, missionRows);
  _bulkAppend(tSheet, taskRows);
  console.log(`[seed] 完了：範例Issue ${issueRows.length} / Mission ${missionRows.length} / Task ${taskRows.length}` + (skipped.length ? ` skip親不在:${skipped.join(',')}` : ''));
}

function clearJIGDummy() {
  const ss = _getSpreadsheet();
  if (!ss) return;
  let removed = 0;
  // [シート名, 目印列]：その列が DUMMY_MARK で始まる行を削除
  [[MISSION_SHEET_FOR_WRITE, 'Mission進度'], [TASK_SHEET_NAME, '進度'], [ISSUE_SHEET_FOR_WRITE, '事務局備註']].forEach(([name, col]) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const ci = headers.indexOf(col);
    if (ci < 0) return;
    const vals = sh.getRange(2, ci + 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) { // 下から削除（行ズレ防止）
      if (String(vals[i][0] || '').indexOf(DUMMY_MARK) === 0) { sh.deleteRow(i + 2); removed++; }
    }
  });
  SpreadsheetApp.flush();
  console.log(`[clearDummy] ${removed} 行を削除（〔範例〕で始まる行）`);
}

// =====================================================================
// 全 Issue に対し 2 Mission × 2 Task を生成（定義つき・2026-06-02）
//   各 Issue に Issue定義、各 Mission に Mission定義、各 Task に Task定義 を入れる。
//   定義は Issue 名を織り込んだカスケード整合のテンプレ（判斷基準／完成判斷／完成＝）。
//   使い方：setupJIG → seedFullDummy → migrateNumbers（→ リロード）。
//   既存の〔範例〕Mission/Task は実行時に一旦掃除してから作り直す（重複防止）。
// =====================================================================
const FD_POOL = ['王詩雅','張育菱','陳少琪','陳瓊芳','于安平','陳勝朋','陳乃菁','謝惠琪','劉玉珊','陳伊柔',
  '沈美君','藍靜儀','林欣亭','陳筱昀','鍾明雯','陳思嘉','曹舒涵','周宜柔','郭怡梅','張菀庭',
  '賴純美','郭姿伶','鄭心怡','傅貞甄','許珮珊','吳家慶','戴詠','鄧宏毅','蔡品媛','林昀萱'];

function _fdClearDummyMT(ss) {
  let removed = 0;
  [[MISSION_SHEET_FOR_WRITE, 'Mission進度'], [TASK_SHEET_NAME, '進度']].forEach(([name, col]) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const ci = headers.indexOf(col);
    if (ci < 0) return;
    const vals = sh.getRange(2, ci + 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0] || '').indexOf(DUMMY_MARK) === 0) { sh.deleteRow(i + 2); removed++; }
    }
  });
  return removed;
}

function seedFullDummy() {
  const ss = _getSpreadsheet();
  if (!ss) { console.log('[full] SHEET_ID 未設定'); return; }
  const iSheet = ss.getSheetByName(ISSUE_SHEET_FOR_WRITE);
  const mSheet = ss.getSheetByName(MISSION_SHEET_FOR_WRITE);
  if (!iSheet || !mSheet) { console.log('[full] Issue主檔/Mission一覽 が無い。先に setupJIG'); return; }
  let tSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!tSheet) {
    tSheet = ss.insertSheet(TASK_SHEET_NAME);
    tSheet.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS]).setFontWeight('bold').setBackground('#F7F4EC');
    Object.keys(TASK_COL_WIDTHS).forEach(k => tSheet.setColumnWidth(Number(k), TASK_COL_WIDTHS[k]));
    tSheet.setFrozenRows(1);
  }
  // Mission一覽／Task一覽 を全リセット（ヘッダ1行を残してデータ行を全削除）。
  // ※手編集で〔範例〕が消えた行も含めて綺麗にし、編號重複を根絶する。全ダミー前提。
  let cleared = 0;
  [mSheet, tSheet].forEach(sh => {
    const lr = sh.getLastRow();
    if (lr > 1) { sh.deleteRows(2, lr - 1); cleared += (lr - 1); }
  });
  SpreadsheetApp.flush();

  // 全 Issue を読む
  const ih = iSheet.getRange(1, 1, 1, iSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const ci = { num: ih.indexOf('編號'), name: ih.indexOf('Issue'), lead: ih.indexOf('戰略負責人'),
    dept: ih.indexOf('處'), def: ih.indexOf('Issue定義') };
  const iLast = iSheet.getLastRow();
  if (ci.num < 0 || iLast < 2) { console.log('[full] Issue 行なし'); return; }
  const irows = iSheet.getRange(2, 1, iLast - 1, iSheet.getLastColumn()).getValues();

  let np = 0, si = 0;
  const pick = () => FD_POOL[np++ % FD_POOL.length];
  const nextStatus = () => STATUS_VALUES[si++ % STATUS_VALUES.length];
  const ages = [1, 3, 6, 10, 16, 30];
  let ai = 0;
  const nextAge = () => ages[ai++ % ages.length];

  const issueDefOut = []; // Issue定義 列の上書き
  const missionRows = [], taskRows = [];
  let mCount = 0, tCount = 0, iDefCount = 0;

  irows.forEach(r => {
    const issueId = String(r[ci.num]).trim();
    const issueName = ci.name >= 0 ? String(r[ci.name]).trim() : '';
    if (!issueId || !issueName) { issueDefOut.push([ci.def >= 0 ? String(r[ci.def] || '') : '']); return; }
    const lead = ci.lead >= 0 ? String(r[ci.lead] || '').trim() : '';
    // Issue定義（空なら埋める。既にあれば尊重）
    let idef = ci.def >= 0 ? String(r[ci.def] || '').trim() : '';
    if (!idef) { idef = `圍繞「${issueName}」建立可衡量的成果與判斷基準，並確保下層 Mission／Task 對齊落實。判斷基準：關鍵指標達標、推進節奏穩定、可追蹤。`; iDefCount++; }
    issueDefOut.push([idef]);

    // 2 Mission × 2 Task
    const missions = [
      { suf: 'M1', name: `${issueName}－策略與計畫`,
        def: `在「${issueName}」之下，盤點現況、定義策略與計畫。完成判斷：策略文件與里程碑定稿、資源與 KPI 明確。`,
        tasks: [
          { name: '現況盤點與假設', def: `盤點現況數據與限制，提出關鍵假設與優先順位。完成＝盤點報告與假設清單交付。` },
          { name: '方案設計與 KPI', def: `設計可執行方案並定義 KPI 與量測方式。完成＝方案與 KPI 定稿、可進入執行。` },
        ] },
      { suf: 'M2', name: `${issueName}－執行與優化`,
        def: `在「${issueName}」之下，落地執行並依數據持續優化。完成判斷：上線運作、KPI 達標、形成可複用節奏。`,
        tasks: [
          { name: '落地執行', def: `依方案推動落地，建立執行節奏與分工。完成＝主要工作上線運作。` },
          { name: '成效追蹤與改善', def: `追蹤 KPI、檢視成效並提出改善。完成＝成效報告與下一輪改善建議。` },
        ] },
    ];
    missions.forEach(m => {
      const missionId = issueId + '-' + m.suf;
      const owner = pick();
      missionRows.push({
        '編號': missionId, 'Mission': m.name, '親編號': issueId,
        '戰略負責人': lead, '擔當': owner, '狀態': nextStatus(),
        'Mission進度': DUMMY_MARK + '推進中', '更新日': _dummyDate(nextAge()),
        'Confluence URL': '', 'Mission定義': m.def
      });
      mCount++;
      m.tasks.forEach((t, k) => {
        const dri = pick();
        const collab = (k === 0) ? pick() : '';
        taskRows.push({
          '編號': missionId + '-T' + (k + 1), 'Task': t.name, '親編號': missionId,
          'DRI': dri, '協作': collab, '狀態': nextStatus(),
          '進度': DUMMY_MARK + '作業中', '更新日': _dummyDate(nextAge()), '連結': '',
          'Task定義': t.def
        });
        tCount++;
      });
    });
  });

  // Issue定義 列を書き戻し（無ければ列が無い→スキップ）
  if (ci.def >= 0) iSheet.getRange(2, ci.def + 1, issueDefOut.length, 1).setValues(issueDefOut);
  _bulkAppend(mSheet, missionRows);
  _bulkAppend(tSheet, taskRows);
  SpreadsheetApp.flush();
  console.log(`[full] 完了：掃除 ${cleared} 行 → Issue定義 ${iDefCount} 件補完 / Mission ${mCount} / Task ${tCount}`);
}
