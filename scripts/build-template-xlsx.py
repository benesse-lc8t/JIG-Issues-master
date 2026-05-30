#!/usr/bin/env python3
"""
TBBS-JIG スプレッドシートの一括セットアップ（Excel 経由・v0.7.x Mission 対応）

Usage:
    python3 scripts/build-template-xlsx.py [input.xlsx] [output.xlsx]

  - input.xlsx：現状の Google Sheet を「ファイル → ダウンロード → .xlsx」した結果
  - output.xlsx：このスクリプトが整えた xlsx。事務局が Google Sheets に
    「ファイル → インポート → 既存のシートを置き換え」でアップロード。

何をするか:
  1. Issue主檔 シートに、不足列（Confluence URL / 狀態 / 事務局備註 / 更新日）を追加
  2. Issue主檔 内に種別=Task / 種別=Mission の行があれば抽出して Mission一覽 へ移行
     - Issue主檔 から「種別」「親編號」列を撤去
  3. 旧シート名 Task一覽 があれば Mission一覽 にリネーム＋中身も移行
  4. Mission一覽 シートを整える（ヘッダ 8 列、列名 'Mission'、編號 *-M#）
  5. 両シートの狀態列にプルダウン（未開始/進行中/完成）と3色条件付き書式
  6. 両シートの更新日列に鮮度色変化（>7日黄、>14日赤）
  7. 列幅・ヘッダ行 freeze

データ検証と条件付き書式は xlsx ネイティブの形式で保存するため、
Google Sheets にインポートしても継承される。
"""
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

INPUT  = Path(sys.argv[1] if len(sys.argv) > 1 else 'Issue主檔.xlsx')
OUTPUT = Path(sys.argv[2] if len(sys.argv) > 2 else 'JIG-template.xlsx')

# 出力 Mission シート名（旧名 Task一覽 は読み込み時に互換）
MISSION_SHEET_NAME = 'Mission一覽'
LEGACY_TASK_SHEET  = 'Task一覽'

NEW_ISSUE_COLS    = ['Confluence URL', '狀態', '事務局備註', '更新日']
STATUS_VALUES     = ['未開始', '進行中', '完成']
STATUS_COLORS     = {'未開始': 'EEEAE0', '進行中': 'DCEBFB', '完成': 'D4F2DD'}
MISSION_HEADERS   = ['編號', 'Mission', '親編號', '戰略負責人', '狀態', '事務局備註', '更新日', 'Confluence URL']
MISSION_WIDTHS    = [14, 38, 12, 12, 10, 28, 12, 30]
HDR_FILL = PatternFill(start_color='F7F4EC', end_color='F7F4EC', fill_type='solid')
HDR_FONT = Font(bold=True)


def apply_status_dropdown(ws, col_letter, row_count=500):
    dv = DataValidation(
        type='list',
        formula1='"{}"'.format(','.join(STATUS_VALUES)),
        allow_blank=True,
        showErrorMessage=True,
    )
    dv.error = '未開始 / 進行中 / 完成 のいずれかを選択してください'
    dv.errorTitle = '無効な値'
    ws.add_data_validation(dv)
    dv.add('{c}2:{c}{n}'.format(c=col_letter, n=row_count))


def apply_status_colors(ws, col_letter, row_count=500):
    rng = '{c}2:{c}{n}'.format(c=col_letter, n=row_count)
    for v, color in STATUS_COLORS.items():
        rule = CellIsRule(
            operator='equal',
            formula=['"{}"'.format(v)],
            fill=PatternFill(start_color=color, end_color=color, fill_type='solid'),
        )
        ws.conditional_formatting.add(rng, rule)


def apply_freshness_colors(ws, col_letter, row_count=500):
    rng = '{c}2:{c}{n}'.format(c=col_letter, n=row_count)
    warn = FormulaRule(
        formula=['AND(${c}2<>"", TODAY()-${c}2>7, TODAY()-${c}2<=14)'.format(c=col_letter)],
        fill=PatternFill(start_color='FFF3CC', end_color='FFF3CC', fill_type='solid'),
        font=Font(color='7A5A00'),
    )
    stale = FormulaRule(
        formula=['AND(${c}2<>"", TODAY()-${c}2>14)'.format(c=col_letter)],
        fill=PatternFill(start_color='FCD7D7', end_color='FCD7D7', fill_type='solid'),
        font=Font(color='8A0000'),
    )
    ws.conditional_formatting.add(rng, warn)
    ws.conditional_formatting.add(rng, stale)


def header_idx(ws, name):
    for cell in ws[1]:
        if cell.value == name:
            return cell.column
    return None


def collect_legacy_missions_from_sheet(ws):
    """既存 Mission一覽 or Task一覽 シートから行を読み取って dict のリストで返す。"""
    if ws is None or ws.max_row < 2:
        return []
    h_map = {}
    for cell in ws[1]:
        if cell.value:
            h_map[str(cell.value).strip()] = cell.column
    out = []
    for row_num in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=row_num, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v is None or (isinstance(v, str) and v == '') for v in row_vals):
            continue

        def g(name):
            c = h_map.get(name)
            return ws.cell(row=row_num, column=c).value if c else None

        # Mission 内容列：Mission / Task / Issue の順で探す
        content = g('Mission') or g('Task') or g('Issue')
        out.append({
            '編號':              g('編號'),
            'Mission':           content,
            '親編號':            g('親編號'),
            '戰略負責人':        g('戰略負責人'),
            '狀態':              g('狀態'),
            '事務局備註':        g('事務局備註'),
            '更新日':            g('更新日'),
            'Confluence URL':    g('Confluence URL'),
        })
    return out


def renumber_mission(parent_num, counter_per_parent):
    if not parent_num:
        return ''
    counter_per_parent[parent_num] = counter_per_parent.get(parent_num, 0) + 1
    return f'{parent_num}-M{counter_per_parent[parent_num]}'


# ---- 処理開始 ----
if not INPUT.exists():
    print(f'入力ファイルが見つかりません: {INPUT}', file=sys.stderr)
    sys.exit(1)

wb = openpyxl.load_workbook(INPUT)

# 1) Issue主檔 に不足列を追加
if 'Issue主檔' not in wb.sheetnames:
    print('「Issue主檔」シートがありません', file=sys.stderr)
    sys.exit(1)
issue_ws = wb['Issue主檔']
existing = [c.value for c in issue_ws[1]]
for name in NEW_ISSUE_COLS:
    if name not in existing:
        col = issue_ws.max_column + 1
        cell = issue_ws.cell(row=1, column=col, value=name)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        existing.append(name)
        print(f'  + Issue主檔 に列「{name}」を追加')

# 2) Issue主檔 内の Task / Mission 行を抜き出して Mission一覽 へ移行する準備
type_col_idx   = header_idx(issue_ws, '種別')
parent_col_idx = header_idx(issue_ws, '親編號')
migrated = []  # list of dict (Mission 用フォーマットに揃える)

if type_col_idx is not None:
    rows_to_delete = []
    for row_num in range(2, issue_ws.max_row + 1):
        type_val = str(issue_ws.cell(row=row_num, column=type_col_idx).value or '').strip()
        if type_val in ('Task', 'Mission'):
            def get_main(col_name, _rn=row_num):
                idx = header_idx(issue_ws, col_name)
                return issue_ws.cell(row=_rn, column=idx).value if idx else None
            migrated.append({
                '編號':           None,  # 後で再採番
                'Mission':        get_main('Issue'),
                '親編號':         get_main('親編號'),
                '戰略負責人':     get_main('戰略負責人'),
                '狀態':           get_main('狀態'),
                '事務局備註':     get_main('事務局備註'),
                '更新日':         get_main('更新日'),
                'Confluence URL': get_main('Confluence URL'),
            })
            rows_to_delete.append(row_num)
    for rn in reversed(rows_to_delete):
        issue_ws.delete_rows(rn)
    if migrated:
        print(f'  - Issue主檔 から 種別=Task/Mission 行 {len(migrated)} 件を抽出（Mission一覽 へ移行）')

# 3) 旧 Task一覽 or 既存 Mission一覽 から既存データを吸い上げる
for legacy_name in (MISSION_SHEET_NAME, LEGACY_TASK_SHEET):
    if legacy_name in wb.sheetnames:
        prior = collect_legacy_missions_from_sheet(wb[legacy_name])
        if prior:
            print(f'  - 既存「{legacy_name}」から {len(prior)} 件を引き継ぎ')
            migrated.extend(prior)
        del wb[legacy_name]

# 4) Issue主檔 から「種別」「親編號」列を撤去（Mission が別シートに分離されたので不要）
for col_name in ['種別', '親編號']:
    idx = header_idx(issue_ws, col_name)
    if idx:
        issue_ws.delete_cols(idx)
        print(f'  - Issue主檔 から列「{col_name}」を撤去')

# 5) Issue主檔 のフォーマット（列削除後にもう一度 idx を取り直す）
st_idx  = header_idx(issue_ws, '狀態')
upd_idx = header_idx(issue_ws, '更新日')
if st_idx:
    apply_status_dropdown(issue_ws, get_column_letter(st_idx), row_count=500)
    apply_status_colors(issue_ws, get_column_letter(st_idx), row_count=500)
if upd_idx:
    apply_freshness_colors(issue_ws, get_column_letter(upd_idx), row_count=500)
issue_ws.freeze_panes = 'A2'

# 6) Mission一覽 シートを作成
mission_ws = wb.create_sheet(MISSION_SHEET_NAME)
for i, h in enumerate(MISSION_HEADERS, start=1):
    cell = mission_ws.cell(row=1, column=i, value=h)
    cell.font = HDR_FONT
    cell.fill = HDR_FILL
for i, w in enumerate(MISSION_WIDTHS, start=1):
    mission_ws.column_dimensions[get_column_letter(i)].width = w

# 7) Mission 行を 編號 *-M# で書き出し
counter = {}
out_row = 2
for m in migrated:
    parent_num = str(m.get('親編號') or '').strip()
    # 既存の 編號 が *-M# 形式ならそのまま使う、そうでなければ再採番
    cur_num = str(m.get('編號') or '').strip()
    if cur_num and '-M' in cur_num and cur_num.split('-M')[-1].isdigit():
        new_num = cur_num
        # カウンタも追随させて後続の採番がぶつからないように
        suffix = int(cur_num.split('-M')[-1])
        existing_pfx = cur_num.rsplit('-M', 1)[0]
        if existing_pfx == parent_num:
            counter[parent_num] = max(counter.get(parent_num, 0), suffix)
    else:
        new_num = renumber_mission(parent_num, counter)

    mission_ws.cell(row=out_row, column=1, value=new_num)
    mission_ws.cell(row=out_row, column=2, value=m.get('Mission'))
    mission_ws.cell(row=out_row, column=3, value=parent_num)
    mission_ws.cell(row=out_row, column=4, value=m.get('戰略負責人'))
    mission_ws.cell(row=out_row, column=5, value=m.get('狀態'))
    mission_ws.cell(row=out_row, column=6, value=m.get('事務局備註'))
    mission_ws.cell(row=out_row, column=7, value=m.get('更新日'))
    # Confluence URL：ダッシュボード側で親 Issue から継承するため空欄が既定。
    # 値が数式（=...）なら旧行を参照していて壊れているので捨てる。
    cu = m.get('Confluence URL')
    if isinstance(cu, str) and cu.startswith('='):
        cu = None
    mission_ws.cell(row=out_row, column=8, value=cu)
    out_row += 1

apply_status_dropdown(mission_ws, 'E', row_count=500)
apply_status_colors(mission_ws, 'E', row_count=500)
apply_freshness_colors(mission_ws, 'G', row_count=500)
mission_ws.freeze_panes = 'A2'
print(f'  + Mission一覽 シートを作成（8列）／Mission {len(migrated)} 件を投入（編號は *-M# 形式）')

wb.save(OUTPUT)
print(f'\n✅ 出力完了: {OUTPUT}')
print('   Google Sheets で「ファイル → インポート → 既存のシートを置き換え」でアップロード')
