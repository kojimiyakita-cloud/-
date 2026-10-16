# -*- coding: utf-8 -*-
"""
2026年度 中途人事会議用 資料　実績最新化スクリプト
=====================================================
実行すると:
  1. 令和８年度 応募者名簿の最新データをキャッシュシートに同期
     （古いキャッシュシートは削除してから新規作成）
  2. 自動計算元シートの P1 を本日の日付に更新
  3. 自動計算元シートの数式を正しい形式に自動修正
     - 採用数/入社数: 「下見営業」と「業務」のみカウント（空欄はカウントしない）
     - 入社取消は実際に入社済みのため、採用数・入社数の双方に含める

→ Excel で開けば P1 に連動する D2/J2/P2/V2 が自動計算され、
  全 COUNTIFS 集計が本日時点の最新実績に更新されます。

実行方法:
  python 今日付けで実績更新.py
  または「今日付けで実績更新.bat」をダブルクリック
"""
from datetime import datetime
from pathlib import Path

import openpyxl

TARGET = Path("c:/Users/honbu19/Desktop/中途分析資料/2026年度　中途人事会議用 資料.xlsx")
SRC    = Path("c:/Users/honbu19/Desktop/中途分析資料/令和８年度　応募者名簿.xlsx")

# 応募者名簿のシート名
S_EIGYO    = "8年度　営業"
S_EVENT    = "8年度　イベント"
S_UKETSUKE = "8年度　受付"

# 資料ファイル内のキャッシュシート名
CACHE_EIGYO    = "_R8営業DATA"
CACHE_EVENT    = "_R8イベントDATA"
CACHE_UKETSUKE = "_R8受付DATA"

# ====================================================================
# 数式定義: 入社数は「下見営業」と「業務」の合計、入社取消も含める
# ====================================================================

# 媒体別 入社数の正しい数式（R列=下見営業 または 業務をカウント）
FORMULA_R5 = (
    '=COUNTIFS(_R8営業DATA!$C$2:$C$5000,"掲載型",_R8営業DATA!$R$2:$R$5000,"下見営業",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    '+COUNTIFS(_R8営業DATA!$C$2:$C$5000,"掲載型",_R8営業DATA!$R$2:$R$5000,"業務",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
)

FORMULA_R16 = (
    '=COUNTIFS(_R8営業DATA!$C$2:$C$5000,"紹介型",_R8営業DATA!$R$2:$R$5000,"下見営業",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    '+COUNTIFS(_R8営業DATA!$C$2:$C$5000,"紹介型",_R8営業DATA!$R$2:$R$5000,"業務",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
)

FORMULA_R27 = (
    '=COUNTIFS(_R8営業DATA!$C$2:$C$5000,"転職イベント",_R8営業DATA!$R$2:$R$5000,"下見営業",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    '+COUNTIFS(_R8営業DATA!$C$2:$C$5000,"転職イベント",_R8営業DATA!$R$2:$R$5000,"業務",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
)

# ====================================================================
# 媒体別 採用数の正しい数式（R列=下見営業 または 業務 をカウント、入社取消も採用に含む）
# ====================================================================

FORMULA_P5 = (
    '=COUNTIFS(_R8営業DATA!$C$2:$C$5000,"掲載型",_R8営業DATA!$R$2:$R$5000,"下見営業",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    '+COUNTIFS(_R8営業DATA!$C$2:$C$5000,"掲載型",_R8営業DATA!$R$2:$R$5000,"業務",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
)

FORMULA_P16 = (
    '=COUNTIFS(_R8営業DATA!$C$2:$C$5000,"紹介型",_R8営業DATA!$R$2:$R$5000,"下見営業",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    '+COUNTIFS(_R8営業DATA!$C$2:$C$5000,"紹介型",_R8営業DATA!$R$2:$R$5000,"業務",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
)

FORMULA_P27 = (
    '=COUNTIFS(_R8営業DATA!$C$2:$C$5000,"転職イベント",_R8営業DATA!$R$2:$R$5000,"下見営業",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    '+COUNTIFS(_R8営業DATA!$C$2:$C$5000,"転職イベント",_R8営業DATA!$R$2:$R$5000,"業務",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
)

# ====================================================================
# 支店別 入社数の数式（R49～R62、R67～R80）
# M列=支店名、R列=下見営業 または 業務、Q列=採用、日付範囲
# ====================================================================

# 支店名リスト（R67～R80の順序）
SHITEN_LIST = [
    "名古屋",    # R67
    "名古屋北",  # R68
    "名古屋東",  # R69
    "静岡",      # R70
    "横浜",      # R71
    "神奈川",    # R72
    "千葉",      # R73
    "市川",      # R74
    "埼玉",      # R75
    "関東統括",  # R76
    "仙台",      # R77
    "郡山",      # R78
    "大阪",      # R79
    "南大阪",    # R80
]

def make_shiten_formula(shiten_name: str) -> str:
    """令和8年度の支店別入社数を生成（下見営業+業務、入社取消を含む）"""
    return (
        f'=COUNTIFS(_R8営業DATA!$M$2:$M$5000,"{shiten_name}",_R8営業DATA!$R$2:$R$5000,"下見営業",'
        '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
        '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
        f'+COUNTIFS(_R8営業DATA!$M$2:$M$5000,"{shiten_name}",_R8営業DATA!$R$2:$R$5000,"業務",'
        '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
        '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    )

def make_previous_shiten_formula(row: int) -> str:
    """令和7年度の支店別入社数を生成（R37と同じ条件、入社取消を含む）"""
    return (
        f'=COUNTIFS(_R7営業DATA!$M$2:$M$5000,$A{row},'
        '_R7営業DATA!$R$2:$R$5000,"下見営業",'
        '_R7営業DATA!$Q$2:$Q$5000,"採用",'
        '_R7営業DATA!$K$2:$K$5000,">="&$D$2,'
        '_R7営業DATA!$K$2:$K$5000,"<="&$J$2)'
    )


# R81: 全体の入社数合計（支店関係なく、下見営業+業務、入社取消を含む）
FORMULA_R81 = (
    '=COUNTIFS(_R8営業DATA!$R$2:$R$5000,"下見営業",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
    '+COUNTIFS(_R8営業DATA!$R$2:$R$5000,"業務",'
    '_R8営業DATA!$Q$2:$Q$5000,"採用",_R8営業DATA!$B$2:$B$5000,">="&$P$2,'
    '_R8営業DATA!$B$2:$B$5000,"<="&$V$2)'
)

# 数式の対応表 (セルアドレス: 正しい数式)
FORMULAS_TO_FIX = {
    "P5":  FORMULA_P5,
    "P16": FORMULA_P16,
    "P27": FORMULA_P27,
    "R5":  FORMULA_R5,
    "R16": FORMULA_R16,
    "R27": FORMULA_R27,
    "R81": FORMULA_R81,
}

# 支店別の数式を追加（R67～R80）
for i, shiten in enumerate(SHITEN_LIST):
    cell_addr = f"R{67 + i}"
    FORMULAS_TO_FIX[cell_addr] = make_shiten_formula(shiten)

# 前年度の支店別数式も追加（R49～R62）
for row in range(49, 63):
    FORMULAS_TO_FIX[f"R{row}"] = make_previous_shiten_formula(row)


def recreate_cache_sheet(wb, name: str):
    """古いキャッシュシートを削除し、空の非表示シートを新規作成して返す。

    上書き同期だと空欄セルに古い値が残ることがあるため、
    毎回削除して作り直す。
    """
    if name in wb.sheetnames:
        del wb[name]
        print(f"  {name}: 古いキャッシュを削除")
    ws = wb.create_sheet(name)
    ws.sheet_state = "hidden"
    return ws


def sync_sheet(src_ws, dst_ws) -> int:
    """ソースシートの全データを新規キャッシュシートにコピー"""
    max_col = max(src_ws.max_column or 1, 30)
    src_max_row = src_ws.max_row

    for row_idx, row_data in enumerate(
        src_ws.iter_rows(
            min_row=1, max_row=src_max_row,
            min_col=1, max_col=max_col,
            values_only=True,
        ),
        start=1,
    ):
        for col_idx in range(1, max_col + 1):
            val = row_data[col_idx - 1] if col_idx - 1 < len(row_data) else None
            dst_ws.cell(row=row_idx, column=col_idx, value=val)

    return src_max_row - 1  # ヘッダー除くデータ行数


def fix_nyusha_torikeshi(ws_cache, cache_name: str) -> int:
    """
    入社取消の人のQ列（採否）を「採用」に自動修正
    
    入社取消 = 一度採用されて入社したが、その後取り消しになった
    → 採用数・入社数の双方に含める
    
    対象列:
    - Q列(17列目): 採否
    - W列(23列目): 備考（ここに「入社取消」が入る）
    
    Returns:
        修正した行数
    """
    COL_Q = 17  # 採否
    COL_W = 23  # 備考
    
    fixed_count = 0
    max_row = ws_cache.max_row
    
    for row in range(2, max_row + 1):  # ヘッダー(1行目)をスキップ
        w_val = ws_cache.cell(row=row, column=COL_W).value
        q_val = ws_cache.cell(row=row, column=COL_Q).value
        
        # W列が「入社取消」の場合、Q列を「採用」に設定
        if w_val and "入社取消" in str(w_val):
            if q_val != "採用":
                ws_cache.cell(row=row, column=COL_Q).value = "採用"
                fixed_count += 1
                # 氏名を取得して表示（E列=5列目）
                name = ws_cache.cell(row=row, column=5).value or f"行{row}"
                print(f"    {cache_name} 行{row} ({name}): Q列を「採用」に修正（入社取消のため）")
    
    return fixed_count


def is_file_locked(path: Path) -> bool:
    """ファイルが（Excel などで）開かれていて書き込みできない状態か判定"""
    if not path.exists():
        return False
    # Excel が開いていると同じフォルダに一時ロックファイル「~$xxxx.xlsx」が出来る
    lock_marker = path.with_name("~$" + path.name)
    if lock_marker.exists():
        return True
    # 実際に書き込みモードで開けるか試す（開ければロックされていない）
    try:
        with open(path, "r+b"):
            return False
    except (PermissionError, OSError):
        return True


def wait_until_writable(path: Path) -> None:
    """対象ファイルが書き込み可能になるまで、Excel を閉じるよう促して待つ"""
    while is_file_locked(path):
        print()
        print("[!] このファイルは現在 Excel で開かれているため保存できません:")
        print(f"    {path.name}")
        print("    → Excel で開いている場合は必ず閉じてください。")
        ans = input("    閉じたら Enter を押して続行（中止する場合は q + Enter）: ")
        if ans.strip().lower() == "q":
            raise SystemExit("処理を中止しました。")


def save_with_retry(wb, path: Path) -> None:
    """保存に失敗（ファイルが開かれている等）した場合はリトライを促す"""
    while True:
        try:
            wb.save(path)
            return
        except PermissionError:
            print()
            print("[!] 保存に失敗しました。ファイルが Excel で開かれている可能性があります:")
            print(f"    {path.name}")
            print("    → Excel を閉じてから再試行してください。")
            ans = input("    閉じたら Enter を押して再試行（中止する場合は q + Enter）: ")
            if ans.strip().lower() == "q":
                raise SystemExit("保存を中止しました。変更は保存されていません。")


def fix_formulas(ws_auto) -> int:
    """
    自動計算元シートの数式を正しい形式に修正
    
    入社数の数式（R5/R16/R27など）：
    - R列（職種）で「下見営業」と「業務」の両方をカウント
    - W列（備考）が「入社取消」の場合もカウント
    
    Returns:
        修正したセルの数
    """
    fixed_count = 0
    
    for cell_addr, correct_formula in FORMULAS_TO_FIX.items():
        cell = ws_auto[cell_addr]
        current = cell.value
        
        # 正規化して比較（スペースや大文字小文字の違いを無視）
        current_normalized = str(current).replace(" ", "").lower() if current else ""
        correct_normalized = correct_formula.replace(" ", "").lower()
        
        if current_normalized != correct_normalized:
            cell.value = correct_formula
            fixed_count += 1
            print(f"    {cell_addr}: 数式を修正しました")
    
    return fixed_count


def main():
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    today_str = today.strftime("%Y/%m/%d")
    print(f"本日の日付: {today_str}")
    print()

    # ---- ファイル読み込み ----
    if not TARGET.exists():
        print(f"[エラー] ファイルが見つかりません: {TARGET}")
        input("Enterキーを押して終了...")
        return
    if not SRC.exists():
        print(f"[エラー] 応募者名簿が見つかりません: {SRC}")
        input("Enterキーを押して終了...")
        return

    # ---- 事前チェック: 対象ファイルが Excel で開かれていないか ----
    if is_file_locked(TARGET):
        print("[注意] 保存先ファイルが開かれています。先に Excel を閉じてください。")
        wait_until_writable(TARGET)

    print(f"読み込み中: {TARGET.name}")
    wb = openpyxl.load_workbook(TARGET)

    print(f"読み込み中: {SRC.name}")
    src = openpyxl.load_workbook(SRC, data_only=True, read_only=True)

    # ---- キャッシュシートを最新データで更新（古いキャッシュは削除して作り直し）----
    print()
    print("【データ同期】応募者名簿 → キャッシュシート（古いキャッシュは削除）")
    synced_any = False
    synced_caches = []  # 同期したキャッシュシートを記録
    for cache_name, sheet_name in [
        (CACHE_EIGYO,    S_EIGYO),
        (CACHE_EVENT,    S_EVENT),
        (CACHE_UKETSUKE, S_UKETSUKE),
    ]:
        if sheet_name in src.sheetnames:
            ws_cache = recreate_cache_sheet(wb, cache_name)
            n = sync_sheet(src[sheet_name], ws_cache)
            print(f"  {cache_name}: {n:,} 行を新規作成")
            synced_any = True
            synced_caches.append((cache_name, ws_cache))
        else:
            print(f"  スキップ: シート '{sheet_name}' が応募者名簿にありません")

    src.close()

    if not synced_any:
        print("[警告] 同期できたシートがありませんでした。")

    # ---- 入社取消の人のQ列を「採用」に自動修正 ----
    print()
    print("【データ補正】入社取消者の採否を自動修正")
    print("  ※ W列(備考)が「入社取消」の場合、Q列(採否)を「採用」に設定")
    print("  ※ 入社取消 = 採用数・入社数の双方にカウント")
    total_fixed = 0
    for cache_name, ws_cache in synced_caches:
        fixed = fix_nyusha_torikeshi(ws_cache, cache_name)
        total_fixed += fixed
    if total_fixed == 0:
        print("    修正対象なし")
    else:
        print(f"    計 {total_fixed} 件を修正しました")

    # ---- P1 を本日の日付に更新 ----
    print()
    print("【日付更新】自動計算元シート P1")
    ws_auto = wb["自動計算元"]

    old_p1 = ws_auto["P1"].value
    old_str = old_p1.strftime("%Y/%m/%d") if isinstance(old_p1, datetime) else str(old_p1)
    ws_auto["P1"] = today
    ws_auto["P1"].number_format = "yyyy/m/d"
    print(f"  P1: {old_str}  →  {today_str}")

    # ---- 数式を正しい形式に修正 ----
    print()
    print("【数式確認・修正】入社数の計算式")
    print("  ※ R列(職種)が「下見営業」または「業務」をカウント")
    print("  ※ W列(備考)が「入社取消」の場合も入社数に含める")
    fixed = fix_formulas(ws_auto)
    if fixed == 0:
        print("    すべての数式は正しい状態です")
    else:
        print(f"    {fixed} 件の数式を修正しました")

    # ---- 保存 ----
    print()
    print(f"保存中: {TARGET.name}")
    save_with_retry(wb, TARGET)

    print()
    print("=" * 50)
    print("完了！")
    print(f"  {TARGET.name} を保存しました。")
    print()
    print("Excel で開くと P1 の日付に連動して")
    print("  D2 / J2 / P2 / V2 が自動計算され、")
    print("  全 COUNTIFS 集計が最新実績に更新されます。")
    print("=" * 50)
    print()
    input("Enterキーを押して終了...")


if __name__ == "__main__":
    main()
