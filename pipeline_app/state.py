# -*- coding: utf-8 -*-
"""処理の進み具合の判定と記録。

各ステップは、前のステップが作ったファイルを受け取って加工する。想定と違う段階のファイルに
実行すると、エラーにならないまま誤った出力ができてしまう。それを防ぐために次の 2 つを行う。

1. ファイルの見出し行から「いまどの段階か」を判定し、実行できないステップを止める。
2. 作業フォルダの pipeline_state.json に、見出しからは分からない情報
   (商品名の標準化を実行したか、どのファイルを取り込んだか) を記録する。

このモジュールは pandas を読み込まない (画面の実行前チェックから使うため)。
"""

import json
import os
from pathlib import Path

DATA_SHEET = "売上データ"
WORK_FILE = "sales_work.xlsx"          # STEP 0〜5 が順に加工していく作業ファイル
OUTPUT_FILE = "sales_for_bi.xlsx"      # STEP 6〜7 が作る、BI 取り込み用の完成ファイル
STATE_FILE = "pipeline_state.json"

# ---- 列の定義 (出力の形はここで一元管理する) ----
# 仕入先ごとの変換部品が返す「標準形式」。STEP 0 直後の作業ファイルはこの 10 列。
MERGED_COLUMNS = ["元ファイル", "計上月", "卸名", "得意先名称", "商品名称", "容量", "入数", "販売本数", "住所", "業態"]
ACCOUNT_COLUMNS = ["企業名", "担当部署", "担当者"]        # STEP 2 が追加
PRICE_COLUMNS = ["上代", "仕切"]                           # STEP 3 が追加
# STEP 4 (最終調整) のあとの 17 列
FINAL_COLUMNS = [
    "卸名", "計上月", "担当部署", "担当者", "得意先名称", "企業名",
    "商品名称", "容量", "入数", "上代", "販売本数", "仕切",
    "売上金額", "年度", "商品区分", "住所", "業態",
]
GEO_COLUMNS = ["都道府県", "都内区分"]                     # STEP 5 が「住所」の右に追加
ID_COLUMNS = ["商品ID", "都道府県ID", "エリアID", "卸ID", "担当部署ID", "担当者ID"]   # STEP 6 が追加
GEO_LAYOUT = FINAL_COLUMNS[:16] + GEO_COLUMNS + ["業態"]                   # STEP 5 のあと (19 列)
OUTPUT_COLUMNS = FINAL_COLUMNS[:16] + GEO_COLUMNS + ID_COLUMNS + ["業態"]   # 完成ファイル (25 列)

STEP_NAMES = {
    0: "データ統合", 1: "商品名の標準化", 2: "顧客情報の付与", 3: "価格の付与", 4: "最終調整",
    5: "都道府県の抽出", 6: "ID の付与", 7: "テーブル設定", 8: "後処理",
}

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


# ---------------------------------------------------------------- 見出し行の読み取り
def _column_index(ref, fallback):
    """セル番地 ("C1" など) から 0 始まりの列番号を求める。"""
    letters = "".join(ch for ch in (ref or "") if ch.isalpha())
    if not letters:
        return fallback
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _shared_string(si):
    """共有文字列 1 件の本文 (ふりがなは含めない)。"""
    parts = [t.text or "" for t in si.findall(_NS + "t")]
    parts += [t.text or "" for r in si.findall(_NS + "r") for t in r.findall(_NS + "t")]
    return "".join(parts)


def _read_header_fast(path, sheet):
    """xlsx の中身を直接たどって 1 行目だけを読む (行数の多いファイルでも一瞬で済むように)。"""
    import xml.etree.ElementTree as ET
    import zipfile

    with zipfile.ZipFile(path) as z:
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rel_id = next((s.get(_RNS + "id") for s in workbook.iter(_NS + "sheet") if s.get("name") == sheet), None)
        if rel_id is None:
            return None
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next((r.get("Target") for r in rels if r.get("Id") == rel_id), None)
        if not target:
            raise ValueError("sheet part not found")
        part = target.lstrip("/") if target.startswith("/") else "xl/" + target

        cells = []
        with z.open(part) as f:
            for _event, element in ET.iterparse(f, events=("end",)):
                if element.tag == _NS + "row":
                    for position, cell in enumerate(element.findall(_NS + "c")):
                        kind = cell.get("t")
                        value = cell.find(_NS + "v")
                        inline = "".join(t.text or "" for t in cell.iter(_NS + "t")) if kind == "inlineStr" else None
                        cells.append((_column_index(cell.get("r"), position), kind,
                                      value.text if value is not None else None, inline))
                    break
        wanted = {int(v) for _, kind, v, _ in cells if kind == "s" and v is not None}
        strings = {}
        if wanted:
            last, index = max(wanted), 0
            with z.open("xl/sharedStrings.xml") as f:
                for _event, element in ET.iterparse(f, events=("end",)):
                    if element.tag == _NS + "si":
                        if index in wanted:
                            strings[index] = _shared_string(element)
                        index += 1
                        element.clear()
                        if index > last:
                            break
    header = []
    for column, kind, value, inline in sorted(cells, key=lambda c: c[0]):
        while len(header) < column:
            header.append(None)
        if kind == "s":
            header.append(strings.get(int(value)) if value is not None else None)
        elif kind == "inlineStr":
            header.append(inline)
        else:
            header.append(value)
    while header and header[-1] is None:
        header.pop()
    return header


def read_header(path, sheet=DATA_SHEET):
    """ブックの 1 行目を読む。読めなければ None。"""
    try:
        return _read_header_fast(path, sheet)
    except Exception:
        return _read_header_openpyxl(path, sheet)      # 想定外の形のファイルは openpyxl に任せる


def _read_header_openpyxl(path, sheet):
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception:
        return None
    try:
        if sheet not in wb.sheetnames:
            return None
        for row in wb[sheet].iter_rows(min_row=1, max_row=1, values_only=True):
            header = list(row)
            while header and header[-1] is None:
                header.pop()
            return header
        return []
    finally:
        wb.close()


# ---------------------------------------------------------------- 段階の判定
def work_stage(header):
    """作業ファイルの見出しから段階を判定する。

    戻り値: {"layout": "merged" | "adjusted" | "geo" | "unknown", "accounts": bool, "prices": bool}
      merged   … STEP 0 の出力の形。STEP 1〜3 を実行できる (STEP 2・3 が済むと列が増える)。
      adjusted … STEP 4 の出力の形 (17 列)。STEP 5 を実行できる。
      geo      … STEP 5 の出力の形 (19 列)。STEP 6 を実行できる。
    """
    stage = {"layout": "unknown", "accounts": False, "prices": False}
    if header is None:
        return stage
    if header == FINAL_COLUMNS:
        stage["layout"] = "adjusted"
    elif header == GEO_LAYOUT:
        stage["layout"] = "geo"
    elif header[:len(MERGED_COLUMNS)] == MERGED_COLUMNS:
        rest = header[len(MERGED_COLUMNS):]
        for accounts in (False, True):
            for prices in (False, True):
                if rest == (ACCOUNT_COLUMNS if accounts else []) + (PRICE_COLUMNS if prices else []):
                    stage.update(layout="merged", accounts=accounts, prices=prices)
    return stage


def output_ok(header):
    """完成ファイルの売上データが STEP 6 の出力の形か。"""
    return header == OUTPUT_COLUMNS


def actual_state(work_dir):
    """作業フォルダの現在の状態を調べる。"""
    work = Path(work_dir)
    saved = load_state(work)
    state = {"work": None, "output": None, "step1": None}
    path = work / WORK_FILE
    if path.exists():
        state["work"] = work_stage(read_header(path))
        record = saved.get("work") or {}
        if record.get("signature") == signature(path) and record.get("from_step0"):
            state["step1"] = 1 in record.get("steps", [])      # STEP 0 からの履歴が分かるときだけ確定
    path = work / OUTPUT_FILE
    if path.exists():
        state["output"] = output_ok(read_header(path))
    return state


def problem(step_id, state):
    """そのステップを実行できない理由 (実行できるなら None)。"""
    work = state.get("work")
    if step_id in (1, 2, 3, 4, 5, 6):
        if work is None:
            return f"作業フォルダに {WORK_FILE} がありません。先に STEP 0 を実行してください。"
        if work["layout"] == "unknown":
            return (f"{WORK_FILE} の列の構成が想定と違うため、STEP {step_id} を実行できません。"
                    "STEP 0 から実行し直してください。")
    if step_id in (1, 2, 3, 4):
        if work["layout"] != "merged":
            return (f"{WORK_FILE} は、すでに STEP 4 (最終調整) まで済んでいます。"
                    f"STEP {step_id} をやり直すには、STEP 0 から実行してください。")
        if step_id in (3, 4) and state.get("step1") is False:
            return "STEP 1 (商品名の標準化) がまだ実行されていません。先に STEP 1 を実行してください。"
        if step_id == 4:
            missing = [f"STEP {n} ({STEP_NAMES[n]})" for n, done in ((2, work["accounts"]), (3, work["prices"])) if not done]
            if missing:
                return "STEP 4 の前に " + " と ".join(missing) + " を実行してください。"
    elif step_id == 5:
        if work["layout"] == "merged":
            return f"{WORK_FILE} は、まだ STEP 4 (最終調整) が済んでいません。先に STEP 4 まで実行してください。"
    elif step_id == 6:
        if work["layout"] != "geo":
            return f"{WORK_FILE} は、まだ STEP 5 (都道府県の抽出) が済んでいません。先に STEP 5 まで実行してください。"
    elif step_id == 7:
        if state.get("output") is None:
            return f"作業フォルダに {OUTPUT_FILE} がありません。先に STEP 6 を実行してください。"
        if not state["output"]:
            return f"{OUTPUT_FILE} の列の構成が想定と違うため、STEP 7 を実行できません。STEP 6 から実行し直してください。"
    return None


def advance(step_id, state):
    """そのステップが成功したあとの状態 (実行前チェックで先の段階を見通すため)。"""
    state = dict(state)
    work = dict(state["work"]) if state.get("work") else None
    if step_id == 0:
        work = {"layout": "merged", "accounts": False, "prices": False}
        state["step1"] = False
    elif step_id == 1:
        state["step1"] = True
        if work:
            work["prices"] = False        # 商品名が変わるので、価格は付け直しになる
    elif step_id == 2 and work:
        work["accounts"] = True
    elif step_id == 3 and work:
        work["prices"] = True
    elif step_id == 4:
        work = {"layout": "adjusted", "accounts": False, "prices": False}
    elif step_id == 5:
        work = {"layout": "geo", "accounts": False, "prices": False}
    elif step_id == 6:
        state["output"] = True
    state["work"] = work
    return state


# ---------------------------------------------------------------- 記録ファイル
def signature(path):
    try:
        st = os.stat(path)
    except OSError:
        return None
    return [st.st_size, st.st_mtime_ns]


def load_state(work_dir):
    try:
        data = json.loads((Path(work_dir) / STATE_FILE).read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_state(work_dir, data):
    try:
        (Path(work_dir) / STATE_FILE).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass          # 記録できなくても処理は続ける (判定が「不明」になるだけ)


def clear_state(work_dir):
    try:
        (Path(work_dir) / STATE_FILE).unlink()
    except OSError:
        pass


def record_step(work_dir, step_id, signature_before=None, merged_files=None):
    """STEP 0〜5 が成功したあとに呼ぶ。作業ファイルにどのステップを実行したかを記録する。

    signature_before: ステップ開始前の作業ファイルの signature()。記録と同じファイルに続けて
    実行したのかを確かめるために使う (別のファイルに差し替えられていたら履歴は不明とする)。
    """
    work = Path(work_dir)
    data = load_state(work)
    if step_id == 0:
        data = {"version": 1, "work": {"steps": [0], "from_step0": True}, "merged_files": merged_files or {}}
    else:
        record = data.get("work") or {}
        known = record.get("signature") is not None and record.get("signature") == signature_before
        steps = list(record.get("steps", [])) if known else []
        data["work"] = {"steps": steps + [step_id], "from_step0": bool(known and record.get("from_step0"))}
    data["work"]["signature"] = signature(work / WORK_FILE)
    save_state(work, data)


def merged_files(work_dir):
    """直近の STEP 0 で取り込んだファイル名 {仕入先名: [ファイル名]}。記録が無ければ None。"""
    data = load_state(work_dir)
    files = data.get("merged_files")
    return files if isinstance(files, dict) and data.get("work") else None
