# -*- coding: utf-8 -*-
"""各処理ステップの本体。

どの関数も「ファイルのパスを受け取り、結果を辞書で返す」形にそろえてある。画面には依存せず、
進み具合は print で出力する (画面側が標準出力を実行ログ欄に流す)。

共通の約束ごと:
- 列は位置ではなく見出しの名前で扱う。
- 照合に使うキー (商品名・得意先名など) は normalize_key() で表記ゆれをならしてから比べる。
- 保存は一時ファイルに書き出してから置き換える (途中で失敗しても元のファイルを壊さない)。
- 利用者に確かめてほしい点は、黙って捨てたり補ったりせず、結果の辞書に入れて返す。
"""

import csv
import importlib.util
import math
import os
import re
import shutil
import time
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from . import state
from .plugin import OPTIONAL_COLUMNS as PLUGIN_OPTIONAL      # 変換部品が返してよい列 (無ければ空として補う)
from .plugin import REQUIRED_COLUMNS as PLUGIN_REQUIRED      # 変換部品が必ず返す列
from .state import (ACCOUNT_COLUMNS, DATA_SHEET, FINAL_COLUMNS, GEO_LAYOUT, MERGED_COLUMNS, OUTPUT_COLUMNS,
                    OUTPUT_FILE, PRICE_COLUMNS, WORK_FILE)

UNMATCHED_PRODUCTS = "unmatched_products.xlsx"      # 商品マスタに無かった商品名の一覧
UNMATCHED_CUSTOMERS = "unmatched_customers.xlsx"    # 顧客マスタに無かった得意先の一覧

# ---- 業務ルール (組織に合わせて変える場所) ----
FISCAL_YEAR_START_MONTH = 1          # 年度の開始月。1 なら暦年、4 なら 4 月始まり
DISCONTINUED_MARK = "終売"           # 商品名にこの文字があれば「終売商品」とする
PRODUCT_CLASS_NORMAL, PRODUCT_CLASS_DISCONTINUED = "通常商品", "終売商品"
BLANK_COMPANY = "(空白)"             # 企業名が分からない行に入れる文字 (BI で空欄の項目にならないように)
KEY_REGION = "東京都"                # この都道府県だけ、さらに細かい区分 (都内区分) を付ける


class StepError(Exception):
    """処理を続行できないエラー (メッセージをそのまま利用者に見せる)。"""


# =============================================================================
# 共通の道具
# =============================================================================
def is_blank(value):
    """None / NaN / 空白文字のみ なら True。"""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def normalize_key(value):
    """照合用のキーを作る: 全角/半角の違いをなくし、前後の空白を除き、連続する空白を 1 つにする。

    元のデータは書き換えない。マスタとデータの両方をこの関数に通してから比べる。
    """
    if is_blank(value):
        return ""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(value))).strip()


def as_number(value):
    """数値として読めれば数値 (整数になるものは int)、読めなければ None。"1,200" や全角数字も読む。"""
    if isinstance(value, bool) or is_blank(value):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = unicodedata.normalize("NFKC", str(value)).strip().replace(",", "")
        try:
            number = float(text)
        except ValueError:
            return None
    if not math.isfinite(number):
        return None
    return int(number) if number == int(number) else number


_CAPACITY_RE = re.compile(r"^(\d+(?:\.\d+)?)(ml|cc|l|ミリリットル|リットル)?$")


def parse_capacity(value):
    """容量を ml 単位の数値にする。"720ml"→720、"1.8L"→1800、"1L"→1000、単位なしはそのまま。"""
    number = as_number(value)
    if number is not None:
        return number
    text = unicodedata.normalize("NFKC", str(value)).lower().replace(",", "").replace(" ", "")
    match = _CAPACITY_RE.match(text)
    if not match:
        return None
    number = float(match.group(1)) * (1000 if match.group(2) in ("l", "リットル") else 1)
    return int(number) if number == int(number) else round(number, 3)


def fiscal_year(month):
    """計上月 (YYYYMM の整数) から年度を求める。"""
    year, mm = divmod(int(month), 100)
    return year if mm >= FISCAL_YEAR_START_MONTH else year - 1


def _tidy(value):
    """セル 1 つを保存に適した形にする (空は None、文字は前後の空白を除く、12.0 は 12)。"""
    if is_blank(value):
        return None
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return value
    if hasattr(value, "item") and not isinstance(value, (int, float)):       # numpy の数値
        try:
            value = value.item()
        except (AttributeError, ValueError):
            return value
    if isinstance(value, float):
        return int(value) if value == int(value) else float(value)
    return value


def _column(values, index):
    """Python の値をそのまま持つ列 (pandas に勝手に小数へ変えさせない)。"""
    return pd.Series(list(values), index=index, dtype=object)


def _sample(values, limit=3):
    values = [str(v) for v in values]
    return "、".join(values[:limit]) + (f" ほか{len(values) - limit}件" if len(values) > limit else "")


def _load_module(path, name):
    import sys
    from .config import APP_DIR
    if str(APP_DIR) not in sys.path:          # 変換部品から pipeline_app.plugin を読み込めるようにする
        sys.path.insert(0, str(APP_DIR))
    spec = importlib.util.spec_from_file_location(name, str(path))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# ---- 読み込み ----
def _read_table(path, sheet, label):
    """シートを表として読む。値の型は変えない (数値は数値、文字は文字、空のセルは NaN)。"""
    try:
        df = pd.read_excel(path, sheet_name=sheet, dtype=object, keep_default_na=False, na_values=[""])
    except ValueError as e:
        if "Worksheet" in str(e):
            raise StepError(f"{label} ({os.path.basename(str(path))}) に「{sheet}」シートがありません。") from e
        raise
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _require_columns(df, columns, label, sheet):
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise StepError(f"{label} の「{sheet}」シートに必要な列がありません: " + "、".join(missing))


def _read_work(path):
    return _read_table(path, DATA_SHEET, WORK_FILE)


def _check_ready(step_id, work_path=None, output_path=None):
    """対象ファイルが、そのステップを実行できる段階かを見出し行で確かめる。"""
    current = {"work": None, "output": None, "step1": None}
    if work_path is not None and os.path.exists(work_path):
        current["work"] = state.work_stage(state.read_header(work_path))
    if output_path is not None and os.path.exists(output_path):
        current["output"] = state.output_ok(state.read_header(output_path))
    message = state.problem(step_id, current)
    if message:
        raise StepError(message)


# ---- 保存 ----
def _temp_path(path):
    """保存先と同じフォルダに作る一時ファイルの名前。"""
    p = Path(path)
    return str(p.with_name(p.stem + ".saving" + p.suffix))


def _remove_quietly(path):
    try:
        os.remove(path)
    except OSError:
        pass


def _replace_with_retry(tmp, path, retries=3):
    """一時ファイルを本来の名前に置き換える。Excel で開かれている場合に備えてやり直す。"""
    for attempt in range(retries):
        try:
            os.replace(tmp, path)
            print(f"保存しました: {path}")
            return
        except PermissionError as e:
            if attempt < retries - 1:
                print(f"保存できませんでした (ファイルが使用中の可能性があります: {os.path.basename(str(path))})。"
                      f"やり直します...({attempt + 1}/{retries})")
                time.sleep(2)
            else:
                _remove_quietly(tmp)
                raise StepError(f"ファイルを保存できません。Excel で開いていないか確認してください:\n{path}") from e


def _save_with_retry(save, path, retries=3):
    """一時ファイルに書き出してから置き換える。save は保存先のパスを受け取って書き出す関数。"""
    path = str(path)
    tmp = _temp_path(path)
    try:
        save(tmp)
    except BaseException:
        _remove_quietly(tmp)
        raise
    _replace_with_retry(tmp, path, retries)


def _save_sheets(sheets, path):
    """{シート名: DataFrame} を 1 つのブックとして保存する。"""
    def save(tmp):
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            for name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=name, index=False)

    _save_with_retry(save, path)


def _save_work(df, path):
    _save_sheets({DATA_SHEET: df}, path)


def _save_list(rows, columns, path):
    """確認用の一覧を保存する。1 件も無ければ、前回の一覧が残らないようにファイルを消す。"""
    if not rows:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
        except OSError as e:
            raise StepError(f"前回の一覧を消せません。Excel で開いていないか確認してください:\n{path}") from e
        return
    _save_sheets({"一覧": pd.DataFrame(rows, columns=columns)}, path)


def _unmatched_rows(df, column, keys, is_unmatched):
    """照合できなかった値を [値, 行数, 卸名] の一覧にまとめる。"""
    found = {}
    for value, key, supplier in zip(df[column], keys, df["卸名"]):
        if key and is_unmatched(key):
            entry = found.setdefault(key, [str(value).strip(), 0, set()])
            entry[1] += 1
            if not is_blank(supplier):
                entry[2].add(str(supplier))
    return [[name, count, "、".join(sorted(suppliers))] for name, count, suppliers in
            sorted(found.values(), key=lambda entry: entry[0])]


# =============================================================================
# STEP 0: データ統合
# =============================================================================
_MONTH_RE = re.compile(r"^(\d{4})(0[1-9]|1[0-2])$")
_FILE_MONTH_RE = re.compile(r"^(\d{4}(?:0[1-9]|1[0-2]))[_\-]")


def _input_files(folder):
    """入力フォルダにある xlsx (Excel の一時ファイル ~$ は除く)。"""
    try:
        return sorted((p for p in Path(folder).iterdir()
                       if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~")),
                      key=lambda p: p.name)
    except OSError:
        return []


def _month_value(value):
    text = normalize_key(as_number(value) if as_number(value) is not None else value)
    return int(text) if _MONTH_RE.match(text) else None


def _standardize(frame, supplier, path):
    """変換部品が返した表を点検し、統合用の 10 列にそろえる。戻り値: (表, 確認が必要な点)"""
    if not isinstance(frame, pd.DataFrame):
        raise StepError("load_file() が表 (DataFrame) を返しませんでした")
    frame = frame.copy()
    frame.columns = [str(c).strip() for c in frame.columns]
    unknown = [c for c in frame.columns if c not in PLUGIN_REQUIRED + PLUGIN_OPTIONAL]
    missing = [c for c in PLUGIN_REQUIRED if c not in frame.columns]
    if unknown or missing:
        detail = (["足りない列: " + "、".join(missing)] if missing else []) + \
                 (["標準形式に無い列: " + "、".join(unknown)] if unknown else [])
        raise StepError("変換結果の列が標準形式と違います (" + " / ".join(detail) + ")")

    index = range(len(frame))
    data = {c: [_tidy(v) for v in frame[c]] for c in frame.columns}
    keep = [i for i in index if any(data[c][i] is not None for c in PLUGIN_REQUIRED)]     # 中身の無い行は捨てる
    out = pd.DataFrame({c: _column((values[i] for i in keep), range(len(keep))) for c, values in data.items()},
                       index=range(len(keep)), columns=list(data), dtype=object)
    for column in PLUGIN_OPTIONAL:
        if column not in out.columns:
            out[column] = _column([None] * len(out), out.index)

    issues = []
    name = f"{supplier.name}: {path.name}"
    if "計上月" in frame.columns:
        months = [_month_value(v) for v in out["計上月"]]
        if any(m is None for m in months):
            issues.append(f"{name} に、年月 (YYYYMM) として読めない計上月が {sum(m is None for m in months)} 行あります")
    else:
        match = _FILE_MONTH_RE.match(path.name)
        months = [int(match.group(1)) if match else None] * len(out)
        if not match:
            issues.append(f"{name} のファイル名の先頭が年月 6 桁 (例: 202610_) になっていないため、計上月が空になります")
    out["計上月"] = _column(months, out.index)
    for column, label in (("販売本数", "販売本数"), ("商品名称", "商品名"), ("得意先名称", "得意先名")):
        if len(out) and all(v is None for v in out[column]):
            issues.append(f"{name} の{label}が全行で空です (見出しが変わった可能性があります)")
    out["元ファイル"] = path.name
    out["卸名"] = supplier.name
    return out[MERGED_COLUMNS], issues


def merge_suppliers(supplier_dirs, output_path):
    """各仕入先の変換部品でファイルを標準形式に直し、1 つの作業ファイルに統合する。

    supplier_dirs: [(Supplier, 入力フォルダ)]
    戻り値: {"rows": 行数, "errors": [1 件も取り込めなかった仕入先], "empty": [入力ファイルが無い仕入先],
             "issues": [確認が必要な点], "merged_files": {仕入先名: [取り込んだファイル名]}}
    """
    print("データ統合を開始します。")
    frames, errors, empty, issues, merged = [], [], [], [], {}
    for supplier, folder in supplier_dirs:
        files = _input_files(folder)
        if not files:
            empty.append(supplier.name)
            print(f"  {supplier.name}: 入力ファイルがありません ({folder})")
            continue
        try:
            module = _load_module(supplier.script, f"pipeline_supplier_{supplier.key}")
            loader = getattr(module, "load_file", None)
            if not callable(loader):
                raise StepError("load_file(path) が定義されていません")
        except Exception as e:
            errors.append(supplier.name)
            issues.append(f"{supplier.name}: 変換部品 {supplier.script.name} を読み込めません ({type(e).__name__}: {e})")
            continue

        loaded, rows = [], 0
        for path in files:
            try:
                frame, found = _standardize(loader(str(path)), supplier, path)
            except Exception as e:
                issues.append(f"{supplier.name}: {path.name} を取り込めませんでした ({type(e).__name__}: {e})")
                continue
            if frame.empty:
                issues.append(f"{supplier.name}: {path.name} に有効な行がありませんでした")
                continue
            frames.append(frame)
            issues.extend(found)
            loaded.append(path.name)
            rows += len(frame)
        if loaded:
            merged[supplier.name] = loaded
            print(f"  {supplier.name}: {len(loaded)} ファイル、{rows} 行を取り込みました。")
        else:
            errors.append(supplier.name)
            print(f"  {supplier.name}: 取り込めたファイルがありません。")

    for issue in issues:
        print(f"注意: {issue}")
    if not frames:
        raise StepError("統合できるデータがありませんでした。"
                        "入力データのルートフォルダと、入力サブフォルダ名を確認してください。")
    result = pd.concat(frames, ignore_index=True).astype(object)
    _save_work(result, output_path)
    print(f"合計 {len(result)} 行を統合しました。")
    return {"rows": len(result), "errors": errors, "empty": empty, "issues": issues, "merged_files": merged}


# =============================================================================
# STEP 1: 商品名の標準化
# =============================================================================
def standardize_products(work_path, product_master_path, unmatched_path):
    """商品マスタの「変換表」で商品名を標準の名前に直し、「標準商品」に無い商品名を一覧にする。

    戻り値: {"replaced": 直した行数, "unmatched": [商品マスタに無い商品名], "conflicts": [変換表の食い違い]}
    """
    print("商品名の標準化を開始します。")
    _check_ready(1, work_path)
    table = _read_table(product_master_path, "変換表", "商品マスタ")
    _require_columns(table, ["元の表記", "標準商品名"], "商品マスタ", "変換表")
    valid = _read_table(product_master_path, "標準商品", "商品マスタ")
    _require_columns(valid, ["商品名称"], "商品マスタ", "標準商品")

    mapping, conflicts = {}, []
    for raw, standard in zip(table["元の表記"], table["標準商品名"]):
        key = normalize_key(raw)
        if not key or is_blank(standard):
            continue
        standard = str(standard).strip()
        if mapping.setdefault(key, standard) != standard:          # 同じ表記に別の標準名 → 上の行を使う
            conflicts.append(f"{str(raw).strip()} → {mapping[key]} / {standard}")
    names = {normalize_key(v): str(v).strip() for v in valid["商品名称"] if not is_blank(v)}
    print(f"変換表 {len(mapping)} 件、標準商品 {len(names)} 件を読み込みました。")

    df = _read_work(work_path)
    converted, replaced = [], 0
    for value in df["商品名称"]:
        if is_blank(value):
            converted.append(None)
            continue
        key = normalize_key(value)
        new = mapping.get(key) or names.get(key) or str(value).strip()     # 表記ゆれだけの違いも標準名にそろえる
        replaced += new != value
        converted.append(new)
    df["商品名称"] = _column(converted, df.index)
    if any(column in df.columns for column in PRICE_COLUMNS):
        # 価格は商品名で引いている。商品名を直したら前の価格は当てにならないので、列ごと外して付け直してもらう
        df = df.drop(columns=[c for c in PRICE_COLUMNS if c in df.columns])
        print("商品名を直したため、付いていた価格を外しました。STEP 3 (価格の付与) をもう一度実行してください。")

    keys = [normalize_key(v) for v in df["商品名称"]]
    rows = _unmatched_rows(df, "商品名称", keys, lambda key: key not in names)
    _save_list(rows, ["商品名称", "行数", "卸名"], unmatched_path)     # 一覧を先に保存する (失敗しても作業ファイルは元のまま)
    _save_work(df, work_path)
    print(f"商品名を {replaced} 行で標準の名前に直しました。")
    if rows:
        print(f"商品マスタに無い商品名が {len(rows)} 件あります ({os.path.basename(str(unmatched_path))} に出力)。")
        for name, count, suppliers in rows[:20]:
            print(f"    {name}  ({count} 行 / {suppliers})")
    else:
        print("すべての商品名が商品マスタにありました。")
    for conflict in conflicts:
        print(f"注意: 変換表で同じ表記に別々の標準商品名が登録されています: {conflict}")
    return {"replaced": replaced, "unmatched": [row[0] for row in rows], "conflicts": conflicts}


# =============================================================================
# STEP 2: 顧客情報の付与
# =============================================================================
def _lookup_table(master, key_column, value_columns, convert=lambda column, value: _tidy(value)):
    """マスタから {照合キー: 値の組} を作る。同じキーが複数あれば上の行を使い、食い違いを報告する。"""
    lookup, conflicts = {}, []
    for row in zip(master[key_column], *(master[c] for c in value_columns)):
        key = normalize_key(row[0])
        if not key:
            continue
        values = tuple(convert(column, value) for column, value in zip(value_columns, row[1:]))
        if lookup.setdefault(key, values) != values and str(row[0]).strip() not in conflicts:
            conflicts.append(str(row[0]).strip())
    return lookup, conflicts


def _apply_lookup(df, key_column, lookup, value_columns, layout):
    """照合結果を列として付ける (すでに列があれば作り直すので、何度実行しても同じ結果になる)。"""
    keys = [normalize_key(v) for v in df[key_column]]
    empty = (None,) * len(value_columns)
    for position, column in enumerate(value_columns):
        df[column] = _column((lookup.get(key, empty)[position] for key in keys), df.index)
    return df[[c for c in layout if c in df.columns]], keys


def add_customer_info(work_path, customer_master_path, unmatched_path):
    """顧客マスタを得意先名称で引き、企業名・担当部署・担当者を付ける。

    顧客マスタにあった得意先は、名前の書き方もマスタの表記にそろえる (全角/半角の違いなどで、
    同じ得意先が BI 上で別々に集計されないようにするため)。
    戻り値: {"matched_rows": 付与できた行数, "unmatched": [顧客マスタに無い得意先], "conflicts": [食い違いのある得意先]}
    """
    print("顧客情報の付与を開始します。")
    _check_ready(2, work_path)
    master = _read_table(customer_master_path, "顧客", "顧客マスタ")
    _require_columns(master, ["得意先名称"] + ACCOUNT_COLUMNS, "顧客マスタ", "顧客")
    lookup, conflicts = _lookup_table(master, "得意先名称", ACCOUNT_COLUMNS)
    names = {}
    for name in master["得意先名称"]:
        names.setdefault(normalize_key(name), str(name).strip())
    print(f"顧客マスタ {len(lookup)} 件を読み込みました。")

    df = _read_work(work_path)
    df, keys = _apply_lookup(df, "得意先名称", lookup, ACCOUNT_COLUMNS,
                             MERGED_COLUMNS + ACCOUNT_COLUMNS + PRICE_COLUMNS)
    df["得意先名称"] = _column((names[key] if key in lookup else _tidy(value)
                               for key, value in zip(keys, df["得意先名称"])), df.index)

    matched = sum(1 for key in keys if key in lookup)
    rows = _unmatched_rows(df, "得意先名称", keys, lambda key: key not in lookup)
    _save_list(rows, ["得意先名称", "行数", "卸名"], unmatched_path)   # 一覧を先に保存する (失敗しても作業ファイルは元のまま)
    _save_work(df, work_path)
    print(f"{matched} / {len(df)} 行に顧客情報を付けました。")
    if rows:
        print(f"顧客マスタに無い得意先が {len(rows)} 件あります ({os.path.basename(str(unmatched_path))} に出力)。"
              f"これらの行の企業名は「{BLANK_COMPANY}」になります。")
    for name in conflicts:
        print(f"注意: 顧客マスタに同じ得意先が複数あり、内容が食い違っています (上の行を使いました): {name}")
    return {"matched_rows": matched, "unmatched": [row[0] for row in rows], "conflicts": conflicts}


# =============================================================================
# STEP 3: 価格の付与
# =============================================================================
def add_prices(work_path, price_master_path):
    """価格マスタを商品名称で引き、上代・仕切を付ける。

    価格は数値でなければならない (文字が混ざると金額の計算が壊れる)。数値として読めない価格は
    書き込まずに "invalid" で報告する。
    戻り値: {"priced_rows": 仕切が付いた行数, "unpriced": [仕切が付かなかった商品名],
             "invalid": [(商品名, 列名, 値)], "conflicts": [食い違いのある商品名]}
    """
    print("価格の付与を開始します。")
    _check_ready(3, work_path)
    master = _read_table(price_master_path, "価格", "価格マスタ")
    _require_columns(master, ["商品名称"] + PRICE_COLUMNS, "価格マスタ", "価格")

    invalid = [(str(name).strip(), column, str(value).strip())
               for column in PRICE_COLUMNS
               for name, value in zip(master["商品名称"], master[column])
               if normalize_key(name) and not is_blank(value) and as_number(value) is None]
    lookup, conflicts = _lookup_table(master, "商品名称", PRICE_COLUMNS, convert=lambda column, value: as_number(value))
    print(f"価格マスタ {len(lookup)} 件を読み込みました。")

    df = _read_work(work_path)
    df, keys = _apply_lookup(df, "商品名称", lookup, PRICE_COLUMNS,
                             MERGED_COLUMNS + ACCOUNT_COLUMNS + PRICE_COLUMNS)
    _save_work(df, work_path)

    priced = sum(1 for value in df["仕切"] if value is not None)
    unpriced = sorted({str(name).strip() for name, cost in zip(df["商品名称"], df["仕切"])
                       if cost is None and not is_blank(name)})
    print(f"{priced} / {len(df)} 行に仕切を付けました。")
    if unpriced:
        print(f"注意: 仕切が付かなかった商品が {len(unpriced)} 件あります (価格マスタに無い、または価格が空)。")
        for name in unpriced[:20]:
            print(f"    {name}")
    for name, column, raw in invalid[:20]:
        print(f"注意: 価格マスタに数値でない価格があります: {name} の{column} = {raw}")
    for name in conflicts:
        print(f"注意: 価格マスタに同じ商品が複数あり、価格が食い違っています (上の行を使いました): {name}")
    return {"priced_rows": priced, "unpriced": unpriced, "invalid": invalid, "conflicts": conflicts}


# =============================================================================
# STEP 4: 最終調整
# =============================================================================
def final_adjustment(work_path):
    """数値の列を数値にそろえ、計算列を加えて、17 列の形に整える。

    戻り値: {"rows": 行数, "non_numeric": {列名: 数値として読めなかった行数}}  読めなかった値は空にする。
    """
    print("最終調整を開始します。")
    _check_ready(4, work_path)
    df = _read_work(work_path)

    non_numeric = {}
    for column, convert in (("容量", parse_capacity), ("入数", as_number), ("販売本数", as_number),
                            ("上代", as_number), ("仕切", as_number)):
        values = [None if is_blank(v) else convert(v) for v in df[column]]
        bad = sum(1 for before, after in zip(df[column], values) if after is None and not is_blank(before))
        if bad:
            non_numeric[column] = bad
            print(f"注意: 「{column}」に数値として読めない値が {bad} 行ありました。その行の「{column}」は空にします。")
        df[column] = _column(values, df.index)

    df["売上金額"] = _column((None if quantity is None or cost is None else _tidy(quantity * cost)
                             for quantity, cost in zip(df["販売本数"], df["仕切"])), df.index)
    months = [_month_value(v) for v in df["計上月"]]
    df["計上月"] = _column(months, df.index)
    df["年度"] = _column((None if m is None else fiscal_year(m) for m in months), df.index)
    df["商品区分"] = _column((None if is_blank(name) else
                             (PRODUCT_CLASS_DISCONTINUED if DISCONTINUED_MARK in str(name) else PRODUCT_CLASS_NORMAL)
                             for name in df["商品名称"]), df.index)
    df["企業名"] = _column((BLANK_COMPANY if is_blank(v) else v for v in df["企業名"]), df.index)

    df = df[FINAL_COLUMNS]
    _save_work(df, work_path)
    print(f"売上金額・年度・商品区分を加え、{len(df.columns)} 列に整えました。({len(df)} 行)")
    return {"rows": len(df), "non_numeric": non_numeric}


# =============================================================================
# STEP 5: 都道府県の抽出
# =============================================================================
PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
    "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
    "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
    "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]


def find_prefecture(address, city_pairs):
    """住所から都道府県を求める。都道府県名が書かれていなければ、市区町村名から引く。"""
    text = normalize_key(address)
    if not text:
        return None
    for prefecture in PREFECTURES:
        if prefecture in text:
            return prefecture
    for city, prefecture in city_pairs:           # 長い名前から先に試す
        if city in text:
            return prefecture
    return None


def extract_prefecture(work_path, prefecture_table_path):
    """住所から都道府県と都内区分を取り出し、「住所」の右に 2 列を加える。

    戻り値: {"rows": 行数, "found": 都道府県が分かった行数, "not_found": [都道府県が分からなかった住所]}
    """
    print("都道府県の抽出を開始します。")
    _check_ready(5, work_path)
    cities = _read_table(prefecture_table_path, "市区町村", "都道府県テーブル")
    _require_columns(cities, ["都道府県", "市区町村"], "都道府県テーブル", "市区町村")
    divisions = _read_table(prefecture_table_path, "都内区分", "都道府県テーブル")
    _require_columns(divisions, ["都内区分"], "都道府県テーブル", "都内区分")

    city_pairs = sorted({(normalize_key(city), str(prefecture).strip())
                         for prefecture, city in zip(cities["都道府県"], cities["市区町村"])
                         if not is_blank(prefecture) and normalize_key(city)},
                        key=lambda pair: (-len(pair[0]), pair))
    division_names = sorted({str(v).strip() for v in divisions["都内区分"] if not is_blank(v)},
                            key=lambda name: (-len(name), name))
    print(f"市区町村 {len(city_pairs)} 件、都内区分 {len(division_names)} 件を読み込みました。")

    df = _read_work(work_path)
    prefectures = [find_prefecture(address, city_pairs) for address in df["住所"]]

    def division(address):
        text = normalize_key(address)
        return next((name for name in division_names if normalize_key(name) in text), None)

    df["都道府県"] = _column(prefectures, df.index)
    df["都内区分"] = _column((division(address) if prefecture == KEY_REGION else None
                             for address, prefecture in zip(df["住所"], prefectures)), df.index)
    df = df[GEO_LAYOUT]
    _save_work(df, work_path)

    found = sum(1 for p in prefectures if p is not None)
    not_found = sorted({str(address).strip() for address, p in zip(df["住所"], prefectures)
                        if p is None and not is_blank(address)})
    print(f"都道府県を {found} / {len(df)} 行で取り出しました。")
    if not_found:
        print(f"注意: 都道府県が分からなかった住所が {len(not_found)} 件あります。")
        for address in not_found[:20]:
            print(f"    {address}")
    return {"rows": len(df), "found": found, "not_found": not_found}


# =============================================================================
# STEP 6: ID の付与
# =============================================================================
MASTER_SHEETS = ["商品マスタ", "都道府県マスタ", "エリアマスタ", "卸マスタ", "部署マスタ", "担当者マスタ", "カレンダーマスタ"]
# (売上データ側のキー列, 付ける ID 列, マスタのシート, マスタ側のキー列, マスタ側の ID 列)
LOOKUPS = [
    ("商品名称", "商品ID", "商品マスタ", "商品名称", "商品ID"),
    ("都道府県", "都道府県ID", "都道府県マスタ", "都道府県", "都道府県ID"),
    ("都内区分", "エリアID", "エリアマスタ", "エリア", "エリアID"),
    ("卸名", "卸ID", "卸マスタ", "卸名", "卸ID"),
    ("担当部署", "担当部署ID", "部署マスタ", "担当部署", "担当部署ID"),
    ("担当者", "担当者ID", "担当者マスタ", "担当者", "担当者ID"),
]


def add_ids(work_path, id_master_path, output_path):
    """マスタの ID を売上データに付け、マスタのシートと一緒に完成ファイルとして保存する。

    戻り値: {"id_counts": {ID列: 付いた行数}, "missing": {ID列: [マスタに無かったキー]},
             "conflicts": {マスタのシート: [同じキーに別々の ID が登録されている値]}}
    """
    print("ID の付与を開始します。")
    _check_ready(6, work_path)
    try:
        book = pd.read_excel(id_master_path, sheet_name=None, dtype=object, keep_default_na=False, na_values=[""])
    except ValueError as e:
        raise StepError(f"ID マスタを読み込めません: {e}") from e
    absent = [sheet for sheet in MASTER_SHEETS if sheet not in book]
    if absent:
        raise StepError("ID マスタに必要なシートがありません: " + "、".join(absent))
    masters = {}
    for sheet in MASTER_SHEETS:
        master = book[sheet]
        master.columns = [str(c).strip() for c in master.columns]
        masters[sheet] = master

    df = _read_work(work_path)
    id_counts, missing, conflicts = {}, {}, {}
    for sales_key, id_column, sheet, master_key, master_id in LOOKUPS:
        _require_columns(masters[sheet], [master_key, master_id], "ID マスタ", sheet)
        lookup, duplicated = _lookup_table(masters[sheet], master_key, [master_id])
        if duplicated:
            conflicts[sheet] = duplicated
            print(f"  注意: 「{sheet}」で、同じ{master_key}に別々の ID が登録されています (上の行を使いました): "
                  f"{_sample(duplicated)}")
        keys = [normalize_key(v) for v in df[sales_key]]
        df[id_column] = _column((lookup.get(key, (None,))[0] for key in keys), df.index)
        id_counts[id_column] = sum(1 for v in df[id_column] if v is not None)
        absent_keys = sorted({key for key, v in zip(keys, df[id_column]) if key and v is None})
        if absent_keys:
            missing[id_column] = absent_keys
        print(f"  {id_column}: {id_counts[id_column]} / {len(df)} 行に付与 (「{sheet}」{len(lookup)} 件)")
        if absent_keys:
            print(f"  注意: 「{sheet}」に無い{sales_key}が {len(absent_keys)} 件あります ({_sample(absent_keys)})")

    sheets = {DATA_SHEET: df[OUTPUT_COLUMNS]}
    sheets.update(masters)
    _save_sheets(sheets, output_path)
    return {"id_counts": id_counts, "missing": missing, "conflicts": conflicts}


# =============================================================================
# STEP 7: テーブル設定
# =============================================================================
def _table_name_part(sheet_name):
    name = re.sub(r"[^\w_]", "", sheet_name.strip().replace(" ", "_"))
    if name and name[0].isdigit():
        name = "T_" + name
    return (name or "Sheet")[:100]


def _display_width(value):
    """セルの表示幅のめやす (全角文字は 2 文字分)。"""
    return sum(2 if unicodedata.east_asian_width(ch) in "WFA" else 1 for ch in str(value))


def apply_tables(output_path, prefix="Data", style_name="TableStyleMedium9"):
    """完成ファイルの全シートを Excel のテーブル (Data_<シート名>) にする。BI から名前で参照できる。

    戻り値: {"sheets": テーブルにしたシート数, "skipped": [データ行が無いためテーブルにしなかったシート]}
    """
    print("テーブル設定を開始します。")
    _check_ready(7, output_path=output_path)
    workbook = openpyxl.load_workbook(output_path)
    used, skipped = set(), []
    for sheet in workbook.worksheets:
        for old in list(sheet.tables):
            del sheet.tables[old]
        if sheet.max_row < 2 or sheet.max_column < 1 or sheet.cell(1, 1).value is None:
            # データ行の無いテーブルを作ると、Excel がファイルを開けなくなる
            skipped.append(sheet.title)
            print(f"  注意: {sheet.title} にはデータ行が無いため、テーブルにしません。")
            continue
        base = f"{prefix}_{_table_name_part(sheet.title)}"
        name, n = base, 1
        while name in used:
            name, n = f"{base}_{n}", n + 1
        used.add(name)
        ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"
        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name=style_name, showFirstColumn=False, showLastColumn=False,
                                              showRowStripes=True, showColumnStripes=False)
        sheet.add_table(table)
        for column, cells in enumerate(sheet.iter_cols(min_row=1, max_row=min(sheet.max_row, 200)), start=1):
            width = max((_display_width(c.value) for c in cells if c.value is not None), default=8)
            sheet.column_dimensions[get_column_letter(column)].width = min(max(width + 4, 8), 48)
        print(f"  {sheet.title}: テーブル {name} ({ref})")
    if not used:
        raise StepError("テーブルを設定できるシートがありません。ファイルは変更していません。")
    _save_with_retry(workbook.save, output_path)
    return {"sheets": len(used), "skipped": skipped}


# =============================================================================
# STEP 8: 後処理
# =============================================================================
def _unique_destination(folder, name):
    """folder 内でまだ使われていないファイル名を返す (同名があれば日時を付ける)。"""
    target = Path(folder) / name
    if not target.exists():
        return target, False
    stem, suffix = Path(name).stem, Path(name).suffix
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate, n = Path(folder) / f"{stem}__{stamp}{suffix}", 2
    while candidate.exists():
        candidate, n = Path(folder) / f"{stem}__{stamp}_{n}{suffix}", n + 1
    return candidate, True


def archive_inputs(folder_pairs, merged_files=None):
    """取り込み済みの入力ファイルを processed フォルダへ移す。

    folder_pairs: [(仕入先名, 入力フォルダ, processed フォルダ)]
    merged_files: 直近の STEP 0 で取り込んだファイル名 {仕入先名: [ファイル名]}。
                  ここに載っているファイルだけを移す (取り込めなかったファイル、あとから届いたファイルは残す)。
                  記録が無い (None) ときは、どれを取り込んだのか分からないので 1 つも移さない。
    戻り値: {"moved": 件数, "renamed": [別名で保存したファイル], "left": [残したファイル],
             "failed": [移せなかったファイル], "unknown": 記録が無くて移さなかったなら True}
    """
    print("取り込み済みの入力ファイルを移動します。")
    moved, renamed, left, failed = 0, [], [], []
    if merged_files is None:
        print("  注意: STEP 0 の記録が無いため、どのファイルを取り込んだのか分かりません。入力ファイルは移動しません。")
        return {"moved": 0, "renamed": [], "left": [], "failed": [], "unknown": True}
    for name, source, destination in folder_pairs:
        source, destination = Path(source), Path(destination)
        if not source.is_dir():
            continue
        wanted = set(merged_files.get(name, []))
        targets, done = [], 0
        for path in sorted(p for p in source.iterdir() if p.is_file()):         # フォルダは移さない
            if path.name.startswith("~"):                                       # Excel の一時ファイル
                continue
            if path.suffix.lower() == ".xlsx" and path.name in wanted:
                targets.append(path)
            else:
                left.append(f"{name}/{path.name}")
        if not targets:
            continue
        destination.mkdir(parents=True, exist_ok=True)
        for path in targets:
            target, was_renamed = _unique_destination(destination, path.name)
            try:
                shutil.move(str(path), str(target))
                moved, done = moved + 1, done + 1
                if was_renamed:
                    renamed.append(f"{name}/{path.name} → {target.name}")
            except Exception as e:
                failed.append(f"{name}/{path.name}")
                print(f"  エラー: {path.name} を移動できませんでした: {e}")
        print(f"  {name}: {done} 件を processed へ移動しました。")
    for item in renamed:
        print(f"  注意: processed に同名のファイルがあったため、別名で保存しました: {item}")
    for item in left:
        print(f"  移動しなかったファイル (今回の統合に含まれていません): {item}")
    return {"moved": moved, "renamed": renamed, "left": left, "failed": failed, "unknown": False}


def backup_outputs(work_dir, names=(WORK_FILE, OUTPUT_FILE, UNMATCHED_PRODUCTS, UNMATCHED_CUSTOMERS)):
    """成果物を <作業フォルダ>/backup/<日時>/ へ移す。

    戻り値: {"folder": 移動先 (対象が無ければ None), "failed": [移せなかったファイル名]}
    """
    print("成果物をバックアップします。")
    work_dir = Path(work_dir)
    existing = [work_dir / name for name in names if (work_dir / name).exists()]
    if not existing:
        print("  バックアップする成果物がありませんでした。")
        return {"folder": None, "failed": []}
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    folder, n = work_dir / "backup" / stamp, 2
    while folder.exists():                              # 同じ秒に 2 回実行した場合
        folder, n = work_dir / "backup" / f"{stamp}_{n}", n + 1
    folder.mkdir(parents=True)
    failed = []
    for path in existing:
        try:
            shutil.move(str(path), str(folder))
            print(f"  {path.name} を移動しました。")
        except Exception as e:
            failed.append(path.name)
            print(f"  エラー: {path.name} を移動できませんでした: {e}")
    print(f"  バックアップ先: {folder}")
    return {"folder": folder, "failed": failed}


# =============================================================================
# 前処理: csv / xls / xlsm → xlsx
# =============================================================================
CONVERTIBLE = (".csv", ".xls", ".xlsm")
ORIGINALS_FOLDER = "converted_originals"
_SKIP_FOLDERS = {ORIGINALS_FOLDER, "processed", "backup"}


def _csv_rows(path):
    """CSV を行のリストとして読む (文字コードは UTF-8 → Shift_JIS の順に試す)。値は文字のまま。"""
    for encoding in ("utf-8-sig", "cp932"):
        try:
            with open(path, newline="", encoding=encoding) as f:
                return [row for row in csv.reader(f)]
        except UnicodeDecodeError:
            continue
    raise StepError("文字コードを判別できません (UTF-8 か Shift_JIS で保存してください)")


def _convert_one(path, target):
    """1 ファイルを値だけの xlsx に変換して target に保存する。"""
    suffix = path.suffix.lower()
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    if suffix == ".csv":
        sheet = workbook.create_sheet("Sheet1")
        for row in _csv_rows(path):
            sheet.append([None if cell == "" else cell for cell in row])     # 先頭の 0 を落とさないよう文字のまま
    elif suffix == ".xls":
        book = pd.read_excel(path, sheet_name=None, header=None, dtype=object, keep_default_na=False, na_values=[""])
        for name, frame in book.items():
            sheet = workbook.create_sheet(str(name)[:31])
            for row in frame.itertuples(index=False):
                sheet.append([_tidy(v) if not isinstance(v, pd.Timestamp) else v.to_pydatetime() for v in row])
    else:                                                                    # xlsm: 数式は計算結果の値にする
        source = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            for source_sheet in source.worksheets:
                sheet = workbook.create_sheet(source_sheet.title)
                for row in source_sheet.iter_rows(values_only=True):
                    sheet.append(list(row))
        finally:
            source.close()
    if not workbook.worksheets:
        workbook.create_sheet("Sheet1")
    _save_with_retry(workbook.save, target)


def convert_folder(folder):
    """フォルダ配下 (サブフォルダを含む) の csv / xls / xlsm を xlsx に変換する。

    元のファイルは消さずに、同じ場所の converted_originals フォルダへ移す。すでにある xlsx には触れない。
    戻り値: {"ok": 変換できた件数, "failed": [変換できなかったファイル]}
    """
    folder = Path(folder)
    files = [p for p in sorted(folder.rglob("*"))
             if p.is_file() and p.suffix.lower() in CONVERTIBLE and not p.name.startswith("~")
             and not _SKIP_FOLDERS & {part.lower() for part in p.relative_to(folder).parts[:-1]}]
    if any(p.suffix.lower() == ".xls" for p in files) and importlib.util.find_spec("xlrd") is None:
        raise StepError(".xls ファイルの変換には xlrd が必要です。コマンド プロンプトで pip install xlrd を実行してください。")
    ok, failed = 0, []
    for path in files:
        target = path.with_suffix(".xlsx")
        shown = str(path.relative_to(folder))
        if target.exists():
            failed.append(f"{shown} (同じ名前の xlsx がすでにあります)")
            continue
        try:
            _convert_one(path, target)
            originals = path.parent / ORIGINALS_FOLDER
            originals.mkdir(exist_ok=True)
            shutil.move(str(path), str(_unique_destination(originals, path.name)[0]))
            ok += 1
            print(f"  変換しました: {shown} → {target.name}")
        except Exception as e:
            failed.append(f"{shown} ({type(e).__name__}: {e})")
    print(f"変換できたファイル: {ok} 件、変換できなかったファイル: {len(failed)} 件")
    for item in failed:
        print(f"  変換できなかったファイル: {item}")
    return {"ok": ok, "failed": failed}
