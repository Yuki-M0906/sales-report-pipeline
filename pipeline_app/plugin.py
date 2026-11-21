# -*- coding: utf-8 -*-
"""仕入先ごとの変換部品 (suppliers/process_*.py) から使う小さな道具。

変換部品の役目は 1 つだけ:
    load_file(path) で 1 つのファイルを読み、「標準形式」の表 (DataFrame) にして返す。

標準形式の列:
    必須 … 得意先名称、商品名称、販売本数
    任意 … 容量、入数、住所、業態、計上月   (無い列は空として扱われる)
元ファイル名・卸名・計上月 (ファイル名の先頭 6 桁) は、呼び出し側が自動で付ける。
"""

import pandas as pd

REQUIRED_COLUMNS = ["得意先名称", "商品名称", "販売本数"]
OPTIONAL_COLUMNS = ["容量", "入数", "住所", "業態", "計上月"]


def read_sheet(path, sheet=0, header_row=1):
    """シートを表として読む。

    sheet      : シート名、または 0 始まりの番号
    header_row : 見出しがある行 (Excel の行番号、1 始まり)。None なら見出し無しで全行を読む。
    値の型は変えない (空のセルは NaN)。見出しの前後の空白は取り除く。
    """
    header = None if header_row is None else header_row - 1
    df = pd.read_excel(path, sheet_name=sheet, header=header, dtype=object, keep_default_na=False, na_values=[""])
    if header is not None:
        df.columns = [str(c).strip() for c in df.columns]
    return df


def select(df, mapping):
    """元の表から必要な列だけを取り出し、標準形式の列名に付け替える。

    mapping: {標準形式の列名: 元の表の見出し}
    元の表に見出しが無ければ、どの見出しが無いかを示してエラーにする
    (仕入先が様式を変えたことに、黙って空の列を作らずに気づけるようにするため)。
    """
    missing = [source for source in mapping.values() if source not in df.columns]
    if missing:
        raise ValueError("ファイルに次の見出しがありません: " + "、".join(str(m) for m in missing)
                         + " (様式が変わっていないか確認してください)")
    out = df[list(mapping.values())].copy()
    out.columns = list(mapping.keys())
    return out


def drop_rows_containing(df, column, words):
    """指定した列に words のいずれかを含む行 (合計行・小計行など) を取り除く。"""
    text = df[column].map(lambda v: "" if pd.isna(v) else str(v))
    return df[~text.map(lambda t: any(word in t for word in words))].reset_index(drop=True)
