# -*- coding: utf-8 -*-
"""サンプル卸C: 店舗 × 商品 の「横持ち」の集計表。

    | 店舗名 | 所在地 | 業種 | 商品1 | 商品2 | ... |     ← 1 行目: 見出し (4 列目から右は商品名)
    |        |        |      | 1L    | 500ml | ... |     ← 2 行目: 各商品の容量
    | 店舗X  | ...    | ...  | 12    |       | ... |     ← 3 行目から: 店舗ごとの販売本数

標準形式は「1 行 = 1 店舗 × 1 商品」なので、数量が入っているセルごとに 1 行を作る (縦持ちに直す)。
"""

import pandas as pd

from pipeline_app import plugin

SUPPLIER_NAME = "サンプル卸C"
FIXED_COLUMNS = 3          # 左から 3 列 (店舗名・所在地・業種) は店舗の情報
EXPECTED_HEAD = ["店舗名", "所在地", "業種"]


def load_file(path):
    raw = plugin.read_sheet(path, sheet=0, header_row=None)
    head = [str(v).strip() for v in raw.iloc[0, :FIXED_COLUMNS]]
    if head != EXPECTED_HEAD:
        raise ValueError(f"1 行目の見出しが想定と違います: {head} (想定: {EXPECTED_HEAD})")
    products = raw.iloc[0, FIXED_COLUMNS:]
    capacities = raw.iloc[1, FIXED_COLUMNS:]

    rows = []
    for _, record in raw.iloc[2:].iterrows():
        store, address, business = record.iloc[:FIXED_COLUMNS]
        if pd.isna(store):
            continue
        for product, capacity, quantity in zip(products, capacities, record.iloc[FIXED_COLUMNS:]):
            if pd.isna(product) or pd.isna(quantity):
                continue
            rows.append({"得意先名称": store, "商品名称": product, "容量": capacity,
                         "販売本数": quantity, "住所": address, "業態": business})
    return pd.DataFrame(rows, columns=["得意先名称", "商品名称", "容量", "販売本数", "住所", "業態"])
