# -*- coding: utf-8 -*-
"""サンプル卸A: 1 行目が見出しの、素直な明細表。

いちばん単純な形。必要な列を取り出して、標準形式の列名に付け替えるだけ。
"""

from pipeline_app import plugin

SUPPLIER_NAME = "サンプル卸A"


def load_file(path):
    df = plugin.read_sheet(path, sheet="売上明細", header_row=1)
    return plugin.select(df, {
        # 標準形式の列名: このファイルでの見出し
        "得意先名称": "得意先名",
        "商品名称": "商品名",
        "容量": "規格",            # "1L" や "500ml" のような文字のままでよい (STEP 4 が ml の数値にそろえる)
        "入数": "ケース入数",
        "販売本数": "数量",
        "住所": "届け先住所",
        "業態": "業種",
    })
