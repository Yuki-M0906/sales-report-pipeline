# -*- coding: utf-8 -*-
"""サンプル卸B: 表の上に表題の行があり、末尾に合計行がある帳票。

- 見出しは 4 行目にある (1〜3 行目は表題と対象月)。
- 最後の「合計」行は明細ではないので取り除く (残すと販売本数が二重に数えられる)。
- 業態の列は無い (標準形式の任意の列は、返さなければ空として扱われる)。
"""

from pipeline_app import plugin

SUPPLIER_NAME = "サンプル卸B"


def load_file(path):
    df = plugin.read_sheet(path, sheet=0, header_row=4)
    df = plugin.select(df, {
        "得意先名称": "店舗",
        "商品名称": "品名",
        "容量": "容量(ml)",
        "入数": "入数",
        "販売本数": "出荷本数",
        "住所": "住所",
    })
    return plugin.drop_rows_containing(df, "得意先名称", ["合計"])
