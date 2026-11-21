# -*- coding: utf-8 -*-
"""サンプル一式 (sample/ フォルダ) を作り直す。

    python tools/make_sample_data.py                 … 問題の無い、きれいなサンプル
    python tools/make_sample_data.py --with-issues   … 「要確認」になる行を混ぜたサンプル

作るもの:
    sample/input/<仕入先名>/input/*.xlsx   仕入先ごとに形の違う販売実績ファイル (3 社)
    sample/convert_demo/*.csv              前処理 (csv → xlsx 変換) を試すためのファイル
    sample/masters/*.xlsx                  マスタ 5 種
    sample/expected/expected_output.xlsx   正しく処理できたときの出力 (検証用の期待値)
    sample/pipeline_config.json            サンプル用の設定 (相対パス)
    sample/work/                           作業フォルダ (中身は空にする)

データはすべて架空のもの。期待値は、処理プログラム (pipeline_app) を使わずに、
このスクリプトの中の「元になる事実の表」から直接組み立てている (検証の独立性を保つため)。
"""

import argparse
import csv
import json
import os
import shutil
import stat
import sys
import time
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "sample"
MARKER = "SAMPLE_DATA_FOLDER.txt"      # サンプル用のフォルダであることの目印

# ----------------------------------------------------------------------------- マスタの中身
# (商品ID, 標準の商品名, カテゴリ, 容量ml, 入数, 上代, 仕切)
PRODUCTS = [
    ("P001", "りんごジュース 1L", "飲料", 1000, 6, 480, 312),
    ("P002", "みかんジュース 1L", "飲料", 1000, 6, 480, 312),
    ("P003", "緑茶 500ml", "飲料", 500, 24, 160, 104),
    ("P004", "ほうじ茶 500ml", "飲料", 500, 24, 160, 104),
    ("P005", "だししょうゆ 300ml", "調味料", 300, 12, 420, 273),
    ("P006", "ぽん酢 360ml", "調味料", 360, 12, 380, 247),
    ("P007", "ごまドレッシング 200ml", "調味料", 200, 12, 350, 227.5),
    ("P008", "麦茶 2L（終売）", "飲料", 2000, 6, 300, 195),
]
# 仕入先ごとの商品名の書き方 (標準の商品名の並びと同じ順)
NAMES_A = ["りんごジュース 1L", "みかんジュース　１Ｌ", "緑茶500", "焙じ茶 500ml",
           "だし醤油 300", "ポン酢 360ml", "胡麻ドレ 200ml", "麦茶 2L（終売）"]
NAMES_B = ["APPLE JUICE 1000", "ORANGE JUICE 1000", "GREEN TEA 500", "ROASTED TEA 500",
           "DASHI SOY 300", "PONZU 360", "SESAME DRESSING 200", "BARLEY TEA 2000"]
NAMES_C = ["りんごJ 1L", "みかんJ 1L", "緑茶 500", "ほうじ茶 500", "だししょうゆ 300",
           "ぽん酢 360", "ごまドレ 200", "麦茶 2L"]
CAPACITY_TEXT = ["1L", "1L", "500ml", "500ml", "300ml", "360ml", "200ml", "2L"]

# (得意先名称, 企業名, 担当部署, 担当者, 住所, 業態)   企業名などが None の得意先は、顧客マスタに載せない
CUSTOMERS = {
    "C1": ("A食堂 本店", "A社", "営業1課", "担当A", "東京都港区見本町1-2-3", "飲食店"),
    "C2": ("A食堂 新宿店", "A社", "営業1課", "担当A", "東京都新宿区見本町4-5", "飲食店"),
    "C3": ("Bカフェ", "B社", "営業1課", "担当B", "横浜市中区見本町6-7", "飲食店"),
    "C4": ("Cスーパー 川崎店", "C社", "営業2課", "担当C", "神奈川県川崎市川崎区見本町8", "小売店"),
    "C5": ("Cスーパー 千葉店", "C社", "営業2課", "担当C", "千葉市中央区見本町9-10", "小売店"),
    "C6": ("Dホテル", "D社", "営業2課", "担当D", "静岡県熱海市見本町11", "ホテル"),
    "C7": ("和食処E", "", "営業3課", "担当E", "大阪府大阪市北区見本町12-13", "飲食店"),
    "C8": ("F商店", "F社", "営業3課", "担当E", "大阪市中央区見本町14", "小売店"),
    "C9": ("G台所 札幌店", "G社", "営業3課", "担当F", "北海道札幌市中央区見本町15", "飲食店"),
    "C10": ("Hマルシェ", "H社", "営業1課", "担当B", "東京都八王子市見本町16", "小売店"),
    "C11": ("Iデリ", None, None, None, "東京都町田市見本町17", "飲食店"),
    "C12": ("Jキッチン", "J社", "営業3課", "担当F", "福岡市博多区見本町18", "飲食店"),
}
SUPPLIERS = [("W01", "サンプル卸A"), ("W02", "サンプル卸B"), ("W03", "サンプル卸C")]
DEPARTMENTS = [("D1", "営業1課"), ("D2", "営業2課"), ("D3", "営業3課")]
STAFF = [("S01", "担当A", "営業1課"), ("S02", "担当B", "営業1課"), ("S03", "担当C", "営業2課"),
         ("S04", "担当D", "営業2課"), ("S05", "担当E", "営業3課"), ("S06", "担当F", "営業3課")]

PREFECTURES = [
    ("北海道", "北海道"), ("青森県", "東北"), ("岩手県", "東北"), ("宮城県", "東北"), ("秋田県", "東北"),
    ("山形県", "東北"), ("福島県", "東北"), ("茨城県", "関東"), ("栃木県", "関東"), ("群馬県", "関東"),
    ("埼玉県", "関東"), ("千葉県", "関東"), ("東京都", "関東"), ("神奈川県", "関東"), ("新潟県", "中部"),
    ("富山県", "中部"), ("石川県", "中部"), ("福井県", "中部"), ("山梨県", "中部"), ("長野県", "中部"),
    ("岐阜県", "中部"), ("静岡県", "中部"), ("愛知県", "中部"), ("三重県", "近畿"), ("滋賀県", "近畿"),
    ("京都府", "近畿"), ("大阪府", "近畿"), ("兵庫県", "近畿"), ("奈良県", "近畿"), ("和歌山県", "近畿"),
    ("鳥取県", "中国"), ("島根県", "中国"), ("岡山県", "中国"), ("広島県", "中国"), ("山口県", "中国"),
    ("徳島県", "四国"), ("香川県", "四国"), ("愛媛県", "四国"), ("高知県", "四国"), ("福岡県", "九州・沖縄"),
    ("佐賀県", "九州・沖縄"), ("長崎県", "九州・沖縄"), ("熊本県", "九州・沖縄"), ("大分県", "九州・沖縄"),
    ("宮崎県", "九州・沖縄"), ("鹿児島県", "九州・沖縄"), ("沖縄県", "九州・沖縄"),
]
# 住所に都道府県名が書かれていないときに、市区町村名から都道府県を引くための表 (主な市の例。実運用では全国分を用意する)
CITIES = [
    ("北海道", "札幌市"), ("宮城県", "仙台市"), ("埼玉県", "さいたま市"), ("千葉県", "千葉市"),
    ("神奈川県", "横浜市"), ("神奈川県", "川崎市"), ("神奈川県", "相模原市"), ("新潟県", "新潟市"),
    ("静岡県", "静岡市"), ("静岡県", "浜松市"), ("愛知県", "名古屋市"), ("京都府", "京都市"),
    ("大阪府", "大阪市"), ("大阪府", "堺市"), ("兵庫県", "神戸市"), ("岡山県", "岡山市"),
    ("広島県", "広島市"), ("福岡県", "北九州市"), ("福岡県", "福岡市"), ("熊本県", "熊本市"),
    ("東京都", "八王子市"), ("東京都", "町田市"), ("東京都", "立川市"), ("東京都", "武蔵野市"),
]
TOKYO_DIVISIONS = [
    "千代田区", "中央区", "港区", "新宿区", "文京区", "台東区", "墨田区", "江東区", "品川区", "目黒区",
    "大田区", "世田谷区", "渋谷区", "中野区", "杉並区", "豊島区", "北区", "荒川区", "板橋区", "練馬区",
    "足立区", "葛飾区", "江戸川区", "八王子市", "町田市", "立川市", "武蔵野市",
]

# ----------------------------------------------------------------------------- 元になる事実
# 仕入先ファイルごとの明細: (得意先, [(商品の番号 0〜7, 販売本数), ...])
SALES_A_09 = [("C1", [(0, 24), (2, 48), (4, 12)]), ("C2", [(1, 12), (3, 24)]), ("C3", [(0, 6), (2, 24), (6, 7)]),
              ("C4", [(0, 60), (1, 60), (7, 18)]), ("C10", [(4, 24), (5, 24)])]
SALES_A_10 = [("C1", [(0, 30), (2, 72), (4, 12)]), ("C2", [(1, 18), (3, 24), (6, 12)]), ("C3", [(0, 12), (2, 24)]),
              ("C4", [(0, 66), (1, 54), (7, 6)]), ("C10", [(4, 36), (5, 12)]), ("C11", [(2, 24), (6, 5)])]
SALES_B_10 = [("C4", [(2, 240), (3, 120)]), ("C5", [(0, 48), (2, 1200), (5, 36)]), ("C6", [(1, 30), (4, 24), (6, 24)]),
              ("C7", [(4, 60), (5, 48)]), ("C8", [(0, 18), (7, 12)])]
SALES_C_10 = [("C7", [(2, 48), (4, 24), (6, 12)]), ("C8", [(0, 12), (1, 12)]), ("C9", [(0, 36), (3, 48), (5, 24)]),
              ("C12", [(2, 96), (6, 36)]), ("C1", [(1, 6), (7, 6)])]

# 仕入先ごとに違う、得意先名の書き方のゆれ (照合では同じ得意先として扱われることを確かめる)
NAME_VARIANTS_B = {"C4": "Ｃスーパー　川崎店", "C5": "Cスーパー 千葉店 ", "C7": "和食処Ｅ"}
BLANK_ADDRESS_C = {"C1"}            # サンプル卸C のファイルでは住所が空の得意先

# --with-issues で サンプル卸A の 10 月分に加える行: (得意先名, 住所, 業種, 商品名, 規格, 入数, 数量)
ISSUE_ROWS = [
    ("A食堂 本店", "東京都港区見本町1-2-3", "飲食店", "季節限定 ゆずサイダー 350ml", "350ml", 24, 48),    # 商品マスタに無い商品
    ("Bカフェ", "横浜市中区見本町6-7", "飲食店", "緑茶500", "500ml", 24, "未定"),                          # 数値でない販売本数
    ("Kベーカリー", "見本郡見本町99", "小売店", "りんごジュース 1L", "1L", 6, 12),                         # 都道府県が分からない住所
]

OUTPUT_COLUMNS = [
    "卸名", "計上月", "担当部署", "担当者", "得意先名称", "企業名", "商品名称", "容量", "入数", "上代",
    "販売本数", "仕切", "売上金額", "年度", "商品区分", "住所", "都道府県", "都内区分",
    "商品ID", "都道府県ID", "エリアID", "卸ID", "担当部署ID", "担当者ID", "業態",
]


# ----------------------------------------------------------------------------- 道具
def save_book(path, sheets):
    """{シート名: 行のリスト} を xlsx として保存する。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    book = openpyxl.Workbook()
    book.remove(book.active)
    for name, rows in sheets.items():
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(list(row))
    book.properties.creator = "sample"
    book.save(path)


def whole(number):
    return int(number) if number == int(number) else number


def prefecture_of(address):
    if not address:
        return None
    for name, _region in PREFECTURES:
        if name in address:
            return name
    for prefecture, city in sorted(CITIES, key=lambda pair: -len(pair[1])):
        if city in address:
            return prefecture
    return None


def expected_row(supplier, month, customer_name, raw_product, capacity, case_size, quantity, address, business):
    """1 明細について、正しく処理されたときの出力 25 列を組み立てる。"""
    customer = next((c for c in CUSTOMERS.values() if c[0] == customer_name and c[1] is not None), None)
    product = next((p for p, names in zip(PRODUCTS, zip(NAMES_A, NAMES_B, NAMES_C, [p[1] for p in PRODUCTS]))
                    if raw_product in names), None)
    prefecture = prefecture_of(address)
    division = next((d for d in sorted(TOKYO_DIVISIONS, key=lambda d: -len(d)) if d in (address or "")), None) \
        if prefecture == "東京都" else None
    quantity = quantity if isinstance(quantity, (int, float)) else None
    cost = product[6] if product else None
    name = product[1] if product else raw_product
    department = customer[2] if customer else None
    person = customer[3] if customer else None
    return [
        supplier, month, department, person, customer_name,
        customer[1] if customer and customer[1] else "(空白)",
        name, capacity, case_size, product[5] if product else None, quantity, cost,
        whole(quantity * cost) if quantity is not None and cost is not None else None,
        month // 100, "終売商品" if "終売" in name else "通常商品", address,
        prefecture, division,
        product[0] if product else None,
        next((i for i, (p, _r) in enumerate(PREFECTURES, start=1) if p == prefecture), None),
        f"A{TOKYO_DIVISIONS.index(division) + 1:02d}" if division else None,
        next(i for i, n in SUPPLIERS if n == supplier),
        next((i for i, n in DEPARTMENTS if n == department), None),
        next((i for i, n, _d in STAFF if n == person), None),
        business,
    ]


# ----------------------------------------------------------------------------- 仕入先ファイル
def build_supplier_a(month, sales, extra=()):
    """サンプル卸A: 1 行目が見出しの素直な明細表。"""
    rows, expected = [["得意先名", "届け先住所", "業種", "商品名", "規格", "ケース入数", "数量"]], []
    for code, items in sales:
        name, _company, _dept, _person, address, business = CUSTOMERS[code]
        for index, quantity in items:
            product = PRODUCTS[index]
            rows.append([name, address, business, NAMES_A[index], CAPACITY_TEXT[index], product[4], quantity])
            expected.append(expected_row("サンプル卸A", month, name, NAMES_A[index], product[3], product[4],
                                         quantity, address, business))
    for name, address, business, product_name, capacity, case_size, quantity in extra:
        rows.append([name, address, business, product_name, capacity, case_size, quantity])
        capacity_ml = {"350ml": 350, "500ml": 500, "1L": 1000}[capacity]
        expected.append(expected_row("サンプル卸A", month, name, product_name, capacity_ml, case_size,
                                     quantity, address, business))
    return {"売上明細": rows}, expected


def build_supplier_b(month, sales):
    """サンプル卸B: 表題の行が上にあり、末尾に合計行がある帳票。業態の列は無い。"""
    rows = [["販売実績報告書"], [f"対象月: {month // 100}年{month % 100}月"], [],
            ["店舗", "住所", "品名", "容量(ml)", "入数", "出荷本数"]]
    expected, total = [], 0
    for code, items in sales:
        name, _company, _dept, _person, address, _business = CUSTOMERS[code]
        for index, quantity in items:
            product = PRODUCTS[index]
            shown = f"{quantity:,}" if quantity >= 1000 else quantity          # 1,200 のような文字の数量も混ぜる
            rows.append([NAME_VARIANTS_B.get(code, name), address, NAMES_B[index], product[3], product[4], shown])
            expected.append(expected_row("サンプル卸B", month, name, NAMES_B[index], product[3], product[4],
                                         quantity, address, None))
            total += quantity
    rows.append(["合計", None, None, None, None, total])
    return {"報告書": rows}, expected


def build_supplier_c(month, sales):
    """サンプル卸C: 店舗 × 商品 の横持ちの集計表 (2 行目は容量)。入数の情報は無い。"""
    rows = [["店舗名", "所在地", "業種"] + NAMES_C, [None, None, None] + CAPACITY_TEXT]
    expected = []
    for code, items in sales:
        name, _company, _dept, _person, address, business = CUSTOMERS[code]
        address = None if code in BLANK_ADDRESS_C else address
        quantities = dict(items)
        rows.append([name, address, business] + [quantities.get(i) for i in range(len(PRODUCTS))])
        for index in sorted(quantities):
            expected.append(expected_row("サンプル卸C", month, name, NAMES_C[index], PRODUCTS[index][3], None,
                                         quantities[index], address, business))
    return {"集計": rows}, expected


# ----------------------------------------------------------------------------- マスタ
def build_masters(folder):
    conversions = [["元の表記", "標準商品名"]]
    for product, *names in zip(PRODUCTS, NAMES_A, NAMES_B, NAMES_C):
        for raw in names:
            if raw != product[1] and [raw, product[1]] not in conversions:
                conversions.append([raw, product[1]])
    conversions = [row for row in conversions if row[0] != "みかんジュース　１Ｌ"]      # 全角/半角だけの違いは登録不要
    save_book(folder / "product_master.xlsx", {
        "変換表": conversions,
        "標準商品": [["商品名称"]] + [[p[1]] for p in PRODUCTS],
    })
    save_book(folder / "customer_master.xlsx", {
        "顧客": [["得意先名称", "企業名", "担当部署", "担当者"]]
        + [[c[0], c[1] or None, c[2], c[3]] for c in CUSTOMERS.values() if c[1] is not None],
    })
    save_book(folder / "price_master.xlsx", {
        "価格": [["商品名称", "上代", "仕切"]] + [[p[1], p[5], p[6]] for p in PRODUCTS],
    })
    save_book(folder / "prefecture_table.xlsx", {
        "市区町村": [["都道府県", "市区町村"]] + [list(pair) for pair in CITIES],
        "都内区分": [["都内区分"]] + [[d] for d in TOKYO_DIVISIONS],
    })
    months = [202600 + m for m in range(1, 13)]
    save_book(folder / "id_master.xlsx", {
        "商品マスタ": [["商品ID", "商品名称", "カテゴリ"]] + [[p[0], p[1], p[2]] for p in PRODUCTS],
        "都道府県マスタ": [["都道府県ID", "都道府県", "地方"]]
        + [[i, name, region] for i, (name, region) in enumerate(PREFECTURES, start=1)],
        "エリアマスタ": [["エリアID", "エリア"]] + [[f"A{i:02d}", d] for i, d in enumerate(TOKYO_DIVISIONS, start=1)],
        "卸マスタ": [["卸ID", "卸名"]] + [list(s) for s in SUPPLIERS],
        "部署マスタ": [["担当部署ID", "担当部署"]] + [list(d) for d in DEPARTMENTS],
        "担当者マスタ": [["担当者ID", "担当者", "担当部署"]] + [list(s) for s in STAFF],
        "カレンダーマスタ": [["計上月", "年", "月", "四半期", "年月表示"]]
        + [[m, m // 100, m % 100, f"Q{(m % 100 - 1) // 3 + 1}", f"{m // 100}年{m % 100}月"] for m in months],
    })


# ----------------------------------------------------------------------------- 本体
def clear_folder(folder):
    """サンプルの置き場所の中の、決まった名前のフォルダを空にする (それ以外は決して消さない)。"""
    folder = folder.resolve()
    if folder.parent != SAMPLE.resolve():
        raise RuntimeError(f"サンプルの置き場所の外は消しません: {folder}")
    def writable_then_retry(function, path, _error):
        """クラウド同期のフォルダなどに付く「読み取り専用」の印を外してから、もう一度消す。"""
        os.chmod(path, stat.S_IWRITE)
        function(path)

    handler = {"onexc": writable_then_retry} if sys.version_info >= (3, 12) else {"onerror": writable_then_retry}
    for attempt in range(5):             # 同期ソフトやウイルス対策が、一時的にフォルダをつかんでいることがある
        try:
            if folder.exists():
                shutil.rmtree(folder, **handler)
            break
        except PermissionError:
            if attempt == 4:
                raise
            time.sleep(1)
    folder.mkdir(parents=True)


def main(argv=None):
    parser = argparse.ArgumentParser(description="サンプル一式を作り直します。")
    parser.add_argument("--with-issues", action="store_true", help="「要確認」になる行を混ぜる")
    parser.add_argument("--out", default=None, help="作る場所 (省略すると、このリポジトリの sample フォルダ)")
    args = parser.parse_args(argv)
    global SAMPLE
    if args.out:
        SAMPLE = Path(args.out).resolve()
        SAMPLE.mkdir(parents=True, exist_ok=True)

    # 中身を消して作り直すので、サンプル用のフォルダだと分かっている場所でしか動かさない
    # (本番の input や work があるフォルダを、うっかり指定しても消さないため)
    marker = SAMPLE / MARKER
    if any(SAMPLE.iterdir()) and not marker.exists():
        parser.error(f"{SAMPLE} は空ではなく、サンプル用のフォルダの目印 ({MARKER}) もありません。"
                     "中身を消してしまわないように中止します。空のフォルダを指定してください。")
    marker.write_text("このフォルダは tools/make_sample_data.py が作り直します。自分のデータを置かないでください。\n",
                      encoding="utf-8")

    for name in ("input", "masters", "expected", "work", "convert_demo"):
        clear_folder(SAMPLE / name)
    (SAMPLE / "work" / ".gitkeep").write_text("", encoding="utf-8")

    files = [
        ("サンプル卸A", "202609_サンプル卸A.xlsx", build_supplier_a(202609, SALES_A_09)),
        ("サンプル卸A", "202610_サンプル卸A.xlsx",
         build_supplier_a(202610, SALES_A_10, ISSUE_ROWS if args.with_issues else ())),
        ("サンプル卸B", "202610_サンプル卸B.xlsx", build_supplier_b(202610, SALES_B_10)),
        ("サンプル卸C", "202610_サンプル卸C.xlsx", build_supplier_c(202610, SALES_C_10)),
    ]
    expected = [OUTPUT_COLUMNS]
    for supplier, file_name, (sheets, rows) in files:
        save_book(SAMPLE / "input" / supplier / "input" / file_name, sheets)
        expected.extend(rows)
    save_book(SAMPLE / "expected" / "expected_output.xlsx", {"売上データ": expected})
    build_masters(SAMPLE / "masters")

    # 前処理を試すための CSV (サンプル卸C の 11 月分。Shift_JIS)
    sheets, _rows = build_supplier_c(202611, SALES_C_10[:3])
    with open(SAMPLE / "convert_demo" / "202611_サンプル卸C.csv", "w", newline="", encoding="cp932") as f:
        csv.writer(f).writerows([["" if v is None else v for v in row] for row in sheets["集計"]])

    config = {
        "work_dir": "work", "input_root": "input", "input_subdir": "input",
        "product_master": "masters/product_master.xlsx", "customer_master": "masters/customer_master.xlsx",
        "price_master": "masters/price_master.xlsx", "prefecture_table": "masters/prefecture_table.xlsx",
        "id_master": "masters/id_master.xlsx", "disabled_suppliers": [],
    }
    (SAMPLE / "pipeline_config.json").write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"サンプルを作りました: {SAMPLE}  (明細 {len(expected) - 1} 行"
          + ("、要確認の行あり)" if args.with_issues else ")"))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
