# 流用ガイド（他の組織で使うために） / Adaptation Guide

<a id="japanese"></a>

**日本語** | [English](#english)

このツールを、自分の組織の仕入先・マスタ・業務ルールに合わせて使うための手引きです。
Python を少し書ける人が、1 人で導入できることを想定しています。

- [1. 全体像: 何を用意し、何を変えるか](#1-全体像-何を用意し何を変えるか)
- [2. 導入の手順](#2-導入の手順)
- [3. 仕入先の変換部品を書く](#3-仕入先の変換部品を書く)
- [4. マスタを用意する](#4-マスタを用意する)
- [5. 業務ルールを変える](#5-業務ルールを変える)
- [6. 列の構成を変える](#6-列の構成を変える)
- [7. BI ツールに取り込む](#7-bi-ツールに取り込む)
- [8. 自分たちのテストを持つ](#8-自分たちのテストを持つ)
- [9. 機密情報を混ぜないために](#9-機密情報を混ぜないために)
- [10. 向いている場面、向いていない場面](#10-向いている場面向いていない場面)

---

## 1. 全体像: 何を用意し、何を変えるか

| | 何を | どこに | 作業量のめやす |
| --- | --- | --- | --- |
| **必ず用意する** | 仕入先ごとの変換部品 | `suppliers/process_*.py` | 1 社あたり 10〜50 行 |
| **必ず用意する** | マスタ 5 種 | 好きな場所（Excel） | 手持ちの台帳を、決まった見出しに整える |
| **必ず設定する** | フォルダとマスタの場所 | 画面の「設定」タブ | 数分 |
| 必要なら変える | 業務ルール（年度の開始月など） | `pipeline_app/steps.py` の冒頭 | 定数を書き換える |
| 必要なら変える | 出力の列の構成 | `pipeline_app/state.py` と `steps.py` | [6 章](#6-列の構成を変える) |
| 変えなくてよい | 画面、実行の制御、段階の判定、保存の仕組み | `gui.py` `runner.py` `state.py` | — |

**サンプルの仕入先とデータは、自分のものに置き換えてください。**
`suppliers/process_sample_*.py` を消す（または残したまま、画面の「設定」タブでチェックを外す）と、自分の仕入先だけが処理されます。

## 2. 導入の手順

1. **サンプルで動きを確かめる。** `run_sample.bat` で一通り動かし、入力（`sample/input/`）・マスタ（`sample/masters/`）・出力を見比べます。
2. **仕入先を 1 社だけ選び、変換部品を書く。**（[3 章](#3-仕入先の変換部品を書く)）いちばん素直な形の帳票から始めます。
3. **マスタを用意する。**（[4 章](#4-マスタを用意する)）最初は、サンプルのマスタをコピーして中身を入れ替えるのが早道です。
4. **`run_gui.bat` で起動し、「設定」タブを埋める。** 作業フォルダは空のフォルダを新しく作って指定します。
5. **STEP 0 だけを実行し、`sales_work.xlsx` を見る。** その仕入先のデータが、標準形式の 10 列に正しく入っているかを確かめます。
6. **STEP 1〜7 を実行し、「要確認」をつぶしていく。** マスタに足りないものが、要確認として順に出てきます。
7. **手で作った正解と突き合わせる。** 1 か月分を、これまでの方法（手作業や既存の仕組み）の結果と比べます。合計の本数と金額、行数、いくつかの行の中身。
8. **仕入先を 1 社ずつ増やす。** 増やすたびに 5〜7 を繰り返します。

一度に全社を載せようとしないでください。1 社ずつなら、数字が合わないときに原因がすぐ分かります。

## 3. 仕入先の変換部品を書く

### 3.1 部品の約束

`suppliers/_template.py` を `process_<英数字の名前>.py` という名前でコピーし、2 つのものを書きます。

```python
from pipeline_app import plugin

SUPPLIER_NAME = "サンプル卸A"          # ① 仕入先の名前


def load_file(path):                    # ② 1 つのファイルを読み、標準形式の表にして返す
    df = plugin.read_sheet(path, sheet="売上明細", header_row=1)
    return plugin.select(df, {
        "得意先名称": "得意先名",       # 標準形式の列名: このファイルでの見出し
        "商品名称": "商品名",
        "販売本数": "数量",
        "容量": "規格",
        "入数": "ケース入数",
        "住所": "届け先住所",
        "業態": "業種",
    })
```

**① `SUPPLIER_NAME`** は、次の 3 つに使われます。3 つが一致している必要があります。

- 入力フォルダの名前（`<入力データのルート>/<SUPPLIER_NAME>/input/`）
- 出力の「卸名」
- ID マスタ「卸マスタ」の「卸名」（卸 ID を引くキー）

**② `load_file(path)`** の戻り値は、「1 行 = 1 得意先 × 1 商品」の表（pandas の DataFrame）です。

| 列 | 必須 | 内容 |
| --- | --- | --- |
| 得意先名称 | ○ | 店・納品先の名前 |
| 商品名称 | ○ | 商品の名前（仕入先の書き方のままでよい） |
| 販売本数 | ○ | 本数 |
| 容量 | | 「720ml」「1.8L」のような文字のままでよい |
| 入数 | | 1 ケースの本数 |
| 住所 | | 得意先の住所 |
| 業態 | | 飲食店・小売店などの区分 |
| 計上月 | | ファイルの中の年月を使いたいときだけ（YYYYMM）。返さなければ、ファイル名の先頭 6 桁が使われる |

これ以外の名前の列を返すと、そのファイルはエラーになります（列名の書き間違いに気づくためです）。

### 3.2 部品でやること、やらなくてよいこと

**やること**

- 必要な列を取り出して、標準形式の列名にする
- **明細でない行を取り除く**（合計、小計、注記、空行）
- 横持ちの表を縦持ちに直す
- 想定と違うと分かったら、**例外を出す**

**やらなくてよいこと**（後の工程が全社共通で行います）

- 前後の空白の除去、全角/半角の統一
- 「1,200」→ 1200、「1.8L」→ 1800 などの数値化
- 商品名や得意先名を標準の名前に直すこと
- 元ファイル名・卸名・計上月を付けること

部品は「形を直す」ことだけに集中させ、「値を直す」ことは共通の処理に任せる、という分担です。
値の直し方を部品ごとに書くと、卸によって直し方が食い違います。

### 3.3 よくある形と書き方

同梱の 3 つの部品が、そのまま見本になります。

| 帳票の形 | 見本 | 要点 |
| --- | --- | --- |
| 1 行目が見出しの、素直な明細表 | `process_sample_a.py` | `read_sheet` と `select` だけ |
| 表題の行が上にある／合計行が下にある | `process_sample_b.py` | `header_row=4` で見出しの行を指定。`drop_rows_containing` で合計行を除く |
| 店舗 × 商品の横持ちの集計表 | `process_sample_c.py` | `header_row=None` で全行を読み、数量の入ったセルごとに 1 行を作る |

そのほかの形への対処:

| 帳票の形 | 対処 |
| --- | --- |
| 得意先名が、グループの最初の行にしか書かれていない | `df["得意先名称"] = df["得意先名称"].ffill()` で下の行へ埋める |
| 1 つのブックに、得意先ごと・月ごとのシートがある | `pd.read_excel(path, sheet_name=None)` で全シートを読み、`pd.concat` でつなぐ |
| 住所が「都道府県」「市区町村」「番地」の列に分かれている | 文字としてつないで「住所」にする |
| 数量が「ケース数」と「バラ本数」に分かれている | 本数に換算して「販売本数」にする（ケース数 × 入数 + バラ本数） |
| 返品がマイナスで入っている | そのまま返す（合計で相殺される）。除外するなら、ここで明示的に除く |
| ファイルが csv / xls で届く | 画面の「前処理」で xlsx に変換してから取り込む（部品は xlsx だけを相手にすればよい） |

### 3.4 黙って読み飛ばさない

仕入先は、予告なく様式を変えます。そのとき、**部品がエラーを出さずに空の表や欠けた表を返す**のが最悪の事態です。

- `plugin.select()` は、見出しが見つからなければ「どの見出しが無いか」を示して止まります。`df.get("数量")` のように、無ければ黙って空を返す書き方は避けてください。
- 位置で読む場合（横持ちの表など）は、**目印になる見出しを確かめてから**読みます（`process_sample_c.py` は、1 行目が「店舗名・所在地・業種」であることを確かめています）。
- `try: ... except: pass` で例外を握りつぶさないでください。

例外が出たファイルは取り込まれず、入力フォルダに残り、画面に「要確認」として理由が表示されます。これが正しい動きです。

### 3.5 部品だけを試す

部品は、画面を通さずに単体で試せます。リポジトリのフォルダで、

```
python -c "import sys; sys.path.insert(0, '.'); from suppliers import process_sample_a as p; print(p.load_file(r'sample\input\サンプル卸A\input\202610_サンプル卸A.xlsx'))"
```

返ってきた表を元のファイルと見比べ、行数・最初の行・最後の行・数量の合計が合っているかを確かめます。
**数量の合計を、元のファイルの合計行と比べる**のが、いちばん手軽で確実な検算です。

## 4. マスタを用意する

マスタは 5 つの Excel ファイルです。**シート名と見出しは、下の表のとおりにしてください**（見出しの順番は問いません。ほかの列があってもかまいません）。
サンプル（`sample/masters/`）を開きながら読むと分かりやすいです。

### 4.1 商品マスタ（`product_master.xlsx`）

| シート | 見出し | 内容 |
| --- | --- | --- |
| 変換表 | 元の表記, 標準商品名 | 仕入先の書き方 → 標準の商品名 |
| 標準商品 | 商品名称 | 有効な商品名の一覧 |

- 仕入先の書き方が標準の商品名と**全角/半角・空白だけ**違う場合は、変換表に載せる必要はありません（自動で標準の表記にそろいます）。
- 「標準商品」に無い商品名は、STEP 1 で「要確認」になり、`unmatched_products.xlsx` に出ます。毎月ここに出たものを変換表に足していくのが、マスタの育て方です。

### 4.2 顧客マスタ（`customer_master.xlsx`）

| シート | 見出し |
| --- | --- |
| 顧客 | 得意先名称, 企業名, 担当部署, 担当者 |

- 載っていない得意先は、企業名が「(空白)」、担当部署・担当者が空になります（エラーにはなりません。`unmatched_customers.xlsx` に一覧が出ます）。
- 同じ得意先が複数行あって内容が違うと、上の行が使われ、「要確認」になります。

### 4.3 価格マスタ（`price_master.xlsx`）

| シート | 見出し |
| --- | --- |
| 価格 | 商品名称, 上代, 仕切 |

- 商品名称は、**標準の商品名**で書きます。
- 価格は数値で入れます。「1,200」や全角の数字は数値として読みますが、「要問合せ」のような文字は読めず、「要確認」になります。

### 4.4 都道府県テーブル（`prefecture_table.xlsx`）

| シート | 見出し | 内容 |
| --- | --- | --- |
| 市区町村 | 都道府県, 市区町村 | 住所に都道府県名が無いとき、市区町村名から都道府県を引くための表 |
| 都内区分 | 都内区分 | 東京都の住所から取り出す区・市の名前 |

- 住所に「東京都」「大阪府」などが書かれていれば、それがそのまま使われます。この表が使われるのは、「横浜市中区…」のように都道府県名が省かれた住所だけです。
- **同梱のものは、主な市だけを載せた見本です。** 実運用では、全国の市区町村の一覧（公的機関が公開しています）から作ってください。
- 同じ名前の市区町村が複数の都道府県にある場合は、長い名前から先に照合されることを利用して、区別できる形で登録します。
- この部分は日本の住所に固有です。ほかの国で使う場合は、`steps.py` の `PREFECTURES`、`find_prefecture()`、`KEY_REGION` を自分たちの地域の区分に置き換えます。

### 4.5 ID マスタ（`id_master.xlsx`）

| シート | 必要な見出し | 売上データのどの列で引くか |
| --- | --- | --- |
| 商品マスタ | 商品名称, 商品ID | 商品名称 |
| 都道府県マスタ | 都道府県, 都道府県ID | 都道府県 |
| エリアマスタ | エリア, エリアID | 都内区分 |
| 卸マスタ | 卸名, 卸ID | 卸名 |
| 部署マスタ | 担当部署, 担当部署ID | 担当部署 |
| 担当者マスタ | 担当者, 担当者ID | 担当者 |
| カレンダーマスタ | （自由） | 引かない。BI 用にそのまま出力へ入れる |

- 7 つのシートは、完成ファイルにそのまま複写されます。BI で使いたい属性（商品のカテゴリ、都道府県の地方区分、四半期など）は、ここに列を足しておけば BI 側で使えます。
- ID は文字でも数値でもかまいません。マスタに入っている型のまま出力されます。

### 4.6 マスタ同士のつじつま

マスタは互いに名前でつながっています。食い違うと、どこかが空になります（そして「要確認」で分かります）。

```
商品マスタ「標準商品」の商品名称  =  価格マスタの商品名称  =  ID マスタ「商品マスタ」の商品名称
顧客マスタの担当部署              =  ID マスタ「部署マスタ」の担当部署
顧客マスタの担当者                =  ID マスタ「担当者マスタ」の担当者
変換部品の SUPPLIER_NAME          =  ID マスタ「卸マスタ」の卸名
都道府県テーブルの都内区分        =  ID マスタ「エリアマスタ」のエリア
```

商品を 1 つ増やすときは、「標準商品」「価格」「商品マスタ（ID）」の 3 か所に足します。

## 5. 業務ルールを変える

`pipeline_app/steps.py` の冒頭に、業務ルールを定数としてまとめてあります。

```python
FISCAL_YEAR_START_MONTH = 1          # 年度の開始月。1 なら暦年、4 なら 4 月始まり
DISCONTINUED_MARK = "終売"           # 商品名にこの文字があれば「終売商品」とする
PRODUCT_CLASS_NORMAL, PRODUCT_CLASS_DISCONTINUED = "通常商品", "終売商品"
BLANK_COMPANY = "(空白)"             # 企業名が分からない行に入れる文字
KEY_REGION = "東京都"                # この都道府県だけ、さらに細かい区分 (都内区分) を付ける
```

| 変えたいこと | 変える場所 |
| --- | --- |
| 年度を 4 月始まりにする | `FISCAL_YEAR_START_MONTH = 4` |
| 細かく見たい地域を、東京都以外にする | `KEY_REGION` と、都道府県テーブルの「都内区分」シート、ID マスタの「エリアマスタ」の中身 |
| 売上金額を、仕切ではなく上代で計算する | `final_adjustment()` の「売上金額」の式 |
| 容量の単位を増やす（g、kg など） | `parse_capacity()` |
| 「要確認」にする条件を増やす／減らす | `runner.py` の `_run_step()`（各ステップの結果を見て、文面を組み立てているところ） |
| 照合の「同じ」と見なす範囲を変える（大文字/小文字、株式会社の有無 など） | `normalize_key()`。ここを変えれば、すべての照合に効きます |

定数を変えたら、テスト（[8 章](#8-自分たちのテストを持つ)）で期待値も合わせて直してください。

## 6. 列の構成を変える

出力の列を足す・減らす・名前を変える場合に触る場所は、決まっています。

**例: 仕入先のファイルにある「JAN コード」を、出力まで通したい**

1. `pipeline_app/state.py`
   - `MERGED_COLUMNS` に「JANコード」を足す（STEP 0 の出力）
   - `FINAL_COLUMNS` の、出したい位置に「JANコード」を足す（STEP 4 以降の出力）
   - `GEO_LAYOUT` と `OUTPUT_COLUMNS` は `FINAL_COLUMNS` から作られていますが、**「住所までの 16 列」という切り方**をしているので、住所より前に列を足したら、その数（`[:16]`）も直します
2. `pipeline_app/plugin.py` の `OPTIONAL_COLUMNS` に「JANコード」を足す（変換部品が返してよい列になる）。ひな形の説明も合わせる
3. 変換部品で「JANコード」を返す
4. `runner.py` の STEP 4 の説明文（「17 列に整理」）を合わせる
5. テストの期待値を直す

**例: マスタから引く属性を増やしたい（得意先に「ランク」を付ける）**

1. 顧客マスタに「ランク」列を足す
2. `state.py` の `ACCOUNT_COLUMNS` に「ランク」を足し、`FINAL_COLUMNS` の出したい位置にも足す（住所より前に足したら、上と同じく `[:16]` も直す）

STEP 2 は `ACCOUNT_COLUMNS` に書かれた列を顧客マスタから引くので、処理の本体は直さなくて済みます。

**例: ID を増やしたい**

`steps.py` の `LOOKUPS` に 1 行、`MASTER_SHEETS` にシート名を足し、`state.py` の `ID_COLUMNS` に ID 列の名前を足します。

**列の名前を別の言語にしたい**場合も同じです。名前が書かれているのは、`state.py`（列の一覧）、`plugin.py`（変換部品の列）、`steps.py`（シート名、マスタの見出し、`LOOKUPS`）と、マスタそのものです。

列を変えると、**変える前に作った作業ファイルは「列の構成が想定と違う」と判定されます**。STEP 0 からやり直してください。
これは不具合ではなく、古い形のファイルに新しい処理を掛けないための仕組みです。

## 7. BI ツールに取り込む

完成ファイル `sales_for_bi.xlsx` の各シートは、Excel のテーブルになっています。

| テーブル名 | 中身 |
| --- | --- |
| `Data_売上データ` | 売上の明細（事実テーブル） |
| `Data_商品マスタ` ほか 6 つ | マスタ（次元テーブル） |

BI 側で、売上データの ID とマスタの ID を 1 対多で結びます（商品ID、都道府県ID、エリアID、卸ID、担当部署ID、担当者ID。カレンダーは計上月）。

**月をまたいでデータを蓄積する方法**は、運用に合わせて選んでください。このツール自体は「今月、入力フォルダにあるファイル」を処理します。

- **毎月の完成ファイルを 1 つのフォルダに貯め、BI の「フォルダから取り込む」機能で結合する。**
  STEP 8 が成果物を `backup/<日時>/` に移すので、そこから `sales_for_bi.xlsx` を月の名前を付けて所定のフォルダへコピーします。
- **過去分も含めて、毎回すべてを処理し直す。**
  `processed` の中のファイルを `input` に戻して STEP 0 から実行します。マスタの修正が過去分にも反映されるのが利点です。数万行であれば数分で済みます。

## 8. 自分たちのテストを持つ

`tests/test_pipeline.py` は、同梱のサンプルに対するテストです。自分の組織用に変えたら、**自分たちの期待値**でテストできるようにしてください。

おすすめは、`tools/make_sample_data.py` のやり方をそのまま使うことです。

1. **架空の小さなデータ**を、自分たちの仕入先の形で作る（実データを切り出すのではなく、スクリプトで作る）
2. 同じスクリプトの中で、**処理プログラムを使わずに**期待値を組み立てる
3. 処理の出力と期待値を、セル単位で比べる

実データを使わないのは、機密を守るためだけではありません。
架空のデータなら、「全角の店名」「文字の数量」「都道府県の無い住所」「合計行」といった**確かめたい条件を、狙って入れられる**からです。

仕入先の様式が変わったとき、マスタの持ち方を変えたとき、pandas を更新したとき。
テストがあれば、変更の前後で出力が変わっていないことを数十秒で確かめられます。

## 9. 機密情報を混ぜないために

このツールを組織の外と共有する（公開する、他社に渡す、質問サイトに貼る）ときに、気をつける点です。
販売実績は、取引先・価格・数量という、もっとも機密性の高い情報のかたまりです。

**リポジトリに入れないもの**

- 実際の入力ファイル、マスタ、出力ファイル、ログ（ログには、ファイル名・商品名・得意先名が出ます）
- 自分の環境の設定ファイル（`pipeline_config.json`。フォルダのパスに、利用者名や組織名が含まれることがあります）
- 実在の仕入先の名前が入った変換部品（`SUPPLIER_NAME`、見出しの名前、ファイル名の例）

同梱の `.gitignore` は、ルートの `pipeline_config.json`、`work/`、`data/`、`logs/` を除外しています。
**実データと実際のマスタは、リポジトリのフォルダの外に置く**のがいちばん確実です。

**見落としやすいもの**

| 場所 | 何が残るか |
| --- | --- |
| Excel・Word・PDF のプロパティ | 作成者・最終更新者の氏名、会社名 |
| 画面の画像 | タイトルバーやパスの中の利用者名、ログに出た得意先名・商品名 |
| プログラムの中の文字 | 絶対パス（利用者名・組織名を含む）、コメントに書いた取引先名や担当者名 |
| テスト用のデータ | 実データを「少しだけ」切り出したもの。少量でも実データです |
| ショートカット（.lnk） | 作成した PC のパス |
| バージョン管理の履歴 | 一度入れたファイルは、あとで消しても履歴に残ります |
| コミットの情報 | すべてのコミットに記録される、作者の名前とメールアドレス |

**公開の前に、機械的に確かめる**

「入っていないはず」を目で確かめるのには限界があります。次のような一覧を作り、公開するフォルダの全ファイル（Excel の中身と、画像以外のバイナリも含めて）を検索してください。

- 自社名、グループ会社名、利用者名、社員番号、メールアドレスのドメイン
- 仕入先の名前（全社）
- マスタにある得意先名・商品名・担当者名（全件。マスタから自動で一覧を作って照合する）
- 社内のシステム名、フォルダ名、プロジェクト名

実データと 1 件も一致しないことを確かめてから公開します。
画像の中の文字は、この方法では調べられません。画面の画像は、1 枚ずつ自分の目で確かめてください。
このリポジトリのサンプルが「サンプル卸A」「A食堂」「担当A」のような記号的な名前になっているのは、実在の名前と偶然一致することを避けるためです。

## 10. 向いている場面、向いていない場面

**向いている**

- 報告元が数社〜数十社で、形がそれぞれ違う
- 月に 1 回など、決まった周期でまとめて処理する
- 1 回のデータが数万行まで
- 利用者が、結果を Excel で開いて確かめたい
- マスタを、業務の担当者が Excel で管理している

**向いていない（別の作りを考えたほうがよい）**

- 数百万行を超える、または毎日・毎時の処理が必要 → データベースと定期実行の仕組みへ
- 複数人が同時に操作する → サーバー側で動かす仕組みへ
- 報告元が自動連携（EDI や API）に対応している → 連携を使うほうが確実

とはいえ、[設計思想](design.md) に書いた考え方 ——
**違いは入口で吸収する、知識はマスタに置く、あやしい点は必ず人に見せる、誤った順番では実行させない、保存は書き出してから置き換える**
—— は、道具が変わってもそのまま使えます。

---

<a id="english"></a>

# Adaptation Guide (using it in another organisation)

[日本語](#japanese) | **English**

How to fit this tool to your own wholesalers, master data and business rules.
It assumes one person who can write a little Python.

The column names, sheet names and headers that the code expects are Japanese. They are shown here as they must be typed,
with the English meaning beside them. Renaming them is possible ([E6](#e6-changing-the-columns)) but not required.

- [E1. Overview: what to prepare, what to change](#e1-overview-what-to-prepare-what-to-change)
- [E2. Steps to introduce it](#e2-steps-to-introduce-it)
- [E3. Writing a converter](#e3-writing-a-converter)
- [E4. Preparing the master files](#e4-preparing-the-master-files)
- [E5. Changing the business rules](#e5-changing-the-business-rules)
- [E6. Changing the columns](#e6-changing-the-columns)
- [E7. Loading into a BI tool](#e7-loading-into-a-bi-tool)
- [E8. Keeping your own tests](#e8-keeping-your-own-tests)
- [E9. Keeping confidential data out](#e9-keeping-confidential-data-out)
- [E10. Where it fits and where it does not](#e10-where-it-fits-and-where-it-does-not)

## E1. Overview: what to prepare, what to change

| | What | Where | Effort |
| --- | --- | --- | --- |
| **Must prepare** | A converter per wholesaler | `suppliers/process_*.py` | 10–50 lines each |
| **Must prepare** | Five master files | Anywhere (Excel) | Arrange your existing lists under the fixed headers |
| **Must set** | Locations of folders and masters | The 「設定」 (Settings) tab | Minutes |
| Change if needed | Business rules (first month of the fiscal year, etc.) | Top of `pipeline_app/steps.py` | Edit constants |
| Change if needed | The output columns | `pipeline_app/state.py` and `steps.py` | See [E6](#e6-changing-the-columns) |
| No need to change | GUI, execution control, stage detection, saving | `gui.py` `runner.py` `state.py` | — |

**Replace the sample wholesalers and data with your own.**
Delete `suppliers/process_sample_*.py` (or keep them and untick them on the Settings tab) and only your wholesalers are processed.

## E2. Steps to introduce it

1. **See it work with the sample.** Run it once with `run_sample.bat` and compare the inputs (`sample/input/`), the masters (`sample/masters/`) and the output.
2. **Pick one wholesaler and write its converter** ([E3](#e3-writing-a-converter)). Start with the plainest form.
3. **Prepare the masters** ([E4](#e4-preparing-the-master-files)). Copying the sample masters and replacing their contents is the quickest start.
4. **Start with `run_gui.bat` and fill in the Settings tab.** Create a new empty folder as the work folder.
5. **Run only STEP 0 and look at `sales_work.xlsx`.** Check that the wholesaler's data is in the ten standard columns correctly.
6. **Run STEP 1–7 and clear the "needs review" items one by one.** Whatever the masters lack shows up there.
7. **Compare with an answer made by hand.** Check one month against your existing method: total quantity, total amount, row count, and a few rows in detail.
8. **Add wholesalers one at a time**, repeating 5–7 each time.

Do not load every wholesaler at once. One at a time, the cause of a mismatch is obvious.

## E3. Writing a converter

### E3.1 The contract

Copy `suppliers/_template.py` to `process_<alphanumeric name>.py` and write two things.

```python
from pipeline_app import plugin

SUPPLIER_NAME = "サンプル卸A"          # (1) the wholesaler's name


def load_file(path):                    # (2) read one file and return a table in the standard form
    df = plugin.read_sheet(path, sheet="売上明細", header_row=1)
    return plugin.select(df, {
        "得意先名称": "得意先名",       # standard column name: header in this file
        "商品名称": "商品名",
        "販売本数": "数量",
        "容量": "規格",
        "入数": "ケース入数",
        "住所": "届け先住所",
        "業態": "業種",
    })
```

**(1) `SUPPLIER_NAME`** is used in three places, which must agree:

- the name of the input folder (`<input root>/<SUPPLIER_NAME>/input/`)
- the 「卸名」 (wholesaler) column of the output
- 「卸名」 in the 「卸マスタ」 sheet of the ID master (the key for the wholesaler ID)

**(2) `load_file(path)`** returns a pandas DataFrame with one row per customer × product.

| Column | Required | Meaning |
| --- | --- | --- |
| 得意先名称 | yes | Customer (store) name |
| 商品名称 | yes | Product name, as the wholesaler writes it |
| 販売本数 | yes | Quantity sold (units) |
| 容量 | | Capacity; text such as "720ml" or "1.8L" is fine |
| 入数 | | Units per case |
| 住所 | | Customer address |
| 業態 | | Business type (restaurant, retailer …) |
| 計上月 | | Month (YYYYMM), only if you want to take it from inside the file. Otherwise the first six digits of the file name are used |

Returning a column with any other name is an error for that file (so that a mistyped column name is noticed).

### E3.2 What a converter does, and what it need not do

**It does**

- pick the needed columns and give them the standard names
- **remove rows that are not detail** (totals, subtotals, notes, blank rows)
- turn a cross-tab into one row per sale
- **raise an exception** when the file is not what it expects

**It need not** (the common steps do these for every wholesaler)

- trim spaces or unify full-width and half-width characters
- turn "1,200" into 1200 or "1.8L" into 1800
- replace product or customer names with standard names
- add the source file name, the wholesaler name or the month

A converter fixes the *shape*; fixing the *values* is left to the common steps.
If each converter cleaned values in its own way, the wholesalers would end up cleaned differently.

### E3.3 Common shapes

The three bundled converters serve as examples.

| Shape of the form | Example | Key point |
| --- | --- | --- |
| A plain detail table with headers on row 1 | `process_sample_a.py` | Just `read_sheet` and `select` |
| Title rows above, a total row below | `process_sample_b.py` | `header_row=4` gives the header row; `drop_rows_containing` removes the total |
| A store × product cross-tab | `process_sample_c.py` | Read everything with `header_row=None` and create one row per filled cell |

Other shapes:

| Shape | Approach |
| --- | --- |
| The customer name appears only on the first row of each group | Fill down: `df["得意先名称"] = df["得意先名称"].ffill()` |
| One workbook has a sheet per customer or per month | Read all sheets with `pd.read_excel(path, sheet_name=None)` and `pd.concat` them |
| The address is split into prefecture, city and street columns | Join them as text into 「住所」 |
| Quantity is split into cases and loose units | Convert to units (cases × units per case + loose units) |
| Returns are entered as negative numbers | Pass them through (they cancel in totals). If they must be excluded, exclude them explicitly here |
| The file arrives as csv / xls | Convert it with the pre-processing button first; a converter only ever deals with xlsx |

### E3.4 Never skip silently

Wholesalers change their forms without notice. The worst outcome is **a converter that raises no error and returns an empty or incomplete table.**

- `plugin.select()` stops and names the missing header. Avoid code such as `df.get("数量")` that silently yields nothing.
- When reading by position (a cross-tab), **check a landmark first** (`process_sample_c.py` checks that row 1 starts with the three expected headers).
- Do not swallow exceptions with `try: ... except: pass`.

A file that raises an exception is not merged, stays in the input folder, and is reported on screen as "needs review" with the reason. That is the correct behaviour.

### E3.5 Trying a converter on its own

A converter can be tried without the GUI. In the repository folder:

```
python -c "import sys; sys.path.insert(0, '.'); from suppliers import process_sample_a as p; print(p.load_file(r'sample\input\サンプル卸A\input\202610_サンプル卸A.xlsx'))"
```

Compare the returned table with the original file: row count, first row, last row, and the sum of the quantities.
**Comparing the quantity sum with the total row of the original file** is the easiest reliable check.

## E4. Preparing the master files

There are five Excel files. **Use the sheet names and headers exactly as below** (header order does not matter, and extra columns are allowed).
It helps to have the samples in `sample/masters/` open while reading.

### E4.1 Product master (`product_master.xlsx`)

| Sheet | Headers | Contents |
| --- | --- | --- |
| 変換表 (conversion table) | 元の表記 (original spelling), 標準商品名 (standard product name) | Wholesaler's spelling → standard product name |
| 標準商品 (standard products) | 商品名称 (product name) | The list of valid product names |

- A spelling that differs from the standard name **only in full-width/half-width characters or spaces** need not be listed; it is aligned automatically.
- A name not in 「標準商品」 becomes "needs review" at STEP 1 and is written to `unmatched_products.xlsx`. Adding those to the conversion table each month is how the master grows.

### E4.2 Customer master (`customer_master.xlsx`)

| Sheet | Headers |
| --- | --- |
| 顧客 (customers) | 得意先名称 (customer name), 企業名 (company), 担当部署 (department), 担当者 (person in charge) |

- A customer not listed gets 「(空白)」 as company and blank department and person (not an error; a list is written to `unmatched_customers.xlsx`).
- If the same customer appears on several rows with different contents, the upper row is used and the step shows "needs review".

### E4.3 Price master (`price_master.xlsx`)

| Sheet | Headers |
| --- | --- |
| 価格 (prices) | 商品名称 (product name), 上代 (list price), 仕切 (wholesale price) |

- Write product names in their **standard** form.
- Prices must be numbers. "1,200" and full-width digits are read as numbers; text such as "ask us" is not, and becomes "needs review".

### E4.4 Prefecture table (`prefecture_table.xlsx`)

| Sheet | Headers | Contents |
| --- | --- | --- |
| 市区町村 (municipalities) | 都道府県 (prefecture), 市区町村 (municipality) | Used to find the prefecture from a municipality name when the address omits the prefecture |
| 都内区分 (Tokyo divisions) | 都内区分 | Names of wards and cities to extract from Tokyo addresses |

- If the address contains a prefecture name, that is used as it is. The table is consulted only for addresses that omit it.
- **The bundled table lists only major cities as an example.** For real use, build it from a nationwide list of municipalities (published by public bodies).
- Where the same municipality name exists in more than one prefecture, register it in a distinguishable form; longer names are matched first.
- This part is specific to Japanese addresses. For another country, replace `PREFECTURES`, `find_prefecture()` and `KEY_REGION` in `steps.py` with your own regions.

### E4.5 ID master (`id_master.xlsx`)

| Sheet | Required headers | Looked up by this sales column |
| --- | --- | --- |
| 商品マスタ (products) | 商品名称, 商品ID | 商品名称 (product name) |
| 都道府県マスタ (prefectures) | 都道府県, 都道府県ID | 都道府県 (prefecture) |
| エリアマスタ (areas) | エリア, エリアID | 都内区分 (Tokyo ward/city) |
| 卸マスタ (wholesalers) | 卸名, 卸ID | 卸名 (wholesaler) |
| 部署マスタ (departments) | 担当部署, 担当部署ID | 担当部署 (department) |
| 担当者マスタ (people) | 担当者, 担当者ID | 担当者 (person in charge) |
| カレンダーマスタ (calendar) | (free) | Not looked up; copied to the output for the BI tool |

- The seven sheets are copied to the finished file as they are. Add any attributes you want in the BI tool (product category, region of a prefecture, quarter …) as extra columns here.
- IDs may be text or numbers; they are output in the type they have in the master.

### E4.6 Consistency between the masters

The masters are linked to each other by name. Where they disagree, something ends up blank (and "needs review" tells you).

```
product name in 「標準商品」  =  product name in the price master  =  product name in 「商品マスタ」 of the ID master
department in the customer master  =  department in 「部署マスタ」
person in the customer master      =  person in 「担当者マスタ」
SUPPLIER_NAME of the converter     =  wholesaler in 「卸マスタ」
「都内区分」 in the prefecture table  =  area in 「エリアマスタ」
```

To add one product, add it in three places: 「標準商品」, the price master and 「商品マスタ」 of the ID master.

## E5. Changing the business rules

The business rules are constants at the top of `pipeline_app/steps.py`.

```python
FISCAL_YEAR_START_MONTH = 1          # first month of the fiscal year: 1 = calendar year, 4 = April start
DISCONTINUED_MARK = "終売"           # a product whose name contains this is classed as discontinued
PRODUCT_CLASS_NORMAL, PRODUCT_CLASS_DISCONTINUED = "通常商品", "終売商品"
BLANK_COMPANY = "(空白)"             # written where the company is unknown
KEY_REGION = "東京都"                # the one prefecture that gets a finer division
```

| To change | Change this |
| --- | --- |
| Fiscal year starting in April | `FISCAL_YEAR_START_MONTH = 4` |
| A different region to see in fine detail | `KEY_REGION`, plus the contents of the 「都内区分」 sheet and of 「エリアマスタ」 |
| Sales amount based on list price instead of wholesale price | The 「売上金額」 expression in `final_adjustment()` |
| More units of capacity (g, kg …) | `parse_capacity()` |
| More or fewer conditions for "needs review" | `_run_step()` in `runner.py`, where each step's result is turned into a message |
| What counts as "the same" when matching (letter case, company suffixes …) | `normalize_key()`; a change there applies to every lookup |
| The labels written into the data (「通常商品」, 「(空白)」 …) | The constants above |

After changing a constant, update the expected values of the tests ([E8](#e8-keeping-your-own-tests)).

## E6. Changing the columns

The places to touch when adding, removing or renaming an output column are fixed.

**Example: pass a barcode column from the wholesalers' files through to the output**

1. `pipeline_app/state.py`
   - add the column to `MERGED_COLUMNS` (output of STEP 0)
   - add it to `FINAL_COLUMNS` at the position you want (output from STEP 4 on)
   - `GEO_LAYOUT` and `OUTPUT_COLUMNS` are built from `FINAL_COLUMNS` by slicing **"the first 16 columns, up to the address"**; if you add a column before the address, change that number (`[:16]`) too
2. `pipeline_app/plugin.py`: add the column to `OPTIONAL_COLUMNS` (converters may then return it), and mention it in the template
3. Return the column from the converters
4. Update the description of STEP 4 in `runner.py` ("17 columns")
5. Update the expected values of the tests

**Example: add another attribute from a master (a rank for each customer)**

1. Add the column to the customer master
2. Add it to `ACCOUNT_COLUMNS` in `state.py`, and to `FINAL_COLUMNS` at the position you want (if that is before the address, change `[:16]` as above)

STEP 2 looks up whatever `ACCOUNT_COLUMNS` lists, so the processing itself needs no change.

**Example: add another ID**

Add a line to `LOOKUPS` and the sheet name to `MASTER_SHEETS` in `steps.py`, and the ID column name to `ID_COLUMNS` in `state.py`.

**Renaming columns to another language** works the same way: the names appear in `state.py` (column lists), `plugin.py` (converter columns), `steps.py` (sheet names, master headers, `LOOKUPS`) and in the masters themselves.

After a change, **a work file created before the change is judged "unexpected column layout".** Start again from STEP 0.
This is not a defect; it is the mechanism that keeps new processing away from files in the old shape.

## E7. Loading into a BI tool

Every sheet of `sales_for_bi.xlsx` is an Excel table.

| Table name | Contents |
| --- | --- |
| `Data_売上データ` | Sales detail (the fact table) |
| `Data_商品マスタ` and six others | Masters (dimension tables) |

In the BI tool, relate the IDs in the sales data to the IDs in the masters, one-to-many
(product, prefecture, area, wholesaler, department, person; the calendar by month).

**How to accumulate data across months** is up to you. The tool itself processes "the files in the input folders this month".

- **Collect each month's finished file in one folder and combine them with the BI tool's "load from folder" feature.**
  STEP 8 moves the outputs to `backup/<date and time>/`; copy `sales_for_bi.xlsx` from there to the collection folder under a name that includes the month.
- **Reprocess everything every time.**
  Move the files in `processed` back to `input` and run from STEP 0. Corrections to the masters then apply to past months too. For tens of thousands of rows this takes minutes.

## E8. Keeping your own tests

`tests/test_pipeline.py` tests the bundled sample. Once you have adapted the tool, test against **your own expected values.**

The recommended way is the one `tools/make_sample_data.py` uses.

1. Generate **small fictional data** in the shapes of your wholesalers (by script, not by cutting out real data)
2. In the same script, build the expected values **without calling the processing code**
3. Compare the output with the expected values cell by cell

Fictional data is not only for confidentiality.
With fictional data you can **plant exactly the conditions you want to check**: a store name in full-width characters, a quantity entered as text, an address without a prefecture, a total row.

When a wholesaler changes its form, when the masters are restructured, when pandas is upgraded —
with tests, you can confirm in under a minute that the output has not changed.

## E9. Keeping confidential data out

Points to watch when sharing this tool outside your organisation (publishing it, handing it to another company, pasting it into a question site).
Sales reports are a concentration of the most sensitive information there is: trading partners, prices and quantities.

**Do not put these in the repository**

- real input files, masters, output files or logs (logs contain file names, product names and customer names)
- your own settings file (`pipeline_config.json`; folder paths can contain a user name or an organisation name)
- converters that carry the names of real wholesalers (`SUPPLIER_NAME`, header names, example file names)

The bundled `.gitignore` excludes `pipeline_config.json`, `work/`, `data/` and `logs/` at the root.
**Keeping real data and real masters outside the repository folder** is the surest way.

**Easily overlooked**

| Place | What remains |
| --- | --- |
| Properties of Excel, Word and PDF files | Names of the author and of the last editor, company name |
| Screenshots | A user name in the title bar or in a path; customer and product names in the log |
| Text inside programs | Absolute paths (with user or organisation names); partner or staff names in comments |
| Test data | "Just a small extract" of real data is still real data |
| Shortcuts (.lnk) | Paths of the PC they were created on |
| Version-control history | A file that was once committed stays in the history after it is deleted |
| Commit metadata | The author name and e-mail address recorded in every commit |

**Check mechanically before publishing**

Looking for "what should not be there" by eye has limits. Build a list like the following and search every file in the folder to be published (including the contents of Excel files and of binaries other than images).

- your company name, group company names, user names, employee numbers, the domain of your e-mail addresses
- the names of all your wholesalers
- every customer, product and staff name in the masters (generate the list from the masters and match automatically)
- names of internal systems, folders and projects

Publish only after confirming that nothing matches the real data.
Images cannot be searched this way; look at every screenshot yourself.
The sample in this repository uses symbolic names for wholesalers, stores and staff so that they cannot coincide with real ones.

## E10. Where it fits and where it does not

**It fits when**

- there are a handful to a few dozen report sources, each with a different form
- processing happens in batches on a fixed cycle, such as monthly
- one run is up to tens of thousands of rows
- users want to open the result in Excel and check it
- business staff maintain the master data in Excel

**It does not fit (consider a different design) when**

- there are millions of rows, or processing is needed daily or hourly → a database and scheduled jobs
- several people operate it at the same time → a server-side system
- the report sources support automated exchange (EDI or an API) → use that instead

Even so, the ideas in the [design document](design.md#english) —
**absorb differences at the entrance, keep knowledge in master data, always show a person what is doubtful, refuse to run in a wrong order, save by writing and then replacing**
— hold whatever tools you use.
