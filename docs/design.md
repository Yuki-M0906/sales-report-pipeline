# 設計思想 / Design

<a id="japanese"></a>

**日本語** | [English](#english)

このツールが何のためにあり、どういう考え方で作られているかをまとめます。
コードを読む前に、あるいは自分の組織向けに作り変える前に読んでください。

- [1. 目的](#1-目的)
- [2. 技術の核](#2-技術の核)
- [3. データの流れ](#3-データの流れ)
- [4. 設計の原則](#4-設計の原則)
- [5. 「要確認」という状態](#5-要確認という状態)
- [6. 誤った操作を止める仕組み](#6-誤った操作を止める仕組み)
- [7. 構成（どこに何が書いてあるか）](#7-構成どこに何が書いてあるか)
- [8. あえてやらないこと](#8-あえてやらないこと)
- [9. 起きやすい問題と備え](#9-起きやすい問題と備え)

---

## 1. 目的

メーカーや販売元は、自社の商品が「どの店に・何が・何本売れたか」を、卸（仕入先）から届く販売実績で知ります。
ところが、その報告は卸ごとに形が違います。

- 見出しの名前が違う（「得意先名」「店舗」「店舗名」…）
- 表の上に表題や対象月の行がある。末尾に合計行がある
- 1 行 1 明細の表もあれば、店舗 × 商品の集計表（横持ち）もある
- 同じ商品・同じ店でも、卸によって書き方が違う
- 価格・自社の担当者・地域といった、分析に必要な情報は入っていない

BI ツールが必要とするのは、その反対のものです。
**全社が同じ列を持ち、同じものは同じ名前で書かれ、マスタと結べる ID が付いた、1 枚の表**。

このツールの目的は、毎月その変換を **速く・同じ手順で・間違いに気づける形で** 行うことです。
最終的な成果物は、BI に取り込むための標準形式のファイル（`sales_for_bi.xlsx`、25 列 + マスタのシート）です。

## 2. 技術の核

核になっている考え方は 3 つです。

### 2.1 違いは入口で吸収し、あとは一本道にする

卸ごとの違いを、**卸 1 社につき 1 つの小さな変換部品**（`suppliers/process_*.py`）に閉じ込めます。
部品の仕事は「その卸のファイルを読んで、標準形式の表にして返す」ことだけです。

```
卸A のファイル ─→ 変換部品 A ─┐
卸B のファイル ─→ 変換部品 B ─┼─→ 標準形式（全社同じ列）─→ 共通の処理（STEP 1〜7）─→ 完成ファイル
卸C のファイル ─→ 変換部品 C ─┘
```

標準形式に合流したあとは、どの卸のデータも同じ処理を通ります。
だから、卸が 1 社増えても書くのは部品 1 つ（数十行）で、共通の処理には手を入れません。
逆に、共通の処理を直せば全社に同時に効きます。

### 2.2 マスタと照合して「意味」をそろえる

形をそろえただけでは分析できません。同じものを同じ名前にし、足りない情報を補います。

| 何をそろえる／補う | 使うマスタ | ステップ |
| --- | --- | --- |
| 商品名を標準の名前にする | 商品マスタ（変換表・標準商品） | STEP 1 |
| 得意先に、企業名・担当部署・担当者を付ける | 顧客マスタ | STEP 2 |
| 商品に、上代・仕切を付ける | 価格マスタ | STEP 3 |
| 住所から、都道府県と都内区分を取り出す | 都道府県テーブル | STEP 5 |
| 商品・地域・卸・部署・担当者に ID を付ける | ID マスタ | STEP 6 |

知識（どの表記がどの商品か、どの店を誰が担当しているか）は**プログラムではなくマスタ（Excel）に置きます**。
業務を知っている人が、プログラムに触れずに直せるようにするためです。

### 2.3 「間違った出力を作らない」ことを仕組みで守る

データの加工で本当に怖いのは、エラーで止まることではありません。
**エラーにならないまま、間違った数字ができあがること**です。そこで次の 3 つを仕組みにしています。

1. **あやしい点は必ず人に見せる（「要確認」）** — マスタに無い商品、数値でない数量、読めなかったファイルなどを、黙って捨てたり補ったりしません。（[5 章](#5-要確認という状態)）
2. **実行できない順番では、実行させない** — ファイルの見出しから「いまどの段階か」を判定し、想定と違う段階のファイルには処理を掛けません。（[6 章](#6-誤った操作を止める仕組み)）
3. **保存の途中で失敗しても、ファイルを壊さない** — 一時ファイルに書き出してから置き換えます。

## 3. データの流れ

作業ファイル `sales_work.xlsx` を STEP 0〜5 が順に加工し、STEP 6 が完成ファイル `sales_for_bi.xlsx` を作ります。

| ステップ | すること | 終わったあとの列 |
| --- | --- | --- |
| STEP 0 データ統合 | 各卸のファイルを変換部品で標準形式にし、1 つにまとめる | 10 列: 元ファイル, 計上月, 卸名, 得意先名称, 商品名称, 容量, 入数, 販売本数, 住所, 業態 |
| STEP 1 商品名の標準化 | 商品名を標準の名前に直す。マスタに無い商品名を一覧にする | （列は同じ） |
| STEP 2 顧客情報の付与 | 得意先名称で顧客マスタを引く | + 企業名, 担当部署, 担当者 |
| STEP 3 価格の付与 | 商品名称で価格マスタを引く | + 上代, 仕切 |
| STEP 4 最終調整 | 数値の列を数値にそろえる。売上金額・年度・商品区分を計算。列を並べ直す | 17 列 |
| STEP 5 都道府県の抽出 | 住所から都道府県・都内区分を取り出す | 19 列（+ 都道府県, 都内区分） |
| STEP 6 ID の付与 | 6 種類の ID を付け、マスタのシートと一緒に完成ファイルに保存 | 25 列（+ ID 6 列） |
| STEP 7 テーブル設定 | 全シートを Excel のテーブル（`Data_<シート名>`）にする | （列は同じ） |
| STEP 8 後処理 | 入力ファイルを `processed` へ、成果物を `backup` へ移す | — |

完成ファイルの 25 列:

```
卸名, 計上月, 担当部署, 担当者, 得意先名称, 企業名, 商品名称, 容量, 入数, 上代,
販売本数, 仕切, 売上金額, 年度, 商品区分, 住所, 都道府県, 都内区分,
商品ID, 都道府県ID, エリアID, 卸ID, 担当部署ID, 担当者ID, 業態
```

列の定義は `pipeline_app/state.py` の 1 か所にまとめてあります。

## 4. 設計の原則

### 原則 1: 黙って捨てない、黙って補わない

想定と違うデータに出会ったとき、プログラムが勝手に「たぶんこうだろう」と処理すると、間違いが出力に紛れ込みます。
このツールは、判断できないものは**空のままにして、その事実を報告**します。

- 数値として読めない価格 → 書き込まない。どの商品の何という値かを表示する
- マスタに無い商品 → 価格も ID も付けない。一覧をファイルに出す
- 見出しが変わって読めないファイル → そのファイルだけ取り込まない。入力フォルダに残す

### 原則 2: 列は「位置」ではなく「名前」で扱う

「左から 5 列目が商品名」という前提で書かれた処理は、列が 1 つ増えただけで、**エラーにならないまま別の列を加工**します。
このツールは、すべての列を見出しの名前で扱い、必要な見出しが無ければその場で止まります。

### 原則 3: 照合に使うキーは、1 か所で正規化する

「Ｃスーパー　川崎店」と「Cスーパー 川崎店」は、人が見れば同じ店ですが、文字としては別物です。
照合のたびに個別の対処をすると、対処が漏れた箇所で「マスタにあるのに一致しない」が起きます。

`normalize_key()` という 1 つの関数が、全角/半角の違い・前後の空白・連続する空白をならします。
商品名・得意先名・ID のキーなど、**マスタとデータを比べるところは必ずこの関数を通します**。
元のデータは書き換えず、比べるときだけ使います（一致したら、マスタ側の表記にそろえます）。

### 原則 4: 何度実行しても、同じ結果になる

STEP 2（顧客情報）と STEP 3（価格）は、付ける列を毎回作り直します。
マスタを直してもう一度実行すれば、直した内容がそのまま反映されます。「前回の結果に上書き」が積み重なって、状態が分からなくなることがありません。

一方、STEP 1（商品名の書き換え）や STEP 4（列の並べ直し）は、やり直すと元の情報が失われる処理です。
これらは「やり直すなら STEP 0 から」と決め、途中からの再実行は仕組みで止めます（[6 章](#6-誤った操作を止める仕組み)）。

### 原則 5: 保存は「書き出してから置き換える」

Excel ファイルの保存は、行数が多いと数十秒かかります。その途中で失敗すると、壊れたファイルが残ります。
すべての保存は、同じフォルダの一時ファイルに書き出し、成功してから本来の名前に置き換えます。
途中で失敗しても、元のファイルは実行前のまま残ります。

### 原則 6: 元のデータを消さない

- 入力ファイルは削除せず、`processed` フォルダへ**移動**する（同名があれば、上書きせず日時を付ける）
- 変換前の csv / xls は削除せず、`converted_originals` フォルダへ移動する
- 成果物は `backup/<日時>/` に残す

あとから「あの月の元データを見たい」「変換が正しかったか確かめたい」に応えられるようにするためです。

### 原則 7: 変わるものを、変わる理由ごとに分けて置く

| 変わるもの | 置き場所 | 変える人 |
| --- | --- | --- |
| 卸ごとの帳票の形 | `suppliers/process_*.py` | 変換部品を書く人 |
| 商品・得意先・価格・担当・ID などの知識 | マスタ（Excel） | 業務の担当者 |
| フォルダやマスタの場所 | 設定ファイル（画面の「設定」タブ） | 利用者 |
| 年度の開始月などの業務ルール | `steps.py` 冒頭の定数 | 導入する人 |
| 出力の列の構成 | `state.py` の列の定義 | 導入する人 |

### 原則 8: 画面と処理を分ける

処理（`steps.py`、`runner.py`）は画面を知りません。進み具合は `print` するだけで、画面側がそれを実行ログ欄に流します。
このおかげで、同じ処理を画面なしで呼び出して、自動テストで検証できます。

### 原則 9: 期待値と突き合わせて検証する

「動いた」と「正しい」は別です。テストでは、出力のすべてのセルを、値と型まで期待値と比べます。
期待値は処理プログラムを通さず、サンプルの元になる事実から別のコードで組み立てています
（同じ間違いが期待値にも入ってしまうのを避けるため）。
さらに、処理をわざと壊してテストが失敗することを確かめ、テストが形だけになっていないことを確認しています。

## 5. 「要確認」という状態

各ステップの結果は、次のどれかで表示されます。

| 表示 | 意味 |
| --- | --- |
| 完了 | 問題なく終わった |
| **要確認** | **最後まで実行したが、人が確かめるべき点がある** |
| エラー | 続行できないため止まった（作業ファイルと完成ファイルは、そのステップを実行する前のまま） |
| 未実行 | 前のステップが止まった、または中止されたため、実行していない |

「要確認」は、このツールでいちばん大事な状態です。
多くのツールは「成功」か「失敗」しか持ちませんが、実際のデータ処理の大半は**「できたが、気になる点がある」**です。
それを成功に丸めると問題が埋もれ、失敗に丸めると毎回止まって使いものになりません。

要確認になると、内容がログに「△ 確認」として表示され、次のステップへ進むかどうかを尋ねます。
進めても止めてもかまいません。判断するのは、内容を知っている人です。

何を要確認にするか:

| ステップ | 要確認になるとき |
| --- | --- |
| STEP 0 | 取り込めなかったファイルがある／入力ファイルの無い卸がある／ファイル名から計上月を読めない／商品名や数量が全行で空 |
| STEP 1 | 商品マスタに無い商品名がある／変換表で同じ表記に別々の標準名がある |
| STEP 2 | 顧客マスタに同じ得意先が複数あり、内容が食い違っている |
| STEP 3 | 数値でない価格がある／同じ商品に違う価格がある／仕切が付かなかった商品がある |
| STEP 4 | 容量・入数・販売本数・価格に、数値として読めない値がある |
| STEP 5 | 住所は書いてあるのに、都道府県が分からない |
| STEP 6 | ID マスタに無いため、ID を付けられなかった値がある／ID マスタで同じ値に別々の ID が登録されている |
| STEP 7 | データ行の無いシートがあり、テーブルにしなかった |
| STEP 8 | 移動先に同名のファイルがあった／今回の統合に含まれないファイルを入力フォルダに残した／どのファイルを取り込んだかの記録が無く、入力ファイルを移動しなかった |

どこまでを要確認にするかは、業務の判断です。
たとえばこの実装では、顧客マスタに無い得意先は要確認にしていません（新しい店は毎月現れるもので、企業名を「(空白)」として集計できるため。一覧はファイルに出します）。
自分の組織では止めて確かめたい、という場合は `runner.py` の該当箇所に数行足すだけで変えられます。

## 6. 誤った操作を止める仕組み

各ステップは、前のステップが作ったファイルを受け取って加工します。
順番を飛ばしたり、同じステップを 2 回掛けたりすると、エラーにならないまま誤った出力ができることがあります。

そこで、**ファイルそのものから「いまどの段階か」を判定**します。

- 作業ファイルの見出し行を読み、列の構成から段階を判定する
  （10 列なら STEP 0 の直後、企業名があれば STEP 2 済み、17 列なら STEP 4 済み、19 列なら STEP 5 済み）
- 見出しからは分からない情報（STEP 1 を実行したか、STEP 0 でどのファイルを取り込んだか）は、作業フォルダの `pipeline_state.json` に記録する
- 記録には作業ファイルの大きさと更新時刻を一緒に残し、ファイルが差し替えられていたら記録を信用しない

実行ボタンを押すと、選んだステップを最初から順に進めた場合を見通して、実行できない組み合わせなら**何も実行せずに**理由を表示します。

| しようとしたこと | 結果 |
| --- | --- |
| STEP 0 を実行せずに STEP 1 | 止める（作業ファイルが無い） |
| STEP 1 を飛ばして STEP 3 | 止める（商品名が標準化されていないと価格が付かない） |
| STEP 2・3 を飛ばして STEP 4 | 止める（企業名・価格の無い完成ファイルができてしまう） |
| STEP 4 が済んだファイルに、もう一度 STEP 1〜4 | 止める（「やり直すなら STEP 0 から」） |
| STEP 2 や STEP 3 を、STEP 4 の前にもう一度 | 実行できる（結果は作り直される） |
| STEP 3 のあとで、STEP 1 をもう一度 | 実行できるが、付いていた価格を外す（価格は商品名で引いているため）。STEP 3 をやり直すまで STEP 4 へ進めない |
| STEP 8 で、統合のあとに届いたファイル | 移動しない（取り込んだファイルだけを移す） |
| STEP 8 で、どのファイルを取り込んだかの記録が無い | 入力ファイルを 1 つも移動せず、そのことを知らせる |

ファイルの名前や「前回どこまで実行したかの記憶」ではなく、**ファイルの中身で判定する**ことが要点です。
ファイルを手で差し替えたり、別の月のファイルを置いたりしても、判定は中身に従います。

## 7. 構成（どこに何が書いてあるか）

```
pipeline_app/
  state.py    列の定義 / 段階の判定 / 記録ファイル          ← 「正しい形」の定義はここだけ
  steps.py    各ステップの処理（ファイルを受け取り、結果を辞書で返す）
  runner.py   ステップの一覧 / 実行前チェック / 順番どおりの実行 / 「要確認」の文面
  gui.py      画面（処理は別スレッドで動かし、キューでやり取りする）
  config.py   設定ファイル / 変換部品の検出
  plugin.py   変換部品から使う小さな道具
suppliers/
  process_*.py  卸ごとの変換部品
```

依存の向きは一方向です: `gui → runner → steps → state`。下の層は上の層を知りません。

画面は、処理を別のスレッドで動かします。処理スレッドは画面の部品に直接触れず、
ログ・状態の変化・確認の問い合わせをキューに入れ、画面側が 0.1 秒ごとに取り出して反映します。
確認ダイアログは、処理スレッドが答えを待つ形で画面側が表示します。

## 8. あえてやらないこと

小さく保つために、次のことはしていません。

- **データベースを使わない** — 入力も出力もマスタも Excel です。利用者が中身を自分で開いて確かめられることを優先しています。
- **自動で取りに行かない** — メールやサイトからのファイル取得はしません。ファイルを置くのは人です。
- **帳票の形を自動で推測しない** — 形の解釈は変換部品に明示的に書きます。推測は外れたときに気づけないためです。
- **同時に複数人で使わない** — 1 台の PC で、1 人が、月に 1 回動かす想定です。
- **大量データを想定しない** — 数万行までを想定しています（すべてメモリに読み込みます）。数百万行なら別の作りが必要です。

## 9. 起きやすい問題と備え

販売実績のようなデータを Excel のファイルで受け渡しする処理で起きやすい問題と、それへの備えです。この設計の各所は、これらを前提に作ってあります。

- **合計行・小計行の扱いは、卸ごとに違う。**
  「小計」と書かれた行が、明細そのものである帳票もあります。取り除く行の条件は、共通の処理ではなく変換部品に、卸ごとに書きます。
- **単位の付いた数値は、数字だけ抜き出してはいけない。**
  「1.8L」から数字だけを取ると 1.8 に、「720ml」は 720 になり、同じ列で単位が混ざります。単位を解釈して 1 つの単位にそろえます。
- **1 つでも文字が混ざった数値の列は、列全体が文字になる。**
  価格の列に「要問合せ」が 1 件あるだけで、掛け算が文字の繰り返しになることがあります。数値の列は、計算の前に 1 セルずつ数値として読み、読めないものは報告します。
- **マスタには重複がある。**
  同じ得意先が 2 行あり、担当者が食い違っている、ということは普通に起きます。どちらを使うかを決め（このツールでは上の行）、食い違いを報告します。
- **「空のセルにだけ書く」処理は、やり直しに弱い。**
  マスタを直して再実行しても、前回の値が残っていて反映されません。付ける列は毎回作り直すほうが安全です。
- **同じ月を 2 回処理してしまうことがある。**
  処理済みファイルの移動先に同名のファイルがあったら、上書きせずに別名で残し、そのことを知らせます。
- **データ行の無いテーブルは、Excel が開けないファイルになる。**
  見出しだけのシートをテーブルにすると、保存はできても Excel が開けません。作ったファイルは、作った道具ではなく、使う道具（Excel）で開けることを確かめます。
- **クラウド同期のフォルダは、ふつうのフォルダと少し違う。**
  同期ソフトがフォルダに「読み取り専用」の印を付けたり、一時的にファイルをつかんだりします。保存や削除は、やり直しと分かりやすいメッセージを前提に書きます。
- **スクリプトのコピーが、処理に紛れ込む。**
  フォルダ内のスクリプトを自動で読み込む仕組みは、「〜 - コピー.py」も読み込みます。名前の形を決め、同じ卸の部品が 2 つあれば片方しか使いません。
- **作った人の PC でしか動かない原因は、たいていパスとライブラリ。**
  プログラムに書かれた絶対パスは設定ファイルへ。必要なライブラリは起動時に確かめ、足りなければ入れ方を表示します。
- **ダブルクリックで起動したプログラムは、エラーを表示する場所が無い。**
  起動に失敗したら、内容をファイルに残してメッセージ ボックスで知らせます。画面のログは、ワンクリックでコピーできるようにします（問い合わせに貼ってもらうため）。

---

<a id="english"></a>

# Design

[日本語](#japanese) | **English**

This document explains what the tool is for and the thinking behind it.
Read it before reading the code or adapting the tool to your organisation.

- [E1. Purpose](#e1-purpose)
- [E2. The core ideas](#e2-the-core-ideas)
- [E3. Data flow](#e3-data-flow)
- [E4. Design principles](#e4-design-principles)
- [E5. The "needs review" state](#e5-the-needs-review-state)
- [E6. Stopping wrong operations](#e6-stopping-wrong-operations)
- [E7. Structure of the code](#e7-structure-of-the-code)
- [E8. What it deliberately does not do](#e8-what-it-deliberately-does-not-do)
- [E9. Common pitfalls and safeguards](#e9-common-pitfalls-and-safeguards)

## E1. Purpose

A manufacturer or a distributor learns "which store sold what, and how many" from the sales reports its wholesalers send.
Those reports all look different.

- The header names differ (the customer column may be called by three different names)
- There are title rows above the table and a total row below it
- Some are detail tables with one row per sale; others are store × product cross-tabs
- The same product or store is written differently by each wholesaler
- What analysis needs — prices, the company's own person in charge, region — is not in the report

A BI tool needs the opposite:
**one table in which every wholesaler has the same columns, the same thing has the same name, and every row carries IDs that link to master tables.**

The purpose of this tool is to do that conversion every month **quickly, by the same procedure, and in a way that makes mistakes visible.**
The final product is a file in a standard form for the BI tool (`sales_for_bi.xlsx`: 25 columns plus the master sheets).

## E2. The core ideas

### E2.1 Absorb the differences at the entrance; one road after that

The differences between wholesalers are confined to **one small converter per wholesaler** (`suppliers/process_*.py`).
A converter does one thing: read that wholesaler's file and return a table in the standard form.

```
Wholesaler A file ─→ converter A ─┐
Wholesaler B file ─→ converter B ─┼─→ standard form (same columns) ─→ common steps (STEP 1–7) ─→ finished file
Wholesaler C file ─→ converter C ─┘
```

After the merge, every wholesaler's data goes through the same processing.
Adding a wholesaler means writing one converter (a few dozen lines); the common steps are not touched.
Fixing a common step fixes it for every wholesaler at once.

### E2.2 Align the meaning by matching against master data

Aligning the shape is not enough. The same thing must get the same name, and the missing information must be added.

| What is aligned or added | Master used | Step |
| --- | --- | --- |
| Product names become standard names | Product master (conversion table, standard products) | STEP 1 |
| Company, department and person in charge are added to each customer | Customer master | STEP 2 |
| List price and wholesale price are added to each product | Price master | STEP 3 |
| Prefecture and Tokyo ward/city are extracted from the address | Prefecture table | STEP 5 |
| IDs are added for product, region, wholesaler, department and person | ID master | STEP 6 |

Knowledge — which spelling means which product, who is in charge of which store — lives **in the master files (Excel), not in the program**,
so that the people who know the business can correct it without touching code.

### E2.3 Protect against "wrong output" by mechanism

The real danger in data processing is not that it stops with an error.
It is that **it finishes without an error and the numbers are wrong.** Three mechanisms address this.

1. **Anything doubtful is always shown to a person ("needs review").** Unknown products, non-numeric quantities and unreadable files are never silently dropped or filled in. ([E5](#e5-the-needs-review-state))
2. **A step cannot run in an order in which it must not run.** The stage of a file is determined from its header row, and a step refuses a file at an unexpected stage. ([E6](#e6-stopping-wrong-operations))
3. **A failed save does not corrupt a file.** Everything is written to a temporary file and then swapped in.

## E3. Data flow

STEP 0–5 process the work file `sales_work.xlsx` in turn; STEP 6 creates the finished file `sales_for_bi.xlsx`.
The column names are Japanese; the English meaning of each is listed in the [user guide](user-guide.md#english).

| Step | What it does | Columns afterwards |
| --- | --- | --- |
| STEP 0 Merge | Converts each wholesaler's files to the standard form and merges them | 10 columns: source file, month, wholesaler, customer, product, capacity, units per case, quantity, address, business type |
| STEP 1 Standardise product names | Replaces product names with standard names; lists names missing from the master | (same) |
| STEP 2 Add customer information | Looks up the customer master by customer name | + company, department, person in charge |
| STEP 3 Add prices | Looks up the price master by product name | + list price, wholesale price |
| STEP 4 Final adjustment | Makes numeric columns numeric; calculates sales amount, fiscal year and product class; reorders columns | 17 columns |
| STEP 5 Extract prefecture | Extracts prefecture and Tokyo ward/city from the address | 19 columns |
| STEP 6 Add IDs | Adds six IDs and saves the finished file together with the master sheets | 25 columns |
| STEP 7 Create tables | Turns every sheet into an Excel table (`Data_<sheet name>`) | (same) |
| STEP 8 Clean up | Moves the input files to `processed` and the outputs to `backup` | — |

All column definitions are in one place: `pipeline_app/state.py`.

## E4. Design principles

### Principle 1: Never drop silently, never fill in silently

When the program meets data it does not expect and processes it with a guess, the mistake ends up in the output.
This tool **leaves blank what it cannot decide and reports the fact.**

- A price that is not a number → not written; the product and the value are shown
- A product missing from the master → no price, no ID; a list is written to a file
- A file whose headers have changed → that file is not merged; it stays in the input folder

### Principle 2: Handle columns by name, not by position

Code written as "the fifth column from the left is the product name" keeps running when a column is added — on the wrong column.
Every column here is addressed by its header name, and processing stops immediately if a required header is missing.

### Principle 3: Normalise matching keys in one place

To a person, a store name written in full-width characters and the same name in half-width characters are the same store. As text they are different.
If each lookup deals with this on its own, one of them will be forgotten and "it is in the master but does not match" will happen.

One function, `normalize_key()`, evens out full-width/half-width differences, leading and trailing spaces and repeated spaces.
**Every comparison between master data and report data goes through it** — product names, customer names, ID keys.
The original data is not rewritten; the function is used only for comparison (on a match, the spelling in the master is adopted).

### Principle 4: Running a step again gives the same result

STEP 2 (customer information) and STEP 3 (prices) rebuild their columns every time.
Correct the master and run the step again, and the correction is reflected. Results never pile up on top of earlier results.

STEP 1 (rewriting product names) and STEP 4 (reordering columns) lose the original information when repeated.
For these the rule is "start again from STEP 0", and repeating them from the middle is blocked ([E6](#e6-stopping-wrong-operations)).

### Principle 5: Save by "write, then replace"

Saving a large Excel file can take tens of seconds, and a failure in the middle leaves a broken file.
Every save writes to a temporary file in the same folder and renames it only after it has succeeded.
If it fails, the previous file is still there, untouched.

### Principle 6: Never delete the original data

- Input files are not deleted; they are **moved** to a `processed` folder (never overwritten — a timestamp is added if the name exists)
- The csv / xls files that were converted are moved to `converted_originals`, not deleted
- Outputs are kept under `backup/<date and time>/`

This makes it possible to answer later questions such as "what did that month's original look like?" or "was the conversion right?".

### Principle 7: Put each thing that changes where its reason for changing lives

| What changes | Where it lives | Who changes it |
| --- | --- | --- |
| The shape of each wholesaler's form | `suppliers/process_*.py` | Whoever writes converters |
| Knowledge about products, customers, prices, staff, IDs | Master files (Excel) | Business staff |
| Locations of folders and master files | Settings file (the Settings tab) | The user |
| Business rules such as the first month of the fiscal year | Constants at the top of `steps.py` | Whoever introduces the tool |
| The output columns | Column definitions in `state.py` | Whoever introduces the tool |

### Principle 8: Keep the GUI and the processing apart

The processing (`steps.py`, `runner.py`) knows nothing about the GUI. It only prints its progress, and the GUI shows that in the log pane.
Because of this the same processing can be called without a GUI and verified by automated tests.

### Principle 9: Verify against expected values

"It ran" and "it is right" are different things. The tests compare every output cell, value and type, with the expected values.
The expected values are built by separate code from the facts behind the sample, not by the processing code
(so that the same mistake cannot enter both).
The processing was also broken on purpose to confirm that the tests fail — that the tests are not merely decorative.

## E5. The "needs review" state

Each step ends in one of these states.

| On screen | Meaning |
| --- | --- |
| 完了 (Done) | Finished without problems |
| **要確認 (Needs review)** | **Ran to the end, but there is something a person should check** |
| エラー (Error) | Stopped because it could not continue (the work file and the finished file are as they were before the step) |
| 未実行 (Not run) | Not executed because an earlier step stopped or the run was cancelled |

"Needs review" is the most important state in this tool.
Most tools have only success and failure, but most real data processing ends as **"done, with something that looks odd"**.
Rounding that up to success buries the problem; rounding it down to failure stops the work every month.

When a step needs review, the details appear in the log and the user is asked whether to continue to the next step.
Either answer is acceptable. The decision belongs to the person who knows the data.

| Step | Becomes "needs review" when |
| --- | --- |
| STEP 0 | A file could not be merged / a wholesaler has no input file / the month cannot be read from a file name / product names or quantities are blank in every row |
| STEP 1 | A product name is missing from the product master / the conversion table maps one spelling to two standard names |
| STEP 2 | The customer master has the same customer twice with different contents |
| STEP 3 | A price is not a number / one product has two different prices / a product got no wholesale price |
| STEP 4 | Capacity, units per case, quantity or price contains a value that cannot be read as a number |
| STEP 5 | An address is present but its prefecture cannot be determined |
| STEP 6 | A value is missing from the ID master / the ID master has two different IDs for the same value |
| STEP 7 | A sheet has no data rows and was not made into a table |
| STEP 8 | A file of the same name already existed at the destination / files not part of this run were left in the input folder / there is no record of which files were merged, so no input file was moved |

Where to draw the line is a business decision.
In this implementation a customer missing from the customer master is not flagged (new stores appear every month, and they can still be aggregated under a blank-company label; a list is written to a file).
If your organisation wants to stop and check, a few lines in `runner.py` change that.

## E6. Stopping wrong operations

Each step takes the file produced by the previous step.
Skipping a step or applying one twice can produce a wrong output without any error.

So the tool **decides the current stage from the file itself.**

- It reads the header row of the work file and determines the stage from the set of columns
  (10 columns: just after STEP 0; a company column: STEP 2 done; 17 columns: STEP 4 done; 19 columns: STEP 5 done)
- What the header cannot tell (whether STEP 1 has run, which files STEP 0 merged) is recorded in `pipeline_state.json` in the work folder
- The record stores the size and modification time of the work file; if the file has been replaced, the record is not trusted

When the Run button is pressed, the tool looks ahead through the selected steps in order and, if the combination cannot run, shows the reason **without running anything.**

| Attempt | Result |
| --- | --- |
| STEP 1 without STEP 0 | Blocked (no work file) |
| STEP 3 skipping STEP 1 | Blocked (prices cannot be matched to unstandardised names) |
| STEP 4 skipping STEP 2 or 3 | Blocked (the finished file would lack company or prices) |
| STEP 1–4 again on a file that has finished STEP 4 | Blocked ("start again from STEP 0") |
| STEP 2 or STEP 3 again before STEP 4 | Allowed (the result is rebuilt) |
| STEP 1 again after STEP 3 | Allowed, but the prices are removed (they were looked up by product name); STEP 4 is blocked until STEP 3 is run again |
| STEP 8 with a file that arrived after the merge | Not moved (only merged files are moved) |
| STEP 8 with no record of which files were merged | No input file is moved, and the user is told |

The point is that the decision rests on **the contents of the file**, not on its name or on a memory of what was run last time.
If a file is swapped by hand, the decision follows the contents.

## E7. Structure of the code

```
pipeline_app/
  state.py    column definitions / stage detection / record file      ← the only definition of "the right shape"
  steps.py    the steps (each takes file paths and returns a dictionary)
  runner.py   step list / pre-run check / ordered execution / wording of "needs review"
  gui.py      the GUI (processing runs in a worker thread and talks through a queue)
  config.py   settings file / discovery of converters
  plugin.py   small helpers for converters
suppliers/
  process_*.py  one converter per wholesaler
```

Dependencies point one way: `gui → runner → steps → state`. A lower layer knows nothing about the layers above it.

The GUI runs the processing in a separate thread. The worker never touches a widget;
it puts log text, status changes and confirmation requests on a queue, and the GUI takes them off every 0.1 second.
A confirmation dialog is shown by the GUI while the worker waits for the answer.

## E8. What it deliberately does not do

- **No database.** Inputs, outputs and master data are Excel files, so users can open and check them.
- **No automatic collection.** It does not fetch files from mail or web sites. A person places the files.
- **No guessing of layouts.** The interpretation of a form is written explicitly in a converter; a wrong guess would go unnoticed.
- **No concurrent use.** One person, one PC, once a month.
- **No large data.** Up to tens of thousands of rows (everything is loaded into memory). Millions of rows need a different design.

## E9. Common pitfalls and safeguards

Problems that tend to occur when sales data is exchanged as Excel files, and how to guard against them. The design above is built with these in mind.

- **How total and subtotal rows should be treated differs by wholesaler.**
  In some forms a row labelled "subtotal" is the detail itself. The rule for dropping rows belongs in each converter, not in the common steps.
- **Do not strip the unit from a number.**
  Taking only the digits turns "1.8L" into 1.8 and "720ml" into 720 — two units in one column. Interpret the unit and convert to one unit.
- **One text value turns a whole numeric column into text.**
  A single "ask us" in a price column can make multiplication repeat a string. Read numeric columns cell by cell before calculating, and report what cannot be read.
- **Master data contains duplicates.**
  The same customer on two rows with different people in charge is normal. Decide which one wins (here: the upper row) and report the conflict.
- **"Write only into empty cells" does not survive a re-run.**
  After the master is corrected, the old value is still there. Rebuilding the added columns every time is safer.
- **The same month gets processed twice.**
  If a file of the same name is already in the archive folder, keep both under different names and say so.
- **A table without data rows makes a file Excel cannot open.**
  The save succeeds, but Excel refuses the file. Check the output with the tool that will use it (Excel), not only with the tool that wrote it.
- **A cloud-synchronised folder is not quite an ordinary folder.**
  The sync client may mark folders read-only or hold a file briefly. Write saves and deletions with retries and clear messages.
- **Copies of scripts slip into the processing.**
  A mechanism that loads every script in a folder also loads "… - copy.py". Fix the naming pattern, and use only one converter per wholesaler.
- **"Works only on the author's PC" is usually paths and libraries.**
  Move absolute paths into a settings file. Check required libraries at start-up and show how to install what is missing.
- **A program started by double-click has nowhere to show an error.**
  If start-up fails, write the details to a file and show a message box. Make the on-screen log copyable in one click, so that users can paste it into a support request.
