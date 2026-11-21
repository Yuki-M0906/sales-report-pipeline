# -*- coding: utf-8 -*-
"""処理全体の検証。

    python -m unittest discover -s tests -v

サンプルデータを一時フォルダに作り、画面を使わずに全ステップを実行して、
出力を期待値 (tools/make_sample_data.py が処理プログラムとは別に組み立てたもの) と
セル単位 (値と型) で比べる。あわせて、誤った操作を止める仕組みが働くことを確かめる。
"""

import contextlib
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

import make_sample_data                                   # noqa: E402
from pipeline_app import config, runner, state, steps     # noqa: E402


def rows_of(path, sheet):
    book = openpyxl.load_workbook(path)
    try:
        return [list(row) for row in book[sheet].iter_rows(values_only=True)]
    finally:
        book.close()


class PipelineCase(unittest.TestCase):
    with_issues = False

    def setUp(self):
        self.base = Path(tempfile.mkdtemp(prefix="pipeline_test_"))
        self.addCleanup(shutil.rmtree, self.base, ignore_errors=True)
        with contextlib.redirect_stdout(io.StringIO()):
            make_sample_data.main(["--out", str(self.base)] + (["--with-issues"] if self.with_issues else []))
        cfg = json.loads((self.base / "pipeline_config.json").read_text(encoding="utf-8"))
        for key, value in cfg.items():
            if isinstance(value, str) and key != "input_subdir":
                cfg[key] = os.path.normpath(self.base / value)
        self.cfg = cfg
        self.work = Path(cfg["work_dir"])
        self.suppliers = config.discover_suppliers()

    def run_steps(self, step_ids, confirm=True):
        """戻り値: (最後まで実行できたか, {ステップ: 状態}, 実行ログ)"""
        statuses, log = {}, io.StringIO()
        with contextlib.redirect_stdout(log):
            ok = runner.run_steps(step_ids, self.cfg, self.suppliers,
                                  on_status=lambda step, status: statuses.__setitem__(step, status),
                                  confirm=lambda _message: confirm)
        return ok, statuses, log.getvalue()

    def assert_output_matches_expected(self):
        actual = rows_of(self.work / state.OUTPUT_FILE, state.DATA_SHEET)
        expected = rows_of(self.base / "expected" / "expected_output.xlsx", state.DATA_SHEET)
        self.assertEqual(len(actual), len(expected), "行数が違います")
        for number, (a, e) in enumerate(zip(actual, expected), start=1):
            self.assertEqual(a, e, f"{number} 行目の値が違います")
            self.assertEqual([type(v) for v in a], [type(v) for v in e], f"{number} 行目の型が違います")


class CleanSampleTest(PipelineCase):
    def test_all_steps_produce_expected_output(self):
        ok, statuses, log = self.run_steps(range(8))
        self.assertTrue(ok, log)
        self.assertEqual(statuses, {step: "done" for step in range(8)})
        self.assert_output_matches_expected()

    def test_output_has_tables_and_master_sheets(self):
        self.run_steps(range(8))
        book = openpyxl.load_workbook(self.work / state.OUTPUT_FILE)
        self.assertEqual(book.sheetnames, [state.DATA_SHEET] + steps.MASTER_SHEETS)
        self.assertEqual(rows_of(self.work / state.OUTPUT_FILE, state.DATA_SHEET)[0], state.OUTPUT_COLUMNS)
        for sheet in book.worksheets:
            self.assertEqual(list(sheet.tables), [f"Data_{sheet.title}"])
        self.assertEqual(book[state.DATA_SHEET].tables[f"Data_{state.DATA_SHEET}"].ref, "A1:Y53")

    def test_one_step_at_a_time_gives_the_same_result(self):
        for step in range(8):
            ok, statuses, log = self.run_steps([step])
            self.assertTrue(ok, log)
        self.assert_output_matches_expected()

    def test_repeating_steps_2_and_3_does_not_change_the_result(self):
        self.run_steps([0, 1, 2, 3])
        ok, _statuses, log = self.run_steps([2, 3, 2, 3])
        self.assertTrue(ok, log)
        self.run_steps([4, 5, 6, 7])
        self.assert_output_matches_expected()

    def test_wrong_order_is_refused_before_anything_runs(self):
        self.assertTrue(any("STEP 0" in m for m in runner.precheck([1], self.cfg)))       # 作業ファイルが無い
        self.run_steps([0])
        self.assertTrue(any("STEP 1" in m for m in runner.precheck([3], self.cfg)))       # 商品名の標準化がまだ
        self.assertTrue(any("STEP 2" in m for m in runner.precheck([1, 4], self.cfg)))    # 顧客・価格がまだ
        self.assertEqual(runner.precheck([1, 2, 3, 4, 5, 6, 7], self.cfg), [])

    def test_a_finished_file_cannot_be_processed_twice(self):
        self.run_steps(range(5))
        before = rows_of(self.work / state.WORK_FILE, state.DATA_SHEET)
        for step in (1, 2, 3, 4):
            ok, statuses, _log = self.run_steps([step])
            self.assertFalse(ok)
            self.assertEqual(statuses[step], "error")
        self.assertEqual(rows_of(self.work / state.WORK_FILE, state.DATA_SHEET), before)   # ファイルは変わらない
        self.run_steps([5])
        ok, statuses, _log = self.run_steps([5])
        self.assertTrue(ok)                                                               # STEP 5 はやり直せる
        self.run_steps([6, 7])
        self.assert_output_matches_expected()

    def test_a_replaced_work_file_is_not_trusted(self):
        """記録と違うファイルに差し替えられたら、「STEP 1 は済んでいる」とは見なさない。"""
        self.run_steps([0, 1])
        self.assertIs(state.actual_state(self.work)["step1"], True)
        book = openpyxl.load_workbook(self.work / state.WORK_FILE)
        book.save(self.work / state.WORK_FILE)                    # 中身は同じだが、別のファイルとして保存し直す
        self.assertIsNone(state.actual_state(self.work)["step1"])

    def test_failed_save_keeps_the_previous_file(self):
        self.run_steps([0, 1, 2, 3])
        before = (self.work / state.WORK_FILE).read_bytes()
        original = steps._save_sheets

        def broken(sheets, path):
            def save(tmp):
                Path(tmp).write_bytes(b"half written")
                raise OSError("disk full")
            steps._save_with_retry(save, path)

        steps._save_sheets = broken
        try:
            ok, statuses, _log = self.run_steps([4])
        finally:
            steps._save_sheets = original
        self.assertFalse(ok)
        self.assertEqual((self.work / state.WORK_FILE).read_bytes(), before)
        self.assertEqual([p.name for p in self.work.glob("*.saving.*")], [])

    def test_step8_moves_only_what_was_merged(self):
        stray = Path(config.input_dir(self.cfg, self.suppliers[0])) / "メモ.txt"
        stray.write_text("memo", encoding="utf-8")
        self.run_steps(range(8))
        late = Path(config.input_dir(self.cfg, self.suppliers[0])) / "202611_サンプル卸A.xlsx"
        shutil.copy(late.with_name("202610_サンプル卸A.xlsx"), late)      # 統合のあとに届いたファイル
        ok, statuses, log = self.run_steps([8])
        self.assertTrue(ok, log)
        self.assertEqual(statuses[8], "warn")                             # 残したファイルがあることを知らせる
        self.assertTrue(late.exists() and stray.exists())
        for supplier in self.suppliers:
            self.assertEqual(sorted(p.name for p in Path(config.input_dir(self.cfg, supplier)).glob("*.xlsx")),
                             ["202611_サンプル卸A.xlsx"] if supplier is self.suppliers[0] else [])
        self.assertEqual(len(list(Path(self.cfg["input_root"]).glob("*/processed/*.xlsx"))), 4)
        backups = list((self.work / "backup").iterdir())
        self.assertEqual(len(backups), 1)
        self.assertTrue((backups[0] / state.OUTPUT_FILE).exists())
        self.assertFalse((self.work / state.OUTPUT_FILE).exists())
        self.assertFalse((self.work / state.STATE_FILE).exists())

    def test_step8_without_a_record_moves_no_input(self):
        """どのファイルを取り込んだのか分からないときは、入力ファイルを動かさない。"""
        self.run_steps(range(8))
        (self.work / state.STATE_FILE).unlink()
        ok, statuses, log = self.run_steps([8])
        self.assertTrue(ok, log)
        self.assertEqual(statuses[8], "warn")
        self.assertEqual(len(list(Path(self.cfg["input_root"]).glob("*/input/*.xlsx"))), 4)
        self.assertEqual(list(Path(self.cfg["input_root"]).glob("*/processed/*")), [])

    def test_step8_needs_the_input_folder(self):
        self.run_steps(range(8))
        self.cfg["input_root"] = str(self.base / "no_such_folder")
        self.assertTrue(any("入力データのルートフォルダ" in m for m in runner.precheck([8], self.cfg)))
        self.assertTrue((self.work / state.OUTPUT_FILE).exists())

    def test_rerunning_step1_makes_prices_be_added_again(self):
        """商品名を直したら、前に付けた価格は外れ、STEP 3 をやり直すまで先へ進めない。"""
        self.run_steps([0, 1, 2, 3])
        self.assertTrue(any("STEP 3" in m for m in runner.precheck([1, 4], self.cfg)))
        ok, _statuses, log = self.run_steps([1])
        self.assertTrue(ok, log)
        self.assertEqual(state.actual_state(self.work)["work"], {"layout": "merged", "accounts": True, "prices": False})
        self.assertTrue(any("STEP 3" in m for m in runner.precheck([4], self.cfg)))
        ok, _statuses, log = self.run_steps([3, 4, 5, 6, 7])
        self.assertTrue(ok, log)
        self.assert_output_matches_expected()

    def test_a_list_that_cannot_be_saved_leaves_the_work_file_alone(self):
        self.run_steps([0])
        before = (self.work / state.WORK_FILE).read_bytes()
        original = steps._save_list

        def broken(rows, columns, path):
            raise steps.StepError("一覧を保存できません")

        steps._save_list = broken
        try:
            ok, statuses, _log = self.run_steps([1])
        finally:
            steps._save_list = original
        self.assertEqual((ok, statuses[1]), (False, "error"))
        self.assertEqual((self.work / state.WORK_FILE).read_bytes(), before)
        self.assertIs(state.actual_state(self.work)["step1"], False)

    def test_a_broken_file_is_reported_and_left_in_place(self):
        folder = Path(config.input_dir(self.cfg, self.suppliers[0]))
        make_sample_data.save_book(folder / "202611_サンプル卸A.xlsx", {"売上明細": [["得意先", "商品"], ["x", "y"]]})
        ok, statuses, log = self.run_steps([0], confirm=True)
        self.assertEqual(statuses[0], "warn")
        self.assertIn("ファイルに次の見出しがありません", log)
        self.assertNotIn("202611_サンプル卸A.xlsx", state.merged_files(self.work)["サンプル卸A"])
        self.run_steps([1, 2, 3, 4, 5, 6, 7])
        self.assert_output_matches_expected()                             # 読めたファイルの分は正しく処理される

    def test_csv_is_converted_and_originals_are_kept(self):
        demo = self.base / "convert_demo"
        with contextlib.redirect_stdout(io.StringIO()):
            result = steps.convert_folder(demo)
        self.assertEqual((result["ok"], result["failed"]), (1, []))
        self.assertTrue((demo / steps.ORIGINALS_FOLDER / "202611_サンプル卸C.csv").exists())
        converted = demo / "202611_サンプル卸C.xlsx"
        module = steps._load_module(self.suppliers[2].script, "test_supplier_c")
        frame = module.load_file(str(converted))
        self.assertEqual(len(frame), 8)
        with contextlib.redirect_stdout(io.StringIO()):
            again = steps.convert_folder(demo)                            # もう一度実行しても何も起きない
        self.assertEqual((again["ok"], again["failed"]), (0, []))


class SampleWithIssuesTest(PipelineCase):
    with_issues = True

    def test_problems_are_flagged_not_hidden(self):
        ok, statuses, log = self.run_steps(range(8))
        self.assertTrue(ok, log)
        self.assertEqual(statuses, {0: "done", 1: "warn", 2: "done", 3: "warn", 4: "warn", 5: "warn", 6: "warn", 7: "done"})
        self.assertIn("季節限定 ゆずサイダー 350ml", log)
        self.assertIn("見本郡見本町99", log)
        self.assert_output_matches_expected()
        unmatched = rows_of(self.work / steps.UNMATCHED_PRODUCTS, "一覧")
        self.assertEqual(unmatched, [["商品名称", "行数", "卸名"], ["季節限定 ゆずサイダー 350ml", 1, "サンプル卸A"]])

    def test_answering_no_stops_the_run(self):
        ok, statuses, _log = self.run_steps(range(8), confirm=False)
        self.assertFalse(ok)
        self.assertEqual(statuses[1], "warn")
        self.assertEqual({statuses[step] for step in range(2, 8)}, {"skipped"})


class MasterProblemTest(PipelineCase):
    def test_non_numeric_price_is_reported_and_not_written(self):
        path = Path(self.cfg["price_master"])
        book = openpyxl.load_workbook(path)
        book["価格"]["C2"] = "要問合せ"          # りんごジュース 1L の仕切
        book["価格"]["B3"] = "４８０"            # 全角の数字は数値として読む
        book.save(path)
        ok, statuses, log = self.run_steps(range(8))
        self.assertTrue(ok, log)
        self.assertEqual(statuses[3], "warn")
        self.assertIn("りんごジュース 1L: 仕切 = 要問合せ", log)
        header, *rows = rows_of(self.work / state.OUTPUT_FILE, state.DATA_SHEET)
        name, price, cost, amount = (header.index(c) for c in ("商品名称", "上代", "仕切", "売上金額"))
        for row in rows:
            if row[name] == "りんごジュース 1L":
                self.assertEqual((row[cost], row[amount]), (None, None))
            elif row[name] == "みかんジュース 1L":
                self.assertEqual(row[price], 480)
            else:
                self.assertIsInstance(row[amount], (int, float))

    def test_duplicate_ids_are_reported(self):
        path = Path(self.cfg["id_master"])
        book = openpyxl.load_workbook(path)
        book["商品マスタ"].append(["PXXX", "りんごジュース 1L", "飲料"])
        book.save(path)
        ok, statuses, log = self.run_steps(range(8))
        self.assertTrue(ok, log)
        self.assertEqual(statuses[6], "warn")
        self.assertIn("同じ値に別々の ID が登録されています", log)
        self.assert_output_matches_expected()                 # 上の行 (P001) が使われる

    def test_sheet_without_data_rows_gets_no_table(self):
        """データ行の無いテーブルを作ると Excel が開けないファイルになるので、作らずに知らせる。"""
        path = Path(self.cfg["id_master"])
        book = openpyxl.load_workbook(path)
        book["カレンダーマスタ"].delete_rows(2, 100)
        book.save(path)
        ok, statuses, log = self.run_steps(range(8))
        self.assertTrue(ok, log)
        self.assertEqual(statuses[7], "warn")
        out = openpyxl.load_workbook(self.work / state.OUTPUT_FILE)
        self.assertEqual(list(out["カレンダーマスタ"].tables), [])
        self.assertEqual(list(out[state.DATA_SHEET].tables), [f"Data_{state.DATA_SHEET}"])
        ok, statuses, _log = self.run_steps([7])               # もう一度実行しても同じ
        self.assertEqual(statuses[7], "warn")

    def test_missing_master_sheet_is_a_clear_error(self):
        path = Path(self.cfg["id_master"])
        book = openpyxl.load_workbook(path)
        del book["卸マスタ"]
        book.save(path)
        ok, statuses, log = self.run_steps(range(8))
        self.assertFalse(ok)
        self.assertEqual((statuses[6], statuses[7]), ("error", "skipped"))
        self.assertIn("ID マスタに必要なシートがありません: 卸マスタ", log)
        self.assertFalse((self.work / state.OUTPUT_FILE).exists())


class HelperTest(unittest.TestCase):
    def test_parse_capacity(self):
        cases = {"720ml": 720, "1.8L": 1800, "１Ｌ": 1000, "500 ML": 500, "360cc": 360, 300: 300, "2リットル": 2000,
                 "1,000ml": 1000, "0.5l": 500, "大瓶": None, "ml": None}
        for value, expected in cases.items():
            self.assertEqual(steps.parse_capacity(value), expected, value)

    def test_as_number(self):
        cases = {"1,200": 1200, "１２": 12, " 7 ": 7, "2.5": 2.5, 12.0: 12, "未定": None, "": None, None: None, "nan": None}
        for value, expected in cases.items():
            self.assertEqual(steps.as_number(value), expected, value)
            self.assertIs(type(steps.as_number(value)), type(expected), value)

    def test_normalize_key(self):
        self.assertEqual(steps.normalize_key(" Ｃスーパー　川崎店 "), "Cスーパー 川崎店")
        self.assertEqual(steps.normalize_key(None), "")
        self.assertEqual(steps.normalize_key(float("nan")), "")

    def test_fiscal_year(self):
        self.assertEqual(steps.fiscal_year(202603), 2026)
        original = steps.FISCAL_YEAR_START_MONTH
        steps.FISCAL_YEAR_START_MONTH = 4
        try:
            self.assertEqual((steps.fiscal_year(202603), steps.fiscal_year(202604)), (2025, 2026))
        finally:
            steps.FISCAL_YEAR_START_MONTH = original

    def test_find_prefecture(self):
        cities = [("京都市", "京都府"), ("横浜市", "神奈川県")]
        self.assertEqual(steps.find_prefecture("東京都港区見本町1", cities), "東京都")       # 「京都」に引きずられない
        self.assertEqual(steps.find_prefecture("横浜市中区見本町", cities), "神奈川県")
        self.assertIsNone(steps.find_prefecture("見本郡見本町", cities))
        self.assertIsNone(steps.find_prefecture(None, cities))

    def test_every_script_compiles(self):
        """テストが読み込まないファイル (起動スクリプトなど) も、文法の誤りが無いことを確かめる。"""
        for path in sorted(ROOT.rglob("*.py")):
            compile(path.read_text(encoding="utf-8"), str(path), "exec")

    def test_sample_generator_refuses_a_folder_that_is_not_a_sample(self):
        base = Path(tempfile.mkdtemp(prefix="pipeline_test_"))
        self.addCleanup(shutil.rmtree, base, ignore_errors=True)
        precious = base / "input" / "real_supplier" / "202610_real.xlsx"
        precious.parent.mkdir(parents=True)
        precious.write_text("real data", encoding="utf-8")
        with self.assertRaises(SystemExit), contextlib.redirect_stderr(io.StringIO()):
            make_sample_data.main(["--out", str(base)])
        self.assertTrue(precious.exists())

    def test_only_properly_named_scripts_are_plugins(self):
        names = [s.script.name for s in config.discover_suppliers()]
        self.assertEqual(names, ["process_sample_a.py", "process_sample_b.py", "process_sample_c.py"])


if __name__ == "__main__":
    unittest.main()
