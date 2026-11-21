# -*- coding: utf-8 -*-
"""ステップの定義と、順番どおりの実行の制御 (画面には依存しない)。"""

import os
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List

from . import config, state


@dataclass
class StepDef:
    id: int
    title: str
    detail: str
    default: bool = True                             # 既定で実行対象にするか
    needs: List[str] = field(default_factory=list)   # 必要な設定キー (ファイル/フォルダ)


STEPS = [
    StepDef(0, "データ統合", f"各仕入先のファイル → {state.WORK_FILE}", needs=["input_root"]),
    StepDef(1, "商品名の標準化", "商品名をそろえる / マスタに無い商品名を一覧に", needs=["product_master"]),
    StepDef(2, "顧客情報の付与", "企業名・担当部署・担当者を付ける", needs=["customer_master"]),
    StepDef(3, "価格の付与", "上代・仕切を付ける", needs=["price_master"]),
    StepDef(4, "最終調整", "売上金額・年度・商品区分を加え、17 列に整理"),
    StepDef(5, "都道府県の抽出", "住所 → 都道府県・都内区分", needs=["prefecture_table"]),
    StepDef(6, "ID の付与", f"ID 6 列を付けて {state.OUTPUT_FILE} を作る", needs=["id_master"]),
    StepDef(7, "テーブル設定", "全シートを Excel のテーブルにする"),
    StepDef(8, "後処理 (移動・バックアップ)", "入力ファイルを processed へ、成果物を backup へ", default=False,
            needs=["input_root"]),
]
STEP_BY_ID = {s.id: s for s in STEPS}

FIELD_LABELS = {key: label for key, label, _ in config.SETTING_FIELDS}


def _lines(items, limit=8):
    """確認ダイアログに出す箇条書き (多すぎるときは省略し、全件は実行ログで見てもらう)。"""
    items = list(items)
    shown = ["・" + str(item) for item in items[:limit]]
    if len(items) > limit:
        shown.append(f"・ほか {len(items) - limit} 件 (実行ログを確認してください)")
    return "\n".join(shown)


def _run_step(step_id, cfg, suppliers):
    """1 ステップを実行する。戻り値は「利用者に確認してほしい内容」(無ければ None)。"""
    from . import steps          # pandas を使うので、必要になってから読み込む

    work = Path(cfg["work_dir"])
    work_file = work / state.WORK_FILE
    output_file = work / state.OUTPUT_FILE
    before = state.signature(work_file)
    notes = []

    if step_id == 0:
        result = steps.merge_suppliers([(s, config.input_dir(cfg, s)) for s in suppliers], str(work_file))
        state.record_step(work, 0, merged_files=result["merged_files"])
        if result["errors"]:
            notes.append("1 件も取り込めなかった仕入先: " + "、".join(result["errors"]))
        if result["empty"]:
            notes.append("入力ファイルが無かった仕入先: " + "、".join(result["empty"]))
        if result["issues"]:
            notes.append("入力ファイルで確認が必要な点:\n" + _lines(result["issues"]))

    elif step_id == 1:
        result = steps.standardize_products(str(work_file), cfg["product_master"], str(work / steps.UNMATCHED_PRODUCTS))
        state.record_step(work, 1, signature_before=before)
        if result["unmatched"]:
            notes.append(f"商品マスタに無い商品名が {len(result['unmatched'])} 件あります "
                         f"({steps.UNMATCHED_PRODUCTS} に出力しました)。\n" + _lines(result["unmatched"])
                         + "\nこのまま進めると、これらの商品には価格と ID が付きません。"
                           "商品マスタに追加して STEP 0 からやり直すのが確実です。")
        if result["conflicts"]:
            notes.append("商品マスタの変換表で、同じ表記に別々の標準商品名が登録されています (上の行を使いました)。\n"
                         + _lines(result["conflicts"]))

    elif step_id == 2:
        result = steps.add_customer_info(str(work_file), cfg["customer_master"], str(work / steps.UNMATCHED_CUSTOMERS))
        state.record_step(work, 2, signature_before=before)
        if result["conflicts"]:
            notes.append("顧客マスタに同じ得意先が複数あり、内容が食い違っています (上の行を使いました)。\n"
                         + _lines(result["conflicts"]))

    elif step_id == 3:
        result = steps.add_prices(str(work_file), cfg["price_master"])
        state.record_step(work, 3, signature_before=before)
        if result["invalid"]:
            notes.append(f"価格マスタに数値でない価格が {len(result['invalid'])} 件あります (その価格は空のままにしました)。\n"
                         + _lines(f"{name}: {column} = {raw}" for name, column, raw in result["invalid"]))
        if result["conflicts"]:
            notes.append("価格マスタに同じ商品が複数あり、価格が食い違っています (上の行を使いました)。\n"
                         + _lines(result["conflicts"]))
        if result["unpriced"]:
            notes.append(f"仕切が付かなかった商品が {len(result['unpriced'])} 件あります (売上金額が空になります)。\n"
                         + _lines(result["unpriced"]))

    elif step_id == 4:
        result = steps.final_adjustment(str(work_file))
        state.record_step(work, 4, signature_before=before)
        if result["non_numeric"]:
            notes.append("数値として読めない値が入っていた列があります (その行の値は空にしました)。\n"
                         + _lines(f"{column}: {count} 行" for column, count in result["non_numeric"].items()))

    elif step_id == 5:
        result = steps.extract_prefecture(str(work_file), cfg["prefecture_table"])
        state.record_step(work, 5, signature_before=before)
        if result["not_found"]:
            notes.append(f"都道府県が分からなかった住所が {len(result['not_found'])} 件あります。"
                         "都道府県テーブルの「市区町村」シートに追加すると読み取れるようになります。\n"
                         + _lines(result["not_found"]))

    elif step_id == 6:
        result = steps.add_ids(str(work_file), cfg["id_master"], str(output_file))
        if result["missing"]:
            notes.append("ID マスタに無いため、ID を付けられなかった値があります。\n"
                         + _lines(f"{column}: {len(keys)} 件 ({'、'.join(keys[:3])}{' ほか' if len(keys) > 3 else ''})"
                                  for column, keys in result["missing"].items()))
        if result["conflicts"]:
            notes.append("ID マスタで、同じ値に別々の ID が登録されています (上の行を使いました)。\n"
                         + _lines(f"{sheet}: {'、'.join(keys[:3])}{' ほか' if len(keys) > 3 else ''}"
                                  for sheet, keys in result["conflicts"].items()))

    elif step_id == 7:
        result = steps.apply_tables(str(output_file))
        if result["skipped"]:
            notes.append("データ行が無いため、テーブルにしなかったシートがあります (BI からテーブル名で参照できません)。\n"
                         + _lines(result["skipped"]))

    elif step_id == 8:
        archived = steps.archive_inputs(
            [(s.name, config.input_dir(cfg, s), config.processed_dir(cfg, s)) for s in suppliers],
            merged_files=state.merged_files(work))
        backup = steps.backup_outputs(work)
        failed = archived["failed"] + backup["failed"]
        if failed:
            raise steps.StepError(
                f"移動できなかったファイルが {len(failed)} 件あります。\n" + _lines(failed)
                + "\nExcel で開いていないか確認してから、もう一度 STEP 8 を実行してください。")
        state.clear_state(work)
        if archived["unknown"]:
            notes.append("STEP 0 の記録が無く、どのファイルを取り込んだのか分からないため、入力ファイルは移動していません。"
                         "入力フォルダを確かめ、処理が済んだファイルは手で processed フォルダへ移してください。")
        if archived["renamed"]:
            notes.append("processed フォルダに同名のファイルがあったため、別名で保存しました "
                         "(同じ月を 2 回処理していないか確認してください)。\n" + _lines(archived["renamed"]))
        if archived["left"]:
            notes.append("今回の統合に含まれていないため、入力フォルダに残したファイルがあります。\n"
                         + _lines(archived["left"]))
    return "\n".join(notes) or None


def settings_problem(step_id, cfg):
    """設定 (フォルダ・マスタファイル) の問題。無ければ None。"""
    step = STEP_BY_ID[step_id]
    work_dir = str(cfg.get("work_dir") or "").strip()
    if not work_dir:
        return "作業フォルダが設定されていません。"
    if not Path(work_dir).is_dir():
        return f"作業フォルダが存在しません: {work_dir}"
    for key in step.needs:
        path = cfg.get(key) or ""
        if not path or not os.path.exists(path):
            return f"「{FIELD_LABELS.get(key, key)}」が見つかりません: {path or '(未設定)'}"
    return None


def check_requirements(step_id, cfg):
    """そのステップをいま実行できない理由 (設定の問題、またはファイルの段階の問題)。無ければ None。"""
    message = settings_problem(step_id, cfg)
    if message:
        return message
    return state.problem(step_id, state.actual_state(cfg["work_dir"]))


def precheck_detail(step_ids, cfg):
    """選択されたステップ全体の実行前チェック。

    戻り値: (設定の問題のリスト, 実行順序の問題のリスト)
    実行順序は、いまのファイルの段階から選択したステップを順に進めた場合を見通して判定する。
    """
    ordered = sorted(step_ids)
    settings, seen = [], set()
    for step_id in ordered:
        message = settings_problem(step_id, cfg)
        if message and message not in seen:
            seen.add(message)
            settings.append(message if "作業フォルダ" in message else f"STEP {step_id}: {message}")
    order = []
    work_dir = str(cfg.get("work_dir") or "").strip()
    if work_dir and Path(work_dir).is_dir():
        current = state.actual_state(work_dir)
        for step_id in ordered:
            message = state.problem(step_id, current)
            if message:
                order.append(f"STEP {step_id}: {message}")
                break                     # 以降は連鎖して同じ理由になるので、最初の 1 件だけ示す
            current = state.advance(step_id, current)
    return settings, order


def precheck(step_ids, cfg):
    """実行前チェックの問題をまとめたリスト (問題が無ければ空)。"""
    settings, order = precheck_detail(step_ids, cfg)
    return settings + order


def run_steps(step_ids, cfg, suppliers,
              on_status: Callable[[int, str], None] = lambda *_: None,
              confirm: Callable[[str], bool] = lambda _msg: True,
              should_stop: Callable[[], bool] = lambda: False):
    """ステップを番号順に実行する。最後まで実行できれば True。

    on_status(step_id, 状態): 状態は running / done / warn / error / skipped
        warn は「完了したが、確認が必要な点がある」。
    confirm(メッセージ): 確認が必要なときに呼ばれる。False なら中断。
    """
    ordered = sorted(step_ids)
    started = time.time()
    attention = []
    for position, step_id in enumerate(ordered):
        step = STEP_BY_ID[step_id]
        if should_stop():
            print("\n■ 中止しました。")
            for rest in ordered[position:]:
                on_status(rest, "skipped")
            return False

        print(f"\n{'=' * 78}\n◆ STEP {step.id}: {step.title}\n{'=' * 78}")
        on_status(step_id, "running")
        t0 = time.time()
        try:
            message = check_requirements(step_id, cfg)
            if message:
                from .steps import StepError
                raise StepError(message)
            question = _run_step(step_id, cfg, suppliers)
        except Exception as e:
            on_status(step_id, "error")
            if type(e).__name__ == "StepError":
                print(f"\n× STEP {step.id} エラー: {e}")
            elif isinstance(e, ModuleNotFoundError):
                print(f"\n× STEP {step.id} エラー: 必要なライブラリ「{e.name}」が入っていません。\n"
                      f"  「{config.SETUP_BAT_NAME}」をダブルクリックして入れてから、もう一度実行してください。\n"
                      f"  手動で入れる場合:  {config.install_command([e.name or ''])}")
            elif isinstance(e, PermissionError):
                print(f"\n× STEP {step.id} エラー: ファイルにアクセスできません。Excel で開いていないか確認してください。\n  {e}")
            else:
                print(f"\n× STEP {step.id} で予期しないエラーが発生しました: {type(e).__name__}: {e}")
                print(traceback.format_exc())
            for rest in ordered[position + 1:]:
                on_status(rest, "skipped")
            return False

        elapsed = time.time() - t0
        if question:
            attention.append(step_id)
            on_status(step_id, "warn")
            print(f"△ STEP {step.id} 完了 (要確認) ({elapsed:.1f}秒)")
            print(f"\n△ 確認: {question}")
        else:
            on_status(step_id, "done")
            print(f"○ STEP {step.id} 完了 ({elapsed:.1f}秒)")

        has_next = position + 1 < len(ordered)
        if question and has_next:
            if should_stop():             # 中止が押されていれば、確認は出さずに止める
                continue
            if not confirm(question + "\n\n次のステップへ進みますか？"):
                print("■ 利用者の指示により中断しました。")
                for rest in ordered[position + 1:]:
                    on_status(rest, "skipped")
                return False

    total = time.time() - started
    if attention:
        names = "、".join(f"STEP {n}" for n in attention)
        print(f"\n★ 選択したステップは最後まで実行しました。(合計 {total:.1f}秒)\n"
              f"△ ただし、確認が必要なステップがあります: {names}。上のログの「△ 確認」を見てください。")
    else:
        print(f"\n★ 選択したステップがすべて完了しました。(合計 {total:.1f}秒)")
    return True
