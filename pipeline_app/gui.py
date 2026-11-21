# -*- coding: utf-8 -*-
"""画面 (tkinter)。

処理は別のスレッドで動かし、画面とはキューでやり取りする
(処理中も画面が固まらず、print の出力がそのまま実行ログ欄に流れる)。
"""

import os
import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from . import APP_TITLE, __version__, config, runner, state

STATUS_TEXT = {
    "idle": ("", "#666666"),
    "running": ("実行中…", "#0b62c4"),
    "done": ("完了", "#1a7f37"),
    "warn": ("要確認", "#b45309"),          # 完了したが、確認が必要な点がある
    "error": ("エラー", "#c62828"),
    "skipped": ("未実行", "#8a6d00"),
}
UNMATCHED_PRODUCTS = "unmatched_products.xlsx"


class QueueWriter:
    """print の出力を画面の実行ログ欄へ流すための、標準出力の代役。"""

    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(("log", str(text)))
        return len(text) if text else 0

    def flush(self):
        pass

    def isatty(self):
        return False


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  v{__version__}" + (f"  [{config.PROFILE}]" if config.PROFILE else ""))
        self.geometry("1120x720")
        self.minsize(1010, 700)

        self.cfg = config.load_config()
        self.suppliers = config.discover_suppliers()
        self.queue = queue.Queue()
        self.worker = None
        self.stop_requested = threading.Event()
        self.last_run_ok = None             # 直近の実行結果 (自動操作のスクリプトから結果を読めるようにしてある)
        self.running_steps = False          # いま動いている (直近に動かした) のがステップの実行か、前処理か

        sys.stdout = sys.stderr = QueueWriter(self.queue)

        self._build()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll)
        print(f"{APP_TITLE} v{__version__} を起動しました。")
        print(f"検出した仕入先の変換部品: {len(self.suppliers)} 件")
        if config.CONFIG_OVERRIDDEN:
            print(f"設定ファイル: {config.CONFIG_PATH}")
            print(f"作業フォルダ: {self.cfg.get('work_dir')}")
        if not self.cfg.get("input_root"):
            print("△ 「設定」タブで、入力データのルートフォルダとマスタファイルを指定してください。")
        for note in config.LOAD_NOTES + config.DISCOVERY_NOTES:
            print("△ " + note)
        missing = config.missing_libraries()
        if missing:
            print("△ " + config.missing_libraries_message(missing))

    def _libraries_ready(self):
        """必要なライブラリがそろっていれば True。足りなければ案内を出して False。"""
        missing = config.missing_libraries()
        if not missing:
            return True
        message = config.missing_libraries_message(missing)
        print("\n× " + message)
        messagebox.showerror("必要なライブラリがありません", message, parent=self)
        return False

    # ------------------------------------------------------------------ 画面の組み立て
    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Run.TButton", font=("Yu Gothic UI", 11, "bold"), padding=8)
        style.configure("Head.TLabel", font=("Yu Gothic UI", 10, "bold"))

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_run = ttk.Frame(notebook, padding=8)
        self.tab_settings = ttk.Frame(notebook, padding=8)
        notebook.add(self.tab_run, text="  実行  ")
        notebook.add(self.tab_settings, text="  設定  ")
        self.notebook = notebook
        self._build_run_tab()
        self._build_settings_tab()

    def _build_run_tab(self):
        tab = self.tab_run
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(0, weight=1)

        left = ttk.Frame(tab)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10))

        box = ttk.LabelFrame(left, text="処理ステップ (上から順に実行)", padding=8)
        box.pack(fill="x")
        self.step_vars, self.step_status = {}, {}
        for row, step in enumerate(runner.STEPS):
            var = tk.BooleanVar(value=step.default)
            self.step_vars[step.id] = var
            ttk.Checkbutton(box, variable=var, text=f"STEP {step.id}  {step.title}").grid(
                row=row * 2, column=0, sticky="w")
            status = ttk.Label(box, text="", width=9, anchor="e")
            status.grid(row=row * 2, column=1, sticky="e", padx=(8, 0))
            self.step_status[step.id] = status
            ttk.Label(box, text=step.detail, foreground="#666666").grid(
                row=row * 2 + 1, column=0, columnspan=2, sticky="w", padx=(22, 0), pady=(0, 4))

        pick = ttk.Frame(left)
        pick.pack(fill="x", pady=(6, 0))
        ttk.Button(pick, text="0〜7 を選択", command=lambda: self._select_steps(range(0, 8))).pack(side="left")
        ttk.Button(pick, text="すべて外す", command=lambda: self._select_steps([])).pack(side="left", padx=4)

        self.run_button = ttk.Button(left, text="選択したステップを実行", style="Run.TButton", command=self.start_run)
        self.run_button.pack(fill="x", pady=(12, 4))
        self.stop_button = ttk.Button(left, text="現在のステップ終了後に中止", command=self._request_stop, state="disabled")
        self.stop_button.pack(fill="x")
        self.progress = ttk.Progressbar(left, mode="indeterminate")
        self.progress.pack(fill="x", pady=(8, 0))

        right = ttk.Frame(tab)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        head = ttk.Frame(right)
        head.grid(row=0, column=0, sticky="ew")
        ttk.Label(head, text="実行ログ", style="Head.TLabel").pack(side="left")
        ttk.Button(head, text="ログを消去", command=self._clear_log).pack(side="right")
        ttk.Button(head, text="ログをコピー", command=self._copy_log).pack(side="right", padx=(0, 4))
        self.convert_button = ttk.Button(head, text="前処理: csv/xls → xlsx 変換…", command=self.start_convert)
        self.convert_button.pack(side="right", padx=4)
        ttk.Button(head, text="未一致の商品名リスト", command=self._open_unmatched).pack(side="right")
        ttk.Button(head, text="作業フォルダを開く", command=self._open_work_dir).pack(side="right", padx=4)
        self.log = ScrolledText(right, wrap="word", font=("MS Gothic", 10), state="disabled")
        self.log.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        self.log.tag_configure("error", foreground="#c62828")
        self.log.tag_configure("warn", foreground="#8a6d00")
        self.log.tag_configure("ok", foreground="#1a7f37")
        self.log.tag_configure("head", foreground="#0b62c4")
        self.status_bar = ttk.Label(tab, text="待機中", anchor="w")
        self.status_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

    def _build_settings_tab(self):
        tab = self.tab_settings
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        paths = ttk.LabelFrame(tab, text="フォルダとマスタファイル", padding=8)
        paths.grid(row=0, column=0, sticky="ew")
        paths.columnconfigure(1, weight=1)
        self.setting_vars = {}
        for row, (key, label, kind) in enumerate(config.SETTING_FIELDS):
            ttk.Label(paths, text=label).grid(row=row, column=0, sticky="w", pady=2)
            var = tk.StringVar(value=self.cfg.get(key, ""))
            self.setting_vars[key] = var
            ttk.Entry(paths, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6, pady=2)
            if kind in ("dir", "file"):
                ttk.Button(paths, text="参照…", width=8,
                           command=lambda k=key, t=kind: self._browse(k, t)).grid(row=row, column=2)
            else:
                ttk.Label(paths, text="(通常は input。空欄なら仕入先フォルダの直下)", foreground="#666666").grid(
                    row=row, column=2, sticky="w")

        sup = ttk.LabelFrame(tab, text="処理対象の仕入先  ―  入力フォルダ: <ルート>/<仕入先名>/<サブフォルダ>", padding=8)
        sup.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        grid = ttk.Frame(sup)
        grid.pack(fill="both", expand=True)
        disabled = set(self.cfg.get("disabled_suppliers", []))
        self.supplier_vars, self.supplier_checks = {}, {}
        columns = 4
        for i, supplier in enumerate(self.suppliers):
            var = tk.BooleanVar(value=supplier.key not in disabled)
            self.supplier_vars[supplier.key] = var
            check = ttk.Checkbutton(grid, variable=var, text=supplier.name)
            check.grid(row=i // columns, column=i % columns, sticky="w", padx=(0, 24), pady=2)
            self.supplier_checks[supplier.key] = check
        buttons = ttk.Frame(sup)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="すべて選択", command=lambda: self._select_suppliers(True)).pack(side="left")
        ttk.Button(buttons, text="すべて外す", command=lambda: self._select_suppliers(False)).pack(side="left", padx=4)
        ttk.Button(buttons, text="入力ファイル件数を再確認", command=self._refresh_counts).pack(side="left", padx=12)

        bottom = ttk.Frame(tab)
        bottom.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bottom, text="設定を保存", command=self._save_settings_clicked).pack(side="left")
        ttk.Label(bottom, text=f"保存先: {config.CONFIG_PATH}", foreground="#666666").pack(side="left", padx=10)
        self._refresh_counts()

    # ------------------------------------------------------------------ 設定
    def _browse(self, key, kind):
        current = self.setting_vars[key].get()
        if kind == "dir":
            chosen = filedialog.askdirectory(title="フォルダを選択", initialdir=current or str(config.APP_DIR))
        else:
            start = os.path.dirname(current) if current else str(config.APP_DIR)
            chosen = filedialog.askopenfilename(
                title="ファイルを選択", initialdir=start, filetypes=[("Excel ファイル", "*.xlsx")])
        if chosen:
            self.setting_vars[key].set(os.path.normpath(chosen))
            if key in ("input_root", "input_subdir"):
                self._refresh_counts()

    def _collect_settings(self):
        for key, var in self.setting_vars.items():
            self.cfg[key] = var.get().strip().strip('"')
        self.cfg["disabled_suppliers"] = [k for k, v in self.supplier_vars.items() if not v.get()]
        return self.cfg

    def _save_settings(self):
        """設定を保存する。保存できたら True。"""
        self._collect_settings()
        try:
            config.save_config(self.cfg)
            return True
        except OSError as e:
            print(f"△ 設定を保存できませんでした: {e}")
            return False

    def _save_settings_clicked(self):
        saved = self._save_settings()
        self._refresh_counts()
        if saved:
            messagebox.showinfo("設定", "設定を保存しました。", parent=self)
        else:
            messagebox.showerror("設定", f"設定を保存できませんでした。\n保存先: {config.CONFIG_PATH}", parent=self)

    def _select_suppliers(self, value):
        for var in self.supplier_vars.values():
            var.set(value)

    def _refresh_counts(self):
        cfg = dict(self._collect_settings())
        for supplier in self.suppliers:
            count = config.count_input_files(cfg, supplier)
            suffix = "フォルダなし" if count is None else f"{count}件"
            self.supplier_checks[supplier.key].configure(text=f"{supplier.name}  ({suffix})")

    def _select_steps(self, ids):
        ids = set(ids)
        for step_id, var in self.step_vars.items():
            var.set(step_id in ids)

    # ------------------------------------------------------------------ 実行
    def _busy(self):
        return self.worker is not None and self.worker.is_alive()

    def start_run(self, step_ids=None):
        if self._busy():
            return
        if step_ids is None:
            step_ids = [i for i, v in self.step_vars.items() if v.get()]
        if not step_ids:
            messagebox.showwarning("実行", "実行するステップを選択してください。", parent=self)
            return
        if not self._libraries_ready():
            return
        self._save_settings()
        cfg = dict(self.cfg)
        suppliers = [s for s in self.suppliers if s.key not in set(cfg["disabled_suppliers"])]

        settings_problems, order_problems = runner.precheck_detail(step_ids, cfg)
        if (0 in step_ids or 8 in step_ids) and not suppliers:
            settings_problems.append("処理対象の仕入先が 1 つも選択されていません。")
        if settings_problems or order_problems:
            message = "\n".join(settings_problems + order_problems)
            if settings_problems:
                message += "\n\n「設定」タブを確認してください。"
            messagebox.showerror("実行できません", message, parent=self)
            return
        if 8 in step_ids and not messagebox.askyesno(
                "後処理の確認",
                "STEP 8 を実行すると、各仕入先の入力ファイルが processed フォルダへ、\n"
                f"成果物 ({state.WORK_FILE} / {state.OUTPUT_FILE}) が backup フォルダへ移動されます。\n\n"
                "BI への取り込みに使うファイルも backup フォルダ内に移動します。実行しますか？",
                parent=self):
            return

        for step_id in self.step_status:
            self._set_status(step_id, "idle")
        self.running_steps = True
        self._start_worker(lambda: runner.run_steps(
            step_ids, cfg, suppliers,
            on_status=lambda i, s: self.queue.put(("status", i, s)),
            confirm=self._confirm_from_worker,
            should_stop=self.stop_requested.is_set))

    def start_convert(self):
        if self._busy():
            return
        if not self._libraries_ready():
            return
        self._collect_settings()
        folder = filedialog.askdirectory(
            title="変換するフォルダを選択 (サブフォルダも対象)",
            initialdir=self.cfg.get("input_root") or str(config.APP_DIR))
        if not folder:
            return
        if not messagebox.askyesno(
                "前処理の確認",
                f"対象フォルダ:\n{folder}\n\n"
                "サブフォルダを含むすべての csv / xls / xlsm を xlsx に変換します。\n"
                "・変換元のファイルは、同じ場所の converted_originals フォルダへ移動します (削除はしません)\n"
                "・すでにある xlsx には触れません\n\n"
                "実行しますか？",
                parent=self):
            return

        def job():
            from . import steps
            print(f"\n{'=' * 78}\n◆ 前処理: csv / xls → xlsx 変換\n{'=' * 78}")
            try:
                result = steps.convert_folder(folder)
            except steps.StepError as e:
                print(f"× 前処理エラー: {e}")
                return False
            except Exception as e:
                print(f"× 前処理でエラーが発生しました: {type(e).__name__}: {e}")
                return False
            if result["failed"]:
                print(f"× 前処理は終わりましたが、変換できなかったファイルが {len(result['failed'])} 件あります (上のログを確認してください)。")
                return False
            print("○ 前処理が完了しました。")
            return True

        self.running_steps = False
        self._start_worker(job, can_stop=False)

    def _start_worker(self, job, can_stop=True):
        self.stop_requested.clear()
        self.last_run_ok = None
        self.run_button.configure(state="disabled")
        self.convert_button.configure(state="disabled")
        self.stop_button.configure(state="normal" if can_stop else "disabled")   # 前処理は途中で止められない
        self.progress.start(12)
        self._set_bar("実行中…")
        self.notebook.select(self.tab_run)

        def target():
            ok = False
            try:
                ok = bool(job())
            except ModuleNotFoundError as e:
                print("× " + config.missing_libraries_message([e.name or "?"]))
            except Exception as e:
                import traceback
                print(f"× 予期しないエラー: {type(e).__name__}: {e}\n{traceback.format_exc()}")
            finally:
                self.queue.put(("finished", ok))

        self.worker = threading.Thread(target=target, daemon=True)
        self.worker.start()

    def _request_stop(self):
        self.stop_requested.set()
        self.stop_button.configure(state="disabled")
        self._set_bar("現在のステップが終わりしだい中止します…")

    def _confirm_from_worker(self, message):
        """処理スレッドから呼ばれる。画面側で確認ダイアログを出して結果を待つ。"""
        answer = {}
        done = threading.Event()
        self.queue.put(("confirm", message, answer, done))
        done.wait()
        return answer.get("value", False)

    # ------------------------------------------------------------------ 画面の更新
    def _poll(self):
        """処理スレッドからの連絡 (ログ・状態・確認・終了) を画面に反映する。

        1 件の処理で例外が出ても止まらないようにする (止まると画面が更新されなくなり、
        確認待ちの処理スレッドも進めなくなるため)。
        """
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                try:
                    if kind == "log":
                        self._append_log(item[1])
                    elif kind == "status":
                        self._set_status(item[1], item[2])
                    elif kind == "confirm":
                        _, message, answer, done = item
                        try:
                            answer["value"] = self.ask_confirm(message)
                        finally:
                            done.set()            # 何があっても処理スレッドを待たせたままにしない
                    elif kind == "finished":
                        self._finished(item[1])
                except Exception as e:            # noqa: BLE001  画面更新の失敗で全体を止めない
                    try:
                        self._write_log_file(f"\n(画面の更新中にエラー: {type(e).__name__}: {e})\n")
                    except Exception:
                        pass
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def ask_confirm(self, message):
        return messagebox.askyesno("確認", message, icon="warning", parent=self)

    def _finished(self, ok):
        self.last_run_ok = ok
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.convert_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        attention = self.running_steps and any(
            label.cget("text") == STATUS_TEXT["warn"][0] for label in self.step_status.values())
        if not ok:
            self._set_bar("中断またはエラーで終了しました (ログを確認してください)")
        elif attention:
            self._set_bar("完了しました (「要確認」のステップがあります。ログの「△ 確認」を見てください)")
        else:
            self._set_bar("完了しました")
        self._refresh_counts()

    def _set_bar(self, text):
        """画面下部の状態表示を変える (一時的なメッセージの表示中でも、元に戻す先として覚えておく)。"""
        self._bar_text = text
        self.status_bar.configure(text=text)

    def _set_status(self, step_id, status):
        text, color = STATUS_TEXT[status]
        self.step_status[step_id].configure(text=text, foreground=color)

    def _append_log(self, text):
        self._write_log_file(text)
        self.log.configure(state="normal")
        for line in text.splitlines(keepends=True):
            tag = ()
            if line.startswith("×") or "Traceback" in line or "エラー" in line:
                tag = ("error",)
            elif line.startswith("△") or "注意" in line:
                tag = ("warn",)
            elif line.startswith(("○", "★")):
                tag = ("ok",)
            elif line.startswith("◆"):
                tag = ("head",)
            self.log.insert("end", line, tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _write_log_file(self, text):
        try:
            base = Path(self.cfg.get("work_dir") or config.APP_DIR)
            if not base.is_dir():             # 入力途中の存在しないフォルダを勝手に作らない
                return
            folder = base / "logs"
            folder.mkdir(exist_ok=True)
            name = f"pipeline_{datetime.now().strftime('%Y-%m-%d')}.log"
            with open(folder / name, "a", encoding="utf-8") as f:
                f.write(text)
        except OSError:
            pass

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _copy_log(self):
        """画面に表示されているログ全体をクリップボードにコピーする (問い合わせ時に貼り付けるため)。"""
        text = self.log.get("1.0", "end-1c").rstrip("\n")
        if not text.strip():
            self._flash_status("コピーするログがありません")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()        # アプリを閉じたあとも貼り付けられるように、クリップボードへ確実に渡す
        self._flash_status(f"ログをコピーしました ({text.count(chr(10)) + 1} 行)。メールやチャットに貼り付けられます")

    def _flash_status(self, message, millis=4000):
        """画面下部に一時的なメッセージを出し、しばらくしたら本来の表示に戻す。"""
        self.status_bar.configure(text=message)

        def restore():
            if self.status_bar.cget("text") == message:
                self.status_bar.configure(text=getattr(self, "_bar_text", "待機中"))

        if getattr(self, "_flash_job", None):
            self.after_cancel(self._flash_job)
        self._flash_job = self.after(millis, restore)

    # ------------------------------------------------------------------ 道具
    def _open_path(self, path):
        if not os.path.exists(path):
            messagebox.showinfo("開く", f"見つかりません:\n{path}", parent=self)
            return
        try:
            os.startfile(path)
        except (OSError, AttributeError) as e:
            messagebox.showerror("開く", f"開けませんでした:\n{path}\n{e}", parent=self)

    def _open_work_dir(self):
        self._collect_settings()
        self._open_path(self.cfg.get("work_dir", ""))

    def _open_unmatched(self):
        self._collect_settings()
        self._open_path(os.path.join(self.cfg.get("work_dir", ""), UNMATCHED_PRODUCTS))

    def _on_close(self):
        if self._busy() and not messagebox.askyesno(
                "終了の確認", "処理の実行中です。終了すると出力ファイルが壊れる可能性があります。\n終了しますか？", parent=self):
            return
        self._save_settings()
        self.destroy()


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
