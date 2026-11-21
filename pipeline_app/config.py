# -*- coding: utf-8 -*-
"""設定の読み書きと、仕入先ごとの変換部品の検出。

このモジュールは pandas を読み込まない (画面をすぐに表示するため)。
"""

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent
SUPPLIER_SCRIPT_DIR = APP_DIR / "suppliers"

# 環境変数 PIPELINE_CONFIG で設定ファイルを切り替えられる (サンプルや検証用の環境など)。
# 切り替えたときは、設定ファイル内のパスをその設定ファイルのフォルダからの相対パスで書ける。
_ENV_CONFIG = os.environ.get("PIPELINE_CONFIG", "").strip().strip('"')
CONFIG_OVERRIDDEN = bool(_ENV_CONFIG)
CONFIG_PATH = Path(_ENV_CONFIG).resolve() if _ENV_CONFIG else APP_DIR / "pipeline_config.json"
PROFILE = os.environ.get("PIPELINE_PROFILE", "").strip()      # 画面タイトルに表示する名前 (例: SAMPLE)

# 設定項目: (キー, 画面表示名, 種別)   種別: dir / file / text
SETTING_FIELDS = [
    ("work_dir", "作業フォルダ (出力先)", "dir"),
    ("input_root", "入力データのルートフォルダ", "dir"),
    ("input_subdir", "仕入先フォルダ内の入力サブフォルダ名", "text"),
    ("product_master", "商品マスタ", "file"),
    ("customer_master", "顧客マスタ", "file"),
    ("price_master", "価格マスタ", "file"),
    ("prefecture_table", "都道府県テーブル", "file"),
    ("id_master", "ID マスタ", "file"),
]
_PATH_KEYS = [key for key, _label, kind in SETTING_FIELDS if kind in ("dir", "file")]

REQUIRED_LIBRARIES = ["pandas", "numpy", "openpyxl"]
SETUP_BAT_NAME = "setup.bat"

_SCRIPT_NAME_RE = re.compile(r"^process_[A-Za-z0-9_]+\.py$")
_SUPPLIER_NAME_RE = re.compile(r"""^SUPPLIER_NAME\s*=\s*["'](.+?)["']""", re.MULTILINE)
DISCOVERY_NOTES = []      # 直近の discover_suppliers() で気づいた点 (画面の起動時に表示する)
LOAD_NOTES = []           # 直近の load_config() で気づいた点


@dataclass
class Supplier:
    key: str          # 例: "sample_a"  (process_sample_a.py)
    name: str         # 例: "サンプル卸A"  (入力フォルダ名、出力の「卸名」)
    script: Path


def missing_libraries():
    """処理に必要なのに入っていないライブラリ名のリスト (ライブラリ自体は読み込まない)。"""
    import importlib.util
    return [name for name in REQUIRED_LIBRARIES if importlib.util.find_spec(name) is None]


def install_command(names=None):
    """いま動いている Python に対してライブラリを入れるコマンド。"""
    import sys
    exe = sys.executable or "python"
    if exe.lower().endswith("pythonw.exe"):
        exe = exe[:-len("pythonw.exe")] + "python.exe"
    return f'"{exe}" -m pip install ' + " ".join(names or REQUIRED_LIBRARIES)


def missing_libraries_message(missing):
    return ("必要なライブラリが入っていません: " + ", ".join(missing) + "\n\n"
            f"「{SETUP_BAT_NAME}」をダブルクリックして入れてから、もう一度実行してください。\n\n"
            "手動で入れる場合は、コマンド プロンプトで次を実行します:\n" + install_command())


def _read_script(script):
    for encoding in ("utf-8-sig", "cp932"):
        try:
            return script.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
        except OSError:
            return ""
    return ""


def discover_suppliers():
    """suppliers/process_*.py を検出して Supplier のリストを返す (ファイル名順)。

    各スクリプトの先頭付近にある  SUPPLIER_NAME = "…"  を仕入先名として使う。
    スクリプトのコピー (例: 「process_sample_a - コピー.py」) を取り込むと同じファイルを二重に
    集計してしまうため、名前が process_<英数字>.py の形のものだけを対象にし、
    同じ仕入先名のスクリプトが複数あるときは最初の 1 つだけを使う。
    """
    notes, found, by_name = [], [], {}
    if SUPPLIER_SCRIPT_DIR.is_dir():
        for script in sorted(SUPPLIER_SCRIPT_DIR.glob("process_*.py")):
            if not _SCRIPT_NAME_RE.match(script.name):
                notes.append(f"仕入先の変換部品として扱わないファイルがあります "
                             f"(名前が process_英数字.py の形ではありません): {script.name}")
                continue
            match = _SUPPLIER_NAME_RE.search(_read_script(script))
            if not match:
                notes.append(f"SUPPLIER_NAME が書かれていないため使いません: {script.name}")
                continue
            name = match.group(1).strip()
            if name in by_name:
                notes.append(f"同じ仕入先 ({name}) の変換部品が複数あります。"
                             f"{by_name[name].script.name} を使い、{script.name} は使いません。")
                continue
            supplier = Supplier(script.stem[len("process_"):], name, script)
            by_name[name] = supplier
            found.append(supplier)
    DISCOVERY_NOTES[:] = notes
    return found


def default_config():
    return {
        "work_dir": "",
        "input_root": "",
        "input_subdir": "input",
        "product_master": "",
        "customer_master": "",
        "price_master": "",
        "prefecture_table": "",
        "id_master": "",
        "disabled_suppliers": [],
    }


def _set_aside(path):
    """読み込めなかった設定ファイルを、上書きされる前に別名で残す。"""
    backup = path.with_name(path.name + ".broken")
    try:
        backup.write_bytes(path.read_bytes())
        return backup
    except OSError:
        return None


def load_config():
    del LOAD_NOTES[:]
    cfg = default_config()
    if CONFIG_PATH.exists():
        try:
            saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
            if not isinstance(saved, dict):
                raise ValueError("設定ファイルの形式が違います")
        except (OSError, ValueError) as e:
            backup = _set_aside(CONFIG_PATH)
            LOAD_NOTES.append(
                f"設定ファイルを読み込めなかったため、初期設定で起動しました ({CONFIG_PATH.name}: {e})。"
                + (f"元のファイルは {backup.name} として残しました。" if backup else ""))
            saved = {}
        for key, value in saved.items():
            if key not in cfg:
                continue
            if key == "disabled_suppliers":
                if isinstance(value, list):
                    cfg[key] = [str(v) for v in value]
            elif isinstance(value, str):
                cfg[key] = value
    elif CONFIG_OVERRIDDEN:
        LOAD_NOTES.append(f"指定された設定ファイルがありません: {CONFIG_PATH} (初期設定で起動しました)")
    for key in _PATH_KEYS:        # 相対パスは設定ファイルのフォルダを基準に解決する
        value = cfg.get(key) or ""
        if value and not os.path.isabs(value):
            cfg[key] = os.path.normpath(os.path.join(str(CONFIG_PATH.parent), value))
    return cfg


def save_config(cfg):
    out = dict(cfg)
    if CONFIG_OVERRIDDEN:         # 切り替えた設定ファイルは相対パスで保存 (フォルダごと移動しても使えるように)
        for key in _PATH_KEYS:
            value = out.get(key) or ""
            if value and os.path.isabs(value):
                try:
                    out[key] = os.path.relpath(value, str(CONFIG_PATH.parent)).replace(os.sep, "/")
                except ValueError:        # ドライブが違う場合は絶対パスのまま
                    pass
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")


def input_dir(cfg, supplier):
    """仕入先の入力フォルダ (<ルート>/<仕入先名>/<サブフォルダ>)。"""
    base = Path(cfg["input_root"]) / supplier.name
    sub = (cfg.get("input_subdir") or "").strip()
    return base / sub if sub else base


def processed_dir(cfg, supplier):
    """処理済みファイルの移動先 (<ルート>/<仕入先名>/processed)。"""
    return Path(cfg["input_root"]) / supplier.name / "processed"


def count_input_files(cfg, supplier):
    """入力フォルダ内の xlsx 件数。フォルダが無ければ None。"""
    folder = input_dir(cfg, supplier)
    if not cfg.get("input_root") or not folder.is_dir():
        return None
    return len([f for f in folder.glob("*.xlsx") if not f.name.startswith("~")])
