"""Excel 単語帳 → JSON 変換スクリプト.

新しい単語帳を追加するときの手順:

    python scripts/convert_excel.py <input.xlsx> <book_id> "<表示名>"

    例:
        python scripts/convert_excel.py ~/Downloads/ターゲット1900.xlsx target1900 "ターゲット1900"

前提とする Excel 形式:
    ヘッダ行 (1行目): No., 単語, 意味
    データ行: 連番 / 英単語 / 日本語の意味

出力:
    data/<book_id>.json
    data/books.json (レジストリを自動更新)
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    import openpyxl  # type: ignore
except ImportError:
    sys.stderr.write("openpyxl が見つかりません。`python -m pip install openpyxl` を実行してください。\n")
    sys.exit(1)


# ------ パス定義 -------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
REGISTRY_PATH = DATA_DIR / "books.json"
REGISTRY_JS_PATH = DATA_DIR / "books.js"

# 表記ゆれを吸収するためのヘッダ別名表(列順ではなく列名でマッピングする)
HEADER_ALIASES = {
    "no": ("no.", "no", "番号", "id"),
    "word": ("単語", "英単語", "word", "english", "英語"),
    "meaning": ("意味", "意 味", "和訳", "meaning", "japanese", "日本語"),
}


# ------ ユーティリティ -------------------------------------------------------
def _normalize_header(value: Any) -> str:
    """ヘッダセルを比較しやすい形に正規化する."""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).lower()


def _resolve_columns(header_row: tuple[Any, ...]) -> dict[str, int]:
    """ヘッダ行から論理列名 → 列インデックスを解決する."""
    normalized = [_normalize_header(h) for h in header_row]
    resolved: dict[str, int] = {}
    for logical, aliases in HEADER_ALIASES.items():
        for idx, cell in enumerate(normalized):
            if cell in aliases:
                resolved[logical] = idx
                break
    missing = [k for k in HEADER_ALIASES if k not in resolved]
    if missing:
        raise ValueError(
            f"Excel ヘッダに必須列が見つかりません: {missing}. "
            f"実際のヘッダ: {header_row}"
        )
    return resolved


def _read_book(xlsx_path: Path) -> list[dict[str, Any]]:
    """Excel を読み、空行を除いた単語データのリストを返す.

    1 行目がヘッダ行のファイル (ターゲット1400 等) と、
    ヘッダなしでいきなり No./単語/意味 が並ぶファイル (パス単シリーズ等) を
    自動判別する。判定基準は単純に「1 行目の 1 列目が整数か」。
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        raise ValueError("空の Excel ファイルです。")

    first = all_rows[0]
    if isinstance(first[0], int):
        # ヘッダなしファイル: 列順は [No., 単語, 意味] と仮定する
        cols = {"no": 0, "word": 1, "meaning": 2}
        data_rows = all_rows
        print(f"  (ヘッダ行なしと判定: 列順 [No., 単語, 意味] を仮定)")
    else:
        cols = _resolve_columns(first)
        data_rows = all_rows[1:]

    entries: list[dict[str, Any]] = []
    for raw in data_rows:
        no = raw[cols["no"]]
        word = raw[cols["word"]]
        meaning = raw[cols["meaning"]]

        # Excel が "false"/"true" を bool に自動変換してくるケースを救済する
        # (パス単2級 No.677 の "false" など)
        if isinstance(word, bool):
            word = "false" if word is False else "true"

        # 列が 1 つ左にズレた行を救済する。
        # No. 列に単語、単語列に意味が入り、意味列が空のケース
        # (システム英単語 No.1 "follow" など)。
        # 「No. が数値化できない文字列」かつ「意味が空」かつ「単語が存在」で判定。
        if isinstance(no, str) and word and not meaning:
            try:
                int(no)
            except (TypeError, ValueError):
                no, word, meaning = None, no, word
                # No. は連番で補完(直前の番号+1、無ければ1)
                no = (entries[-1]["no"] + 1) if entries else 1

        # No. と単語が空で意味だけある行は「前の単語の続き」とみなして結合する
        # (パス単1級 No.937 dub の意味が 2 行に分割されているケースなど)
        if no is None and not word and meaning and entries:
            entries[-1]["meaning"] = f'{entries[-1]["meaning"]} {str(meaning).strip()}'
            continue

        # それ以外で全列空はテンプレートの余白とみなして黙ってスキップ
        if not word and not meaning:
            continue
        # 単語または意味の片方だけが欠けている行はデータ不整合として警告
        if no is None or not word or not meaning:
            sys.stderr.write(f"  ! 不完全な行をスキップ: {raw}\n")
            continue
        try:
            no_int = int(no)
        except (TypeError, ValueError):
            sys.stderr.write(f"  ! No. が数値でない行をスキップ: {raw}\n")
            continue
        entries.append({
            "no": no_int,
            "word": str(word).strip(),
            "meaning": str(meaning).strip(),
        })
    # 念のため No. 順にソート
    entries.sort(key=lambda x: x["no"])
    return entries


def _upsert_registry(book_id: str, title: str, count: int, kind: str) -> None:
    """books.json (正本) と books.js (file:// 直開き用) を同時に更新する."""
    DATA_DIR.mkdir(exist_ok=True)
    registry: dict[str, Any] = {"books": []}
    if REGISTRY_PATH.exists():
        registry = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))

    books = registry.setdefault("books", [])
    existing = next((b for b in books if b["id"] == book_id), None)
    payload = {
        "id": book_id,
        "title": title,
        # JS ローダ側はこの相対パスを <script src> として使う
        "file": f"{book_id}.js",
        "count": count,
        # 出題形式ラベルの出し分けに使う種別 (en / kobun / generic 等)
        "kind": kind,
    }
    if existing:
        existing.update(payload)
    else:
        books.append(payload)

    REGISTRY_PATH.write_text(
        json.dumps(registry, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    _write_registry_js(registry)


def _write_registry_js(registry: dict[str, Any]) -> None:
    """books.json と等価な内容を <script> 経由で読める JS に書き出す.

    モジュール構文を使わないのは、file:// で開かれた HTML から
    type=module のロードが拒否されるブラウザがあるため。
    """
    payload_js = json.dumps(registry, ensure_ascii=False, indent=2)
    content = (
        "/* AUTO-GENERATED by scripts/convert_excel.py — 編集禁止 */\n"
        "(function (global) {\n"
        "  var ns = global.WTG = global.WTG || {};\n"
        f"  ns.registry = {payload_js};\n"
        "})(window);\n"
    )
    REGISTRY_JS_PATH.write_text(content, encoding="utf-8")


def _write_book_js(book_id: str, payload: dict[str, Any]) -> Path:
    """data/<id>.js を生成する.

    file:// で開いた HTML から <script src> 経由で読み込んでもらうことを想定し、
    JS は IIFE でグローバル汚染を最小限にしつつ window.WTG.books に登録する。
    """
    out_path = DATA_DIR / f"{book_id}.js"
    payload_js = json.dumps(payload, ensure_ascii=False, indent=2)
    content = (
        "/* AUTO-GENERATED by scripts/convert_excel.py — 編集禁止 */\n"
        "(function (global) {\n"
        "  var ns = global.WTG = global.WTG || {};\n"
        "  ns.books = ns.books || {};\n"
        f"  ns.books[{json.dumps(book_id)}] = {payload_js};\n"
        "})(window);\n"
    )
    out_path.write_text(content, encoding="utf-8")
    return out_path


# ------ エントリポイント -----------------------------------------------------
def main() -> int:
    parser = argparse.ArgumentParser(description="Excel 単語帳 → JSON 変換")
    parser.add_argument("xlsx", type=Path, help="入力 Excel ファイル")
    parser.add_argument("book_id", help='半角英数の ID 例: "target1400"')
    parser.add_argument("title", help='表示名 例: "ターゲット1400"')
    parser.add_argument(
        "--kind",
        default="en",
        choices=["en", "kobun", "generic"],
        help="出題形式ラベルの種別。en=英単語(英語→日本語)、"
             "kobun=古文(古語→訳)、generic=汎用(単語→意味)。既定: en",
    )
    args = parser.parse_args()

    # ID は URL/ファイル名に使われるので英数とハイフン/アンダースコアに制限
    if not re.fullmatch(r"[a-zA-Z0-9_\-]+", args.book_id):
        sys.stderr.write("book_id は英数字・ハイフン・アンダースコアのみ許可しています。\n")
        return 2

    if not args.xlsx.exists():
        sys.stderr.write(f"ファイルが見つかりません: {args.xlsx}\n")
        return 2

    print(f"読み込み中: {args.xlsx}")
    entries = _read_book(args.xlsx)
    print(f"  → {len(entries)} 語を抽出")

    out_path = DATA_DIR / f"{args.book_id}.json"
    out_path.parent.mkdir(exist_ok=True)
    payload = {
        "id": args.book_id,
        "title": args.title,
        "count": len(entries),
        "kind": args.kind,
        "words": entries,
    }
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"  → JSON 書き出し: {out_path}")

    js_path = _write_book_js(args.book_id, payload)
    print(f"  → JS 書き出し: {js_path}")

    _upsert_registry(args.book_id, args.title, len(entries), args.kind)
    print(f"  → レジストリ更新: {REGISTRY_PATH} / {REGISTRY_JS_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
