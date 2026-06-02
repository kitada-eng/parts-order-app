# -*- coding: utf-8 -*-
"""
folder_loader.py
「発注履歴見積書」フォルダ(Googleドライブ パソコン版でローカル同期されたフォルダ)を
走査し、中のファイルから発注履歴を統合的に構築する。Google Cloud / API は不要。

取得元は2通り:
  (1) ローカルフォルダ(Googleドライブ パソコン版で同期されたパス) … scan_folder()
  (2) 共有DriveフォルダのWeb URL(リンク共有・閲覧可)              … scan_drive_folder_url()

対応ファイル:
  - .xlsx / .xlsm / .csv : 表形式の履歴(列: 型式 品名 発注先 単価 備考)。見出し名で列を自動判別。
  - Googleスプレッドシート: CSV公開URLから取得。
  - .pdf                 : 見積書。表を抽出して 型式→単価 を取得し、発注先はファイル名等から推定。

同じ型式が複数ファイルに出てきた場合は、既定で新しいもの(URL時は後勝ち)を優先する。
"""

from __future__ import annotations

import io
import json
import os
import re
import urllib.request
from dataclasses import dataclass, field
from typing import Optional

import order_core as core
from order_core import HistoryRecord, _find_header_indices, _parse_price, _normalize_model

TABLE_EXT = {".xlsx", ".xlsm", ".csv"}
GSHEET_EXT = {".gsheet"}
PDF_EXT = {".pdf"}

_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
_GAPP_FOLDER = "application/vnd.google-apps.folder"
_GAPP_SHEET = "application/vnd.google-apps.spreadsheet"


@dataclass
class ScanReport:
    """走査結果のサマリ(画面表示・トラブルシュート用)"""
    files_seen: int = 0
    tables_loaded: int = 0
    gsheets_loaded: int = 0
    pdfs_loaded: int = 0
    records_total: int = 0
    pdf_items: int = 0
    sources: list = field(default_factory=list)      # 読み込めたファイル名
    skipped: list = field(default_factory=list)       # (ファイル名, 理由)
    errors: list = field(default_factory=list)        # (ファイル名, エラー内容)


# ----------------------------------------------------------------------------
# フォルダ走査
# ----------------------------------------------------------------------------
def scan_folder(
    folder: str,
    *,
    recursive: bool = True,
    conflict: str = "latest",
) -> tuple[dict[str, HistoryRecord], ScanReport]:
    """
    フォルダを走査して、統合した履歴 dict と ScanReport を返す。

    conflict : 'latest' = 更新日時が新しいファイルを優先(既定)
               'first'  = 先に見つかったものを優先
    """
    if not os.path.isdir(folder):
        raise NotADirectoryError(f"フォルダが見つかりません: {folder}")

    # 走査対象ファイルを収集
    paths: list[str] = []
    if recursive:
        for root, _dirs, files in os.walk(folder):
            for f in files:
                paths.append(os.path.join(root, f))
    else:
        paths = [os.path.join(folder, f) for f in os.listdir(folder)
                 if os.path.isfile(os.path.join(folder, f))]

    # 常に古い→新しいの順に処理する。
    #   latest優先: 後勝ち(history[k]=v)で新しいファイルが上書き
    #   first優先 : 先勝ち(setdefault)で古いファイルが残る
    paths.sort(key=lambda p: os.path.getmtime(p))

    history: dict[str, HistoryRecord] = {}
    report = ScanReport()

    def merge(sub: dict[str, HistoryRecord]):
        for k, v in sub.items():
            if conflict == "first":
                history.setdefault(k, v)
            else:
                history[k] = v  # latest: 後勝ち

    for path in paths:
        ext = os.path.splitext(path)[1].lower()
        name = os.path.basename(path)
        if ext not in TABLE_EXT | GSHEET_EXT | PDF_EXT:
            continue
        report.files_seen += 1
        try:
            if ext in TABLE_EXT:
                sub = core.load_history_from_file(path)
                _stamp_source(sub, name)
                merge(sub)
                report.tables_loaded += 1
                report.sources.append(name)
            elif ext in GSHEET_EXT:
                sub = _load_gsheet_shortcut(path)
                _stamp_source(sub, name)
                merge(sub)
                report.gsheets_loaded += 1
                report.sources.append(name)
            elif ext in PDF_EXT:
                vendor, sub = parse_quote_pdf(path)
                _stamp_source(sub, name)
                merge(sub)
                report.pdfs_loaded += 1
                report.pdf_items += len(sub)
                report.sources.append(f"{name} (発注先推定: {vendor or '不明'})")
        except PermissionError as e:
            report.skipped.append((name, str(e)))
        except Exception as e:  # noqa: BLE001
            report.errors.append((name, f"{type(e).__name__}: {e}"))

    report.records_total = len(history)
    return history, report


def _stamp_source(sub: dict[str, HistoryRecord], source: str):
    for rec in sub.values():
        if not rec.source:
            rec.source = source


# ----------------------------------------------------------------------------
# .gsheet ショートカットの読み込み
# ----------------------------------------------------------------------------
def _load_gsheet_shortcut(path: str) -> dict[str, HistoryRecord]:
    """.gsheet(JSON)から対象スプレッドシートのURLを取り出しCSVで取得する。"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    url = data.get("url", "")
    doc_id = data.get("doc_id", "") or data.get("resource_key", "")
    if "/spreadsheets/d/" in url:
        return core.load_history_from_csv_url(url)
    if doc_id:
        synthetic = f"https://docs.google.com/spreadsheets/d/{doc_id}/edit"
        return core.load_history_from_csv_url(synthetic)
    raise PermissionError(
        "ネイティブGoogleスプレッドシートはローカルに実データを持たないため、"
        "対象シートを『リンクを知っている全員が閲覧可』にする必要があります。"
    )


# ----------------------------------------------------------------------------
# PDF 見積書のパーサ(雛形・実物で要較正)
# ----------------------------------------------------------------------------
def _vendor_from_filename(path: str) -> str:
    """ファイル名の先頭トークンを発注先(メーカー)名として推定する。
    例: 'ミスミ_見積書_20250601.pdf' -> 'ミスミ'
    """
    name = os.path.splitext(os.path.basename(path))[0]
    # 日付や「見積」等のノイズを除いた先頭トークン
    token = re.split(r"[ _\-　]+", name)[0]
    token = re.sub(r"(御?見積書?|quote|estimate)$", "", token, flags=re.IGNORECASE)
    return token.strip()


def _records_from_table(table, vendor: str) -> dict[str, HistoryRecord]:
    """PDFから抽出した1つの表(list[list])を 型式→HistoryRecord に変換する。"""
    out: dict[str, HistoryRecord] = {}
    if not table:
        return out
    # 見出し行を探す(先頭数行のうち型式列が見つかる行)
    header_row = None
    header_idx = {}
    for i, row in enumerate(table[:3]):
        idx = _find_header_indices([("" if c is None else c) for c in row])
        if "model" in idx and ("price" in idx or "name" in idx):
            header_row = i
            header_idx = idx
            break
    if header_row is None:
        return out

    def cell(row, key, default=None):
        i = header_idx.get(key, default)
        if i is None or i >= len(row):
            return ""
        return "" if row[i] is None else row[i]

    for row in table[header_row + 1:]:
        model = _normalize_model(cell(row, "model"))
        if not model:
            continue
        out[model] = HistoryRecord(
            name=str(cell(row, "name")).strip(),
            vendor=vendor,
            price=_parse_price(cell(row, "price")),
            note=str(cell(row, "note")).strip(),
        )
    return out


_NUMERIC_RE = re.compile(r"^[\d,，.\-\s¥￥]+$")
_UNIT_WORDS = {"台", "個", "本", "式", "枚", "組", "セット", "点", "箱", "巻", "m", "ｍ", "kg", "pcs"}


def _is_numeric_cell(text: str) -> bool:
    return bool(_NUMERIC_RE.match(text))


def _pick_item_text(row, exclude_idx: set[int]) -> str:
    """
    明細行から「品名/型式」とみなせる最長テキストセルを選ぶ。
    数値だけのセル(行番号/数量/単価/金額)や単位語(台/式…)は除外する。
    見積書は列ズレが起きやすいため、ヘッダー位置ではなく中身で判定する。
    """
    best = ""
    for i, c in enumerate(row):
        if i in exclude_idx or c is None:
            continue
        s = str(c).strip()
        if not s or _is_numeric_cell(s) or s in _UNIT_WORDS:
            continue
        if len(s) > len(best):
            best = s
    return best


def _records_from_quote_table(table, vendor: str) -> dict[str, HistoryRecord]:
    """
    見積書PDFの明細表(list[list])から 型式→HistoryRecord を抽出する。
    手順: 単価列を見出しから特定 → 各行で単価が取れる行のみ明細とみなし、
          行内の最長テキストを型式/品名として採用する。
    """
    out: dict[str, HistoryRecord] = {}
    if not table:
        return out

    # 単価列を含む見出し行を探す(先頭4行以内)
    price_col = None
    qty_col = None
    header_row = None
    for i, row in enumerate(table[:4]):
        idx = _find_header_indices([("" if c is None else c) for c in row])
        if "price" in idx:
            price_col = idx["price"]
            qty_col = idx.get("qty")
            header_row = i
            break
    if price_col is None:
        return out

    exclude = {price_col}
    if qty_col is not None:
        exclude.add(qty_col)

    for row in table[header_row + 1:]:
        if price_col >= len(row):
            continue
        price = _parse_price(row[price_col])
        if price is None:
            continue  # 単価が無い行(継続説明・合計など)は除外
        item = _pick_item_text(row, exclude)
        # 行番号がセルに混入した場合(例 '13 SHIPXOL_M01-205')、先頭の番号を除去
        item = re.sub(r"^\d{1,3}[\s　]+(?=\S)", "", item)
        key = _normalize_model(item)
        if not key:
            continue
        out[key] = HistoryRecord(name=key, vendor=vendor, price=price)
    return out


def parse_quote_pdf(source, vendor_hint: Optional[str] = None, *, filename: Optional[str] = None) -> tuple[str, dict[str, HistoryRecord]]:
    """
    PDF見積書から (発注先, {型式: HistoryRecord}) を抽出する(ベストエフォート)。
    source はファイルパス(str)・bytes・file-like のいずれでも可。発注先は呼び出し側で
    サブフォルダ名(vendor_hint)を渡すのが最も確実。pdfplumber が必要。

    ※ 見積書の様式はメーカーごとに異なるため、実物に合わせた較正が必要になることがあります。
    """
    import pdfplumber

    if vendor_hint:
        vendor = vendor_hint
    elif filename:
        vendor = _vendor_from_filename(filename)
    elif isinstance(source, str):
        vendor = _vendor_from_filename(source)
    else:
        vendor = ""

    if isinstance(source, bytes):
        source = io.BytesIO(source)

    records: dict[str, HistoryRecord] = {}
    with pdfplumber.open(source) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                # まず見積書ヒューリスティック、ダメなら汎用表パーサ
                recs = _records_from_quote_table(table, vendor)
                if not recs:
                    recs = _records_from_table(table, vendor)
                records.update(recs)
    return vendor, records


# ----------------------------------------------------------------------------
# 共有DriveフォルダのWeb URLから走査(Drive for Desktop不要)
# ----------------------------------------------------------------------------
def extract_folder_id(url: str) -> str:
    """DriveフォルダのURLからフォルダIDを取り出す。"""
    m = re.search(r"/folders/([a-zA-Z0-9\-_]+)", url)
    if not m:
        m = re.search(r"[?&]id=([a-zA-Z0-9\-_]+)", url)
    if not m:
        raise ValueError("DriveフォルダのURLからフォルダIDを取得できませんでした。")
    return m.group(1)


def _http_get(url: str, timeout: int = 60) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _fix_mojibake(s):
    """_DRIVE_ivd を unicode_escape 復号した際の latin-1 byte列を正しいUTF-8へ戻す。"""
    if not isinstance(s, str):
        return s
    try:
        return s.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return s


def _list_public_drive_folder(folder_id: str) -> list[dict]:
    """
    公開(リンク共有)されたDriveフォルダのWebページを解析し、
    エントリ情報 [{id, name, mime}] の一覧を返す(サブフォルダも含む)。
    Google Cloud / API 不要。

    ※ Google非公式の手法(ページ内の _DRIVE_ivd データを解析)のため、
       Drive側の仕様変更で動かなくなる可能性があります。
    """
    html = _http_get(f"https://drive.google.com/drive/folders/{folder_id}").decode(
        "utf-8", errors="replace"
    )
    m = re.search(r"window\['_DRIVE_ivd'\]\s*=\s*'(.*?)';", html, re.DOTALL)
    if not m:
        if "accounts.google.com/v3/signin" in html or "/ServiceLogin" in html:
            raise PermissionError(
                "フォルダを取得できませんでした。フォルダの共有設定を"
                "「リンクを知っている全員が閲覧可」にしてください。"
            )
        raise RuntimeError(
            "フォルダの内容を解析できませんでした(Drive側の仕様変更の可能性)。"
            "代わりに『Googleドライブ パソコン版』のローカルパス指定をご利用ください。"
        )
    decoded = m.group(1).encode("utf-8").decode("unicode_escape")
    data = json.loads(decoded)

    entries = []
    for entry in (data[0] or []):
        try:
            fid, name, mime = entry[0], _fix_mojibake(entry[2]), entry[3]
        except (IndexError, TypeError):
            continue
        entries.append({"id": fid, "name": name, "mime": mime})
    return entries


def _download_drive_file(file_id: str) -> bytes:
    """Drive上のファイルIDから実体(bytes)をダウンロードする。"""
    return _http_get(f"https://drive.google.com/uc?id={file_id}&export=download")


def scan_drive_folder_url(
    url: str,
    *,
    conflict: str = "latest",
    max_depth: int = 5,
) -> tuple[dict[str, HistoryRecord], ScanReport]:
    """
    共有DriveフォルダのURLから、中のファイル(サブフォルダ含む)を取得して履歴を統合する。
    Drive for Desktop / Google Cloud 不要(フォルダはリンク共有・閲覧可であること)。

    フォルダがメーカー別サブフォルダで構成されている場合、
    そのサブフォルダ名を「発注先」として採用する。
    """
    folder_id = extract_folder_id(url)
    history: dict[str, HistoryRecord] = {}
    report = ScanReport()

    def merge(sub):
        for k, v in sub.items():
            if conflict == "first":
                history.setdefault(k, v)
            else:
                history[k] = v

    def walk(fid: str, vendor: str, depth: int):
        if depth > max_depth:
            return
        for e in _list_public_drive_folder(fid):
            name, mime, eid = e["name"], e["mime"], e["id"]
            if mime == _GAPP_FOLDER:
                # サブフォルダ名を発注先として、その中へ降りる
                walk(eid, name, depth + 1)
                continue
            ext = os.path.splitext(name)[1].lower()
            if not (mime == _GAPP_SHEET or ext in TABLE_EXT or ext in PDF_EXT):
                continue
            report.files_seen += 1
            try:
                if mime == _GAPP_SHEET:
                    synthetic = f"https://docs.google.com/spreadsheets/d/{eid}/edit"
                    sub = core.load_history_from_csv_url(synthetic)
                    _override_vendor(sub, vendor)
                    _stamp_source(sub, name)
                    merge(sub)
                    report.gsheets_loaded += 1
                    report.sources.append(name)
                elif ext == ".csv":
                    raw = _download_drive_file(eid)
                    import csv as _csv
                    text = raw.decode("utf-8-sig", errors="replace")
                    sub = core._build_history_from_rows(list(_csv.reader(io.StringIO(text))))
                    _override_vendor(sub, vendor)
                    _stamp_source(sub, name)
                    merge(sub)
                    report.tables_loaded += 1
                    report.sources.append(name)
                elif ext in {".xlsx", ".xlsm"}:
                    raw = _download_drive_file(eid)
                    sub = core.load_history_from_file(raw)
                    _override_vendor(sub, vendor)
                    _stamp_source(sub, name)
                    merge(sub)
                    report.tables_loaded += 1
                    report.sources.append(name)
                elif ext == ".pdf":
                    raw = _download_drive_file(eid)
                    used_vendor, sub = parse_quote_pdf(raw, vendor_hint=vendor or None, filename=name)
                    if not sub:
                        report.skipped.append(
                            (name, f"発注先{used_vendor or '?'}: 明細を抽出できませんでした"
                                    "(画像PDF/様式差の可能性。OCRまたは表形式が必要)")
                        )
                    else:
                        _stamp_source(sub, name)
                        merge(sub)
                        report.pdfs_loaded += 1
                        report.pdf_items += len(sub)
                        report.sources.append(f"{name} (発注先: {used_vendor or '不明'}, {len(sub)}件)")
            except PermissionError as e2:
                report.skipped.append((name, str(e2)))
            except Exception as e2:  # noqa: BLE001
                report.errors.append((name, f"{type(e2).__name__}: {e2}"))

    walk(folder_id, vendor="", depth=0)
    report.records_total = len(history)
    return history, report


def _override_vendor(sub: dict[str, HistoryRecord], vendor: str):
    """構造化ファイル(表)では、発注先が空のときだけサブフォルダ名で補う。
    ファイル自身に発注先列があればそちらを優先する。"""
    if not vendor:
        return
    for rec in sub.values():
        if not rec.vendor:
            rec.vendor = vendor
