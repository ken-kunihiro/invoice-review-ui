"""
請求書OCRレビューサーバー
Usage: python server.py --folder <証憑フォルダパス> [--port 3470] [--no-browser]
"""
import argparse
import asyncio
import base64
import csv
import hashlib
import io
import json
import mimetypes
import os
import pathlib
import re
import shutil
import sys
import threading
import time
import unicodedata
import webbrowser
from collections import Counter
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse
from urllib.request import urlopen

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import fitz  # PyMuPDF。OCRマーカー表示（/api/highlight）のページ画像化・テキスト検索に使う
    HAS_FITZ = True
except ImportError:
    fitz = None
    HAS_FITZ = False

try:
    from PIL import Image  # 画像ファイルの実サイズ取得用（/api/highlight）
    HAS_PIL = True
except ImportError:
    Image = None
    HAS_PIL = False

try:
    import vision_ocr
    HAS_VISION_OCR = True
except ImportError:
    vision_ocr = None
    HAS_VISION_OCR = False

try:
    # Windows内蔵OCR（winrt経由）。スキャンPDF・画像の単語ごとの実測位置を取得するために使う
    # （/api/highlight のマーカー精度向上。サブPC等で未導入の場合は静かにスキップしAI推定へ）
    from winrt.windows.globalization import Language as _WinLanguage
    from winrt.windows.graphics.imaging import BitmapDecoder as _WinBitmapDecoder
    from winrt.windows.media.ocr import OcrEngine as _WinOcrEngine
    from winrt.windows.storage.streams import (InMemoryRandomAccessStream as _WinStream,
                                                DataWriter as _WinDataWriter)
    HAS_WINOCR = True
except ImportError:
    HAS_WINOCR = False

SCRIPT_DIR = pathlib.Path(__file__).parent
INDEX_HTML = SCRIPT_DIR / "index.html"
DATA_FILE_NAME = "_invoice_review.json"
SUPPORTED_EXTS = {".pdf", ".jpg", ".jpeg", ".png"}

# 取引先（発行元）抽出
COMPANY_SUFFIXES = ("株式会社", "合同会社", "有限会社", "合名会社", "合資会社",
                    "一般社団法人", "一般財団法人", "公益社団法人", "公益財団法人",
                    "社会福祉法人", "医療法人", "特定非営利活動法人", "NPO法人")
COMPANY_SUFFIXES_SORTED = sorted(COMPANY_SUFFIXES, key=len, reverse=True)
NAME_CHARS = r"[A-Za-zＡ-Ｚａ-ｚ一-龥ぁ-んァ-ヶー0-9０-９・&'\. 　]"
NAME_LABEL = re.compile(r'(?:Name|name|氏名|担当)\s*[:：]?\s*(.*)')
PERSON_RE = re.compile(r'^[\s　]*([一-龥]{1,4}[\s　]?[一-龥]{1,4})[\s　]*$')
BANK_LINE = re.compile(r'(銀行|信用金庫|信用組合|労働金庫|農協|ゆうちょ)')
ACCOUNT_NAME_LABEL = re.compile(r'(?:口座名義|名義人|名義)\s*[:：/／]?\s*(.+)')
ADDRESS_NG = re.compile(r'(〒|[都道府県].{0,8}[市区町村]|丁目|番地|TEL|Tel|FAX|Email|@|銀行|支店|口座|登録番号)')
TITLE_NG = re.compile(r'^(御?請求書|見積書|納品書|領収書|注文書|発注書|請求|発行|合計|小計|品名|摘要|数量|単価|金額|備考|内訳|明細)')

# 金額（通貨マーカー付き数字のみ。電話番号・郵便番号を弾く）
CURRENCY_NUM = re.compile(r'[¥￥]\s*([\d,]+)|([\d,]+)\s*円')
TOTAL_KEYWORDS = re.compile(r'(合計金額|ご請求金額|御請求金額|請求金額|お支払金額|請求額|御請.*金額)')
SUBTOTAL_NG = re.compile(r'(小計|消費税|内訳|税抜|税率)')

# 明細（品名）抽出
ITEM_HEADER = re.compile(r'(品[\s　]*名|品目|摘要|内容|内容・仕様)')
ITEM_HEADER_AMT = re.compile(r'(金[\s　]*額|価格|明細金額)')
ITEM_STOP = re.compile(r'^(小計|合計|消費税|税率|内訳|備考|お?振込|【|A[：:]|A\+B|うち消費税|非課税|支払|検印)')
NUM_TOKEN = re.compile(r'[¥￥]?([\d,]{2,})')
UNIT_WORDS = re.compile(r'(一式|式|個|時間|時|本|枚|点|か所|箇所|セット|単位|円)')

# 日付パターン
DATE_PATTERNS = [
    re.compile(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日'),
    re.compile(r'(\d{4})/(\d{1,2})/(\d{1,2})'),
    re.compile(r'(\d{4})-(\d{1,2})-(\d{1,2})'),
]
# 源泉パターン
WITHHOLDING_PATTERNS = [
    re.compile(r'源泉[徴収税額所得税]*[\s　]*[-△▲¥￥][\s　]*([\d,]+)'),
    re.compile(r'源泉[徴収税額所得税]*[\s　]*([\d,]+)'),
]
# 勘定科目予測（品名・取引先・本文のキーワード → 候補科目）
# 上から順に優先。最初にヒットした科目を「予測（要確認）」として提示する。
# 税務判断はしないため、あくまで編集可能な候補。グレーは無理に埋めない。
ACCOUNT_KEYWORDS = [
    ("支払報酬料", ["税理士", "会計士", "公認会計士", "社労士", "社会保険労務士",
                "弁護士", "司法書士", "行政書士", "弁理士", "顧問料", "監査報酬"]),
    ("旅費交通費", ["旅費", "交通費", "特急料金", "新幹線", "航空券", "宿泊", "ホテル",
                "タクシー", "出張", "ガソリン", "高速代", "駐車場", "電車", "バス代",
                "etc", "ｅｔｃ"]),
    ("広告宣伝費", ["広告", "宣伝", "プロモーション", "マーケティング費", "出稿", "掲載料",
                "リスティング", "prtimes", "pr times"]),
    # 「お振込手数料はご負担」等の定型文は除外。役務としての手数料のみ拾う
    ("支払手数料", ["決済手数料", "為替手数料", "stripe", "ストライプ",
                "paypal", "ペイパル", "収納代行"]),
    ("地代家賃", ["家賃", "賃料", "地代", "テナント", "レンタルオフィス",
                "バーチャルオフィス", "コワーキング"]),
    ("通信費", ["ソフトウェア保守", "システム利用料", "サーバー", "ドメイン", "クラウド",
              "ライセンス", "サブスク", "saas", "回線", "通信費", "通信料",
              "openai", "anthropic", "claude", "aws", "google", "figma",
              "notion", "slack", "adobe", "zoom", "github"]),
    ("会議費", ["会議費", "打合せ", "打ち合わせ", "ミーティング", "カフェ"]),
    ("交際費", ["接待", "会食", "懇親", "贈答", "手土産"]),
    ("新聞図書費", ["書籍", "図書", "新聞", "購読", "資料代"]),
    ("消耗品費", ["消耗品", "事務用品", "文具", "備品", "印刷", "名刺"]),
    ("業務委託費", ["業務委託"]),
    ("外注費", ["サポート報酬", "サポ報酬", "mod報酬", "制作", "デザイン", "開発",
              "コーディング", "ライティング", "執筆", "編集", "動画編集", "撮影",
              "運用代行", "外注", "委託料", "代行"]),
]


def predict_account(partner: str, item_names: list, full_text: str) -> str:
    """品名・取引先・本文から勘定科目を予測。確信が持てなければ空文字（要確認）。

    あくまで編集可能な候補。税務判断はせず、グレーは埋めない。
    取引先＋品名を主に見て、補助的に本文先頭を見る。
    """
    primary = (partner + " " + " ".join(item_names)).lower()
    fallback = (primary + " " + full_text[:600]).lower()
    for account, keywords in ACCOUNT_KEYWORDS:
        if any(kw.lower() in primary for kw in keywords):
            return account
    for account, keywords in ACCOUNT_KEYWORDS:
        if any(kw.lower() in fallback for kw in keywords):
            return account
    return ""


# 外貨パターン
FOREIGN_PATTERN = re.compile(r'USD|EUR|CNY|\$|€')
# キーワード近傍検出用
ISSUE_KEYWORDS = re.compile(r'(請求日|発行日|発行年月日)')
DUE_KEYWORDS = re.compile(r'(支払期限|お支払期日|振込期限)')
WITHHOLDING_KEYWORDS = re.compile(r'(源泉|所得税)')


def file_id(path: pathlib.Path) -> str:
    return hashlib.md5(str(path).encode()).hexdigest()[:12]


def unique_path(dest: pathlib.Path) -> pathlib.Path:
    if not dest.exists():
        return dest
    stem, suffix = dest.stem, dest.suffix
    for i in range(1, 1000):
        cand = dest.with_name(f"{stem} ({i}){suffix}")
        if not cand.exists():
            return cand
    raise FileExistsError(str(dest))


def parse_date(text: str) -> str:
    for pat in DATE_PATTERNS:
        m = pat.search(text)
        if m:
            y, mo, d = m.group(1), m.group(2), m.group(3)
            return f"{y}/{int(mo):02d}/{int(d):02d}"
    return ""


def find_date_near_keyword(lines: list, keyword_re) -> str:
    for i, line in enumerate(lines):
        if keyword_re.search(line):
            context = "\n".join(lines[max(0, i-1):i+3])
            d = parse_date(context)
            if d:
                return d
    return ""


def _marked_nums(line: str) -> list:
    """通貨マーカー（¥/￥/円）付きの数値だけを返す"""
    vals = []
    for m in CURRENCY_NUM.finditer(line):
        s = m.group(1) or m.group(2)
        if s:
            vals.append(int(s.replace(",", "")))
    return vals


def extract_amount(lines: list) -> tuple:
    """(金額int or None, 確信度bool)。振込合計＝マーカー付き数字。電話番号は除外"""
    kw = []
    for line in lines:
        if TOTAL_KEYWORDS.search(line) and not SUBTOTAL_NG.search(line):
            kw += _marked_nums(line)
    if kw:
        return max(kw), True
    allnums = []
    for line in lines:
        allnums += _marked_nums(line)
    if allnums:
        return max(allnums), False
    return None, False


def extract_withholding(lines: list) -> int:
    text = "\n".join(lines)
    if not WITHHOLDING_KEYWORDS.search(text):
        return 0
    for pat in WITHHOLDING_PATTERNS:
        m = pat.search(text)
        if m:
            return int(m.group(1).replace(",", ""))
    return 0


def _clean(s: str) -> str:
    return s.replace("　", " ").strip()


def split_by_marker(line: str) -> list:
    """行を御中/様の直後で区切り、(断片, 宛先か) のリストを返す"""
    res, last = [], 0
    for m in re.finditer(r'(御中|様)', line):
        res.append((line[last:m.end()], True))
        last = m.end()
    if last < len(line):
        res.append((line[last:], False))
    return res or [(line, False)]


def find_companies(part: str) -> list:
    """断片から会社名を組み立てて返す（前置・後置両対応）"""
    names = []
    for suf in COMPANY_SUFFIXES_SORTED:
        start = 0
        while True:
            i = part.find(suf, start)
            if i < 0:
                break
            am = re.match(NAME_CHARS + r'{1,25}', part[i+len(suf):])
            bm = re.search(NAME_CHARS + r'{1,25}$', part[:i])
            after_name = _clean(am.group(0)) if am else ""
            before_name = _clean(bm.group(0)) if bm else ""
            if after_name:
                names.append(_clean(suf + after_name))
            elif before_name:
                names.append(_clean(before_name + suf))
            start = i + len(suf)
    return names


def extract_partner(lines: list) -> str:
    """取引先＝発行元。宛先（御中/様）は除外し、会社→Name:→口座名義→振込先→個人名の順で拾う"""
    issuers, recipients = [], set()
    for line in lines:
        for part, is_recip in split_by_marker(line):
            for name in find_companies(part):
                (recipients.add(name) if is_recip else issuers.append(name))
    for n in issuers:
        if n not in recipients:
            return n
    # Name: ラベル
    for i, line in enumerate(lines):
        m = NAME_LABEL.search(line)
        if m:
            val = _clean(m.group(1))
            if val and not ADDRESS_NG.search(val) and not val.startswith(("[", "〒")):
                return val
            for nxt in lines[i+1:i+3]:
                nv = _clean(nxt)
                if nv and not ADDRESS_NG.search(nv):
                    return nv
    # 口座名義ラベル
    for line in lines:
        m = ACCOUNT_NAME_LABEL.search(line)
        if m:
            val = re.sub(r'^[/／\s]*', '', _clean(m.group(1)))
            if val:
                return val
    # 振込先の銀行行 末尾の個人名（口座番号のあと）
    for line in lines:
        if BANK_LINE.search(line):
            m = re.search(r'\d{6,8}[\s　]+([一-龥]{1,4}[\s　]?[一-龥]{1,4})\s*$', line)
            if m:
                return _clean(m.group(1))
    # 上部の個人名ブロック（宛先・住所・タイトル行でない姓名）
    for line in lines[:15]:
        if ADDRESS_NG.search(line) or "御中" in line or "様" in line or TITLE_NG.match(_clean(line)):
            continue
        pm = PERSON_RE.match(line)
        if pm:
            return _clean(pm.group(1))
    return ""


def extract_line_items(lines: list, total: int) -> list:
    """品名と明細金額を抽出。合計が請求額と整合する場合のみ採用（帳尻合わせ防止）"""
    items = []
    started = False
    for line in lines:
        l = _clean(line)
        if not l:
            continue
        if not started:
            if ITEM_HEADER.search(l) and (ITEM_HEADER_AMT.search(l) or "単価" in l or "数量" in l):
                started = True
            continue
        if ITEM_STOP.match(l):
            break
        nums = [int(x.replace(",", "")) for x in NUM_TOKEN.findall(l)]
        if not nums:
            continue
        # 品名抽出: 日付（2/3, 2/18, 2/3〜2/4）は品名の一部なので保護してから数字列を除去
        holds = []
        def _hold(m):
            holds.append(m.group(0))
            return f"\x00{len(holds)-1}\x00"
        tmp = re.sub(r'\d{1,2}/\d{1,2}(?:[〜~\-]\d{1,2}/\d{1,2})?', _hold, l)
        name = re.sub(r'[¥￥]?[\d,]{2,}', '', tmp)       # 単価・金額（2桁以上）を除去
        name = UNIT_WORDS.sub('', name)                  # 単位語（式・個・時間…）を除去
        name = re.sub(r'(?:^|\s)\d(?=\s|$)', ' ', name)  # 空白で挟まれた1桁の数量を除去
        name = re.sub(r'\x00(\d+)\x00', lambda m: holds[int(m.group(1))], name)  # 日付を復元
        name = re.sub(r'\s+', ' ', name)
        name = _clean(name).strip("・/／-－ ")
        amount = max(nums)
        if name and amount >= 10:
            items.append({"name": name[:40], "amount": amount, "account": "",
                          "tax_class": "課対仕入10%", "tax_calc": "内税"})
    # 整合チェック: 明細合計が請求額（税込）または税抜相当に一致し、2件以上なら採用。
    # 税抜内訳しか拾えない請求書（小計＝税抜）でも明細化できるよう税抜ベースも許容。
    # 税抜で一致した場合は各行を「外税」にし、freee側で消費税を足して総額一致させる。
    # 帳尻合わせはせず、いずれかに一致した場合のみ採用する。
    if total and len(items) >= 2:
        s = sum(i["amount"] for i in items)
        if abs(s - total) <= max(1, int(total * 0.02)):
            return items  # 税込内訳 → 内税のまま
        for base in (round(total / 1.1), round(total / 1.08)):
            if abs(s - base) <= max(1, int(base * 0.02)):
                for it in items:
                    it["tax_calc"] = "外税"  # 税抜内訳 → 外税で出力
                return items
    return []


def ocr_pdf(path: pathlib.Path) -> dict:
    result = {"partner": "", "issue_date": "", "due_date": "",
              "amount": 0, "withholding": 0, "memo": "", "status": "check",
              "line_items": []}
    if not HAS_PDFPLUMBER:
        return result
    try:
        with pdfplumber.open(str(path)) as pdf:
            full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        result["memo"] = f"PDF読み取りエラー: {e}"
        return result

    if not full_text.strip():
        return result

    # 外貨チェック
    if FOREIGN_PATTERN.search(full_text):
        result["memo"] = "外貨建て・換算要"
        result["status"] = "check"
        return result

    lines = full_text.splitlines()

    partner = extract_partner(lines)
    issue_date = find_date_near_keyword(lines, ISSUE_KEYWORDS) or parse_date(full_text)
    due_date = find_date_near_keyword(lines, DUE_KEYWORDS)
    amount, certain = extract_amount(lines)
    withholding = extract_withholding(lines)
    line_items = extract_line_items(lines, amount or 0)

    # 勘定科目を予測（要確認の候補。明細ごとにも予測）
    item_names = [i["name"] for i in line_items]
    # 明細ごとの予測は品名のみで判定。確信が持てなければ空欄（要確認）。
    # 全科目を一律に埋めると誤りが伝播するため fallback はしない。
    for it in line_items:
        it["account"] = predict_account("", [it["name"]], "")
    # フォーム上段の科目は、明細があれば金額最大の行（主たる役務）を基準に予測。
    # 経費立替等の少額行に引っ張られないようにする。明細がなければ取引先＋本文で予測。
    if line_items:
        dominant = max(line_items, key=lambda i: i["amount"])
        account = dominant["account"] or predict_account(partner, item_names, full_text)
    else:
        account = predict_account(partner, item_names, full_text)

    result["partner"] = partner
    result["issue_date"] = issue_date
    result["due_date"] = due_date
    result["amount"] = amount or 0
    result["withholding"] = withholding
    result["line_items"] = line_items
    result["account"] = account

    if withholding > 0:
        result["status"] = "gensen"
    elif partner and (amount or 0) > 0:
        result["status"] = "ok" if certain else "check"
    else:
        result["status"] = "check"

    return result


def ai_ocr_available() -> bool:
    return bool(HAS_VISION_OCR and vision_ocr.is_available())


def _iso_to_slash(s) -> str:
    """"YYYY-MM-DD" (AI-OCR出力) -> "YYYY/MM/DD"（本ツールの日付表記）"""
    if not s:
        return ""
    m = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', str(s))
    if not m:
        return ""
    y, mo, d = m.groups()
    return f"{y}/{int(mo):02d}/{int(d):02d}"


def apply_ai_ocr_to_record(rec: dict, path: pathlib.Path) -> tuple:
    """AI-OCR結果をレコードにマッピングする。

    戻り値: (成功したか, エラーメッセージ or None)
    失敗時も rec は壊さず、ai_notes にエラー内容だけ残す。
    """
    if not HAS_VISION_OCR:
        return False, "vision_ocr未導入"
    prev_status = rec.get("status")
    result = vision_ocr.ocr_file(path)
    if result.get("error"):
        rec["ai_notes"] = f"AI-OCR失敗: {result['error']}"
        return False, result["error"]
    # 読取値が更新されるため、旧値ベースのマーカー位置キャッシュを破棄
    rec.pop("ai_bboxes", None)

    vendor = result.get("vendor") or ""
    amount = result.get("amount_total") or 0
    issue = _iso_to_slash(result.get("service_month_end") or result.get("issue_date"))
    due = _iso_to_slash(result.get("due_date"))
    withholding = result.get("withholding_tax") or 0
    confidence = result.get("confidence") or "mid"
    notes = result.get("notes") or ""

    line_items = []
    for it in result.get("line_items") or []:
        amt = int(it.get("amount") or 0)
        desc = it.get("description") or ""
        if amt <= 0 and not desc:
            continue
        line_items.append({
            "name": desc[:40],
            "account": it.get("account_item") or "",
            "tax_class": it.get("tax_class") or "課対仕入10%",
            "tax_calc": "内税",
            "amount": amt,
        })

    if vendor:
        rec["partner"] = vendor
    if amount:
        rec["amount"] = int(amount)
    if issue:
        rec["issue_date"] = issue
    if due:
        rec["due_date"] = due
    if withholding:
        rec["withholding"] = int(withholding)
    items_sum_ok = True
    if line_items:
        rec["line_items"] = line_items
        dominant = max(line_items, key=lambda i: i["amount"])
        if dominant.get("account"):
            rec["account"] = dominant["account"]
        if dominant.get("tax_class"):
            rec["tax_class"] = dominant["tax_class"]
        # 明細合計と税込合計の突合（AI読取時は全行内税なので単純合計でよい）
        items_sum = sum(i["amount"] for i in line_items)
        if amount and items_sum != int(amount):
            items_sum_ok = False
            notes = (notes + f" ※明細合計{items_sum}円が合計{int(amount)}円と不一致（要修正）").strip()

    if notes and not (rec.get("memo") or "").strip():
        rec["memo"] = notes

    rec["ai_confidence"] = confidence
    rec["ai_notes"] = notes
    rec["source"] = "ai_ocr"

    if confidence == "low":
        rec["status"] = "check"
    elif withholding > 0:
        rec["status"] = "gensen"
    elif rec.get("partner") and (rec.get("amount") or 0) > 0:
        rec["status"] = "ok"
    else:
        rec["status"] = "check"

    # 確認済みレコードを明示的に再読取した場合、確信度が高ければ確認済みのまま維持する
    # （再読取のたびに確認が解除されると出力対象から外れてしまうため）
    if prev_status == "confirmed" and confidence != "low":
        rec["status"] = "confirmed"

    # 明細合計が合わないままCSVに乗せない（帳尻合わせ禁止。人の修正を挟む）
    if not items_sum_ok:
        rec["status"] = "check"

    return True, None


# ===== OCR根拠マーカー表示（/api/highlight） =====
# 令和換算: 令和N年 = 西暦(2018+N)年
WAREKI_ERA_START = 2018


def _date_search_variants(date_str: str) -> list:
    """"YYYY/MM/DD" 形式の日付から、請求書上の表記ゆれ検索バリアントを順に生成"""
    m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})', date_str or "")
    if not m:
        return []
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    variants = [
        f"{y}/{mo:02d}/{d:02d}",
        f"{y}/{mo}/{d}",
        f"{y}-{mo:02d}-{d:02d}",
        f"{y}年{mo}月{d}日",
        f"{y}年{mo:02d}月{d:02d}日",
    ]
    reiwa = y - WAREKI_ERA_START
    if reiwa >= 1:
        variants.append(f"令和{reiwa}年{mo}月{d}日")
    return variants


def _amount_search_variants(amount: int) -> list:
    """金額の検索バリアント（カンマ区切り→¥付き→カンマなしの順）"""
    if not amount:
        return []
    s = f"{int(amount):,}"
    return [s, f"¥{s}", str(int(amount))]


def _partner_search_variants(partner: str) -> list:
    """取引先名の検索バリアント（そのまま→空白除去→法人格を除いた本体）"""
    if not partner:
        return []
    variants = [partner]
    stripped = partner.replace(" ", "").replace("　", "")
    if stripped not in variants:
        variants.append(stripped)
    for suf in COMPANY_SUFFIXES_SORTED:
        if suf in partner:
            body = partner.replace(suf, "").strip()
            if body and body not in variants:
                variants.append(body)
            break
    return variants


def _highlight_field_variants(rec: dict) -> dict:
    """レコードの読取値から、各項目の検索バリアントリストを組み立てる（値が無い項目は含めない）"""
    variants = {
        "partner": _partner_search_variants(rec.get("partner", "")),
        "issue_date": _date_search_variants(rec.get("issue_date", "")),
        "due_date": _date_search_variants(rec.get("due_date", "")),
        "amount": _amount_search_variants(rec.get("amount") or 0),
    }
    if rec.get("withholding"):
        variants["withholding"] = _amount_search_variants(rec.get("withholding"))
    return {k: v for k, v in variants.items() if v}


def _search_first_hit(page, variants: list) -> list:
    """バリアントを順に試し、最初にヒットしたRectのリストを返す（見つからなければ空リスト）"""
    for v in variants:
        try:
            rects = page.search_for(v)
        except Exception:
            rects = []
        if rects:
            return rects
    return []


def _highlight_from_text_pdf(path: pathlib.Path, rec: dict):
    """テキスト層のあるPDFから各項目の正確な位置を検索する。

    1項目もヒットしなければ None を返し、呼び出し側でAI推定（_highlight_from_ai）へ
    フォールバックさせる。
    """
    field_variants = _highlight_field_variants(rec)
    if not field_variants:
        return None
    try:
        doc = fitz.open(str(path))
    except Exception:
        return None

    chosen_page_rects = None
    try:
        for page in doc:
            page_rects = {}
            for field, variants in field_variants.items():
                rects = _search_first_hit(page, variants)
                if rects:
                    page_rects[field] = rects[:5] if field == "amount" else rects[:1]
            if page_rects:
                chosen_page_rects = (page, page_rects)
                break

        if chosen_page_rects is None:
            return None

        page, page_rects = chosen_page_rects
        zoom = 2
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        png_bytes = pix.tobytes("png")
        width, height = pix.width, pix.height
    finally:
        doc.close()

    boxes = []
    for field, rects in page_rects.items():
        for r in rects:
            boxes.append({
                "field": field,
                "x0": r.x0 * zoom, "y0": r.y0 * zoom,
                "x1": r.x1 * zoom, "y1": r.y1 * zoom,
            })
    not_found = [f for f in field_variants if f not in page_rects]

    return {
        "ok": True,
        "image": "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii"),
        "width": width,
        "height": height,
        "approx": False,
        "boxes": boxes,
        "not_found": not_found,
    }


def _highlight_background_image(path: pathlib.Path):
    """マーカーを重ねる背景画像を作る。PDFはページ1をPNG化、画像はファイルそのまま。

    戻り値: (data_uri, width, height) または例外
    """
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        doc = fitz.open(str(path))
        try:
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
            png_bytes = pix.tobytes("png")
            width, height = pix.width, pix.height
        finally:
            doc.close()
        data_uri = "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii")
        return data_uri, width, height

    # JPG/PNG: ファイルそのままを返す（再エンコードによる劣化を避ける）
    raw = path.read_bytes()
    mime = "image/jpeg" if suffix in (".jpg", ".jpeg") else "image/png"
    if HAS_PIL:
        with Image.open(path) as im:
            width, height = im.size
    else:
        # PIL未導入時のフォールバック（fitzで再エンコードしてサイズだけ取得）
        doc = fitz.open(str(path))
        try:
            pix = doc[0].get_pixmap()
            width, height = pix.width, pix.height
        finally:
            doc.close()
    data_uri = f"data:{mime};base64," + base64.b64encode(raw).decode("ascii")
    return data_uri, width, height


def _highlight_from_ai(path: pathlib.Path, rec: dict, office_name: str) -> dict:
    """AI-OCR（Claude Vision）で各項目のおおよその位置を推定する（approx: true）。

    一度推定した位置は rec["ai_bboxes"] にキャッシュし、2回目以降はAPIを呼ばない。
    """
    if not ai_ocr_available():
        return {"ok": False, "error": "AI-OCRが利用できません（.envにANTHROPIC_API_KEYまたはCLAUDE_API_KEYが必要です。またはClaude Codeのインストールが必要です）"}

    fields = {
        "partner": rec.get("partner") or "",
        "issue_date": rec.get("issue_date") or "",
        "due_date": rec.get("due_date") or "",
        "amount": rec.get("amount") or 0,
    }
    if rec.get("withholding"):
        fields["withholding"] = rec.get("withholding")

    cached = rec.get("ai_bboxes")
    if cached is not None:
        boxes_norm = cached
    else:
        result = vision_ocr.locate_fields(path, fields)
        if result.get("error"):
            return {"ok": False, "error": result["error"]}
        boxes_norm = result.get("boxes") or {}
        with LOCK:
            rec["ai_bboxes"] = boxes_norm
            save_office(office_name)

    try:
        image_data_uri, width, height = _highlight_background_image(path)
    except Exception as e:
        return {"ok": False, "error": f"背景画像の生成に失敗しました: {e}"}

    boxes = []
    not_found = []
    for field in fields:
        b = boxes_norm.get(field)
        if b:
            x0, y0, x1, y1 = b
            boxes.append({
                "field": field,
                "x0": x0 / 1000 * width, "y0": y0 / 1000 * height,
                "x1": x1 / 1000 * width, "y1": y1 / 1000 * height,
            })
        else:
            not_found.append(field)

    return {
        "ok": True,
        "image": image_data_uri,
        "width": width,
        "height": height,
        "approx": True,
        "boxes": boxes,
        "not_found": not_found,
    }


def _ocr_normalize(s: str) -> str:
    """OCR実測位置マッチング用の正規化。全角英数字→半角、空白・カンマ・¥記号を除去"""
    s = unicodedata.normalize("NFKC", s or "")
    return re.sub(r"[\s,¥￥]", "", s)


async def _win_ocr_recognize(png_bytes: bytes):
    """WindowsOCRエンジンでPNGバイト列を認識する（言語パック等が無ければNoneを返す）"""
    stream = _WinStream()
    writer = _WinDataWriter(stream.get_output_stream_at(0))
    writer.write_bytes(png_bytes)
    await writer.store_async()
    decoder = await _WinBitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()

    engine = _WinOcrEngine.try_create_from_language(_WinLanguage("ja"))
    if engine is None:
        engine = _WinOcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        return None
    return await engine.recognize_async(bitmap)


def _win_ocr_render_png(path: pathlib.Path):
    """OCR入力・画面表示の両方に使うPNGを生成する（座標系を完全一致させるため共用）。

    戻り値: (png_bytes, width, height)
    """
    max_dim = _WinOcrEngine.max_image_dimension
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        doc = fitz.open(str(path))
        try:
            page = doc[0]
            # 基本zoom=2.0だとレシート等の小さいページで文字が潰れてOCR誤読が
            # 多発したため（file_1325で検証）、OCR用途の標準的な300dpi相当
            # （zoom=300/72≈4.17）を基本値にする。max_image_dimensionは超えない。
            zoom = 300.0 / 72.0
            w, h = page.rect.width, page.rect.height
            if max(w, h) * zoom > max_dim:
                zoom = max_dim / max(w, h)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            png_bytes = pix.tobytes("png")
            return png_bytes, pix.width, pix.height
        finally:
            doc.close()

    # JPG/PNG: PILで開いてPNG化（拡大はしない。縮小はOCR最大サイズを超える場合のみ）
    with Image.open(path) as im:
        im = im.convert("RGB")
        w, h = im.size
        if max(w, h) > max_dim:
            scale = max_dim / max(w, h)
            im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        return buf.getvalue(), im.size[0], im.size[1]


def _ocr_build_lines(result):
    """OCR結果を行単位に分解し、正規化文字列→単語インデックスの対応表を作る。

    戻り値: [(line_norm, char_to_word, words), ...]
    words の各要素は {"x","y","w","h"}（PNGピクセル座標）
    """
    lines_data = []
    for line in result.lines:
        words = []
        char_to_word = []
        norm_parts = []
        for widx, word in enumerate(line.words):
            r = word.bounding_rect
            words.append({"x": r.x, "y": r.y, "w": r.width, "h": r.height})
            norm = _ocr_normalize(word.text)
            norm_parts.append(norm)
            char_to_word.extend([widx] * len(norm))
        line_norm = "".join(norm_parts)
        lines_data.append((line_norm, char_to_word, words))
    return lines_data


def _ocr_search_first_hit(lines_data: list, variants: list) -> list:
    """バリアントを順に試し、最初にヒットしたバウンディングボックス（複数行分）のリストを返す"""
    for v in variants:
        vn = _ocr_normalize(v)
        if not vn:
            continue
        hits = []
        for line_norm, char_to_word, words in lines_data:
            start = 0
            while True:
                idx = line_norm.find(vn, start)
                if idx == -1:
                    break
                w_start = char_to_word[idx]
                w_end = char_to_word[idx + len(vn) - 1]
                touched = words[w_start:w_end + 1]
                hits.append({
                    "x0": min(w["x"] for w in touched),
                    "y0": min(w["y"] for w in touched),
                    "x1": max(w["x"] + w["w"] for w in touched),
                    "y1": max(w["y"] + w["h"] for w in touched),
                })
                start = idx + len(vn)
        if hits:
            return hits
    return []


def _highlight_from_local_ocr(path: pathlib.Path, rec: dict):
    """Windows内蔵OCRで実測した単語位置から各項目の正確な位置を検索する。

    テキスト層の無いスキャンPDF・画像向け。AI推定より座標精度が高い。
    1項目もヒットしなければ None を返し、呼び出し側でAI推定へフォールバックさせる。
    """
    field_variants = _highlight_field_variants(rec)
    if not field_variants:
        return None

    png_bytes, width, height = _win_ocr_render_png(path)
    result = asyncio.run(_win_ocr_recognize(png_bytes))
    if result is None:
        return None

    lines_data = _ocr_build_lines(result)

    boxes = []
    not_found = []
    for field, variants in field_variants.items():
        hits = _ocr_search_first_hit(lines_data, variants)
        if not hits:
            not_found.append(field)
            continue
        hits = hits[:5] if field == "amount" else hits[:1]
        for h in hits:
            boxes.append({"field": field, "x0": h["x0"], "y0": h["y0"], "x1": h["x1"], "y1": h["y1"]})

    if not boxes:
        return None

    return {
        "ok": True,
        "image": "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii"),
        "width": width,
        "height": height,
        "approx": False,
        "boxes": boxes,
        "not_found": not_found,
    }


def build_highlight(path: pathlib.Path, rec: dict, office_name: str) -> dict:
    """OCR根拠マーカー表示用のレスポンスを組み立てる。

    1. テキスト層のあるPDFはまず正確な位置検索を試みる（PDFのみ）
    2. ヒットしなければ、Windows内蔵OCRによる実測位置検索を試みる（導入されていれば）
    3. それでもヒットしない・例外が出た場合はAI-OCRによる位置推定にフォールバックする
       （マーカー機能はUI補助のため、途中で例外が出ても落とさずAI推定へ逃がす）
    """
    if not HAS_FITZ:
        return {"ok": False, "error": "PyMuPDF(fitz)が見つかりません（pip install pymupdf で有効化）"}

    result = None
    if path.suffix.lower() == ".pdf":
        try:
            result = _highlight_from_text_pdf(path, rec)
        except Exception:
            result = None

    if result is None and HAS_WINOCR:
        try:
            result = _highlight_from_local_ocr(path, rec)
        except Exception:
            result = None

    if result is None:
        result = _highlight_from_ai(path, rec, office_name)
    return result


def _is_archived_path(folder: pathlib.Path, p: pathlib.Path) -> bool:
    """_exported 等、アンダースコア始まりのサブフォルダ配下か"""
    rel = p.relative_to(folder).parts
    return any(part.startswith("_") for part in rel[:-1])


def scan_folder(folder: pathlib.Path, existing: dict) -> list:
    records = []
    existing_paths = {r["path"]: r for r in existing}

    found_paths = set()
    for p in sorted(folder.rglob("*")):
        if p.suffix.lower() not in SUPPORTED_EXTS or DATA_FILE_NAME in p.name:
            continue
        # _exported/ 等のアーカイブ領域は新規スキャン対象外
        if _is_archived_path(folder, p):
            continue
        found_paths.add(str(p))
        if str(p) in existing_paths:
            rec = existing_paths[str(p)]
            rec.pop("missing", None)
            records.append(rec)
        else:
            rec = {
                "id": file_id(p),
                "file": p.name,
                "path": str(p),
                "partner": "",
                "issue_date": "",
                "due_date": "",
                "amount": 0,
                "account": "",
                "tax_class": "課対仕入10%",
                "withholding": 0,
                "memo": "",
                "status": "check",
                "missing": False,
                "line_items": [],
                "source": "none",
                "ai_confidence": None,
                "ai_notes": "",
                "user_edited": False,
            }
            if p.suffix.lower() == ".pdf":
                extracted = ocr_pdf(p)
                rec.update(extracted)
                if extracted.get("partner") and extracted.get("amount"):
                    rec["source"] = "text"
            incomplete = not (rec.get("partner") and rec.get("amount"))
            if incomplete and ai_ocr_available():
                apply_ai_ocr_to_record(rec, p)
            records.append(rec)

    # スキャンに出てこなかった既存レコードの後始末
    for path_str, rec in existing_paths.items():
        if path_str in found_paths:
            continue
        if rec.get("status") == "exported":
            # 出力済み（_exported配下へ移動済み）。missingにせず保持
            rec.pop("missing", None)
            records.append(rec)
        else:
            rec["missing"] = True
            records.append(rec)

    return records


def load_data(folder: pathlib.Path) -> list:
    data_file = folder / DATA_FILE_NAME
    if data_file.exists():
        try:
            return json.loads(data_file.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def save_data(folder: pathlib.Path, records: list):
    data_file = folder / DATA_FILE_NAME
    data_file.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


def most_common_ym(records: list) -> str:
    months = []
    for r in records:
        if r.get("status") == "confirmed" and r.get("issue_date"):
            months.append(r["issue_date"][:7].replace("/", ""))
    if months:
        return Counter(months).most_common(1)[0][0]
    return datetime.now().strftime("%Y%m")


def item_gross(it: dict) -> int:
    """明細の税込換算額。外税行は税区分の税率を足す（freeeの計算に合わせる）。"""
    amt = int(it.get("amount") or 0)
    if it.get("tax_calc") == "外税":
        tc = it.get("tax_class", "")
        if "8%" in tc:
            return round(amt * 1.08)
        if "10%" in tc:
            return round(amt * 1.1)
    return amt


def export_csv(office_name: str, ids=None) -> dict:
    """確認済み（または選択ID）をfreee CSV化し、対象PDFを _exported へ移動する。"""
    office = OFFICES.get(office_name)
    if not office:
        return {"error": "事業所が見つかりません"}
    records = office["records"]
    folder = office["path"]

    confirmed = [r for r in records if r.get("status") == "confirmed" and not r.get("missing")]
    if ids:
        idset = set(ids)
        confirmed = [r for r in confirmed if r["id"] in idset]
    unconfirmed = len([r for r in records
                       if r.get("status") not in ("confirmed", "exported")
                       and not r.get("missing")])

    if not confirmed:
        return {"error": "出力対象がありません（確認済みを選択してね）",
                "count": 0, "rows": 0, "total": 0, "moved": 0,
                "unconfirmed": unconfirmed, "gensen_list": [], "mismatch_list": []}

    ym = most_common_ym(records)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_dir = folder / "_exported" / ym
    archive_dir.mkdir(parents=True, exist_ok=True)
    out_path = archive_dir / f"freee_import_{ym}_{stamp}.csv"

    header = ["収支区分", "管理番号", "発生日", "決済期日", "取引先",
              "勘定科目", "税区分", "金額", "税計算区分", "備考",
              "決済日", "決済口座", "決済金額"]

    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\r\n")
    writer.writerow(header)

    total = 0
    rows_out = 0
    gensen_list = []
    mismatch_list = []
    for r in confirmed:
        amount = int(r.get("amount") or 0)
        withholding = int(r.get("withholding") or 0)
        memo = r.get("memo", "")
        partner = r.get("partner", "")
        issue = r.get("issue_date", "")
        due = r.get("due_date", "")
        items = [it for it in (r.get("line_items") or []) if int(it.get("amount") or 0) != 0]

        if withholding > 0:
            memo = (memo + " ※源泉あり").strip()
            gensen_list.append({
                "partner": partner, "amount": amount,
                "withholding": withholding, "net": amount - withholding,
            })

        if items:
            # 明細があれば必ず明細行から出力（1件でも明細側の勘定科目・税区分を使う。税区分混在対応）。
            # 税込換算（外税行は税率分を足す）で請求額と比較し、ズレたら帳尻合わせせず警告。
            items_sum = sum(int(it.get("amount") or 0) for it in items)
            gross_sum = sum(item_gross(it) for it in items)
            if gross_sum != amount:
                mismatch_list.append({"partner": partner, "amount": amount,
                                      "items_sum": items_sum, "gross_sum": gross_sum})
            for it in items:
                line_memo = it.get("name", "")
                if withholding > 0:
                    line_memo = (line_memo + " ※源泉あり").strip()
                writer.writerow([
                    "支出", "", issue, due, partner,
                    it.get("account") or r.get("account", ""),
                    it.get("tax_class") or r.get("tax_class", "課対仕入10%"),
                    int(it.get("amount") or 0), it.get("tax_calc") or "内税",
                    line_memo, "", "", ""
                ])
                rows_out += 1
            total += gross_sum
        else:
            writer.writerow([
                "支出", "", issue, due, partner,
                r.get("account", ""), r.get("tax_class", "課対仕入10%"),
                amount, "内税", memo, "", "", ""
            ])
            rows_out += 1
            total += amount

    raw = buf.getvalue().encode("cp932", errors="replace")
    out_path.write_bytes(raw)

    # アーカイブ移動: 対象PDFを _exported/YYYYMM/ へ移し、status=exported に
    moved = 0
    for r in confirmed:
        src = pathlib.Path(r["path"])
        try:
            if src.exists() and src.parent != archive_dir:
                dest = unique_path(archive_dir / src.name)
                src.rename(dest)
                r["path"] = str(dest)
                r["file"] = dest.name
                r["id"] = file_id(dest)  # idはパス由来なので更新
                moved += 1
        except OSError:
            pass  # 移動できなくても出力自体は成立させる
        r["status"] = "exported"
        r["exported_at"] = stamp
        r["export_csv"] = out_path.name
    save_data(folder, records)
    office["id_map"] = {r["id"]: r for r in records}

    return {
        "path": str(out_path),
        "csv_name": out_path.name,
        "count": len(confirmed),
        "rows": rows_out,
        "total": total,
        "moved": moved,
        "unconfirmed": unconfirmed,
        "gensen_list": gensen_list,
        "mismatch_list": mismatch_list,
    }


# グローバル状態（マルチテナント＝事業所別の箱）
ROOT: pathlib.Path = None        # 事業所ルートフォルダ
OFFICES: dict = {}               # name -> {"path", "records", "id_map"}
LOCK = threading.Lock()


def safe_office_name(name: str) -> str:
    """事業所名をフォルダ名として安全化（パス区切り・先頭アンダースコア等を除去）"""
    name = (name or "").strip()
    name = re.sub(r'[\\/:*?"<>|]', "", name).strip().strip(". 　")
    name = name.lstrip("_")
    return name[:60]


def list_office_names() -> list:
    if not ROOT or not ROOT.is_dir():
        return []
    return [p.name for p in sorted(ROOT.iterdir())
            if p.is_dir() and not p.name.startswith("_") and not p.name.startswith(".")]


def load_office(name: str) -> dict:
    """事業所をスキャン（必要ならOCR）してメモリに載せる"""
    path = ROOT / name
    existing = load_data(path)
    records = scan_folder(path, existing)
    save_data(path, records)
    office = {"path": path, "records": records,
              "id_map": {r["id"]: r for r in records}}
    OFFICES[name] = office
    return office


def get_office(name: str, scan: bool = True):
    name = safe_office_name(name)
    if not name or not (ROOT / name).is_dir():
        return None
    if name in OFFICES:
        return OFFICES[name]
    return load_office(name) if scan else None


def save_office(name: str):
    office = OFFICES.get(name)
    if office:
        save_data(office["path"], office["records"])
        office["id_map"] = {r["id"]: r for r in office["records"]}


def office_summary(name: str) -> dict:
    """軽量サマリー（OCRはせず、キャッシュ＋ディスク件数からカウント）"""
    path = ROOT / name
    cached = load_data(path)
    by = {}
    cached_paths = set()
    for r in cached:
        st = r.get("status", "check")
        by[st] = by.get(st, 0) + 1
        cached_paths.add(r.get("path"))
    new_files = 0
    if path.is_dir():
        for p in path.rglob("*"):
            if p.suffix.lower() not in SUPPORTED_EXTS or DATA_FILE_NAME in p.name:
                continue
            if _is_archived_path(path, p):
                continue
            if str(p) not in cached_paths:
                new_files += 1
    pending = (by.get("check", 0) + by.get("ok", 0)
               + by.get("gensen", 0) + new_files)
    return {
        "name": name,
        "pending": pending,                  # 未確認（要レビュー）
        "ready": by.get("confirmed", 0),     # 確認済み（出力待ち）
        "exported": by.get("exported", 0),   # 出力済み
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # アクセスログ抑制

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/":
            content = INDEX_HTML.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", len(content))
            # UI更新が古いキャッシュで見えなくなるのを防ぐ
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(content)

        elif path == "/api/offices":
            # 事業所一覧（軽量サマリー）
            offices = [office_summary(n) for n in list_office_names()]
            self.send_json({"root": str(ROOT), "offices": offices,
                             "ai_ocr_available": ai_ocr_available()})

        elif path == "/api/invoices":
            qs = parse_qs(parsed.query)
            office_name = safe_office_name(qs.get("office", [""])[0])
            office = get_office(office_name)
            if not office:
                self.send_json({"error": "事業所が見つかりません"}, 404)
                return
            self.send_json({
                "office": office_name,
                "folder": str(office["path"]),
                "records": office["records"],
                "ai_ocr_available": ai_ocr_available(),
            })

        elif path == "/api/highlight":
            qs = parse_qs(parsed.query)
            office_name = safe_office_name(qs.get("office", [""])[0])
            rid = qs.get("id", [None])[0]
            office = get_office(office_name, scan=False) or get_office(office_name)
            rec = office["id_map"].get(rid) if office else None
            if not rec:
                self.send_json({"ok": False, "error": "レコードが見つかりません"}, 404)
                return
            file_path = pathlib.Path(rec["path"])
            if not file_path.exists():
                self.send_json({"ok": False, "error": "元ファイルが見つかりません"}, 404)
                return
            result = build_highlight(file_path, rec, office_name)
            self.send_json(result, 200 if result.get("ok") else 400)

        elif path == "/pdf":
            qs = parse_qs(parsed.query)
            office_name = safe_office_name(qs.get("office", [""])[0])
            rid = qs.get("id", [None])[0]
            office = get_office(office_name, scan=False) or get_office(office_name)
            rec = office["id_map"].get(rid) if office else None
            if not rec:
                self.send_response(404)
                self.end_headers()
                return
            file_path = pathlib.Path(rec["path"])
            if not file_path.exists():
                self.send_response(404)
                self.end_headers()
                return
            suffix = file_path.suffix.lower()
            if suffix == ".pdf":
                ctype = "application/pdf"
            elif suffix in (".jpg", ".jpeg"):
                ctype = "image/jpeg"
            elif suffix == ".png":
                ctype = "image/png"
            else:
                ctype = "application/octet-stream"
            data = file_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Disposition", "inline")
            self.send_header("Content-Length", len(data))
            self.end_headers()
            self.wfile.write(data)

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get("Content-Length", 0))

        if path == "/api/upload":
            qs = parse_qs(parsed.query)
            office_name = safe_office_name(qs.get("office", [""])[0])
            office = get_office(office_name)
            if not office:
                self.send_json({"error": "事業所が見つかりません"}, 404)
                return
            name = qs.get("name", [""])[0]
            name = pathlib.Path(name).name  # パス成分を除去
            suffix = pathlib.Path(name).suffix.lower()
            if not name or suffix not in SUPPORTED_EXTS:
                self.send_json({"error": "対応形式は PDF / JPG / PNG のみです"}, 400)
                return
            if length <= 0 or length > 50 * 1024 * 1024:
                self.send_json({"error": "ファイルサイズが不正です（上限50MB）"}, 400)
                return
            body = self.rfile.read(length)
            with LOCK:
                dest = unique_path(office["path"] / name)
                dest.write_bytes(body)
                rec = {
                    "id": file_id(dest),
                    "file": dest.name,
                    "path": str(dest),
                    "partner": "",
                    "issue_date": "",
                    "due_date": "",
                    "amount": 0,
                    "account": "",
                    "tax_class": "課対仕入10%",
                    "withholding": 0,
                    "memo": "",
                    "status": "check",
                    "missing": False,
                    "line_items": [],
                    "source": "none",
                    "ai_confidence": None,
                    "ai_notes": "",
                    "user_edited": False,
                }
                if dest.suffix.lower() == ".pdf":
                    extracted = ocr_pdf(dest)
                    rec.update(extracted)
                    if extracted.get("partner") and extracted.get("amount"):
                        rec["source"] = "text"
                incomplete = not (rec.get("partner") and rec.get("amount"))
                if incomplete and ai_ocr_available():
                    apply_ai_ocr_to_record(rec, dest)
                office["records"].append(rec)
                save_office(office_name)
            self.send_json({"ok": True, "record": rec})
            return

        body = self.rfile.read(length)
        try:
            payload = json.loads(body)
        except Exception:
            self.send_json({"error": "invalid json"}, 400)
            return

        if path == "/api/offices":
            # 事業所を新規作成（フォルダを作る）
            name = safe_office_name(payload.get("name", ""))
            if not name:
                self.send_json({"error": "事業所名が不正です"}, 400)
                return
            new_path = ROOT / name
            if new_path.exists():
                self.send_json({"error": "同名の事業所が既にあります"}, 400)
                return
            with LOCK:
                new_path.mkdir(parents=True, exist_ok=True)
            self.send_json({"ok": True, "name": name})
            return

        if path == "/api/office_delete":
            # 事業所フォルダごと _trash/ へ退避（元に戻せる。他エンドポイントと違い
            # get_office()は呼ばない＝削除対象を余計にスキャン・OCRしない）
            name = safe_office_name(payload.get("office", ""))
            src = ROOT / name if name else None
            if not name or not src.is_dir():
                self.send_json({"error": "事業所が見つかりません"}, 404)
                return
            with LOCK:
                trash_root = ROOT / "_trash"
                trash_root.mkdir(exist_ok=True)
                dest = unique_path(trash_root / name)
                try:
                    shutil.move(str(src), str(dest))
                except Exception as e:
                    self.send_json({"error": f"フォルダの移動に失敗しました: {e}"}, 500)
                    return
                OFFICES.pop(name, None)  # メモリ上のキャッシュからも除去
            self.send_json({"ok": True})
            return

        # 以降は office を要求するエンドポイント
        office_name = safe_office_name(payload.get("office", ""))
        office = get_office(office_name)
        if not office:
            self.send_json({"error": "事業所が見つかりません"}, 404)
            return

        if path == "/api/save":
            rid = payload.get("id")
            rec = office["id_map"].get(rid)
            if not rec:
                self.send_json({"error": "not found"}, 404)
                return
            allowed = {"partner", "issue_date", "due_date", "amount",
                       "account", "tax_class", "withholding", "memo", "line_items"}
            with LOCK:
                marker_fields = {"partner", "issue_date", "due_date", "amount", "withholding"}
                for k in allowed:
                    if k in payload:
                        rec[k] = payload[k]
                        if k in marker_fields:
                            # 値が変わるとAI推定のマーカー位置が古くなるためキャッシュを破棄
                            rec.pop("ai_bboxes", None)
                rec["user_edited"] = True  # 手動編集済み。以後の自動AI-OCRで上書きしない
                save_office(office_name)
            self.send_json({"ok": True})

        elif path == "/api/confirm":
            rid = payload.get("id")
            rec = office["id_map"].get(rid)
            if not rec:
                self.send_json({"error": "not found"}, 404)
                return
            with LOCK:
                if rec.get("status") == "confirmed":
                    # 元ステータスに戻す
                    prev = rec.get("_prev_status", "ok")
                    rec["status"] = prev
                    rec.pop("_prev_status", None)
                else:
                    rec["_prev_status"] = rec.get("status", "ok")
                    rec["status"] = "confirmed"
                save_office(office_name)
            self.send_json({"ok": True, "status": rec["status"]})

        elif path == "/api/export":
            ids = payload.get("ids") or None
            with LOCK:
                result = export_csv(office_name, ids)
            self.send_json(result)

        elif path == "/api/ocr":
            # 1件を明示的にAI-OCRで再読取（確認済み・手動編集済みでも明示操作なので実行OK）
            rid = payload.get("id")
            rec = office["id_map"].get(rid)
            if not rec:
                self.send_json({"error": "not found"}, 404)
                return
            if not ai_ocr_available():
                self.send_json({"error": "AI-OCRが利用できません（.envにANTHROPIC_API_KEYまたはCLAUDE_API_KEYが必要です。またはClaude Codeのインストールが必要です）"}, 400)
                return
            file_path = pathlib.Path(rec["path"])
            if not file_path.exists():
                self.send_json({"error": "元ファイルが見つかりません"}, 404)
                return
            with LOCK:
                ok, err = apply_ai_ocr_to_record(rec, file_path)
                save_office(office_name)
            if ok:
                self.send_json({"ok": True, "record": rec})
            else:
                self.send_json({"ok": False, "error": err or "AI-OCRに失敗しました", "record": rec})

        elif path == "/api/ocr_pending":
            # 「要確認」全件（確認済み・手動編集済み・欠落は除外）をまとめてAI-OCR
            if not ai_ocr_available():
                self.send_json({"error": "AI-OCRが利用できません（.envにANTHROPIC_API_KEYまたはCLAUDE_API_KEYが必要です。またはClaude Codeのインストールが必要です）"}, 400)
                return
            targets = [r for r in office["records"]
                       if r.get("status") in ("check", "ok")
                       and not r.get("missing")
                       and not r.get("user_edited")]
            processed, improved, failed = 0, 0, 0
            with LOCK:
                for rec in targets:
                    file_path = pathlib.Path(rec["path"])
                    if not file_path.exists():
                        continue
                    prev_status = rec.get("status")
                    ok, _err = apply_ai_ocr_to_record(rec, file_path)
                    processed += 1
                    if ok:
                        if prev_status == "check" and rec.get("status") in ("ok", "gensen"):
                            improved += 1
                    else:
                        failed += 1
                    time.sleep(0.2)  # レート制限対策の小休止
                save_office(office_name)
            self.send_json({"ok": True, "processed": processed,
                             "improved": improved, "failed": failed})

        elif path == "/api/delete":
            # レコードを削除する。実ファイルは消さず _trash/ へ退避（元に戻せる）
            rid = payload.get("id")
            rec = office["id_map"].get(rid)
            if not rec:
                self.send_json({"ok": False, "error": "レコードが見つかりません"}, 404)
                return
            with LOCK:
                file_path = pathlib.Path(rec["path"])
                if file_path.exists():
                    trash_dir = office["path"] / "_trash"
                    trash_dir.mkdir(exist_ok=True)
                    dest = unique_path(trash_dir / file_path.name)
                    try:
                        shutil.move(str(file_path), str(dest))
                    except Exception as e:
                        self.send_json({"ok": False, "error": f"ファイルの移動に失敗しました: {e}"}, 500)
                        return
                # ファイルが既に無い場合（_exported済み等）もエラーにせずレコード削除は続行
                office["records"] = [r for r in office["records"] if r["id"] != rid]
                save_office(office_name)
            self.send_json({"ok": True})

        else:
            self.send_response(404)
            self.end_headers()


def _is_own_server_running(port: int, timeout: float = 1.5) -> bool:
    """指定ポートに何か応答するサーバーがいるか確認する（二重起動時のブラウザ再オープン判定用）。"""
    try:
        with urlopen(f"http://127.0.0.1:{port}/", timeout=timeout):
            return True
    except OSError:
        return False


def main():
    global ROOT

    parser = argparse.ArgumentParser(description="請求書OCRレビューサーバー（事業所別）")
    parser.add_argument("--root", help="事業所ルートフォルダ（省略時はツール直下 _offices）")
    parser.add_argument("--folder", help="(互換) 単一フォルダ。その親をルート、自身を1事業所として扱う")
    parser.add_argument("--port", type=int, default=3470)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    if args.root:
        ROOT = pathlib.Path(args.root).resolve()
    elif args.folder:
        # 後方互換: 単一フォルダ指定 → その親をルートに（フォルダ自体が1事業所）
        ROOT = pathlib.Path(args.folder).resolve().parent
    else:
        ROOT = (SCRIPT_DIR / "_offices").resolve()
    ROOT.mkdir(parents=True, exist_ok=True)

    names = list_office_names()
    print(f"事業所ルート: {ROOT}")
    print(f"  事業所 {len(names)} 件: {('、'.join(names)) if names else '（まだありません。ホーム画面から追加してね）'}")

    if not HAS_PDFPLUMBER:
        print("警告: pdfplumber が見つかりません。PDF抽出はスキップします（pip install pdfplumber で有効化）")

    if not HAS_FITZ:
        print("警告: PyMuPDF(fitz) が見つかりません。OCR根拠マーカー表示は無効です（pip install pymupdf で有効化）")

    ocr_mode = vision_ocr.ocr_mode() if HAS_VISION_OCR else "none"
    if ocr_mode == "api":
        print("AI-OCR: API（スキャンPDF・画像は自動でClaude Visionにフォールバックします）")
    elif ocr_mode == "cli":
        print("AI-OCR: Claude CLI（APIキー未設定のためClaude Code CLI経由でフォールバックします。位置特定機能のみ非対応）")
    else:
        print("AI-OCR: 無効（.env に ANTHROPIC_API_KEY / CLAUDE_API_KEY が無く、Claude CLIも見つかりません。テキスト抽出のみで動作します）")

    print(f"サーバー起動: http://127.0.0.1:{args.port}")

    # Windowsではallow_reuse_address=Trueだと同一ポートに二重起動できてしまい、
    # 古いプロセスが新旧混在でリクエストを受ける事故になるため無効化（二重起動は即エラーで落とす）
    ThreadingHTTPServer.allow_reuse_address = False
    try:
        server = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    except OSError:
        # ポートが使用中 → 既にこのツールが起動済みの可能性が高いので疎通確認する。
        # 応答があればエラー扱いにせず、そのままブラウザを開いて正常終了（start.batの二重起動対策）。
        if _is_own_server_running(args.port):
            print(f"既にサーバーが起動しています（http://127.0.0.1:{args.port}）。ブラウザを開きます。")
            if not args.no_browser:
                webbrowser.open(f"http://127.0.0.1:{args.port}")
            return
        print(f"ポート{args.port}は使用中ですが応答がありません（別のアプリが使用中の可能性があります）")
        sys.exit(1)
    server.daemon_threads = True

    if not args.no_browser:
        webbrowser.open(f"http://127.0.0.1:{args.port}")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nサーバーを停止しました")


if __name__ == "__main__":
    main()
