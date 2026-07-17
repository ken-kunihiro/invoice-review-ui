#!/usr/bin/env python3
"""
汎用 請求書/領収書 AI-OCR（Claude Vision）

PDF・画像（JPG/PNG/WEBP）から経理情報を抽出する。
invoice-review-ui/server.py の補助として、pdfplumberのテキスト抽出が
不完全な場合（スキャンPDF・画像ファイル）にフォールバックで使う。

PII保護：住所・電話・メールアドレス・口座番号はJSONに含めない
（vendor名とインボイス登録番号のみ残す）。

単体実行（動作確認用）:
  python vision_ocr.py <ファイルパス>
"""
import base64
import json
import mimetypes
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile

MODEL = "claude-sonnet-5"
MAX_TOKENS = 8192  # thinkingを返すモデルは思考分もmax_tokensを消費するため大きめに確保
MAX_PDF_BYTES = 20 * 1024 * 1024  # 20MB

SUPPORTED_EXTS = {".pdf", ".jpg", ".jpeg", ".png", ".webp"}

# 汎用勘定科目リスト（畠山謙人税理士事務所の分類ルールに準拠）
GENERIC_ACCOUNT_ITEMS = [
    "外注費", "業務委託費", "支払報酬料", "旅費交通費", "通信費",
    "会議費", "交際費", "消耗品費", "地代家賃", "支払手数料",
    "新聞図書費", "広告宣伝費", "福利厚生費", "車両費", "諸会費", "雑費",
]

TAX_CLASS_OPTIONS = ["課対仕入10%", "課対仕入8%（軽）", "対象外", "非課税", "不課税"]

SCRIPT_DIR = pathlib.Path(__file__).resolve().parent
RULES_PATH = SCRIPT_DIR / "rules.txt"
RULES_LOCAL_PATH = SCRIPT_DIR / "rules.local.txt"

# rules.txt が見つからない場合のフォールバック（旧SYSTEM_PROMPT直書き内容と同等。
# 個人ルールは含めない＝個人ルールは rules.local.txt 側で別途読み込む）
FALLBACK_RULES_TEXT = """- 税理士・公認会計士・社労士・弁護士・司法書士等への顧問料・監査報酬 → 支払報酬料
- 電車・タクシー・高速代・宿泊・駐車場・出張費 → 旅費交通費
- SaaS・サブスク・サーバー・ドメイン・クラウド利用料 → 通信費
- 飲食1万円以下・カフェでの打合せ → 会議費
- 飲食1万円超・接待・贈答 → 交際費
- 文具・事務用品・少額備品・名刺 → 消耗品費
- 家賃・賃料・レンタルオフィス・バーチャルオフィス → 地代家賃
- 振込手数料・決済手数料 → 支払手数料
- 書籍・新聞・購読料 → 新聞図書費
- 広告・宣伝・掲載料 → 広告宣伝費
- 制作・開発・デザイン・執筆・運用代行等の外部委託 → 外注費 または 業務委託費
- クレジットカードの年会費 → 会議費（支払手数料や諸会費にしない）
- Amazonプライム会費など、サービスの年会費・月会費 → 諸会費
- ウォーターサーバー → 福利厚生費
- 出張でないホテル・宿泊施設の利用 → 福利厚生費（出張なら旅費交通費）
- ガソリンスタンド（ENEOS・出光・コスモ石油等） → 車両費
- PR TIMES・Indeed 等の広告掲載・求人掲載 → 広告宣伝費
- Stripe 等の決済手数料 → 支払手数料
- Amazon・ASKUL・モノタロウでの物品購入 → 消耗品費"""


def _read_rules_stripped(path: pathlib.Path):
    """rules.txt / rules.local.txt を読み、# コメント行を除去して返す。
    無い/読めない場合は None。キャッシュしない（毎回ディスクから読む＝
    サーバー起動中に編集しても次回OCRから即反映される）。
    """
    try:
        if not path.exists():
            return None
        raw = path.read_text(encoding="utf-8")
    except OSError:
        return None
    lines = [ln for ln in raw.splitlines() if not ln.strip().startswith("#")]
    text = "\n".join(lines).strip()
    return text or None


def _load_account_item_rules() -> str:
    """勘定科目ルール本文を組み立てる（rules.txt＋rules.local.txt、フォールバック込み）。"""
    base = _read_rules_stripped(RULES_PATH)
    if base is None:
        base = FALLBACK_RULES_TEXT
    local = _read_rules_stripped(RULES_LOCAL_PATH)
    if local:
        base = base + "\n" + local
    return base


def build_system_prompt() -> str:
    """SYSTEM_PROMPTを毎回組み立てる（rules.txt/rules.local.txtを都度読み込むため）。

    安全弁（「判断できなければ空文字にする」）は rules.txt の内容に関わらず
    このPython側に必ず残す。rules.txtが壊れていてもOCRの安全側デフォルトが
    崩れないようにするため。
    """
    account_item_rules = _load_account_item_rules()
    return f"""あなたは経理書類のOCR専門家です。請求書・領収書の画像/PDFから経理情報を抽出します。

# 出力フォーマット (JSON)
{{
  "doc_type": "invoice" | "receipt" | "expense_list" | "unknown",
  "issue_date": "YYYY-MM-DD" or null,
  "service_month_end": "YYYY-MM-DD" or null,
  "due_date": "YYYY-MM-DD" or null,
  "vendor": "取引先名（発行元）",
  "registration_number": "T13桁" or null,
  "amount_total": 税込合計（整数）or null,
  "withholding_tax": 源泉所得税額（整数、記載がなければnull）,
  "line_items": [  // 必ず1行以上返す（下記「明細の扱い」参照）
    {{ "description": "品目", "amount": 税込金額（整数）, "account_item": "勘定科目名", "tax_class": "課対仕入10%|課対仕入8%（軽）|対象外|非課税|不課税" }}
  ],
  "confidence": "high" | "mid" | "low",
  "notes": "備考（曖昧な点・要確認事項など。なければ空文字）"
}}

# 取引先（vendor）の判定 ― 最重要
- 取引先＝請求書・領収書の発行元（差出人）。宛先（「〇〇御中」「〇〇様」が付く側）は取引先ではない
- 宛先側に見える名前（受領者・自社らしき名前）は絶対にvendorに入れない
- 個人事業主・フリーランスの氏名でもよい（会社名でなくてよい）

# doc_type の判定
- 「請求書」「Invoice」の体裁 → "invoice"
- 店舗のレシート・領収書、合計のみ → "receipt"
- 複数の店舗・SaaS利用がまとまった経費精算リスト → "expense_list"
- 判定できなければ "unknown"

# 勘定科目（account_item）の選択ルール
{', '.join(GENERIC_ACCOUNT_ITEMS)} から選ぶ。
{account_item_rules}

- 上記のどれにも当てはまらず判断できなければ account_item は空文字にする（無理に埋めない）

# 税区分（tax_class）
- 国内の課税取引（標準10%） → "課対仕入10%"
- 軽減税率8%（飲食料品等） → "課対仕入8%（軽）"
- 海外SaaS・海外事業者への支払い、外貨建て取引 → "対象外"
- 判断できなければ "課対仕入10%" をデフォルトにする

# 計上日（service_month_end）
役務提供月の末日。
- 請求書発行日が月末 → 同月末
- 請求書発行日が翌月初旬（〜10日） → 前月末
- 請求書発行日が翌月中旬以降 → 同月末
- 領収書（receipt）は issue_date と同じでよい

# 源泉所得税（withholding_tax）
書類に源泉徴収額の記載があれば整数で抽出する。記載がなければ null（0ではなくnull）。

# confidence
- high: 取引先・金額・日付すべて明確に読み取れた
- mid: 一部推測が入った
- low: 手書き・不鮮明・情報が乏しく確信が持てない

# 明細（line_items）の扱い ― 必ず1行以上返すこと
- 単一品目・単一税率の書類でも、合計を1行の明細として必ず返す
- **税率が混在する書類（8%軽減と10%が両方ある等）は、必ず税率ごとに行を分ける**
- レシート末尾の税率内訳欄（「8%対象」「10%対象」「外税8%対象額」等）を最優先で確認し、両方の税率に対象額があれば必ず2行に分ける
- 外税表記（対象額と税額が別記載）の場合、各行の amount は「対象額＋その税率の税額」の税込額にする（例: 8%対象10,000円・税800円 → amount 10800）
- 品目が多い場合は「勘定科目×税区分」の組み合わせごとに1行へまとめてよい（同じ勘定科目でも税区分が違えば必ず別行）
- 各行の amount は税込金額。全行の合計が amount_total と一致するように読み取る
- 一致させられない（内訳が読み取れない）場合は、数字を作って帳尻合わせせず、notes にその旨を書いて confidence を下げる

# PII保護（重要・必ず守ること）
- 出力JSONに **住所・電話番号・メールアドレス・口座番号** を一切含めないこと
- vendor名（会社名・個人名）と登録番号（インボイスT番号）は含めてOK

JSONのみを返してください（説明文・コードフェンス・マークダウン不要）。"""


def _load_env():
    """.env候補を順に探してAPIキーを読み込む。値は一切ログ出力しない。

    ANTHROPIC_API_KEY が無い環境（メインPCの .env はダッシュボード用
    CLAUDE_API_KEY のみ）では CLAUDE_API_KEY をフォールバックとして使う。
    """
    if os.environ.get("ANTHROPIC_API_KEY"):
        return
    try:
        from dotenv import load_dotenv
    except ImportError:
        load_dotenv = None
    if load_dotenv is not None:
        # このフォルダ→親フォルダの順に .env を探す（リポジトリ直下の .env も拾える）
        candidates = [SCRIPT_DIR / ".env"]
        candidates += [parent / ".env" for parent in SCRIPT_DIR.parents]
        for env_path in candidates:
            try:
                if env_path.exists():
                    load_dotenv(env_path)
            except OSError:
                continue
    if not os.environ.get("ANTHROPIC_API_KEY") and os.environ.get("CLAUDE_API_KEY"):
        os.environ["ANTHROPIC_API_KEY"] = os.environ["CLAUDE_API_KEY"]


_load_env()

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


NO_OCR_ERROR = ("AI-OCRが利用できません"
                 "（ANTHROPIC_API_KEY/CLAUDE_API_KEY未設定 または anthropicライブラリ未インストール。"
                 "またはClaude Codeのインストールが必要です）")


def _api_key_available() -> bool:
    return HAS_ANTHROPIC and bool(os.environ.get("ANTHROPIC_API_KEY"))


def _cli_path():
    """Claude Code CLI（`claude`コマンド）の実行パスを返す。無ければNone。

    Windowsでは claude.EXE / claude.CMD 等に解決される想定。呼び出す都度
    探す（PATH変更やインストール直後の反映を即拾えるようにするため）。
    """
    return shutil.which("claude")


def ocr_mode() -> str:
    """現在使えるAI-OCRの実行方式を返す。

    "api": ANTHROPIC_API_KEY（またはCLAUDE_API_KEY代替）でSDK直叩き
    "cli": APIキーは無いがClaude Code CLIがインストール済み（受講生向けフォールバック）
    "none": どちらも無い
    """
    if _api_key_available():
        return "api"
    if _cli_path():
        return "cli"
    return "none"


def is_available() -> bool:
    """AI-OCRが何らかの方式（API直叩き or Claude CLI）で使える状態か"""
    return ocr_mode() != "none"


def _media_type(path: pathlib.Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return "application/pdf"
    if suffix in (".jpg", ".jpeg"):
        return "image/jpeg"
    if suffix == ".png":
        return "image/png"
    if suffix == ".webp":
        return "image/webp"
    guessed, _ = mimetypes.guess_type(str(path))
    return guessed or "application/octet-stream"


def _int_or_none(v):
    try:
        return int(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _normalize(data: dict, path: pathlib.Path) -> dict:
    """欠けているキーを補い、型を安全にそろえる（想定外の値が来ても壊れないように）"""
    items = []
    for it in (data.get("line_items") or []):
        if not isinstance(it, dict):
            continue
        tax_class = it.get("tax_class") if it.get("tax_class") in TAX_CLASS_OPTIONS else "課対仕入10%"
        items.append({
            "description": str(it.get("description") or "")[:60],
            "amount": _int_or_none(it.get("amount")) or 0,
            "account_item": str(it.get("account_item") or "")[:20],
            "tax_class": tax_class,
        })

    return {
        "doc_type": data.get("doc_type") if data.get("doc_type") in
                    ("invoice", "receipt", "expense_list", "unknown") else "unknown",
        "issue_date": data.get("issue_date") or None,
        "service_month_end": data.get("service_month_end") or None,
        "due_date": data.get("due_date") or None,
        "vendor": (data.get("vendor") or "").strip() or None,
        "registration_number": data.get("registration_number") or None,
        "amount_total": _int_or_none(data.get("amount_total")),
        "withholding_tax": _int_or_none(data.get("withholding_tax")),
        "line_items": items,
        "confidence": data.get("confidence") if data.get("confidence") in ("high", "mid", "low") else "mid",
        "notes": str(data.get("notes") or "")[:300],
        "file": path.name,
    }


def _ocr_file_cli(path: pathlib.Path) -> dict:
    """Claude Code CLI（ヘッドレスモード）経由でOCRする。

    ANTHROPIC_API_KEYが無い環境（Claude Codeは持っているが個別のAPIキーは
    持っていない受講生向け）のフォールバック。CLIのReadツールに画像/PDFを
    読ませる必要があるため、いったんtempfileに書き出してパスを渡す。
    """
    cli = _cli_path()
    if not cli:
        return {"error": "Claude CLIが見つかりません"}

    tmp_suffix = path.suffix.lower()
    if tmp_suffix not in SUPPORTED_EXTS:
        tmp_suffix = ".png"

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(prefix="vision_ocr_", suffix=tmp_suffix, delete=False) as tf:
            tf.write(path.read_bytes())
            tmp_path = pathlib.Path(tf.name)

        prompt = (
            build_system_prompt()
            + "\n\n# 実行手順\n"
            + f"まず Read ツールで次のファイルを読み込んでください: {tmp_path}\n"
            + "読み込んだ画像/PDFから経理情報を抽出し、出力フォーマットの通りJSONのみを出力してください"
              "（説明文・前置き・コードフェンス・マークダウン一切不要。JSON以外の文字を含めないこと）。"
        )

        try:
            result = subprocess.run(
                [cli, "-p", prompt, "--model", MODEL, "--output-format", "text", "--allowedTools", "Read"],
                capture_output=True, timeout=180, encoding="utf-8", errors="replace",
            )
        except subprocess.TimeoutExpired:
            return {"error": "Claude CLI呼び出しがタイムアウトしました（180秒）"}
        except Exception as e:
            return {"error": f"Claude CLI呼び出しエラー: {e}"}

        if result.returncode != 0:
            stderr = (result.stderr or "").strip()[:200]
            return {"error": f"Claude CLI呼び出しに失敗しました（終了コード{result.returncode}）: {stderr}"}

        raw = (result.stdout or "").strip()
    finally:
        if tmp_path is not None:
            try:
                tmp_path.unlink(missing_ok=True)
            except OSError:
                pass

    if raw.startswith("```"):
        lines = raw.splitlines()
        raw = "\n".join(lines[1:])
        if raw.endswith("```"):
            raw = raw[:-3]
        raw = raw.strip()

    # CLIの出力は前置き文が付くことがあるため、最初の{〜最後の}を抽出する
    start, end = raw.find("{"), raw.rfind("}")
    if start == -1 or end == -1 or end < start:
        return {"error": f"Claude CLI結果にJSONが見つかりませんでした: {raw[:120]}"}
    json_str = raw[start:end + 1]

    try:
        data = json.loads(json_str)
    except json.JSONDecodeError:
        return {"error": f"Claude CLI結果のJSON解析に失敗しました: {json_str[:120]}"}

    if not isinstance(data, dict):
        return {"error": "Claude CLI結果の形式が不正です"}

    return _normalize(data, path)


def ocr_file(path) -> dict:
    """1ファイルをAI-OCRする。失敗時も例外を投げず {"error": "..."}（を含むdict）を返す。"""
    path = pathlib.Path(path)
    if not path.exists():
        return {"error": f"ファイルが見つかりません: {path.name}"}
    if path.suffix.lower() not in SUPPORTED_EXTS:
        return {"error": f"未対応の形式です: {path.suffix}"}
    if path.suffix.lower() == ".pdf" and path.stat().st_size > MAX_PDF_BYTES:
        return {"error": "PDFのサイズが大きすぎます（上限20MB）"}

    mode = ocr_mode()
    if mode == "none":
        return {"error": NO_OCR_ERROR}
    if mode == "cli":
        return _ocr_file_cli(path)

    try:
        media_type = _media_type(path)
        b64 = base64.standard_b64encode(path.read_bytes()).decode("utf-8")
        if media_type == "application/pdf":
            content_block = {"type": "document",
                              "source": {"type": "base64", "media_type": media_type, "data": b64}}
        else:
            content_block = {"type": "image",
                              "source": {"type": "base64", "media_type": media_type, "data": b64}}

        client = anthropic.Anthropic()
        msg = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=build_system_prompt(),
            messages=[{
                "role": "user",
                "content": [
                    content_block,
                    {"type": "text", "text": "このファイルから経理情報を抽出してください。"},
                ],
            }],
        )
        # thinkingブロックを返すモデルでも壊れないよう、textブロックだけを拾う
        raw = "".join(b.text for b in msg.content
                      if getattr(b, "type", "") == "text").strip()
    except Exception as e:
        return {"error": f"AI-OCR呼び出しエラー: {e}"}

    if raw.startswith("```"):
        lines = raw.splitlines()
        raw = "\n".join(lines[1:])
        if raw.endswith("```"):
            raw = raw[:-3]
        raw = raw.strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return {"error": f"AI-OCR結果のJSON解析に失敗しました: {raw[:120]}"}

    if not isinstance(data, dict):
        return {"error": "AI-OCR結果の形式が不正です"}

    return _normalize(data, path)


LOCATE_SYSTEM_PROMPT = """あなたは画像内の文字列の位置を特定するアシスタントです。
画像全体を幅1000×高さ1000の正規化座標（左上原点、右方向がx増加、下方向がy増加）とみなし、
指定された各項目の値が書かれている位置を矩形 [x0, y0, x1, y1]（整数、0〜1000）で返してください。

# 注意
- 値の再解釈・再抽出はしないでください。与えられた値がそのまま（表記ゆれ・全角半角・カンマの有無を含めて）書かれている箇所を探すことだけに専念してください
- 画像内に見つからない項目は null にしてください
- 矩形はその文字列を囲む最小限の範囲にしてください（行全体や項目欄全体を囲まない）

出力は次の形式のJSONのみ（説明文・コードフェンス・マークダウン不要）:
{"項目名": [x0, y0, x1, y1] または null, ...}
"""


def locate_fields(path, fields: dict) -> dict:
    """指定フィールドの値が画像/PDF内のどこにあるかをAIで推定する（マーカー表示用）。

    server.py の /api/highlight から、テキスト層検索で見つからなかった場合の
    フォールバックとして呼ばれる。値の再抽出はせず、位置探索のみ行う。

    fields: {"partner": "六花亭帯広空港店", "issue_date": "2026/07/31",
             "due_date": "...", "amount": 1650, "withholding": 500} のように
            既読取済みの値を渡す（空・0の項目は自動で除外）
    戻り値: {"boxes": {field: [x0,y0,x1,y1]（0-1000正規化）or None, ...}, "error": None}
            失敗時は {"boxes": {}, "error": "エラーメッセージ"}
    """
    path = pathlib.Path(path)
    if not path.exists():
        return {"boxes": {}, "error": f"ファイルが見つかりません: {path.name}"}
    if path.suffix.lower() not in SUPPORTED_EXTS:
        return {"boxes": {}, "error": f"未対応の形式です: {path.suffix}"}

    mode = ocr_mode()
    if mode == "none":
        return {"boxes": {}, "error": NO_OCR_ERROR}

    targets = {k: v for k, v in (fields or {}).items() if v not in (None, "", 0)}
    if not targets:
        return {"boxes": {}, "error": None}

    if mode == "cli":
        # 位置特定（locate_fields）はCLI経由だと1回180秒近くかかりうり、
        # ハイライト表示のためだけに使うには重すぎる。CLIフォールバック環境では
        # 位置特定を諦め「位置特定できず」として全項目Noneを返す（呼び出し元の
        # server.py側でハイライト無し表示にフォールバックする）。
        return {"boxes": {k: None for k in targets}, "error": "位置特定できず（Claude CLI環境では非対応）"}

    try:
        media_type = _media_type(path)
        b64 = base64.standard_b64encode(path.read_bytes()).decode("utf-8")
        if media_type == "application/pdf":
            content_block = {"type": "document",
                              "source": {"type": "base64", "media_type": media_type, "data": b64}}
        else:
            content_block = {"type": "image",
                              "source": {"type": "base64", "media_type": media_type, "data": b64}}

        prompt_lines = ["次の項目それぞれについて、画像内でその値が書かれている位置を教えてください。"]
        for k, v in targets.items():
            prompt_lines.append(f"- {k}: {v}")
        prompt_text = "\n".join(prompt_lines)

        client = anthropic.Anthropic()
        msg = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,  # thinkingを返すモデルでも途切れないよう本文と同じ上限
            system=LOCATE_SYSTEM_PROMPT,
            messages=[{
                "role": "user",
                "content": [
                    content_block,
                    {"type": "text", "text": prompt_text},
                ],
            }],
        )
        raw = "".join(b.text for b in msg.content
                      if getattr(b, "type", "") == "text").strip()
    except Exception as e:
        return {"boxes": {}, "error": f"AI-OCR呼び出しエラー: {e}"}

    if raw.startswith("```"):
        lines = raw.splitlines()
        raw = "\n".join(lines[1:])
        if raw.endswith("```"):
            raw = raw[:-3]
        raw = raw.strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return {"boxes": {}, "error": f"位置推定結果のJSON解析に失敗しました: {raw[:120]}"}

    if not isinstance(data, dict):
        return {"boxes": {}, "error": "位置推定結果の形式が不正です"}

    boxes = {}
    for k in targets:
        v = data.get(k)
        if (isinstance(v, list) and len(v) == 4
                and all(isinstance(n, (int, float)) for n in v)):
            x0, y0, x1, y1 = (float(n) for n in v)
            x0, x1 = sorted((x0, x1))
            y0, y1 = sorted((y0, y1))
            boxes[k] = [max(0.0, min(1000.0, x0)), max(0.0, min(1000.0, y0)),
                        max(0.0, min(1000.0, x1)), max(0.0, min(1000.0, y1))]
        else:
            boxes[k] = None

    return {"boxes": boxes, "error": None}


def main():
    """動作確認用: python vision_ocr.py <file>"""
    if len(sys.argv) < 2:
        print("Usage: python vision_ocr.py <path>")
        return
    result = ocr_file(sys.argv[1])
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
