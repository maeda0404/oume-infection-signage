#!/usr/bin/env python3
"""
西多摩地域 感染症流行状況サイネージ用
data/status.json 生成スクリプト

取得対象
・西多摩保健所管内 インフルエンザ定点当たり報告数
・西多摩保健所管内 新型コロナ定点当たり報告数
・東京都全体 インフルエンザ定点当たり報告数
・東京都全体 新型コロナ定点当たり報告数

データ取得元
・東京都西多摩保健所 感染症週報
・東京都感染症情報センター インフルエンザ患者報告数Excel
・東京都感染症情報センター 東京都感染症週報PDF

取得失敗時は、前回のstatus.jsonに有効な値があれば維持する。
取得できなかった値を0.00として出力しない。
"""

import gzip
import io
import json
import re
import time
import urllib.request
from datetime import date, datetime, timedelta, timezone
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urljoin, urlparse

from openpyxl import load_workbook
from pypdf import PdfReader


# =============================================================================
# 基本設定
# =============================================================================

JST = timezone(timedelta(hours=9))

ROOT_DIR = Path(__file__).resolve().parents[1]
STATUS_FILE = ROOT_DIR / "data" / "status.json"

NISHITAMA_WEEKLY_PAGE_URL = (
    "https://www.hokeniryo.metro.tokyo.lg.jp/"
    "shisetsu/jigyosyo/hokenjyo/nisitama/"
    "kansen/byouki/nishitama"
)

TOKYO_FLU_TOP_URL = (
    "https://idsc.tmiph.metro.tokyo.lg.jp/"
    "diseases/flu/flu/"
)

TOKYO_WEEKLY_PAGE_URL = (
    "https://idsc.tmiph.metro.tokyo.lg.jp/"
    "weekly/"
)

TOKYO_WEEKLY_PDF_BASE_URL = (
    "https://idsc.tmiph.metro.tokyo.lg.jp/"
    "assets/weekly"
)

USER_AGENT = (
    "Mozilla/5.0 "
    "Nishitama-Infection-Signage/3.0"
)


# =============================================================================
# HTMLリンク解析
# =============================================================================

class LinkParser(HTMLParser):
    """HTML内のリンク先と表示文字列を抽出する。"""

    def __init__(self):
        super().__init__()

        self.links = []
        self.current_href = None
        self.current_text = []

    def handle_starttag(self, tag, attrs):
        if tag.lower() != "a":
            return

        attributes = dict(attrs)

        self.current_href = attributes.get("href")
        self.current_text = []

    def handle_data(self, data):
        if self.current_href is not None:
            self.current_text.append(data)

    def handle_endtag(self, tag):
        if (
            tag.lower() == "a"
            and self.current_href is not None
        ):
            self.links.append(
                {
                    "href": self.current_href,
                    "text": " ".join(
                        self.current_text
                    ),
                }
            )

            self.current_href = None
            self.current_text = []


# =============================================================================
# 通信処理
# =============================================================================

def fetch_bytes(url, retries=3):
    """指定URLからバイトデータを取得する。"""

    last_error = None

    for attempt in range(1, retries + 1):
        try:
            separator = "&" if "?" in url else "?"

            cache_buster = int(
                datetime.now(
                    timezone.utc
                ).timestamp()
            )

            request = urllib.request.Request(
                f"{url}{separator}_={cache_buster}",
                headers={
                    "User-Agent": USER_AGENT,
                    "Accept": (
                        "application/pdf,"
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet,"
                        "application/vnd.ms-excel,"
                        "text/html,"
                        "*/*"
                    ),
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                },
            )

            with urllib.request.urlopen(
                request,
                timeout=45,
            ) as response:
                raw = response.read()

            if raw[:2] == b"\x1f\x8b":
                return gzip.decompress(raw)

            return raw

        except Exception as error:
            last_error = error

            print(
                f"通信失敗 {attempt}/{retries}: "
                f"{url}: {error}"
            )

            if attempt < retries:
                time.sleep(5 * attempt)

    raise last_error


def decode_html(raw):
    """HTMLバイト列を文字列へ変換する。"""

    for encoding in (
        "utf-8",
        "cp932",
        "shift_jis",
    ):
        try:
            return raw.decode(encoding)

        except UnicodeDecodeError:
            continue

    return raw.decode(
        "utf-8",
        errors="replace",
    )


def fetch_text(url):
    """指定URLからHTMLテキストを取得する。"""

    return decode_html(
        fetch_bytes(url)
    )


# =============================================================================
# 共通処理
# =============================================================================

def normalize_text(value):
    """文字列の全角記号と空白を正規化する。"""

    if value is None:
        return ""

    text = str(value)

    replacements = {
        "\u3000": " ",
        "（": "(",
        "）": ")",
        "－": "-",
        "―": "-",
        "〜": "～",
        "~": "～",
        "／": "/",
    }

    for before, after in replacements.items():
        text = text.replace(
            before,
            after,
        )

    return re.sub(
        r"[ \t\r\f\v]+",
        " ",
        text,
    ).strip()


def compact_text(value):
    """解析用にすべての空白を除去する。"""

    return re.sub(
        r"\s+",
        "",
        normalize_text(value),
    )


def extract_links(html, base_url):
    """HTMLから絶対URLとリンク文字列を取得する。"""

    parser = LinkParser()
    parser.feed(html)

    results = []

    for item in parser.links:
        href = item.get("href")

        if not href:
            continue

        results.append(
            {
                "url": urljoin(
                    base_url,
                    unescape(href),
                ),
                "label": normalize_text(
                    unescape(
                        item.get("text", "")
                    )
                ),
            }
        )

    return results


def number(value):
    """数値、セル値、文字列をfloatへ変換する。"""

    if (
        value is None
        or isinstance(value, bool)
    ):
        return None

    if isinstance(
        value,
        (int, float),
    ):
        return float(value)

    text = compact_text(value)

    text = (
        text
        .replace(",", "")
        .replace("人/定点", "")
        .replace("人", "")
    )

    if not re.fullmatch(
        r"-?\d+(?:\.\d+)?",
        text,
    ):
        return None

    return float(text)


def iso_week_period(year, week):
    """ISO週番号から月曜日～日曜日の期間を生成する。"""

    monday = date.fromisocalendar(
        year,
        week,
        1,
    )

    sunday = date.fromisocalendar(
        year,
        week,
        7,
    )

    return (
        f"{monday.month}月{monday.day}日～"
        f"{sunday.month}月{sunday.day}日"
    )


def previous_iso_week(year, week):
    """指定したISO週の前週を返す。"""

    current_monday = date.fromisocalendar(
        year,
        week,
        1,
    )

    previous_monday = (
        current_monday
        - timedelta(days=7)
    )

    previous_iso = (
        previous_monday.isocalendar()
    )

    return (
        previous_iso.year,
        previous_iso.week,
    )


def determine_level(value):
    """
    定点当たり報告数を表示区分へ変換する。

    1未満：非流行
    1以上10未満：流行中
    10以上30未満：注意報レベル
    30以上：警報レベル
    """

    if value is None:
        return "確認中"

    if value >= 30:
        return "警報レベル"

    if value >= 10:
        return "注意報レベル"

    if value >= 1:
        return "流行中"

    return "非流行"


def make_record(
    value,
    previous_value,
    year,
    week,
    period,
    source_name,
    source_url,
):
    """感染症1件分のデータを生成する。"""

    difference = None

    if (
        value is not None
        and previous_value is not None
    ):
        difference = round(
            value - previous_value,
            2,
        )

    return {
        "level": determine_level(value),
        "year": year,
        "week": f"第{week}週",
        "weekNumber": week,
        "period": period,
        "perSentinel": value,
        "previousPerSentinel": previous_value,
        "difference": difference,
        "trend": source_name,
        "provisional": True,
        "dataStatus": "取得成功",
        "sourceUrl": source_url,
    }


def extract_pdf_text(pdf_bytes):
    """PDF全ページからテキストを抽出する。"""

    if not pdf_bytes.startswith(b"%PDF"):
        raise ValueError(
            "取得したデータがPDFではありません"
        )

    reader = PdfReader(
        io.BytesIO(pdf_bytes)
    )

    pages = []

    for page_number, page in enumerate(
        reader.pages,
        start=1,
    ):
        try:
            text = page.extract_text() or ""

        except Exception as error:
            print(
                f"PDF {page_number}ページ目の"
                f"テキスト抽出エラー: {error}"
            )

            text = ""

        if text.strip():
            pages.append(text)

    if not pages:
        raise ValueError(
            "PDFからテキストを抽出できません"
        )

    return "\n".join(pages)


# =============================================================================
# 前回データ
# =============================================================================

def empty_status():
    """初期状態を生成する。"""

    return {
        "nishitama": {
            "influenza": {},
            "covid19": {},
        },
        "tokyo": {
            "influenza": {},
            "covid19": {},
        },
    }


def load_previous_status():
    """前回のstatus.jsonを読み込む。"""

    if not STATUS_FILE.exists():
        return empty_status()

    try:
        value = json.loads(
            STATUS_FILE.read_text(
                encoding="utf-8"
            )
        )

        if isinstance(value, dict):
            return value

    except Exception as error:
        print(
            "前回status.json読込エラー:",
            repr(error),
        )

    return empty_status()


def get_previous_record(
    previous,
    group_name,
    infection_name,
):
    """前回データから指定感染症のデータを取得する。"""

    group = previous.get(
        group_name,
        {},
    )

    if not isinstance(group, dict):
        return {}

    record = group.get(
        infection_name,
        {},
    )

    if isinstance(record, dict):
        return record

    return {}


def fallback_record(
    previous,
    group_name,
    infection_name,
    source_url,
):
    """取得失敗時に前回値を維持する。"""

    old = get_previous_record(
        previous,
        group_name,
        infection_name,
    )

    if old.get("perSentinel") is not None:
        return {
            **old,
            "dataStatus": "取得失敗・前回値を維持",
        }

    return {
        "level": "確認中",
        "year": None,
        "week": "最新発表",
        "weekNumber": None,
        "period": "確認中",
        "perSentinel": None,
        "previousPerSentinel": None,
        "difference": None,
        "trend": "公式情報を確認中",
        "provisional": True,
        "dataStatus": "取得失敗",
        "sourceUrl": source_url,
    }


# =============================================================================
# 西多摩保健所
# =============================================================================

def calculate_link_score(label, url):
    """西多摩週報らしいリンクを点数化する。"""

    text = compact_text(
        f"{label} {url}"
    ).lower()

    score = 0

    if "感染症週報" in text:
        score += 100

    if "西多摩" in text:
        score += 50

    if "感染症" in text:
        score += 20

    if "週報" in text:
        score += 20

    if ".pdf" in text:
        score += 10

    if "/documents/" in text:
        score += 5

    return score


def find_latest_nishitama_report_url(page_html):
    """西多摩保健所ページから最新週報リンクを取得する。"""

    candidates = []

    for position, link in enumerate(
        extract_links(
            page_html,
            NISHITAMA_WEEKLY_PAGE_URL,
        )
    ):
        url = link["url"]
        label = link["label"]

        score = calculate_link_score(
            label,
            url,
        )

        if score < 40:
            continue

        combined = normalize_text(
            f"{label} {url}"
        )

        week_match = re.search(
            r"第?\s*(\d{1,2})\s*週",
            combined,
        )

        year_match = re.search(
            r"(20\d{2})",
            combined,
        )

        week = (
            int(week_match.group(1))
            if week_match
            else -1
        )

        year = (
            int(year_match.group(1))
            if year_match
            else -1
        )

        candidates.append(
            {
                "score": score,
                "year": year,
                "week": week,
                "position": position,
                "url": url,
                "label": label,
            }
        )

    if not candidates:
        raise ValueError(
            "西多摩保健所ページから"
            "感染症週報リンクを検出できません"
        )

    candidates.sort(
        key=lambda item: (
            item["year"],
            item["week"],
            item["score"],
            -item["position"],
        ),
        reverse=True,
    )

    selected = candidates[0]

    print(
        "採用した西多摩週報リンク:",
        json.dumps(
            selected,
            ensure_ascii=False,
        ),
    )

    return selected["url"]


def resolve_pdf_bytes(report_url):
    """
    週報リンクからPDF本体を取得する。

    URLが直接PDFでない場合は、
    遷移先HTML内のPDFリンクを探す。
    """

    first_bytes = fetch_bytes(
        report_url
    )

    if first_bytes.startswith(b"%PDF"):
        return first_bytes, report_url

    first_html = decode_html(
        first_bytes
    )

    pdf_candidates = []

    for link in extract_links(
        first_html,
        report_url,
    ):
        url = link["url"]
        label = link["label"]

        search_text = compact_text(
            f"{label} {url}"
        ).lower()

        if (
            ".pdf" in search_text
            or "ダウンロード" in search_text
            or "感染症週報" in search_text
        ):
            pdf_candidates.append(url)

    for pdf_url in pdf_candidates:
        try:
            pdf_bytes = fetch_bytes(
                pdf_url
            )

            if pdf_bytes.startswith(b"%PDF"):
                return pdf_bytes, pdf_url

        except Exception as error:
            print(
                "PDF候補取得失敗:",
                pdf_url,
                repr(error),
            )

    raise ValueError(
        "西多摩週報リンクからPDF本体を取得できません"
    )


def extract_week_information(pdf_text):
    """PDFから年、週番号、期間を取得する。"""

    normalized = normalize_text(
        pdf_text
    )

    compact = compact_text(
        pdf_text
    )

    week_patterns = [
        r"(20\d{2})年第?(\d{1,2})週",
        r"(20\d{2})年.*?第?(\d{1,2})週",
    ]

    year = None
    week = None

    for pattern in week_patterns:
        match = re.search(
            pattern,
            compact,
        )

        if match:
            year = int(
                match.group(1)
            )

            week = int(
                match.group(2)
            )

            break

    if year is None or week is None:
        raise ValueError(
            "西多摩週報から年・週番号を取得できません"
        )

    complete_match = re.search(
        (
            r"(\d{1,2})\s*月\s*(\d{1,2})\s*日"
            r"\s*[～\-]\s*"
            r"(\d{1,2})\s*月\s*(\d{1,2})\s*日"
        ),
        normalized,
    )

    period = None

    if complete_match:
        period = (
            f"{int(complete_match.group(1))}月"
            f"{int(complete_match.group(2))}日～"
            f"{int(complete_match.group(3))}月"
            f"{int(complete_match.group(4))}日"
        )

    if period is None:
        short_match = re.search(
            (
                r"(\d{1,2})\s*月\s*(\d{1,2})\s*日"
                r"\s*[～\-]\s*"
                r"(\d{1,2})\s*日"
            ),
            normalized,
        )

        if short_match:
            month = int(
                short_match.group(1)
            )

            period = (
                f"{month}月"
                f"{int(short_match.group(2))}日～"
                f"{month}月"
                f"{int(short_match.group(3))}日"
            )

    if period is None:
        period = iso_week_period(
            year,
            week,
        )

    return year, week, period


def extract_value_pair(
    pdf_text,
    disease_patterns,
):
    """
    PDFから最新値と前週値を取得する。

    想定形式:
    インフルエンザ 5.07（3.07）
    新型コロナウイルス感染症 0.79（0.29）
    """

    compact = compact_text(
        pdf_text
    )

    for disease_pattern in disease_patterns:
        patterns = [
            (
                disease_pattern
                + r"[:：]?"
                + r".{0,50}?"
                + r"(\d+(?:\.\d+)?)"
                + r"\("
                + r"(\d+(?:\.\d+)?)"
                + r"\)"
            ),
            (
                disease_pattern
                + r"[:：]?"
                + r".{0,50}?"
                + r"(\d+(?:\.\d+)?)"
                + r"[前週比]"
                + r".{0,20}?"
                + r"(\d+(?:\.\d+)?)"
            ),
        ]

        for pattern in patterns:
            match = re.search(
                pattern,
                compact,
                flags=re.IGNORECASE,
            )

            if not match:
                continue

            current_value = float(
                match.group(1)
            )

            previous_value = float(
                match.group(2)
            )

            if (
                current_value < 0
                or previous_value < 0
                or current_value > 500
                or previous_value > 500
            ):
                continue

            return (
                current_value,
                previous_value,
            )

    return None, None


def get_nishitama_data():
    """西多摩保健所公式週報から感染症情報を取得する。"""

    page_html = fetch_text(
        NISHITAMA_WEEKLY_PAGE_URL
    )

    report_url = find_latest_nishitama_report_url(
        page_html
    )

    pdf_bytes, actual_pdf_url = resolve_pdf_bytes(
        report_url
    )

    pdf_text = extract_pdf_text(
        pdf_bytes
    )

    year, week, period = extract_week_information(
        pdf_text
    )

    influenza_value, influenza_previous = (
        extract_value_pair(
            pdf_text,
            [
                r"インフルエンザ",
            ],
        )
    )

    covid_value, covid_previous = (
        extract_value_pair(
            pdf_text,
            [
                r"新型コロナウイルス感染症",
                r"新型コロナ感染症",
                r"新型コロナ",
                r"COVID-?19",
            ],
        )
    )

    if influenza_value is None:
        raise ValueError(
            "西多摩週報から"
            "インフルエンザの最新値・前週値を"
            "取得できません"
        )

    if covid_value is None:
        raise ValueError(
            "西多摩週報から"
            "新型コロナの最新値・前週値を"
            "取得できません"
        )

    influenza = make_record(
        value=influenza_value,
        previous_value=influenza_previous,
        year=year,
        week=week,
        period=period,
        source_name="西多摩保健所感染症週報",
        source_url=actual_pdf_url,
    )

    covid19 = make_record(
        value=covid_value,
        previous_value=covid_previous,
        year=year,
        week=week,
        period=period,
        source_name="西多摩保健所感染症週報",
        source_url=actual_pdf_url,
    )

    return {
        "influenza": influenza,
        "covid19": covid19,
        "debug": {
            "reportPageUrl": report_url,
            "sourceUrl": actual_pdf_url,
            "year": year,
            "week": week,
            "period": period,
            "influenzaValue": influenza_value,
            "influenzaPrevious": influenza_previous,
            "covid19Value": covid_value,
            "covid19Previous": covid_previous,
        },
    }


# =============================================================================
# 東京都インフルエンザ
# =============================================================================

def find_tokyo_flu_excel_url(top_html):
    """東京都公式ページから患者報告数Excelを取得する。"""

    candidates = []

    for link in extract_links(
        top_html,
        TOKYO_FLU_TOP_URL,
    ):
        url = link["url"]
        label = link["label"]
        lower_url = url.lower()

        is_excel = (
            ".xlsx" in lower_url
            or ".xls" in lower_url
        )

        is_report_file = (
            "hasseidoko" in lower_url
            or "患者報告数" in label
        )

        if is_excel and is_report_file:
            candidates.append(url)

    if not candidates:
        raise ValueError(
            "東京都公式ページから"
            "インフルエンザ患者報告数Excelを"
            "取得できません"
        )

    return candidates[0]


def parse_week_header(value):
    """Excelの週番号見出しを解析する。"""

    text = compact_text(
        value
    ).lower()

    match = re.fullmatch(
        r"(?:第)?(\d{1,2})(?:w|週)",
        text,
    )

    if not match:
        return None

    week = int(
        match.group(1)
    )

    if 1 <= week <= 53:
        return week

    return None


def find_tokyo_flu_records(workbook):
    """東京都Excelから最新週と前週を取得する。"""

    now = datetime.now(JST)
    candidates = []

    for sheet in workbook.worksheets:
        print(
            "東京都Excel解析シート:",
            sheet.title,
        )

        header_row = None
        week_columns = {}

        for row in range(
            1,
            min(sheet.max_row, 40) + 1,
        ):
            found = {}

            for column in range(
                1,
                sheet.max_column + 1,
            ):
                week = parse_week_header(
                    sheet.cell(
                        row=row,
                        column=column,
                    ).value
                )

                if week is not None:
                    found[column] = week

            if len(found) > len(week_columns):
                header_row = row
                week_columns = found

        if not header_row or not week_columns:
            continue

        total_row = None
        sentinel_row = None

        for row in range(
            header_row + 1,
            sheet.max_row + 1,
        ):
            labels = [
                compact_text(
                    sheet.cell(
                        row=row,
                        column=column,
                    ).value
                ).lower()
                for column in range(
                    1,
                    min(
                        sheet.max_column,
                        10,
                    ) + 1,
                )
            ]

            if any(
                label in {
                    "total",
                    "合計",
                    "東京都計",
                }
                for label in labels
            ):
                total_row = row

            if any(
                (
                    "定点数" in label
                    or "テイテンスウ" in label
                )
                for label in labels
            ):
                sentinel_row = row

        if (
            total_row is None
            or sentinel_row is None
        ):
            continue

        sheet_text = " ".join(
            normalize_text(
                sheet.cell(
                    row=row,
                    column=column,
                ).value
            )
            for row in range(
                1,
                min(sheet.max_row, 12) + 1,
            )
            for column in range(
                1,
                min(sheet.max_column, 15) + 1,
            )
        )

        season_match = re.search(
            r"(20\d{2})\s*[-－]",
            sheet_text,
        )

        if season_match:
            season_start_year = int(
                season_match.group(1)
            )

        else:
            season_start_year = (
                now.year
                if now.month >= 8
                else now.year - 1
            )

        for column, week in week_columns.items():
            total = number(
                sheet.cell(
                    row=total_row,
                    column=column,
                ).value
            )

            sentinel_count = number(
                sheet.cell(
                    row=sentinel_row,
                    column=column,
                ).value
            )

            if (
                total is None
                or sentinel_count is None
                or sentinel_count <= 0
            ):
                continue

            year = (
                season_start_year
                if week >= 36
                else season_start_year + 1
            )

            try:
                week_end = date.fromisocalendar(
                    year,
                    week,
                    7,
                )

            except ValueError:
                continue

            if week_end > now.date():
                continue

            candidates.append(
                {
                    "year": year,
                    "week": week,
                    "value": round(
                        total / sentinel_count,
                        2,
                    ),
                    "reportedCases": int(total),
                    "sentinelCount": int(
                        sentinel_count
                    ),
                    "sheet": sheet.title,
                }
            )

    if not candidates:
        raise ValueError(
            "東京都インフルエンザExcelから"
            "最新週を抽出できません"
        )

    unique_records = {}

    for item in candidates:
        key = (
            item["year"],
            item["week"],
        )

        unique_records[key] = item

    records = sorted(
        unique_records.values(),
        key=lambda item: (
            item["year"],
            item["week"],
        ),
    )

    latest = records[-1]

    previous = (
        records[-2]
        if len(records) >= 2
        else None
    )

    print(
        "東京都インフルエンザ最新週:",
        latest,
    )

    print(
        "東京都インフルエンザ前週:",
        previous,
    )

    return latest, previous


def get_tokyo_influenza_data():
    """東京都公式Excelからインフルエンザ情報を取得する。"""

    top_html = fetch_text(
        TOKYO_FLU_TOP_URL
    )

    excel_url = find_tokyo_flu_excel_url(
        top_html
    )

    print(
        "東京都インフルエンザExcel:",
        excel_url,
    )

    excel_bytes = fetch_bytes(
        excel_url
    )

    if not excel_bytes.startswith(b"PK"):
        raise ValueError(
            "取得した東京都ファイルが"
            "有効なxlsx形式ではありません"
        )

    workbook = load_workbook(
        io.BytesIO(excel_bytes),
        data_only=True,
        read_only=True,
    )

    try:
        latest, previous = (
            find_tokyo_flu_records(
                workbook
            )
        )

    finally:
        workbook.close()

    previous_value = (
        previous["value"]
        if previous
        else None
    )

    record = make_record(
        value=latest["value"],
        previous_value=previous_value,
        year=latest["year"],
        week=latest["week"],
        period=iso_week_period(
            latest["year"],
            latest["week"],
        ),
        source_name="東京都感染症情報センター",
        source_url=excel_url,
    )

    record["reportedCases"] = (
        latest["reportedCases"]
    )

    record["sentinelCount"] = (
        latest["sentinelCount"]
    )

    return record


# =============================================================================
# 東京都新型コロナ
# =============================================================================

def find_latest_tokyo_weekly_pdf(page_html):
    """
    東京都感染症週報ページから最新PDFを検出する。

    想定URL:
    https://idsc.tmiph.metro.tokyo.lg.jp/assets/weekly/2026/46.pdf
    """

    now = datetime.now(JST)
    candidates = []

    for link in extract_links(
        page_html,
        TOKYO_WEEKLY_PAGE_URL,
    ):
        url = link["url"]
        label = link["label"]

        parsed = urlparse(url)

        match = re.search(
            r"/assets/weekly/(20\d{2})/(\d{1,2})\.pdf$",
            parsed.path,
            flags=re.IGNORECASE,
        )

        if not match:
            continue

        year = int(
            match.group(1)
        )

        week = int(
            match.group(2)
        )

        try:
            week_end = date.fromisocalendar(
                year,
                week,
                7,
            )

        except ValueError:
            continue

        if week_end > now.date():
            continue

        candidates.append(
            {
                "year": year,
                "week": week,
                "url": url,
                "label": label,
            }
        )

    if not candidates:
        raise ValueError(
            "東京都感染症週報ページから"
            "最新週報PDFを検出できません"
        )

    candidates.sort(
        key=lambda item: (
            item["year"],
            item["week"],
        )
    )

    selected = candidates[-1]

    print(
        "採用した東京都感染症週報:",
        json.dumps(
            selected,
            ensure_ascii=False,
        ),
    )

    return selected


def make_tokyo_weekly_pdf_url(year, week):
    """年と週番号から東京都感染症週報PDFのURLを生成する。"""

    return (
        f"{TOKYO_WEEKLY_PDF_BASE_URL}/"
        f"{year}/{week}.pdf"
    )


def find_tokyo_rate_section(pdf_text):
    """
    東京都感染症週報から、
    保健所別の定点医療機関当たり報告数表を取得する。
    """

    compact = compact_text(
        pdf_text
    )

    start_marker = (
        "定点把握対象疾患"
        "定点医療機関当たり報告数"
        "【保健所別】"
    )

    start = compact.find(
        start_marker
    )

    if start < 0:
        start_marker = (
            "定点把握対象疾患"
            "定点医療機関当たり報告数"
        )

        start = compact.find(
            start_marker
        )

    if start < 0:
        raise ValueError(
            "東京都感染症週報から"
            "保健所別定点当たり報告数表を"
            "検出できません"
        )

    end_markers = [
        "定点把握対象疾患報告数【保健所別】",
        "定点把握対象疾患報告数",
    ]

    end_candidates = []

    for end_marker in end_markers:
        position = compact.find(
            end_marker,
            start + len(start_marker),
        )

        if position >= 0:
            end_candidates.append(position)

    if end_candidates:
        end = min(end_candidates)
        return compact[start:end]

    return compact[start:]


def extract_tokyo_covid_rate(pdf_text):
    """
    東京都感染症週報から東京都全体の
    新型コロナ定点当たり報告数を取得する。

    対象行の列順:
    東京都
    1列目 インフルエンザ
    2列目 新型コロナ
    3列目 急性呼吸器感染症
    """

    section = find_tokyo_rate_section(
        pdf_text
    )

    candidates = []

    for match in re.finditer(
        r"東京都(?!合計)",
        section,
    ):
        nearby = section[
            match.end():
            match.end() + 250
        ]

        values = [
            float(value)
            for value in re.findall(
                r"(?<![\d.])"
                r"(\d+\.\d+)"
                r"(?![\d.])",
                nearby,
            )
        ]

        if len(values) < 2:
            continue

        influenza_value = values[0]
        covid_value = values[1]

        if not (
            0 <= influenza_value <= 500
            and 0 <= covid_value <= 100
        ):
            continue

        candidates.append(
            {
                "influenza": round(
                    influenza_value,
                    2,
                ),
                "covid19": round(
                    covid_value,
                    2,
                ),
                "values": values[:12],
                "nearby": nearby[:250],
            }
        )

    if not candidates:
        raise ValueError(
            "東京都感染症週報から"
            "東京都行の定点当たり報告数を"
            "取得できません"
        )

    # 小児科系の東京都行が先にあり、
    # インフルエンザ・コロナ系の東京都行が後にあるため、
    # 最後の有効候補を採用する。
    selected = candidates[-1]

    print(
        "東京都定点値・採用候補:",
        json.dumps(
            selected,
            ensure_ascii=False,
        ),
    )

    return selected["covid19"]


def extract_tokyo_covid_reported_cases(pdf_text):
    """
    東京都感染症週報から東京都全体の
    新型コロナ報告数を取得する。

    対象行の列順:
    東京都合計
    1列目 インフルエンザ
    2列目 新型コロナ
    3列目 急性呼吸器感染症
    """

    compact = compact_text(
        pdf_text
    )

    candidates = []

    for match in re.finditer(
        r"東京都合計",
        compact,
    ):
        nearby = compact[
            match.end():
            match.end() + 220
        ]

        values = [
            int(value.replace(",", ""))
            for value in re.findall(
                r"(?<![\d.])"
                r"(\d{1,3}(?:,\d{3})+|\d+)"
                r"(?![\d.])",
                nearby,
            )
        ]

        if len(values) < 2:
            continue

        influenza_count = values[0]
        covid_count = values[1]

        if not (
            0 <= influenza_count <= 1000000
            and 0 <= covid_count <= 1000000
        ):
            continue

        candidates.append(
            {
                "influenza": influenza_count,
                "covid19": covid_count,
                "values": values[:12],
            }
        )

    if not candidates:
        print(
            "東京都新型コロナ報告数を"
            "取得できませんでした"
        )

        return None

    # インフルエンザ・コロナ系の東京都合計行は
    # PDF後半に掲載されるため最後の候補を採用する。
    selected = candidates[-1]

    print(
        "東京都報告数・採用候補:",
        json.dumps(
            selected,
            ensure_ascii=False,
        ),
    )

    return selected["covid19"]


def get_tokyo_covid19_data():
    """
    最新週と前週の東京都感染症週報PDFから
    東京都全体の新型コロナ情報を取得する。
    """

    page_html = fetch_text(
        TOKYO_WEEKLY_PAGE_URL
    )

    latest_report = find_latest_tokyo_weekly_pdf(
        page_html
    )

    latest_pdf_bytes = fetch_bytes(
        latest_report["url"]
    )

    latest_pdf_text = extract_pdf_text(
        latest_pdf_bytes
    )

    current_value = extract_tokyo_covid_rate(
        latest_pdf_text
    )

    reported_cases = (
        extract_tokyo_covid_reported_cases(
            latest_pdf_text
        )
    )

    previous_year, previous_week = (
        previous_iso_week(
            latest_report["year"],
            latest_report["week"],
        )
    )

    previous_url = make_tokyo_weekly_pdf_url(
        previous_year,
        previous_week,
    )

    print(
        "東京都新型コロナ前週PDF:",
        previous_url,
    )

    previous_pdf_bytes = fetch_bytes(
        previous_url
    )

    previous_pdf_text = extract_pdf_text(
        previous_pdf_bytes
    )

    previous_value = extract_tokyo_covid_rate(
        previous_pdf_text
    )

    record = make_record(
        value=current_value,
        previous_value=previous_value,
        year=latest_report["year"],
        week=latest_report["week"],
        period=iso_week_period(
            latest_report["year"],
            latest_report["week"],
        ),
        source_name="東京都感染症週報",
        source_url=latest_report["url"],
    )

    record["reportedCases"] = (
        reported_cases
    )

    record["previousSourceUrl"] = (
        previous_url
    )

    return record


# =============================================================================
# status.json生成
# =============================================================================

def main():
    """感染症情報を取得してstatus.jsonを更新する。"""

    previous = load_previous_status()

    errors = []
    debug = {}

    now_text = datetime.now(
        JST
    ).strftime(
        "%Y-%m-%d %H:%M"
    )

    # -------------------------------------------------------------------------
    # 西多摩
    # -------------------------------------------------------------------------

    try:
        nishitama_result = (
            get_nishitama_data()
        )

        nishitama_influenza = (
            nishitama_result["influenza"]
        )

        nishitama_covid19 = (
            nishitama_result["covid19"]
        )

        debug["nishitama"] = (
            nishitama_result["debug"]
        )

    except Exception as error:
        print(
            "西多摩感染症情報取得エラー:",
            repr(error),
        )

        errors.append(
            "西多摩感染症情報"
        )

        nishitama_influenza = fallback_record(
            previous=previous,
            group_name="nishitama",
            infection_name="influenza",
            source_url=NISHITAMA_WEEKLY_PAGE_URL,
        )

        nishitama_covid19 = fallback_record(
            previous=previous,
            group_name="nishitama",
            infection_name="covid19",
            source_url=NISHITAMA_WEEKLY_PAGE_URL,
        )

        debug["nishitama"] = {
            "errorType": type(error).__name__,
            "errorMessage": str(error),
            "fallback": "previousStatus",
        }

    # -------------------------------------------------------------------------
    # 東京都インフルエンザ
    # -------------------------------------------------------------------------

    try:
        tokyo_influenza = (
            get_tokyo_influenza_data()
        )

    except Exception as error:
        print(
            "東京都インフルエンザ取得エラー:",
            repr(error),
        )

        errors.append(
            "東京都インフルエンザ"
        )

        tokyo_influenza = fallback_record(
            previous=previous,
            group_name="tokyo",
            infection_name="influenza",
            source_url=TOKYO_FLU_TOP_URL,
        )

        debug["tokyoInfluenza"] = {
            "errorType": type(error).__name__,
            "errorMessage": str(error),
            "fallback": "previousStatus",
        }

    # -------------------------------------------------------------------------
    # 東京都新型コロナ
    # -------------------------------------------------------------------------

    try:
        tokyo_covid19 = (
            get_tokyo_covid19_data()
        )

        debug["tokyoCovid19"] = {
            "sourceUrl": tokyo_covid19.get(
                "sourceUrl"
            ),
            "previousSourceUrl": (
                tokyo_covid19.get(
                    "previousSourceUrl"
                )
            ),
            "year": tokyo_covid19.get(
                "year"
            ),
            "week": tokyo_covid19.get(
                "week"
            ),
            "value": tokyo_covid19.get(
                "perSentinel"
            ),
            "previousValue": (
                tokyo_covid19.get(
                    "previousPerSentinel"
                )
            ),
            "reportedCases": (
                tokyo_covid19.get(
                    "reportedCases"
                )
            ),
        }

    except Exception as error:
        print(
            "東京都新型コロナ取得エラー:",
            repr(error),
        )

        errors.append(
            "東京都新型コロナ"
        )

        tokyo_covid19 = fallback_record(
            previous=previous,
            group_name="tokyo",
            infection_name="covid19",
            source_url=TOKYO_WEEKLY_PAGE_URL,
        )

        debug["tokyoCovid19"] = {
            "errorType": type(error).__name__,
            "errorMessage": str(error),
            "fallback": "previousStatus",
        }

    # -------------------------------------------------------------------------
    # status.json
    # -------------------------------------------------------------------------

    status = {
        "updated": now_text,
        "sourceStatus": (
            "正常"
            if not errors
            else (
                "一部取得失敗："
                + "・".join(errors)
            )
        ),
        "area": {
            "displayName": "西多摩地域",
            "dataArea": "西多摩保健所管内",
        },
        "nishitama": {
            "influenza": nishitama_influenza,
            "covid19": nishitama_covid19,
        },
        "tokyo": {
            "influenza": tokyo_influenza,
            "covid19": tokyo_covid19,
        },
        "debug": debug,
    }

    STATUS_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary_file = (
        STATUS_FILE.with_suffix(".tmp")
    )

    temporary_file.write_text(
        json.dumps(
            status,
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    # 書き出したファイルが正常なJSONか確認する。
    json.loads(
        temporary_file.read_text(
            encoding="utf-8"
        )
    )

    # 正常なJSONの場合だけstatus.jsonへ置き換える。
    temporary_file.replace(
        STATUS_FILE
    )

    print(
        json.dumps(
            status,
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
