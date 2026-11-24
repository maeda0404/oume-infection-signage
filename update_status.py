#!/usr/bin/env python3
"""
西多摩地域 感染症流行状況サイネージ用 status.json 生成スクリプト

【主データ】
東京都西多摩保健所「感染症週報」

取得情報
・西多摩保健所管内 インフルエンザ定点当たり報告数
・西多摩保健所管内 新型コロナ定点当たり報告数
・各感染症の前週値、増減、対象週、対象期間

【補助データ】
・東京都感染症情報センター
  インフルエンザ患者報告数Excel
・東京都感染症週報PDF
  新型コロナウイルス感染症の東京都値

取得失敗時は、前回の status.json の値を維持する。
取得に失敗した値を 0.00 として出力しない。
"""

import gzip
import io
import json
import re
import time
import urllib.request
from datetime import date, datetime, timedelta, timezone
from html import unescape
from pathlib import Path
from urllib.parse import urljoin, urlparse

from openpyxl import load_workbook
from pypdf import PdfReader


# ---------------------------------------------------------------------------
# 基本設定
# ---------------------------------------------------------------------------

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
    "https://idsc.tmiph.metro.tokyo.lg.jp/weekly/"
)

USER_AGENT = (
    "Mozilla/5.0 "
    "Nishitama-Infection-Signage/1.0"
)


# ---------------------------------------------------------------------------
# 通信処理
# ---------------------------------------------------------------------------

def fetch_bytes(url, retries=3):
    """指定URLからバイトデータを取得する。"""

    last_error = None

    for attempt in range(1, retries + 1):
        try:
            separator = "&" if "?" in url else "?"
            cache_buster = int(
                datetime.now(timezone.utc).timestamp()
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
                        "text/html,*/*"
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


def fetch_text(url):
    """指定URLからHTMLテキストを取得する。"""

    raw = fetch_bytes(url)

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


# ---------------------------------------------------------------------------
# 共通処理
# ---------------------------------------------------------------------------

def normalize_text(value):
    """文字列内の空白や全角記号を正規化する。"""

    if value is None:
        return ""

    text = str(value)

    text = text.replace("\u3000", " ")
    text = text.replace("（", "(")
    text = text.replace("）", ")")
    text = text.replace("－", "-")
    text = text.replace("〜", "～")
    text = text.replace("~", "～")

    return re.sub(
        r"\s+",
        " ",
        text,
    ).strip()


def compact_text(value):
    """解析用に改行と余分な空白を除去する。"""

    return re.sub(
        r"\s+",
        "",
        normalize_text(value),
    )


def html_to_text(html):
    """HTMLタグを除去して表示文字列を取得する。"""

    text = re.sub(
        r"<script[\s\S]*?</script>",
        " ",
        html,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"<style[\s\S]*?</style>",
        " ",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"<[^>]+>",
        " ",
        text,
    )

    return normalize_text(
        unescape(text)
    )


def number(value):
    """セル値や文字列を数値へ変換する。"""

    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = compact_text(value)

    text = (
        text
        .replace(",", "")
        .replace("人", "")
        .replace("人/定点", "")
        .replace("人／定点", "")
    )

    if re.fullmatch(
        r"-?\d+(?:\.\d+)?",
        text,
    ):
        return float(text)

    return None


def extract_links(html, base_url):
    """HTML内のリンクと表示文字列を抽出する。"""

    links = []

    pattern = re.compile(
        r'<a\b[^>]*?href\s*=\s*["\']([^"\']+)["\'][^>]*>'
        r'([\s\S]*?)</a>',
        flags=re.IGNORECASE,
    )

    for href, label_html in pattern.findall(html):
        url = urljoin(
            base_url,
            unescape(href),
        )

        label = html_to_text(
            label_html
        )

        links.append(
            {
                "url": url,
                "label": label,
            }
        )

    return links


def extract_pdf_text(pdf_bytes):
    """PDF全ページからテキストを取得する。"""

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
                f"テキスト抽出失敗: {error}"
            )
            text = ""

        if text:
            pages.append(text)

    if not pages:
        raise ValueError(
            "PDFからテキストを抽出できません"
        )

    return "\n".join(pages)


def load_previous_status():
    """前回のstatus.jsonを読み込む。"""

    empty_status = {
        "nishitama": {
            "influenza": {},
            "covid19": {},
        },
        "tokyo": {
            "influenza": {},
            "covid19": {},
        },
    }

    if not STATUS_FILE.exists():
        return empty_status

    try:
        value = json.loads(
            STATUS_FILE.read_text(
                encoding="utf-8"
            )
        )

        if not isinstance(value, dict):
            return empty_status

        return value

    except Exception as error:
        print(
            "前回データ読込エラー:",
            repr(error),
        )

        return empty_status


def iso_week_period(year, week):
    """ISO週から月曜日～日曜日の期間を生成する。"""

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


def determine_level(value):
    """
    定点当たり報告数から表示区分を決定する。

    1未満     非流行
    1以上10未満 流行中
    10以上30未満 注意報レベル
    30以上    警報レベル
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


def make_infection_record(
    value,
    previous_value,
    year,
    week,
    period,
    source_name,
    source_url,
):
    """感染症1件分のJSONデータを生成する。"""

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


# ---------------------------------------------------------------------------
# 西多摩保健所週報
# ---------------------------------------------------------------------------

def find_latest_nishitama_pdf_url(page_html):
    """西多摩保健所ページから最新週報PDFを取得する。"""

    candidates = []

    for link in extract_links(
        page_html,
        NISHITAMA_WEEKLY_PAGE_URL,
    ):
        url = link["url"]
        label = link["label"]

        search_text = (
            compact_text(label)
            + compact_text(url)
        )

        if (
            "西多摩" not in search_text
            and "感染症週報" not in search_text
        ):
            continue

        week_match = re.search(
            r"第?(\d{1,2})週",
            normalize_text(label),
        )

        week = (
            int(week_match.group(1))
            if week_match
            else -1
        )

        candidates.append(
            {
                "week": week,
                "url": url,
                "label": label,
            }
        )

    if not candidates:
        raise ValueError(
            "西多摩保健所ページから"
            "感染症週報リンクを取得できません"
        )

    candidates.sort(
        key=lambda item: item["week"],
        reverse=True,
    )

    selected = candidates[0]

    print(
        "西多摩週報リンク:",
        selected,
    )

    return selected["url"]


def extract_nishitama_week_information(pdf_text):
    """西多摩週報から年・週・期間を取得する。"""

    normalized = normalize_text(pdf_text)

    week_match = re.search(
        r"(20\d{2})\s*年\s*第?\s*(\d{1,2})\s*週",
        normalized,
    )

    if not week_match:
        compact = compact week_match = re.search(
            r"(20\d{2})年第?(\d{1,2})週",
            compact,
        )

    if not week_match:
        raise ValueError(
            "西多摩週報から年と週番号を取得できません"
        )

    year = int(
        week_match.group(1)
    )

    week = int(
        week_match.group(2)
    )

    period_match = re.search(
        r"(\d{1,2})\s*月\s*(\d{1,2})\s*日"
        r"\s*[～\-]\s*"
        r"(?:(\d{1,2})\s*月\s*)?"
        r"(\d{1,2})\s*日",
        normalized,
    )

    if period_match:
        start_month = int(
            period_match.group(1)
        )

        start_day = int(
            period_match.group(2)
        )

        end_month = int(
            period_match.group(3)
        ) if period_match.group(3) else start_month

        end_day = int(
            period_match.group(4)
        )

        period = (
            f"{start_month}月{start_day}日～"
            f"{end_month}月{end_day}日"
        )

    else:
        period = iso_week_period(
            year,
            week,
        )

    return year, week, period


def extract_current_previous_pair(
    pdf_text,
    disease_patterns,
):
    """
    疾患名の後にある「最新値（前週値）」を取得する。

    例：
    インフルエンザ5.07（3.07）
    新型コロナウイルス感染症0.79（0.29）
    """

    compact = compact_text(pdf_text)

    for disease_pattern in disease_patterns:
        pattern = re.compile(
            disease_pattern
            + r"[:：]?"
            + r"(\d+(?:\.\d+)?)"
            + r"\("
            + r"(\d+(?:\.\d+)?)"
            + r"\)"
        )

        match = pattern.search(
            compact
        )

        if not match:
            continue

        current_value = float(
            match.group(1)
        )

        previous_value = float(
            match.group(2)
        )

        return current_value, previous_value

    return None, None


def get_nishitama_data():
    """西多摩保健所公式週報から感染症情報を取得する。"""

    page_html = fetch_text(
        NISHITAMA_WEEKLY_PAGE_URL
    )

    pdf_url = find_latest_nishitama_pdf_url(
        page_html
    )

    pdf_bytes = fetch_bytes(
        pdf_url
    )

    if not pdf_bytes.startswith(b"%PDF"):
        raise ValueError(
            "西多摩週報リンクからPDFを取得できません"
        )

    pdf_text = extract_pdf_text(
        pdf_bytes
    )

    year, week, period = (
        extract_nishitama_week_information(
            pdf_text
        )
    )

    flu_value, flu_previous = (
        extract_current_previous_pair(
            pdf_text,
            [
                r"インフルエンザ",
            ],
        )
    )

    covid_value, covid_previous = (
        extract_current_previous_pair(
            pdf_text,
            [
                r"新型コロナウイルス感染症",
                r"新型コロナ感染症",
                r"新型コロナ",
                r"COVID-?19",
            ],
        )
    )

    if flu_value is None:
        raise ValueError(
            "西多摩週報から"
            "インフルエンザ値を取得できません"
        )

    if covid_value is None:
        raise ValueError(
            "西多摩週報から"
            "新型コロナ値を取得できません"
        )

    influenza = make_infection_record(
        value=flu_value,
        previous_value=flu_previous,
        year=year,
        week=week,
        period=period,
        source_name="西多摩保健所感染症週報",
        source_url=pdf_url,
    )

    covid19 = make_infection_record(
        value=covid_value,
        previous_value=covid_previous,
        year=year,
        week=week,
        period=period,
        source_name="西多摩保健所感染症週報",
        source_url=pdf_url,
    )

    return {
        "influenza": influenza,
        "covid19": covid19,
        "debug": {
            "sourceUrl": pdf_url,
            "year": year,
            "week": week,
            "period": period,
            "influenzaMatched": True,
            "covid19Matched": True,
        },
    }


# ---------------------------------------------------------------------------
# 東京都インフルエンザ
# ---------------------------------------------------------------------------

def find_tokyo_flu_excel_url(top_html):
    """東京都公式ページから患者報告数Excelを取得する。"""

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
            return url

    raise ValueError(
        "東京都公式ページから"
        "インフルエンザ患者報告数Excelを"
        "取得できません"
    )


def parse_week_header(value):
    """Excelの週番号見出しを解析する。"""

    text = compact_text(value).lower()

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
    """
    東京都インフルエンザExcelから
    最新週と前週を取得する。
    """

    now = datetime.now(JST)
    candidates = []

    for sheet in workbook.worksheets:
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
                    min(sheet.max_column, 10) + 1,
                )
            ]

            if any(
                label in {"total", "合計"}
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

        if total_row is None or sentinel_row is None:
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
                min(sheet.max_row, 10) + 1,
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

            points = number(
                sheet.cell(
                    row=sentinel_row,
                    column=column,
                ).value
            )

            if (
                total is None
                or points is None
                or points <= 0
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
                        total / points,
                        2,
                    ),
                    "reportedCases": int(total),
                    "sentinelCount": int(points),
                    "sheet": sheet.title,
                }
            )

    if not candidates:
        raise ValueError(
            "東京都インフルエンザExcelから"
            "有効な週次データを取得できません"
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

    return latest, previous


def get_tokyo_influenza_data():
    """東京都のインフルエンザ情報を取得する。"""

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

    record = make_infection_record(
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


# ---------------------------------------------------------------------------
# 東京都新型コロナ
# ---------------------------------------------------------------------------

def find_latest_tokyo_weekly_pdf_url(page_html):
    """東京都感染症週報ページから最新PDFを取得する。"""

    now = datetime.now(JST)

    candidates = []

    for link in extract_links(
        page_html,
        TOKYO_WEEKLY_PAGE_URL,
    ):
        url = link["url"]
        label = normalize_text(
            link["label"]
        )

        parsed = urlparse(url)
        path = parsed.path.lower()

        match = re.search(
            r"/assets/weekly/(20\d{2})/"
            r"(\d{1,2})\.pdf$",
            path,
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
            "東京都感染症週報の"
            "最新PDFを取得できません"
        )

    candidates.sort(
        key=lambda item: (
            item["year"],
            item["week"],
        ),
    )

    return candidates[-1]


def extract_tokyo_covid_value(pdf_text):
    """
    東京都感染症週報の定点把握対象疾患表から
    新型コロナの東京都値を取得する。

    PDFの表抽出順序が変わる可能性があるため、
    複数パターンを試す。
    """

    lines = [
        normalize_text(line)
        for line in pdf_text.splitlines()
        if normalize_text(line)
    ]

    disease_names = (
        "新型コロナウイルス感染症",
        "新型コロナ",
        "COVID-19",
    )

    candidates = []

    for index, line in enumerate(lines):
        if not any(
            name in line
            for name in disease_names
        ):
            continue

        nearby = " ".join(
            lines[
                index:
                min(index + 5, len(lines))
            ]
        )

        values = [
            float(value)
            for value in re.findall(
                r"(?<!\d)(\d+\.\d+)(?!\d)",
                nearby,
            )
        ]

        for value in values:
            if 0 <= value <= 200:
                candidates.append(value)

        if candidates:
            break

    if not candidates:
        compact = compact_text(
            pdf_text
        )

        for disease_name in disease_names:
            pattern = re.compile(
                re.escape(disease_name)
                + r".{0,120}?"
                + r"(\d+\.\d+)"
            )

            match = pattern.search(
                compact
            )

            if match:
                value = float(
                    match.group(1)
                )

                if 0 <= value <= 200:
                    candidates.append(value)
                    break

    if not candidates:
        raise ValueError(
            "東京都感染症週報から"
            "新型コロナの東京都値を取得できません"
        )

    return candidates[0]


def get_tokyo_covid_data():
    """東京都感染症週報から新型コロナ値を取得する。"""

    page_html = fetch_text(
        TOKYO_WEEKLY_PAGE_URL
    )

    latest_info = (
        find_latest_tokyo_weekly_pdf_url(
            page_html
        )
    )

    latest_pdf_bytes = fetch_bytes(
        latest_info["url"]
    )

    latest_pdf_text = extract_pdf_text(
        latest_pdf_bytes
    )

    latest_value = extract_tokyo_covid_value(
        latest_pdf_text
    )

    previous_value = None

    try:
        previous_year = latest_info["year"]
        previous_week = latest_info["week"] - 1

        if previous_week < 1:
            previous_year -= 1
            previous_week = date(
                previous_year,
                12,
                28,
            ).isocalendar().week

        previous_url = (
            "https://idsc.tmiph.metro.tokyo.lg.jp/"
            f"assets/weekly/{previous_year}/"
            f"{previous_week:02d}.pdf"
        )

        previous_pdf_bytes = fetch_bytes(
            previous_url
        )

        previous_pdf_text = extract_pdf_text(
            previous_pdf_bytes
        )

        previous_value = extract_tokyo_covid_value(
            previous_pdf_text
        )

    except Exception as error:
        print(
            "東京都コロナ前週値取得エラー:",
            repr(error),
        )

    return make_infection_record(
        value=latest_value,
        previous_value=previous_value,
        year=latest_info["year"],
        week=latest_info["week"],
        period=iso_week_period(
            latest_info["year"],
            latest_info["week"],
        ),
        source_name="東京都感染症週報",
        source_url=latest_info["url"],
    )


# ---------------------------------------------------------------------------
# 取得失敗時の維持処理
# ---------------------------------------------------------------------------

def previous_record(
    previous,
    group_name,
    infection_name,
):
    """前回データから感染症情報を取得する。"""

    group = previous.get(
        group_name,
        {}
    )

    if not isinstance(group, dict):
        return {}

    record = group.get(
        infection_name,
        {}
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
    """前回値を維持するか、確認中データを生成する。"""

    old = previous_record(
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


# ---------------------------------------------------------------------------
# status.json生成
# ---------------------------------------------------------------------------

def main():
    """公式データを取得し、status.jsonを生成する。"""

    previous = load_previous_status()

    errors = []
    debug = {}

    now_text = datetime.now(
        JST
    ).strftime(
        "%Y-%m-%d %H:%M"
    )

    # -----------------------------------------------------------------------
    # 西多摩
    # -----------------------------------------------------------------------

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

        nishitama_influenza = (
            fallback_record(
                previous,
                "nishitama",
                "influenza",
                NISHITAMA_WEEKLY_PAGE_URL,
            )
        )

        nishitama_covid19 = (
            fallback_record(
                previous,
                "nishitama",
                "covid19",
                NISHITAMA_WEEKLY_PAGE_URL,
            )
        )

        debug["nishitama"] = {
            "errorType": type(error).__name__,
            "errorMessage": str(error),
            "fallback": "previousStatus",
        }

    # -----------------------------------------------------------------------
    # 東京都インフルエンザ
    # -----------------------------------------------------------------------

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
            previous,
            "tokyo",
            "influenza",
            TOKYO_FLU_TOP_URL,
        )

        debug["tokyoInfluenza"] = {
            "errorType": type(error).__name__,
            "errorMessage": str(error),
            "fallback": "previousStatus",
        }

    # -----------------------------------------------------------------------
    # 東京都新型コロナ
    # -----------------------------------------------------------------------

    try:
        tokyo_covid19 = (
            get_tokyo_covid_data()
        )

    except Exception as error:
        print(
            "東京都新型コロナ取得エラー:",
            repr(error),
        )

        errors.append(
            "東京都新型コロナ"
        )

        tokyo_covid19 = fallback_record(
            previous,
            "tokyo",
            "covid19",
            TOKYO_WEEKLY_PAGE_URL,
        )

        debug["tokyoCovid19"] = {
            "errorType": type(error).__name__,
            "errorMessage": str(error),
            "fallback": "previousStatus",
        }

    # -----------------------------------------------------------------------
    # status.json
    # -----------------------------------------------------------------------

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

    # JSONとして正常か確認
    json.loads(
        temporary_file.read_text(
            encoding="utf-8"
        )
    )

    # 正常な場合だけstatus.jsonへ置き換える
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
