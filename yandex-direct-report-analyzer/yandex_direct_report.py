#!/usr/bin/env python3
"""Выгрузка детализированного отчета из Reports API Яндекс Директа в CSV."""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys
import time
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable, Iterator, Mapping, Sequence

import requests
from dotenv import load_dotenv


load_dotenv(Path(__file__).resolve().parent / ".env")


# ---------------------------------------------------------------------------
# НАСТРОЙКИ. Учётные данные загружаются только из окружения или файла .env.
# ---------------------------------------------------------------------------

# Не храните OAuth-токен в исходном коде и не передавайте его аргументом CLI.
TOKEN = os.getenv("YANDEX_DIRECT_TOKEN", "").strip()

# Логин клиента Директа. Нужен для агентского аккаунта.
# Для собственного аккаунта можно оставить пустую строку.
CLIENT_LOGIN = os.getenv("YANDEX_DIRECT_CLIENT_LOGIN", "").strip()

# ID целей: YANDEX_DIRECT_GOAL_IDS=111,222,333 или --goals 111 222 333.
GOAL_IDS = [
    value.strip()
    for value in os.getenv("YANDEX_DIRECT_GOAL_IDS", "").split(",")
    if value.strip()
]

# Допустимые модели: AUTO, LC, FCCD, LSCCD.
ATTRIBUTION_MODEL = "AUTO"

# Оставьте пустыми для последних 30 завершенных дней или задайте YYYY-MM-DD.
DATE_FROM = os.getenv("YANDEX_DIRECT_DATE_FROM", "").strip()
DATE_TO = os.getenv("YANDEX_DIRECT_DATE_TO", "").strip()

# Пустое значение создает direct_report_ДАТА_ДАТА.csv в рабочей папке.
OUTPUT_FILE = os.getenv("YANDEX_DIRECT_OUTPUT_FILE", "").strip()

# CSV для русской версии Excel: UTF-8 с BOM, разделитель ;, десятичная запятая.
CSV_DELIMITER = ";"
CSV_DECIMAL_SEPARATOR = ","

# Если True, кроме суммарных метрик добавляются столбцы по каждой цели.
INCLUDE_GOAL_BREAKDOWN = False

# Заменять ID аудиторной корректировки на "Название [ID]".
RESOLVE_ADJUSTMENT_NAMES = True

# API по умолчанию обрезает отчет на 1 000 000 строк. Скрипт выгружает отчет
# страницами указанного размера, пока не получит все строки.
REPORT_PAGE_LIMIT = 250_000


REPORTS_URL = "https://api.direct.yandex.com/json/v501/reports"
RETARGETING_LISTS_URL = (
    "https://api.direct.yandex.com/json/v501/retargetinglists"
)

REPORT_FIELDS = [
    "CampaignName",
    "AdGroupName",
    "CriterionType",
    "Criterion",
    "RlAdjustmentId",
    "AutotargetingBrandOption",
    "MatchType",
    "TargetingCategory",
    "AdId",
    "AdFormat",
    "AdNetworkType",
    "Slot",
    "TargetingLocationName",
    "Device",
    "Gender",
    "Age",
    "IncomeGrade",
    "CampaignType",
    "Impressions",
    "Clicks",
    "Cost",
    "Conversions",
    "Revenue",
]

BASE_OUTPUT_FIELDS = [
    "Название кампании",
    "Название группы",
    "Тип условия показа",
    "Ключевая фраза",
    "Условие подбора аудитории",
    "Корректировки для целевой аудитории",
    "Упоминание брендов в запросах",
    "Тип соответствия",
    "Категория запроса",
    "№ объявления",
    "Формат",
    "Тип площадки",
    "Вид размещения",
    "Регион таргетинга",
    "Тип устройства",
    "Пол",
    "Возраст",
    "Уровень платёжеспособности",
    "Тип кампании",
    "Показы",
    "Клики",
    "Расход",
    "Конверсии",
    "Ценность конверсий",
]

AUDIENCE_CRITERION_TYPES = {
    "RETARGETING",
    "INTERESTS_AND_DEMOGRAPHICS",
    "MOBILE_APP_CATEGORY",
    "OFFER_RETARGETING",
}

EMPTY_API_VALUES = {"", "-", "--", "null", "None"}

ENUM_TRANSLATIONS: dict[str, dict[str, str]] = {
    "CriterionType": {
        "KEYWORD": "Ключевая фраза",
        "AUTOTARGETING": "Автотаргетинг",
        "RETARGETING": "Ретаргетинг и подбор аудитории",
        "INTERESTS_AND_DEMOGRAPHICS": "Профиль пользователей",
        "MOBILE_APP_CATEGORY": "Интерес к категории мобильных приложений",
        "WEBPAGE_FILTER": "Условие для динамического объявления",
        "FEED_FILTER": "Фильтр по фиду",
        "OFFER_RETARGETING": "Офферный ретаргетинг",
    },
    "AutotargetingBrandOption": {
        "WITHOUT_BRANDS": "Без упоминания брендов",
        "WITHOUT_BRAND": "Без упоминания брендов",
        "WITH_ADVERTISER_BRAND": "С упоминанием бренда рекламодателя",
        "WITH_COMPETITORS_BRAND": "С упоминанием брендов конкурентов",
        "WITH_COMPETITOR_BRAND": "С упоминанием брендов конкурентов",
    },
    "MatchType": {
        "RELATED_KEYWORD": "Дополнительная релевантная фраза",
        "SYNONYM": "Семантическое соответствие",
        "KEYWORD": "Ключевая фраза",
        "NONE": "Нет",
    },
    "TargetingCategory": {
        "EXACT": "Целевые/узкие запросы",
        "NARROW": "Узкие запросы",
        "ALTERNATIVE": "Альтернативные запросы",
        "COMPETITOR": "Запросы с упоминанием конкурентов",
        "BROADER": "Широкие запросы",
        "ACCESSORY": "Сопутствующие запросы",
    },
    "AdFormat": {
        "IMAGE": "Графический",
        "TEXT": "Текстовый",
        "VIDEO": "Видео",
        "SMART_MULTIPLE": "Смарт-баннер",
        "SMART_SINGLE": "Смарт-объявление",
        "ADAPTIVE_IMAGE": "Адаптивный графический",
        "SMART_TILE": "Смарт-плитка",
    },
    "AdNetworkType": {
        "SEARCH": "Поиск",
        "AD_NETWORK": "Сети",
    },
    "Slot": {
        "ALONE": "Эксклюзивное размещение",
        "PREMIUMBLOCK": "Спецразмещение",
        "SUGGEST": "Саджест",
        "PRODUCT_GALLERY": "Товарная галерея",
        "OTHER": "Другие блоки",
    },
    "Device": {
        "DESKTOP": "Компьютеры",
        "MOBILE": "Смартфоны",
        "TABLET": "Планшеты",
        "SMART_TV": "Smart TV",
    },
    "Gender": {
        "GENDER_MALE": "Мужской",
        "GENDER_FEMALE": "Женский",
        "UNKNOWN": "Не определён",
    },
    "Age": {
        "AGE_0_17": "0–17",
        "AGE_18_24": "18–24",
        "AGE_25_34": "25–34",
        "AGE_35_44": "35–44",
        "AGE_45": "45+",
        "AGE_45_54": "45–54",
        "AGE_55": "55+",
        "UNKNOWN": "Не определён",
    },
    "IncomeGrade": {
        "VERY_HIGH": "Топ 1%",
        "HIGH": "2–5%",
        "ABOVE_AVERAGE": "6–10%",
        "OTHER": "Остальные",
    },
    "CampaignType": {
        "TEXT_CAMPAIGN": "Текстово-графическая",
        "MOBILE_APP_CAMPAIGN": "Реклама мобильных приложений",
        "DYNAMIC_TEXT_CAMPAIGN": "Динамическая",
        "SMART_CAMPAIGN": "Смарт-кампания",
        "MCBANNER_CAMPAIGN": "Медийная",
        "CPM_BANNER_CAMPAIGN": "Медийная баннерная",
        "CPM_DEALS_CAMPAIGN": "Медийная сделка",
        "CPM_FRONTPAGE_CAMPAIGN": "Медийная на главной",
        "CPM_PRICE": "Медийная с оплатой за показы",
        "UNIFIED_CAMPAIGN": "Единая перфоманс-кампания",
    },
}


class DirectApiError(RuntimeError):
    """Ошибка запроса или ответа API Яндекс Директа."""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Выгрузить детализированный отчет Яндекс Директа в CSV."
    )
    parser.add_argument("--date-from", help="Начало периода, YYYY-MM-DD")
    parser.add_argument("--date-to", help="Конец периода, YYYY-MM-DD")
    parser.add_argument(
        "--goals",
        nargs="+",
        help="ID целей через пробел или запятую (не более 10)",
    )
    parser.add_argument(
        "--attribution",
        choices=("AUTO", "LC", "FCCD", "LSCCD"),
        default=ATTRIBUTION_MODEL,
        help="Модель атрибуции (по умолчанию AUTO)",
    )
    parser.add_argument("--output", type=Path, help="Путь к итоговому CSV")
    parser.add_argument(
        "--goal-breakdown",
        action="store_true",
        default=INCLUDE_GOAL_BREAKDOWN,
        help="Добавить отдельные столбцы метрик для каждой цели",
    )
    parser.add_argument(
        "--no-adjustment-names",
        action="store_true",
        help="Не запрашивать названия аудиторных корректировок",
    )
    return parser.parse_args()


def split_goal_ids(values: Sequence[str] | None) -> list[str]:
    raw_values: Iterable[str] = values if values is not None else GOAL_IDS
    goals = [part.strip() for value in raw_values for part in value.split(",")]
    goals = [goal for goal in goals if goal]

    if not goals:
        raise ValueError("Укажите хотя бы один ID цели.")
    if len(goals) > 10:
        raise ValueError("Reports API принимает не более 10 целей за запрос.")
    if len(set(goals)) != len(goals):
        raise ValueError("ID целей не должны повторяться.")
    if any(not goal.isdigit() for goal in goals):
        raise ValueError("ID целей должны состоять только из цифр.")
    return goals


def resolve_dates(date_from_arg: str | None, date_to_arg: str | None) -> tuple[str, str]:
    yesterday = date.today() - timedelta(days=1)
    default_from = yesterday - timedelta(days=29)
    raw_from = date_from_arg or DATE_FROM or default_from.isoformat()
    raw_to = date_to_arg or DATE_TO or yesterday.isoformat()

    try:
        parsed_from = date.fromisoformat(raw_from)
        parsed_to = date.fromisoformat(raw_to)
    except ValueError as exc:
        raise ValueError("Даты должны быть в формате YYYY-MM-DD.") from exc

    if parsed_from > parsed_to:
        raise ValueError("Начальная дата не может быть позже конечной.")
    if parsed_to >= date.today():
        print(
            "Предупреждение: данные за сегодня могут быть неполными.",
            file=sys.stderr,
        )
    return parsed_from.isoformat(), parsed_to.isoformat()


def resolve_output_path(
    output_arg: Path | None, date_from: str, date_to: str
) -> Path:
    if output_arg is not None:
        return output_arg.expanduser().resolve()
    if OUTPUT_FILE:
        return Path(OUTPUT_FILE).expanduser().resolve()
    return Path.cwd() / f"direct_report_{date_from}_{date_to}.csv"


def common_headers(token: str, client_login: str) -> dict[str, str]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept-Language": "ru",
        "Content-Type": "application/json; charset=utf-8",
        "User-Agent": "yandex-direct-csv-export/1.0",
    }
    if client_login:
        headers["Client-Login"] = client_login
    return headers


def build_report_body(
    date_from: str,
    date_to: str,
    goal_ids: Sequence[str],
    attribution_model: str,
    page_limit: int = REPORT_PAGE_LIMIT,
    page_offset: int = 0,
) -> dict:
    unique_suffix = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    return {
        "params": {
            "SelectionCriteria": {
                "DateFrom": date_from,
                "DateTo": date_to,
            },
            "Goals": list(goal_ids),
            "AttributionModels": [attribution_model],
            "FieldNames": REPORT_FIELDS,
            "Page": {
                "Limit": page_limit,
                "Offset": page_offset,
            },
            "ReportName": (
                f"direct-csv-{date_from}-{date_to}-"
                f"offset-{page_offset}-{unique_suffix}"
            ),
            "ReportType": "CUSTOM_REPORT",
            "DateRangeType": "CUSTOM_DATE",
            "Format": "TSV",
            "IncludeVAT": "YES",
            "IncludeDiscount": "NO",
        }
    }


def error_details(response: requests.Response) -> str:
    request_id = response.headers.get("RequestId", "не указан")
    try:
        payload = response.json()
    except ValueError:
        payload = response.text.strip()[:1000]
    return (
        f"HTTP {response.status_code}; RequestId: {request_id}; "
        f"ответ: {payload}"
    )


def download_report(
    token: str,
    client_login: str,
    body: Mapping,
    max_wait_seconds: int = 3600,
) -> str:
    headers = common_headers(token, client_login)
    headers.update(
        {
            "processingMode": "auto",
            "returnMoneyInMicros": "false",
            "skipReportHeader": "true",
            "skipReportSummary": "true",
            "Accept-Encoding": "gzip",
        }
    )
    started_at = time.monotonic()
    server_error_attempts = 0

    while True:
        try:
            response = requests.post(
                REPORTS_URL,
                headers=headers,
                json=body,
                timeout=(20, 300),
            )
        except requests.RequestException as exc:
            raise DirectApiError(f"Не удалось обратиться к Reports API: {exc}") from exc

        if response.status_code == 200:
            response.encoding = "utf-8"
            return response.text

        if response.status_code in (201, 202):
            elapsed = time.monotonic() - started_at
            if elapsed >= max_wait_seconds:
                raise DirectApiError(
                    "Отчет не сформировался за отведенный час. Запустите скрипт "
                    "повторно или уменьшите период."
                )
            retry_in = max(1, int(response.headers.get("retryIn", "60")))
            status = (
                "поставлен в очередь"
                if response.status_code == 201
                else "ещё формируется"
            )
            print(f"Отчет {status}; повтор через {retry_in} сек.")
            time.sleep(min(retry_in, max_wait_seconds - int(elapsed)))
            continue

        if response.status_code in (500, 502) and server_error_attempts < 2:
            server_error_attempts += 1
            delay = 5 * server_error_attempts
            print(
                f"Временная ошибка сервера HTTP {response.status_code}; "
                f"повтор через {delay} сек.",
                file=sys.stderr,
            )
            time.sleep(delay)
            continue

        raise DirectApiError(f"Reports API вернул ошибку: {error_details(response)}")


def iter_report_rows(report_text: str) -> Iterator[dict[str, str]]:
    reader = csv.DictReader(io.StringIO(report_text), delimiter="\t")
    if not reader.fieldnames:
        raise DirectApiError("Reports API вернул пустой отчет без заголовка столбцов.")
    for row in reader:
        yield {key: (value or "") for key, value in row.items() if key is not None}


def collect_adjustment_ids(report_text: str) -> set[str]:
    return {
        row["RlAdjustmentId"].strip()
        for row in iter_report_rows(report_text)
        if row.get("RlAdjustmentId", "").strip() not in EMPTY_API_VALUES
    }


def chunks(values: Sequence[str], size: int) -> Iterator[list[str]]:
    for start in range(0, len(values), size):
        yield list(values[start : start + size])


def load_adjustment_names(
    token: str, client_login: str, adjustment_ids: set[str]
) -> dict[str, str]:
    if not adjustment_ids:
        return {}

    numeric_ids = sorted(
        (value for value in adjustment_ids if value.isdigit()), key=int
    )
    names: dict[str, str] = {}
    headers = common_headers(token, client_login)

    try:
        for id_chunk in chunks(numeric_ids, 10_000):
            body = {
                "method": "get",
                "params": {
                    "SelectionCriteria": {"Ids": [int(value) for value in id_chunk]},
                    "FieldNames": ["Id", "Name"],
                },
            }
            response = requests.post(
                RETARGETING_LISTS_URL,
                headers=headers,
                json=body,
                timeout=(20, 120),
            )
            if response.status_code != 200:
                raise DirectApiError(error_details(response))
            payload = response.json()
            if "error" in payload:
                raise DirectApiError(str(payload["error"]))
            for item in payload.get("result", {}).get("RetargetingLists", []):
                names[str(item["Id"])] = item.get("Name", "")
    except (requests.RequestException, ValueError, KeyError, DirectApiError) as exc:
        print(
            "Предупреждение: не удалось получить названия аудиторных "
            f"корректировок ({exc}). В CSV останутся их ID.",
            file=sys.stderr,
        )
    return names


def translate(field_name: str, value: str) -> str:
    value = value.strip()
    if value in EMPTY_API_VALUES:
        return ""
    return ENUM_TRANSLATIONS.get(field_name, {}).get(value, value)


def parse_decimal(value: str | None) -> Decimal | None:
    if value is None:
        return None
    cleaned = value.strip().replace("\xa0", "").replace(" ", "")
    if cleaned in EMPTY_API_VALUES:
        return None
    # Reports API обычно использует точку, но обработаем и десятичную запятую.
    if "," in cleaned and "." not in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def format_decimal(value: Decimal | None) -> str:
    if value is None:
        return ""
    result = format(value, "f")
    if "." in result:
        result = result.rstrip("0").rstrip(".")
    if CSV_DECIMAL_SEPARATOR != ".":
        result = result.replace(".", CSV_DECIMAL_SEPARATOR)
    return result


def normalize_numeric(value: str) -> str:
    parsed = parse_decimal(value)
    return format_decimal(parsed) if parsed is not None else translate("", value)


def goal_metric(
    row: Mapping[str, str], metric: str, goal_id: str, attribution_model: str
) -> Decimal | None:
    return parse_decimal(row.get(f"{metric}_{goal_id}_{attribution_model}"))


def sum_goal_metric(
    row: Mapping[str, str],
    metric: str,
    goal_ids: Sequence[str],
    attribution_model: str,
) -> Decimal | None:
    values = [
        goal_metric(row, metric, goal_id, attribution_model)
        for goal_id in goal_ids
    ]
    present_values = [value for value in values if value is not None]
    return sum(present_values, Decimal(0)) if present_values else None


def output_fieldnames(
    goal_ids: Sequence[str], include_goal_breakdown: bool
) -> list[str]:
    fields = BASE_OUTPUT_FIELDS.copy()
    if include_goal_breakdown:
        for goal_id in goal_ids:
            fields.extend(
                [
                    f"Конверсии (цель {goal_id})",
                    f"Ценность конверсий (цель {goal_id})",
                ]
            )
    return fields


def build_output_row(
    row: Mapping[str, str],
    goal_ids: Sequence[str],
    attribution_model: str,
    adjustment_names: Mapping[str, str],
    include_goal_breakdown: bool,
) -> dict[str, str]:
    criterion_type = row.get("CriterionType", "").strip()
    criterion = translate("", row.get("Criterion", ""))
    adjustment_id = row.get("RlAdjustmentId", "").strip()

    if adjustment_id in EMPTY_API_VALUES:
        adjustment = ""
    elif adjustment_id in adjustment_names and adjustment_names[adjustment_id]:
        adjustment = f"{adjustment_names[adjustment_id]} [{adjustment_id}]"
    else:
        adjustment = adjustment_id

    output_row = {
        "Название кампании": translate("", row.get("CampaignName", "")),
        "Название группы": translate("", row.get("AdGroupName", "")),
        "Тип условия показа": translate("CriterionType", criterion_type),
        "Ключевая фраза": criterion if criterion_type == "KEYWORD" else "",
        "Условие подбора аудитории": (
            criterion if criterion_type in AUDIENCE_CRITERION_TYPES else ""
        ),
        "Корректировки для целевой аудитории": adjustment,
        "Упоминание брендов в запросах": translate(
            "AutotargetingBrandOption",
            row.get("AutotargetingBrandOption", ""),
        ),
        "Тип соответствия": translate("MatchType", row.get("MatchType", "")),
        "Категория запроса": translate(
            "TargetingCategory", row.get("TargetingCategory", "")
        ),
        "№ объявления": translate("", row.get("AdId", "")),
        "Формат": translate("AdFormat", row.get("AdFormat", "")),
        "Тип площадки": translate(
            "AdNetworkType", row.get("AdNetworkType", "")
        ),
        "Вид размещения": translate("Slot", row.get("Slot", "")),
        "Регион таргетинга": translate(
            "", row.get("TargetingLocationName", "")
        ),
        "Тип устройства": translate("Device", row.get("Device", "")),
        "Пол": translate("Gender", row.get("Gender", "")),
        "Возраст": translate("Age", row.get("Age", "")),
        "Уровень платёжеспособности": translate(
            "IncomeGrade", row.get("IncomeGrade", "")
        ),
        "Тип кампании": translate("CampaignType", row.get("CampaignType", "")),
        "Показы": normalize_numeric(row.get("Impressions", "")),
        "Клики": normalize_numeric(row.get("Clicks", "")),
        "Расход": normalize_numeric(row.get("Cost", "")),
        "Конверсии": format_decimal(
            sum_goal_metric(row, "Conversions", goal_ids, attribution_model)
        ),
        "Ценность конверсий": format_decimal(
            sum_goal_metric(row, "Revenue", goal_ids, attribution_model)
        ),
    }

    if include_goal_breakdown:
        for goal_id in goal_ids:
            output_row[f"Конверсии (цель {goal_id})"] = format_decimal(
                goal_metric(row, "Conversions", goal_id, attribution_model)
            )
            output_row[f"Ценность конверсий (цель {goal_id})"] = format_decimal(
                goal_metric(row, "Revenue", goal_id, attribution_model)
            )
    return output_row


def write_csv(
    report_text: str,
    output_path: Path,
    goal_ids: Sequence[str],
    attribution_model: str,
    adjustment_names: Mapping[str, str],
    include_goal_breakdown: bool,
    append: bool = False,
) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fields = output_fieldnames(goal_ids, include_goal_breakdown)
    row_count = 0

    mode = "a" if append else "w"
    # BOM нужен только при создании файла. При дописывании используем обычный UTF-8.
    encoding = "utf-8" if append else "utf-8-sig"
    with output_path.open(mode, encoding=encoding, newline="") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=fields,
            delimiter=CSV_DELIMITER,
            lineterminator="\n",
            quoting=csv.QUOTE_MINIMAL,
        )
        if not append:
            writer.writeheader()
        for row in iter_report_rows(report_text):
            writer.writerow(
                build_output_row(
                    row,
                    goal_ids,
                    attribution_model,
                    adjustment_names,
                    include_goal_breakdown,
                )
            )
            row_count += 1
    return row_count


def main() -> int:
    args = parse_args()
    token = TOKEN
    if not token:
        print(
            "Ошибка: задайте YANDEX_DIRECT_TOKEN в файле .env или окружении.",
            file=sys.stderr,
        )
        return 2

    try:
        goal_ids = split_goal_ids(args.goals)
        date_from, date_to = resolve_dates(args.date_from, args.date_to)
        output_path = resolve_output_path(args.output, date_from, date_to)
        print(
            f"Запрашиваю отчет за {date_from} — {date_to}; "
            f"цели: {', '.join(goal_ids)}."
        )
        temporary_output = output_path.with_suffix(output_path.suffix + ".part")
        total_rows = 0
        page_offset = 0
        page_number = 1

        while True:
            print(
                f"Страница {page_number}: строки с {page_offset + 1}, "
                f"лимит {REPORT_PAGE_LIMIT}."
            )
            body = build_report_body(
                date_from,
                date_to,
                goal_ids,
                args.attribution,
                page_limit=REPORT_PAGE_LIMIT,
                page_offset=page_offset,
            )
            report_text = download_report(token, CLIENT_LOGIN, body)

            adjustment_names: dict[str, str] = {}
            if RESOLVE_ADJUSTMENT_NAMES and not args.no_adjustment_names:
                adjustment_ids = collect_adjustment_ids(report_text)
                adjustment_names = load_adjustment_names(
                    token, CLIENT_LOGIN, adjustment_ids
                )

            page_rows = write_csv(
                report_text,
                temporary_output,
                goal_ids,
                args.attribution,
                adjustment_names,
                args.goal_breakdown,
                append=page_number > 1,
            )
            total_rows += page_rows
            print(
                f"Страница {page_number} записана: {page_rows} строк; "
                f"всего {total_rows}."
            )

            if page_rows < REPORT_PAGE_LIMIT:
                break
            page_offset += page_rows
            page_number += 1

        temporary_output.replace(output_path)
        row_count = total_rows
    except (ValueError, OSError, DirectApiError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    print(f"Готово: {row_count} строк записано в {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
