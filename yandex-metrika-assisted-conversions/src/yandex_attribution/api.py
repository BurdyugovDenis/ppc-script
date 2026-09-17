from __future__ import annotations

import hashlib
import json
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Iterator

import requests

from .config import Settings


DIRECT_REPORT_URL = "https://api.direct.yandex.com/json/v5/reports"
METRIKA_BASE_URL = "https://api-metrika.yandex.net/management/v1"

LOG_FIELDS = [
    "ym:s:visitID",
    "ym:s:clientID",
    "ym:s:counterUserIDHash",
    "ym:s:date",
    "ym:s:dateTime",
    "ym:s:goalsID",
    "ym:s:<attribution>TrafficSource",
    "ym:s:<attribution>AdvEngine",
    "ym:s:<attribution>DirectClickOrder",
    "ym:s:<attribution>DirectBannerGroup",
    "ym:s:<attribution>DirectClickOrderName",
    "ym:s:<attribution>ClickBannerGroupName",
    "ym:s:<attribution>DirectPhraseOrCond",
    "ym:s:<attribution>DirectPlatformType",
    "ym:s:<attribution>DirectPlatform",
    "ym:s:<attribution>DirectConditionType",
    "ym:s:<attribution>UTMSource",
    "ym:s:<attribution>UTMMedium",
    "ym:s:<attribution>UTMCampaign",
    "ym:s:<attribution>UTMContent",
    "ym:s:<attribution>UTMTerm",
]

ASSOCIATED_REVENUE_LOG_FIELDS = [
    "ym:s:goalsPrice",
    "ym:s:goalsCurrency",
]


def metrika_log_fields(settings: Settings) -> list[str]:
    """Return Logs API fields required by the selected report options."""
    fields = list(LOG_FIELDS)
    if settings.include_associated_revenue:
        fields.extend(ASSOCIATED_REVENUE_LOG_FIELDS)
    return fields

DIRECT_FIELDS = [
    "Date",
    "CampaignId",
    "CampaignName",
    "AdGroupId",
    "AdGroupName",
    "CriterionId",
    "Criterion",
    "CriterionType",
    "AdNetworkType",
    "Placement",
    "Impressions",
    "Clicks",
    "Cost",
    "Revenue",
    "Conversions",
]


class APIError(RuntimeError):
    pass


def _raise_for_api(response: requests.Response) -> None:
    if response.ok:
        return
    try:
        body = response.content.decode("utf-8")[:2000]
    except UnicodeDecodeError:
        body = response.text[:2000]
    raise APIError(f"HTTP {response.status_code} from {response.url}: {body}")


def _direct_report_name(params: dict[str, object], date_from: date, date_to: date) -> str:
    """Return a stable name that changes whenever report parameters change."""
    signature = json.dumps(
        params,
        ensure_ascii=True,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    digest = hashlib.sha256(signature).hexdigest()[:12]
    return f"associated_{date_from.isoformat()}_{date_to.isoformat()}_{digest}"


def date_chunks(date_from: date, date_to: date, days: int) -> Iterator[tuple[date, date]]:
    cursor = date_from
    while cursor <= date_to:
        chunk_to = min(cursor + timedelta(days=days - 1), date_to)
        yield cursor, chunk_to
        cursor = chunk_to + timedelta(days=1)


class DirectClient:
    def __init__(self, token: str, client_login: str, timeout: int = 180) -> None:
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Client-Login": client_login,
                "Accept-Language": "ru",
                "processingMode": "auto",
                "returnMoneyInMicros": "false",
                "skipReportHeader": "true",
                "skipReportSummary": "true",
            }
        )

    def download_report(
        self,
        output_path: Path,
        date_from: date,
        date_to: date,
        settings: Settings,
    ) -> None:
        goal_ids = settings.require_goal_ids()
        params: dict[str, object] = {
            "SelectionCriteria": {
                "DateFrom": date_from.isoformat(),
                "DateTo": date_to.isoformat(),
            },
            "Goals": [str(goal_id) for goal_id in goal_ids],
            "AttributionModels": [settings.direct_conversion_model],
            "FieldNames": DIRECT_FIELDS,
            "ReportType": "CUSTOM_REPORT",
            "DateRangeType": "CUSTOM_DATE",
            "Format": "TSV",
            "IncludeVAT": "YES" if settings.include_vat else "NO",
            "IncludeDiscount": "NO",
        }
        params["ReportName"] = _direct_report_name(params, date_from, date_to)
        payload = {"params": params}

        deadline = time.monotonic() + settings.max_wait_minutes * 60
        while True:
            response = self.session.post(
                DIRECT_REPORT_URL, json=payload, timeout=self.timeout
            )
            if response.status_code == 200:
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_bytes(response.content)
                return
            if response.status_code in {201, 202} and time.monotonic() < deadline:
                retry_after = int(response.headers.get("retryIn", settings.poll_seconds))
                time.sleep(max(1, min(retry_after, 60)))
                continue
            _raise_for_api(response)
            raise APIError("Direct report did not become ready before the deadline")


class MetrikaClient:
    def __init__(self, token: str, timeout: int = 180) -> None:
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({"Authorization": f"OAuth {token}"})

    def list_goals(self, counter_id: int) -> list[dict]:
        response = self.session.get(
            f"{METRIKA_BASE_URL}/counter/{counter_id}/goals", timeout=self.timeout
        )
        _raise_for_api(response)
        return response.json().get("goals", [])

    def create_log_request(
        self,
        counter_id: int,
        date_from: date,
        date_to: date,
        settings: Settings,
    ) -> int:
        response = self.session.post(
            f"{METRIKA_BASE_URL}/counter/{counter_id}/logrequests",
            params={
                "date1": date_from.isoformat(),
                "date2": date_to.isoformat(),
                "fields": ",".join(metrika_log_fields(settings)),
                "source": "visits",
                "attribution": settings.metrika_attribution_model,
            },
            timeout=self.timeout,
        )
        _raise_for_api(response)
        return int(response.json()["log_request"]["request_id"])

    def wait_for_log(
        self, counter_id: int, request_id: int, settings: Settings
    ) -> dict:
        deadline = time.monotonic() + settings.max_wait_minutes * 60
        while time.monotonic() < deadline:
            response = self.session.get(
                f"{METRIKA_BASE_URL}/counter/{counter_id}/logrequest/{request_id}",
                timeout=self.timeout,
            )
            _raise_for_api(response)
            log_request = response.json()["log_request"]
            status = log_request["status"]
            if status == "processed":
                return log_request
            if status in {
                "canceled",
                "processing_failed",
                "cleaned_by_user",
                "cleaned_automatically_as_too_old",
            }:
                raise APIError(f"Metrika log request {request_id} ended with {status}")
            time.sleep(max(1, min(settings.poll_seconds, 60)))
        raise APIError(f"Metrika log request {request_id} timed out")

    def download_log_parts(
        self, counter_id: int, request_id: int, log_request: dict, output_dir: Path
    ) -> list[Path]:
        output_dir.mkdir(parents=True, exist_ok=True)
        paths: list[Path] = []
        for part in log_request.get("parts", []):
            number = int(part["part_number"])
            response = self.session.get(
                (
                    f"{METRIKA_BASE_URL}/counter/{counter_id}/logrequest/"
                    f"{request_id}/part/{number}/download"
                ),
                timeout=self.timeout,
            )
            _raise_for_api(response)
            path = output_dir / f"part_{number:05d}.tsv"
            path.write_bytes(response.content)
            paths.append(path)
        return paths

    def clean_log(self, counter_id: int, request_id: int) -> None:
        response = self.session.post(
            (
                f"{METRIKA_BASE_URL}/counter/{counter_id}/logrequest/"
                f"{request_id}/clean"
            ),
            timeout=self.timeout,
        )
        _raise_for_api(response)
