from __future__ import annotations

import tempfile
import unittest
from dataclasses import replace
from datetime import date
from pathlib import Path

import requests

from yandex_attribution.api import (
    APIError,
    LOG_FIELDS,
    DirectClient,
    MetrikaClient,
    _raise_for_api,
    metrika_log_fields,
)
from yandex_attribution.config import Settings


class _Response:
    status_code = 200
    content = b"Date\tCost\n"


class _Session:
    def __init__(self) -> None:
        self.payload: dict | None = None

    def post(self, _url: str, *, json: dict, timeout: int) -> _Response:
        self.payload = json
        return _Response()


class _MetrikaResponse:
    ok = True

    def json(self) -> dict:
        return {"log_request": {"request_id": 42}}


class _MetrikaSession:
    def __init__(self) -> None:
        self.params: dict | None = None

    def post(self, _url: str, *, params: dict, timeout: int) -> _MetrikaResponse:
        self.params = params
        return _MetrikaResponse()


class DirectClientTest(unittest.TestCase):
    def test_report_requests_cost_with_vat_by_default(self) -> None:
        settings = Settings(
            metrika_counter_ids=(1,),
            direct_client_login="client",
            goal_ids=(99,),
        )
        client = DirectClient("token", "client")
        session = _Session()
        client.session = session  # type: ignore[assignment]

        with tempfile.TemporaryDirectory() as tmp:
            client.download_report(
                Path(tmp) / "direct.tsv",
                date(2026, 8, 1),
                date(2026, 8, 31),
                settings,
            )

        assert session.payload is not None
        self.assertEqual(session.payload["params"]["IncludeVAT"], "YES")
        self.assertIn("Revenue", session.payload["params"]["FieldNames"])
        report_name_with_vat = session.payload["params"]["ReportName"]
        self.assertRegex(
            report_name_with_vat,
            r"^associated_2026-08-01_2026-08-31_[0-9a-f]{12}$",
        )

        settings_without_vat = replace(settings, include_vat=False)
        with tempfile.TemporaryDirectory() as tmp:
            client.download_report(
                Path(tmp) / "direct.tsv",
                date(2026, 8, 1),
                date(2026, 8, 31),
                settings_without_vat,
            )
        assert session.payload is not None
        self.assertNotEqual(
            session.payload["params"]["ReportName"],
            report_name_with_vat,
        )

    def test_api_error_is_decoded_as_utf8(self) -> None:
        response = requests.Response()
        response.status_code = 400
        response.url = "https://api.direct.yandex.com/json/v5/reports"
        response.encoding = "ISO-8859-1"
        response._content = (
            '{"error":{"error_detail":"Измените ReportName"}}'
        ).encode("utf-8")

        with self.assertRaises(APIError) as context:
            _raise_for_api(response)
        self.assertIn("Измените ReportName", str(context.exception))


class MetrikaClientTest(unittest.TestCase):
    def test_log_request_uses_automatic_attribution(self) -> None:
        settings = Settings(
            metrika_counter_ids=(1,),
            direct_client_login="client",
            goal_ids=(99,),
            metrika_attribution_model="AUTOMATIC",
        )
        client = MetrikaClient("token")
        session = _MetrikaSession()
        client.session = session  # type: ignore[assignment]

        request_id = client.create_log_request(
            1,
            date(2026, 8, 1),
            date(2026, 8, 31),
            settings,
        )

        self.assertEqual(request_id, 42)
        assert session.params is not None
        self.assertEqual(session.params["attribution"], "AUTOMATIC")
        self.assertEqual(session.params["fields"], ",".join(LOG_FIELDS))
        self.assertIn("ym:s:<attribution>TrafficSource", LOG_FIELDS)

        self.assertNotIn("ym:s:goalsPrice", session.params["fields"])
        self.assertNotIn("ym:s:goalsCurrency", session.params["fields"])

        revenue_settings = replace(settings, include_associated_revenue=True)
        client.create_log_request(
            1,
            date(2026, 8, 1),
            date(2026, 8, 31),
            revenue_settings,
        )
        assert session.params is not None
        self.assertEqual(
            session.params["fields"],
            ",".join(metrika_log_fields(revenue_settings)),
        )
        self.assertIn("ym:s:goalsPrice", session.params["fields"])
        self.assertIn("ym:s:goalsCurrency", session.params["fields"])


if __name__ == "__main__":
    unittest.main()
