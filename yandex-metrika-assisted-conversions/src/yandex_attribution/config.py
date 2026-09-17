from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping


@dataclass(frozen=True)
class Settings:
    metrika_counter_ids: tuple[int, ...]
    direct_client_login: str
    goal_ids: tuple[int, ...] = ()
    lookback_days: int = 90
    direct_conversion_model: str = "LC"
    # UA Multi-Channel Funnels used the actual source of every interaction in
    # the path. LAST is the closest Metrika Logs API equivalent: it does not
    # inherit Yandex Direct from a previous visit as AUTOMATIC can do.
    metrika_attribution_model: str = "LAST"
    direct_chunk_days: int = 31
    poll_seconds: int = 10
    max_wait_minutes: int = 60
    clean_metrika_log_after_download: bool = True
    include_vat: bool = True
    include_associated_revenue: bool = False

    @classmethod
    def from_env(cls, environ: Mapping[str, str] | None = None) -> "Settings":
        """Read every user-editable setting from one .env-backed mapping."""
        values = os.environ if environ is None else environ

        def value(name: str, default: str = "") -> str:
            return values.get(name, default).strip()

        counter_text = value("YANDEX_METRIKA_COUNTER_IDS") or value(
            "YANDEX_METRIKA_COUNTER_ID"
        )
        if not counter_text:
            raise ValueError(
                "YANDEX_METRIKA_COUNTER_IDS is empty. Fill it in the .env file."
            )

        settings = cls(
            metrika_counter_ids=parse_integer_ids(
                "YANDEX_METRIKA_COUNTER_IDS", counter_text
            ),
            direct_client_login=value("YANDEX_DIRECT_CLIENT_LOGIN"),
            goal_ids=parse_goal_ids(value("YANDEX_GOAL_IDS")),
            lookback_days=int(value("LOOKBACK_DAYS", "90")),
            direct_conversion_model=value("DIRECT_CONVERSION_MODEL", "LC").upper(),
            metrika_attribution_model=normalize_metrika_attribution_model(
                value("METRIKA_ATTRIBUTION_MODEL", "LAST")
            ),
            direct_chunk_days=int(value("DIRECT_CHUNK_DAYS", "31")),
            poll_seconds=int(value("POLL_SECONDS", "10")),
            max_wait_minutes=int(value("MAX_WAIT_MINUTES", "60")),
            clean_metrika_log_after_download=parse_bool(
                "CLEAN_METRIKA_LOG_AFTER_DOWNLOAD",
                value("CLEAN_METRIKA_LOG_AFTER_DOWNLOAD", "true"),
            ),
            include_vat=parse_bool("INCLUDE_VAT", value("INCLUDE_VAT", "true")),
            include_associated_revenue=parse_bool(
                "INCLUDE_ASSOCIATED_REVENUE",
                value("INCLUDE_ASSOCIATED_REVENUE", "false"),
            ),
        )
        settings.validate()
        return settings

    def validate(self) -> None:
        if not self.metrika_counter_ids:
            raise ValueError("YANDEX_METRIKA_COUNTER_IDS must not be empty")
        if any(counter_id <= 0 for counter_id in self.metrika_counter_ids):
            raise ValueError(
                "All YANDEX_METRIKA_COUNTER_IDS values must be positive integers"
            )
        if not self.direct_client_login.strip():
            raise ValueError(
                "YANDEX_DIRECT_CLIENT_LOGIN is empty. Fill it in the .env file."
            )
        if any(goal_id <= 0 for goal_id in self.goal_ids):
            raise ValueError("All YANDEX_GOAL_IDS values must be positive integers")
        if len(self.goal_ids) > 10:
            raise ValueError("YANDEX_GOAL_IDS supports at most 10 goals per run")
        if not 1 <= self.lookback_days <= 365:
            raise ValueError("LOOKBACK_DAYS must be between 1 and 365")
        if self.direct_conversion_model not in {"LC", "LSCCD", "FCCD", "AUTO"}:
            raise ValueError(
                "DIRECT_CONVERSION_MODEL must be LC, LSCCD, FCCD, or AUTO"
            )
        if self.metrika_attribution_model not in {
            "FIRST",
            "LAST",
            "LASTSIGN",
            "LAST_YANDEX_DIRECT_CLICK",
            "CROSS_DEVICE_LAST_SIGNIFICANT",
            "CROSS_DEVICE_FIRST",
            "CROSS_DEVICE_LAST_YANDEX_DIRECT_CLICK",
            "CROSS_DEVICE_LAST",
            "AUTOMATIC",
        }:
            raise ValueError(
                "METRIKA_ATTRIBUTION_MODEL must be AUTO/AUTOMATIC or another "
                "model supported by Metrika Logs API"
            )
        if not 1 <= self.direct_chunk_days <= 366:
            raise ValueError("DIRECT_CHUNK_DAYS must be between 1 and 366")
        if self.poll_seconds <= 0:
            raise ValueError("POLL_SECONDS must be positive")
        if self.max_wait_minutes <= 0:
            raise ValueError("MAX_WAIT_MINUTES must be positive")

    def require_goal_ids(self) -> tuple[int, ...]:
        if not self.goal_ids:
            raise ValueError(
                "YANDEX_GOAL_IDS is empty. Add goal IDs to .env separated by commas."
            )
        return self.goal_ids


def load_dotenv(path: Path = Path(".env")) -> None:
    """Load a small KEY=VALUE .env file without adding another dependency."""
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def get_token(name: str) -> str:
    token = os.environ.get(name, "").strip()
    if not token:
        raise RuntimeError(
            f"Environment variable {name} is empty. Fill .env or export it first."
        )
    return token


def parse_goal_ids(value: str) -> tuple[int, ...]:
    return parse_integer_ids("YANDEX_GOAL_IDS", value)


def normalize_metrika_attribution_model(value: str) -> str:
    """Return the exact uppercase value expected by Metrika Logs API."""
    normalized = value.strip().upper()
    return "AUTOMATIC" if normalized == "AUTO" else normalized


def parse_integer_ids(name: str, value: str) -> tuple[int, ...]:
    if not value.strip():
        return ()
    result: list[int] = []
    for item in value.split(","):
        text = item.strip()
        if not text:
            continue
        try:
            goal_id = int(text)
        except ValueError as exc:
            raise ValueError(f"Invalid integer {text!r} in {name}") from exc
        if goal_id not in result:
            result.append(goal_id)
    return tuple(result)


def parse_bool(name: str, value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"{name} must be true or false")
