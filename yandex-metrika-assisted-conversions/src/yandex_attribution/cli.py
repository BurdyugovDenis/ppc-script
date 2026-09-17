from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
from datetime import date, timedelta
from pathlib import Path

from .api import DirectClient, MetrikaClient, date_chunks
from .config import Settings, get_token, load_dotenv
from .pipeline import build


DIRECT_REPORT_CACHE_VERSION = 3
METRIKA_LOG_CACHE_VERSION = 2


def parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Expected YYYY-MM-DD") from exc


def run_id(
    date_from: date,
    date_to: date,
    counter_ids: tuple[int, ...],
    goal_ids: tuple[int, ...],
) -> str:
    counters = "_".join(str(counter_id) for counter_id in counter_ids)
    goals = "_".join(str(goal_id) for goal_id in goal_ids)
    return (
        f"{date_from.isoformat()}_{date_to.isoformat()}_"
        f"counters_{counters}_goals_{goals}"
    )


def output_filename(
    date_from: date,
    date_to: date,
    client_login: str,
) -> str:
    """Return a filesystem-safe report filename with period and client login."""
    safe_login = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", client_login.strip())
    safe_login = re.sub(r"\s+", "_", safe_login)
    safe_login = re.sub(r"_+", "_", safe_login).strip(" ._") or "client"
    return f"{date_from.isoformat()}_{date_to.isoformat()}_{safe_login}.xlsx"


def read_run_metadata(path: Path) -> dict[str, object]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def direct_cache_matches(
    metadata: dict[str, object], settings: Settings
) -> bool:
    direct = metadata.get("direct")
    if not isinstance(direct, dict):
        return False
    return (
        direct.get("report_cache_version") == DIRECT_REPORT_CACHE_VERSION
        and direct.get("client_login") == settings.direct_client_login
        and direct.get("attribution_model") == settings.direct_conversion_model
        and direct.get("include_vat") == settings.include_vat
    )


def metrika_cache_matches(
    metadata: dict[str, object], settings: Settings
) -> bool:
    metrika = metadata.get("metrika")
    if not isinstance(metrika, dict):
        return False
    return (
        metrika.get("log_cache_version") == METRIKA_LOG_CACHE_VERSION
        and metrika.get("attribution_model")
        == settings.metrika_attribution_model
        and metrika.get("lookback_days") == settings.lookback_days
        and (
            not settings.include_associated_revenue
            or bool(metrika.get("include_associated_revenue", False))
        )
    )


def validate_dates(date_from: date, date_to: date) -> None:
    if date_from > date_to:
        raise ValueError("date-from must not be after date-to")
    if date_to >= date.today():
        raise ValueError("Metrika Logs API does not provide the current day; use yesterday or earlier")


def extract(args: argparse.Namespace) -> Path:
    settings = Settings.from_env()
    goal_ids = settings.require_goal_ids()
    validate_dates(args.date_from, args.date_to)
    metrika_token = get_token("YANDEX_METRIKA_TOKEN")
    direct_token = get_token("YANDEX_DIRECT_TOKEN")

    rid = run_id(
        args.date_from, args.date_to, settings.metrika_counter_ids, goal_ids
    )
    raw_dir = Path(args.data_dir) / "raw" / rid
    if raw_dir.exists() and args.force:
        shutil.rmtree(raw_dir)
    metadata_path = raw_dir / "run.json"
    existing_metadata = read_run_metadata(metadata_path)
    reuse_direct_cache = direct_cache_matches(existing_metadata, settings)
    reuse_metrika_cache = metrika_cache_matches(existing_metadata, settings)
    direct_dir = raw_dir / "direct"
    direct_dir.mkdir(parents=True, exist_ok=True)

    extract_from = args.date_from - timedelta(days=settings.lookback_days)
    direct = DirectClient(direct_token, settings.direct_client_login)
    for chunk_from, chunk_to in date_chunks(
        args.date_from, args.date_to, settings.direct_chunk_days
    ):
        path = direct_dir / f"{chunk_from.isoformat()}_{chunk_to.isoformat()}.tsv"
        if not path.exists() or args.force or not reuse_direct_cache:
            vat_label = "с НДС" if settings.include_vat else "без НДС"
            print(
                f"Direct ({vat_label}): {chunk_from}..{chunk_to}",
                flush=True,
            )
            direct.download_report(path, chunk_from, chunk_to, settings)

    metrika = MetrikaClient(metrika_token)
    metrika_root = raw_dir / "metrika"
    metrika_root.mkdir(parents=True, exist_ok=True)
    for counter_id in settings.metrika_counter_ids:
        metrika_dir = metrika_root / str(counter_id)
        metrika_dir.mkdir(parents=True, exist_ok=True)
        existing_parts = sorted(metrika_dir.glob("part_*.tsv"))
        if existing_parts and not args.force and reuse_metrika_cache:
            continue
        print(
            f"Metrika counter {counter_id} "
            f"({settings.metrika_attribution_model}): "
            f"{extract_from}..{args.date_to}",
            flush=True,
        )
        request_id = metrika.create_log_request(
            counter_id, extract_from, args.date_to, settings
        )
        log_request = metrika.wait_for_log(
            counter_id, request_id, settings
        )
        staging_dir = metrika_root / f".{counter_id}.download"
        if staging_dir.exists():
            shutil.rmtree(staging_dir)
        downloaded_parts = metrika.download_log_parts(
            counter_id, request_id, log_request, staging_dir
        )
        if not downloaded_parts:
            raise RuntimeError(
                f"Metrika returned no log parts for counter {counter_id}"
            )
        if metrika_dir.exists():
            shutil.rmtree(metrika_dir)
        staging_dir.replace(metrika_dir)
        if settings.clean_metrika_log_after_download:
            metrika.clean_log(counter_id, request_id)

    metadata = {
        "date_from": args.date_from.isoformat(),
        "date_to": args.date_to.isoformat(),
        "extract_from": extract_from.isoformat(),
        "metrika_counter_ids": list(settings.metrika_counter_ids),
        "goal_ids": list(goal_ids),
        "lookback_days": settings.lookback_days,
        "direct": {
            "report_cache_version": DIRECT_REPORT_CACHE_VERSION,
            "client_login": settings.direct_client_login,
            "attribution_model": settings.direct_conversion_model,
            "include_vat": settings.include_vat,
        },
        "metrika": {
            "log_cache_version": METRIKA_LOG_CACHE_VERSION,
            "attribution_model": settings.metrika_attribution_model,
            "lookback_days": settings.lookback_days,
            "include_associated_revenue": settings.include_associated_revenue,
        },
    }
    metadata_path.write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return raw_dir


def build_run(args: argparse.Namespace, raw_dir: Path | None = None) -> Path:
    settings = Settings.from_env()
    goal_ids = settings.require_goal_ids()
    validate_dates(args.date_from, args.date_to)
    rid = run_id(
        args.date_from, args.date_to, settings.metrika_counter_ids, goal_ids
    )
    raw_dir = raw_dir or Path(args.data_dir) / "raw" / rid
    metadata = read_run_metadata(raw_dir / "run.json")
    if not direct_cache_matches(metadata, settings):
        vat_label = "с НДС" if settings.include_vat else "без НДС"
        raise ValueError(
            "Кэш Директа не соответствует текущим настройкам "
            f"({vat_label}, модель {settings.direct_conversion_model}). "
            "Установите RUN_MODE=run: скрипт перевыгрузит только "
            "статистику Директа и сохранит уже скачанные логи Метрики."
        )
    if not metrika_cache_matches(metadata, settings):
        revenue_label = (
            "с ассоциированным доходом"
            if settings.include_associated_revenue
            else "без ассоциированного дохода"
        )
        raise ValueError(
            "Кэш Метрики не соответствует текущей модели атрибуции, окну "
            "ассиста или настройке дохода "
            f"({settings.metrika_attribution_model}, "
            f"{settings.lookback_days} дней, {revenue_label}). "
            "Установите RUN_MODE=run: "
            "скрипт перевыгрузит только логи Метрики и сохранит статистику "
            "Директа."
        )
    direct_paths = sorted((raw_dir / "direct").glob("*.tsv"))
    metrika_paths = sorted((raw_dir / "metrika").glob("*/part_*.tsv"))
    if not metrika_paths:
        # Compatibility with data downloaded by the previous single-counter version.
        metrika_paths = sorted((raw_dir / "metrika").glob("part_*.tsv"))
    if not direct_paths or not metrika_paths:
        raise FileNotFoundError(
            f"Raw Direct or Metrika files are missing under {raw_dir}. Run extract first."
        )
    output_path = Path(args.data_dir) / "output" / output_filename(
        args.date_from,
        args.date_to,
        settings.direct_client_login,
    )
    database_path = Path(args.data_dir) / f"{rid}.sqlite3"
    qa = build(
        direct_paths,
        metrika_paths,
        database_path,
        output_path,
        settings,
        args.date_from,
        args.date_to,
    )
    totals = qa["combined_totals"]
    print(f"Лиды Метрики: {totals['unique_goal_visits']:,}".replace(",", " "))
    print(f"Ассоциированные конверсии Директа: {totals['assisted_metrika_conversions_unique']:,}".replace(",", " "))
    print(f"Листов в XLSX: {len(qa['excel_workbook']['sheets'])}")
    return output_path


def list_goals(args: argparse.Namespace) -> None:
    settings = Settings.from_env()
    metrika = MetrikaClient(get_token("YANDEX_METRIKA_TOKEN"))
    print("counter_id\tgoal_id\tname\ttype\tstatus")
    for counter_id in settings.metrika_counter_ids:
        goals = metrika.list_goals(counter_id)
        for goal in goals:
            status = goal.get("status", "")
            print(
                f"{counter_id}\t{goal.get('id')}\t{goal.get('name')}\t"
                f"{goal.get('type')}\t{status}"
            )


def make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="yandex-attribution",
        description="Join Yandex Direct spend with Metrika assisted conversions.",
    )
    parser.add_argument("--data-dir", default="data")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("goals", help="List goals from all configured counters")

    for name in ("extract", "build", "run"):
        child = subparsers.add_parser(name)
        child.add_argument("--date-from", required=True, type=parse_date)
        child.add_argument("--date-to", required=True, type=parse_date)
        if name in {"extract", "run"}:
            child.add_argument("--force", action="store_true")
    return parser


def env_arguments() -> list[str]:
    """Use the same entry point for PyCharm and terminal commands."""
    load_dotenv()
    mode = os.environ.get("RUN_MODE", "run").strip().lower()
    if mode not in {"run", "build", "extract", "goals"}:
        raise ValueError("RUN_MODE must be run, build, extract or goals")
    args = ["--data-dir", str(Path(os.environ.get("DATA_DIR", "data").strip()).expanduser()), mode]
    if mode != "goals":
        for key, flag in (("DATE_FROM", "--date-from"), ("DATE_TO", "--date-to")):
            value = os.environ.get(key, "").strip()
            if not value:
                raise ValueError(f"Fill {key} in the .env file")
            args.extend([flag, value])
    return args


def main(argv: list[str] | None = None) -> None:
    load_dotenv()
    parser = make_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "goals":
            list_goals(args)
        elif args.command == "extract":
            raw_dir = extract(args)
            print(f"Raw data: {raw_dir}")
        elif args.command == "build":
            output_path = build_run(args)
            print(f"Output: {output_path}")
        elif args.command == "run":
            raw_dir = extract(args)
            output_path = build_run(args, raw_dir)
            print(f"Output: {output_path}")
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
