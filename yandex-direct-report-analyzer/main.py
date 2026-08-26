#!/usr/bin/env python3
"""Единая точка запуска для выгрузки и анализа данных Яндекс Директа в PyCharm."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent

# Настройки запуска из PyCharm. Их можно переопределить аргументами командной строки.
DOWNLOAD_BEFORE_ANALYSIS = False
# Если INPUT_FILE равен None, скрипт автоматически выберет самый новый CSV из папки data/.
# Здесь можно указать конкретный файл или передать его через аргумент --input.
INPUT_FILE: Path | None = None
OUTPUT_FILE = PROJECT_DIR / "output" / "direct_report_analysis.xlsx"

# Оставьте PERIOD равным None, чтобы определить период из имени файла,
# например direct_report_2025-10-10_2026-08-18.csv.
PERIOD: str | None = None

# Указывайте лист только для XLSX, если нужные данные находятся не на первом листе.
SOURCE_SHEET: str | None = None

# Переключатели финансовых метрик. Безопасный режим по умолчанию считает только CR и CPL.
# Яндекс Директ может возвращать ценность лидовой конверсии, которая не является реальной
# выручкой бизнеса. Включайте выручку только тогда, когда исходный столбец имеет реальный бизнес-смысл.
CALCULATE_REVENUE = False
CALCULATE_DRR = True
CALCULATE_ROMI = True
CALCULATE_AOV = False

# Включайте только тогда, когда каждая выбранная конверсия равна одному заказу.
# В остальных случаях AOV должен использовать отдельный столбец с подтверждёнными заказами.
AOV_FROM_CONVERSIONS = False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Выгрузить отчёт Яндекс Директа и/или собрать аналитический Excel."
        )
    )
    parser.add_argument("--input", type=Path, help="Готовый CSV/XLSX для анализа")
    parser.add_argument("--output", type=Path, help="Путь к итоговому XLSX")
    parser.add_argument(
        "--download",
        action="store_true",
        default=DOWNLOAD_BEFORE_ANALYSIS,
        help="Сначала выгрузить CSV через Reports API",
    )
    parser.add_argument(
        "--export-output",
        type=Path,
        help="Путь к CSV, создаваемому при --download",
    )
    parser.add_argument("--date-from", help="Начало выгрузки, YYYY-MM-DD")
    parser.add_argument("--date-to", help="Конец выгрузки, YYYY-MM-DD")
    parser.add_argument("--goals", nargs="+", help="ID целей через пробел")
    parser.add_argument(
        "--attribution",
        choices=("AUTO", "LC", "FCCD", "LSCCD"),
        default="AUTO",
        help="Модель атрибуции",
    )
    parser.add_argument("--goal-breakdown", action="store_true")
    parser.add_argument("--no-adjustment-names", action="store_true")
    parser.add_argument("--sheet", help="Лист исходного XLSX")
    parser.add_argument("--period", help="Подпись периода в итоговом отчёте")
    parser.add_argument(
        "--revenue",
        action=argparse.BooleanOptionalAction,
        default=CALCULATE_REVENUE,
        help="Включить/выключить выручку (доход) из выгрузки",
    )
    parser.add_argument(
        "--drr",
        action=argparse.BooleanOptionalAction,
        default=CALCULATE_DRR,
        help="Включить/выключить ДРР (требует выручку)",
    )
    parser.add_argument(
        "--romi",
        action=argparse.BooleanOptionalAction,
        default=CALCULATE_ROMI,
        help="Включить/выключить ROMI (требует выручку)",
    )
    parser.add_argument(
        "--aov",
        action=argparse.BooleanOptionalAction,
        default=CALCULATE_AOV,
        help="Включить/выключить AOV (требует выручку и заказы)",
    )
    parser.add_argument(
        "--aov-from-conversions",
        action=argparse.BooleanOptionalAction,
        default=AOV_FROM_CONVERSIONS,
        help="Считать каждую конверсию заказом для AOV",
    )
    return parser.parse_args()


def inferred_period(path: Path) -> str:
    match = re.search(r"(20\d{2}-\d{2}-\d{2})[_—-]+(20\d{2}-\d{2}-\d{2})", path.stem)
    if match:
        return f"{match.group(1)} — {match.group(2)} (по имени файла)"
    return "Не указан в источнике"


def run(command: list[str]) -> None:
    print("\n$", " ".join(command))
    subprocess.run(command, check=True)


def find_csv_in_data(project_dir: Path) -> Path:
    """Найти самый новый непустой CSV, находящийся непосредственно в папке data/."""
    data_dir = project_dir / "data"
    if not data_dir.is_dir():
        raise FileNotFoundError(f"Не найдена папка с данными: {data_dir}")
    candidates = [
        path
        for path in data_dir.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".csv"
        and not path.name.startswith(".")
        and not path.name.lower().endswith(".part.csv")
        and path.stat().st_size > 0
    ]
    if not candidates:
        raise FileNotFoundError(
            f"В {data_dir} нет доступных CSV. "
            "Положите CSV в data/ или передайте --input."
        )
    return max(candidates, key=lambda path: (path.stat().st_mtime_ns, path.name))


def configure_metrics(
    analysis: dict,
    *,
    include_revenue: bool,
    include_drr: bool,
    include_romi: bool,
    include_aov: bool,
    allow_conversion_proxy: bool,
) -> None:
    metrics = analysis.setdefault("metrics_available", {})
    source_metrics = dict(metrics)
    analysis["source_metrics_available"] = source_metrics
    source_revenue = bool(source_metrics.get("revenue"))
    source_orders = bool(source_metrics.get("orders"))

    revenue_available = bool(include_revenue and source_revenue)
    aov_available = bool(
        revenue_available
        and include_aov
        and (source_orders or allow_conversion_proxy)
    )
    aov_from_conversions = bool(aov_available and not source_orders and allow_conversion_proxy)
    metrics.update(
        {
            "revenue": revenue_available,
            "drr": bool(revenue_available and include_drr),
            "romi": bool(revenue_available and include_romi),
            "orders": bool(aov_available and source_orders),
            "aov": aov_available,
        }
    )
    analysis["aov_from_conversions"] = aov_from_conversions
    if aov_available and source_orders:
        analysis["aov_basis"] = "Выручка / подтверждённые заказы из исходной выгрузки"
    elif aov_from_conversions:
        analysis["aov_basis"] = (
            "Прокси: ценность конверсий / конверсии; "
            "каждая конверсия считается заказом для расчёта AOV"
        )
    else:
        analysis["aov_basis"] = "Отключён" if not include_aov else "Недоступен"

    disabled: list[str] = []
    if not include_revenue:
        disabled.extend(["Выручка", "ДРР", "ROMI", "AOV"])
    else:
        if not include_drr:
            disabled.append("ДРР")
        if not include_romi:
            disabled.append("ROMI")
        if not include_aov:
            disabled.append("AOV")
    unavailable: list[str] = []
    if include_revenue and not source_revenue:
        unavailable.extend(["Выручка", "ДРР", "ROMI"])
        if include_aov:
            unavailable.append("AOV")
    elif include_aov and not aov_available:
        unavailable.append("AOV")
    analysis["disabled_metrics"] = list(dict.fromkeys(disabled))
    analysis["unavailable_metrics"] = list(dict.fromkeys(unavailable))


def export_report(args: argparse.Namespace, project_dir: Path) -> tuple[Path, str]:
    from yandex_direct_report import resolve_dates

    date_from, date_to = resolve_dates(args.date_from, args.date_to)
    export_output = (
        args.export_output.expanduser().resolve()
        if args.export_output
        else (project_dir / "data" / f"direct_report_{date_from}_{date_to}.csv")
    )
    export_output.parent.mkdir(parents=True, exist_ok=True)
    command = [
        sys.executable,
        str(project_dir / "yandex_direct_report.py"),
        "--date-from",
        date_from,
        "--date-to",
        date_to,
        "--attribution",
        args.attribution,
        "--output",
        str(export_output),
    ]
    if args.goals:
        command.extend(["--goals", *args.goals])
    if args.goal_breakdown:
        command.append("--goal-breakdown")
    if args.no_adjustment_names:
        command.append("--no-adjustment-names")
    run(command)
    if not export_output.is_file():
        raise FileNotFoundError(f"Экспортёр не создал CSV: {export_output}")
    return export_output, f"{date_from} — {date_to}"


def main() -> None:
    args = parse_args()
    project_dir = PROJECT_DIR
    exported_period = None
    if args.download:
        source, exported_period = export_report(args, project_dir)
    elif args.input:
        source = args.input.expanduser().resolve()
    elif INPUT_FILE is not None:
        source = INPUT_FILE.expanduser().resolve()
    else:
        source = find_csv_in_data(project_dir).resolve()
        print(f"Автовыбор CSV из data/: {source}")
    output = (args.output or OUTPUT_FILE).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(
            f"Не найден исходный файл: {source}. "
            "Передайте --input или запустите с --download."
        )
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="yandex-direct-analysis-") as temporary:
        analysis_path = Path(temporary) / "analysis.json"
        aggregate_command = [
            sys.executable,
            str(project_dir / "aggregate_export.py"),
            "--input",
            str(source),
            "--output-json",
            str(analysis_path),
        ]
        source_sheet = args.sheet or SOURCE_SHEET
        if source_sheet:
            aggregate_command.extend(["--sheet", source_sheet])
        run(aggregate_command)

        analysis = json.loads(analysis_path.read_text(encoding="utf-8"))
        analysis["period"] = (
            args.period or PERIOD or exported_period or inferred_period(source)
        )
        configure_metrics(
            analysis,
            include_revenue=args.revenue,
            include_drr=args.drr,
            include_romi=args.romi,
            include_aov=args.aov,
            allow_conversion_proxy=args.aov_from_conversions,
        )
        enabled = [
            label
            for label, key in (
                ("Выручка", "revenue"),
                ("ДРР", "drr"),
                ("ROMI", "romi"),
                ("AOV", "aov"),
            )
            if analysis["metrics_available"].get(key)
        ]
        print(f"\nФинансовые метрики: {', '.join(enabled) or 'отключены'}")
        if args.aov:
            print(f"AOV: {analysis.get('aov_basis') or 'недоступен'}")
        analysis_path.write_text(
            json.dumps(analysis, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )

        run(
            [
                sys.executable,
                str(project_dir / "build_report.py"),
                "--analysis",
                str(analysis_path),
                "--output",
                str(output),
            ]
        )
        if not output.is_file():
            raise FileNotFoundError(f"Генератор не создал итоговый Excel: {output}")
        run(
            [
                sys.executable,
                str(project_dir / "validate_report.py"),
                "--analysis",
                str(analysis_path),
                "--workbook",
                str(output),
            ]
        )

    print(f"\nГотово: {output}")


if __name__ == "__main__":
    main()
