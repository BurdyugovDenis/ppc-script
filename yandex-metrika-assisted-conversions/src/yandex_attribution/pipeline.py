from __future__ import annotations

import csv
import re
import sqlite3
from collections.abc import Iterable
from datetime import date, timedelta
from pathlib import Path

from .config import Settings
from .excel import export_excel_workbook


NULLS = {"", "--", "-", "\\N", "None", "null"}

YANDEX_DIRECT_ADV_ENGINES = {
    "ya_direct",
    "ya_undefined",
    "yandex direct",
    "yandex.direct",
    "yandex direct: undefined",
    "яндекс директ",
    "яндекс.директ",
    "яндекс: директ",
    "яндекс.директ: не определено",
}

def clean(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return "" if text in NULLS else text


def norm_text(value: object) -> str:
    return " ".join(clean(value).casefold().split())


def norm_id(value: object) -> str:
    text = clean(value)
    if not text or text == "0":
        return ""
    try:
        return str(int(float(text)))
    except ValueError:
        return text


def norm_network(value: object) -> str:
    raw = norm_text(value)
    if not raw:
        return ""
    if "search" in raw or "поиск" in raw:
        return "SEARCH"
    if any(x in raw for x in ("network", "context", "реклам", "рся", "content")):
        return "AD_NETWORK"
    return clean(value).upper()


def classify_traffic_source(
    traffic_source: object,
    adv_engine: object,
    campaign_id: str,
) -> str:
    """Normalize both identified and undefined Yandex Direct visits."""
    source = norm_text(traffic_source)
    engine = norm_text(adv_engine)
    if campaign_id or (
        source == "ad" and engine in YANDEX_DIRECT_ADV_ENGINES
    ):
        return "yandex_direct"
    return source or "unknown"


def association_window(
    conversion_date_from: date,
    conversion_date_to: date,
    lookback_days: int,
) -> tuple[date, date]:
    """Return the raw-data envelope needed by per-conversion lookback paths."""
    return (
        conversion_date_from - timedelta(days=lookback_days),
        conversion_date_to,
    )


def number(value: object) -> float:
    text = clean(value).replace(" ", "").replace(",", ".")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_goal_ids(value: object) -> list[int]:
    return [int(item) for item in re.findall(r"\d+", clean(value))]


def parse_number_array(value: object) -> list[float]:
    """Parse a numeric Logs API array without depending on JSON quoting style."""
    pattern = r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?"
    return [float(item) for item in re.findall(pattern, clean(value))]


def parse_text_array(value: object) -> list[str]:
    """Parse simple quoted string arrays returned by Metrika Logs API."""
    text = clean(value).strip()
    if not text or text == "[]":
        return []
    inner = text[1:-1] if text.startswith("[") and text.endswith("]") else text
    return [item.strip().strip("'\"") for item in inner.split(",")]


def pick(row: dict[str, str], *names: str) -> str:
    for name in names:
        if name in row:
            return clean(row[name])
    return ""


def metrika_field_names(settings: Settings, suffix: str) -> tuple[str, str]:
    """Return possible TSV headers for an attribution-parameterized field."""
    expanded = settings.metrika_attribution_model.lower()
    return (
        f"ym:s:{expanded}{suffix}",
        f"ym:s:<attribution>{suffix}",
    )


def connect(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    return connection


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        DROP TABLE IF EXISTS visits;
        DROP TABLE IF EXISTS goal_visits;
        DROP TABLE IF EXISTS direct_stats;
        DROP TABLE IF EXISTS conversion_touches;

        CREATE TABLE visits (
            counter_id INTEGER NOT NULL,
            visit_id TEXT NOT NULL,
            user_id TEXT NOT NULL,
            visit_datetime TEXT NOT NULL,
            event_date TEXT NOT NULL,
            traffic_source TEXT NOT NULL,
            adv_engine TEXT NOT NULL,
            campaign_id TEXT NOT NULL,
            campaign_name TEXT NOT NULL,
            ad_group_id TEXT NOT NULL,
            ad_group_name TEXT NOT NULL,
            criterion TEXT NOT NULL,
            criterion_key TEXT NOT NULL,
            criterion_type TEXT NOT NULL,
            network_type TEXT NOT NULL,
            placement TEXT NOT NULL,
            placement_key TEXT NOT NULL,
            utm_source TEXT NOT NULL,
            utm_medium TEXT NOT NULL,
            utm_campaign TEXT NOT NULL,
            utm_content TEXT NOT NULL,
            utm_term TEXT NOT NULL,
            PRIMARY KEY (counter_id, visit_id)
        );

        CREATE TABLE goal_visits (
            counter_id INTEGER NOT NULL,
            visit_id TEXT NOT NULL,
            user_id TEXT NOT NULL,
            visit_datetime TEXT NOT NULL,
            conversion_date TEXT NOT NULL,
            goal_id INTEGER NOT NULL,
            goal_reaches INTEGER NOT NULL,
            goal_revenue REAL NOT NULL,
            goal_currency TEXT NOT NULL,
            PRIMARY KEY (counter_id, visit_id, goal_id)
        );

        CREATE TABLE direct_stats (
            event_date TEXT NOT NULL,
            campaign_id TEXT NOT NULL,
            campaign_name TEXT NOT NULL,
            ad_group_id TEXT NOT NULL,
            ad_group_name TEXT NOT NULL,
            criterion_id TEXT NOT NULL,
            criterion TEXT NOT NULL,
            criterion_key TEXT NOT NULL,
            criterion_type TEXT NOT NULL,
            network_type TEXT NOT NULL,
            placement TEXT NOT NULL,
            placement_key TEXT NOT NULL,
            impressions REAL NOT NULL,
            clicks REAL NOT NULL,
            cost REAL NOT NULL,
            direct_api_converted_visits REAL NOT NULL,
            direct_api_revenue REAL NOT NULL
        );
        """
    )


def load_metrika_visits(
    connection: sqlite3.Connection,
    paths: Iterable[Path],
    settings: Settings,
) -> tuple[int, int]:
    visit_count = 0
    goal_visit_count = 0
    visit_sql = """
        INSERT OR REPLACE INTO visits VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
    """
    goal_sql = (
        "INSERT OR REPLACE INTO goal_visits "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
    )
    goal_ids = settings.require_goal_ids()

    def attributed(row: dict[str, str], *suffixes: str) -> str:
        names = [name for suffix in suffixes for name in metrika_field_names(settings, suffix)]
        return pick(row, *names)

    for path in paths:
        parent_name = path.parent.name
        if parent_name.isdigit():
            counter_id = int(parent_name)
            if counter_id not in settings.metrika_counter_ids:
                raise ValueError(
                    f"Metrika file {path} belongs to unconfigured counter {counter_id}"
                )
        elif len(settings.metrika_counter_ids) == 1:
            # Compatibility with the previous flat raw/metrika directory.
            counter_id = settings.metrika_counter_ids[0]
        else:
            raise ValueError(
                f"Cannot determine Metrika counter for {path}; expected a numeric "
                "parent directory"
            )
        with path.open("r", encoding="utf-8", newline="") as fh:
            reader = csv.DictReader(fh, delimiter="\t")
            if reader.fieldnames is None:
                continue
            required_attribution_fields = ("TrafficSource", "DirectClickOrder")
            missing_fields = [
                suffix
                for suffix in required_attribution_fields
                if not any(
                    name in reader.fieldnames
                    for name in metrika_field_names(settings, suffix)
                )
            ]
            if missing_fields:
                raise ValueError(
                    f"Metrika file {path} does not contain fields for model "
                    f"{settings.metrika_attribution_model}. Run with "
                    "RUN_MODE=run to download fresh Logs API data."
                )
            if settings.include_associated_revenue:
                missing_revenue_fields = [
                    name
                    for name in ("ym:s:goalsPrice", "ym:s:goalsCurrency")
                    if name not in reader.fieldnames
                ]
                if missing_revenue_fields:
                    raise ValueError(
                        f"Metrika file {path} does not contain "
                        f"{', '.join(missing_revenue_fields)}. Set RUN_MODE=run "
                        "to download fresh Logs API data with goal revenue."
                    )
            for row in reader:
                visit_id = pick(row, "ym:s:visitID")
                user_id = pick(row, "ym:s:counterUserIDHash") or pick(
                    row, "ym:s:clientID"
                )
                visit_datetime = pick(row, "ym:s:dateTime")
                event_date = pick(row, "ym:s:date") or visit_datetime[:10]
                if not visit_id or not user_id or not visit_datetime:
                    continue

                campaign_id = norm_id(attributed(row, "DirectClickOrder"))
                adv_engine = attributed(row, "AdvEngine")
                traffic_source = classify_traffic_source(
                    attributed(row, "TrafficSource"), adv_engine, campaign_id,
                )
                criterion = attributed(row, "DirectPhraseOrCond")
                placement = attributed(row, "DirectPlatform")
                values = (
                    counter_id,
                    visit_id,
                    user_id,
                    visit_datetime,
                    event_date,
                    traffic_source,
                    adv_engine,
                    campaign_id,
                    attributed(row, "DirectClickOrderName"),
                    norm_id(attributed(row, "DirectBannerGroup")),
                    attributed(row, "ClickBannerGroupName", "DirectBannerGroupName"),
                    criterion,
                    norm_text(criterion),
                    attributed(row, "DirectConditionType"),
                    norm_network(attributed(row, "DirectPlatformType")),
                    placement,
                    norm_text(placement),
                    *(attributed(row, f"UTM{field}") for field in (
                        "Source", "Medium", "Campaign", "Content", "Term"
                    )),
                )
                connection.execute(visit_sql, values)
                visit_count += 1

                row_goal_ids = parse_goal_ids(row.get("ym:s:goalsID"))
                row_goal_prices = (
                    parse_number_array(row.get("ym:s:goalsPrice"))
                    if settings.include_associated_revenue
                    else []
                )
                row_goal_currencies = (
                    parse_text_array(row.get("ym:s:goalsCurrency"))
                    if settings.include_associated_revenue
                    else []
                )
                if settings.include_associated_revenue and row_goal_ids:
                    if len(row_goal_prices) != len(row_goal_ids):
                        raise ValueError(
                            f"Metrika visit {counter_id}:{visit_id} has "
                            "different lengths for goalsID and goalsPrice."
                        )
                    if row_goal_currencies and len(row_goal_currencies) != len(
                        row_goal_ids
                    ):
                        raise ValueError(
                            f"Metrika visit {counter_id}:{visit_id} has "
                            "different lengths for goalsID and goalsCurrency."
                        )
                for goal_id in goal_ids:
                    selected_indexes = [
                        index
                        for index, item in enumerate(row_goal_ids)
                        if item == goal_id
                    ]
                    if not selected_indexes:
                        continue
                    goal_revenue = (
                        sum(
                            row_goal_prices[index] / 1000.0
                            for index in selected_indexes
                        )
                        if settings.include_associated_revenue
                        else 0.0
                    )
                    currencies = {
                        row_goal_currencies[index]
                        for index in selected_indexes
                        if index < len(row_goal_currencies)
                        and row_goal_currencies[index]
                    }
                    if len(currencies) == 1:
                        goal_currency = next(iter(currencies))
                    elif len(currencies) > 1:
                        goal_currency = "MIXED"
                    else:
                        goal_currency = ""
                    connection.execute(
                        goal_sql,
                        (
                            counter_id,
                            visit_id,
                            user_id,
                            visit_datetime,
                            event_date,
                            goal_id,
                            len(selected_indexes),
                            goal_revenue,
                            goal_currency,
                        ),
                    )
                    goal_visit_count += 1
        connection.commit()
    if settings.include_associated_revenue:
        currencies = {
            row[0]
            for row in connection.execute(
                """
                SELECT DISTINCT goal_currency
                FROM goal_visits
                WHERE goal_currency <> ''
                """
            )
        }
        if "MIXED" in currencies or len(currencies) > 1:
            raise ValueError(
                "Selected Metrika goals use more than one revenue currency "
                f"({', '.join(sorted(currencies))}). The script cannot sum "
                "different currencies; align goal currencies or set "
                "INCLUDE_ASSOCIATED_REVENUE=false."
            )
    return visit_count, goal_visit_count


def load_direct_stats(
    connection: sqlite3.Connection, paths: Iterable[Path], settings: Settings
) -> int:
    count = 0
    goal_ids = settings.require_goal_ids()
    sql = "INSERT INTO direct_stats VALUES (" + ",".join("?" for _ in range(17)) + ")"
    for path in paths:
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh, delimiter="\t")
            fields = reader.fieldnames or []
            for row in reader:
                criterion = clean(row.get("Criterion"))
                placement = clean(row.get("Placement"))
                base_values = (
                        clean(row.get("Date")),
                        norm_id(row.get("CampaignId")),
                        clean(row.get("CampaignName")),
                        norm_id(row.get("AdGroupId")),
                        clean(row.get("AdGroupName")),
                        norm_id(row.get("CriterionId")),
                        criterion,
                        norm_text(criterion),
                        clean(row.get("CriterionType")),
                        norm_network(row.get("AdNetworkType")),
                        placement,
                        norm_text(placement),
                        number(row.get("Impressions")),
                        number(row.get("Clicks")),
                        number(row.get("Cost")),
                )
                total_direct_conversions = 0.0
                total_direct_revenue = 0.0
                for goal_id in goal_ids:
                    conversion_field = (
                        f"Conversions_{goal_id}_{settings.direct_conversion_model}"
                    )
                    if conversion_field not in fields:
                        candidates = [
                            name
                            for name in fields
                            if name.startswith(f"Conversions_{goal_id}_")
                        ]
                        if len(candidates) == 1:
                            conversion_field = candidates[0]
                    total_direct_conversions += number(row.get(conversion_field))
                    revenue_field = (
                        f"Revenue_{goal_id}_{settings.direct_conversion_model}"
                    )
                    if revenue_field not in fields:
                        candidates = [
                            name
                            for name in fields
                            if name.startswith(f"Revenue_{goal_id}_")
                        ]
                        if len(candidates) == 1:
                            revenue_field = candidates[0]
                    total_direct_revenue += number(row.get(revenue_field))
                connection.execute(
                    sql,
                    (
                        *base_values,
                        total_direct_conversions,
                        total_direct_revenue,
                    ),
                )
                count += 1
        connection.commit()
    return count


def build_conversion_touches(
    connection: sqlite3.Connection,
    settings: Settings,
    conversion_date_from: date,
    conversion_date_to: date,
) -> int:
    connection.execute("DROP TABLE IF EXISTS conversion_touches")
    connection.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_visits_user_dt
            ON visits(counter_id, user_id, visit_datetime);
        CREATE INDEX IF NOT EXISTS idx_goal_visits_date
            ON goal_visits(counter_id, goal_id, conversion_date);
        """
    )
    connection.execute(
        """
        CREATE TABLE conversion_touches AS
        WITH conversion_basket AS (
            SELECT
                counter_id,
                visit_id,
                MAX(user_id) AS user_id,
                MAX(visit_datetime) AS visit_datetime,
                MAX(conversion_date) AS conversion_date,
                SUM(goal_reaches) AS goal_reaches,
                SUM(goal_revenue) AS goal_revenue
            FROM goal_visits
            WHERE conversion_date BETWEEN :date_from AND :date_to
            GROUP BY counter_id, visit_id
        ), eligible_conversions AS (
            SELECT c.*
            FROM conversion_basket c
            JOIN visits closing
              ON closing.counter_id = c.counter_id
             AND closing.visit_id = c.visit_id
            WHERE closing.traffic_source <> 'yandex_direct'
        ), ranked AS (
            SELECT
                c.counter_id,
                c.visit_id AS conversion_visit_id,
                c.goal_reaches,
                c.goal_revenue,
                c.conversion_date,
                t.visit_id AS touch_visit_id,
                t.visit_datetime AS touch_datetime,
                t.event_date,
                t.traffic_source,
                t.adv_engine,
                t.campaign_id,
                t.campaign_name,
                t.ad_group_id,
                t.ad_group_name,
                t.criterion,
                t.criterion_key,
                t.criterion_type,
                t.network_type,
                t.placement,
                t.placement_key,
                t.utm_source,
                t.utm_medium,
                t.utm_campaign,
                t.utm_content,
                t.utm_term,
                ROW_NUMBER() OVER (
                    PARTITION BY
                        c.counter_id, c.visit_id, t.traffic_source, t.adv_engine,
                        t.campaign_id, t.ad_group_id, t.criterion_key,
                        t.network_type, t.placement_key, t.utm_source,
                        t.utm_medium, t.utm_campaign, t.utm_content, t.utm_term
                    ORDER BY t.visit_datetime DESC, t.visit_id DESC
                ) AS touch_rank
            FROM eligible_conversions c
            JOIN visits t
              ON t.counter_id = c.counter_id
             AND t.user_id = c.user_id
             AND t.traffic_source = 'yandex_direct'
             AND t.visit_datetime < c.visit_datetime
             AND t.visit_datetime >= datetime(
                    c.visit_datetime,
                    '-' || :lookback_days || ' days'
                 )
        ), deduplicated AS (
            SELECT * FROM ranked WHERE touch_rank = 1
        )
        SELECT
            deduplicated.*,
            1.0 / COUNT(*) OVER (
                PARTITION BY counter_id, conversion_visit_id
            ) AS linear_conversion_credit,
            1.0 * goal_reaches / COUNT(*) OVER (
                PARTITION BY counter_id, conversion_visit_id
            ) AS linear_goal_reach_credit,
            1.0 * goal_revenue / COUNT(*) OVER (
                PARTITION BY counter_id, conversion_visit_id
            ) AS linear_revenue_credit
        FROM deduplicated
        """,
        {
            "date_from": conversion_date_from.isoformat(),
            "date_to": conversion_date_to.isoformat(),
            "lookback_days": settings.lookback_days,
        },
    )
    connection.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_conversion_touches_direct
            ON conversion_touches(
                counter_id, event_date, campaign_id, ad_group_id, criterion_key,
                network_type, placement_key
            );
        CREATE INDEX IF NOT EXISTS idx_conversion_touches_conversion
            ON conversion_touches(
                conversion_date, counter_id, conversion_visit_id
            );
        """
    )
    connection.commit()
    return int(connection.execute("SELECT COUNT(*) FROM conversion_touches").fetchone()[0])


def build_conversion_paths(
    connection: sqlite3.Connection,
    settings: Settings,
    conversion_date_from: date,
    conversion_date_to: date,
) -> int:
    """Build channel paths and channel roles from the same goal-visit cohort."""
    revenue_select = (
        ", SUM(goal_revenue) AS goal_revenue"
        if settings.include_associated_revenue
        else ""
    )
    params = {
        "date_from": conversion_date_from.isoformat(),
        "date_to": conversion_date_to.isoformat(),
        "counter_ids": ",".join(map(str, settings.metrika_counter_ids)),
        "goal_ids": ",".join(map(str, settings.goal_ids)),
        "metrika_model": settings.metrika_attribution_model,
        "lookback_days": settings.lookback_days,
    }
    # Materialize the shared history once. Roles use individual visits; the
    # readable paths below still collapse consecutive occurrences of a channel.
    connection.execute("DROP TABLE IF EXISTS temp.conversion_path_steps")
    connection.execute(
        """
        CREATE TEMP TABLE conversion_path_steps AS
        WITH conversion_basket AS (
            SELECT
                counter_id,
                visit_id,
                MAX(user_id) AS user_id,
                MAX(visit_datetime) AS visit_datetime,
                MAX(conversion_date) AS conversion_date,
                SUM(goal_reaches) AS goal_reaches,
                SUM(goal_revenue) AS goal_revenue
            FROM goal_visits
            WHERE conversion_date BETWEEN :date_from AND :date_to
            GROUP BY counter_id, visit_id
        ), path_steps AS (
            SELECT
                c.counter_id,
                c.visit_id AS conversion_visit_id,
                c.conversion_date,
                c.goal_reaches,
                c.goal_revenue,
                v.visit_id AS step_visit_id,
                v.visit_datetime,
                v.traffic_source,
                CASE v.traffic_source
                    WHEN 'yandex_direct' THEN 'Яндекс Директ'
                    WHEN 'organic' THEN 'Поисковые системы'
                    WHEN 'direct' THEN 'Прямые заходы'
                    WHEN 'referral' THEN 'Ссылки на сайтах'
                    WHEN 'social' THEN 'Социальные сети'
                    WHEN 'messenger' THEN 'Мессенджеры'
                    WHEN 'internal' THEN 'Внутренние переходы'
                    WHEN 'recommendation' THEN 'Рекомендательные системы'
                    WHEN 'email' THEN 'Email'
                    WHEN 'saved' THEN 'Сохранённые страницы'
                    WHEN 'ad' THEN 'Другая реклама'
                    WHEN 'unknown' THEN 'Не определено'
                    ELSE v.traffic_source
                END AS channel
            FROM conversion_basket c
            JOIN visits v
              ON v.counter_id = c.counter_id
             AND v.user_id = c.user_id
             AND (
                    v.visit_datetime < c.visit_datetime
                    OR v.visit_id = c.visit_id
                 )
             AND v.visit_datetime >= datetime(
                    c.visit_datetime,
                    '-' || :lookback_days || ' days'
                 )
        )
        SELECT *,
            ROW_NUMBER() OVER (
                PARTITION BY counter_id, conversion_visit_id
                ORDER BY visit_datetime, step_visit_id
            ) AS visit_position,
            COUNT(*) OVER (
                PARTITION BY counter_id, conversion_visit_id
            ) AS visits_in_path
        FROM path_steps
        """,
        params,
    )
    build_channel_roles(connection)
    connection.execute("DROP TABLE IF EXISTS report_conversion_paths")
    connection.execute(
        f"""
        CREATE TABLE report_conversion_paths AS
        WITH marked AS (
            SELECT
                *,
                LAG(channel) OVER (
                    PARTITION BY counter_id, conversion_visit_id
                    ORDER BY visit_datetime, step_visit_id
                ) AS previous_channel
            FROM conversion_path_steps
        ), collapsed AS (
            SELECT *
            FROM marked
            WHERE previous_channel IS NULL OR channel <> previous_channel
        ), path_windows AS (
            SELECT
                *,
                GROUP_CONCAT(channel, ' → ') OVER (
                    PARTITION BY counter_id, conversion_visit_id
                    ORDER BY visit_datetime, step_visit_id
                    ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
                ) AS conversion_path,
                COUNT(*) OVER (
                    PARTITION BY counter_id, conversion_visit_id
                ) AS path_length,
                MAX(CASE WHEN traffic_source = 'yandex_direct' THEN 1 ELSE 0 END)
                    OVER (
                        PARTITION BY counter_id, conversion_visit_id
                    ) AS contains_yandex_direct,
                ROW_NUMBER() OVER (
                    PARTITION BY counter_id, conversion_visit_id
                    ORDER BY visit_datetime DESC, step_visit_id DESC
                ) AS last_step_rank
            FROM collapsed
        ), per_conversion AS (
            SELECT
                counter_id,
                conversion_visit_id,
                goal_reaches,
                goal_revenue,
                conversion_path,
                path_length,
                channel AS closing_channel,
                contains_yandex_direct
            FROM path_windows
            WHERE last_step_rank = 1
        )
        SELECT
            :date_from AS date_from,
            :date_to AS date_to,
            :counter_ids AS counter_ids,
            :goal_ids AS goal_ids,
            'converted_visits' AS lead_basis,
            :metrika_model AS metrika_attribution_model,
            conversion_path,
            path_length,
            closing_channel,
            CASE contains_yandex_direct WHEN 1 THEN 'Да' ELSE 'Нет' END
                AS contains_yandex_direct,
            COUNT(*) AS converted_visits,
            ROUND(
                100.0 * COUNT(*) / (SELECT COUNT(*) FROM per_conversion),
                4
            ) AS conversion_share_pct,
            SUM(goal_reaches) AS goal_reaches
            {revenue_select}
        FROM per_conversion
        GROUP BY
            conversion_path, path_length, closing_channel,
            contains_yandex_direct
        ORDER BY converted_visits DESC, conversion_path
        """,
        params,
    )
    connection.commit()
    return int(
        connection.execute(
            "SELECT COUNT(*) FROM report_conversion_paths"
        ).fetchone()[0]
    )


def build_channel_roles(connection: sqlite3.Connection) -> None:
    """Summarize raw visit positions, counting channel presence once per path.

    Requires conversion_path_steps prepared by build_conversion_paths. A path
    is identified by (counter_id, conversion_visit_id), not by its channel text.
    The weight is the pooled share of touches, not an average of path shares.
    """
    connection.execute("DROP TABLE IF EXISTS report_channel_roles")
    connection.execute(
        """
        CREATE TABLE report_channel_roles AS
        WITH totals AS (
            SELECT
                COUNT(*) AS touches,
                SUM(visit_position = 1) AS paths,
                SUM(visit_position = 1 AND visits_in_path >= 2) AS paths_2plus,
                SUM(visit_position = 1 AND visits_in_path >= 3) AS paths_3plus
            FROM conversion_path_steps
        ), per_channel_path AS (
            SELECT counter_id, conversion_visit_id, channel, visits_in_path,
                COUNT(*) AS channel_touches,
                MAX(visit_position = 1) AS is_first,
                MAX(visit_position = visits_in_path) AS is_last,
                MAX(visit_position > 1 AND visit_position < visits_in_path)
                    AS is_middle
            FROM conversion_path_steps
            GROUP BY counter_id, conversion_visit_id, channel, visits_in_path
        )
        SELECT channel AS source,
            ROUND(100.0 * SUM(channel_touches) / NULLIF(t.touches, 0), 4)
                AS chain_weight_pct,
            ROUND(100.0 * COUNT(*) / NULLIF(t.paths, 0), 4)
                AS unique_chain_share_pct,
            ROUND(100.0 * SUM(is_first) / NULLIF(t.paths, 0), 4)
                AS first_touch_share_pct,
            ROUND(100.0 * SUM(is_first AND visits_in_path >= 2)
                / NULLIF(t.paths_2plus, 0), 4) AS first_touch_2plus_share_pct,
            ROUND(100.0 * SUM(is_middle) / NULLIF(t.paths_3plus, 0), 4)
                AS middle_touch_share_pct,
            ROUND(100.0 * SUM(is_last) / NULLIF(t.paths, 0), 4)
                AS last_touch_share_pct,
            ROUND(100.0 * SUM(is_last AND visits_in_path >= 2)
                / NULLIF(t.paths_2plus, 0), 4) AS last_touch_2plus_share_pct
        FROM per_channel_path
        CROSS JOIN totals t
        GROUP BY channel
        ORDER BY unique_chain_share_pct DESC, chain_weight_pct DESC, source
        """
    )




def channel_assists_sql(include_associated_revenue: bool = False) -> str:
    revenue_output = (
        ", SUM(goal_revenue) AS assisted_metrika_revenue"
        if include_associated_revenue
        else ""
    )
    return f"""
    WITH eligible AS (
        SELECT t.*
        FROM conversion_touches t
        JOIN visits closing
          ON closing.counter_id = t.counter_id
         AND closing.visit_id = t.conversion_visit_id
        WHERE t.traffic_source <> closing.traffic_source
    ), weighted AS (
        SELECT
            eligible.*,
            1.0 / COUNT(*) OVER (
                PARTITION BY counter_id, conversion_visit_id
            ) AS eligible_linear_conversion_credit
        FROM eligible
    )
    SELECT
        counter_id,
        traffic_source AS source,
        adv_engine,
        campaign_id,
        MAX(campaign_name) AS campaign_name,
        ad_group_id,
        MAX(ad_group_name) AS ad_group_name,
        MAX(criterion) AS criterion,
        MAX(criterion_type) AS criterion_type,
        network_type AS placement_type,
        MAX(placement) AS placement,
        utm_source, utm_medium, utm_campaign, utm_content, utm_term,
        COUNT(*) AS assisted_converted_visits,
        SUM(goal_reaches) AS assisted_goal_reaches,
        ROUND(SUM(eligible_linear_conversion_credit), 6)
            AS assisted_conversions_linear
        {revenue_output}
    FROM weighted
    GROUP BY
        counter_id, traffic_source, adv_engine, campaign_id, ad_group_id,
        criterion_key, network_type, placement_key, utm_source, utm_medium,
        utm_campaign, utm_content, utm_term
    ORDER BY assisted_converted_visits DESC, source
    """


def conversion_touches_sql(include_associated_revenue: bool = False) -> str:
    revenue_output = """
        , t.goal_revenue
        , ROUND(t.linear_revenue_credit, 8) AS linear_revenue_credit
    """ if include_associated_revenue else ""
    return f"""
    SELECT
        t.counter_id, t.conversion_visit_id, t.goal_reaches, t.conversion_date,
        closing.traffic_source AS closing_source,
        CASE
            WHEN t.traffic_source <> closing.traffic_source THEN 'eligible_assist'
            ELSE 'same_closing_source'
        END AS source_assist_status,
        t.touch_visit_id, t.touch_datetime, t.event_date AS touch_date,
        t.traffic_source AS source, t.adv_engine, t.campaign_id, t.campaign_name,
        t.ad_group_id, t.ad_group_name, t.criterion, t.criterion_type,
        t.network_type AS placement_type, t.placement,
        t.utm_source, t.utm_medium, t.utm_campaign, t.utm_content, t.utm_term,
        ROUND(t.linear_conversion_credit, 8) AS linear_conversion_credit,
        ROUND(t.linear_goal_reach_credit, 8) AS linear_goal_reach_credit
        {revenue_output}
    FROM conversion_touches t
    JOIN visits closing
      ON closing.counter_id = t.counter_id
     AND closing.visit_id = t.conversion_visit_id
    ORDER BY t.counter_id, t.conversion_date, t.conversion_visit_id,
             t.touch_datetime
    """


DIRECT_PERFORMANCE_BASE_SQL = """
WITH
direct AS (
    SELECT
        campaign_id, MAX(campaign_name) AS campaign_name,
        ad_group_id, MAX(ad_group_name) AS ad_group_name,
        criterion_key, MAX(criterion_id) AS criterion_id,
        MAX(criterion) AS criterion, MAX(criterion_type) AS criterion_type,
        network_type, placement_key, MAX(placement) AS placement,
        SUM(impressions) AS impressions,
        SUM(clicks) AS clicks,
        SUM(cost) AS cost,
        SUM(direct_api_converted_visits) AS direct_api_converted_visits,
        SUM(direct_api_revenue) AS direct_api_revenue
    FROM direct_stats
    WHERE event_date BETWEEN :date_from AND :date_to
    GROUP BY campaign_id, ad_group_id, criterion_key, network_type, placement_key
),
assists AS (
    SELECT
        campaign_id, MAX(campaign_name) AS campaign_name,
        ad_group_id, MAX(ad_group_name) AS ad_group_name,
        criterion_key, MAX(criterion) AS criterion,
        MAX(criterion_type) AS criterion_type,
        network_type, placement_key, MAX(placement) AS placement
    FROM conversion_touches
    WHERE traffic_source = 'yandex_direct'
    GROUP BY campaign_id, ad_group_id, criterion_key, network_type, placement_key
),
ordinary AS (
    SELECT
        v.campaign_id, MAX(v.campaign_name) AS campaign_name,
        v.ad_group_id, MAX(v.ad_group_name) AS ad_group_name,
        v.criterion_key, MAX(v.criterion) AS criterion,
        MAX(v.criterion_type) AS criterion_type,
        v.network_type, v.placement_key, MAX(v.placement) AS placement,
        COUNT(DISTINCT printf('%d:%s', c.counter_id, c.visit_id))
            AS ordinary_converted_visits,
        SUM(c.goal_reaches) AS ordinary_goal_reaches
    FROM goal_visits c
    JOIN visits v
      ON v.counter_id = c.counter_id
     AND v.visit_id = c.visit_id
    WHERE v.traffic_source = 'yandex_direct'
      AND c.conversion_date BETWEEN :date_from AND :date_to
    GROUP BY v.campaign_id, v.ad_group_id, v.criterion_key,
             v.network_type, v.placement_key
),
keys AS (
    SELECT campaign_id, ad_group_id, criterion_key, network_type, placement_key
    FROM direct
    UNION
    SELECT campaign_id, ad_group_id, criterion_key, network_type, placement_key
    FROM assists
    UNION
    SELECT campaign_id, ad_group_id, criterion_key, network_type, placement_key
    FROM ordinary
)
SELECT
    'yandex_direct' AS source,
    k.campaign_id,
    COALESCE(d.campaign_name, a.campaign_name, o.campaign_name, '') AS campaign_name,
    k.ad_group_id,
    COALESCE(d.ad_group_name, a.ad_group_name, o.ad_group_name, '') AS ad_group_name,
    COALESCE(d.criterion_id, '') AS criterion_id,
    COALESCE(d.criterion, a.criterion, o.criterion, '') AS criterion,
    k.criterion_key,
    COALESCE(d.criterion_type, a.criterion_type, o.criterion_type, '') AS criterion_type,
    k.network_type,
    COALESCE(d.placement, a.placement, o.placement, '') AS placement,
    k.placement_key,
    COALESCE(d.impressions, 0) AS impressions,
    COALESCE(d.clicks, 0) AS clicks,
    COALESCE(d.cost, 0) AS cost,
    COALESCE(d.direct_api_converted_visits, 0) AS direct_api_converted_visits,
    COALESCE(d.direct_api_revenue, 0) AS direct_api_revenue,
    COALESCE(o.ordinary_converted_visits, 0) AS ordinary_converted_visits,
    COALESCE(o.ordinary_goal_reaches, 0) AS ordinary_goal_reaches
FROM keys k
LEFT JOIN direct d USING (
    campaign_id, ad_group_id, criterion_key, network_type, placement_key
)
LEFT JOIN assists a USING (
    campaign_id, ad_group_id, criterion_key, network_type, placement_key
)
LEFT JOIN ordinary o USING (
    campaign_id, ad_group_id, criterion_key, network_type, placement_key
)
"""


SOURCE_PERFORMANCE_BASE_SQL = """
WITH
direct AS (
    SELECT
        'yandex_direct' AS source,
        SUM(impressions) AS impressions,
        SUM(clicks) AS clicks,
        SUM(cost) AS cost,
        SUM(direct_api_converted_visits) AS direct_api_converted_visits,
        SUM(direct_api_revenue) AS direct_api_revenue
    FROM direct_stats
    WHERE event_date BETWEEN :date_from AND :date_to
    GROUP BY source
),
ordinary AS (
    SELECT
        v.traffic_source AS source,
        COUNT(DISTINCT printf('%d:%s', c.counter_id, c.visit_id))
            AS ordinary_converted_visits,
        SUM(c.goal_reaches) AS ordinary_goal_reaches
    FROM goal_visits c
    JOIN visits v
      ON v.counter_id = c.counter_id
     AND v.visit_id = c.visit_id
    WHERE c.conversion_date BETWEEN :date_from AND :date_to
    GROUP BY v.traffic_source
),
assists AS (
    SELECT
        t.traffic_source AS source
    FROM conversion_touches t
    JOIN visits closing
      ON closing.counter_id = t.counter_id
     AND closing.visit_id = t.conversion_visit_id
    WHERE t.traffic_source <> closing.traffic_source
    GROUP BY t.traffic_source
),
keys AS (
    SELECT source FROM direct
    UNION SELECT source FROM ordinary
    UNION SELECT source FROM assists
)
SELECT
    k.source,
    COALESCE(d.impressions, 0) AS impressions,
    COALESCE(d.clicks, 0) AS clicks,
    COALESCE(d.cost, 0) AS cost,
    COALESCE(d.direct_api_converted_visits, 0) AS direct_api_converted_visits,
    COALESCE(d.direct_api_revenue, 0) AS direct_api_revenue,
    COALESCE(o.ordinary_converted_visits, 0) AS ordinary_converted_visits,
    COALESCE(o.ordinary_goal_reaches, 0) AS ordinary_goal_reaches
FROM keys k
LEFT JOIN direct d USING (source)
LEFT JOIN ordinary o USING (source)
LEFT JOIN assists a USING (source)
"""


REPORT_SLICES: dict[str, tuple[list[str], list[str]]] = {
    "campaign": (
        ["campaign_id"],
        ["campaign_id", "MAX(campaign_name) AS campaign_name"],
    ),
    "ad_group": (
        ["campaign_id", "ad_group_id"],
        [
            "campaign_id", "MAX(campaign_name) AS campaign_name",
            "ad_group_id", "MAX(ad_group_name) AS ad_group_name",
        ],
    ),
    "criterion": (
        ["campaign_id", "ad_group_id", "criterion_key"],
        [
            "campaign_id", "MAX(campaign_name) AS campaign_name",
            "ad_group_id", "MAX(ad_group_name) AS ad_group_name",
            "MAX(criterion_id) AS criterion_id", "MAX(criterion) AS criterion",
            "MAX(criterion_type) AS criterion_type",
        ],
    ),
    "placement_type": (
        ["network_type"],
        ["network_type AS placement_type"],
    ),
    "placement": (
        ["network_type", "placement_key"],
        ["network_type AS placement_type", "MAX(placement) AS placement"],
    ),
    "full_detail": (
        [
            "campaign_id", "ad_group_id", "criterion_key",
            "network_type", "placement_key",
        ],
        [
            "campaign_id", "MAX(campaign_name) AS campaign_name",
            "ad_group_id", "MAX(ad_group_name) AS ad_group_name",
            "MAX(criterion_id) AS criterion_id", "MAX(criterion) AS criterion",
            "MAX(criterion_type) AS criterion_type",
            "network_type AS placement_type", "MAX(placement) AS placement",
        ],
    ),
}


def _kpi_select_sql(
    dimension_selects: list[str],
    group_columns: list[str],
    source_table: str,
    settings: Settings,
    conversion_date_from: date,
    conversion_date_to: date,
) -> str:
    dimensions = ",\n            ".join(dimension_selects)
    metric_join_dimensions = ",\n                ".join(
        f"{column} AS _join_{column}" for column in group_columns
    )
    dimension_columns = ",\n            ".join(
        re.search(r"\s+AS\s+(\w+)\s*$", expression, re.IGNORECASE).group(1)
        if re.search(r"\s+AS\s+(\w+)\s*$", expression, re.IGNORECASE)
        else expression.strip()
        for expression in dimension_selects
    )
    group_by = ", ".join(group_columns)
    if source_table == "performance_source_base":
        assist_dimension_selects = ["t.traffic_source AS source"]
        assist_group_columns = ["t.traffic_source"]
        assist_output_columns = ["source"]
        assist_source_filter = ""
        same_closing_slice = "t.traffic_source = closing.traffic_source"
    else:
        assist_dimension_selects = [
            f"t.{column} AS {column}" for column in group_columns
        ]
        assist_group_columns = [f"t.{column}" for column in group_columns]
        assist_output_columns = group_columns
        assist_source_filter = "AND t.traffic_source = 'yandex_direct'"
        same_closing_slice = " AND ".join(
            f"t.{column} = closing.{column}" for column in group_columns
        )
    assist_dimensions = ", ".join(assist_dimension_selects)
    assist_group_by = ", ".join(assist_group_columns)
    assist_outputs = ", ".join(assist_output_columns)
    assist_join = " AND ".join(
        f"m._join_{column} = a.{column}" for column in assist_output_columns
    )
    counter_ids_text = ",".join(
        str(counter_id) for counter_id in settings.metrika_counter_ids
    )
    goal_ids_text = ",".join(str(goal_id) for goal_id in settings.goal_ids)
    assist_deduplicated_revenue = (
        ", MAX(t.goal_revenue) AS goal_revenue"
        if settings.include_associated_revenue
        else ""
    )
    assist_metrics_revenue = (
        ", SUM(goal_revenue) AS assisted_metrika_revenue"
        if settings.include_associated_revenue
        else ""
    )
    metrics_with_assists_revenue = (
        ", COALESCE(a.assisted_metrika_revenue, 0) AS assisted_metrika_revenue"
        if settings.include_associated_revenue
        else ""
    )
    report_revenue_column = (
        ", ROUND(assisted_metrika_revenue, 6) AS assisted_metrika_revenue"
        if settings.include_associated_revenue
        else ""
    )
    return f"""
        WITH metrics AS (
            SELECT
                {dimensions},
                {metric_join_dimensions},
                SUM(impressions) AS impressions,
                SUM(clicks) AS clicks,
                ROUND(SUM(cost), 2) AS cost,
                SUM(direct_api_converted_visits) AS direct_leads,
                ROUND(SUM(direct_api_revenue), 6) AS direct_revenue,
                SUM(ordinary_converted_visits) AS metrika_leads,
                SUM(ordinary_goal_reaches) AS ordinary_goal_reaches
            FROM {source_table}
            GROUP BY {group_by}
        ), assist_deduplicated AS (
            SELECT
                {assist_dimensions},
                t.counter_id,
                t.conversion_visit_id,
                MAX(t.goal_reaches) AS goal_reaches
                {assist_deduplicated_revenue}
            FROM conversion_touches t
            JOIN visits closing
              ON closing.counter_id = t.counter_id
             AND closing.visit_id = t.conversion_visit_id
            WHERE t.conversion_date BETWEEN '{conversion_date_from.isoformat()}'
                                       AND '{conversion_date_to.isoformat()}'
              {assist_source_filter}
              AND NOT ({same_closing_slice})
            GROUP BY
                {assist_group_by}, t.counter_id, t.conversion_visit_id
        ), assist_metrics AS (
            SELECT
                {assist_outputs},
                COUNT(*) AS assisted_metrika_conversions,
                SUM(goal_reaches) AS assisted_goal_reaches
                {assist_metrics_revenue}
            FROM assist_deduplicated
            GROUP BY {assist_outputs}
        ), metrics_with_assists AS (
            SELECT
                m.*,
                COALESCE(a.assisted_metrika_conversions, 0)
                    AS assisted_metrika_conversions,
                COALESCE(a.assisted_goal_reaches, 0)
                    AS assisted_goal_reaches
                {metrics_with_assists_revenue}
            FROM metrics m
            LEFT JOIN assist_metrics a ON {assist_join}
        ), kpi_with_total AS (
            SELECT
                *,
                metrika_leads + assisted_metrika_conversions
                    AS total_metrika_conversions
            FROM metrics_with_assists
        )
        SELECT
            '{conversion_date_from.isoformat()}' AS date_from,
            '{conversion_date_to.isoformat()}' AS date_to,
            '{counter_ids_text}' AS counter_ids,
            '{goal_ids_text}' AS goal_ids,
            'converted_visits' AS lead_basis,
            '{settings.direct_conversion_model}' AS direct_attribution_model,
            '{settings.metrika_attribution_model}' AS metrika_attribution_model,
            {dimension_columns},
            impressions,
            clicks,
            cost,
            direct_revenue,
            CASE WHEN direct_revenue > 0
                 THEN ROUND(100.0 * cost / direct_revenue, 4)
                 END AS drr_direct_pct,
            CASE WHEN cost > 0
                 THEN ROUND(100.0 * (direct_revenue - cost) / cost, 4)
                 END AS romi_direct_pct,
            direct_leads,
            CASE WHEN clicks > 0
                 THEN ROUND(100.0 * direct_leads / clicks, 4) END AS cr_direct_pct,
            CASE WHEN direct_leads > 0
                 THEN ROUND(cost / direct_leads + 0.000000001, 2) END AS cpl_direct,
            metrika_leads,
            CASE WHEN clicks > 0
                 THEN ROUND(100.0 * metrika_leads / clicks, 4) END AS cr_metrika_pct,
            CASE WHEN metrika_leads > 0
                 THEN ROUND(cost / metrika_leads + 0.000000001, 2) END AS cpl_metrika,
            assisted_metrika_conversions,
            total_metrika_conversions,
            CASE WHEN clicks > 0
                 THEN ROUND(100.0 * total_metrika_conversions / clicks, 4)
                 END AS cr_total_metrika_pct,
            CASE WHEN total_metrika_conversions > 0
                 THEN ROUND(cost / total_metrika_conversions + 0.000000001, 2)
                 END AS cpl_total_metrika,
            CASE WHEN metrika_leads > 0
                 THEN ROUND(
                     1.0 * assisted_metrika_conversions / metrika_leads, 4
                 ) END
                 AS association_coefficient,
            CASE
                WHEN metrika_leads > 0 THEN 'ok'
                WHEN assisted_metrika_conversions > 0 THEN 'only_assisted'
                ELSE 'no_leads'
            END AS association_status,
            ordinary_goal_reaches,
            assisted_goal_reaches
            {report_revenue_column}
        FROM kpi_with_total
        WHERE impressions <> 0
           OR clicks <> 0
           OR cost <> 0
           OR direct_revenue <> 0
           OR direct_leads <> 0
           OR metrika_leads <> 0
           OR assisted_metrika_conversions <> 0
    """


def create_performance_tables(
    connection: sqlite3.Connection,
    settings: Settings,
    conversion_date_from: date,
    conversion_date_to: date,
) -> list[str]:
    params = {
        "date_from": conversion_date_from.isoformat(),
        "date_to": conversion_date_to.isoformat(),
    }
    connection.execute("DROP TABLE IF EXISTS performance_direct_base")
    connection.execute(
        f"CREATE TABLE performance_direct_base AS {DIRECT_PERFORMANCE_BASE_SQL}",
        params,
    )
    connection.execute("DROP TABLE IF EXISTS performance_source_base")
    connection.execute(
        f"CREATE TABLE performance_source_base AS {SOURCE_PERFORMANCE_BASE_SQL}",
        params,
    )

    table_names = ["report_by_source"]
    source_sql = _kpi_select_sql(
        ["source"],
        ["source"],
        "performance_source_base",
        settings,
        conversion_date_from,
        conversion_date_to,
    )
    connection.execute("DROP TABLE IF EXISTS report_by_source")
    connection.execute(f"CREATE TABLE report_by_source AS {source_sql}")

    for name, (groups, selects) in REPORT_SLICES.items():
        table_name = f"report_by_{name}"
        connection.execute(f"DROP TABLE IF EXISTS {table_name}")
        report_sql = _kpi_select_sql(
            selects,
            groups,
            "performance_direct_base",
            settings,
            conversion_date_from,
            conversion_date_to,
        )
        connection.execute(f"CREATE TABLE {table_name} AS {report_sql}")
        table_names.append(table_name)
    connection.commit()
    return table_names


def export_outputs(
    connection: sqlite3.Connection,
    output_path: Path,
    settings: Settings,
    conversion_date_from: date,
    conversion_date_to: date,
    counts: dict[str, int],
) -> dict[str, object]:
    cohort_date_from, cohort_date_to = association_window(
        conversion_date_from,
        conversion_date_to,
        settings.lookback_days,
    )
    performance_tables = create_performance_tables(
        connection,
        settings,
        conversion_date_from,
        conversion_date_to,
    )
    performance_tables.extend(["report_conversion_paths", "report_channel_roles"])
    channel_sql = channel_assists_sql(settings.include_associated_revenue)
    touches_sql = conversion_touches_sql(settings.include_associated_revenue)
    output_counts: dict[str, int] = {}
    for table_name in performance_tables:
        file_stem = table_name.removeprefix("report_")
        output_counts[f"{file_stem}_rows"] = int(
            connection.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0]
        )
    output_counts.update(
        {
            "channel_assist_rows": int(
                connection.execute(
                    f"SELECT COUNT(*) FROM ({channel_sql})"
                ).fetchone()[0]
            ),
            "conversion_touch_rows": int(
                connection.execute("SELECT COUNT(*) FROM conversion_touches").fetchone()[0]
            ),
        }
    )
    selected_goals_detail: dict[str, dict[str, int | float | str]] = {}
    for goal_id in settings.require_goal_ids():
        goal_totals = connection.execute(
            """
            SELECT
                COUNT(*),
                COALESCE(SUM(goal_reaches), 0),
                COALESCE(SUM(goal_revenue), 0),
                GROUP_CONCAT(DISTINCT NULLIF(goal_currency, ''))
            FROM goal_visits
            WHERE goal_id = :goal_id
              AND conversion_date BETWEEN :date_from AND :date_to
            """,
            {
                "goal_id": goal_id,
                "date_from": conversion_date_from.isoformat(),
                "date_to": conversion_date_to.isoformat(),
            },
        ).fetchone()
        selected_goals_detail[str(goal_id)] = {
            "goal_visits": goal_totals[0],
            "goal_reaches": goal_totals[1],
        }
        if settings.include_associated_revenue:
            selected_goals_detail[str(goal_id)].update(
                {
                    "goal_revenue": round(float(goal_totals[2]), 6),
                    "goal_currency": goal_totals[3] or "",
                }
            )
    totals = connection.execute(
        """
        WITH eligible_source_rows AS (
            SELECT
                t.counter_id,
                t.conversion_visit_id,
                t.traffic_source,
                MAX(t.goal_revenue) AS goal_revenue
            FROM conversion_touches t
            JOIN visits closing
              ON closing.counter_id = t.counter_id
             AND closing.visit_id = t.conversion_visit_id
            WHERE t.conversion_date BETWEEN :date_from AND :date_to
              AND t.traffic_source <> closing.traffic_source
            GROUP BY
                t.counter_id, t.conversion_visit_id, t.traffic_source
        ), weighted_source_rows AS (
            SELECT
                *,
                1.0 / COUNT(*) OVER (
                    PARTITION BY counter_id, conversion_visit_id
                ) AS linear_credit
            FROM eligible_source_rows
        )
        SELECT
            (SELECT COALESCE(SUM(impressions), 0)
               FROM direct_stats
              WHERE event_date BETWEEN :date_from AND :date_to),
            (SELECT COALESCE(SUM(clicks), 0)
               FROM direct_stats
              WHERE event_date BETWEEN :date_from AND :date_to),
            (SELECT COALESCE(SUM(cost), 0)
               FROM direct_stats
              WHERE event_date BETWEEN :date_from AND :date_to),
            (SELECT COALESCE(SUM(direct_api_converted_visits), 0)
               FROM direct_stats
              WHERE event_date BETWEEN :date_from AND :date_to),
            (SELECT COUNT(DISTINCT printf('%d:%s', counter_id, visit_id))
               FROM goal_visits
              WHERE conversion_date BETWEEN :date_from AND :date_to),
            (SELECT COALESCE(SUM(goal_reaches), 0)
               FROM goal_visits
              WHERE conversion_date BETWEEN :date_from AND :date_to),
            (SELECT COUNT(*)
               FROM conversion_touches
              WHERE conversion_date BETWEEN :date_from AND :date_to),
            (SELECT COUNT(DISTINCT printf(
                       '%d:%s', counter_id, conversion_visit_id
                   ))
               FROM eligible_source_rows),
            (SELECT COALESCE(SUM(linear_credit), 0)
               FROM weighted_source_rows),
            (SELECT COALESCE(SUM(goal_revenue), 0)
               FROM eligible_source_rows),
            (SELECT COALESCE(SUM(direct_api_revenue), 0)
               FROM direct_stats
              WHERE event_date BETWEEN :date_from AND :date_to)
        """,
        {
            "date_from": conversion_date_from.isoformat(),
            "date_to": conversion_date_to.isoformat(),
        },
    ).fetchone()
    direct_revenue_total = float(totals[10])
    direct_cost_total = float(totals[2])
    qa = {
        "metrika_counter_ids": list(settings.metrika_counter_ids),
        "goal_ids": list(settings.goal_ids),
        "conversion_cohort": {
            "date_from": conversion_date_from.isoformat(),
            "date_to": conversion_date_to.isoformat(),
            "lookback_days": settings.lookback_days,
            "association_date_from": cohort_date_from.isoformat(),
            "association_date_to": cohort_date_to.isoformat(),
            "association_logic": "ua_mcf_prior_yandex_direct_interactions",
            "metrika_lead_basis": "converted_visits",
            "direct_attribution_model": settings.direct_conversion_model,
            "metrika_attribution_model": settings.metrika_attribution_model,
            "direct_cost_includes_vat": settings.include_vat,
            "include_associated_revenue": settings.include_associated_revenue,
        },
        "loaded": counts,
        "outputs": output_counts,
        "sqlite_report_tables": performance_tables,
        "selected_goals_detail": selected_goals_detail,
        "combined_totals": {
            "direct_impressions": totals[0],
            "direct_clicks": totals[1],
            "direct_cost": totals[2],
            "direct_revenue": round(direct_revenue_total, 6),
            "direct_drr_pct": (
                round(100.0 * direct_cost_total / direct_revenue_total, 6)
                if direct_revenue_total > 0
                else None
            ),
            "direct_romi_pct": (
                round(
                    100.0
                    * (direct_revenue_total - direct_cost_total)
                    / direct_cost_total,
                    6,
                )
                if direct_cost_total > 0
                else None
            ),
            "direct_api_converted_visits_sum": totals[3],
            "unique_goal_visits": totals[4],
            "goal_reaches": totals[5],
            "conversion_touch_rows_non_additive": totals[6],
            "assisted_metrika_conversions_unique": totals[7],
            "metrika_conversions_with_assists": totals[4] + totals[7],
            "linear_conversion_credit_total": round(float(totals[8]), 6),
        },
        "notes": [
            "Все выбранные цели объединяются до расчёта KPI.",
            "Визиты разных счётчиков изолируются по counter_id до объединения.",
            "Лиды Директа, лиды Метрики и ассоциированные конверсии выводятся отдельно.",
            "Лиды Метрики и ассоциированные конверсии считаются по визитам.",
            "Общий показатель Метрики = лиды Метрики + ассоциированные конверсии Метрики; лиды Директа в эту сумму не входят.",
            "Достижения целей сохраняются как справочные показатели.",
            "Лиды Директа рассчитаны по модели из DIRECT_CONVERSION_MODEL.",
            "Источники конверсионных и предыдущих визитов определяются по модели из METRIKA_ATTRIBUTION_MODEL; для UA-style логики используется LAST.",
            "Все расходы и CPL Директа выгружены с НДС."
            if settings.include_vat
            else "Все расходы и CPL Директа выгружены без НДС.",
            "Доход Директа — сумма Revenue всех выбранных целей по модели DIRECT_CONVERSION_MODEL.",
            "ДРР Директа = расход / доход × 100%; ROMI Директа = (доход − расход) / расход × 100%.",
            "ROMI Директа учитывает только расход на рекламу; себестоимость, маржа и другие затраты не входят в расчёт.",
            "При сумме нескольких целей Директ может учитывать один визит более одного раза.",
            "Ассоциированная конверсия Директа учитывается, только если конверсионный визит по LAST не относится к Директу.",
            "Для каждого целевого визита ищутся только более ранние визиты из Директа или «Директ: Не определено» не дальше чем за LOOKBACK_DAYS до конверсии.",
            "Окно LOOKBACK_DAYS пересчитывается отдельно для каждой конверсии; визиты Директа после конверсии не учитываются.",
            "Лист «Основные пути» включает все каналы до целевого визита и сам закрывающий канал; соседние повторы канала схлопываются.",
            "Лист «Роль каналов» использует те же целевые визиты и окно LOOKBACK_DAYS, но каждое касание — отдельный визит, без схлопывания повторов.",
            "На листе «Роль каналов» одна цепочка соответствует одному целевому визиту; одинаковые последовательности разных конверсий считаются отдельно. Промежуточные касания считаются среди цепочек из 3 и более визитов.",
            "Директ с неопределённой кампанией входит в показатель на уровне источника, но не распределяется по кампании, группе, ключу или площадке.",
            "Полные ассисты неаддитивны между строками среза.",
            "Линейный кредит сохраняется на технических листах и делит конверсию между допустимыми касаниями.",
            "Контрольные конверсии Директа и обычные конверсии Logs API могут различаться.",
            "CR и CPL считаются для лидов Директа, лидов Метрики и общего показателя Метрики.",
            "Коэффициент ассоциированности = ассоциированные конверсии Метрики / лиды Метрики.",
            "Каждый лист со срезом рассчитывается независимо на своём уровне агрегации.",
        ],
    }
    if settings.include_associated_revenue:
        qa["combined_totals"]["assisted_metrika_revenue_unique"] = round(
            float(totals[9]), 6
        )
        qa["notes"].append(
            "Ассоциированный доход — сумма дохода выбранных целей из целевых визитов; технический коэффициент Logs API 1000 удалён. Между строками среза показатель неаддитивен."
        )
    excel_info = export_excel_workbook(
        connection,
        output_path,
        qa,
        channel_sql,
        touches_sql,
    )
    qa["excel_workbook"] = excel_info
    return qa


def build(
    direct_paths: list[Path],
    metrika_paths: list[Path],
    database_path: Path,
    output_path: Path,
    settings: Settings,
    conversion_date_from: date,
    conversion_date_to: date,
) -> dict[str, object]:
    connection = connect(database_path)
    try:
        create_schema(connection)
        visit_count, goal_visit_count = load_metrika_visits(
            connection, metrika_paths, settings
        )
        direct_count = load_direct_stats(connection, direct_paths, settings)
        touch_count = build_conversion_touches(
            connection, settings, conversion_date_from, conversion_date_to
        )
        path_count = build_conversion_paths(
            connection, settings, conversion_date_from, conversion_date_to
        )
        return export_outputs(
            connection,
            output_path,
            settings,
            conversion_date_from,
            conversion_date_to,
            {
                "metrika_visits": visit_count,
                "selected_goal_visit_rows": goal_visit_count,
                "direct_rows": direct_count,
                "conversion_touches": touch_count,
                "conversion_paths": path_count,
            },
        )
    finally:
        connection.close()
