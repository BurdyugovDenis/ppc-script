from __future__ import annotations

import sqlite3
import unittest
from datetime import date

from yandex_attribution.config import Settings
from yandex_attribution.pipeline import build_conversion_paths, create_schema


class ChannelRolesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.connection = sqlite3.connect(":memory:")
        self.connection.row_factory = sqlite3.Row
        create_schema(self.connection)
        self.settings = Settings(
            metrika_counter_ids=(1, 2),
            direct_client_login="test-client",
            goal_ids=(99, 100),
            lookback_days=90,
        )

    def tearDown(self) -> None:
        self.connection.close()

    def add_visit(
        self, visit_id: str, user: str, source: str, timestamp: str,
        *, counter: int = 1, goals: tuple[int, ...] = (),
    ) -> None:
        self.connection.execute(
            "INSERT INTO visits VALUES (" + ",".join(["?"] * 22) + ")",
            (counter, visit_id, user, timestamp, timestamp[:10], source)
            + ("",) * 16,
        )
        for goal in goals:
            self.connection.execute(
                "INSERT INTO goal_visits VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (counter, visit_id, user, timestamp, timestamp[:10], goal, 2, 0, ""),
            )

    def build_roles(self) -> dict[str, dict]:
        build_conversion_paths(
            self.connection, self.settings, date(2026, 8, 1), date(2026, 8, 31)
        )
        return {
            row["source"]: dict(row)
            for row in self.connection.execute("SELECT * FROM report_channel_roles")
        }

    def test_touch_weights_presence_and_all_positions(self) -> None:
        paths = [
            ["yandex_direct"],
            ["yandex_direct", "organic"],
            ["organic", "yandex_direct", "organic"],
            ["yandex_direct", "yandex_direct", "email", "yandex_direct", "organic"],
        ]
        for path_number, path in enumerate(paths):
            for step, source in enumerate(path, start=1):
                self.add_visit(
                    f"{path_number}-{step}", f"user-{path_number}", source,
                    f"2026-08-{step:02d} 10:00:00",
                    goals=(99, 100) if step == len(path) else (),
                )
        roles = self.build_roles()
        direct = roles["Яндекс Директ"]
        search = roles["Поисковые системы"]
        email = roles["Email"]
        self.assertEqual(direct["chain_weight_pct"], 54.5455)  # 6 / 11 touches
        self.assertEqual(search["chain_weight_pct"], 36.3636)
        self.assertEqual(email["chain_weight_pct"], 9.0909)
        self.assertEqual(direct["unique_chain_share_pct"], 100.0)
        self.assertEqual(search["unique_chain_share_pct"], 75.0)
        self.assertEqual(email["unique_chain_share_pct"], 25.0)
        self.assertEqual(direct["first_touch_share_pct"], 75.0)
        self.assertEqual(search["first_touch_share_pct"], 25.0)
        self.assertEqual(direct["first_touch_2plus_share_pct"], 66.6667)
        self.assertEqual(search["first_touch_2plus_share_pct"], 33.3333)
        # Exactly three visits qualify. Multiple middle Direct visits count once.
        self.assertEqual(direct["middle_touch_share_pct"], 100.0)
        self.assertEqual(email["middle_touch_share_pct"], 50.0)
        self.assertEqual(search["middle_touch_share_pct"], 0.0)
        self.assertEqual(direct["last_touch_share_pct"], 25.0)
        self.assertEqual(search["last_touch_share_pct"], 75.0)
        self.assertEqual(direct["last_touch_2plus_share_pct"], 0.0)
        self.assertEqual(search["last_touch_2plus_share_pct"], 100.0)
        for column in (
            "chain_weight_pct", "first_touch_share_pct", "last_touch_share_pct",
            "first_touch_2plus_share_pct", "last_touch_2plus_share_pct",
        ):
            self.assertAlmostEqual(sum(row[column] for row in roles.values()), 100, places=3)
        self.assertEqual(self.connection.execute(
            "SELECT SUM(converted_visits) FROM report_conversion_paths"
        ).fetchone()[0], 4)

    def test_counter_isolation_90_day_boundary_and_future_exclusion(self) -> None:
        for counter in (1, 2):
            # Same user/visit IDs on different counters must remain separate.
            self.add_visit("old", "user", "social", "2026-05-12 09:59:59", counter=counter)
            self.add_visit("start", "user", "yandex_direct", "2026-05-12 10:00:00", counter=counter)
            self.add_visit("end", "user", "organic", "2026-08-10 10:00:00",
                           counter=counter, goals=(99, 100))
            self.add_visit("same-time", "user", "email", "2026-08-10 10:00:00", counter=counter)
            self.add_visit("future", "user", "email", "2026-08-11 10:00:00", counter=counter)
        # An older target visit is history, not an extra August conversion.
        self.add_visit("prior-goal", "other-user", "email", "2026-07-31 10:00:00", goals=(99,))
        roles = self.build_roles()
        self.assertEqual(set(roles), {"Яндекс Директ", "Поисковые системы"})
        self.assertEqual(roles["Яндекс Директ"]["chain_weight_pct"], 50.0)
        self.assertEqual(roles["Яндекс Директ"]["first_touch_2plus_share_pct"], 100.0)
        self.assertIsNone(roles["Яндекс Директ"]["middle_touch_share_pct"])
        path = self.connection.execute("SELECT * FROM report_conversion_paths").fetchone()
        self.assertEqual(path["converted_visits"], 2)
        self.assertEqual(path["path_length"], 2)
        self.assertEqual(self.connection.execute("SELECT COUNT(*) FROM conversion_path_steps").fetchone()[0], 4)

    def test_rebuild_empty_and_single_visit_paths(self) -> None:
        self.assertEqual(self.build_roles(), {})
        self.add_visit("single", "user", "direct", "2026-08-10 10:00:00", goals=(99,))
        role = self.build_roles()["Прямые заходы"]
        for column in ("chain_weight_pct", "unique_chain_share_pct", "first_touch_share_pct", "last_touch_share_pct"):
            self.assertEqual(role[column], 100.0)
        for column in ("first_touch_2plus_share_pct", "middle_touch_share_pct", "last_touch_2plus_share_pct"):
            self.assertIsNone(role[column])
        self.connection.execute("DELETE FROM goal_visits")
        self.assertEqual(self.build_roles(), {})


if __name__ == "__main__":
    unittest.main()
