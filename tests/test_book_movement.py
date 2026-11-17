import json

from ingestion.book_movement import observed_book_movement


def _ledger(kind, observed, price):
    return {
        "league": "NBA",
        "target_date": "20261120",
        "ledger_kind": kind,
        "games": [
            {"game_id": "g1", "scheduled_utc": "2026-11-20T22:00:00Z"}
        ],
        "book_snapshots": [
            {
                "game_id": "g1",
                "provider": "book-feed",
                "book": "Example",
                "market": "moneyline",
                "period": "full_game",
                "observed_utc": observed,
                "scheduled_utc": "2026-11-20T22:00:00Z",
                "pre_match": True,
                "home_price_decimal": price,
                "away_price_decimal": 2.2,
            }
        ],
    }


def test_observed_book_movement_uses_only_distinct_pre_tip_snapshot_times(tmp_path):
    first = tmp_path / "pregame-1.json"
    later = tmp_path / "freeze.json"
    duplicate = tmp_path / "freeze-duplicate.json"
    late = tmp_path / "post.json"
    first_payload = _ledger("pregame", "2026-11-20T16:00:00Z", 1.8)
    first_payload["book_snapshots"][0]["frozen"] = False
    first.write_text(json.dumps(first_payload))
    duplicate_payload = _ledger("freeze", "2026-11-20T16:00:00Z", 1.8)
    duplicate_payload["book_snapshots"][0]["frozen"] = True
    duplicate.write_text(json.dumps(duplicate_payload))
    later_payload = _ledger("freeze", "2026-11-20T20:00:00Z", 1.9)
    later_payload["book_snapshots"][0]["frozen"] = True
    later.write_text(json.dumps(later_payload))
    after_tip = _ledger("results", "2026-11-20T22:01:00Z", 1.7)
    after_tip["book_snapshots"][0]["pre_match"] = False
    late.write_text(json.dumps(after_tip))

    report = observed_book_movement(
        [first, duplicate, later, late], league="NBA", target_date="2026-11-20"
    )

    assert report["slots_with_multiple_timestamped_observations"] == 1
    row = report["movements"][0]
    assert row["observations"] == 2
    assert report["slots_with_conflicting_values_at_same_timestamp"] == 0
    assert row["deltas"]["home_price_decimal"] == 0.1
    assert "not CLV" in report["interpretation"]
    assert report["exclusions"] == [
        "untimestamped scoreboard open/close context",
        "post-match snapshots",
    ]


def test_observed_book_movement_fails_closed_on_conflicting_same_time_rows(tmp_path):
    first = _ledger("pregame", "2026-11-20T16:00:00Z", 1.8)
    conflict = _ledger("freeze", "2026-11-20T16:00:00Z", 1.85)
    later = _ledger("results", "2026-11-20T20:00:00Z", 1.9)
    paths = []
    for index, payload in enumerate((first, conflict, later)):
        path = tmp_path / f"ledger-{index}.json"
        path.write_text(json.dumps(payload))
        paths.append(path)

    report = observed_book_movement(paths)

    assert report["slots_with_multiple_timestamped_observations"] == 0
    assert report["slots_with_conflicting_values_at_same_timestamp"] == 1
