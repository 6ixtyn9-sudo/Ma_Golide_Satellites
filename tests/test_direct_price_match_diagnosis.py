"""Why a lane has no price: name disagreement, not an absent source.

``BOOK_PROVIDERS_BY_LEAGUE`` registers ``('betexplorer', 'betexplorer_direct')``
for ABA League, KBL and Japan B.League alike -- there is no privileged ABA price
path. The direct fallback is league-agnostic: it reads the day board and binds
rows by Paris date plus guarded ordered team names. So the difference between a
priced lane and an unpriced one is usually whether our team name agrees with the
board's name, not whether the source covers the league.

The adapter already records that verdict per candidate, but the pipeline used to
collapse a zero-snapshot run into ``fetched_no_snapshots``, which reads like "no
prices exist for this league" -- the silent-empty-result this project refuses,
and the label standing between a lane and its prices.

These tests use *receipts*, i.e. our own published data structure, not scraped
markup. The receipt shape asserted here is the one ``fetch_odds`` writes.
"""

from unittest.mock import patch

from ingestion.models import Game, GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import betexplorer, betexplorer_direct
from ingestion.store import InMemoryStore

RECEIPT = {
    "status": "success",
    "candidate_games": 3,
    "day_rows_seen": 311,
    "fixtures_matched": 0,
    "fixtures_unmatched": 3,
    "candidate_fixture_diagnostics": [
        {
            "game_id": "kbl-1",
            "status": "unmatched",
            "reason": "ordered_team_pair_not_found_on_board_date",
            "home": "WONJU DB PROMY",
            "away": "SEOUL SAMSUNG THUNDERS",
            "board_date": "2026-11-25",
            "board_rows_on_date": 12,
            "nearest_board_rows": [
                {
                    "home": "Wonju Dongbu Promy",
                    "away": "Seoul Samsung Thunders",
                    "shared_token_count": 3,
                    "shared_tokens": ["promy", "seoul", "thunders"],
                    "orientation": "ordered",
                }
            ],
        },
        {
            "game_id": "kbl-2",
            "status": "unmatched",
            "reason": "no_fixture_rows_on_paris_board_date",
            "home": "Busan KCC Egis",
            "away": "Ulsan Hyundai Mobis",
            "board_date": "2026-11-26",
            "board_rows_on_date": 0,
            "nearest_board_rows": [],
        },
        {
            "game_id": "kbl-3",
            "status": "ambiguous",
            "reason": "ambiguous_ordered_team_pair",
            "home": "Anyang Jung Kwan Jang",
            "away": "Suwon KT Sonicboom",
            "board_date": "2026-11-25",
            "board_rows_on_date": 12,
            "nearest_board_rows": [],
        },
    ],
}


def test_the_verdict_names_the_counts_and_the_dominant_reason():
    diagnosis = betexplorer_direct.summarize_match_failure(RECEIPT)

    assert diagnosis["candidates"] == 3
    assert diagnosis["board_rows"] == 311
    assert diagnosis["matched"] == 0
    assert diagnosis["reasons"] == {
        "ordered_team_pair_not_found_on_board_date": 1,
        "no_fixture_rows_on_paris_board_date": 1,
        "ambiguous_ordered_team_pair": 1,
    }
    assert "0 of 3 candidate game(s) matched 311 board row(s)" in diagnosis["verdict"]


def test_the_examples_print_the_pair_that_did_not_agree():
    diagnosis = betexplorer_direct.summarize_match_failure(RECEIPT, example_limit=1)

    assert len(diagnosis["examples"]) == 1
    example = diagnosis["examples"][0]
    assert example["scheduled_home"] == "WONJU DB PROMY"
    assert example["nearest_board_home"] == "Wonju Dongbu Promy"
    assert "promy" in example["shared_tokens"]


def test_examples_are_bounded():
    assert len(betexplorer_direct.summarize_match_failure(RECEIPT, example_limit=2)["examples"]) == 2


def test_a_matched_run_is_not_reported_as_a_failure():
    receipt = {
        "status": "success",
        "candidate_games": 2,
        "day_rows_seen": 40,
        "fixtures_matched": 2,
        "candidate_fixture_diagnostics": [
            {"game_id": "a", "status": "matched", "reason": "exact_ordered_pair_on_paris_board_date"},
            {"game_id": "b", "status": "matched", "reason": "exact_ordered_pair_on_paris_board_date"},
        ],
    }
    diagnosis = betexplorer_direct.summarize_match_failure(receipt)

    assert diagnosis["reasons"] == {}
    assert diagnosis["examples"] == []
    assert diagnosis["verdict"] == "2 of 2 candidate game(s) matched"


def test_a_missing_receipt_says_so_instead_of_guessing():
    diagnosis = betexplorer_direct.summarize_match_failure(None)

    assert diagnosis["status"] == "no_receipt"
    assert diagnosis["candidates"] == 0
    assert "receipt" in diagnosis["verdict"]


def _game() -> Game:
    return Game(
        game_id="kbl-1",
        league=League.SOUTH_KOREA_KBL,
        season="2026",
        home_team_id="66",
        away_team_id="50",
        home_name="WONJU DB PROMY",
        away_name="SEOUL SAMSUNG THUNDERS",
        scheduled_utc="2026-11-25T11:00:00Z",
        status=GameStatus.SCHEDULED,
        provider_ids={"kbl": "S47G01N49"},
    )


def test_a_zero_snapshot_run_is_named_in_the_report_and_a_gap():
    pipeline = Pipeline(
        PipelineConfig(league="SOUTH_KOREA_KBL", date_yyyymmdd="20261124"),
        InMemoryStore(),
        fetch=lambda url: {},
        allow_writes=False,
    )
    with patch.object(betexplorer_direct, "fetch_odds", return_value=[]), patch.object(
        betexplorer_direct, "last_receipt", return_value=RECEIPT
    ), patch.object(betexplorer_direct, "last_gaps", return_value=[]), patch.object(
        betexplorer_direct, "last_quota", return_value={}
    ), patch.object(betexplorer, "fetch_odds", return_value=[]):
        pipeline._acquire_book_lines([_game()], paid_sources_deferred=True)

    status = pipeline.book_provider_status["betexplorer_direct"]
    assert status["status"] == "fetched_no_snapshots"
    diagnosis = status["match_diagnosis"]
    assert diagnosis["reasons"]["ordered_team_pair_not_found_on_board_date"] == 1
    gap = next(g for g in pipeline.gaps if g.required_field == "book_lines")
    assert gap.provider == "betexplorer_direct"
    assert "board row(s)" in gap.detail


def test_a_cooldown_receipt_names_the_refusal_in_the_outer_status():
    """A 429-cooldown receipt must not render as ``fetched_no_snapshots``.

    Run 45385819819: NCAA_W / LITHUANIA_NKL / ISRAEL_LIGA_LEUMIT receipts all
    said ``http_429_cooldown`` while the outer status string stayed
    ``fetched_no_snapshots`` -- a refused read and a names-did-not-agree read
    were indistinguishable by status alone.
    """

    receipt = dict(RECEIPT)
    receipt["status"] = "http_429_cooldown"
    pipeline = Pipeline(
        PipelineConfig(league="SOUTH_KOREA_KBL", date_yyyymmdd="20261124"),
        InMemoryStore(),
        fetch=lambda url: {},
        allow_writes=False,
    )
    with patch.object(betexplorer_direct, "fetch_odds", return_value=[]), patch.object(
        betexplorer_direct, "last_receipt", return_value=receipt
    ), patch.object(
        betexplorer_direct, "last_gaps", return_value=[]
    ), patch.object(
        betexplorer_direct, "last_quota", return_value={}
    ), patch.object(betexplorer, "fetch_odds", return_value=[]):
        pipeline._acquire_book_lines([_game()], paid_sources_deferred=True)

    status = pipeline.book_provider_status["betexplorer_direct"]
    assert status["status"] == "http_429_cooldown"


def test_a_deferred_paid_source_does_not_switch_off_the_free_one():
    """The bug behind ``book_snapshots: 0`` on every fixtures-only lane.

    ``--skip-book-acquisition`` is set whenever the run holds no Odds API permit.
    It used to collapse the provider list to the direct fallback alone, so the
    free keyless ``betexplorer`` page that LITHUANIA_NKL, CHINA_WCBA_W,
    AUSTRALIA_WNBL_W and GREECE configure as their only price source was never
    called: five games stored, zero snapshots, therefore zero priced picks and
    nothing a ticket could ever be built from.
    """

    from unittest.mock import MagicMock

    pipeline = Pipeline(
        PipelineConfig(league="LITHUANIA_NKL", date_yyyymmdd="20261124"),
        InMemoryStore(),
        fetch=lambda url: {},
        allow_writes=False,
    )
    free = MagicMock(return_value=[])
    free.return_value = []

    with patch.object(betexplorer, "fetch_odds", free), patch.object(
        betexplorer, "last_quota", return_value={}
    ), patch.object(
        betexplorer_direct, "fetch_odds", return_value=[]
    ), patch.object(
        betexplorer_direct, "last_receipt", return_value=RECEIPT
    ), patch.object(
        betexplorer_direct, "last_gaps", return_value=[]
    ), patch.object(
        betexplorer_direct, "last_quota", return_value={}
    ):
        pipeline._acquire_book_lines([_game()], paid_sources_deferred=True)

    assert free.called, "the lane's own free price source must still be read"
    assert "theoddsapi" not in pipeline.book_provider_status


def test_the_deferral_names_the_paid_providers_it_skipped():
    """A lane that yields no priced pick must say which source was deferred."""

    pipeline = Pipeline(
        PipelineConfig(league="NBA", date_yyyymmdd="20261124"),
        InMemoryStore(),
        fetch=lambda url: {},
        allow_writes=False,
    )
    with patch.object(betexplorer_direct, "fetch_odds", return_value=[]), patch.object(
        betexplorer_direct, "last_receipt", return_value=RECEIPT
    ), patch.object(
        betexplorer_direct, "last_gaps", return_value=[]
    ), patch.object(
        betexplorer_direct, "last_quota", return_value={}
    ):
        pipeline._acquire_book_lines([_game()], paid_sources_deferred=True)

    deferred = pipeline.book_provider_status["deferred"]
    assert deferred["providers"] == ["theoddsapi"]
    # The direct fallback is free and was still attempted, so the lane's own
    # status survives beside the deferral instead of being replaced by it.
    assert "betexplorer_direct" in pipeline.book_provider_status
