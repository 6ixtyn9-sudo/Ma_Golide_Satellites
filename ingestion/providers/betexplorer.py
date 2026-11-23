"""BetExplorer aggregate-moneyline adapter.

Evidence status (see docs/PROVIDERS.md)
--------------------------------------------------
``LIVE_VERIFIED (live captures 2026-11-16; source terms unreviewed)`` applies
only to the server-rendered EuroLeague and CBA fixture tables captured under
``tests/fixtures/betexplorer/``. ``ABA_LEAGUE`` uses the same constrained
full-game moneyline adapter shape as a **probationary research-only route**:
its price policy remains aggregate/unresolved/probation and cannot supply a
named-book ticket, receipt, bank, ROI, or execution quote. Its route must not
be mistaken for a validated ABA schedule/result contract.

Verified conventions for the captured EuroLeague/CBA pages:

- One moneyline pair per match (header columns ``1``/``2``), decimal odds in
  ``data-odd`` attributes. Capped/empty odds cells are skipped — a missing
  quote stays missing and is never invented.
- The quoted price pair is the site's representative (average) odds across
  contributing bookmakers; per-bookmaker detail, opening/closing archives,
  and handicap/total markets load client-side (JS) and are NOT parsed here.
- Datetime cells are site-local ``Europe/Paris`` wall clock (CET/CEST).
  Verified against the official EuroLeague schedule on 2026-11-16:
  ``Tomorrow 17:00`` on the fixture page matched official
  ``2026-11-17T16:00:00Z``. A blank datetime cell repeats the previous
  row's datetime. Relative labels (``Today``/``Tomorrow``) resolve against
  the observation time; ``dd.mm.[yyyy] HH:MM`` is absolute with the year
  inferred when absent.
- Team display names differ from official spellings; the identity layer
  maps them through the canonical alias table. Names without an alias stay
  unmapped and quarantine (they never join by order or score).

Prices are observed at fetch time. This adapter deliberately skips rows
whose tip time has already passed; it does not archive post-tip quotes or
historical prices. ``pre_match`` is true only for a snapshot observed before
tip. Source terms have not been reviewed, so automated reuse or
redistribution is not cleared by this adapter.
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

from ..http import HttpClient
from ..identity import canonical_team_name, now_utc_iso, parse_utc, to_utc_iso
from ..models import (
    BookSnapshot,
    League,
    Market,
    OperationalValidation,
    Period,
    PriceProvenance,
    SourcePermission,
    payload_hash,
)
from . import IdentityError

PROVIDER = "betexplorer"
BOOK = "average"
SITE_TZ = "Europe/Paris"
EVIDENCE = (
    "LIVE_VERIFIED (fixtures pages and representative prices captured 2026-11-23; "
    "source terms unreviewed, prices are the site average, not a named book)"
)
ABA_EVIDENCE = "PROBATIONARY_RESEARCH_ONLY (no named-book, ticket, ROI, or execution use)"

BASE = "https://www.betexplorer.com/basketball"

#: Competition pages this parser reads. Every entry below the first three was
#: verified on 2026-11-23 by fetching the fixtures URL and matching its served
#: ``<h1>`` title, not by guessing the slug — the slugs are irregular
#: (``italy/lega-a``, ``turkey/super-lig``, ``japan/b-league-one``,
#: ``south-korea/kbl``, ``argentina/liga-a``), and ``turkey/tbl`` is the SECOND
#: tier despite matching the old top-flight name.
FIXTURES_URLS = {
    "EUROLEAGUE": BASE + "/europe/euroleague/fixtures/",
    "CBA": BASE + "/china/cba/fixtures/",
    "ABA_LEAGUE": BASE + "/europe/aba-league/fixtures/",
    "SPAIN_ACB": BASE + "/spain/acb/fixtures/",
    "ITALY_LBA": BASE + "/italy/lega-a/fixtures/",
    "TURKEY_BSL": BASE + "/turkey/super-lig/fixtures/",
    "FRANCE_LNB_PRO_A": BASE + "/france/lnb/fixtures/",
    "GERMANY_BBL": BASE + "/germany/bbl/fixtures/",
    "GREECE_BASKET_LEAGUE": BASE + "/greece/basket-league/fixtures/",
    "ISRAEL_PREMIER_LEAGUE": BASE + "/israel/super-league/fixtures/",
    "JAPAN_B_LEAGUE": BASE + "/japan/b-league-one/fixtures/",
    "SOUTH_KOREA_KBL": BASE + "/south-korea/kbl/fixtures/",
    "BRAZIL_NBB": BASE + "/brazil/nbb/fixtures/",
    "ARGENTINA_LIGA_NACIONAL": BASE + "/argentina/liga-a/fixtures/",
    "BASKETBALL_CHAMPIONS_LEAGUE": BASE + "/europe/champions-league/fixtures/",
}
#: ``BASKETBALL_AFRICA_LEAGUE`` is deliberately absent: its fixtures table serves
#: ``No matches found`` for the whole season (verified 2026-11-23), so registering
#: it would read as a healthy source that can never quote. Its RESULTS page does
#: carry odds, but this adapter reads fixtures, so the lane stays unpriced by
#: this provider rather than silently empty.
SUPPORTED_LEAGUES = frozenset(FIXTURES_URLS)

#: BetExplorer reports one representative price per market. It publishes the
#: fixture row's odds only in the days before tip-off (empty ``data-oid`` further
#: out) and uses ``1.00`` as a "no market" placeholder, which is never a price.
MIN_DECIMAL_ODD = 1.01


def evidence_for(league: League | str) -> str:
    """Return the competition-specific evidence label without overclaiming.

    The HTML parser is shared, but its EuroLeague/CBA live captures do not
    prove a separate ABA competition page or grant named-book evidence.
    """

    league_key = league.value if isinstance(league, League) else str(league).strip().upper()
    _require_league(league_key)
    return ABA_EVIDENCE if league_key == "ABA_LEAGUE" else EVIDENCE

_ROW_RE = re.compile(r"<tr>(?:(?!</tr>).)*?</tr>", re.S)
_LINK_RE = re.compile(
    r'<a href="/basketball/[^"]+/(?P<slug>[a-z0-9-]+)/(?P<match_id>[A-Za-z0-9]{8})/"'
    r' class="in-match">'
)
_SPANS_RE = re.compile(r"<span>(?:<strong>)?(?P<name>[^<]+?)(?:</strong>)?</span>")
_ODD_RE = re.compile(r'data-odd="(?P<odd>[0-9.]+)"')
_DATETIME_RE = re.compile(
    r'class="table-main__datetime"[^>]*>(?P<label>[^<]*)</td>'
)
_ABSOLUTE_RE = re.compile(
    r"^(?P<day>\d{1,2})\.(?P<month>\d{1,2})\.(?P<year>\d{4})?\s+(?P<hour>\d{1,2}):(?P<minute>\d{2})$"
)
_CLOCK_RE = re.compile(r"^(?P<today>Today|Tomorrow)?\s*(?P<hour>\d{1,2}):(?P<minute>\d{2})$")

TextFetch = Callable[[str], str]


def is_enabled() -> bool:
    return True  # keyless public pages


def last_quota() -> Dict[str, Any]:
    return {}


def _require_league(league: League | str) -> str:
    value = league.value if isinstance(league, League) else str(league).strip().upper()
    if value not in SUPPORTED_LEAGUES:
        raise ValueError(
            f"{PROVIDER} supports {sorted(SUPPORTED_LEAGUES)}, not {value!r}"
        )
    return value


def _text_fetch(fetch_factory: Optional[Callable[[Any], Any]], text_fetch: Optional[TextFetch]):
    """HTML needs raw text, not the pipeline's JSON-parsing fetch.

    The pipeline passes its ``fetch_factory`` (JSON ``get_json`` callables).
    Production therefore builds its own bounded ``HttpClient`` for text; a
    caller may inject ``text_fetch`` explicitly (tests use recorded HTML).
    """
    if text_fetch is not None:
        return text_fetch
    return HttpClient().get_text


def _parse_datetime(label: str, observed: datetime) -> Optional[datetime]:
    """Resolve one datetime-cell label to a site-local aware datetime.

    ``label`` is the raw cell text (``&nbsp;``/empty = repeat previous).
    Relative ``Today``/``Tomorrow`` resolve against the site-local date of
    ``observed``; absolute ``dd.mm.[yyyy] HH:MM`` infers the year only when
    absent (nearest year consistent with the observation date).
    """
    text = label.replace("&nbsp;", " ").strip()
    if not text:
        return None
    site_now = observed.astimezone(ZoneInfo(SITE_TZ))
    match = _ABSOLUTE_RE.match(text)
    if match:
        month = int(match.group("month"))
        day = int(match.group("day"))
        if match.group("year"):
            year = int(match.group("year"))
        else:
            year = site_now.year
            if (month, day) < (site_now.month, site_now.day):
                year += 1
        try:
            return datetime(
                year, month, day, int(match.group("hour")), int(match.group("minute")),
                tzinfo=ZoneInfo(SITE_TZ),
            )
        except ValueError:
            return None
    match = _CLOCK_RE.match(text)
    if match:
        day_offset = 1 if match.group("today") == "Tomorrow" else 0
        base = (site_now + timedelta(days=day_offset)).date()
        try:
            return datetime(
                base.year,
                base.month,
                base.day,
                int(match.group("hour")),
                int(match.group("minute")),
                tzinfo=ZoneInfo(SITE_TZ),
            )
        except ValueError:
            return None
    return None


def _rows(
    html: str, observed: datetime
) -> List[Tuple[Optional[datetime], str, str, str, str, Tuple[str, str]]]:
    """Parse fixture rows into (local_dt, match_id, slug, home, away, odds).

    Only rows carrying BOTH decimal quotes produce an odds tuple; rows with
    capped or absent prices keep an empty tuple (a missing quote stays
    missing). Datetime labels resolve against the observation time so a
    blank cell repeats the previous row's datetime and relative
    ``Today``/``Tomorrow`` labels stay deterministic for a pinned ``now``.
    """
    output: List[Tuple[Optional[datetime], str, str, str, str, Tuple[str, str]]] = []
    previous: Optional[datetime] = None
    for row in _ROW_RE.findall(html):
        link = _LINK_RE.search(row)
        if link is None:
            continue
        match_id = link.group("match_id")
        slug = link.group("slug")
        names = _SPANS_RE.findall(row)
        if len(names) != 2:
            continue
        cell = _DATETIME_RE.search(row)
        local_dt = None
        if cell is not None:
            local_dt = _parse_datetime(cell.group("label"), observed)
        if local_dt is None:
            local_dt = previous
        else:
            previous = local_dt
        if local_dt is None:
            continue
        odds = tuple(_ODD_RE.findall(row))
        if len(odds) != 2:
            odds = ()
        else:
            try:
                if any(float(value) < MIN_DECIMAL_ODD for value in odds):
                    odds = ()
            except ValueError:
                odds = ()
        output.append((local_dt, match_id, slug, names[0], names[1], odds))
    return output


def fetch_odds(
    league: str,
    date_yyyymmdd: str,
    fetch_factory: Optional[Callable[[Any], Any]] = None,
    now_utc: Optional[str] = None,
    tz_name: Optional[str] = None,
    text_fetch: Optional[TextFetch] = None,
) -> List[BookSnapshot]:
    """Moneyline snapshots for the slate date from the fixtures table.

    Rows whose site-local date differs from ``date_yyyymmdd`` are ignored.
    The league-local slate date and the site date coincide because the site
    publishes each league's fixtures in its CET/CEST wall clock, which is
    also the configured acquisition timezone for EuroLeague and the ledger
    timezone for CBA display.
    """
    league_key = _require_league(league)
    observed_text = now_utc or now_utc_iso()
    observed = parse_utc(observed_text)
    local_now = observed.astimezone(ZoneInfo(SITE_TZ))
    target = date_yyyymmdd.replace("-", "")
    fetch_text = _text_fetch(fetch_factory, text_fetch)
    html = fetch_text(FIXTURES_URLS[league_key])
    if not isinstance(html, str) or "table-main" not in html:
        raise IdentityError(
            f"{PROVIDER} {league_key}: fixtures page did not return a results table"
        )

    snapshots: List[BookSnapshot] = []
    for local_dt, match_id, slug, home, away, odds in _rows(html, observed):
        if not odds:
            continue
        if local_dt.strftime("%Y%m%d") != target:
            continue
        if local_dt <= local_now:
            # A slate-date row already past on the fixtures page carries
            # post-tip quotes; store nothing rather than a mislabeled price.
            continue
        scheduled_utc = to_utc_iso(local_dt)
        snapshots.append(
            BookSnapshot(
                game_id=f"{PROVIDER}:{match_id}",
                provider=PROVIDER,
                book=BOOK,
                market=Market.MONEYLINE,
                period=Period.FULL_GAME,
                observed_utc=observed_text,
                home_price_decimal=float(odds[0]),
                away_price_decimal=float(odds[1]),
                pre_match=True,
                main_line=True,
                payload_hash=payload_hash(
                    {"match_id": match_id, "slug": slug, "odds": odds}
                ),
                # BetExplorer's fixture-table value is explicitly an average
                # quote.  It is retained for research/provenance but cannot
                # become a named-book ticket price until source reuse and a
                # named book contract are independently cleared.
                price_provenance=PriceProvenance.AGGREGATE_CONSENSUS,
                source_permission=SourcePermission.UNRESOLVED,
                operational_validation=OperationalValidation.PROBATION,
                provider_event_id=match_id,
                home_name=canonical_team_name(home.strip()),
                away_name=canonical_team_name(away.strip()),
                scheduled_utc=scheduled_utc,
            )
        )
    return snapshots


def smoke(
    league: str,
    fetch_factory: Optional[Callable[[Any], Any]] = None,
    date_yyyymmdd: Optional[str] = None,
    now_utc: Optional[str] = None,
    tz_name: Optional[str] = None,
    text_fetch: Optional[TextFetch] = None,
) -> Dict[str, Any]:
    """Read-only coverage probe over the fixtures table (no writes)."""
    league_key = _require_league(league)
    observed_text = now_utc or now_utc_iso()
    observed = parse_utc(observed_text)
    local_now = observed.astimezone(ZoneInfo(SITE_TZ))
    target = (date_yyyymmdd or local_now.strftime("%Y%m%d")).replace("-", "")
    fetch_text = _text_fetch(fetch_factory, text_fetch)
    html = fetch_text(FIXTURES_URLS[league_key])
    rows = _rows(html, observed)
    dated = [row for row in rows if row[0] is not None]
    priced = [row for row in dated if row[5]]
    slated = [row for row in priced if row[0].strftime("%Y%m%d") == target]
    return {
        "provider": PROVIDER,
        "league": league_key,
        "verification": evidence_for(league_key),
        "rows_seen": len(rows),
        "rows_with_datetime": len(dated),
        "rows_with_prices": len(priced),
        "slate_date_rows_with_prices": len(slated),
        "note": (
            "representative average odds; per-bookmaker and handicap/total "
            "markets are client-side and not parsed"
        ),
    }


__all__ = [
    "ABA_EVIDENCE",
    "BOOK",
    "EVIDENCE",
    "FIXTURES_URLS",
    "PROVIDER",
    "SITE_TZ",
    "SUPPORTED_LEAGUES",
    "evidence_for",
    "fetch_odds",
    "is_enabled",
    "last_quota",
    "smoke",
]
