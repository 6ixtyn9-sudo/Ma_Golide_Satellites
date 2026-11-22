# League route contracts: verified leads and blockers

This page records what has actually been verified for each lane that does not
yet have an adapter, so the remaining work is mechanical rather than
speculative. A route lead is **not** an integration: nothing here changes a
registry entry, and a lane keeps `source_unavailable` until its adapter,
identity checks, and fixtures exist and pass.

Hosts and endpoint shapes below were read from published documentation and
third-party client source, not from a Ma Golide capture. Every future adapter
must stay labelled `referenced route contract` until a first-party capture
exists, exactly as ABA, EuroCup and ACB do.

## Ready pattern: a documented period-level route exists

| Lane | Host | Documented routes | Notes / blocker |
| --- | --- | --- | --- |
| Spain ACB | `api2.acb.com` | `seasondata/Calendar/season-calendar?season={end_year}`, `matchdata/MatchHeader/match-header?matchId=` | **Implemented** in `ingestion/providers/acb.py`. |
| Italy LBA | `www.legabasket.it/api` | `championships/get-championships?cs_id=&items=&s=` and further championship sub-routes | Season-scoped championship listing; match-level route and period fields still need pinning down. |
| Japan B.League | `b-league.s3.amazonaws.com/web_json` | `contexts/{schedule_key}` | Static JSON on S3; the schedule key discovery step is the outstanding unknown. |
| South Korea KBL | `api.kbl.or.kr` | `season/list` and additional season routes | Simplest host of the group; needs a game/period route and team identity. |
| Australia NBL | module exists in `sportsdataverse-py` (`nbl_au`) | file path and routes not confirmed from documentation | Treat as unverified; The Odds API NBL key stays probationary and never supplies fixtures. |

## No packaged adapter found

These have no maintained open-source period-level route. Each needs its own
site contract (calendar + period scores + stable team identity) or a
browser-automation fallback, which this project has ruled out as a primary
route.

- Basketball Champions League (FIBA family; shares contract style with the
  validated FIBA EuroLeague Women adapter, but its competition route is
  unverified).
- France LNB Pro A, Germany BBL, Turkey BSL, Greece Basket League, Israel
  Premier League, Brazil NBB, Argentina Liga Nacional.

## Evidence that no pricing work is needed first

Every one of the lanes above already registers `betexplorer_direct` as a price
source, and for most it is the only one. Pricing was never the blocker: the
missing half is fixtures plus reconciled periods. BetExplorer itself cannot
close that gap — its adapter parses day boards and full-game moneylines/lines
but no quarter data, so it can never be a result source.

## Order of work

1. Pin the LBA championship and match routes; period fields from the
   championship/box routes.
2. KBL season and game routes (smallest surface of the remaining four).
3. B.League schedule-key discovery against the S3 context index.
4. NBL: confirm whether a documented route exists at all before any work.
5. FIBA family (Champions League) once the ELW adapter's proven pattern is
   generalised.
6. The seven domestic leagues as individual contracts; do not batch them
   behind a generic scraper, because each needs its own identity crosswalk and
   its own failure semantics.
