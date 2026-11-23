# Free basketball-data provider policy and research

## Policy

Ma Golide accepts only providers with a documented, ongoing **$0 tier** for
its optional authenticated integrations. A trial that expires or a source whose
production use requires a paid contract is not an eligible provider under this
policy. A free key does not by itself make a source ticket eligible: the source
must still pass rights, live-schema, identity, coverage, freshness, and
operational-validation review.

The repository keeps two optional authenticated providers under this policy:

| Provider | Credential | Ongoing free allowance | Current role |
| --- | --- | --- | --- |
| The Odds API | `THE_ODDS_API_KEY` | 500 credits/month advertised allowance; 360 local routine-reservation cap | Direct named-book full-game `h2h,totals,spreads` for configured NBA/WNBA/NCAA lanes, subject to bounded credentialed smoke and sport/region/bookmaker coverage |
| Predixsport | `PREDIXSPORT_API_KEY` | 200 shared calls/month for personal, non-commercial use | Optional NBA/EuroLeague prediction input; still documentation-only pending an authenticated schema smoke |

The documented NCAA women's sport key is `basketball_wncaab` (not
`basketball_ncaaw`). A current-odds request is **not** universally one credit:
provider cost is requested markets times effective bookmaker-region groups.
Routine collection is consequently a date-bounded, named-DraftKings group
request for three full-game markets, with a maximum planned cost of three
credits. Native period and historical requests have different/higher documented
cost paths and are outside the routine free-tier contract. The shared key is
never exposed to manual workflow dispatch or to an unpermitted matrix lane.

Trial-only or contract-only integrations are deliberately absent:

- iSportsAPI: temporary evaluation access followed by paid plans.
- Sportradar: trial/evaluation access followed by a commercial contract.

They are not silently replaced by a quote from a different league or an
unverified scraper. Missing coverage remains a visible CoverageGap.

## Free-source search, 2026-11-18

The research was deliberately multilingual. Search phrases included:

- English: `free basketball API schedule results odds`.
- Spanish: `API gratis baloncesto calendario resultados cuotas`.
- Portuguese: `API gratuita basquete calendario resultados odds`.
- French: `API gratuite basket calendrier résultats cotes`.
- Russian: `бесплатный API баскетбол расписание результаты коэффициенты`.
- Chinese: `免费 篮球 数据 API 赛程 比分 赔率`.

Candidate providers below are discovery leads, **not** approved integrations.
No provider may be added to the registry or promoted to ticket evidence until a
free-key smoke records the exact live schema, competition coverage, source
terms, rate limits, event identity, and result/price reconciliation behavior.
The [league readiness and source-adoption roadmap](LEAGUE-ADOPTION-ROADMAP.md)
sets the implementation order: candidate discovery first records a source role
and evidence contract, but does not grant a secret, an adapter, an automated
request, or ticket-price authority.

### International discovery leads that were not promoted

| Lead | Discovery context | Decision |
| --- | --- | --- |
| [AllSports API basketball](https://allsportsapi.com/basketball-api) | Found through multilingual API-list searches; its advertised free access limits a key to two randomly selected leagues. | Excluded from a multi-league canonical workflow: its free allowance cannot guarantee the requested competition or a repeatable fixture lifecycle. |
| [Rinuo free sport API listing](https://www.rinuo.com/free/sport) | Chinese-language free-sport API directory lead. | Directory listing is not proof of an ongoing provider tier, source rights, or basketball schema; retain as a search lead only. |

### Schedule/result candidates

| Candidate | Advertised free allowance | Advertised basketball scope | Current decision |
| --- | --- | --- | --- |
| [API-Basketball](https://www.api-basketball.com/) | 100 requests/day | Worldwide competitions and endpoints | Candidate for a bounded schedule/result smoke. Do not assume free odds access or rights for unattended publication. |
| [SportAPI](https://sportapi.ai/) | 1,000 requests/day; 7,000/month | NBA, WNBA, NCAA, NBA G League, EuroLeague, EuroCup, and other basketball coverage | Candidate for a key-backed comparison against existing official/public lanes. Needs identity and official-result verification. |
| [Highlightly Basketball API](https://highlightly.net/basketball-api/) | 100 requests/day | Advertises 340+ leagues including ACB, BBL, LNB, LBA, Greece, Australia, Japan, CBA, KBL, NBB, NCAA, and FIBA competitions | Best broad schedule/result research lead for currently unsupported lanes. Its free-key response and source rights must be tested before any production use. |
| [API Ninjas Basketball](https://api-ninjas.com/api/basketball) | Current-season free plan | NBA, WNBA, NCAA men, NCAA women schedules/finals/standings | Redundant schedule/result cross-check candidate only; it does not solve named-book odds. |
| [BasketballAPI.net](https://basketballapi.net/documentation) | 100 requests/day | Current-day/yesterday results and current competitions | Too limited for a future-fixture lifecycle; retain only as a possible result-reconciliation comparison. |

### Odds candidates

| Candidate | Advertised free allowance | Potential use | Current decision |
| --- | --- | --- | --- |
| [OddsPapi](https://oddspapi.io/blog/free-sports-data-api/) | 250 requests/month | A very-low-volume schedule/odds smoke across international books | Not suitable for the hourly all-lane workflow at this quota. Require direct terms, live key, and named-book identity verification before adding an adapter. |
| Existing The Odds API | See policy table above | Current approved free-tier integration path | Retain and test first; no new adapter is needed to diagnose its real account coverage. |

The multilingual search produced several directories and unofficial examples,
but no additional primary-source, permanently-free, direct named-book basketball
API that can safely be promoted without a free-key smoke. In particular, a
Portuguese implementation report notes that API-Basketball's free odds endpoint
was unavailable and uses The Odds API for price data instead. That is discovery
evidence only, not an approval of either provider for a new lane.

## Next validation sequence

1. Configure only the permanently-free keys as GitHub Actions secrets.
2. Run one bounded, redacted smoke per existing provider and record provider,
   league, region, bookmaker, market, source timestamp, quota, and event match
   status. Never log a key.
3. Obtain a free key for one schedule/result candidate, beginning with
   Highlightly for broad global-lane coverage or API-Basketball for a compact
   comparison. Capture a fixture and an official final before writing an
   adapter.
4. Evaluate each candidate against a league-specific source contract. A generic
   global API cannot be labelled an official result source merely because it
   returns a score.
5. Add a candidate only after hermetic parser fixtures and a documented live
   smoke prove the contract. Preserve price-source reuse permission and
   operational-validation status as audit metadata; aggregate consensus prices
   remain research-only, while named-book paper candidates still require every
   model, calibration, and ROI gate.

## Price avenues beyond BetExplorer — forge and package reconnaissance (2026-11-23)

The question asked was whether GitLab offers a lower-maintenance bootstrap price
avenue than the GitHub projects already surveyed. The reconnaissance below was
run from a sandbox whose only reachable registries are PyPI, npm and GitHub;
GitLab's public API was read through the fetch proxy only (no direct egress), so
per-project *source* inspection there was not possible — only metadata.

### GitLab: no usable project found

`/api/v4/projects` searches for `basketball odds`, `betting`, `bookmaker`,
`sports odds` and `basketball`, sorted by `last_activity_at`, returned no
sports-betting or odds-feed project: results were poker cheat sheets, a
RuneScape launcher, an unrelated odds-named signal repo, and similar. GitLab's
public corpus is not where this class of project lives.

More important than the search result: **a forge does not change the
maintenance cost.** What makes a bootstrap price avenue expensive here is
scraping an undocumented HTML surface and re-verifying it every time the host
changes. A scraper hosted on GitLab instead of GitHub is the same scraper. The
categories that actually differ in maintenance are:

| Class | Maintenance | Provenance it can honestly claim |
|---|---|---|
| League's own feed | low — but none of these leagues publishes a price (probe 2026-11-23: B.League's S3 payload has no odds field of any name; ACB's match-header has none on a future fixture) | `no_price` |
| Documented commercial API with a key | low: versioned REST, published error model | named book or aggregate, per its licence |
| Exchange API (documented) | low: versioned, authenticated, no HTML | `exchange` (added 2026-11-23, see below) |
| HTML scrape of an aggregate | high: layout changes, no contract, no rate limit published | `aggregate_consensus` |

### Package ecosystem: only one avenue is alive

Checked against PyPI and npm on 2026-11-23 (`pypi.org/pypi/<name>/json`,
`registry.npmjs.org/<name>`):

| Package | Registry | Latest | Last release | Licence | Verdict |
|---|---|---|---|---|---|
| `betfairlightweight` | PyPI | 2.24.0 | 2026-09-04 | MIT | **alive** — wrapper for the documented Betfair API-NG; the only candidate whose maintenance is the upstream's problem, not ours |
| `sportsdataverse` | PyPI | 0.1.4 | 2026-10-26 | unspecified | alive; already the source of our route contracts, not a price feed |
| `oddsapi` | PyPI | 1.2.4 | 2021-10-15 | unspecified | dead; The Odds API is already integrated directly |
| `betfair` | PyPI | 0.1.1 | 2014-12-30 | unknown | dead |
| `the-odds-api` | npm | 2.1.0 | 2019-06-04 | — | dead |
| `betfair-api` | npm | 0.2.0 | 2014-12-02 | — | dead |
| `pinnacle-api` | npm | 0.0.0 | 2017-08-10 | — | dead |

`oddspapi` does not exist on PyPI (404) under that spelling.

### What changed in the repository

`PriceProvenance.EXCHANGE` now exists. Before it, an exchange quote had to be
recorded as `named_book`, `aggregate_consensus` or `no_price`, all wrong — and
`named_book` is exactly what `ticket_price_eligible` trusts, so the
unknown-provider fallback would have laundered an exchange offer into the
paper-ticket lane. `price_evidence_policy` now returns an exchange policy (with
`acquire=False`, unresolved permission, unverified validation) for `betfair` /
`betfair_exchange` identities **even when the league has not registered them**.

No adapter was written, because no authorised capture exists: writing one from
the package's README would be an invented API shape, and the exchange's terms
for automated single-user reuse have not been reviewed.

### The one probe worth running

Betfair API-NG with a free delayed app key, one bounded catalogue read per
league, recording: market present/absent, `marketStartTime`, number of
runners, price range, and whether the league appears at all — basketball's
event type id is 7522. Coverage is the open question, not documentation: if
ACB / LBA / B.League / KBL / BCL are not listed, this avenue costs nothing but
also buys nothing, and the honest label for those lanes stays "no market".

### Bot protection is a boundary, not an obstacle

Cloudflare/CloudFront challenges on league and bookmaker sites are the reason
the "alt lane" work is slow, and the correct response is to stop treating them
as something to get around. Three reasons, in order of weight:

1. **It is somebody's access control.** This project's evidence model already
   carries `SourcePermission`; a route that has to defeat a bot challenge is
   not a source we can label `CLEARED`, and a headless browser that defeats one
   is a deliberate circumvention, not an acquisition method.
2. **It does not survive.** A scraper that works today behind a rotating
   challenge is the highest-maintenance thing in the repository — the exact
   complaint that started this section.
3. **It is not where our gaps are.** The lanes that looked blocked by
   Cloudflare mostly are not: B.League serves its HTML **and** its keyless S3
   contexts, FIBA GDAP serves whole seasons on one call, KBL answers its window
   route, ACB answers both routes with the site-shipped `X-APIKEY`, and
   BetExplorer answered every probe on 2026-11-23 with real `data-odd` values
   for played games. What is missing is *prices for unopened rounds* and
   *coverage of specific leagues* — a calendar and licensing problem, not a
   challenge-response problem.

So the usable class is unchanged and narrow: documented API, documented
exchange, or a licensed aggregate. `www.acb.com` being 403 behind CloudFront
costs nothing, because the price we would have scraped there is not ours to
take and not what the ACB API was asked for.

### Settling provider coverage with the key already held

`scripts/probe_odds_api_coverage.py` reads The Odds API's documented
`GET /v4/sports` catalogue with the existing `THE_ODDS_API_KEY`, writes the raw
body plus `LEDGER.tsv` (timestamp, slug, redacted URL, status, bytes, sha256),
and classifies every basketball key against the registry:

- `already mapped` — the provider lists it and `SPORT_KEYS` uses it;
- `listed, unmapped` — the provider carries the league and the registry has not
  admitted it: needs a credentialed odds smoke and a price policy, not an
  adapter guess;
- `listed, no league` — a basketball competition we do not trade;
- absent from the table — the only reading that means "not carried".

One request, no per-league guessing, and an empty body is reported as "do not
read this as no coverage" rather than as a negative.
