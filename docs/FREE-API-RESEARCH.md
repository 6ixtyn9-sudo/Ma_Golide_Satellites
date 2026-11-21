# Free basketball-data provider policy and research

## Policy

Ma Golide accepts only providers with a documented, ongoing **$0 tier** for
its optional authenticated integrations. A trial that expires or a source whose
production use requires a paid contract is not an eligible provider under this
policy. A free key does not by itself make a source ticket eligible: the source
must still pass rights, live-schema, identity, coverage, freshness, and
operational-validation review.

The repository keeps three optional authenticated providers under this policy:

| Provider | Credential | Ongoing free allowance | Current role |
| --- | --- | --- | --- |
| The Odds API | `THE_ODDS_API_KEY` | 500 credits/month | Direct named-book odds for configured NBA/WNBA/NCAA lanes, subject to sport/region/market coverage |
| SportsGameOdds | `SPORTSGAMEODDS_API_KEY` | 2,500 objects/month | NBA full-game and quarter totals |
| Predixsport | `PREDIXSPORT_API_KEY` | 200 shared calls/month for personal, non-commercial use | Optional NBA/EuroLeague prediction input; still documentation-only pending an authenticated schema smoke |

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
| [OddsPapi](https://oddspapi.io/blog/free-sports-data-api/) | 250 requests/month | A very-low-volume schedule/odds smoke across international books | Not suitable for the 30-minute all-lane workflow at this quota. Require direct terms, live key, and named-book identity verification before adding an adapter. |
| Existing The Odds API / SportsGameOdds | See policy table above | Current approved free-tier integration paths | Retain and test first; no new adapter is needed to diagnose their real account coverage. |

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
