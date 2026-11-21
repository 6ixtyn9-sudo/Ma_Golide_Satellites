# External prediction and named-book price actions

**Purpose:** execute the remaining external work needed to give each configured
basketball lane a lawful opportunity to collect frozen pre-tip prediction and
named-book moneyline evidence.

This is an evidence and outreach pack. It does **not** create a new
paper-ticket lane or relax any existing forecast, calibration, identity,
pricing, or no-bet gate.

## Credential correction

`THE_ODDS_API_KEY` and `PREDIXSPORT_API_KEY` are already configured as GitHub
repository secrets.  They are not outstanding vendor-onboarding tasks.

- The workflow injects `PREDIXSPORT_API_KEY` into each registered league runner.
  The current technical work is to make its authenticated NBA/EuroLeague list
  capture observable, bounded to one request per league run, and identity-safe.
- The workflow injects `THE_ODDS_API_KEY` only after an automated,
  scheduler-issued The Odds API permit.  Manual dispatch deliberately receives
  neither the permit nor the key.  This preserves the three-credit request cap
  and makes current price capture an authorised scheduled action rather than a
  manual probe.

Accordingly, do **not** send new credential/onboarding enquiries to either
provider.  The remaining external enquiries concern expansion providers and
source rights for currently unpriced leagues.

## Current repository facts

The repository contains 24 configured basketball leagues.  The source work here
covers every one:

```text
NBA, WNBA, EuroLeague, EuroLeague Women, Chinese CBA, NCAA men, NCAA women,
EuroCup, Basketball Champions League, NBA G League, Basketball Africa League,
Australia NBL, Spain ACB, France LNB Pro A, Germany BBL, Italy LBA, Turkey BSL,
Greece Basket League, ABA League, Israel Premier League, Japan B.League,
South Korea KBL, Brazil NBB, Argentina Liga Nacional.
```

The currently configured direct named-book adapter is deliberately limited to
NBA, WNBA, NCAA men, and NCAA women.  It is permit-gated and remains
`DOCUMENTATION_ONLY` until a credentialed, bounded smoke proves actual account,
region, bookmaker, event identity, and timestamp coverage.  Australia NBL is a
documented but probationary mapping.  EuroLeague is deliberately *not* enabled
in the adapter.

The current Predixsport adapter is also `DOCUMENTATION_ONLY`: no authenticated
response has been retained, its free tier is not a production right, and
published documentation/schema is not a substitute for a source contract.

The active ABA League results backfill is separate from this work.  More
reconciled results improve empirical diagnostics, but cannot turn aggregate
prices or reconstructed predictions into frozen, named-book evidence.

## What was independently confirmed

The following statements are limited to the cited provider documentation.  They
identify viable contract conversations, not approved integrations or observed
coverage for a particular account or fixture.

### 1. Existing direct price provider: The Odds API

- Its official historical endpoint returns bookmaker-odds snapshots as of a
  requested timestamp; its documentation describes per-market `last_update`
  and paid historical access.
- Its published historical-coverage table includes `basketball_euroleague`
  beginning 2020-09-24 and `basketball_nbl` beginning 2024-09-24.  Historical
  availability begins only when a sport, bookmaker, and market was added.
- Its terms explicitly permit indefinite storage, research/analytical work,
  derived values, and ML training, while prohibiting resale of the raw feed.

Official sources:

- <https://the-odds-api.com/liveapi/guides/v4/>
- <https://the-odds-api.com/historical-odds-data/>
- <https://the-odds-api.com/sports-odds-data/sports-apis.html>
- <https://the-odds-api.com/terms-and-conditions.html>

**Repository decision:** this is already a bootstrapped price route for the
four configured permit-eligible lanes. Its next work is operational evidence:
let the existing scheduled permit path retain real account/book/region/fixture
observations. EuroLeague remains deliberately disabled and Australia NBL
remains probationary; neither may be enabled merely because the key exists.

### 2. Broad named-book price lead: OpticOdds

- Its documented league list includes EuroCup, EuroLeague Women, EuroLeague,
  and NBA G League.
- Its odds API requires one or more named `sportsbook` values.  Its sportsbook
  endpoint exposes supplier identities and active status.
- Its documentation describes a historical endpoint returning price-change,
  lock, unlock, and settlement events with timestamps.

Official sources:

- <https://developer.opticodds.com/docs/leagues>
- <https://developer.opticodds.com/reference/get_sportsbooks>
- <https://developer.opticodds.com/reference/get_fixtures-odds>
- <https://developer.opticodds.com/docs/odds-api-getting-started-guide>
- <https://developer.opticodds.com/reference/odds>

**Repository decision:** a potentially useful contract lead for the unpriced
European/G League lanes, but no implementation or data request is authorised.
The contract must prove the exact target league/book combination, permitted
storage and derived-publication use, the historical retention window, and that
records are fixed-odds bookmaker quotes rather than an exchange, consensus, or
OpticOdds AI value.

### 3. Broad independent prediction lead: Sportradar Global Basketball

- The documented Global Basketball API has a Season Probabilities endpoint for
  pre-match home/away win probabilities.
- Its sport-event probability endpoint documents pre-match and live
  probabilities; the live material exposes a `last_updated` timestamp.
- The documentation is a production/trial API contract, not a free public data
  licence.  It does not, by itself, establish that every configured league is
  covered, that historical probability snapshots are available, or that retained
  publication/research use is allowed.

Official sources:

- <https://developer.sportradar.com/basketball/reference/global-basketball-overview>
- <https://developer.sportradar.com/basketball/reference/global-basketball-season-probabilities>
- <https://developer.sportradar.com/basketball/reference/global-basketball-sport-event-probabilities>
- <https://developer.sportradar.com/basketball/reference/global-basketball-live-probabilities>
- <https://developer.sportradar.com/odds/reference/probabilities-faq>

**Repository decision:** this is the strongest single B2B candidate for
provider predictions across multiple leagues.  It needs a coverage matrix and
written answers about immutable pre-tip timestamps, provider/model versioning,
historical archives, and our retention/derived-output rights before any code.

### 4. Broad historical/prediction lead: Stats Perform / Opta

Stats Perform publicly advertises live and historical data/API delivery,
basketball coverage, and predictive models.  The public pages do not prove an
exact league-by-league probability feed, original forecast timestamps, or the
rights needed for this project.  Those are contract questions, not assumptions.

Official sources:

- <https://www.statsperform.com/products/opta-for-prediction-markets/>
- <https://www.statsperform.com/stats-perform-faqs-apis-and-data-delivery/>
- <https://developers.statsperform.com/stats-perform-products>
- <https://developers.statsperform.com/historical-sports-data-for-pricing-models>

**Repository decision:** pursue in parallel with Sportradar, not as evidence of
coverage or permission.  It is the main alternative if Sportradar cannot
provide an immutable pre-tip forecast archive for the required competitions.

### 5. Narrow current prediction lead: Predixsport

Predixsport publicly presents NBA and EuroLeague prediction content.  The
repository already isolates an authenticated adapter for those two leagues, but
has never captured a production response.  Public pages can establish neither a
commercial retention right nor that a historical response is the exact forecast
published before a particular tip.

Official sources:

- <https://www.predixsport.com/>
- <https://www.predixsport.com/nba_predictions>
- <https://www.predixsport.com/api-access>
- <https://www.predixsport.com/terms>

**Repository decision:** the NBA/EuroLeague credential path is already
bootstrapped. Capture one authenticated upcoming list per relevant league run,
join only exact provider-ID or home/away/UTC-tip matches, and retain matching
pre-tip observations prospectively. Do not treat later-recomputed values or
date-only archives as frozen historical assay evidence.

### 6. ABA League and Japan B.League rights paths

- ABA League publishes a corporate contact (`info@aba-liga.com`) and identifies
  itself as owner of associated competition rights.  Its official news also
  records an exclusive data partnership with Genius Sports.  That makes direct
  permission and/or referral to the official data-rights holder the correct
  route; it is not a licence to automate its public website.
- Data Stadium explicitly sells B.League research/academic data, including
  B1/B2 historical basketball data, and describes a contract process.  Its
  real-time delivery page states that basketball data can be delivered by XML,
  API, or email.  This is a legitimate Japanese-language contract path, not a
  Snowflake sample, public-page, or private-endpoint permission.

Official sources:

- <https://www.aba-liga.com/company-info/>
- <https://www.aba-liga.com/newslist/Official/14/>
- <https://datastadium.co.jp/academic-data>
- <https://datastadium.co.jp/service-detail/PoFfD7Uf>
- <https://datastadium.co.jp/contact>

**Repository decision:** send the ABA and Data Stadium inquiries in parallel.
Neither source is approved for acquisition, prediction, or named-book pricing
until it supplies written scope and terms.

## Required contract answers

These questions apply to a **new** provider, source role, or league expansion.
They do not block the two already-bootstrapped secret paths; those require real
scheduled observations under their existing narrow repository policies.

### Prediction evidence

1. List exact covered competitions from the 24-league list above.
2. State the forecast fields: event identifier, home/away teams, scheduled tip
   time in UTC, win probabilities, provider/model name and model version.
3. Supply an immutable provider-created `published_at` or `generated_at` for
   each forecast.  A client retrieval time alone is insufficient.
4. Confirm forecasts are available before tip and identify update/revision
   behaviour.
5. For historical backfill, confirm the archive contains the value originally
   published before tip, its original timestamp, and no post-result
   recomputation.  If it cannot say that, the archive is not frozen pre-tip
   evidence.
6. Confirm rights for automated collection, durable storage, internal research,
   calibration, and publication of derived non-betting research output.

### Named-book moneyline evidence

1. List exact covered competitions, regions, and named bookmakers for each.
2. Confirm a record preserves the actual bookmaker ID/name, source fixture ID,
   home/away selection, full-game moneyline market, offered price, market state,
   and provider/book timestamp.
3. Confirm fixed-odds bookmaker records can be filtered from exchanges,
   prediction markets, consensus/best-price values, and provider-derived lines.
4. Confirm whether historical records preserve bookmaker identity and a pre-tip
   timestamp.  Ask for start dates and granularity by league/book/market.
5. Confirm the same collection, retention, internal research, and derived-output
   rights.  A non-production trial is not an answer.

## Ready-to-send inquiries

The account owner must send the remaining external enquiries through their
normal business channel; this workspace has no connected mail application.
The wording intentionally says "evaluating" rather than falsely claiming that
a new live production feed is already in use. The Odds API and Predixsport are
explicitly excluded here because their credential paths are already bootstrapped.

### A. Sportradar — global basketball probabilities and archive

**Subject:** Licensing enquiry — pre-tip basketball probability feed and immutable historical snapshots

```text
Hello Sportradar Sales,

We are evaluating a non-commercial, paper-only basketball research system. We
need provider-generated pre-tip win probabilities for the following
competitions:

NBA, WNBA, EuroLeague, EuroLeague Women, Chinese CBA, NCAA men, NCAA women,
EuroCup, Basketball Champions League, NBA G League, Basketball Africa League,
Australia NBL, Spain ACB, France LNB Pro A, Germany BBL, Italy LBA, Turkey BSL,
Greece Basket League, ABA League, Israel Premier League, Japan B.League,
South Korea KBL, Brazil NBB, and Argentina Liga Nacional.

Your Global Basketball documentation describes season and sport-event
probabilities. Before considering a contract, please confirm:

1. Which of the listed competitions are included in the production
   Probabilities package;
2. Whether each pre-match probability record carries a provider-created,
   immutable generated/published timestamp, a source event ID, scheduled UTC
   tip time, home/away mapping, and provider/model version;
3. Whether historical probability snapshots are available and, if so, whether
   they are the originally published pre-tip values rather than later model
   recomputations;
4. The permitted scope for automated capture, long-term internal retention,
   calibration, and publication of our own derived non-commercial research;
5. Available trial, production, historical-data, rate-limit, and pricing terms.

We do not seek exchange prices, betting execution, consumer odds display, or
raw-feed redistribution. Thank you.
```

### B. Stats Perform / Opta — coverage, forecasts, and historical forecast provenance

**Subject:** Data licensing enquiry — basketball pre-match predictions and historical forecast provenance

```text
Hello Stats Perform team,

We are evaluating licensed Opta data and predictive products for a
non-commercial, paper-only basketball research project. We need a lawful route
for pre-tip winner probabilities and, where available, original historical
forecast snapshots.

Could you provide a coverage matrix for the following competitions: NBA, WNBA,
EuroLeague, EuroLeague Women, Chinese CBA, NCAA men, NCAA women, EuroCup,
Basketball Champions League, NBA G League, Basketball Africa League, Australia
NBL, Spain ACB, France LNB Pro A, Germany BBL, Italy LBA, Turkey BSL, Greece
Basket League, ABA League, Israel Premier League, Japan B.League, South Korea
KBL, Brazil NBB, and Argentina Liga Nacional?

For each covered competition, please confirm:

1. Whether Opta Predictions supplies pre-match home/away winner probabilities;
2. Event ID, team mapping, scheduled UTC tip, provider/model revision, and
   immutable publication/generation timestamp fields;
3. Whether a historical archive contains forecasts as originally issued before
   tip-off, rather than post-event recalculations;
4. Rights for automated retrieval, durable internal storage, calibration, and
   publication of our own derived non-commercial analysis; and
5. Commercial model, minimum term, historical-data option, and API delivery
   details.

We do not request raw-feed redistribution, betting execution, or odds display.
Thank you.
```

### C. OpticOdds — named fixed-odds books, historical archive, and rights

**Subject:** Licensing enquiry — named-book pre-game moneyline and historical archive for basketball

```text
Hello OpticOdds team,

We are evaluating your API for a non-commercial, paper-only basketball research
project. We need pre-game full-game moneyline records from actual named
fixed-odds bookmakers, not exchanges, prediction markets, best-price values,
consensus prices, or provider-derived lines.

Please provide a coverage matrix for these competitions: NBA, WNBA, EuroLeague,
EuroLeague Women, Chinese CBA, NCAA men, NCAA women, EuroCup, Basketball
Champions League, NBA G League, Basketball Africa League, Australia NBL, Spain
ACB, France LNB Pro A, Germany BBL, Italy LBA, Turkey BSL, Greece Basket League,
ABA League, Israel Premier League, Japan B.League, South Korea KBL, Brazil NBB,
and Argentina Liga Nacional.

For each supported competition, please confirm:

1. Named fixed-odds sportsbooks available and whether their ID/name is retained
   per quote;
2. Availability of the full-game moneyline market, home/away selection mapping,
   event ID, scheduled UTC tip, odds-update timestamp, and market status;
3. Historical archive coverage by league/book/market, timestamp granularity,
   and whether original bookmaker attribution is retained;
4. Whether exchanges and prediction-market sources can be excluded completely;
5. Production licence terms for automated collection, durable internal storage,
   historical backfill, calibration, and publication of derived non-commercial
   research; and
6. Pricing, rate limits, supported regions, and a contract contact.

We will not redistribute the raw feed or use it to place bets. Thank you.
```

### D. Existing The Odds API and Predixsport paths — no outreach required

Both credentials are already bootstrapped in repository secrets. The immediate
work is to retain and inspect evidence from the existing workflow:

- The Odds API stays limited to its scheduled, permit-gated routine lanes. A
  secret alone does not promote EuroLeague or Australia NBL.
- Predixsport is fetched once per NBA/EuroLeague run. Its rows must join exactly
  before they are retained as forecast observations; unmatched rows are visible
  coverage gaps rather than guessed fixture mappings.
- A historical-provider backfill request remains separate from prospective
  capture. It must not be added unless the provider explicitly supplies
  original pre-tip snapshots and their immutable timestamps.

### E. ABA League — official data-rights inquiry

**To:** `info@aba-liga.com` (copy `headoffice@aba-liga.com` if appropriate)

**Subject:** Request for official ABA League data rights and licensed feed contact

```text
Dear ABA League team,

We operate a non-commercial, paper-only basketball research project and would
like to obtain written permission before automating any use of ABA League
fixture, result, or statistical data.

Could you please advise whether ABA League or its official data-rights partner
licenses a feed for:

- fixture identity and scheduled tip times;
- final scores and period scores;
- historical seasons and corrections/postponements; and
- lawful automated retrieval, durable internal research retention, and derived
  non-commercial publication.

If a data-rights partner manages this scope, please provide the correct
commercial contact and confirm whether the partner can license the above use.
We are not asking to scrape or reuse the public website without permission.

Thank you for your guidance.
```

### F. Data Stadium — B.League official/research data inquiry (Japanese)

**Subject:** B.LEAGUEデータの研究利用・自動取得に関するお問い合わせ

```text
データスタジアム株式会社 ご担当者様

南アフリカで、非営利・ペーパー運用のバスケットボール研究プロジェクトを運営しています。
B.LEAGUEの公開サイトや非公開エンドポイントを無断で取得する意図はなく、正式な契約に基づく
データ提供について相談したく、ご連絡しました。

以下の用途に利用可能なB.LEAGUE（B1・B2を含む）のデータ商品またはAPI／XML配信はありますでしょうか。

・試合ID、対戦カード、開始時刻（UTCを含む）
・試合結果、クォーター別スコア、延期・中止・訂正情報
・過去シーズンの履歴データ
・研究目的での自動取得、長期保存、モデル検証、および自ら作成した非営利の分析結果の公開

可能な場合、対象シーズン・リーグ、データ項目、提供形式、利用許諾範囲、保存・公開条件、
料金、レート制限、契約手続をご案内ください。

オッズ、賭けの執行、データの再販売を目的とするものではありません。
よろしくお願いいたします。
```

**English reference translation:**

```text
Hello Data Stadium team,

We operate a non-commercial, paper-only basketball research project from South
Africa. We do not intend to scrape public B.League sites or use any undocumented
endpoint. We would like to ask about a formal contract for B.League data.

Please advise whether a B1/B2 data product, API, or XML delivery can license:
fixture IDs, opponents and UTC tip times; final and quarter scores plus
postponement/cancellation/correction data; historical seasons; and automated
research collection, long-term retention, model validation, and publication of
our own non-commercial derived analysis.

Please provide covered seasons/leagues, fields, delivery format, permitted
uses/retention/publication conditions, pricing, rate limits, and contract steps.
We do not seek betting execution, odds data, or raw-data resale.

Thank you.
```

## After a provider responds

A positive commercial answer is still not an integration.  For the one
provider/role/league combination approved next, capture a small set of real
upcoming fixtures in an explicitly authorised, no-write validation.  It must
prove all requested fields in an actual response and retain the response
fingerprint/contract reference without credentials.

Only then can an implementation preserve, at minimum:

- **prediction:** provider event ID, provider/model/version, scheduled UTC tip,
  immutable provider time, client observation time, home/away probabilities,
  payload hash, and contract/version reference;
- **price:** provider event ID, bookmaker ID/name, full-game moneyline, home/away
  selection, odds, market state, provider/book timestamp, client observation
  time, payload hash, and contract/version reference.

The system must reject exchange, consensus, best-price, relayed, post-tip,
untimestamped, mismatched, or post-result records.  It must continue to emit no
paper ticket until its existing fixture, forecast, calibration, exact-price,
and freeze rules are independently met.
