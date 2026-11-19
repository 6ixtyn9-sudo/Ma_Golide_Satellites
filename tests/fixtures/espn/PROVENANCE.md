# ESPN fixture provenance

## NBA G League (`basketball/nba-development`)

The two NBA G League fixtures were captured from ESPN's public site API on
2026-10-29 UTC and retained for hermetic parser regressions:

- `scoreboard_gleague_20260410.json` —
  `https://site.api.espn.com/apis/site/v2/sports/basketball/nba-development/scoreboard?dates=20260410`
- `summary_gleague_401867657.json` —
  `https://site.api.espn.com/apis/site/v2/sports/basketball/nba-development/summary?event=401867657`

They are copied from the capture set recorded by
`machina-sports/sports-skills` at commit `8ab0645dcd1dc9a8aca993a5bb27865cc16300e8`.
Its provenance records that fields were pruned for fixture size but values were
not invented or edited. These fixtures demonstrate only the tested route and
schema: event/team identity, UTC time, terminal status, four regulation
linescores, and reconciliation to the final score. They do not establish ESPN
BPI, deep core-event-index history, bookmaker permission, or current upstream
availability.
