# Capture provenance for the committed provider fixtures

Everything here comes from one probe package produced **outside this repository**
by the operator's external agent on 2026-11-22 and delivered as text: each
artifact carried its own byte count and sha256, and every body was re-hashed on
arrival before it was committed. 24 of the 28 transmitted artifacts reproduced
byte-exactly; the four that did not are excluded rather than "fixed up" (see the
bottom of this page).

This file exists so the evidence class of a fixture is checkable rather than
asserted. `tests/test_capture_fidelity.py` re-hashes the two whole bodies on
every run: editing a capture now fails a test instead of quietly turning a
capture back into a shape.

## Committed, byte-exact whole bodies

| Fixture | Provider route | Bytes | sha256 |
| --- | --- | --- | --- |
| `lba/championships_2026.json` | `legabasket.it/api/v1/championships?cs_id=1&items=1000&s=2026` | 867 | `e6d648fb4f57e5a44c6fc3053c974402d103866a95291d758aba6f3a266e70cc` |
| `lba/calendar_602.json` | `legabasket.it/api/v1/championships/get-championships-calendar-by-id?id=602` | 11636 | `9a0e5d28c18eac20d06aeef94b9ca354e0597319301b2ba48628c5853c155101` |

The calendar is the league's current focus day at capture time: eight rows
spanning 2026-11-21 and 2026-11-22, with `filters.days` marking round 9 (id
4494) as focused rather than any single date.

## Committed, container written here around verbatim captured rows

The provider's play-by-play body is 314 KB and its season board is 159 KB; the
package transmitted verbatim byte excerpts instead of whole bodies. The row
objects below are copied unchanged from those excerpts — only the JSON envelope
(`{"pbp": {"actions": [...]}}`, `{"games": [...]}`, `{"data": [...]}`) and
the whitespace of the re-serialisation are this repository's.

| Fixture | Captured source | Rows used | Bytes (as committed) | sha256 (as committed) |
| --- | --- | --- | --- | --- |
| `lba/pbp_25426.json` | period-end actions of `get-championships-matches-play-by-play?id=25426` (Trento 100-86 Napoli) | 4 | 2389 | `6ba45b151318d157fb77fa1a86982ed8...` |
| `lba/pbp_25422_ot.json` | same route, `id=25422` (Cantù 92-106 Tortona, after overtime) | 5 | 2976 | `f4c0d57b6ed61e75e1e319e585ecf6b1...` |
| `israel_ibl/board5_captured_rows.json` | `basket.co.il/ws/ws.asmx/Games?cYear=2027&board_id=5...`, four rows | 4 | 4443 | `90804be6d6a85b0ce4f19fe0955e6058...` |
| `eurocup/U2026_ot_game.json` | `api-live.euroleague.net/v2/competitions/U/seasons/U2026/games`, the first row with `extraPeriods` | 1 | 3021 | `5f96e379955eaa3c546a974280fa16f5...` |

`lba/pbp_synthetic_noncontiguous_ot.json` is **not** a capture: it is a
hand-built refusal case (regulation anchors plus a lone period 6) and is named
accordingly.

## Not committed

Oversized original bodies that were never transmissible whole (retained only by
the external agent): the four LBA play-by-play bodies (313–372 KB), the ACB
2027 calendar (329 KB), the Israel board (159 KB), the EuroCup `U2026` schedule
(457 KB), the BCL GDAP response (309 KB), and three BetExplorer pages
(110–128 KB).

Four transmitted artifacts did **not** reproduce their stated hash on arrival and
are therefore not committed as captures: `acb/matchheader_401.txt` (366/372
bytes), `betexplorer/matchodds_404.txt` (910/917), `israel/board5_rows.slice.txt`
(4723/4725 — the four row objects inside it were intact and are what the Israel
fixture uses), and `nbl/sdist_nbl_listing.txt` (204/206). Their claims are
recorded in `docs/PROVIDERS.md` as *reported by the probe*, not as
repository-verified bytes.
