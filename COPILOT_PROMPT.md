# HiPPO Sales Agent Arena — build specification

Paste everything below the line into Microsoft Copilot (or any code assistant) to have it
rebuild, extend or restyle the board. `index.html` in this repo is a finished, working
implementation of this exact spec — you can use it as-is and only come back here when you
want Copilot to change something.

---

Build a **single self-contained HTML file** — all CSS in one `<style>` block, all JavaScript in
one `<script>` block, no frameworks, no CDN links, no build step, no network calls of any kind.
It must run correctly when opened directly from disk (`file://`) on a locked-down corporate
laptop, and be shareable as one file over email or Teams.

It is a daily sales commission leaderboard — a "Sales Agent Arena" — for the HiPPO-branded
insurance sales floor. Team leads upload a CSV once a day; agents read the board.

## The problem it solves

The floor currently runs a flat commission pool. A flat pool pays everyone "on average", which
builds in an exit clause — an agent can coast to the average and still get paid. It is also
unfair across brands: Budget Insurance receives far more contacts than Auto & General, so a
Budget agent looks better on a shared leaderboard purely because of channel volume.

Replace it with a **bracket system where every target is relative to the agent's own brand**.

## 1. Brands

| Code | Brand | Effectiveness target | Earning |
|---|---|---|---|
| `BUDGET` | Budget Insurance | 31.6% | yes |
| `FFW` | First for Women | 34.7% | yes |
| `AG` | Auto & General | 31.5% | yes |
| `DIALDIRECT` | Dial Direct | not set | **no — display only** |

Dial Direct agents appear on the board with live contacts, effectiveness and their own brand
average, but are excluded from the pool, Super Club and Brand Champion until a target is set.
Setting `earning:true` plus an `effTarget` is all it takes to switch them on.

## 2. The two gates — both or nothing

**Gate 1 — Contacts, brand-relative.**
```
brandAvgContacts = Σ contacts in that brand ÷ headcount in that brand
gate             = brandAvgContacts × 90%          (90% editable per brand)
pass             = agent contacts ≥ gate
```
The average is recomputed from the uploaded file every single day. This is the mechanism that
stops a high-volume brand out-ranking a low-volume one: every agent is measured against their
own channel.

**Gate 2 — Effectiveness, brand-specific.**
```
effectiveness = sales ÷ contacts          (as a %)
pass          = effectiveness ≥ that brand's target
```

**Miss either gate and the agent earns R0.** No bracket, no bonuses, nothing. The board must
show *which* gate failed and exactly how far off — "8 contacts short of gate 352", "3 sales
short of 31.6%". Non-qualifiers stay visible on the leaderboard, greyed out; seeing the gap is
the motivator.

## 3. The bracket

Effectiveness is the single most important metric. It must dominate both by weight and by
driving the multiplier.

```
effDepth     = effectiveness ÷ brandEffTarget       (≥ 1.00 for a qualifier)
contactDepth = contacts ÷ brandAvgContacts
bracketMult  = tier multiplier, from effDepth × 100
points       = bracketMult × (0.70 × effDepth + 0.30 × contactDepth)
share        = points ÷ Σ all qualifiers' points
```

| Bracket | Min % of eff. target | Multiplier |
|---|---|---|
| Legend | 130% | ×2.00 |
| Elite | 120% | ×1.70 |
| Contender | 110% | ×1.45 |
| Riser | 105% | ×1.25 |
| Entry | 100% | ×1.00 |

## 4. The pool

- Base **R450,000**. It opens at base on day one so agents see the full opportunity immediately.
- **May grow above base**, hard-stopped at an upper cap (default R520,000).
- Each day's participation = qualifiers ÷ agents in earning brands. If today's participation is
  at or above yesterday's, the pool compounds **×1.015**; if it drops, **×0.985**. Every day
  counts, and the value visibly expires as people fall off.
- Floored at 70% of base. The cap can never sit below the base and the floor can never sit above
  the cap, whatever someone types into the rules screen.
- Snapshots are stored per reporting date and the flex compounds across them in date order.
  Deleting a bad upload recalculates the pool immediately.

**Unlock factor.** `unlock = 0.35 + 0.65 × participation` (0.35 floor, editable). This is what
stops three qualifiers walking away with the entire R450,000 and is what makes the Winner's
Circle an addition rather than a rounding error.

## 5. Bonuses — all funded FROM the pool, never on top of it

- **Super Club** on bracket points: 1st R15,000, 2nd R10,000, 3rd R8,000.
- **Brand Champion**: R5,000 to the top qualifier in each earning brand (stacks with Super Club).
- **Winner's Circle**: when 5 or fewer agents qualify, each gets an extra 5–10% of pool value,
  sliding by scarcity — 5 qualifiers → 5% each, 1 qualifier → 10%.

Order of operations:
```
reserves      = Super Club + (R5,000 × earning brands with a qualifier)
circleTotal   = qualifiers × pool × circlePct
distributable = (pool − reserves − circleTotal) × unlock
basePay       = share × distributable
gross         = basePay + superClub + champion + circle
net           = min(gross − qaDeduction, individualCap R100,000), floored at 0
```

**Budget guard.** After everything, if total payouts would exceed the live pool, scale every
payout down proportionally. The budget must be impossible to blow. A pool exists to manage
budget, not to exceed it.

## 6. QA error points — editable, 1 to 5

A five-row table. Each row independently chooses **rand** or **percent** mode:

| Points | Mode | Rand | Percent |
|---|---|---|---|
| 1 | rand | R500 | 2% |
| 2 | rand | R1,200 | 5% |
| 3 | rand | R2,500 | 10% |
| 4 | percent | R4,000 | 18% |
| 5 | percent | R6,000 | 30% |

Rand mode deducts the flat value; percent mode deducts that share of gross earnings. A deduction
can never push an agent below R0. QA points come from a CSV column, or are typed straight into
the leaderboard row.

## 7. CSV ingest — South African Excel

The file is exported from SA Excel, which writes **`;` as the field delimiter and `,` as the
decimal separator**. Handle it without the user having to think about it:

- Sniff the delimiter from the header row (`;`, `,`, tab, `|` — `;` wins ties).
- If the delimiter is `;` the comma is a decimal point; if it is `,` the comma can only be a
  thousands mark. `"R 1 072 000,00"` → 1072000. `"32,5"` → 32.5. `"1.234,56"` → 1234.56.
  Strip `R`, `%`, brackets and every kind of space including non-breaking.
- Match headers case-, space- and punctuation-insensitively with aliases:
  business manager, team lead, agent/operator/consultant, brand/channel, contacts, leads,
  conversions, sales/closings, gross sales, gross effectiveness, QA points.
- Match brand names fuzzily: `budget`, `first for women`/`first woman`/`ffw`,
  `auto and general`/`auto & general`/`a&g`, `dial direct`.
- **Always recompute** conversion, closing and effectiveness from the raw counts
  (`leads÷contacts`, `sales÷leads`, `sales÷contacts`) so a mis-exported percentage column cannot
  corrupt the standings. Only fall back to the supplied percentages when the counts are missing.
- Drag-and-drop plus an upload button. Parse entirely in-browser with `FileReader` — the data
  never leaves the machine.

## 8. Screens

**The big number.** The live pool in oversized type at the top of every screen, with its delta
against base, days banked, participation, allocated total, top earner, and a bar showing where
it sits between floor and cap.

**Arena.** Ranked leaderboard: rank, agent, business manager, brand chip in brand colour, team
lead, contacts with gate status and a progress rail, effectiveness with target status and a
rail, conversion, closing, bracket badge, points, an editable QA box, and rand earned. Sortable
by any column, searchable, filterable by brand, team lead, business manager and snapshot date.

**Brands & Averages.** One card per brand showing headcount, running average contacts, the live
contacts gate, the effectiveness target, actual brand effectiveness and qualifiers. Below it, a
per-brand roster ranked on effectiveness. A team lead picks their brand; an agent sees exactly
what their own channel's average is.

**My Path.** Two tools for the selected agent:

*The Coach* — fun and educational, live maths on their own numbers:
> "From your 344 contacts, at your current closing of 75.0%, you need 18 more leads — that is 13
> more sales — to reach 31.6% effectiveness."

Plus sliders on contacts, leads and sales that instantly recompute conversion, closing and
effectiveness.

*The Path Loader* — an input and slider from R0 to R100,000 and a **Load Path** button. It
back-solves against the live field: hold the agent's conversion rate, lift closing (capped at a
believable 95%) and contacts together, always clearing the contacts gate, and binary-search the
effort level that pays the requested rand value. Then push those numbers into every metric —
contacts, leads, sales, conversion, closing, effectiveness — each showing "was X", under a clear
**PROJECTION** banner with a one-click reset. Show the projected rand value in very large type
along with projected rank.

Three honest outcomes, each with its own message:
- exact hit → "Path found — hit these numbers and you land here."
- overshoot → "Nearest achievable point above your target. Earnings step up in jumps — at each
  bracket edge, and again when you break into the Super Club or take your brand."
- unreachable → "Out of reach today: 18 agents are already sharing this pool. Your ceiling
  against the current field is R29,747 — it rises as others fall away and as the pool grows."

**Rules.** Every value above editable in-page with instant recalculation: pool base, cap, floor,
daily flex, individual cap, unlock floor, per-brand targets/gates/earning flag/colour, the
bracket table, the weights, the QA table, Super Club and Winner's Circle values. Plus export and
import of the rules as JSON, and a day-history table with delete.

## 9. Persistence

`localStorage` — one key for the rules, one for the day snapshots (full rows per date, so
changing a rule retroactively recalculates the whole month consistently). Re-uploading the same
date overwrites rather than double-counting. A "Clear Month" action resets to base.

## 10. Look and feel

Dark arena shell in the HiPPO palette, with each brand's own colour driving its rows, chips and
brand card. Large tabular-numeric figures, generous spacing, progress rails on the two gates.
Must be readable on a phone (team leads check it on the move) with no horizontal page scroll —
wide tables scroll inside their own container.

| Brand | Reference site | Primary | Accent |
|---|---|---|---|
| HiPPO shell | hippo.co.za | green `#00A758` / `#00D46A` | pink `#E5007E` |
| Budget Insurance | budget.co.za | blue `#0072CE` | navy `#003C71` |
| First for Women | firstforwomen.co.za | pink `#E6007E` | magenta `#B4006B` |
| Auto & General | autogeneral.co.za | red `#E4002B` | dark red `#A30020` |
| Dial Direct | dialdirect.co.za | violet `#6B2FA0` | — |

Keep every colour in a single config object so exact brand hex codes can be swapped in one
place. Use a system font stack so the file renders identically with no network.
