# Dashboard Logic Reference

What every KPI card, table, and chart on the NeoEngage dashboard actually shows and how it's
calculated. If a number looks surprising, check here before assuming it's wrong — most
"surprising" numbers turn out to be a deliberate rule documented below (e.g. Win rate isn't
what most people guess it is).

This documents behavior as implemented in `app.js`. If the code changes, this file can drift —
treat it as a guide to *how it's meant to work*, and cross-check against the function named in
each section if a number still looks wrong after reading this.

---

## 1. Core rules everything else builds on

Read this section first — every KPI, table, and chart in the app is built out of these few rules.

### 1.1 Currency — everything is entered in SGD directly (as of 2026-08-11)

Both **Estimated Value** (Opportunities) and **Target Revenue** (Companies) are typed in
directly as SGD figures, in every region — there's no conversion happening anywhere anymore.
What you enter is what gets summed and displayed.

```
Estimated Value (SGD)  =  estimated value, as entered
Weighted Value (SGD)   =  estimated value × probability, as entered
Target Revenue (SGD)   =  target revenue, as entered
```

**This is a change from how it used to work**, worth knowing if you're cross-checking against
older exports or screenshots: opportunity values used to be entered in each region's own local
currency (ME/US = USD, APAC = SGD, IND = INR) and converted to SGD using a live FX rate table
that finance maintained, with the rate *frozen* the moment a deal closed (Win/Loss) so two deals
closed months apart at different rates kept their own historical conversion. Target Revenue used
to be entered in USD and converted at today's rate.

That whole FX-conversion mechanism (`FX_RATES`, `fetchFxRates()`, `REGION_CURRENCY`,
`fxRateFor()`, `currencyForOpp()`, and the `fx_rate_locked` field written when a deal closes) is
still sitting in `app.js`, but nothing reads it anymore — `oppEstSgd()`, `oppWeightedSgd()`, and
`targetRevSgd()` are now plain pass-throughs. It's inert, not deleted, in case this ever needs to
come back.

### 1.2 Opportunity stages — the 9 stages, in order

| Stage | Bucket | Probability |
|---|---|---|
| Blue Sky Opportunity | Open | Manual — pick 10/25/75/100% |
| Identified Opportunity | Open | Manual |
| *(stage "03" was retired at some point — the numbering intentionally skips it internally; nothing to look for here)* | — | — |
| Proposal preparation | Open | Manual |
| Proposal submitted | Open | Manual |
| To Win | Open | **Auto-locked to 100%** |
| Win | Closed | **Auto-locked to 100%** |
| Loss | Closed | **Auto-locked to 0%** |
| Client not interested | Closed | **Auto-locked to 0%** |
| Missed Opportunity | Closed | **Auto-locked to 0%** |

- **Open pipeline** = Blue Sky, Identified, Proposal preparation, Proposal submitted, To Win.
- **Closed** = Win, Loss, Client not interested, Missed Opportunity.
- The moment an opportunity's stage is set to To Win, Win, Loss, Client not interested, or
  Missed Opportunity, the Probability field auto-fills and **locks** (greyed out, with a 🔒
  note) — it can't be manually overridden in those stages.
- **To Win is still "open"**, not closed — it keeps counting in open/weighted pipeline totals,
  just at 100% probability (i.e. its full value counts). FX rate does **not** lock at To Win —
  only once it actually becomes Win or Loss.

### 1.3 Weighted value

```
Weighted Value  =  Estimated Value  ×  Probability
```

Computed once and stored on the record (`probability_weighted_value`) at the moment an
opportunity is saved or edited — it isn't recalculated live from probability every time it's
displayed, so if you change probability, weighted value only updates the next time that
record is saved.

### 1.4 Win rate

Shown on the Overview, Opportunities, and All Regions pages. It is **not** wins ÷ all
opportunities:

```
Win Rate  =  (Wins)  ÷  (Opportunities that have reached "Proposal submitted", i.e. have a
              Proposal Submission Date filled in)
```

An opportunity still sitting at Blue Sky or Identified doesn't count in the denominator at
all — only ones that got as far as actually submitting a proposal. This is why Win Rate can
look "high" even with lots of early-stage opportunities in the pipeline — those aren't being
penalized in the ratio.

---

## 2. Overview page

*(`renderOverview()`)*

| KPI tile | Formula |
|---|---|
| Open pipeline | Sum of Estimated Value (SGD) across all open-stage opportunities |
| Weighted pipeline | Sum of Weighted Value (SGD) across all open-stage opportunities |
| Win rate | See §1.4 |
| Overdue CTAs | Count of engagements where Follow-up done ≠ "Yes" **and** CTA Due Date is in the past |

**Engagement by BD Owner / by Department tables** — counts of engagements in the selected
period (Weekly / Monthly / Quarterly / Year-to-date — see §2a below), broken down by BD/PM
owner, cross-tabbed against Engagement Type (first table) or the client's Department (second
table, looked up from the Clients list by matching Company + Client Name).

**Pipeline Funnel row** (bottom of the BD Sales Funnel Tracker) — Estimated Value and Weighted
Value for each of the 6 open-or-won stages (Blue Sky → Win), plus a Total column. Matched by
exact stage name.

**Weekly engagement trend chart** — count of engagements per week, last 10 weeks, weeks start
Monday (ISO week).

**Account Health Matrix** ("Total Touch Points" table) — one row per company:

| Column | Formula |
|---|---|
| Total Touch Points | Count of engagements logged for that company |
| Win | Sum of Estimated Value for that company's Won opportunities |
| To Win | Sum of Estimated Value for that company's To Win opportunities |
| Pipeline | Sum of Estimated Value for *open* opportunities, **excluding** To Win (To Win is already broken out in its own column, so it isn't double-counted here) |
| Lose | Sum of Estimated Value for that company's Loss + Client not interested opportunities |
| Total | Win + To Win + Pipeline + Lose |

Rows with zero engagements, wins, To Win, pipeline, and losses are hidden. Sorted by Total,
highest first.

**Follow-ups due this week** — engagements not yet marked done, with a CTA Due Date within the
next 7 days (including already-overdue ones), soonest first, capped at 8 shown.

**Recent outcomes** — the 10 most recently *decided* Win/Loss opportunities, ranked by the date
the stage actually changed to Win/Loss (pulled from that opportunity's stage history), falling
back to Expected Close Date if no history exists.

### 2a. What "period" means (Weekly/Monthly/Quarterly/YTD filters)

Fiscal year runs **April → March** (matches "FY 2026–27" branding). `periodStart()`:

- **Weekly** = Monday of the current week
- **Monthly** = 1st of the current month
- **Quarterly** = start of whichever Apr/Jul/Oct/Jan block contains today
- **YTD** = April 1st of the current fiscal year (or last April 1st, if today is Jan–Mar)

Only *engagement counts* are period-filtered by this control — pipeline/weighted/win-rate
figures are always all-time regardless of the period picker, everywhere in the app. An open
opportunity doesn't stop counting as open pipeline just because it was created outside the
selected window.

---

## 3. BD Sales Funnel Tracker

Appears identically on both the **Overview** page (single region) and **All Regions Overview**
(merged across every region) — same logic, different data scope. *(`renderBDFunnel()` /
`renderAllRegionsBDFunnel()`)*

**Section 1 — Stakeholder List**: count of clients, "Total" (all clients) or "Active" (status =
Active business or Active engagement only, toggle via the Show dropdown), broken down by
Department. Collapsed to just the Total Clients row by default — "Show breakdown" expands it.

**Section 2 — Engagement Overview**: count of engagements by Engagement Type, filterable by
month (defaults to All Time).

**Section 3 — Opportunity Mapping**: count of *all* opportunities by every one of the 9 stages.

**Section 4 — Pipeline Funnel**: one column per open-or-won stage (Blue Sky → Win, 6 columns),
each showing opportunity count and Estimated Value, plus a Total column. Matched by exact
stage name — see the fix note in §9.

---

## 4. All Regions Overview

Visible only to SuperAdmin, Global User, and Global Viewer roles. *(`renderAllRegionsOverview()`
and friends)* — pulls every region's data in parallel and merges it, tagging each record with
which region it came from so per-region breakdowns stay possible.

| KPI tile | Formula |
|---|---|
| Open pipeline | Same as Overview §2, across the merged (all-region) dataset |
| Weighted pipeline | Same, merged |
| Win rate | Same formula as §1.4, merged |
| Engagements logged | Count of engagements in the selected period (this one *is* period-filtered — see §2a) |

**Region comparison chart** — Open pipeline vs. Weighted pipeline, one bar pair per region.

**Region summary table** — per region: Clients, Active clients, Open Opps, Pipeline, Weighted,
Wins, Win Rate, Engagements (in the selected period) — plus a Grand Total row (Win Rate isn't
summed in the total row, since averaging rates isn't meaningful).

**Pipeline by stage per region table** — same 6 open-or-won stages as §3's Pipeline Funnel,
counts only, one row per region plus a Grand Total row.

**Engagement activity trend chart** — one line per region, engagement count per week, last 12
weeks.

---

## 5. Companies page

*(`renderCompanies()`, `computeCompanyRollup()`)*

| KPI tile | Formula |
|---|---|
| Companies | Count of distinct companies (merges known duplicate SharePoint entries, e.g. "Abbvie"/"AbbVie" → one company) |
| Combined pipeline | Sum of open-stage Estimated Value across every company |
| Total wins | Count of Won opportunities across every company |
| Onboarding on file | Count of companies with an Onboarding Status set |

**Company table columns** — per company, rolled up from Clients/Opportunities/Engagements:

| Column | Meaning |
|---|---|
| Onboarding | From the Companies record directly |
| Target Revenue (SGD) | See §1.1 |
| Clients / Active | Total clients for this company / how many have status Active business or Active engagement |
| Open Opps | Count of open-stage opportunities |
| Pipeline | Sum of Estimated Value across *all* open-stage opportunities (unlike the Account Health Matrix in §2, this figure **does** include To Win — it isn't broken out separately here) |
| Weighted | Sum of Weighted Value across open-stage opportunities |
| Wins | Count of Won opportunities |
| Engagements / Last Engagement | Count of logged engagements / most recent engagement date |

Rows are tinted light red if the company's last engagement was **more than 90 days ago**.

**Health Matrix** (Overall Budget Potential, Overall Client Relationship, Client Perception,
Team Satisfaction, Degree of Innovation) — these are **not calculated**, they're manually
entered scores (1–5, or a category for Client Perception) shown read-only here; editing happens
through that company's Edit modal.

---

## 6. Client Database page

*(`renderClients()`, `renderClientPivot()`)*

KPI tiles are plain counts: Total clients (with distinct-company count as the subtitle), and one
tile per status (Active business, Active engagement, Prospects, Inactive).

**Stakeholder pivot table** — Company → Department (drill down by clicking a company row),
cross-tabbed by Status, with row and grand totals. Optionally filtered to one Therapy Area.

---

## 7. Opportunity Tracker page

*(`renderOpportunities()`, `renderOpPivotTable()`, `renderOpStageChart()`)*

| KPI tile | Formula |
|---|---|
| Total | Count of all opportunities (subtitle: how many are open) |
| Open pipeline | Sum of Estimated Value, open stages only |
| Weighted | Sum of Weighted Value, open stages only |
| Win | Count + Estimated Value sum of Won opportunities |
| Lost | Count + Estimated Value sum of Loss **+** Client not interested opportunities |
| Win rate | See §1.4 |

**Stage pivot table** — one row per stage that has ≥1 opportunity (all 9 stages, not just the
6 open-or-won ones): count, Estimated Value, Weighted Value, plus a Grand Total row. Optionally
filtered to one company.

**Stage chart** — bar = opportunity count per stage (all 9 stages), line = Weighted Value per
stage, same company filter as the pivot table above.

---

## 8. Engagement Log page

*(`renderEngagements()`, `renderEngagementPivotTable()`)*

| KPI tile | Formula |
|---|---|
| Total | Count of all logged engagements |
| This week | Count with an engagement date in the current ISO week (Monday–today) |
| Last week | Count in the prior ISO week, with a ▲/▼/– delta vs. this week shown as the subtitle |
| Follow-up pending | Count where Follow-up done ≠ "Yes" |
| Led to RFP | Count where Engagement Outcome contains the text "led to rfp" (case-insensitive) |

**Engagement pivot table** — Outcome (row group) → Stakeholder Type (drill-down rows), cross-
tabbed by Engagement Type, with grand totals. Filterable by BD/PM, month, company, objective,
and engagement type.

**Weekly bar chart** — engagement count per week, last 12 weeks; clicking a bar drills into that
week's engagements.

---

## 9. Known quirks that look like bugs but aren't

- **Stage "03" doesn't exist.** It was retired from the pipeline at some point; both the tool
  and the Excel templates intentionally skip straight from "02 Identified Opportunity" to
  "04 Proposal preparation" internally. Nothing to fix here.
- **"AbbVie" vs "Abbvie"** and similar near-duplicate company names typed inconsistently across
  regions are merged on read (`canonicalCompany`/`normCo`) so they roll up as one company
  everywhere — you won't see two separate rows for the same real company just because of
  capitalization differences.
- **Win Rate's denominator is "submitted," not "total."** See §1.4 — this is by design, not a
  miscount.
- **Account Health Matrix's "Pipeline" column excludes To Win; the Companies table's "Pipeline"
  column does not.** Two different tables, two slightly different definitions of "Pipeline" —
  documented in §2 and §5 respectively. If the two numbers don't match for the same company,
  this is why.
- **FX conversion no longer runs at all** (as of 2026-08-11) — figures are entered directly in
  SGD everywhere now. The old locking-on-close behavior is retired but not deleted; see §1.1.

---

## 10. Fixed bug (for reference — resolved 2026-08-10)

The "Pipeline Funnel" row (§3) and the All Regions "Pipeline by stage" table (§4) used to build
their columns from array position (`LOOKUPS.opportunityStatus[n-1]` for n = 1..7) while filtering
rows by a *derived stage number* that has a gap at the retired stage "03" (see §9). That mismatch
silently shifted every column from "Proposal submitted" onward by one stage — most seriously, the
**column labeled "Win" was actually showing "To Win" numbers**, and "Proposal preparation" was
always empty. Both were fixed to match columns by exact stage name instead of a number, which
also makes them immune to this class of bug in the future. Nothing else in the app used the
buggy pattern — KPI tiles, Win Rate, and Recent Outcomes all use the proper `isWinStage()` /
`isLossStage()` checks and were never affected.
