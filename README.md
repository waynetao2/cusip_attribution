# CUSIP-level asset attribution: quarterly process

This replaces the manual build of `Cusip_level_rec_ParQuotes.xlsx`. It brings in Bryan Goodwin's
July 2026 (vBG) proposal and adds fixes for the issues raised in recent quarters.

## Quarterly run (about 5 minutes)
1. Copy `config/26Q2.yaml` to `config/26Q3.yaml`. Then:
   * set `prior_date` / `current_date`
   * paste the UST and JGB par quotes for both dates
   * update the JPI FX rates
2. Run `python attribution.py --config config/26Q3.yaml`
3. Review the `Checks` and `DQ exceptions` tabs. Fix anything there, either in the config or upstream in Aladdin/IGC, then re-run.
4. Send `outputs/Cusip_level_attribution_26Q3.xlsx` (the `exhibit` tab matches the old layout, plus the new columns).

Parallel / backfill: `--from-workbook <old file>` rebuilds a past quarter from its own tabs.
Add `--legacy` to reproduce the delivered numbers exactly.

## What the script does
| Step | Old workbook | New process |
|---|---|---|
| Data pull | 4 queries pasted into tabs (Q1/Q2 data, par book, Axis cells) | `sql/01-04` bound to dates from config |
| Universe | Inner join on riskid + portfolio. Portfolio moves dropped out (372 CUSIPs in 26Q1), and they were fixed by hand | Outer join with a status for each row: Continuing / Transferred / New / Sold. Transfers are auto-matched by riskid |
| Duplicate lots | SUMIFS added duration/OAS across lots | MV/par/book summed, analytics MV-weighted |
| Rate shock | Monthly par grid, lookup at WAL | Same. WAL < 1 month now uses the first grid point instead of #N/A -> 0 |
| Convexity | Aladdin value used as-is | `convexity_scale` (Aladdin quotes /100). **Confirm** |
| Actual | dMV - dPar (vBG col W) | Prior par x change in clean price. This removes trades/paydowns at actual price, and strips FX for JPI |
| Time effect | None, so it fell into the miss | Pull-to-par = (par - MV) / (WAL x 4) |
| Bad analytics | Used as-is (OAS 5,000 bp produced a $1.4M miss) | OAS beyond +/-1,000 bp is flagged and spread becomes the plug |
| True-up | vBG cols AA/AB, abs-value weights | `true_up_method`: abs_weight (vBG), duration-weighted, or none. `Other_unallocated` makes every row tie |
| Controls | Hard-coded Validation tab | `Checks` tab: tie-out, duplicates, unmapped MV, coverage vs IGC |

## Items to confirm before production
* `01_positions.sql`: does `igc.positions` hold CurrentFace and US_GAAP_Book_Value? If not, add the join used for the par/book pull.
* `03_axis_cell_map.sql`: add the valuation-date filter.
* `04_igc_block_totals.sql`: the block attribute name.
* Aladdin convexity scaling (`convexity_scale`).
* True-up method. This needs agreement with Joel, Casey and Bryan.
* Optionally, store par quotes in a table and pull them in SQL instead of pasting them into the config.
* Schedule: run on business day +X after Aladdin month-end is final, using Windows Task Scheduler or the team's job runner.
