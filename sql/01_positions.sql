-- IGC position-level analytics for one quarter-end (run once per date).
-- :asofdate is bound by attribution.py (e.g. '2026-06-30').
-- CurrentFace and US_GAAP_Book_Value replace the separate manual "par book" pull.
-- CONFIRM: if CurrentFace / US_GAAP_Book_Value live in a different table, change the join here only.
SELECT  p.riskid,
        p.portfoliocode,
        p.saa_level1,
        p.WAL,
        p.MV_CLEAN,
        p.EFF_DUR,
        p.EFF_CONV,
        p.SPREAD_DUR,
        p.OAS,
        p.CurrentFace,
        p.US_GAAP_Book_Value
FROM    igc.positions p
WHERE   p.asofdate   = :asofdate
  AND   p.sourcetype = 'Monthly'
