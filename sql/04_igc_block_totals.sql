-- Control totals for the coverage check (replaces the hard-coded numbers on the Validation tab).
-- CONFIRM: block/fund attribute name on igc.positions.
SELECT  model_block, SUM(MV_CLEAN) AS MV_CLEAN,
        SUM(CASE WHEN saa_level1 = 'Cash Securities' THEN MV_CLEAN ELSE 0 END) AS CASH_MV
FROM    igc.positions
WHERE   asofdate = :asofdate AND sourcetype = 'Monthly'
GROUP BY model_block
