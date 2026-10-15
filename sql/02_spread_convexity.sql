-- Aladdin spread convexity (not carried on igc.positions).
SELECT  cusip       AS riskid,
        PORTF_LIST  AS portfoliocode,
        SPREAD_CONVEXITY
FROM    brs.monthly_analytics
WHERE   filedate = :asofdate
