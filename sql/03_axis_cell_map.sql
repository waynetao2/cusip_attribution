-- Axis asset-cell to model-block mapping (current quarter asset file).
-- The CUSIP sits in characters 20-28 of cell_name and the 3-char portfolio suffix in 29-31,
-- e.g. BOND---_01CMBS--02_61767CAY4FT3_ON-BNFIT -> 61767CAY4 + TR_FT3.
SELECT  cell_name,
        subfund_name_input_bel,
        SUBSTRING(cell_name, 20, 9)          AS riskid,
        'TR_' + SUBSTRING(cell_name, 29, 3)  AS portfoliocode,
        CASE fund_name_input_bel
             WHEN 'Coins_FIA'       THEN 'Coin'
             WHEN 'Modco_FIA'       THEN 'Modco'
             WHEN 'JPI'             THEN 'JPI'
             WHEN 'TLA Payout'      THEN 'TLA Payout'
             WHEN 'Allianz Surplus' THEN 'TLR Surplus (Non-JPI)'
             WHEN 'Pru'             THEN 'Pru'
             WHEN 'DFL'             THEN 'DFL'
             ELSE ''
        END AS Model
FROM    [AxisInputAssetProd].[rpt].[fact_axis_asset_file]
-- CONFIRM: add the valuation-date / run filter used for the quarter, e.g.
-- WHERE valuation_date = :asofdate
