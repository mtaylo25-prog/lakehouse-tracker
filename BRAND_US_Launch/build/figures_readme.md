# figures.json: every key, its meaning and the workbook cell it comes from

Pulled by build/pull_figures.py from output/03_Financial_Model_5Y.xlsx after LibreOffice recalculation. Scenario blocks are pulled by setting the Scenario selector (Assumptions!C3) to 1, 2 and 3 on temporary copies; the delivered workbook stays at Scenario 2 (Medium), Duty 1 (Copper), FX 1430. Every defined name is listed in build/figures_map.json (name to sheet!cell) and every JSON key path in build/figures_keys.json (path to name, cell and selector state). Numbers are raw; format with build/fmt.py or build/fmt.js.

## Conventions

Model months: Month 1 is the first month with inventory live on Amazon (assumed early 2027). Year 1 is Months 1 to 12, Year 2 is Months 13 to 24.
18 month block (m18): Amazon US only, the August basis, Months 1 to 18: revenue, net after fixed costs, seller fee and launch costs, net margin, cumulative breakeven month (null if never), Month 18 run rate (Month 18 revenue times 12), exit margin (Month 18 net margin), units.
Monthly rows: revenue and net are all channels (Amazon US plus the eBay and Walmart overflow and B2B rows from Month 13); amazon_revenue and amazon_net are Amazon US only; cash_balance is before new funding, cash_balance_after_funding adds the line of credit or investor amount in the month it is needed.
Annual: revenue, net, net margin, COGS, channel fees (referral, payment, commission), fulfillment (FBA, MCF, 3PL, shipping), advertising, returns and storage, contribution, fixed, seller fee, launch, compliance, revenue_by_channel and contribution_by_channel for amazon_us, ebay_walmart, b2b, dtc, amazon_canada, pro_line.
Cash: peak_need = the negative minimum cash balance before funding (Months 0 to 24 and the Year 3 to 5 year ends); peak_need_month (number; 36, 48, 60 mean a year end) and peak_need_month_label; funding_gap = 1.2 times peak need; funding_needed_by_month = the month before the balance first goes negative (null if never); runway_months_day90 = cash at end of Month 3 after funding divided by the average burn of Months 4 to 6 (the string funded if there is no burn); runway_months_day90_unfunded = the same before funding.
Gates: day90 and month12 carry each modeled metric with its value, pass threshold, kill trigger and the PASS or KILL booleans, plus result = PASS, REVIEW or KILL (KILL if any kill trigger fires, PASS if every modeled threshold is met, REVIEW otherwise). Conversion and listing suppression are observed at the gate, not modeled.

## Key list

| JSON key path | Defined name | Workbook cell | Selector state |
|---|---|---|---|
| scenarios.Low.m18.revenue | m18_revenue | 'Monthly PL Y1 Y2'!$C$60 | 1 |
| scenarios.Low.m18.net | m18_net | 'Monthly PL Y1 Y2'!$C$61 | 1 |
| scenarios.Low.m18.net_margin | m18_net_margin | 'Monthly PL Y1 Y2'!$C$62 | 1 |
| scenarios.Low.m18.breakeven_month | m18_breakeven_month | 'Monthly PL Y1 Y2'!$C$63 | 1 |
| scenarios.Low.m18.run_rate | m18_run_rate | 'Monthly PL Y1 Y2'!$C$64 | 1 |
| scenarios.Low.m18.exit_margin | m18_exit_margin | 'Monthly PL Y1 Y2'!$C$65 | 1 |
| scenarios.Low.m18.units | m18_units | 'Monthly PL Y1 Y2'!$C$66 | 1 |
| scenarios.Low.m24.revenue | m24_revenue | 'Monthly PL Y1 Y2'!$C$68 | 1 |
| scenarios.Low.m24.net | m24_net | 'Monthly PL Y1 Y2'!$C$69 | 1 |
| scenarios.Low.m24.breakeven_month | m24_breakeven_month | 'Monthly PL Y1 Y2'!$C$67 | 1 |
| scenarios.Low.monthly[].units | monthly_units | 'Monthly PL Y1 Y2'!$D$7:$AA$7 | 1 |
| scenarios.Low.monthly[].revenue | monthly_revenue | 'Monthly PL Y1 Y2'!$D$54:$AA$54 | 1 |
| scenarios.Low.monthly[].net | monthly_net | 'Monthly PL Y1 Y2'!$D$55:$AA$55 | 1 |
| scenarios.Low.monthly[].cum_net | monthly_cum_net | 'Monthly PL Y1 Y2'!$D$56:$AA$56 | 1 |
| scenarios.Low.monthly[].cash_balance | monthly_cash_balance | 'Cash Flow and Working Capital'!$C$33:$AA$33 | 1 |
| scenarios.Low.monthly[].tacos | monthly_tacos | 'Monthly PL Y1 Y2'!$D$24:$AA$24 | 1 |
| scenarios.Low.monthly[].contribution_before_ads | monthly_contribution_before_ads | 'Monthly PL Y1 Y2'!$D$25:$AA$25 | 1 |
| scenarios.Low.monthly[].contribution_after_ads | monthly_contribution_after_ads | 'Monthly PL Y1 Y2'!$D$26:$AA$26 | 1 |
| scenarios.Low.monthly[].amazon_revenue | monthly_amazon_revenue | 'Monthly PL Y1 Y2'!$D$9:$AA$9 | 1 |
| scenarios.Low.monthly[].amazon_net | monthly_amazon_net | 'Monthly PL Y1 Y2'!$D$21:$AA$21 | 1 |
| scenarios.Low.monthly[].cash_balance_after_funding | monthly_cash_balance_after_funding | 'Cash Flow and Working Capital'!$C$54:$AA$54 | 1 |
| scenarios.Low.monthly[].inventory_on_hand | monthly_inventory_on_hand | 'Cash Flow and Working Capital'!$C$37:$AA$37 | 1 |
| scenarios.Low.monthly[].inventory_in_transit | monthly_inventory_in_transit | 'Cash Flow and Working Capital'!$C$38:$AA$38 | 1 |
| scenarios.Low.annual.Y1.revenue | annual_Y1_revenue | 'Annual PL Y1 to Y5'!$C$69 | 1 |
| scenarios.Low.annual.Y1.net | annual_Y1_net | 'Annual PL Y1 to Y5'!$C$80 | 1 |
| scenarios.Low.annual.Y1.net_margin | annual_Y1_net_margin | 'Annual PL Y1 to Y5'!$C$81 | 1 |
| scenarios.Low.annual.Y1.cogs | annual_Y1_cogs | 'Annual PL Y1 to Y5'!$C$70 | 1 |
| scenarios.Low.annual.Y1.fees | annual_Y1_fees | 'Annual PL Y1 to Y5'!$C$71 | 1 |
| scenarios.Low.annual.Y1.fulfillment | annual_Y1_fulfillment | 'Annual PL Y1 to Y5'!$C$72 | 1 |
| scenarios.Low.annual.Y1.advertising | annual_Y1_advertising | 'Annual PL Y1 to Y5'!$C$73 | 1 |
| scenarios.Low.annual.Y1.returns | annual_Y1_returns | 'Annual PL Y1 to Y5'!$C$74 | 1 |
| scenarios.Low.annual.Y1.contribution | annual_Y1_contribution | 'Annual PL Y1 to Y5'!$C$75 | 1 |
| scenarios.Low.annual.Y1.fixed | annual_Y1_fixed | 'Annual PL Y1 to Y5'!$C$76 | 1 |
| scenarios.Low.annual.Y1.seller_fee | annual_Y1_seller_fee | 'Annual PL Y1 to Y5'!$C$77 | 1 |
| scenarios.Low.annual.Y1.launch | annual_Y1_launch | 'Annual PL Y1 to Y5'!$C$78 | 1 |
| scenarios.Low.annual.Y1.compliance | annual_Y1_compliance | 'Annual PL Y1 to Y5'!$C$79 | 1 |
| scenarios.Low.annual.Y1.cum_net | annual_Y1_cum_net | 'Annual PL Y1 to Y5'!$C$82 | 1 |
| scenarios.Low.annual.Y1.amazon_units | annual_Y1_amazon_units | 'Annual PL Y1 to Y5'!$C$6 | 1 |
| scenarios.Low.annual.Y1.asp | annual_Y1_asp | 'Annual PL Y1 to Y5'!$C$7 | 1 |
| scenarios.Low.annual.Y1.landed_pct | annual_Y1_landed_pct | 'Annual PL Y1 to Y5'!$C$8 | 1 |
| scenarios.Low.annual.Y1.ads_pct | annual_Y1_ads_pct | 'Annual PL Y1 to Y5'!$C$9 | 1 |
| scenarios.Low.annual.Y1.amazon_contribution_pct | annual_Y1_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$C$83 | 1 |
| scenarios.Low.annual.Y1.revenue_by_channel.amazon_us | annual_Y1_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$C$15 | 1 |
| scenarios.Low.annual.Y1.revenue_by_channel.ebay_walmart | annual_Y1_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$C$24 | 1 |
| scenarios.Low.annual.Y1.revenue_by_channel.b2b | annual_Y1_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$C$33 | 1 |
| scenarios.Low.annual.Y1.revenue_by_channel.dtc | annual_Y1_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$C$42 | 1 |
| scenarios.Low.annual.Y1.revenue_by_channel.amazon_canada | annual_Y1_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$C$51 | 1 |
| scenarios.Low.annual.Y1.revenue_by_channel.pro_line | annual_Y1_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$C$60 | 1 |
| scenarios.Low.annual.Y2.revenue | annual_Y2_revenue | 'Annual PL Y1 to Y5'!$D$69 | 1 |
| scenarios.Low.annual.Y2.net | annual_Y2_net | 'Annual PL Y1 to Y5'!$D$80 | 1 |
| scenarios.Low.annual.Y2.net_margin | annual_Y2_net_margin | 'Annual PL Y1 to Y5'!$D$81 | 1 |
| scenarios.Low.annual.Y2.cogs | annual_Y2_cogs | 'Annual PL Y1 to Y5'!$D$70 | 1 |
| scenarios.Low.annual.Y2.fees | annual_Y2_fees | 'Annual PL Y1 to Y5'!$D$71 | 1 |
| scenarios.Low.annual.Y2.fulfillment | annual_Y2_fulfillment | 'Annual PL Y1 to Y5'!$D$72 | 1 |
| scenarios.Low.annual.Y2.advertising | annual_Y2_advertising | 'Annual PL Y1 to Y5'!$D$73 | 1 |
| scenarios.Low.annual.Y2.returns | annual_Y2_returns | 'Annual PL Y1 to Y5'!$D$74 | 1 |
| scenarios.Low.annual.Y2.contribution | annual_Y2_contribution | 'Annual PL Y1 to Y5'!$D$75 | 1 |
| scenarios.Low.annual.Y2.fixed | annual_Y2_fixed | 'Annual PL Y1 to Y5'!$D$76 | 1 |
| scenarios.Low.annual.Y2.seller_fee | annual_Y2_seller_fee | 'Annual PL Y1 to Y5'!$D$77 | 1 |
| scenarios.Low.annual.Y2.launch | annual_Y2_launch | 'Annual PL Y1 to Y5'!$D$78 | 1 |
| scenarios.Low.annual.Y2.compliance | annual_Y2_compliance | 'Annual PL Y1 to Y5'!$D$79 | 1 |
| scenarios.Low.annual.Y2.cum_net | annual_Y2_cum_net | 'Annual PL Y1 to Y5'!$D$82 | 1 |
| scenarios.Low.annual.Y2.amazon_units | annual_Y2_amazon_units | 'Annual PL Y1 to Y5'!$D$6 | 1 |
| scenarios.Low.annual.Y2.asp | annual_Y2_asp | 'Annual PL Y1 to Y5'!$D$7 | 1 |
| scenarios.Low.annual.Y2.landed_pct | annual_Y2_landed_pct | 'Annual PL Y1 to Y5'!$D$8 | 1 |
| scenarios.Low.annual.Y2.ads_pct | annual_Y2_ads_pct | 'Annual PL Y1 to Y5'!$D$9 | 1 |
| scenarios.Low.annual.Y2.amazon_contribution_pct | annual_Y2_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$D$83 | 1 |
| scenarios.Low.annual.Y2.revenue_by_channel.amazon_us | annual_Y2_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$D$15 | 1 |
| scenarios.Low.annual.Y2.revenue_by_channel.ebay_walmart | annual_Y2_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$D$24 | 1 |
| scenarios.Low.annual.Y2.revenue_by_channel.b2b | annual_Y2_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$D$33 | 1 |
| scenarios.Low.annual.Y2.revenue_by_channel.dtc | annual_Y2_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$D$42 | 1 |
| scenarios.Low.annual.Y2.revenue_by_channel.amazon_canada | annual_Y2_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$D$51 | 1 |
| scenarios.Low.annual.Y2.revenue_by_channel.pro_line | annual_Y2_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$D$60 | 1 |
| scenarios.Low.annual.Y3.revenue | annual_Y3_revenue | 'Annual PL Y1 to Y5'!$E$69 | 1 |
| scenarios.Low.annual.Y3.net | annual_Y3_net | 'Annual PL Y1 to Y5'!$E$80 | 1 |
| scenarios.Low.annual.Y3.net_margin | annual_Y3_net_margin | 'Annual PL Y1 to Y5'!$E$81 | 1 |
| scenarios.Low.annual.Y3.cogs | annual_Y3_cogs | 'Annual PL Y1 to Y5'!$E$70 | 1 |
| scenarios.Low.annual.Y3.fees | annual_Y3_fees | 'Annual PL Y1 to Y5'!$E$71 | 1 |
| scenarios.Low.annual.Y3.fulfillment | annual_Y3_fulfillment | 'Annual PL Y1 to Y5'!$E$72 | 1 |
| scenarios.Low.annual.Y3.advertising | annual_Y3_advertising | 'Annual PL Y1 to Y5'!$E$73 | 1 |
| scenarios.Low.annual.Y3.returns | annual_Y3_returns | 'Annual PL Y1 to Y5'!$E$74 | 1 |
| scenarios.Low.annual.Y3.contribution | annual_Y3_contribution | 'Annual PL Y1 to Y5'!$E$75 | 1 |
| scenarios.Low.annual.Y3.fixed | annual_Y3_fixed | 'Annual PL Y1 to Y5'!$E$76 | 1 |
| scenarios.Low.annual.Y3.seller_fee | annual_Y3_seller_fee | 'Annual PL Y1 to Y5'!$E$77 | 1 |
| scenarios.Low.annual.Y3.launch | annual_Y3_launch | 'Annual PL Y1 to Y5'!$E$78 | 1 |
| scenarios.Low.annual.Y3.compliance | annual_Y3_compliance | 'Annual PL Y1 to Y5'!$E$79 | 1 |
| scenarios.Low.annual.Y3.cum_net | annual_Y3_cum_net | 'Annual PL Y1 to Y5'!$E$82 | 1 |
| scenarios.Low.annual.Y3.amazon_units | annual_Y3_amazon_units | 'Annual PL Y1 to Y5'!$E$6 | 1 |
| scenarios.Low.annual.Y3.asp | annual_Y3_asp | 'Annual PL Y1 to Y5'!$E$7 | 1 |
| scenarios.Low.annual.Y3.landed_pct | annual_Y3_landed_pct | 'Annual PL Y1 to Y5'!$E$8 | 1 |
| scenarios.Low.annual.Y3.ads_pct | annual_Y3_ads_pct | 'Annual PL Y1 to Y5'!$E$9 | 1 |
| scenarios.Low.annual.Y3.amazon_contribution_pct | annual_Y3_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$E$83 | 1 |
| scenarios.Low.annual.Y3.revenue_by_channel.amazon_us | annual_Y3_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$E$15 | 1 |
| scenarios.Low.annual.Y3.revenue_by_channel.ebay_walmart | annual_Y3_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$E$24 | 1 |
| scenarios.Low.annual.Y3.revenue_by_channel.b2b | annual_Y3_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$E$33 | 1 |
| scenarios.Low.annual.Y3.revenue_by_channel.dtc | annual_Y3_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$E$42 | 1 |
| scenarios.Low.annual.Y3.revenue_by_channel.amazon_canada | annual_Y3_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$E$51 | 1 |
| scenarios.Low.annual.Y3.revenue_by_channel.pro_line | annual_Y3_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$E$60 | 1 |
| scenarios.Low.annual.Y4.revenue | annual_Y4_revenue | 'Annual PL Y1 to Y5'!$F$69 | 1 |
| scenarios.Low.annual.Y4.net | annual_Y4_net | 'Annual PL Y1 to Y5'!$F$80 | 1 |
| scenarios.Low.annual.Y4.net_margin | annual_Y4_net_margin | 'Annual PL Y1 to Y5'!$F$81 | 1 |
| scenarios.Low.annual.Y4.cogs | annual_Y4_cogs | 'Annual PL Y1 to Y5'!$F$70 | 1 |
| scenarios.Low.annual.Y4.fees | annual_Y4_fees | 'Annual PL Y1 to Y5'!$F$71 | 1 |
| scenarios.Low.annual.Y4.fulfillment | annual_Y4_fulfillment | 'Annual PL Y1 to Y5'!$F$72 | 1 |
| scenarios.Low.annual.Y4.advertising | annual_Y4_advertising | 'Annual PL Y1 to Y5'!$F$73 | 1 |
| scenarios.Low.annual.Y4.returns | annual_Y4_returns | 'Annual PL Y1 to Y5'!$F$74 | 1 |
| scenarios.Low.annual.Y4.contribution | annual_Y4_contribution | 'Annual PL Y1 to Y5'!$F$75 | 1 |
| scenarios.Low.annual.Y4.fixed | annual_Y4_fixed | 'Annual PL Y1 to Y5'!$F$76 | 1 |
| scenarios.Low.annual.Y4.seller_fee | annual_Y4_seller_fee | 'Annual PL Y1 to Y5'!$F$77 | 1 |
| scenarios.Low.annual.Y4.launch | annual_Y4_launch | 'Annual PL Y1 to Y5'!$F$78 | 1 |
| scenarios.Low.annual.Y4.compliance | annual_Y4_compliance | 'Annual PL Y1 to Y5'!$F$79 | 1 |
| scenarios.Low.annual.Y4.cum_net | annual_Y4_cum_net | 'Annual PL Y1 to Y5'!$F$82 | 1 |
| scenarios.Low.annual.Y4.amazon_units | annual_Y4_amazon_units | 'Annual PL Y1 to Y5'!$F$6 | 1 |
| scenarios.Low.annual.Y4.asp | annual_Y4_asp | 'Annual PL Y1 to Y5'!$F$7 | 1 |
| scenarios.Low.annual.Y4.landed_pct | annual_Y4_landed_pct | 'Annual PL Y1 to Y5'!$F$8 | 1 |
| scenarios.Low.annual.Y4.ads_pct | annual_Y4_ads_pct | 'Annual PL Y1 to Y5'!$F$9 | 1 |
| scenarios.Low.annual.Y4.amazon_contribution_pct | annual_Y4_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$F$83 | 1 |
| scenarios.Low.annual.Y4.revenue_by_channel.amazon_us | annual_Y4_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$F$15 | 1 |
| scenarios.Low.annual.Y4.revenue_by_channel.ebay_walmart | annual_Y4_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$F$24 | 1 |
| scenarios.Low.annual.Y4.revenue_by_channel.b2b | annual_Y4_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$F$33 | 1 |
| scenarios.Low.annual.Y4.revenue_by_channel.dtc | annual_Y4_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$F$42 | 1 |
| scenarios.Low.annual.Y4.revenue_by_channel.amazon_canada | annual_Y4_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$F$51 | 1 |
| scenarios.Low.annual.Y4.revenue_by_channel.pro_line | annual_Y4_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$F$60 | 1 |
| scenarios.Low.annual.Y5.revenue | annual_Y5_revenue | 'Annual PL Y1 to Y5'!$G$69 | 1 |
| scenarios.Low.annual.Y5.net | annual_Y5_net | 'Annual PL Y1 to Y5'!$G$80 | 1 |
| scenarios.Low.annual.Y5.net_margin | annual_Y5_net_margin | 'Annual PL Y1 to Y5'!$G$81 | 1 |
| scenarios.Low.annual.Y5.cogs | annual_Y5_cogs | 'Annual PL Y1 to Y5'!$G$70 | 1 |
| scenarios.Low.annual.Y5.fees | annual_Y5_fees | 'Annual PL Y1 to Y5'!$G$71 | 1 |
| scenarios.Low.annual.Y5.fulfillment | annual_Y5_fulfillment | 'Annual PL Y1 to Y5'!$G$72 | 1 |
| scenarios.Low.annual.Y5.advertising | annual_Y5_advertising | 'Annual PL Y1 to Y5'!$G$73 | 1 |
| scenarios.Low.annual.Y5.returns | annual_Y5_returns | 'Annual PL Y1 to Y5'!$G$74 | 1 |
| scenarios.Low.annual.Y5.contribution | annual_Y5_contribution | 'Annual PL Y1 to Y5'!$G$75 | 1 |
| scenarios.Low.annual.Y5.fixed | annual_Y5_fixed | 'Annual PL Y1 to Y5'!$G$76 | 1 |
| scenarios.Low.annual.Y5.seller_fee | annual_Y5_seller_fee | 'Annual PL Y1 to Y5'!$G$77 | 1 |
| scenarios.Low.annual.Y5.launch | annual_Y5_launch | 'Annual PL Y1 to Y5'!$G$78 | 1 |
| scenarios.Low.annual.Y5.compliance | annual_Y5_compliance | 'Annual PL Y1 to Y5'!$G$79 | 1 |
| scenarios.Low.annual.Y5.cum_net | annual_Y5_cum_net | 'Annual PL Y1 to Y5'!$G$82 | 1 |
| scenarios.Low.annual.Y5.amazon_units | annual_Y5_amazon_units | 'Annual PL Y1 to Y5'!$G$6 | 1 |
| scenarios.Low.annual.Y5.asp | annual_Y5_asp | 'Annual PL Y1 to Y5'!$G$7 | 1 |
| scenarios.Low.annual.Y5.landed_pct | annual_Y5_landed_pct | 'Annual PL Y1 to Y5'!$G$8 | 1 |
| scenarios.Low.annual.Y5.ads_pct | annual_Y5_ads_pct | 'Annual PL Y1 to Y5'!$G$9 | 1 |
| scenarios.Low.annual.Y5.amazon_contribution_pct | annual_Y5_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$G$83 | 1 |
| scenarios.Low.annual.Y5.revenue_by_channel.amazon_us | annual_Y5_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$G$15 | 1 |
| scenarios.Low.annual.Y5.revenue_by_channel.ebay_walmart | annual_Y5_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$G$24 | 1 |
| scenarios.Low.annual.Y5.revenue_by_channel.b2b | annual_Y5_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$G$33 | 1 |
| scenarios.Low.annual.Y5.revenue_by_channel.dtc | annual_Y5_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$G$42 | 1 |
| scenarios.Low.annual.Y5.revenue_by_channel.amazon_canada | annual_Y5_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$G$51 | 1 |
| scenarios.Low.annual.Y5.revenue_by_channel.pro_line | annual_Y5_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$G$60 | 1 |
| scenarios.Low.cash.peak_need | cash_peak_need | 'Cash Flow and Working Capital'!$C$48 | 1 |
| scenarios.Low.cash.peak_need_month | cash_peak_need_month | 'Cash Flow and Working Capital'!$C$47 | 1 |
| scenarios.Low.cash.peak_need_month_label | cash_peak_need_month_label | 'Cash Flow and Working Capital'!$C$46 | 1 |
| scenarios.Low.cash.min_balance | cash_min_balance | 'Cash Flow and Working Capital'!$C$45 | 1 |
| scenarios.Low.cash.funding_gap | cash_funding_gap | 'Cash Flow and Working Capital'!$C$49 | 1 |
| scenarios.Low.cash.funding_needed_by_month | cash_funding_needed_by_month | 'Cash Flow and Working Capital'!$C$50 | 1 |
| scenarios.Low.cash.runway_months_day90 | cash_runway_months_day90 | 'Cash Flow and Working Capital'!$C$55 | 1 |
| scenarios.Low.cash.runway_months_day90_unfunded | cash_runway_months_day90_unfunded | 'Cash Flow and Working Capital'!$C$52 | 1 |
| scenarios.Low.cash.opening_capital | cash_opening_capital | 'Assumptions'!$F$106 | 1 |
| scenarios.Low.gates.day90.metrics.tacos.value | gate_d90_tacos_value | 'Funding and Breakeven'!$B$28 | 1 |
| scenarios.Low.gates.day90.metrics.ret.value | gate_d90_ret_value | 'Funding and Breakeven'!$B$29 | 1 |
| scenarios.Low.gates.day90.metrics.cpu.value | gate_d90_cpu_value | 'Funding and Breakeven'!$B$30 | 1 |
| scenarios.Low.gates.day90.metrics.runway.value | gate_d90_runway_value | 'Funding and Breakeven'!$B$31 | 1 |
| scenarios.Low.gates.day90.result | gate_day90_result | 'Funding and Breakeven'!$B$33 | 1 |
| scenarios.Low.gates.month12.metrics.rev.value | gate_m12_rev_value | 'Funding and Breakeven'!$B$36 | 1 |
| scenarios.Low.gates.month12.metrics.cum.value | gate_m12_cum_value | 'Funding and Breakeven'!$B$37 | 1 |
| scenarios.Low.gates.month12.metrics.tacos.value | gate_m12_tacos_value | 'Funding and Breakeven'!$B$38 | 1 |
| scenarios.Low.gates.month12.metrics.ret.value | gate_m12_ret_value | 'Funding and Breakeven'!$B$39 | 1 |
| scenarios.Low.gates.month12.metrics.opcf.value | gate_m12_opcf_value | 'Funding and Breakeven'!$B$40 | 1 |
| scenarios.Low.gates.month12.result | gate_month12_result | 'Funding and Breakeven'!$B$41 | 1 |
| scenarios.Medium.m18.revenue | m18_revenue | 'Monthly PL Y1 Y2'!$C$60 | 2 |
| scenarios.Medium.m18.net | m18_net | 'Monthly PL Y1 Y2'!$C$61 | 2 |
| scenarios.Medium.m18.net_margin | m18_net_margin | 'Monthly PL Y1 Y2'!$C$62 | 2 |
| scenarios.Medium.m18.breakeven_month | m18_breakeven_month | 'Monthly PL Y1 Y2'!$C$63 | 2 |
| scenarios.Medium.m18.run_rate | m18_run_rate | 'Monthly PL Y1 Y2'!$C$64 | 2 |
| scenarios.Medium.m18.exit_margin | m18_exit_margin | 'Monthly PL Y1 Y2'!$C$65 | 2 |
| scenarios.Medium.m18.units | m18_units | 'Monthly PL Y1 Y2'!$C$66 | 2 |
| scenarios.Medium.m24.revenue | m24_revenue | 'Monthly PL Y1 Y2'!$C$68 | 2 |
| scenarios.Medium.m24.net | m24_net | 'Monthly PL Y1 Y2'!$C$69 | 2 |
| scenarios.Medium.m24.breakeven_month | m24_breakeven_month | 'Monthly PL Y1 Y2'!$C$67 | 2 |
| scenarios.Medium.monthly[].units | monthly_units | 'Monthly PL Y1 Y2'!$D$7:$AA$7 | 2 |
| scenarios.Medium.monthly[].revenue | monthly_revenue | 'Monthly PL Y1 Y2'!$D$54:$AA$54 | 2 |
| scenarios.Medium.monthly[].net | monthly_net | 'Monthly PL Y1 Y2'!$D$55:$AA$55 | 2 |
| scenarios.Medium.monthly[].cum_net | monthly_cum_net | 'Monthly PL Y1 Y2'!$D$56:$AA$56 | 2 |
| scenarios.Medium.monthly[].cash_balance | monthly_cash_balance | 'Cash Flow and Working Capital'!$C$33:$AA$33 | 2 |
| scenarios.Medium.monthly[].tacos | monthly_tacos | 'Monthly PL Y1 Y2'!$D$24:$AA$24 | 2 |
| scenarios.Medium.monthly[].contribution_before_ads | monthly_contribution_before_ads | 'Monthly PL Y1 Y2'!$D$25:$AA$25 | 2 |
| scenarios.Medium.monthly[].contribution_after_ads | monthly_contribution_after_ads | 'Monthly PL Y1 Y2'!$D$26:$AA$26 | 2 |
| scenarios.Medium.monthly[].amazon_revenue | monthly_amazon_revenue | 'Monthly PL Y1 Y2'!$D$9:$AA$9 | 2 |
| scenarios.Medium.monthly[].amazon_net | monthly_amazon_net | 'Monthly PL Y1 Y2'!$D$21:$AA$21 | 2 |
| scenarios.Medium.monthly[].cash_balance_after_funding | monthly_cash_balance_after_funding | 'Cash Flow and Working Capital'!$C$54:$AA$54 | 2 |
| scenarios.Medium.monthly[].inventory_on_hand | monthly_inventory_on_hand | 'Cash Flow and Working Capital'!$C$37:$AA$37 | 2 |
| scenarios.Medium.monthly[].inventory_in_transit | monthly_inventory_in_transit | 'Cash Flow and Working Capital'!$C$38:$AA$38 | 2 |
| scenarios.Medium.annual.Y1.revenue | annual_Y1_revenue | 'Annual PL Y1 to Y5'!$C$69 | 2 |
| scenarios.Medium.annual.Y1.net | annual_Y1_net | 'Annual PL Y1 to Y5'!$C$80 | 2 |
| scenarios.Medium.annual.Y1.net_margin | annual_Y1_net_margin | 'Annual PL Y1 to Y5'!$C$81 | 2 |
| scenarios.Medium.annual.Y1.cogs | annual_Y1_cogs | 'Annual PL Y1 to Y5'!$C$70 | 2 |
| scenarios.Medium.annual.Y1.fees | annual_Y1_fees | 'Annual PL Y1 to Y5'!$C$71 | 2 |
| scenarios.Medium.annual.Y1.fulfillment | annual_Y1_fulfillment | 'Annual PL Y1 to Y5'!$C$72 | 2 |
| scenarios.Medium.annual.Y1.advertising | annual_Y1_advertising | 'Annual PL Y1 to Y5'!$C$73 | 2 |
| scenarios.Medium.annual.Y1.returns | annual_Y1_returns | 'Annual PL Y1 to Y5'!$C$74 | 2 |
| scenarios.Medium.annual.Y1.contribution | annual_Y1_contribution | 'Annual PL Y1 to Y5'!$C$75 | 2 |
| scenarios.Medium.annual.Y1.fixed | annual_Y1_fixed | 'Annual PL Y1 to Y5'!$C$76 | 2 |
| scenarios.Medium.annual.Y1.seller_fee | annual_Y1_seller_fee | 'Annual PL Y1 to Y5'!$C$77 | 2 |
| scenarios.Medium.annual.Y1.launch | annual_Y1_launch | 'Annual PL Y1 to Y5'!$C$78 | 2 |
| scenarios.Medium.annual.Y1.compliance | annual_Y1_compliance | 'Annual PL Y1 to Y5'!$C$79 | 2 |
| scenarios.Medium.annual.Y1.cum_net | annual_Y1_cum_net | 'Annual PL Y1 to Y5'!$C$82 | 2 |
| scenarios.Medium.annual.Y1.amazon_units | annual_Y1_amazon_units | 'Annual PL Y1 to Y5'!$C$6 | 2 |
| scenarios.Medium.annual.Y1.asp | annual_Y1_asp | 'Annual PL Y1 to Y5'!$C$7 | 2 |
| scenarios.Medium.annual.Y1.landed_pct | annual_Y1_landed_pct | 'Annual PL Y1 to Y5'!$C$8 | 2 |
| scenarios.Medium.annual.Y1.ads_pct | annual_Y1_ads_pct | 'Annual PL Y1 to Y5'!$C$9 | 2 |
| scenarios.Medium.annual.Y1.amazon_contribution_pct | annual_Y1_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$C$83 | 2 |
| scenarios.Medium.annual.Y1.revenue_by_channel.amazon_us | annual_Y1_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$C$15 | 2 |
| scenarios.Medium.annual.Y1.revenue_by_channel.ebay_walmart | annual_Y1_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$C$24 | 2 |
| scenarios.Medium.annual.Y1.revenue_by_channel.b2b | annual_Y1_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$C$33 | 2 |
| scenarios.Medium.annual.Y1.revenue_by_channel.dtc | annual_Y1_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$C$42 | 2 |
| scenarios.Medium.annual.Y1.revenue_by_channel.amazon_canada | annual_Y1_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$C$51 | 2 |
| scenarios.Medium.annual.Y1.revenue_by_channel.pro_line | annual_Y1_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$C$60 | 2 |
| scenarios.Medium.annual.Y2.revenue | annual_Y2_revenue | 'Annual PL Y1 to Y5'!$D$69 | 2 |
| scenarios.Medium.annual.Y2.net | annual_Y2_net | 'Annual PL Y1 to Y5'!$D$80 | 2 |
| scenarios.Medium.annual.Y2.net_margin | annual_Y2_net_margin | 'Annual PL Y1 to Y5'!$D$81 | 2 |
| scenarios.Medium.annual.Y2.cogs | annual_Y2_cogs | 'Annual PL Y1 to Y5'!$D$70 | 2 |
| scenarios.Medium.annual.Y2.fees | annual_Y2_fees | 'Annual PL Y1 to Y5'!$D$71 | 2 |
| scenarios.Medium.annual.Y2.fulfillment | annual_Y2_fulfillment | 'Annual PL Y1 to Y5'!$D$72 | 2 |
| scenarios.Medium.annual.Y2.advertising | annual_Y2_advertising | 'Annual PL Y1 to Y5'!$D$73 | 2 |
| scenarios.Medium.annual.Y2.returns | annual_Y2_returns | 'Annual PL Y1 to Y5'!$D$74 | 2 |
| scenarios.Medium.annual.Y2.contribution | annual_Y2_contribution | 'Annual PL Y1 to Y5'!$D$75 | 2 |
| scenarios.Medium.annual.Y2.fixed | annual_Y2_fixed | 'Annual PL Y1 to Y5'!$D$76 | 2 |
| scenarios.Medium.annual.Y2.seller_fee | annual_Y2_seller_fee | 'Annual PL Y1 to Y5'!$D$77 | 2 |
| scenarios.Medium.annual.Y2.launch | annual_Y2_launch | 'Annual PL Y1 to Y5'!$D$78 | 2 |
| scenarios.Medium.annual.Y2.compliance | annual_Y2_compliance | 'Annual PL Y1 to Y5'!$D$79 | 2 |
| scenarios.Medium.annual.Y2.cum_net | annual_Y2_cum_net | 'Annual PL Y1 to Y5'!$D$82 | 2 |
| scenarios.Medium.annual.Y2.amazon_units | annual_Y2_amazon_units | 'Annual PL Y1 to Y5'!$D$6 | 2 |
| scenarios.Medium.annual.Y2.asp | annual_Y2_asp | 'Annual PL Y1 to Y5'!$D$7 | 2 |
| scenarios.Medium.annual.Y2.landed_pct | annual_Y2_landed_pct | 'Annual PL Y1 to Y5'!$D$8 | 2 |
| scenarios.Medium.annual.Y2.ads_pct | annual_Y2_ads_pct | 'Annual PL Y1 to Y5'!$D$9 | 2 |
| scenarios.Medium.annual.Y2.amazon_contribution_pct | annual_Y2_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$D$83 | 2 |
| scenarios.Medium.annual.Y2.revenue_by_channel.amazon_us | annual_Y2_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$D$15 | 2 |
| scenarios.Medium.annual.Y2.revenue_by_channel.ebay_walmart | annual_Y2_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$D$24 | 2 |
| scenarios.Medium.annual.Y2.revenue_by_channel.b2b | annual_Y2_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$D$33 | 2 |
| scenarios.Medium.annual.Y2.revenue_by_channel.dtc | annual_Y2_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$D$42 | 2 |
| scenarios.Medium.annual.Y2.revenue_by_channel.amazon_canada | annual_Y2_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$D$51 | 2 |
| scenarios.Medium.annual.Y2.revenue_by_channel.pro_line | annual_Y2_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$D$60 | 2 |
| scenarios.Medium.annual.Y3.revenue | annual_Y3_revenue | 'Annual PL Y1 to Y5'!$E$69 | 2 |
| scenarios.Medium.annual.Y3.net | annual_Y3_net | 'Annual PL Y1 to Y5'!$E$80 | 2 |
| scenarios.Medium.annual.Y3.net_margin | annual_Y3_net_margin | 'Annual PL Y1 to Y5'!$E$81 | 2 |
| scenarios.Medium.annual.Y3.cogs | annual_Y3_cogs | 'Annual PL Y1 to Y5'!$E$70 | 2 |
| scenarios.Medium.annual.Y3.fees | annual_Y3_fees | 'Annual PL Y1 to Y5'!$E$71 | 2 |
| scenarios.Medium.annual.Y3.fulfillment | annual_Y3_fulfillment | 'Annual PL Y1 to Y5'!$E$72 | 2 |
| scenarios.Medium.annual.Y3.advertising | annual_Y3_advertising | 'Annual PL Y1 to Y5'!$E$73 | 2 |
| scenarios.Medium.annual.Y3.returns | annual_Y3_returns | 'Annual PL Y1 to Y5'!$E$74 | 2 |
| scenarios.Medium.annual.Y3.contribution | annual_Y3_contribution | 'Annual PL Y1 to Y5'!$E$75 | 2 |
| scenarios.Medium.annual.Y3.fixed | annual_Y3_fixed | 'Annual PL Y1 to Y5'!$E$76 | 2 |
| scenarios.Medium.annual.Y3.seller_fee | annual_Y3_seller_fee | 'Annual PL Y1 to Y5'!$E$77 | 2 |
| scenarios.Medium.annual.Y3.launch | annual_Y3_launch | 'Annual PL Y1 to Y5'!$E$78 | 2 |
| scenarios.Medium.annual.Y3.compliance | annual_Y3_compliance | 'Annual PL Y1 to Y5'!$E$79 | 2 |
| scenarios.Medium.annual.Y3.cum_net | annual_Y3_cum_net | 'Annual PL Y1 to Y5'!$E$82 | 2 |
| scenarios.Medium.annual.Y3.amazon_units | annual_Y3_amazon_units | 'Annual PL Y1 to Y5'!$E$6 | 2 |
| scenarios.Medium.annual.Y3.asp | annual_Y3_asp | 'Annual PL Y1 to Y5'!$E$7 | 2 |
| scenarios.Medium.annual.Y3.landed_pct | annual_Y3_landed_pct | 'Annual PL Y1 to Y5'!$E$8 | 2 |
| scenarios.Medium.annual.Y3.ads_pct | annual_Y3_ads_pct | 'Annual PL Y1 to Y5'!$E$9 | 2 |
| scenarios.Medium.annual.Y3.amazon_contribution_pct | annual_Y3_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$E$83 | 2 |
| scenarios.Medium.annual.Y3.revenue_by_channel.amazon_us | annual_Y3_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$E$15 | 2 |
| scenarios.Medium.annual.Y3.revenue_by_channel.ebay_walmart | annual_Y3_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$E$24 | 2 |
| scenarios.Medium.annual.Y3.revenue_by_channel.b2b | annual_Y3_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$E$33 | 2 |
| scenarios.Medium.annual.Y3.revenue_by_channel.dtc | annual_Y3_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$E$42 | 2 |
| scenarios.Medium.annual.Y3.revenue_by_channel.amazon_canada | annual_Y3_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$E$51 | 2 |
| scenarios.Medium.annual.Y3.revenue_by_channel.pro_line | annual_Y3_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$E$60 | 2 |
| scenarios.Medium.annual.Y4.revenue | annual_Y4_revenue | 'Annual PL Y1 to Y5'!$F$69 | 2 |
| scenarios.Medium.annual.Y4.net | annual_Y4_net | 'Annual PL Y1 to Y5'!$F$80 | 2 |
| scenarios.Medium.annual.Y4.net_margin | annual_Y4_net_margin | 'Annual PL Y1 to Y5'!$F$81 | 2 |
| scenarios.Medium.annual.Y4.cogs | annual_Y4_cogs | 'Annual PL Y1 to Y5'!$F$70 | 2 |
| scenarios.Medium.annual.Y4.fees | annual_Y4_fees | 'Annual PL Y1 to Y5'!$F$71 | 2 |
| scenarios.Medium.annual.Y4.fulfillment | annual_Y4_fulfillment | 'Annual PL Y1 to Y5'!$F$72 | 2 |
| scenarios.Medium.annual.Y4.advertising | annual_Y4_advertising | 'Annual PL Y1 to Y5'!$F$73 | 2 |
| scenarios.Medium.annual.Y4.returns | annual_Y4_returns | 'Annual PL Y1 to Y5'!$F$74 | 2 |
| scenarios.Medium.annual.Y4.contribution | annual_Y4_contribution | 'Annual PL Y1 to Y5'!$F$75 | 2 |
| scenarios.Medium.annual.Y4.fixed | annual_Y4_fixed | 'Annual PL Y1 to Y5'!$F$76 | 2 |
| scenarios.Medium.annual.Y4.seller_fee | annual_Y4_seller_fee | 'Annual PL Y1 to Y5'!$F$77 | 2 |
| scenarios.Medium.annual.Y4.launch | annual_Y4_launch | 'Annual PL Y1 to Y5'!$F$78 | 2 |
| scenarios.Medium.annual.Y4.compliance | annual_Y4_compliance | 'Annual PL Y1 to Y5'!$F$79 | 2 |
| scenarios.Medium.annual.Y4.cum_net | annual_Y4_cum_net | 'Annual PL Y1 to Y5'!$F$82 | 2 |
| scenarios.Medium.annual.Y4.amazon_units | annual_Y4_amazon_units | 'Annual PL Y1 to Y5'!$F$6 | 2 |
| scenarios.Medium.annual.Y4.asp | annual_Y4_asp | 'Annual PL Y1 to Y5'!$F$7 | 2 |
| scenarios.Medium.annual.Y4.landed_pct | annual_Y4_landed_pct | 'Annual PL Y1 to Y5'!$F$8 | 2 |
| scenarios.Medium.annual.Y4.ads_pct | annual_Y4_ads_pct | 'Annual PL Y1 to Y5'!$F$9 | 2 |
| scenarios.Medium.annual.Y4.amazon_contribution_pct | annual_Y4_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$F$83 | 2 |
| scenarios.Medium.annual.Y4.revenue_by_channel.amazon_us | annual_Y4_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$F$15 | 2 |
| scenarios.Medium.annual.Y4.revenue_by_channel.ebay_walmart | annual_Y4_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$F$24 | 2 |
| scenarios.Medium.annual.Y4.revenue_by_channel.b2b | annual_Y4_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$F$33 | 2 |
| scenarios.Medium.annual.Y4.revenue_by_channel.dtc | annual_Y4_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$F$42 | 2 |
| scenarios.Medium.annual.Y4.revenue_by_channel.amazon_canada | annual_Y4_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$F$51 | 2 |
| scenarios.Medium.annual.Y4.revenue_by_channel.pro_line | annual_Y4_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$F$60 | 2 |
| scenarios.Medium.annual.Y5.revenue | annual_Y5_revenue | 'Annual PL Y1 to Y5'!$G$69 | 2 |
| scenarios.Medium.annual.Y5.net | annual_Y5_net | 'Annual PL Y1 to Y5'!$G$80 | 2 |
| scenarios.Medium.annual.Y5.net_margin | annual_Y5_net_margin | 'Annual PL Y1 to Y5'!$G$81 | 2 |
| scenarios.Medium.annual.Y5.cogs | annual_Y5_cogs | 'Annual PL Y1 to Y5'!$G$70 | 2 |
| scenarios.Medium.annual.Y5.fees | annual_Y5_fees | 'Annual PL Y1 to Y5'!$G$71 | 2 |
| scenarios.Medium.annual.Y5.fulfillment | annual_Y5_fulfillment | 'Annual PL Y1 to Y5'!$G$72 | 2 |
| scenarios.Medium.annual.Y5.advertising | annual_Y5_advertising | 'Annual PL Y1 to Y5'!$G$73 | 2 |
| scenarios.Medium.annual.Y5.returns | annual_Y5_returns | 'Annual PL Y1 to Y5'!$G$74 | 2 |
| scenarios.Medium.annual.Y5.contribution | annual_Y5_contribution | 'Annual PL Y1 to Y5'!$G$75 | 2 |
| scenarios.Medium.annual.Y5.fixed | annual_Y5_fixed | 'Annual PL Y1 to Y5'!$G$76 | 2 |
| scenarios.Medium.annual.Y5.seller_fee | annual_Y5_seller_fee | 'Annual PL Y1 to Y5'!$G$77 | 2 |
| scenarios.Medium.annual.Y5.launch | annual_Y5_launch | 'Annual PL Y1 to Y5'!$G$78 | 2 |
| scenarios.Medium.annual.Y5.compliance | annual_Y5_compliance | 'Annual PL Y1 to Y5'!$G$79 | 2 |
| scenarios.Medium.annual.Y5.cum_net | annual_Y5_cum_net | 'Annual PL Y1 to Y5'!$G$82 | 2 |
| scenarios.Medium.annual.Y5.amazon_units | annual_Y5_amazon_units | 'Annual PL Y1 to Y5'!$G$6 | 2 |
| scenarios.Medium.annual.Y5.asp | annual_Y5_asp | 'Annual PL Y1 to Y5'!$G$7 | 2 |
| scenarios.Medium.annual.Y5.landed_pct | annual_Y5_landed_pct | 'Annual PL Y1 to Y5'!$G$8 | 2 |
| scenarios.Medium.annual.Y5.ads_pct | annual_Y5_ads_pct | 'Annual PL Y1 to Y5'!$G$9 | 2 |
| scenarios.Medium.annual.Y5.amazon_contribution_pct | annual_Y5_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$G$83 | 2 |
| scenarios.Medium.annual.Y5.revenue_by_channel.amazon_us | annual_Y5_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$G$15 | 2 |
| scenarios.Medium.annual.Y5.revenue_by_channel.ebay_walmart | annual_Y5_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$G$24 | 2 |
| scenarios.Medium.annual.Y5.revenue_by_channel.b2b | annual_Y5_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$G$33 | 2 |
| scenarios.Medium.annual.Y5.revenue_by_channel.dtc | annual_Y5_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$G$42 | 2 |
| scenarios.Medium.annual.Y5.revenue_by_channel.amazon_canada | annual_Y5_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$G$51 | 2 |
| scenarios.Medium.annual.Y5.revenue_by_channel.pro_line | annual_Y5_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$G$60 | 2 |
| scenarios.Medium.cash.peak_need | cash_peak_need | 'Cash Flow and Working Capital'!$C$48 | 2 |
| scenarios.Medium.cash.peak_need_month | cash_peak_need_month | 'Cash Flow and Working Capital'!$C$47 | 2 |
| scenarios.Medium.cash.peak_need_month_label | cash_peak_need_month_label | 'Cash Flow and Working Capital'!$C$46 | 2 |
| scenarios.Medium.cash.min_balance | cash_min_balance | 'Cash Flow and Working Capital'!$C$45 | 2 |
| scenarios.Medium.cash.funding_gap | cash_funding_gap | 'Cash Flow and Working Capital'!$C$49 | 2 |
| scenarios.Medium.cash.funding_needed_by_month | cash_funding_needed_by_month | 'Cash Flow and Working Capital'!$C$50 | 2 |
| scenarios.Medium.cash.runway_months_day90 | cash_runway_months_day90 | 'Cash Flow and Working Capital'!$C$55 | 2 |
| scenarios.Medium.cash.runway_months_day90_unfunded | cash_runway_months_day90_unfunded | 'Cash Flow and Working Capital'!$C$52 | 2 |
| scenarios.Medium.cash.opening_capital | cash_opening_capital | 'Assumptions'!$F$106 | 2 |
| scenarios.Medium.gates.day90.metrics.tacos.value | gate_d90_tacos_value | 'Funding and Breakeven'!$B$28 | 2 |
| scenarios.Medium.gates.day90.metrics.ret.value | gate_d90_ret_value | 'Funding and Breakeven'!$B$29 | 2 |
| scenarios.Medium.gates.day90.metrics.cpu.value | gate_d90_cpu_value | 'Funding and Breakeven'!$B$30 | 2 |
| scenarios.Medium.gates.day90.metrics.runway.value | gate_d90_runway_value | 'Funding and Breakeven'!$B$31 | 2 |
| scenarios.Medium.gates.day90.result | gate_day90_result | 'Funding and Breakeven'!$B$33 | 2 |
| scenarios.Medium.gates.month12.metrics.rev.value | gate_m12_rev_value | 'Funding and Breakeven'!$B$36 | 2 |
| scenarios.Medium.gates.month12.metrics.cum.value | gate_m12_cum_value | 'Funding and Breakeven'!$B$37 | 2 |
| scenarios.Medium.gates.month12.metrics.tacos.value | gate_m12_tacos_value | 'Funding and Breakeven'!$B$38 | 2 |
| scenarios.Medium.gates.month12.metrics.ret.value | gate_m12_ret_value | 'Funding and Breakeven'!$B$39 | 2 |
| scenarios.Medium.gates.month12.metrics.opcf.value | gate_m12_opcf_value | 'Funding and Breakeven'!$B$40 | 2 |
| scenarios.Medium.gates.month12.result | gate_month12_result | 'Funding and Breakeven'!$B$41 | 2 |
| scenarios.High.m18.revenue | m18_revenue | 'Monthly PL Y1 Y2'!$C$60 | 3 |
| scenarios.High.m18.net | m18_net | 'Monthly PL Y1 Y2'!$C$61 | 3 |
| scenarios.High.m18.net_margin | m18_net_margin | 'Monthly PL Y1 Y2'!$C$62 | 3 |
| scenarios.High.m18.breakeven_month | m18_breakeven_month | 'Monthly PL Y1 Y2'!$C$63 | 3 |
| scenarios.High.m18.run_rate | m18_run_rate | 'Monthly PL Y1 Y2'!$C$64 | 3 |
| scenarios.High.m18.exit_margin | m18_exit_margin | 'Monthly PL Y1 Y2'!$C$65 | 3 |
| scenarios.High.m18.units | m18_units | 'Monthly PL Y1 Y2'!$C$66 | 3 |
| scenarios.High.m24.revenue | m24_revenue | 'Monthly PL Y1 Y2'!$C$68 | 3 |
| scenarios.High.m24.net | m24_net | 'Monthly PL Y1 Y2'!$C$69 | 3 |
| scenarios.High.m24.breakeven_month | m24_breakeven_month | 'Monthly PL Y1 Y2'!$C$67 | 3 |
| scenarios.High.monthly[].units | monthly_units | 'Monthly PL Y1 Y2'!$D$7:$AA$7 | 3 |
| scenarios.High.monthly[].revenue | monthly_revenue | 'Monthly PL Y1 Y2'!$D$54:$AA$54 | 3 |
| scenarios.High.monthly[].net | monthly_net | 'Monthly PL Y1 Y2'!$D$55:$AA$55 | 3 |
| scenarios.High.monthly[].cum_net | monthly_cum_net | 'Monthly PL Y1 Y2'!$D$56:$AA$56 | 3 |
| scenarios.High.monthly[].cash_balance | monthly_cash_balance | 'Cash Flow and Working Capital'!$C$33:$AA$33 | 3 |
| scenarios.High.monthly[].tacos | monthly_tacos | 'Monthly PL Y1 Y2'!$D$24:$AA$24 | 3 |
| scenarios.High.monthly[].contribution_before_ads | monthly_contribution_before_ads | 'Monthly PL Y1 Y2'!$D$25:$AA$25 | 3 |
| scenarios.High.monthly[].contribution_after_ads | monthly_contribution_after_ads | 'Monthly PL Y1 Y2'!$D$26:$AA$26 | 3 |
| scenarios.High.monthly[].amazon_revenue | monthly_amazon_revenue | 'Monthly PL Y1 Y2'!$D$9:$AA$9 | 3 |
| scenarios.High.monthly[].amazon_net | monthly_amazon_net | 'Monthly PL Y1 Y2'!$D$21:$AA$21 | 3 |
| scenarios.High.monthly[].cash_balance_after_funding | monthly_cash_balance_after_funding | 'Cash Flow and Working Capital'!$C$54:$AA$54 | 3 |
| scenarios.High.monthly[].inventory_on_hand | monthly_inventory_on_hand | 'Cash Flow and Working Capital'!$C$37:$AA$37 | 3 |
| scenarios.High.monthly[].inventory_in_transit | monthly_inventory_in_transit | 'Cash Flow and Working Capital'!$C$38:$AA$38 | 3 |
| scenarios.High.annual.Y1.revenue | annual_Y1_revenue | 'Annual PL Y1 to Y5'!$C$69 | 3 |
| scenarios.High.annual.Y1.net | annual_Y1_net | 'Annual PL Y1 to Y5'!$C$80 | 3 |
| scenarios.High.annual.Y1.net_margin | annual_Y1_net_margin | 'Annual PL Y1 to Y5'!$C$81 | 3 |
| scenarios.High.annual.Y1.cogs | annual_Y1_cogs | 'Annual PL Y1 to Y5'!$C$70 | 3 |
| scenarios.High.annual.Y1.fees | annual_Y1_fees | 'Annual PL Y1 to Y5'!$C$71 | 3 |
| scenarios.High.annual.Y1.fulfillment | annual_Y1_fulfillment | 'Annual PL Y1 to Y5'!$C$72 | 3 |
| scenarios.High.annual.Y1.advertising | annual_Y1_advertising | 'Annual PL Y1 to Y5'!$C$73 | 3 |
| scenarios.High.annual.Y1.returns | annual_Y1_returns | 'Annual PL Y1 to Y5'!$C$74 | 3 |
| scenarios.High.annual.Y1.contribution | annual_Y1_contribution | 'Annual PL Y1 to Y5'!$C$75 | 3 |
| scenarios.High.annual.Y1.fixed | annual_Y1_fixed | 'Annual PL Y1 to Y5'!$C$76 | 3 |
| scenarios.High.annual.Y1.seller_fee | annual_Y1_seller_fee | 'Annual PL Y1 to Y5'!$C$77 | 3 |
| scenarios.High.annual.Y1.launch | annual_Y1_launch | 'Annual PL Y1 to Y5'!$C$78 | 3 |
| scenarios.High.annual.Y1.compliance | annual_Y1_compliance | 'Annual PL Y1 to Y5'!$C$79 | 3 |
| scenarios.High.annual.Y1.cum_net | annual_Y1_cum_net | 'Annual PL Y1 to Y5'!$C$82 | 3 |
| scenarios.High.annual.Y1.amazon_units | annual_Y1_amazon_units | 'Annual PL Y1 to Y5'!$C$6 | 3 |
| scenarios.High.annual.Y1.asp | annual_Y1_asp | 'Annual PL Y1 to Y5'!$C$7 | 3 |
| scenarios.High.annual.Y1.landed_pct | annual_Y1_landed_pct | 'Annual PL Y1 to Y5'!$C$8 | 3 |
| scenarios.High.annual.Y1.ads_pct | annual_Y1_ads_pct | 'Annual PL Y1 to Y5'!$C$9 | 3 |
| scenarios.High.annual.Y1.amazon_contribution_pct | annual_Y1_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$C$83 | 3 |
| scenarios.High.annual.Y1.revenue_by_channel.amazon_us | annual_Y1_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$C$15 | 3 |
| scenarios.High.annual.Y1.revenue_by_channel.ebay_walmart | annual_Y1_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$C$24 | 3 |
| scenarios.High.annual.Y1.revenue_by_channel.b2b | annual_Y1_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$C$33 | 3 |
| scenarios.High.annual.Y1.revenue_by_channel.dtc | annual_Y1_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$C$42 | 3 |
| scenarios.High.annual.Y1.revenue_by_channel.amazon_canada | annual_Y1_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$C$51 | 3 |
| scenarios.High.annual.Y1.revenue_by_channel.pro_line | annual_Y1_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$C$60 | 3 |
| scenarios.High.annual.Y2.revenue | annual_Y2_revenue | 'Annual PL Y1 to Y5'!$D$69 | 3 |
| scenarios.High.annual.Y2.net | annual_Y2_net | 'Annual PL Y1 to Y5'!$D$80 | 3 |
| scenarios.High.annual.Y2.net_margin | annual_Y2_net_margin | 'Annual PL Y1 to Y5'!$D$81 | 3 |
| scenarios.High.annual.Y2.cogs | annual_Y2_cogs | 'Annual PL Y1 to Y5'!$D$70 | 3 |
| scenarios.High.annual.Y2.fees | annual_Y2_fees | 'Annual PL Y1 to Y5'!$D$71 | 3 |
| scenarios.High.annual.Y2.fulfillment | annual_Y2_fulfillment | 'Annual PL Y1 to Y5'!$D$72 | 3 |
| scenarios.High.annual.Y2.advertising | annual_Y2_advertising | 'Annual PL Y1 to Y5'!$D$73 | 3 |
| scenarios.High.annual.Y2.returns | annual_Y2_returns | 'Annual PL Y1 to Y5'!$D$74 | 3 |
| scenarios.High.annual.Y2.contribution | annual_Y2_contribution | 'Annual PL Y1 to Y5'!$D$75 | 3 |
| scenarios.High.annual.Y2.fixed | annual_Y2_fixed | 'Annual PL Y1 to Y5'!$D$76 | 3 |
| scenarios.High.annual.Y2.seller_fee | annual_Y2_seller_fee | 'Annual PL Y1 to Y5'!$D$77 | 3 |
| scenarios.High.annual.Y2.launch | annual_Y2_launch | 'Annual PL Y1 to Y5'!$D$78 | 3 |
| scenarios.High.annual.Y2.compliance | annual_Y2_compliance | 'Annual PL Y1 to Y5'!$D$79 | 3 |
| scenarios.High.annual.Y2.cum_net | annual_Y2_cum_net | 'Annual PL Y1 to Y5'!$D$82 | 3 |
| scenarios.High.annual.Y2.amazon_units | annual_Y2_amazon_units | 'Annual PL Y1 to Y5'!$D$6 | 3 |
| scenarios.High.annual.Y2.asp | annual_Y2_asp | 'Annual PL Y1 to Y5'!$D$7 | 3 |
| scenarios.High.annual.Y2.landed_pct | annual_Y2_landed_pct | 'Annual PL Y1 to Y5'!$D$8 | 3 |
| scenarios.High.annual.Y2.ads_pct | annual_Y2_ads_pct | 'Annual PL Y1 to Y5'!$D$9 | 3 |
| scenarios.High.annual.Y2.amazon_contribution_pct | annual_Y2_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$D$83 | 3 |
| scenarios.High.annual.Y2.revenue_by_channel.amazon_us | annual_Y2_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$D$15 | 3 |
| scenarios.High.annual.Y2.revenue_by_channel.ebay_walmart | annual_Y2_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$D$24 | 3 |
| scenarios.High.annual.Y2.revenue_by_channel.b2b | annual_Y2_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$D$33 | 3 |
| scenarios.High.annual.Y2.revenue_by_channel.dtc | annual_Y2_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$D$42 | 3 |
| scenarios.High.annual.Y2.revenue_by_channel.amazon_canada | annual_Y2_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$D$51 | 3 |
| scenarios.High.annual.Y2.revenue_by_channel.pro_line | annual_Y2_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$D$60 | 3 |
| scenarios.High.annual.Y3.revenue | annual_Y3_revenue | 'Annual PL Y1 to Y5'!$E$69 | 3 |
| scenarios.High.annual.Y3.net | annual_Y3_net | 'Annual PL Y1 to Y5'!$E$80 | 3 |
| scenarios.High.annual.Y3.net_margin | annual_Y3_net_margin | 'Annual PL Y1 to Y5'!$E$81 | 3 |
| scenarios.High.annual.Y3.cogs | annual_Y3_cogs | 'Annual PL Y1 to Y5'!$E$70 | 3 |
| scenarios.High.annual.Y3.fees | annual_Y3_fees | 'Annual PL Y1 to Y5'!$E$71 | 3 |
| scenarios.High.annual.Y3.fulfillment | annual_Y3_fulfillment | 'Annual PL Y1 to Y5'!$E$72 | 3 |
| scenarios.High.annual.Y3.advertising | annual_Y3_advertising | 'Annual PL Y1 to Y5'!$E$73 | 3 |
| scenarios.High.annual.Y3.returns | annual_Y3_returns | 'Annual PL Y1 to Y5'!$E$74 | 3 |
| scenarios.High.annual.Y3.contribution | annual_Y3_contribution | 'Annual PL Y1 to Y5'!$E$75 | 3 |
| scenarios.High.annual.Y3.fixed | annual_Y3_fixed | 'Annual PL Y1 to Y5'!$E$76 | 3 |
| scenarios.High.annual.Y3.seller_fee | annual_Y3_seller_fee | 'Annual PL Y1 to Y5'!$E$77 | 3 |
| scenarios.High.annual.Y3.launch | annual_Y3_launch | 'Annual PL Y1 to Y5'!$E$78 | 3 |
| scenarios.High.annual.Y3.compliance | annual_Y3_compliance | 'Annual PL Y1 to Y5'!$E$79 | 3 |
| scenarios.High.annual.Y3.cum_net | annual_Y3_cum_net | 'Annual PL Y1 to Y5'!$E$82 | 3 |
| scenarios.High.annual.Y3.amazon_units | annual_Y3_amazon_units | 'Annual PL Y1 to Y5'!$E$6 | 3 |
| scenarios.High.annual.Y3.asp | annual_Y3_asp | 'Annual PL Y1 to Y5'!$E$7 | 3 |
| scenarios.High.annual.Y3.landed_pct | annual_Y3_landed_pct | 'Annual PL Y1 to Y5'!$E$8 | 3 |
| scenarios.High.annual.Y3.ads_pct | annual_Y3_ads_pct | 'Annual PL Y1 to Y5'!$E$9 | 3 |
| scenarios.High.annual.Y3.amazon_contribution_pct | annual_Y3_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$E$83 | 3 |
| scenarios.High.annual.Y3.revenue_by_channel.amazon_us | annual_Y3_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$E$15 | 3 |
| scenarios.High.annual.Y3.revenue_by_channel.ebay_walmart | annual_Y3_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$E$24 | 3 |
| scenarios.High.annual.Y3.revenue_by_channel.b2b | annual_Y3_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$E$33 | 3 |
| scenarios.High.annual.Y3.revenue_by_channel.dtc | annual_Y3_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$E$42 | 3 |
| scenarios.High.annual.Y3.revenue_by_channel.amazon_canada | annual_Y3_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$E$51 | 3 |
| scenarios.High.annual.Y3.revenue_by_channel.pro_line | annual_Y3_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$E$60 | 3 |
| scenarios.High.annual.Y4.revenue | annual_Y4_revenue | 'Annual PL Y1 to Y5'!$F$69 | 3 |
| scenarios.High.annual.Y4.net | annual_Y4_net | 'Annual PL Y1 to Y5'!$F$80 | 3 |
| scenarios.High.annual.Y4.net_margin | annual_Y4_net_margin | 'Annual PL Y1 to Y5'!$F$81 | 3 |
| scenarios.High.annual.Y4.cogs | annual_Y4_cogs | 'Annual PL Y1 to Y5'!$F$70 | 3 |
| scenarios.High.annual.Y4.fees | annual_Y4_fees | 'Annual PL Y1 to Y5'!$F$71 | 3 |
| scenarios.High.annual.Y4.fulfillment | annual_Y4_fulfillment | 'Annual PL Y1 to Y5'!$F$72 | 3 |
| scenarios.High.annual.Y4.advertising | annual_Y4_advertising | 'Annual PL Y1 to Y5'!$F$73 | 3 |
| scenarios.High.annual.Y4.returns | annual_Y4_returns | 'Annual PL Y1 to Y5'!$F$74 | 3 |
| scenarios.High.annual.Y4.contribution | annual_Y4_contribution | 'Annual PL Y1 to Y5'!$F$75 | 3 |
| scenarios.High.annual.Y4.fixed | annual_Y4_fixed | 'Annual PL Y1 to Y5'!$F$76 | 3 |
| scenarios.High.annual.Y4.seller_fee | annual_Y4_seller_fee | 'Annual PL Y1 to Y5'!$F$77 | 3 |
| scenarios.High.annual.Y4.launch | annual_Y4_launch | 'Annual PL Y1 to Y5'!$F$78 | 3 |
| scenarios.High.annual.Y4.compliance | annual_Y4_compliance | 'Annual PL Y1 to Y5'!$F$79 | 3 |
| scenarios.High.annual.Y4.cum_net | annual_Y4_cum_net | 'Annual PL Y1 to Y5'!$F$82 | 3 |
| scenarios.High.annual.Y4.amazon_units | annual_Y4_amazon_units | 'Annual PL Y1 to Y5'!$F$6 | 3 |
| scenarios.High.annual.Y4.asp | annual_Y4_asp | 'Annual PL Y1 to Y5'!$F$7 | 3 |
| scenarios.High.annual.Y4.landed_pct | annual_Y4_landed_pct | 'Annual PL Y1 to Y5'!$F$8 | 3 |
| scenarios.High.annual.Y4.ads_pct | annual_Y4_ads_pct | 'Annual PL Y1 to Y5'!$F$9 | 3 |
| scenarios.High.annual.Y4.amazon_contribution_pct | annual_Y4_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$F$83 | 3 |
| scenarios.High.annual.Y4.revenue_by_channel.amazon_us | annual_Y4_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$F$15 | 3 |
| scenarios.High.annual.Y4.revenue_by_channel.ebay_walmart | annual_Y4_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$F$24 | 3 |
| scenarios.High.annual.Y4.revenue_by_channel.b2b | annual_Y4_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$F$33 | 3 |
| scenarios.High.annual.Y4.revenue_by_channel.dtc | annual_Y4_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$F$42 | 3 |
| scenarios.High.annual.Y4.revenue_by_channel.amazon_canada | annual_Y4_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$F$51 | 3 |
| scenarios.High.annual.Y4.revenue_by_channel.pro_line | annual_Y4_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$F$60 | 3 |
| scenarios.High.annual.Y5.revenue | annual_Y5_revenue | 'Annual PL Y1 to Y5'!$G$69 | 3 |
| scenarios.High.annual.Y5.net | annual_Y5_net | 'Annual PL Y1 to Y5'!$G$80 | 3 |
| scenarios.High.annual.Y5.net_margin | annual_Y5_net_margin | 'Annual PL Y1 to Y5'!$G$81 | 3 |
| scenarios.High.annual.Y5.cogs | annual_Y5_cogs | 'Annual PL Y1 to Y5'!$G$70 | 3 |
| scenarios.High.annual.Y5.fees | annual_Y5_fees | 'Annual PL Y1 to Y5'!$G$71 | 3 |
| scenarios.High.annual.Y5.fulfillment | annual_Y5_fulfillment | 'Annual PL Y1 to Y5'!$G$72 | 3 |
| scenarios.High.annual.Y5.advertising | annual_Y5_advertising | 'Annual PL Y1 to Y5'!$G$73 | 3 |
| scenarios.High.annual.Y5.returns | annual_Y5_returns | 'Annual PL Y1 to Y5'!$G$74 | 3 |
| scenarios.High.annual.Y5.contribution | annual_Y5_contribution | 'Annual PL Y1 to Y5'!$G$75 | 3 |
| scenarios.High.annual.Y5.fixed | annual_Y5_fixed | 'Annual PL Y1 to Y5'!$G$76 | 3 |
| scenarios.High.annual.Y5.seller_fee | annual_Y5_seller_fee | 'Annual PL Y1 to Y5'!$G$77 | 3 |
| scenarios.High.annual.Y5.launch | annual_Y5_launch | 'Annual PL Y1 to Y5'!$G$78 | 3 |
| scenarios.High.annual.Y5.compliance | annual_Y5_compliance | 'Annual PL Y1 to Y5'!$G$79 | 3 |
| scenarios.High.annual.Y5.cum_net | annual_Y5_cum_net | 'Annual PL Y1 to Y5'!$G$82 | 3 |
| scenarios.High.annual.Y5.amazon_units | annual_Y5_amazon_units | 'Annual PL Y1 to Y5'!$G$6 | 3 |
| scenarios.High.annual.Y5.asp | annual_Y5_asp | 'Annual PL Y1 to Y5'!$G$7 | 3 |
| scenarios.High.annual.Y5.landed_pct | annual_Y5_landed_pct | 'Annual PL Y1 to Y5'!$G$8 | 3 |
| scenarios.High.annual.Y5.ads_pct | annual_Y5_ads_pct | 'Annual PL Y1 to Y5'!$G$9 | 3 |
| scenarios.High.annual.Y5.amazon_contribution_pct | annual_Y5_amazon_contribution_pct | 'Annual PL Y1 to Y5'!$G$83 | 3 |
| scenarios.High.annual.Y5.revenue_by_channel.amazon_us | annual_Y5_revenue_by_channel_amazon_us | 'Annual PL Y1 to Y5'!$G$15 | 3 |
| scenarios.High.annual.Y5.revenue_by_channel.ebay_walmart | annual_Y5_revenue_by_channel_ebay_walmart | 'Annual PL Y1 to Y5'!$G$24 | 3 |
| scenarios.High.annual.Y5.revenue_by_channel.b2b | annual_Y5_revenue_by_channel_b2b | 'Annual PL Y1 to Y5'!$G$33 | 3 |
| scenarios.High.annual.Y5.revenue_by_channel.dtc | annual_Y5_revenue_by_channel_dtc | 'Annual PL Y1 to Y5'!$G$42 | 3 |
| scenarios.High.annual.Y5.revenue_by_channel.amazon_canada | annual_Y5_revenue_by_channel_amazon_canada | 'Annual PL Y1 to Y5'!$G$51 | 3 |
| scenarios.High.annual.Y5.revenue_by_channel.pro_line | annual_Y5_revenue_by_channel_pro_line | 'Annual PL Y1 to Y5'!$G$60 | 3 |
| scenarios.High.cash.peak_need | cash_peak_need | 'Cash Flow and Working Capital'!$C$48 | 3 |
| scenarios.High.cash.peak_need_month | cash_peak_need_month | 'Cash Flow and Working Capital'!$C$47 | 3 |
| scenarios.High.cash.peak_need_month_label | cash_peak_need_month_label | 'Cash Flow and Working Capital'!$C$46 | 3 |
| scenarios.High.cash.min_balance | cash_min_balance | 'Cash Flow and Working Capital'!$C$45 | 3 |
| scenarios.High.cash.funding_gap | cash_funding_gap | 'Cash Flow and Working Capital'!$C$49 | 3 |
| scenarios.High.cash.funding_needed_by_month | cash_funding_needed_by_month | 'Cash Flow and Working Capital'!$C$50 | 3 |
| scenarios.High.cash.runway_months_day90 | cash_runway_months_day90 | 'Cash Flow and Working Capital'!$C$55 | 3 |
| scenarios.High.cash.runway_months_day90_unfunded | cash_runway_months_day90_unfunded | 'Cash Flow and Working Capital'!$C$52 | 3 |
| scenarios.High.cash.opening_capital | cash_opening_capital | 'Assumptions'!$F$106 | 3 |
| scenarios.High.gates.day90.metrics.tacos.value | gate_d90_tacos_value | 'Funding and Breakeven'!$B$28 | 3 |
| scenarios.High.gates.day90.metrics.ret.value | gate_d90_ret_value | 'Funding and Breakeven'!$B$29 | 3 |
| scenarios.High.gates.day90.metrics.cpu.value | gate_d90_cpu_value | 'Funding and Breakeven'!$B$30 | 3 |
| scenarios.High.gates.day90.metrics.runway.value | gate_d90_runway_value | 'Funding and Breakeven'!$B$31 | 3 |
| scenarios.High.gates.day90.result | gate_day90_result | 'Funding and Breakeven'!$B$33 | 3 |
| scenarios.High.gates.month12.metrics.rev.value | gate_m12_rev_value | 'Funding and Breakeven'!$B$36 | 3 |
| scenarios.High.gates.month12.metrics.cum.value | gate_m12_cum_value | 'Funding and Breakeven'!$B$37 | 3 |
| scenarios.High.gates.month12.metrics.tacos.value | gate_m12_tacos_value | 'Funding and Breakeven'!$B$38 | 3 |
| scenarios.High.gates.month12.metrics.ret.value | gate_m12_ret_value | 'Funding and Breakeven'!$B$39 | 3 |
| scenarios.High.gates.month12.metrics.opcf.value | gate_m12_opcf_value | 'Funding and Breakeven'!$B$40 | 3 |
| scenarios.High.gates.month12.result | gate_month12_result | 'Funding and Breakeven'!$B$41 | 3 |
| unit_economics.hero.price | ue_hero_price | 'SKU Unit Economics'!$B$33 | any |
| unit_economics.hero.list | ue_hero_list | 'SKU Unit Economics'!$B$34 | any |
| unit_economics.hero.landed_cost | ue_hero_landed_cost | 'SKU Unit Economics'!$B$36 | any |
| unit_economics.hero.landed_pct | ue_hero_landed_pct | 'SKU Unit Economics'!$B$35 | any |
| unit_economics.hero.referral | ue_hero_referral | 'SKU Unit Economics'!$B$38 | any |
| unit_economics.hero.fba | ue_hero_fba | 'SKU Unit Economics'!$B$39 | any |
| unit_economics.hero.ads | ue_hero_ads | 'SKU Unit Economics'!$B$41 | any |
| unit_economics.hero.returns_storage | ue_hero_returns_storage | 'SKU Unit Economics'!$B$43 | any |
| unit_economics.hero.net | ue_hero_net | 'SKU Unit Economics'!$B$44 | any |
| unit_economics.hero.net_pct | ue_hero_net_pct | 'SKU Unit Economics'!$B$45 | any |
| unit_economics.hero.korea_wholesale_gross | ue_hero_korea_wholesale_gross | 'SKU Unit Economics'!$B$46 | any |
| unit_economics.hero.contribution_before_ads | ue_hero_contribution_before_ads | 'SKU Unit Economics'!$B$47 | any |
| unit_economics.hero.fob_usd | HeroFobUsd | 'SKU Unit Economics'!$B$49 | any |
| unit_economics.hero.fob_share_of_price | HeroFobShare | 'SKU Unit Economics'!$B$48 | any |
| unit_economics.per_sku | sku_table | 'SKU Unit Economics'!$A$5:$AF$14 | any |
| unit_economics.cost_bases.h2.landed | cost_base_h2_landed | 'SKU Unit Economics'!$D$53 | any |
| unit_economics.cost_bases.h2.net | cost_base_h2_net | 'SKU Unit Economics'!$F$53 | any |
| unit_economics.cost_bases.august.landed | cost_base_august_landed | 'SKU Unit Economics'!$D$54 | any |
| unit_economics.cost_bases.august.net | cost_base_august_net | 'SKU Unit Economics'!$F$54 | any |
| unit_economics.cost_bases.alibaba_low.landed | cost_base_alibaba_low_landed | 'SKU Unit Economics'!$D$55 | any |
| unit_economics.cost_bases.alibaba_low.net | cost_base_alibaba_low_net | 'SKU Unit Economics'!$F$55 | any |
| unit_economics.cost_bases.alibaba_high.landed | cost_base_alibaba_high_landed | 'SKU Unit Economics'!$D$56 | any |
| unit_economics.cost_bases.alibaba_high.net | cost_base_alibaba_high_net | 'SKU Unit Economics'!$F$56 | any |
| duty.copper.rate | duty_copper_rate | 'Landed Cost & Tariff Scenarios'!$G$5 | any |
| duty.copper.landed_pct_medium | duty_copper_landed_pct_medium | 'Landed Cost & Tariff Scenarios'!$B$23 | any |
| duty.copper.hero_net | duty_copper_hero_net | 'Landed Cost & Tariff Scenarios'!$D$28 | any |
| duty.fiber.rate | duty_fiber_rate | 'Landed Cost & Tariff Scenarios'!$G$6 | any |
| duty.fiber.landed_pct_medium | duty_fiber_landed_pct_medium | 'Landed Cost & Tariff Scenarios'!$C$23 | any |
| duty.fiber.hero_net | duty_fiber_hero_net | 'Landed Cost & Tariff Scenarios'!$D$29 | any |
| duty.stress.rate | duty_stress_rate | 'Landed Cost & Tariff Scenarios'!$G$7 | any |
| duty.stress.landed_pct_medium | duty_stress_landed_pct_medium | 'Landed Cost & Tariff Scenarios'!$D$23 | any |
| duty.stress.hero_net | duty_stress_hero_net | 'Landed Cost & Tariff Scenarios'!$D$30 | any |
| duty.fiber_list3_check.rate | duty_fiber_list3_check_rate | 'Landed Cost & Tariff Scenarios'!$G$8 | any |
| duty.waterfall | waterfall_block | 'Landed Cost & Tariff Scenarios'!$E$31:$G$35 | any |
| duty.tariff_stack | tariff_stack_block | 'Landed Cost & Tariff Scenarios'!$A$39:$F$41 | any |
| sensitivity.tornado | tornado_block | 'Sensitivity'!$A$48:$I$53 | any |
| sensitivity.tornado_m18 | tornado18_block | 'Sensitivity'!$A$57:$I$62 | any |
| sensitivity.table | sensitivity_table | 'Sensitivity'!$A$25:$H$43 | any |
| sensitivity.base_y5_net | sens_base_y5_net | 'Sensitivity'!$B$12 | any |
| sensitivity.base_m18_net | sens_base_m18_net | 'Sensitivity'!$B$6 | any |
| levers | levers_block | 'Levers'!$A$5:$J$16 | any |
| gate_thresholds.d90_conv_pass.value | gate_threshold_d90_conv_pass | 'Assumptions'!$F$114 | any |
| gate_thresholds.d90_conv_kill.value | gate_threshold_d90_conv_kill | 'Assumptions'!$F$115 | any |
| gate_thresholds.d90_tacos_pass.value | gate_threshold_d90_tacos_pass | 'Assumptions'!$F$116 | any |
| gate_thresholds.d90_tacos_kill.value | gate_threshold_d90_tacos_kill | 'Assumptions'!$F$117 | any |
| gate_thresholds.d90_ret_pass.value | gate_threshold_d90_ret_pass | 'Assumptions'!$F$118 | any |
| gate_thresholds.d90_ret_kill.value | gate_threshold_d90_ret_kill | 'Assumptions'!$F$119 | any |
| gate_thresholds.d90_cpu_pass.value | gate_threshold_d90_cpu_pass | 'Assumptions'!$F$120 | any |
| gate_thresholds.d90_cpu_kill.value | gate_threshold_d90_cpu_kill | 'Assumptions'!$F$121 | any |
| gate_thresholds.d90_run_pass.value | gate_threshold_d90_run_pass | 'Assumptions'!$F$122 | any |
| gate_thresholds.d90_run_kill.value | gate_threshold_d90_run_kill | 'Assumptions'!$F$123 | any |
| gate_thresholds.d90_ip_days.value | gate_threshold_d90_ip_days | 'Assumptions'!$F$124 | any |
| gate_thresholds.m12_rev_pass.value | gate_threshold_m12_rev_pass | 'Assumptions'!$F$125 | any |
| gate_thresholds.m12_rev_kill.value | gate_threshold_m12_rev_kill | 'Assumptions'!$F$126 | any |
| gate_thresholds.m12_cum_pass.value | gate_threshold_m12_cum_pass | 'Assumptions'!$F$127 | any |
| gate_thresholds.m12_cum_kill.value | gate_threshold_m12_cum_kill | 'Assumptions'!$F$128 | any |
| gate_thresholds.m12_tacos_pass.value | gate_threshold_m12_tacos_pass | 'Assumptions'!$F$129 | any |
| gate_thresholds.m12_tacos_kill.value | gate_threshold_m12_tacos_kill | 'Assumptions'!$F$130 | any |
| gate_thresholds.m12_ret_pass.value | gate_threshold_m12_ret_pass | 'Assumptions'!$F$131 | any |
| gate_thresholds.m12_ret_kill.value | gate_threshold_m12_ret_kill | 'Assumptions'!$F$132 | any |
| gate_thresholds.m12_gap_kill.value | gate_threshold_m12_gap_kill | 'Assumptions'!$F$133 | any |
| august_deal.units | august_deal_units | 'FX'!$B$9 | any |
| august_deal.cost_krw | august_deal_cost_krw | 'FX'!$B$10 | any |
| august_deal.revenue_krw | august_deal_revenue_krw | 'FX'!$B$11 | any |
| august_deal.gross_profit_krw | august_deal_gross_profit_krw | 'FX'!$B$12 | any |
| august_deal.gross_margin | august_deal_gross_margin | 'FX'!$B$13 | any |
| plan_b.Y1.revenue | planb_Y1_revenue | 'Plan B'!$C$28 | any |
| plan_b.Y1.net | planb_Y1_net | 'Plan B'!$C$29 | any |
| plan_b.Y1.b2b_revenue | planb_Y1_b2b_revenue | 'Plan B'!$C$5 | any |
| plan_b.Y1.korea_gross_profit | planb_Y1_korea_gross_profit | 'Plan B'!$C$23 | any |
| plan_b.Y2.revenue | planb_Y2_revenue | 'Plan B'!$D$28 | any |
| plan_b.Y2.net | planb_Y2_net | 'Plan B'!$D$29 | any |
| plan_b.Y2.b2b_revenue | planb_Y2_b2b_revenue | 'Plan B'!$D$5 | any |
| plan_b.Y2.korea_gross_profit | planb_Y2_korea_gross_profit | 'Plan B'!$D$23 | any |
| plan_b.Y3.revenue | planb_Y3_revenue | 'Plan B'!$E$28 | any |
| plan_b.Y3.net | planb_Y3_net | 'Plan B'!$E$29 | any |
| plan_b.Y3.b2b_revenue | planb_Y3_b2b_revenue | 'Plan B'!$E$5 | any |
| plan_b.Y3.korea_gross_profit | planb_Y3_korea_gross_profit | 'Plan B'!$E$23 | any |
| scenario_summary.Low.m18_revenue | ss_Low_m18_revenue | 'Scenario Summary'!$B$6 | any (Scenario Summary engines) |
| scenario_summary.Low.m18_net | ss_Low_m18_net | 'Scenario Summary'!$B$7 | any (Scenario Summary engines) |
| scenario_summary.Low.m18_net_margin | ss_Low_m18_net_margin | 'Scenario Summary'!$B$8 | any (Scenario Summary engines) |
| scenario_summary.Low.m18_breakeven_month | ss_Low_m18_breakeven_month | 'Scenario Summary'!$B$9 | any (Scenario Summary engines) |
| scenario_summary.Low.m18_run_rate | ss_Low_m18_run_rate | 'Scenario Summary'!$B$10 | any (Scenario Summary engines) |
| scenario_summary.Low.m18_exit_margin | ss_Low_m18_exit_margin | 'Scenario Summary'!$B$11 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y1_revenue | ss_Low_annual_Y1_revenue | 'Scenario Summary'!$B$13 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y1_net | ss_Low_annual_Y1_net | 'Scenario Summary'!$B$14 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y2_revenue | ss_Low_annual_Y2_revenue | 'Scenario Summary'!$B$15 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y2_net | ss_Low_annual_Y2_net | 'Scenario Summary'!$B$16 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y3_revenue | ss_Low_annual_Y3_revenue | 'Scenario Summary'!$B$17 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y3_net | ss_Low_annual_Y3_net | 'Scenario Summary'!$B$18 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y4_revenue | ss_Low_annual_Y4_revenue | 'Scenario Summary'!$B$19 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y4_net | ss_Low_annual_Y4_net | 'Scenario Summary'!$B$20 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y5_revenue | ss_Low_annual_Y5_revenue | 'Scenario Summary'!$B$21 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y5_net | ss_Low_annual_Y5_net | 'Scenario Summary'!$B$22 | any (Scenario Summary engines) |
| scenario_summary.Low.annual_Y5_net_margin | ss_Low_annual_Y5_net_margin | 'Scenario Summary'!$B$23 | any (Scenario Summary engines) |
| scenario_summary.Low.cash_peak_need | ss_Low_cash_peak_need | 'Scenario Summary'!$B$25 | any (Scenario Summary engines) |
| scenario_summary.Low.cash_peak_need_month | ss_Low_cash_peak_need_month | 'Scenario Summary'!$B$26 | any (Scenario Summary engines) |
| scenario_summary.Low.cash_peak_need_month_label | ss_Low_cash_peak_need_month_label | 'Scenario Summary'!$B$27 | any (Scenario Summary engines) |
| scenario_summary.Low.cash_funding_gap | ss_Low_cash_funding_gap | 'Scenario Summary'!$B$29 | any (Scenario Summary engines) |
| scenario_summary.Low.cash_funding_needed_by_month | ss_Low_cash_funding_needed_by_month | 'Scenario Summary'!$B$30 | any (Scenario Summary engines) |
| scenario_summary.Low.cash_runway_months_day90 | ss_Low_cash_runway_months_day90 | 'Scenario Summary'!$B$31 | any (Scenario Summary engines) |
| scenario_summary.Low.gate_day90 | ss_Low_gate_day90 | 'Scenario Summary'!$B$32 | any (Scenario Summary engines) |
| scenario_summary.Low.gate_month12 | ss_Low_gate_month12 | 'Scenario Summary'!$B$33 | any (Scenario Summary engines) |
| scenario_summary.Medium.m18_revenue | ss_Medium_m18_revenue | 'Scenario Summary'!$C$6 | any (Scenario Summary engines) |
| scenario_summary.Medium.m18_net | ss_Medium_m18_net | 'Scenario Summary'!$C$7 | any (Scenario Summary engines) |
| scenario_summary.Medium.m18_net_margin | ss_Medium_m18_net_margin | 'Scenario Summary'!$C$8 | any (Scenario Summary engines) |
| scenario_summary.Medium.m18_breakeven_month | ss_Medium_m18_breakeven_month | 'Scenario Summary'!$C$9 | any (Scenario Summary engines) |
| scenario_summary.Medium.m18_run_rate | ss_Medium_m18_run_rate | 'Scenario Summary'!$C$10 | any (Scenario Summary engines) |
| scenario_summary.Medium.m18_exit_margin | ss_Medium_m18_exit_margin | 'Scenario Summary'!$C$11 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y1_revenue | ss_Medium_annual_Y1_revenue | 'Scenario Summary'!$C$13 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y1_net | ss_Medium_annual_Y1_net | 'Scenario Summary'!$C$14 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y2_revenue | ss_Medium_annual_Y2_revenue | 'Scenario Summary'!$C$15 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y2_net | ss_Medium_annual_Y2_net | 'Scenario Summary'!$C$16 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y3_revenue | ss_Medium_annual_Y3_revenue | 'Scenario Summary'!$C$17 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y3_net | ss_Medium_annual_Y3_net | 'Scenario Summary'!$C$18 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y4_revenue | ss_Medium_annual_Y4_revenue | 'Scenario Summary'!$C$19 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y4_net | ss_Medium_annual_Y4_net | 'Scenario Summary'!$C$20 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y5_revenue | ss_Medium_annual_Y5_revenue | 'Scenario Summary'!$C$21 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y5_net | ss_Medium_annual_Y5_net | 'Scenario Summary'!$C$22 | any (Scenario Summary engines) |
| scenario_summary.Medium.annual_Y5_net_margin | ss_Medium_annual_Y5_net_margin | 'Scenario Summary'!$C$23 | any (Scenario Summary engines) |
| scenario_summary.Medium.cash_peak_need | ss_Medium_cash_peak_need | 'Scenario Summary'!$C$25 | any (Scenario Summary engines) |
| scenario_summary.Medium.cash_peak_need_month | ss_Medium_cash_peak_need_month | 'Scenario Summary'!$C$26 | any (Scenario Summary engines) |
| scenario_summary.Medium.cash_peak_need_month_label | ss_Medium_cash_peak_need_month_label | 'Scenario Summary'!$C$27 | any (Scenario Summary engines) |
| scenario_summary.Medium.cash_funding_gap | ss_Medium_cash_funding_gap | 'Scenario Summary'!$C$29 | any (Scenario Summary engines) |
| scenario_summary.Medium.cash_funding_needed_by_month | ss_Medium_cash_funding_needed_by_month | 'Scenario Summary'!$C$30 | any (Scenario Summary engines) |
| scenario_summary.Medium.cash_runway_months_day90 | ss_Medium_cash_runway_months_day90 | 'Scenario Summary'!$C$31 | any (Scenario Summary engines) |
| scenario_summary.Medium.gate_day90 | ss_Medium_gate_day90 | 'Scenario Summary'!$C$32 | any (Scenario Summary engines) |
| scenario_summary.Medium.gate_month12 | ss_Medium_gate_month12 | 'Scenario Summary'!$C$33 | any (Scenario Summary engines) |
| scenario_summary.High.m18_revenue | ss_High_m18_revenue | 'Scenario Summary'!$D$6 | any (Scenario Summary engines) |
| scenario_summary.High.m18_net | ss_High_m18_net | 'Scenario Summary'!$D$7 | any (Scenario Summary engines) |
| scenario_summary.High.m18_net_margin | ss_High_m18_net_margin | 'Scenario Summary'!$D$8 | any (Scenario Summary engines) |
| scenario_summary.High.m18_breakeven_month | ss_High_m18_breakeven_month | 'Scenario Summary'!$D$9 | any (Scenario Summary engines) |
| scenario_summary.High.m18_run_rate | ss_High_m18_run_rate | 'Scenario Summary'!$D$10 | any (Scenario Summary engines) |
| scenario_summary.High.m18_exit_margin | ss_High_m18_exit_margin | 'Scenario Summary'!$D$11 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y1_revenue | ss_High_annual_Y1_revenue | 'Scenario Summary'!$D$13 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y1_net | ss_High_annual_Y1_net | 'Scenario Summary'!$D$14 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y2_revenue | ss_High_annual_Y2_revenue | 'Scenario Summary'!$D$15 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y2_net | ss_High_annual_Y2_net | 'Scenario Summary'!$D$16 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y3_revenue | ss_High_annual_Y3_revenue | 'Scenario Summary'!$D$17 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y3_net | ss_High_annual_Y3_net | 'Scenario Summary'!$D$18 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y4_revenue | ss_High_annual_Y4_revenue | 'Scenario Summary'!$D$19 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y4_net | ss_High_annual_Y4_net | 'Scenario Summary'!$D$20 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y5_revenue | ss_High_annual_Y5_revenue | 'Scenario Summary'!$D$21 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y5_net | ss_High_annual_Y5_net | 'Scenario Summary'!$D$22 | any (Scenario Summary engines) |
| scenario_summary.High.annual_Y5_net_margin | ss_High_annual_Y5_net_margin | 'Scenario Summary'!$D$23 | any (Scenario Summary engines) |
| scenario_summary.High.cash_peak_need | ss_High_cash_peak_need | 'Scenario Summary'!$D$25 | any (Scenario Summary engines) |
| scenario_summary.High.cash_peak_need_month | ss_High_cash_peak_need_month | 'Scenario Summary'!$D$26 | any (Scenario Summary engines) |
| scenario_summary.High.cash_peak_need_month_label | ss_High_cash_peak_need_month_label | 'Scenario Summary'!$D$27 | any (Scenario Summary engines) |
| scenario_summary.High.cash_funding_gap | ss_High_cash_funding_gap | 'Scenario Summary'!$D$29 | any (Scenario Summary engines) |
| scenario_summary.High.cash_funding_needed_by_month | ss_High_cash_funding_needed_by_month | 'Scenario Summary'!$D$30 | any (Scenario Summary engines) |
| scenario_summary.High.cash_runway_months_day90 | ss_High_cash_runway_months_day90 | 'Scenario Summary'!$D$31 | any (Scenario Summary engines) |
| scenario_summary.High.gate_day90 | ss_High_gate_day90 | 'Scenario Summary'!$D$32 | any (Scenario Summary engines) |
| scenario_summary.High.gate_month12 | ss_High_gate_month12 | 'Scenario Summary'!$D$33 | any (Scenario Summary engines) |

## Keys without a single cell

meta: build metadata (fx, scenario_default, duty_default, built, workbook, fx_scenarios, selector state, scenario summary cross check).
scenarios.{S}.gates.*.pass_threshold, kill_trigger, label: text from BUILD_SPEC 2.7 written by pull_figures.py.
kill_criteria: the Day 90 and Month 12 criteria lists (text from BUILD_SPEC 2.7) plus results per scenario (from gate_day90_result and gate_month12_result).
headline: a convenience copy of the scenario headline numbers (same cells as scenarios.{S}.m18, annual, cash, gates).
unit_economics.per_sku: one object per row of the SKU Unit Economics table (sku_table), columns in table order.
duty.waterfall and duty.tariff_stack: chart data blocks from the Landed Cost & Tariff Scenarios tab.
sensitivity.table: the one way sensitivity table rows (driver, setting, setting value, landed % change, 18 month net, Year 5 net, deltas).
scenario_summary: the Scenario Summary tab (three parallel live engines) read from the delivered workbook; used to cross check the selector pulls.
full_recalc: Medium case recalculated at each duty scenario (copper, fiber, stress) and each FX scenario (1300, 1430, 1550): landed %, 18 month revenue and net, breakeven, Year 5 revenue and net, peak need, funding gap, and the deltas vs the Medium base. These are exact full model values; the Sensitivity tab rows are compact approximations.
scenarios.{S}.insurance_trigger_month and insurance_trigger_margin_usd, insurance_trigger_month: the first month whose Amazon US revenue (monthly_amazon_revenue) is at or above facts.amazon.insurance_trigger_monthly_sales_usd, and the amount it passes the threshold by; derived by pull_figures.py, the single source for the deck and the SOPs.
gate_thresholds: the Day 90 and Month 12 thresholds from the Gate thresholds block on Assumptions (one cell each; the gate formulas and their wording read these cells).
august_deal: the August 2026 deal totals (units, cost, revenue, gross profit in KRW, gross margin) from the yellow input cells on the FX tab, sourced from inputs/august_deal_totals.csv.
model_notes: every deviation from BUILD_SPEC and the choice behind it (id, where, default, why, decides, by), written by pull_figures.py from model_core.py for 11_Decision_Log.md, the README tie out and the band slide.
