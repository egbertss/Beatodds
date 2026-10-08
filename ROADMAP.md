# Roadmap

## Part 1 (done, October 2026)

- Data pipeline for Polymarket and Kalshi earnings contracts, Fed contracts and Yahoo Finance prices.
- Excel workbook with every statistic as a visible formula.
- Main finding: when the crowd is at least 80% sure of a beat, a miss costs about −7.0% against −3.0% otherwise.

## Part 2 (in progress)

1. **Beat the crowd.** Train a machine-learning model on past reports only (crowd probability and its weekly change, company record, volatility, option prices). Test it walk-forward on later months. Keep it only if its Brier score beats the crowd's.
2. **Size of the move.** Predict the absolute abnormal return so each report gets a likely range, not only a probability.
3. **Is the crowd too pessimistic?** In Part 1, buying the "beat" side when the crowd gave 60% or less returned about 10% per contract after costs (279 reports). Test it on new reports only.
4. **Early warning.** A watch list of reports where the crowd is very sure of a beat, because there a miss has cost the most.
5. **Oil and crypto.** Kalshi daily WTI contracts and Polymarket monthly and weekly "What will WTI hit?" events, plus Bitcoin and Ethereum on both. Compare the crowd's probability with the actual price (Yahoo: CL=F, BTC-USD, ETH-USD) and with volatility-based benchmarks.
6. **Complete Kalshi download.** Retries and slower requests so every run picks up the same contracts.

## Final submission

8 to 10 slides and a video of 2 minutes at most.
