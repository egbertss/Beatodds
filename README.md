# BeatOdds

The crowd tells you whether a company will beat its earnings. BeatOdds tells you how far the share can fall, and what to do before the report.

BeatOdds reads the prices of prediction-market contracts such as "Will [company] beat quarterly earnings?" on Polymarket and Kalshi. A price of 0.85 means the crowd sees an 85% chance of a beat. For every holding that reports in the next 21 days, BeatOdds combines those odds with option prices, analyst estimates and the company's own record. It returns the chance of a beat, a 1-in-10 worst case for the share, the price of a protective put against its expected payoff, one action (trim, hedge or hold) and a short brief for the client. Every number in the brief is checked against the data.

Built by Jacobus Egberts Sonatori ([@egbertss](https://github.com/egbertss)): idea, research design, data pipeline, models, code and analysis. Started as a course project for Statistics and Financial Data Analysis (AI for Financial Data Analysis Challenge), Master in Finance, International University of Monaco, 2026.

## Main findings

Run of 25 November 2026. 1,426 settled Polymarket earnings reports since September 2025.

| Move of the share against the market | Crowd 80% or more sure of a beat | Crowd less sure |
|---|---|---|
| The company beat | +0.2% (615 reports) | +2.1% (406 reports) |
| The company missed | −6.6% (67 reports) | −2.1% (299 reports) |

- When the crowd is confident, a miss costs 4.5 points more than other misses (t −4.7, errors clustered by report week). With odds read before the release time on SEC EDGAR, the gap is −4.2 points (t −4.1). On the 203 reports after the Part 1 data freeze the gap is −0.3 points (t −0.2), with only 4 confident misses so far.
- The crowd forecasts beats better than the base rate and the company's own record (Brier score 0.142, against 0.193 and 0.179). No machine-learning model beat the crowd on this, so the crowd stays the beat forecast.
- A gradient-boosted quantile model, trained only on earlier months, gives each report a 1-in-10 worst case. It is 16% more accurate than the option-implied move.
- Backtest on 1,024 reports, out of sample, same share of reports trimmed for every policy. The average of the worst 10% of moves is −18.8% holding everything, −14.5% trimming when the crowd is confident and −11.8% trimming on the model's worst case (2.8 points better than the crowd rule, p 0.001). A simple screen that trims the biggest past movers does about as well (−11.6%). The model is better where there is no history to screen on (−20.7% against −41.4%, 10 reports).
- Live calls. 84 calls were saved and hashed (SHA-256) before the companies reported, and 83 have been scored. Holding, the worst 10% lost 23.1%. Following the calls, 14.0%. 10.8% of shares fell below their 1-in-10 worst case, against a target of 10%. This covers two days in one earnings week.
- Reading contracts. On 55 contracts rewritten in four new formats, the rule-based parser read none fully and Gemini read 95%. On 425 contracts it had not seen before, Gemini agreed with the parser on 99% to 100% of fields.
- When the crowd and StarMine both expect a beat, the worst 10% is −20.6%. When StarMine expects a miss, it is −14.5% (p 0.003). This was found after looking at the data and is being tested on new reports.

All results are associations in one sample. They describe expectations and price moves, not cause and effect.

## Repository contents

| File | What it is |
|---|---|
| `beatodds.py` | The full code: downloads the data, fits the models, scores live calls and builds the Excel workbook |
| `BeatOdds_Colab.ipynb` | A notebook that runs `beatodds.py` in Google Colab |
| `beatodds_p1.py` | The Part 1 version (October 2026), kept for reference |
| `requirements.txt` | Python packages |
| `ROADMAP.md` | What was done and what comes next |

The workbook and raw data are not stored here. Run the code to build them from the live sources.

## How to run it

In Google Colab (recommended):

1. Open `BeatOdds_Colab.ipynb` in Colab.
2. In Colab's Secrets panel (the key icon), add `GEMINI_API_KEY` (free at aistudio.google.com) and `SEC_CONTACT` (your name and e-mail, which the SEC asks for). Both are optional. Without them the AI steps or EDGAR are skipped.
3. Run cells 1 to 3. Step 1 asks for saved data zips (press Cancel to start from zero), downloads everything (about 20 to 30 minutes) and gives `GroupA2_LSEG_request.xlsx`.
4. Open that file in Excel with LSEG Workspace signed in, wait until the values fill in, and save it. Without an LSEG licence, press Cancel in the next step and the workbook is built without LSEG data.
5. Run cell 4 and upload the filled file (about 10 to 15 minutes). Colab downloads `GroupA2_P3_Workbook.xlsx` and `GroupA2_P3_RawData.zip`. Keep the zip: the next run restores it, so saved snapshots and live calls are kept and scored.

On your own computer: `pip install -r requirements.txt`, then `python beatodds.py` with `STEP = 1`, and again with `STEP = 2`.

A daily run that only saves the odds, option prices and live calls takes about 2 minutes (`SNAPSHOT_ONLY = True`). It can be scheduled through each earnings season.

## Data sources

| Source | What | Frequency |
|---|---|---|
| Polymarket (Gamma API, CLOB prices-history) | Earnings contracts, their text and prices | Hourly |
| Kalshi (Trade API v2) | Company-figure (KPI) contracts | Hourly |
| Yahoo Finance (yfinance) | Share prices, SPY, EPS history, option prices | Daily, quarterly, snapshot |
| SEC EDGAR | Release times (8-K item 2.02) and last quarter's release text | Per filing |
| LSEG Workspace (Excel add-in, licence needed) | Implied volatility, I/B/E/S consensus and revisions, StarMine SmartEstimate | Per report |
| Gemini 2.5 Flash (Google AI Studio) | Reads contract text and releases, writes client briefs | Per run |

## The workbook

Python downloads and arranges the data and fits the models. Every statistic in the workbook is a visible Excel formula. Python outputs (model forecasts, bootstrap draws) sit on `Raw_` sheets.

- Summary: what holds, in one page.
- Results, Tail, Tests: the crowd's accuracy and the cost of a confident miss.
- Models, Backtest, Bootstrap: walk-forward models, policies compared on their worst 10%, week-clustered bootstrap.
- Crowd_Adds: odds drift, crowd against StarMine and analysts.
- Hedge: the protective put priced against its expected payoff.
- Live, Agent, Scorecard: upcoming reports, client briefs, and live calls scored after the release.
- GenAI, AI_Checks: tests of the contract reading and the number guardrail.
- Business: market size with sources and value per client.
- Assumptions: sources, download dates and every setting (blue cells).

## Settings

At the top of `beatodds.py`: `GEMINI_API_KEY`, `SEC_CONTACT`, `GROUP_NUMBER` (file names), `HIGH_P = 0.80` (the confident-crowd cut-off, fixed in Part 1), and `STEP` (1 or 2). Trim size, trading cost and other inputs are blue cells on the Assumptions sheet.

The language model is called through one function (`genai`), so a bank can point it at its own approved model.

## Use of generative AI

Claude (Anthropic) helped draft and debug the Python code and edit the text. Gemini 2.5 Flash is part of the product: it reads contracts and releases and writes client briefs. I chose the problem, the data, the tests, the thresholds and the decisions, checked the results against the raw data, and I am responsible for the analysis.

## Disclaimer

Coursework for educational purposes. Not investment advice. Prediction-market, Yahoo Finance, SEC and LSEG data belong to their providers. This repository holds only the code that downloads them under each provider's terms.
