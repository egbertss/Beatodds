# BeatOdds

**What prediction markets know before earnings day.**

BeatOdds reads the prices of prediction-market contracts such as "Will [company] beat quarterly earnings?" on Polymarket and Kalshi. A price of 0.85 means the crowd sees an 85% chance of a beat. The project tests whether that probability forecasts earnings results better than simple benchmarks, and whether it tells an investor how the share will react when the results come out.

Group A2, Master in Finance, International University of Monaco. Course project for Statistics and Financial Data Analysis (AI for Financial Data Analysis Challenge), October 2026.

## Main findings (Part 1)

Sample: 1,207 settled Polymarket earnings contracts on 423 companies, September 2025 to October 2026.

| Average move of the share against the market | Crowd 80% or more sure of a beat | Crowd less sure |
|---|---|---|
| The company beat | +0.2% (579 reports) | +1.5% (330 reports) |
| The company missed | **−7.0%** (63 reports) | −3.0% (235 reports) |

- When the crowd is very sure of a beat, a beat adds little and a miss costs more than twice as much. The gap of about 4 points has a 95% interval of 1.3 to 6.8 points.
- The gap holds after adjusting for market sensitivity (beta), with a three-day window, for calm and volatile shares, for the most traded contracts and at a stricter 85% cut-off.
- It is smaller and not significant when the odds are read a day or a week earlier, so the signal is sharpest on the report day.
- The crowd is well calibrated (it says 85%, companies beat 83% of the time) and about 25% more accurate than the usual beat rate, 17% more than the company's own history and 22% more than a model built on each company's record against analysts.
- On US Fed decisions, both markets put about 95% on the actual outcome the day before.

All results are associations in one sample. They describe expectations and price moves, not cause and effect.

## Repository contents

| Path | What it is |
|---|---|
| `beatodds_p1.py` | The full Part 1 code: downloads the data and builds the Excel workbook |
| `BeatOdds_Colab.ipynb` | A two-cell notebook to run the code in Google Colab |
| `requirements.txt` | Python packages |
| `docs/ROADMAP.md` | Plan for Part 2 |

The workbook and raw data are not stored here. Run the code to build them from the live sources. The findings above come from the run of 8 October 2026, 12:09 UTC.

## Data sources

All free and public. No API keys are needed.

| Source | What | Frequency |
|---|---|---|
| Polymarket (Gamma API, CLOB prices-history) | Earnings and Fed decision contracts with their prices | Hourly |
| Kalshi (Trade API v2) | Company-figure (KPI) and Fed decision contracts | Hourly |
| Yahoo Finance (yfinance) | Share prices, SPY, 5-year Treasury yield (^FVX), analyst estimates and reported EPS, option prices | Daily, quarterly, snapshot |
| FRED (public CSV) | 2-year Treasury yield and Fed target, used when reachable | Daily |

## How to run it

### In Google Colab (recommended)

1. Open `BeatOdds_Colab.ipynb` in Colab, or create a new notebook.
2. First cell: `!pip install -q yfinance openpyxl scikit-learn`
3. Second cell: paste the whole of `beatodds_p1.py` and run it.
4. The run takes about 15 minutes, including a step that installs LibreOffice and calculates every formula.
5. Colab downloads two files: `Group<number>_P1_Workbook.xlsx` and `Group<number>_P1_RawData.zip`.

### On your own computer

```bash
pip install -r requirements.txt
python beatodds_p1.py
```

The formula calculation step needs LibreOffice (`soffice`). Without it the workbook is still written, and Excel calculates the formulas when you open it and click Enable Editing.

### Settings

At the top of `beatodds_p1.py`:

- `GROUP_NUMBER` sets the file names.
- `HIGH_P = 0.8` is the "confident crowd" cut-off.
- `USE_KALSHI`, `USE_FED`, `USE_OPTIONS` switch the extra sources on or off.
- `USE_EDGAR` and `USE_FINRA` are off. To use SEC EDGAR, put your own name and email in `SEC_USER_AGENT`.

Each run downloads the latest data, so numbers change slightly from run to run. The Kalshi download is rate-limited, so the number of Kalshi contracts can differ between runs.

## The workbook

Python only downloads and arranges the data and fits one benchmark model. Every statistic in the workbook is a visible Excel formula.

- **Results**: parts A to R (accuracy, calibration, stock reaction, the priced-in table, run-up and drift, liquidity tiers, trading rules, Fed weeks, robustness, cut-off sensitivity).
- **Company**: pick a ticker to see its history.
- **Live**: a scenario card for each upcoming report and the odds for the next Fed meeting.
- **Events**: one row per report.
- **Kalshi_Results, Fed**: the Kalshi and Fed analyses.
- **Assumptions**: sources, download dates and settings.
- **Raw_ sheets**: the data as downloaded.

The formulas work in Excel set to a comma decimal separator (Italian, French, Dutch, German).

## Roadmap

See [docs/ROADMAP.md](docs/ROADMAP.md). In short: a machine-learning model tested against the crowd, a forecast of the size of the share move, an out-of-sample test of whether the crowd is too pessimistic below 60%, a watch list for portfolio managers, and an extension to oil, Bitcoin and Ethereum contracts.

## Use of generative AI

Claude (Anthropic) drafted and debugged the Python code and helped discuss the method and edit the text. The group chose the problem, the data sources, the comparison rules, the 80% threshold and the checks, verified the results against the raw data, and is responsible for the analysis.

## Disclaimer

Coursework for educational purposes. Not investment advice. Prediction-market and Yahoo Finance data belong to their providers. This repository holds only the code that downloads them under each provider's public terms.
