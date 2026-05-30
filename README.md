# 🧠 Aegis Omniscient — Stock Valuation Terminal

A single-file, no-API-key stock valuation CLI built on free [`yfinance`](https://pypi.org/project/yfinance/)
data. It values a company with a **2-stage Free-Cash-Flow-to-Equity (FCFE) DCF**
— the same methodology Simply Wall St uses — and surrounds it with analyst
ratings, peer percentiles, technicals, quality metrics, risk flags, and an
overall conviction call.

## Install

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Requires Python 3.10+. Export formats (`openpyxl`, `fpdf2`) are optional — the
app degrades gracefully if they're absent.

## Commands

```bash
python app.py single AAPL                 # full valuation dashboard
python app.py single AMZN --base ni       # override the FCFE base cash flow
python app.py single TSLA -m              # add a Monte Carlo simulation
python app.py single AAPL -e excel        # export: json | csv | excel | pdf
python app.py single AAPL --no-cache      # ignore the 6h cache and refetch

python app.py compare AAPL MSFT GOOGL     # side-by-side comparison table
python app.py watchlist                   # analyze your whole watchlist
python app.py add NVDA                    # add a ticker to the watchlist
python app.py remove NVDA                 # remove a ticker

python app.py scenario AMZN -g 0.22 -d 0.085 -t 0.036   # custom DCF inputs

python app.py config --show               # print current configuration
python app.py config -s default_base=auto # set a config key
```

Run `python app.py --help` (or `python app.py <command> --help`) for all options.

## Valuation methodology

The intrinsic value is a **2-stage FCFE DCF**:

1. **Base cash flow** — a per-share starting cash flow (see *base modes* below).
2. **Stage 1** — projected ~10 years, with growth seeded from analyst estimates
   and **fading** linearly toward the terminal rate. Stage-1 growth is capped
   (`max_stage1_growth`, default 20%) to stay defensible over a decade.
3. **Terminal value** — Gordon Growth at the long-run rate (5-yr average 10-yr
   Treasury yield, clamped to a sane band).
4. **Discount rate** — **cost of equity** via CAPM:
   `rf + adjusted_beta × equity_risk_premium + country_risk_premium`, using a
   Blume-adjusted beta. The **country-risk premium** (Damodaran-style) adds a
   sovereign/FX premium for non-US companies (e.g. Brazil ≈ +3.5%), so emerging-
   market names aren't discounted as cheaply as US ones.

**Bull / Base / Bear** scenarios vary the growth and discount assumptions around
the base case.

### Base cash-flow modes (`--base`)

| mode   | base cash flow                                  | notes |
|--------|-------------------------------------------------|-------|
| `ocf`  | latest operating cash flow                      | **default** for most sectors; treats heavy capex as growth investment (most SWS-like) |
| `ni`   | latest net income                               | **default for financials** (banks/insurers have no meaningful OCF); most conservative |
| `fcf`  | 3-yr average free cash flow                      | smooths one-off capex spikes |
| `auto` | greater of 3-yr avg FCF and net income           | balanced |

The default base is `ocf` (configurable via `default_base`), **except for the
Financial Services sector, which defaults to `ni`** unless you pass `--base`
explicitly.

## Configuration

Settings live in `aegis_config.json` (auto-created with defaults). Key fields:

| key                 | default | meaning |
|---------------------|---------|---------|
| `default_ticker`    | `AAPL`  | ticker used when `single` is called with no argument |
| `default_base`      | `ocf`   | FCFE base mode (see above) |
| `max_stage1_growth` | `0.20`  | cap on stage-1 growth |
| `min_growth`        | `0.01`  | floor on stage-1 growth |
| `equity_risk_premium` | `0.05` | equity risk premium used in CAPM |
| `use_monte_carlo`   | `false` | run Monte Carlo by default |
| `num_simulations`   | `5000`  | Monte Carlo iterations |
| `confidence_weights`| —       | weights for the conviction score (mos / analyst / technical / quality) |

Watchlist tickers are stored in `aegis_watchlist.json`. Reports and a 6-hour
fetch cache are written under `aegis_reports/`.

## Caveats

- Uses **free** `yfinance` data. It does not have access to paid analyst
  forward-FCF forecasts (e.g. S&P Capital IQ), so valuations tend to be more
  conservative than services that do.
- This is a research/educational tool, **not investment advice**.
