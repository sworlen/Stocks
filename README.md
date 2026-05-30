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

python app.py validate                    # regression harness over a diverse basket
python app.py validate AAPL NU BARC.L     # validate specific tickers
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

### Currency normalization

Financial statements are reported in `financialCurrency` (e.g. EUR for ASML/RACE,
GBP for `.L` listings) while the quoted price is in `currency` (often USD for ADRs,
or **GBp/pence** for London). The engine values in the statement currency, then
converts the per-share result into the price currency (handling the GBp = 1/100
GBP "pence trap") so the upside/MOS comparison is apples-to-apples. The currency
is shown in the dashboard, and an FX-normalization flag is raised when conversion
is applied.

### Net cash

For non-financials, **net cash per share** (`totalCash − totalDebt`) is added to
the operating value — a balance-sheet asset the FCFE stream doesn't capture
(e.g. Nintendo's large cash pile). Disable via `add_net_cash=false`. Skipped for
banks, where debt/deposits are operational, not excess cash.

### Pre-profit companies — revenue→margin model

FCFE is undefined when the base cash flow is negative (early-stage growth names
like OUST/RKLB), so those names are routed to a **revenue→margin** model instead
of printing a misleading `$0.00`:

1. Grow revenue/share from today, holding the seed growth for a few years
   (hyper-growth plateau) then fading to the terminal rate.
2. Ramp the net margin from today's (negative) level to a **sector-normal
   target margin**.
3. Discount positive interim earnings at the cost of equity and capitalize the
   final-year earnings at a **sector exit P/E**.
4. Deflate for stock-comp **dilution**, then add net cash.

The seed growth blends reported YoY revenue growth with the historical revenue
CAGR; target margins / exit multiples come from a built-in sector table
(`SECTOR_MARGIN_EXIT`). This is a deliberately **conservative fundamental
floor** — names priced mostly on long-dated optionality (e.g. RKLB) can still
trade well above it, which the dashboard flags.

### Banks & insurers — residual-income / justified-P/B model

Banks have no meaningful free cash flow (deposits/loans dominate the cash-flow
statement), so an FCFE/NI DCF with a Gordon terminal badly overstates them
(JPM read +57%, Barclays +325%). Financials are instead valued off **equity**:

`value = book value + Σ PV(excess return) + PV(terminal residual income)`

1. Excess return each year = `(ROE − cost of equity) × book value`. With ROE
   equal to the cost of equity the bank is worth exactly book (P/B = 1); a
   durable ROE above it earns a premium to book.
2. ROE fades linearly from today's level to a **15% long-run ROE** over 12
   years; book value compounds at the retention rate (`1 − payout`).
3. The final-year residual income is capitalized as a perpetuity growing at the
   terminal rate.

Routed in for `sector == "Financial Services"` only, and only when the book
value / ROE pass a sanity guard (positive, and `ROE × book ≈ trailing EPS`);
otherwise the name falls back to the FCFE/NI path. Net cash is **not** added
(deposits/debt are operational for a bank). Statement-currency in, price-currency
out — so Barclays' GBP book correctly converts to a GBp (pence) fair value.

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
| `add_net_cash`      | `true`  | add net cash per share to non-financials |
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
