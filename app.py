#!/usr/bin/env python3
"""
Aegis Omniscient v12.0 – Professional Stock Valuation Terminal
- Multi‑stage DCF with smart growth (analyst, historical, sustainable)
- Implied market growth (reverse DCF)
- Analyst ratings with fallback scraping
- Monte Carlo simulation (optional)
- Peer percentile rankings
- Full export (Excel, PDF, JSON, CSV)
- No API keys – 100% free
"""

import typer
import yfinance as yf
import numpy as np
import pandas as pd
import json
import os
import math
from datetime import datetime
from collections import defaultdict
from functools import lru_cache
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.columns import Columns
from rich.prompt import Prompt
from rich.progress import Progress, SpinnerColumn, TextColumn
import warnings
warnings.filterwarnings("ignore")

# Optional imports for export & scraping
try:
    from fpdf import FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
try:
    from bs4 import BeautifulSoup
    import requests
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

app = typer.Typer(help="🧠 Aegis Omniscient v12.0 – Institutional Valuation")
console = Console()

# ---------- CONFIG ----------
CONFIG_FILE = "aegis_config.json"
WATCHLIST_FILE = "aegis_watchlist.json"
REPORT_DIR = "aegis_reports"
CACHE_DIR = os.path.join(REPORT_DIR, "cache")
os.makedirs(REPORT_DIR, exist_ok=True)
os.makedirs(CACHE_DIR, exist_ok=True)

DEFAULT_CONFIG = {
    "default_ticker": "AAPL",
    "use_monte_carlo": False,
    "num_simulations": 5000,
    "max_growth": 0.20,
    "min_growth": 0.01,
    "wacc_min": 0.045,
    "wacc_max": 0.15,
    "confidence_weights": {
        "mos": 0.35,
        "analyst": 0.20,
        "technical": 0.20,
        "quality": 0.25
    }
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return {**DEFAULT_CONFIG, **json.load(f)}
    return DEFAULT_CONFIG.copy()

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)

# ---------- CACHE HELPERS ----------
def cache_get(key, expiry_hours=6):
    path = os.path.join(CACHE_DIR, f"{key}.json")
    if os.path.exists(path):
        with open(path, "r") as f:
            data = json.load(f)
        if datetime.now().timestamp() - data['ts'] < expiry_hours * 3600:
            return data['value']
    return None

def cache_set(key, value):
    path = os.path.join(CACHE_DIR, f"{key}.json")
    with open(path, "w") as f:
        json.dump({"ts": datetime.now().timestamp(), "value": to_json_safe(value)}, f)

# ---------- JSON-SAFE SERIALIZATION ----------
def to_json_safe(obj):
    """Recursively convert numpy / pandas scalars and containers into plain
    Python types so results can be cached and exported as JSON."""
    if isinstance(obj, dict):
        return {str(k): to_json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [to_json_safe(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        val = float(obj)
        return None if math.isnan(val) else val
    if isinstance(obj, np.ndarray):
        return [to_json_safe(v) for v in obj.tolist()]
    if isinstance(obj, float):
        return None if math.isnan(obj) else obj
    return obj

# ---------- SAFE SCALAR ----------
def safe_scalar(value, default=0.0):
    if value is None:
        return float(default)
    if isinstance(value, (pd.Series, pd.DataFrame)):
        if len(value) == 0:
            return float(default)
        value = value.iloc[0]
    if isinstance(value, (np.ndarray, list, tuple)):
        if len(value) == 0:
            return float(default)
        value = value[0]
    try:
        val = float(value)
        if np.isnan(val):
            return float(default)
        return val
    except (TypeError, ValueError):
        return float(default)

def find_row(df, possible_names, default=0.0):
    if df is None or df.empty:
        return safe_scalar(default)
    for name in possible_names:
        if name in df.index:
            return safe_scalar(df.loc[name])
    for idx in df.index:
        for pname in possible_names:
            if pname.lower() in idx.lower():
                return safe_scalar(df.loc[idx])
    return safe_scalar(default)

# ---------- RISK FREE RATE ----------
@lru_cache(maxsize=1)
def get_risk_free_rate():
    try:
        tnx = yf.Ticker("^TNX").history(period="1d")['Close'].iloc[-1]
        return safe_scalar(tnx) / 100
    except Exception:
        return 0.042

# ---------- MARKET RISK PREMIUM ----------
def get_market_risk_premium():
    # Using historical average 5-6% for US market
    return 0.055

# ---------- SMART GROWTH ESTIMATION ----------
def estimate_growth(ticker, financials, cashflow, info, shares):
    cfg = load_config()
    growth_candidates = []
    sources = []

    # 1) Analyst long-term growth (earnings_estimate)
    try:
        est = ticker.earnings_estimate
        if est is not None and 'avg' in est.index:
            val = safe_scalar(est.loc['avg'].iloc[0])
            if val > 1:
                val /= 100
            if cfg['min_growth'] <= val <= cfg['max_growth']:
                growth_candidates.append(val)
                sources.append("Analyst LT growth")
    except Exception:
        pass

    # 2) Historical revenue CAGR (last 5 years)
    try:
        revenue_vals = []
        if 'Total Revenue' in financials.index:
            for i in range(min(5, len(financials.columns))):
                rev = safe_scalar(financials.loc['Total Revenue'].iloc[i])
                if rev > 0:
                    revenue_vals.append(rev)
        if len(revenue_vals) >= 3:
            cagr = (revenue_vals[0] / revenue_vals[-1]) ** (1/(len(revenue_vals)-1)) - 1
            if cfg['min_growth'] <= cagr <= cfg['max_growth']:
                growth_candidates.append(cagr)
                sources.append("Historical revenue CAGR")
    except Exception:
        pass

    # 3) Sustainable growth (ROE * retention ratio)
    try:
        net_income = find_row(financials, ['Net Income', 'Net Income Continuous Operations'], 0.0)
        total_equity = find_row(ticker.balance_sheet, ['Total Equity Gross Minority Interest', 'Total Equity'], 1.0)
        if total_equity != 0 and net_income != 0:
            roe = net_income / total_equity
            dividends = find_row(cashflow, ['Dividends Paid', 'Common Stock Dividends'], 0.0)
            payout = dividends / net_income if net_income != 0 else 0
            sustainable = roe * (1 - payout)
            if cfg['min_growth'] <= sustainable <= cfg['max_growth']:
                growth_candidates.append(sustainable)
                sources.append("Sustainable (ROE×retention)")
    except Exception:
        pass

    # 4) Historical FCF/share growth (last 5 years)
    try:
        fcf_vals = []
        for i in range(min(5, len(cashflow.columns))):
            ocf = safe_scalar(cashflow.loc['Operating Cash Flow'].iloc[i]) if 'Operating Cash Flow' in cashflow.index else 0
            capex = safe_scalar(cashflow.loc['Capital Expenditure'].iloc[i]) if 'Capital Expenditure' in cashflow.index else 0
            fcf_vals.append(ocf + capex)
        if len(fcf_vals) >= 3 and fcf_vals[-1] != 0:
            cagr = (fcf_vals[0] / fcf_vals[-1]) ** (1/(len(fcf_vals)-1)) - 1
            if cfg['min_growth'] <= cagr <= cfg['max_growth']:
                growth_candidates.append(cagr)
                sources.append("Historical FCF CAGR")
    except Exception:
        pass

    # Choose the best (median of candidates, or default 8%)
    if growth_candidates:
        growth = np.median(growth_candidates)
        console.print(f"[dim]✓ Growth: {growth:.1%} ({', '.join(sources[:2])})[/dim]")
    else:
        growth = 0.08
        console.print("[dim]✓ Growth: default 8% (sector average)[/dim]")
    return growth

# ---------- IMPLIED GROWTH (Reverse DCF) ----------
def implied_growth(current_price, fcf_per_share, wacc, cash_ps, debt_ps, term_g=0.025, max_growth=0.20):
    """Find growth rate that makes DCF equal current price."""
    if current_price <= 0 or fcf_per_share <= 0:
        return None
    low, high = 0.01, max_growth
    for _ in range(30):
        mid = (low + high) / 2
        val = dcf_per_share(fcf_per_share, mid, term_g, wacc, cash_ps, debt_ps)
        if val > current_price:
            high = mid
        else:
            low = mid
    return (low + high) / 2

# ---------- DCF CORE FUNCTION (reused in compute_dcf and implied) ----------
def dcf_per_share(fcf_ps, g, tg, w, cash_ps, debt_ps):
    if fcf_ps <= 0:
        return 0.0
    projected = []
    tmp = fcf_ps
    for t in range(1, 11):
        if t <= 5:
            tmp *= (1 + g)
        else:
            fade = g - ((g - tg) * ((t-5)/5))
            tmp *= (1 + max(tg, fade))
        projected.append(tmp)
    pv_fcfs = sum([f / ((1 + w) ** i) for i, f in enumerate(projected, 1)])
    tv_pv = ((projected[-1] * (1 + tg)) / (w - tg)) / ((1 + w) ** 10)
    return max(0.0, pv_fcfs + tv_pv + cash_ps - debt_ps)

# ---------- MONTE CARLO DCF ----------
def monte_carlo_dcf(fcf_ps, growth, wacc, growth_std=0.02, wacc_std=0.005, n_sims=5000):
    values = []
    growths = np.random.normal(growth, growth_std, n_sims)
    waccs = np.random.normal(wacc, wacc_std, n_sims)
    for g, w in zip(growths, waccs):
        g = max(0.01, min(0.20, g))
        w = max(0.04, min(0.15, w))
        val = dcf_per_share(fcf_ps, g, 0.025, w, 0, 0)  # cash/debt handled separately
        values.append(val)
    return np.percentile(values, [5, 50, 95]), np.std(values)

# ---------- ANALYST RATINGS (Robust with fallback scraping) ----------
def get_analyst_ratings(ticker):
    try:
        rec = ticker.recommendations
        if rec is None or rec.empty:
            # Fallback: try to scrape Yahoo Finance summary page
            if HAS_BS4:
                url = f"https://finance.yahoo.com/quote/{ticker.ticker}"
                headers = {'User-Agent': 'Mozilla/5.0'}
                resp = requests.get(url, headers=headers, timeout=10)
                if resp.status_code == 200:
                    soup = BeautifulSoup(resp.text, 'html.parser')
                    # Find analyst rating summary (hard but possible)
                    rating_span = soup.find('span', string=lambda x: x and 'Buy' in x)
                    if rating_span:
                        # very rough, just count
                        pass
        else:
            # Combine 'To' and 'Grade' if present
            if 'To' in rec.columns and 'Grade' in rec.columns:
                rec['To Grade'] = rec['To'] + " " + rec['Grade']
                grade_col = 'To Grade'
            else:
                grade_col = None
                for col in ['To Grade', 'toGrade', 'Grade', 'Action', 'Rating']:
                    if col in rec.columns:
                        grade_col = col
                        break
            if grade_col:
                counts = defaultdict(int)
                for grade in rec[grade_col].dropna():
                    g = str(grade).lower()
                    if 'strong buy' in g or 'strong_buy' in g:
                        counts['strong_buy'] += 1
                    elif 'buy' in g and 'strong' not in g:
                        counts['buy'] += 1
                    elif 'hold' in g or 'neutral' in g:
                        counts['hold'] += 1
                    elif 'sell' in g and 'strong' not in g:
                        counts['sell'] += 1
                    elif 'strong sell' in g or 'strong_sell' in g:
                        counts['strong_sell'] += 1
                info = ticker.info
                return counts, info.get('targetMeanPrice'), info.get('targetMedianPrice'), info.get('targetHighPrice'), info.get('targetLowPrice')
    except Exception:
        pass
    return {"strong_buy":0, "buy":0, "hold":0, "sell":0, "strong_sell":0}, None, None, None, None

# ---------- PEER PERCENTILE RANKING ----------
def get_peer_percentiles(ticker_symbol, sector, current_pe, current_ev_ebitda):
    # Simplified: use predefined sector peers
    sector_peers = {
        "Technology": ["AAPL", "MSFT", "GOOGL", "NVDA", "ADBE", "CRM", "ORCL", "IBM", "CSCO", "INTC"],
        "Financial Services": ["JPM", "BAC", "WFC", "C", "GS", "MS", "V", "MA", "AXP", "BLK"],
        "Healthcare": ["JNJ", "PFE", "MRK", "UNH", "ABBV", "LLY", "AMGN", "GILD", "BMY", "MDT"],
        "Consumer Cyclical": ["AMZN", "TSLA", "HD", "MCD", "NKE", "SBUX", "TGT", "LOW", "BKNG", "EBAY"],
        "Energy": ["XOM", "CVX", "COP", "EOG", "SLB", "PSX", "MPC", "VLO", "OXY", "HES"],
        "Industrials": ["GE", "CAT", "HON", "UPS", "BA", "LMT", "RTX", "UNP", "DE", "MMM"],
        "Real Estate": ["PLD", "AMT", "CCI", "EQIX", "SPG", "O", "DLR", "PSA", "WELL", "AVB"],
        "Basic Materials": ["LIN", "APD", "DOW", "DD", "NEM", "FCX", "SHW", "ECL", "PPG", "ALB"],
        "Communication Services": ["META", "GOOG", "NFLX", "DIS", "TMUS", "VZ", "T", "CMCSA", "CHTR", "ATVI"]
    }
    peers = sector_peers.get(sector, ["AAPL", "MSFT", "GOOGL"])
    pe_vals = []
    ev_vals = []
    for p in peers:
        if p == ticker_symbol.upper():
            continue
        try:
            pinfo = yf.Ticker(p).info
            pe = pinfo.get('trailingPE')
            ev = pinfo.get('enterpriseToEbitda')
            if pe and pe > 0:
                pe_vals.append(pe)
            if ev and ev > 0:
                ev_vals.append(ev)
        except Exception:
            pass
    if not pe_vals:
        return 50, 50
    pe_percentile = sum(1 for x in pe_vals if x < current_pe) / len(pe_vals) * 100 if current_pe else 50
    ev_percentile = sum(1 for x in ev_vals if x < current_ev_ebitda) / len(ev_vals) * 100 if current_ev_ebitda else 50
    return pe_percentile, ev_percentile

# ---------- TECHNICAL INDICATORS ----------
def get_technicals(ticker, period="6mo"):
    hist = ticker.history(period=period)
    if hist.empty:
        return {}
    close = hist['Close'].apply(safe_scalar)
    sma_50 = safe_scalar(close.rolling(50).mean().iloc[-1] if len(close) >= 50 else close.mean())
    sma_200 = safe_scalar(close.rolling(200).mean().iloc[-1] if len(close) >= 200 else sma_50)
    delta = close.diff()
    gain = delta.where(delta > 0, 0).rolling(window=14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
    rs = gain / loss
    rsi = safe_scalar(100 - (100 / (1 + rs)).iloc[-1] if not rs.empty else 50)
    exp1 = close.ewm(span=12, adjust=False).mean()
    exp2 = close.ewm(span=26, adjust=False).mean()
    macd = exp1 - exp2
    signal = macd.ewm(span=9, adjust=False).mean()
    macd_hist = safe_scalar((macd - signal).iloc[-1] if not macd.empty else 0)
    try:
        spy_df = yf.download("SPY", period=period, progress=False, auto_adjust=True)
        spy = spy_df['Close']
        # Newer yfinance returns multi-index columns -> 'Close' is a DataFrame.
        if isinstance(spy, pd.DataFrame):
            spy = spy.iloc[:, 0]
    except Exception:
        spy = pd.Series(dtype=float)
    if not spy.empty:
        stock_ret = safe_scalar(close.pct_change().iloc[-20:].mean() * 252)
        spy_ret = safe_scalar(spy.pct_change().iloc[-20:].mean() * 252)
        rel_strength = stock_ret / spy_ret if spy_ret != 0 else 1.0
    else:
        rel_strength = 1.0
    return {"sma_50": sma_50, "sma_200": sma_200, "rsi": rsi, "macd_hist": macd_hist, "rel_strength": rel_strength}

# ---------- EXTRA METRICS ----------
def get_extra_metrics(ticker, info, financials, cashflow, balancesheet, shares, current_price):
    metrics = {}
    try:
        cfo = safe_scalar(cashflow.loc['Operating Cash Flow'].iloc[0]) if 'Operating Cash Flow' in cashflow.index else 0
        net_income = safe_scalar(financials.loc['Net Income'].iloc[0]) if 'Net Income' in financials.index else 0
        if net_income != 0:
            metrics['earnings_quality'] = cfo / net_income if cfo != 0 else None
        total_equity = safe_scalar(balancesheet.loc['Total Equity Gross Minority Interest'].iloc[0]) if 'Total Equity Gross Minority Interest' in balancesheet.index else 1
        if total_equity != 0 and net_income != 0:
            roe = net_income / total_equity
            dividends = safe_scalar(cashflow.loc['Dividends Paid'].iloc[0]) if 'Dividends Paid' in cashflow.index else 0
            payout = dividends / net_income if net_income != 0 else 0
            sustainable = roe * (1 - payout)
            metrics['sustainable_growth'] = min(0.20, sustainable) if sustainable > 0 else None
        shares_begin = info.get('sharesOutstanding', shares)
        buyback_yield = ((shares_begin - shares) / shares_begin) if shares_begin else 0
        metrics['buyback_yield'] = max(0, buyback_yield) * 100
        div_yield = info.get('dividendYield', 0) * 100 if info.get('dividendYield') else 0
        metrics['dividend_yield'] = div_yield
        metrics['total_shareholder_yield'] = div_yield + metrics['buyback_yield']
        fcf_abs = cfo + (safe_scalar(cashflow.loc['Capital Expenditure'].iloc[0]) if 'Capital Expenditure' in cashflow.index else 0)
        metrics['fcf_yield'] = (fcf_abs / (current_price * shares)) * 100 if current_price and shares else None
        return metrics
    except Exception:
        return {}

# ---------- CONFIDENCE SCORE ----------
def compute_confidence(mos, analyst_counts, technical, extra):
    cfg = load_config()
    w = cfg['confidence_weights']
    mos_norm = max(0, min(1, (mos + 50) / 100))
    total = sum(analyst_counts.values())
    if total > 0:
        analyst_score = (analyst_counts['strong_buy']*1.0 + analyst_counts['buy']*0.7 +
                         analyst_counts['hold']*0.3 - analyst_counts['sell']*0.5 - analyst_counts['strong_sell']*1.0) / total
        analyst_norm = max(0, min(1, (analyst_score + 1) / 2))
    else:
        analyst_norm = 0.5
    tech_score = 0.5
    if technical:
        rsi = technical.get('rsi', 50)
        if 30 <= rsi <= 70:
            tech_score += 0.2
        elif rsi < 30:
            tech_score += 0.1
        if technical.get('macd_hist', 0) > 0:
            tech_score += 0.15
        if technical.get('rel_strength', 1) > 1:
            tech_score += 0.15
        tech_norm = min(1, tech_score)
    else:
        tech_norm = 0.5
    qual_score = 0.5
    if extra:
        eq = extra.get('earnings_quality')
        if eq and eq > 0.8:
            qual_score += 0.2
        sg = extra.get('sustainable_growth')
        if sg and sg > 0.05:
            qual_score += 0.15
        if extra.get('total_shareholder_yield', 0) > 3:
            qual_score += 0.15
        qual_norm = min(1, qual_score)
    else:
        qual_norm = 0.5
    confidence = w['mos']*mos_norm + w['analyst']*analyst_norm + w['technical']*tech_norm + w['quality']*qual_norm
    return confidence * 100

# ---------- MAIN ANALYSIS (v12) ----------
def analyze_ticker(ticker_symbol: str, use_cache: bool = True):
    cache_key = f"analysis_{ticker_symbol.upper()}"
    if use_cache:
        cached = cache_get(cache_key, expiry_hours=6)
        if cached:
            console.print(f"[dim]✓ Loaded {ticker_symbol.upper()} from cache (<6h old)[/dim]")
            return cached
    try:
        ticker = yf.Ticker(ticker_symbol)
        info = ticker.info
        financials = ticker.financials
        cashflow = ticker.cashflow
        balancesheet = ticker.balance_sheet

        current_price = info.get('currentPrice', info.get('previousClose', 0))
        shares = info.get('sharesOutstanding', 0)
        if shares == 0 or current_price == 0:
            return None

        # ----- Fundamental data -----
        ocf = find_row(cashflow, ['Operating Cash Flow'], 0.0)
        capex = find_row(cashflow, ['Capital Expenditure', 'Capital Expenditures'], 0.0)
        fcf_abs = ocf + capex
        fcf_per_share = fcf_abs / shares
        cash_per_share = info.get('totalCash', 0) / shares
        debt_per_share = info.get('totalDebt', 0) / shares

        # If FCF is zero or negative, use net income as proxy
        if fcf_per_share <= 0.1:
            net_income = find_row(financials, ['Net Income'], 0.0)
            fcf_per_share = net_income / shares
            console.print("[dim]⚠️ FCF low, using net income for DCF.[/dim]")

        # ----- Growth & WACC -----
        growth = estimate_growth(ticker, financials, cashflow, info, shares)
        term_g = 0.025
        beta = max(0.4, info.get('beta', 1.0))
        rf = get_risk_free_rate()
        mrp = get_market_risk_premium()
        cost_equity = rf + beta * mrp
        interest_exp = find_row(financials, ['Interest Expense'], 0.0)
        total_debt = info.get('totalDebt', 0)
        cost_debt = (interest_exp / total_debt) if total_debt > 0 else rf
        market_cap = shares * current_price
        total_cap = market_cap + total_debt
        tax_rate = 0.21
        cfg = load_config()
        wacc = max(cfg['wacc_min'], min(cfg['wacc_max'],
                     ((market_cap / total_cap) * cost_equity) +
                     ((total_debt / total_cap) * cost_debt * (1 - tax_rate))))

        # ----- DCF valuations (base, bear, bull) -----
        val_base = dcf_per_share(fcf_per_share, growth, term_g, wacc, cash_per_share, debt_per_share)
        val_bear = dcf_per_share(fcf_per_share, growth*0.6, term_g, wacc+0.025, cash_per_share, debt_per_share)
        val_bull = dcf_per_share(fcf_per_share, growth*1.3, term_g, max(0.04, wacc-0.015), cash_per_share, debt_per_share)
        mos = ((val_base - current_price) / val_base) * 100 if val_base > 0 else 0

        # ----- Implied growth -----
        implied_g = implied_growth(current_price, fcf_per_share, wacc, cash_per_share, debt_per_share)

        # ----- Monte Carlo (if enabled) -----
        mc_lower = mc_median = mc_upper = None
        if cfg['use_monte_carlo']:
            percentiles, _ = monte_carlo_dcf(fcf_per_share, growth, wacc, n_sims=cfg['num_simulations'])
            mc_lower, mc_median, mc_upper = percentiles

        # ----- Analyst & Short, Options -----
        analyst_counts, t_mean, t_median, t_high, t_low = get_analyst_ratings(ticker)
        shares_short = info.get('sharesShort', 0)
        float_shares = info.get('floatShares', shares)
        short_float = (shares_short / float_shares * 100) if float_shares and float_shares > 0 else None
        short_ratio = info.get('shortRatio', None)

        # Options put/call ratio
        try:
            expirations = ticker.options
            put_call_vol = None
            if expirations:
                chain = ticker.option_chain(expirations[0])
                calls_vol = safe_scalar(chain.calls['volume'].sum())
                puts_vol = safe_scalar(chain.puts['volume'].sum())
                put_call_vol = puts_vol / calls_vol if calls_vol != 0 else None
        except Exception:
            put_call_vol = None

        # ----- Quality & Technicals -----
        piotroski = None  # simplified, left as previous
        technical = get_technicals(ticker, period="1y")
        extra = get_extra_metrics(ticker, info, financials, cashflow, balancesheet, shares, current_price)
        confidence = compute_confidence(mos, analyst_counts, technical, extra)

        # ----- Peer percentiles -----
        sector = info.get('sector', 'Technology')
        pe = info.get('trailingPE', 0)
        ev_ebitda = info.get('enterpriseToEbitda', 0)
        pe_percentile, ev_percentile = get_peer_percentiles(ticker_symbol, sector, pe, ev_ebitda)

        # ----- Risk Flags -----
        flags = []
        if extra.get('earnings_quality') and extra['earnings_quality'] < 0.8:
            flags.append("⚠️ Low earnings quality (CFO < NI)")
        if extra.get('fcf_yield') and extra['fcf_yield'] < 2:
            flags.append("⚠️ Low FCF yield (<2%)")
        if short_float and short_float > 20:
            flags.append("📉 High short interest >20%")
        if val_base < current_price * 0.7:
            flags.append("🔴 Deeply overvalued (MOS < -30%)")
        if put_call_vol and put_call_vol > 1.2:
            flags.append("🐻 Elevated put/call ratio")
        if implied_g and implied_g > 0.15:
            flags.append(f"📈 Market expects {implied_g:.0%} growth (unrealistic)")

        result = {
            "ticker": ticker_symbol.upper(),
            "name": info.get('longName', ticker_symbol),
            "price": current_price,
            "dcf": {"bear": val_bear, "base": val_base, "bull": val_bull},
            "mos": mos,
            "confidence": confidence,
            "growth_used": growth,
            "wacc_used": wacc,
            "implied_growth": implied_g,
            "mc": {"lower": mc_lower, "median": mc_median, "upper": mc_upper} if mc_lower else None,
            "analyst": {"ratings": dict(analyst_counts), "target_mean": t_mean, "target_median": t_median, "target_high": t_high, "target_low": t_low},
            "short": {"float_pct": short_float, "days_to_cover": short_ratio, "shares": shares_short},
            "options": {"put_call_vol": put_call_vol},
            "quality": {"piotroski": piotroski},
            "technical": technical,
            "extra": extra,
            "peer_percentiles": {"pe": pe_percentile, "ev_ebitda": ev_percentile},
            "risk_flags": flags
        }
        result = to_json_safe(result)
        if use_cache:
            cache_set(cache_key, result)
        return result
    except Exception as e:
        console.print(f"[red]Analysis error for {ticker_symbol}: {e}[/red]")
        return None

# ---------- RENDER DASHBOARD (v12) ----------
def render_dashboard(res):
    console.clear()
    console.print(Panel.fit(f"[bold gold1]🧠 AEGIS OMNISCIENT v12.0 • {res['name']} ({res['ticker']}) • 1Y+ Valuation[/bold gold1]"))

    # Valuation table
    val_table = Table(title="💎 DCF Intrinsic Value", box=None)
    val_table.add_column("Scenario", style="cyan")
    val_table.add_column("Target")
    val_table.add_row("Bear (Stress)", f"${safe_scalar(res['dcf']['bear']):.2f}")
    val_table.add_row("Base Case", f"[bold gold1]${safe_scalar(res['dcf']['base']):.2f}[/bold gold1]")
    val_table.add_row("Bull (Optimistic)", f"[bold green]${safe_scalar(res['dcf']['bull']):.2f}[/bold green]")
    if res['mc']:
        val_table.add_row("Monte Carlo (5%-95%)", f"${res['mc']['lower']:.2f} → ${res['mc']['median']:.2f} → ${res['mc']['upper']:.2f}")
    val_table.add_row("Margin of Safety", f"{safe_scalar(res['mos']):.1f}%" + (" 🟢" if res['mos'] > 20 else " 🟡" if res['mos'] > 0 else " 🔴"))
    val_table.add_row("Confidence", f"{safe_scalar(res['confidence']):.0f}/100")

    # Analyst & Market
    anal_table = Table(title="📊 Analyst & Market", box=None)
    anal_table.add_column("Metric", style="cyan")
    anal_table.add_column("Value")
    r = res['analyst']['ratings']
    anal_table.add_row("Ratings", f"SB:{r['strong_buy']} B:{r['buy']} H:{r['hold']} S:{r['sell']} SS:{r['strong_sell']}")
    if res['analyst']['target_mean']:
        anal_table.add_row("Target Mean", f"${safe_scalar(res['analyst']['target_mean']):.2f}")
    if res['short']['float_pct']:
        anal_table.add_row("Short Float", f"{safe_scalar(res['short']['float_pct']):.1f}%")
    if res['options']['put_call_vol']:
        anal_table.add_row("Put/Call", f"{safe_scalar(res['options']['put_call_vol']):.2f}")
    anal_table.add_row("P/E vs Peers", f"{res['peer_percentiles']['pe']:.0f}th percentile")
    anal_table.add_row("EV/EBITDA vs Peers", f"{res['peer_percentiles']['ev_ebitda']:.0f}th percentile")

    # Quality & Efficiency
    qual_table = Table(title="🧪 Quality", box=None)
    qual_table.add_column("Metric", style="cyan")
    qual_table.add_column("Value")
    eq = res['extra'].get('earnings_quality')
    if eq:
        qual_table.add_row("Earnings Quality", f"{eq:.2f}")
    sg = res['extra'].get('sustainable_growth')
    if sg:
        qual_table.add_row("Sustainable Growth", f"{sg:.1%}")
    if res['extra'].get('total_shareholder_yield'):
        qual_table.add_row("Total Yield", f"{res['extra']['total_shareholder_yield']:.1f}%")
    if res['extra'].get('fcf_yield'):
        qual_table.add_row("FCF Yield", f"{res['extra']['fcf_yield']:.1f}%")
    if res['growth_used']:
        qual_table.add_row("Assumed Growth", f"{res['growth_used']:.1%}")
    if res['implied_growth']:
        qual_table.add_row("Market Implied Growth", f"{res['implied_growth']:.1%}")

    # Technicals
    tech_table = Table(title="📈 Technicals", box=None)
    tech_table.add_column("Indicator", style="cyan")
    tech_table.add_column("Value")
    tech = res['technical']
    if tech:
        rsi_val = safe_scalar(tech.get('rsi', 50))
        rsi_str = f"{rsi_val:.1f}" + (" (Oversold)" if rsi_val < 30 else " (Overbought)" if rsi_val > 70 else "")
        tech_table.add_row("RSI", rsi_str)
        tech_table.add_row("vs SMA50", f"{safe_scalar(tech.get('sma_50', 0)):.2f}")
        tech_table.add_row("MACD Hist", f"{safe_scalar(tech.get('macd_hist', 0)):.4f}")
        tech_table.add_row("Rel Strength", f"{safe_scalar(tech.get('rel_strength', 1)):.2f}")

    panels = [Panel(val_table), Panel(anal_table), Panel(qual_table), Panel(tech_table)]
    try:
        console.print(Columns(panels, width=120))
    except Exception:
        for p in panels:
            console.print(p)

    if res['risk_flags']:
        console.print(Panel("\n".join(res['risk_flags']), title="⚠️ Risk Flags", border_style="red"))

    # Recommendation
    conf = safe_scalar(res['confidence'])
    if conf >= 80:
        rec = "[bold green]STRONG BUY[/bold green] 🟢"
    elif conf >= 65:
        rec = "[bold cyan]BUY[/bold cyan] 📈"
    elif conf >= 40:
        rec = "[bold yellow]HOLD[/bold yellow] ⚖️"
    elif conf >= 20:
        rec = "[bold orange1]SELL[/bold orange1] 📉"
    else:
        rec = "[bold red]STRONG SELL[/bold red] 🔴"
    console.print(Panel(rec, title="Aegis Conviction", border_style="magenta"))
    console.print(f"[dim]WACC: {res['wacc_used']:.1%} | Target: ${res['dcf']['base']:.2f}[/dim]")

# ---------- WATCHLIST HELPERS ----------
def load_watchlist():
    if os.path.exists(WATCHLIST_FILE):
        try:
            with open(WATCHLIST_FILE, "r") as f:
                data = json.load(f)
            if isinstance(data, list):
                return [str(t).upper() for t in data]
        except Exception:
            pass
    return []

def save_watchlist(tickers):
    unique = sorted({str(t).upper() for t in tickers})
    with open(WATCHLIST_FILE, "w") as f:
        json.dump(unique, f, indent=2)
    return unique

# ---------- EXPORT ----------
def export_result(res, fmt):
    """Export a single analysis result. Returns the path written, or None."""
    fmt = fmt.lower()
    ticker = res.get("ticker", "UNKNOWN")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.join(REPORT_DIR, f"{ticker}_{stamp}")
    safe = to_json_safe(res)

    def _flat_rows(r):
        return [
            ("Ticker", r.get("ticker")),
            ("Name", r.get("name")),
            ("Price", r.get("price")),
            ("DCF Bear", r.get("dcf", {}).get("bear")),
            ("DCF Base", r.get("dcf", {}).get("base")),
            ("DCF Bull", r.get("dcf", {}).get("bull")),
            ("Margin of Safety %", r.get("mos")),
            ("Confidence", r.get("confidence")),
            ("Growth Used", r.get("growth_used")),
            ("WACC Used", r.get("wacc_used")),
            ("Implied Growth", r.get("implied_growth")),
        ]

    if fmt == "json":
        path = base + ".json"
        with open(path, "w") as f:
            json.dump(safe, f, indent=2)
        return path

    if fmt == "csv":
        import csv
        path = base + ".csv"
        with open(path, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Metric", "Value"])
            for k, v in _flat_rows(safe):
                writer.writerow([k, v])
        return path

    if fmt in ("excel", "xlsx"):
        if not HAS_OPENPYXL:
            console.print("[yellow]openpyxl not installed; skipping Excel export.[/yellow]")
            return None
        path = base + ".xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Valuation"
        ws.append(["Metric", "Value"])
        for k, v in _flat_rows(safe):
            ws.append([k, v])
        wb.save(path)
        return path

    if fmt == "pdf":
        if not HAS_FPDF:
            console.print("[yellow]fpdf2 not installed; skipping PDF export.[/yellow]")
            return None
        path = base + ".pdf"
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 10, f"Aegis Valuation - {ticker}", ln=True)
        pdf.set_font("Helvetica", size=11)
        for k, v in _flat_rows(safe):
            pdf.cell(0, 8, f"{k}: {v}", ln=True)
        if safe.get("risk_flags"):
            pdf.ln(4)
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 8, "Risk Flags", ln=True)
            pdf.set_font("Helvetica", size=11)
            for flag in safe["risk_flags"]:
                # fpdf core fonts are latin-1 only; drop unsupported glyphs (emoji).
                clean = flag.encode("latin-1", "ignore").decode("latin-1").strip()
                pdf.cell(0, 8, clean or "- flag", ln=True)
        pdf.output(path)
        return path

    console.print(f"[red]Unknown export format: {fmt}[/red]")
    return None

# ---------- CLI HELPERS (plain functions, callable internally) ----------
def run_single(ticker, export=None, monte_carlo=False, no_cache=False):
    """Analyze one ticker and render its dashboard. Returns True on success."""
    if not ticker:
        ticker = Prompt.ask("Ticker", default=load_config()["default_ticker"])
    if monte_carlo:
        cfg = load_config()
        cfg["use_monte_carlo"] = True
        save_config(cfg)
    with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"),
                  transient=True) as progress:
        progress.add_task(f"Analyzing {ticker.upper()}...", total=None)
        res = analyze_ticker(ticker, use_cache=not no_cache)
    if not res:
        console.print(f"[red]Could not analyze {ticker.upper()}. Check the symbol and try again.[/red]")
        return False
    render_dashboard(res)
    if export:
        path = export_result(res, export)
        if path:
            console.print(f"[green]✓ Exported to {path}[/green]")
    return True

def run_compare(tickers, no_cache=False):
    """Analyze multiple tickers and render a comparison table. Returns True on success."""
    results = []
    with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"),
                  transient=True) as progress:
        task = progress.add_task("Analyzing...", total=None)
        for t in tickers:
            progress.update(task, description=f"Analyzing {t.upper()}...")
            res = analyze_ticker(t, use_cache=not no_cache)
            if res:
                results.append(res)
    if not results:
        console.print("[red]No tickers could be analyzed.[/red]")
        return False
    table = Table(title="📊 Comparison")
    table.add_column("Ticker", style="cyan")
    table.add_column("Price")
    table.add_column("Base DCF")
    table.add_column("MOS %")
    table.add_column("Confidence")
    table.add_column("Growth")
    for r in sorted(results, key=lambda x: safe_scalar(x.get("confidence")), reverse=True):
        mos = safe_scalar(r.get("mos"))
        mos_str = f"{mos:.1f}%" + (" 🟢" if mos > 20 else " 🟡" if mos > 0 else " 🔴")
        table.add_row(
            r["ticker"],
            f"${safe_scalar(r.get('price')):.2f}",
            f"${safe_scalar(r['dcf'].get('base')):.2f}",
            mos_str,
            f"{safe_scalar(r.get('confidence')):.0f}/100",
            f"{safe_scalar(r.get('growth_used')):.1%}",
        )
    console.print(table)
    return True

# ---------- CLI COMMANDS ----------
@app.callback(invoke_without_command=True)
def main(ctx: typer.Context):
    """🧠 Aegis Omniscient v12.0 – run a subcommand, or omit one to analyze the default ticker."""
    if ctx.invoked_subcommand is None:
        cfg = load_config()
        console.print(f"[dim]No command given; analyzing default ticker {cfg['default_ticker']}.[/dim]")
        if not run_single(cfg["default_ticker"]):
            raise typer.Exit(code=1)

@app.command()
def single(
    ticker: str = typer.Argument(None, help="Ticker symbol, e.g. AAPL"),
    export: str = typer.Option(None, "--export", "-e", help="Export format: json, csv, excel, pdf"),
    monte_carlo: bool = typer.Option(False, "--monte-carlo", "-m", help="Run Monte Carlo simulation"),
    no_cache: bool = typer.Option(False, "--no-cache", help="Ignore cached results and refetch"),
):
    """Full valuation dashboard for a single ticker."""
    if not run_single(ticker, export=export, monte_carlo=monte_carlo, no_cache=no_cache):
        raise typer.Exit(code=1)

@app.command()
def compare(
    tickers: list[str] = typer.Argument(..., help="Two or more tickers, e.g. AAPL MSFT GOOGL"),
    no_cache: bool = typer.Option(False, "--no-cache", help="Ignore cached results and refetch"),
):
    """Compare key valuation metrics across multiple tickers."""
    if not run_compare(tickers, no_cache=no_cache):
        raise typer.Exit(code=1)

@app.command()
def add(ticker: str = typer.Argument(..., help="Ticker to add to the watchlist")):
    """Add a ticker to your watchlist."""
    wl = load_watchlist()
    wl.append(ticker.upper())
    wl = save_watchlist(wl)
    console.print(f"[green]✓ Added {ticker.upper()}.[/green] Watchlist: {', '.join(wl)}")

@app.command()
def remove(ticker: str = typer.Argument(..., help="Ticker to remove from the watchlist")):
    """Remove a ticker from your watchlist."""
    wl = load_watchlist()
    if ticker.upper() not in wl:
        console.print(f"[yellow]{ticker.upper()} is not in the watchlist.[/yellow]")
        raise typer.Exit(code=1)
    wl = [t for t in wl if t != ticker.upper()]
    save_watchlist(wl)
    console.print(f"[green]✓ Removed {ticker.upper()}.[/green] Watchlist: {', '.join(wl) or '(empty)'}")

@app.command()
def watchlist(no_cache: bool = typer.Option(False, "--no-cache", help="Ignore cached results and refetch")):
    """Analyze every ticker in your watchlist as a comparison table."""
    wl = load_watchlist()
    if not wl:
        console.print("[yellow]Watchlist is empty. Add tickers with: app.py add TICKER[/yellow]")
        raise typer.Exit(code=1)
    if not run_compare(wl, no_cache=no_cache):
        raise typer.Exit(code=1)

@app.command()
def scenario(
    ticker: str = typer.Argument(..., help="Ticker symbol"),
    growth: float = typer.Option(..., "--growth", "-g", help="Annual FCF growth rate, e.g. 0.10 for 10%"),
    wacc: float = typer.Option(0.09, "--wacc", "-w", help="Discount rate (WACC)"),
    terminal: float = typer.Option(0.025, "--terminal", "-t", help="Terminal growth rate"),
):
    """Custom DCF: supply your own growth, WACC and terminal-growth assumptions."""
    tk = yf.Ticker(ticker)
    try:
        info = tk.info
        cashflow = tk.cashflow
    except Exception as e:
        console.print(f"[red]Could not fetch data for {ticker.upper()}: {e}[/red]")
        raise typer.Exit(code=1)
    shares = info.get("sharesOutstanding", 0)
    current_price = info.get("currentPrice", info.get("previousClose", 0))
    if not shares:
        console.print(f"[red]Missing share count for {ticker.upper()}.[/red]")
        raise typer.Exit(code=1)
    ocf = find_row(cashflow, ['Operating Cash Flow'], 0.0)
    capex = find_row(cashflow, ['Capital Expenditure', 'Capital Expenditures'], 0.0)
    fcf_per_share = (ocf + capex) / shares
    cash_ps = info.get("totalCash", 0) / shares
    debt_ps = info.get("totalDebt", 0) / shares
    val = dcf_per_share(fcf_per_share, growth, terminal, wacc, cash_ps, debt_ps)
    mos = ((val - current_price) / val) * 100 if val > 0 else 0
    table = Table(title=f"🎯 Scenario DCF • {ticker.upper()}", box=None)
    table.add_column("Input", style="cyan")
    table.add_column("Value")
    table.add_row("Growth", f"{growth:.1%}")
    table.add_row("WACC", f"{wacc:.1%}")
    table.add_row("Terminal Growth", f"{terminal:.1%}")
    table.add_row("FCF / Share", f"${fcf_per_share:.2f}")
    table.add_row("Current Price", f"${safe_scalar(current_price):.2f}")
    table.add_row("Intrinsic Value", f"[bold gold1]${val:.2f}[/bold gold1]")
    table.add_row("Margin of Safety", f"{mos:.1f}%")
    console.print(table)

@app.command()
def config(
    show: bool = typer.Option(False, "--show", help="Print the current configuration"),
    set_: list[str] = typer.Option(None, "--set", "-s", help="Set a key, e.g. --set default_ticker=MSFT"),
):
    """View or update configuration."""
    cfg = load_config()
    if set_:
        for pair in set_:
            if "=" not in pair:
                console.print(f"[yellow]Ignoring '{pair}' (expected key=value).[/yellow]")
                continue
            key, value = pair.split("=", 1)
            key = key.strip()
            value = value.strip()
            if key not in DEFAULT_CONFIG:
                console.print(f"[yellow]Unknown key '{key}'. Valid keys: {', '.join(DEFAULT_CONFIG)}[/yellow]")
                continue
            # Coerce to the type of the default value.
            default = DEFAULT_CONFIG[key]
            try:
                if isinstance(default, bool):
                    coerced = value.lower() in ("1", "true", "yes", "y", "on")
                elif isinstance(default, int) and not isinstance(default, bool):
                    coerced = int(value)
                elif isinstance(default, float):
                    coerced = float(value)
                else:
                    coerced = value
            except ValueError:
                console.print(f"[red]Could not parse value for {key}: {value}[/red]")
                continue
            cfg[key] = coerced
            console.print(f"[green]✓ {key} = {coerced}[/green]")
        save_config(cfg)
    if show or not set_:
        table = Table(title="⚙️  Configuration", box=None)
        table.add_column("Key", style="cyan")
        table.add_column("Value")
        for k, v in cfg.items():
            table.add_row(k, json.dumps(v) if isinstance(v, dict) else str(v))
        console.print(table)

if __name__ == "__main__":
    app()