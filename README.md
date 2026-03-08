# Nifty_50_Stock_Analyzer
Nifty_50_Stock_Analyzer will provide Technical and functional report over gmail
[README.md](https://github.com/user-attachments/files/25820609/README.md)
# 💎 NIFTY 100 Stock Analyzer v5.5

A fully automated **Technical + Fundamental stock analysis engine** for NIFTY 100 stocks. Runs daily, generates a dark-themed HTML report, and optionally delivers it to your inbox. Built entirely in Python using `yfinance`.

---

## 📸 What the Report Looks Like

The report opens as a single HTML file in your browser with:

- **Live clock** (IST) and Nifty 50 / Bank Nifty index prices
- **KPI band** — counts of Strong Buy / Buy / Hold / Sell / Strong Sell across all 100 stocks
- **Top Buy Table** — sector-diversified shortlist with full trade setup
- **Top Sell Table** — confirmed downtrend stocks with short setup
- **Full Watchlist** — all 97–100 analyzed stocks, filterable by rating

---

## 🚀 Quickstart

### Install dependencies
```bash
pip install yfinance pandas pytz
```

### Run the script
```bash
python Nifty50_stocksanalyzer_v5_4.py
```

This generates `index.html` in the current directory. Open it in any browser.

### Optional: Email delivery
```bash
export GMAIL_USER="your@gmail.com"
export GMAIL_APP_PASSWORD="your_app_password"
export RECIPIENT_EMAIL="recipient@gmail.com"
python Nifty50_stocksanalyzer_v5_4.py
```

> Use a Gmail **App Password** (not your main password). Generate one at: Google Account → Security → 2-Step Verification → App Passwords.

### Optional: Custom output file
```bash
export OUTPUT_FILE="report.html"
python Nifty50_stocksanalyzer_v5_4.py
```

---

## 🏗️ Architecture

```
generate_complete_report()
    └── analyze_all_stocks()             ← loops all 100 symbols
            └── analyze_stock(symbol)    ← main logic per stock
                    ├── yfinance fetch (1 year price + fundamentals)
                    ├── Technical Score  (−6 to +6)
                    ├── Fundamental Score (0 to 100)
                    ├── Combined Score   (0 to 100)
                    ├── Veto Gate        (hard cap to HOLD if 3+ bearish)
                    └── result dict      (50+ fields)
    └── get_top_recommendations()        ← filters Buy + Sell tables
    └── generate_html()                  ← builds full HTML report
    └── save_html()
    └── send_email()  (optional)
```

---

## 🧮 How Scoring Works

Every stock gets **two independent scores** which are then combined.

---

### 1. Technical Score (−6 to +6)

Measures price momentum and trend strength from chart data. Each signal adds or subtracts points. The score is **capped at +6 / floored at −6**.

| Signal | Condition | Points |
|--------|-----------|--------|
| Price vs SMA20 | Price above SMA20 | +1 |
| Price vs SMA20 | Price below SMA20 | −1 |
| Price vs SMA50 | Price above SMA50 | +1 |
| Price vs SMA50 | Price below SMA50 | −1 |
| Price vs SMA200 | Above SMA200 AND SMA200 rising | +2 |
| Price vs SMA200 | Above SMA200 but SMA200 flat/falling | 0 |
| Price vs SMA200 | Below SMA200 | −2 |
| Double SMA penalty | Price below BOTH SMA20 and SMA50 | −1 |
| SMA20 slope (V53) | SMA20 today < SMA20 five bars ago | −1 |
| Death cross (V53) | SMA20 < SMA50 | −1 |
| RSI zone + direction | See RSI table below | −3 to +2 |
| MACD | MACD line > Signal line | +1 |
| MACD | MACD line < Signal line | −1 |
| ADX strength | ADX > 25 AND price > SMA50 (uptrend) | +1 |
| ADX strength | ADX > 25 AND price < SMA50 (downtrend) | −1 |
| ADX weakness | ADX < 20 (no trend) | −1 |
| Volume ratio | Vol > 1.5× AND price > SMA20 | +1 |
| Volume ratio | Vol < 0.7× (thin volume) | −1 |
| 52-Week High proximity | Within 5% of 52W high AND above SMA200 | +1 |

#### RSI Scoring Detail

RSI is scored based on both its **current value** and its **direction (slope over 5 bars)**:

| RSI Zone | Direction | Score | Label shown |
|----------|-----------|-------|-------------|
| < 30 | Any (above SMA200) | +2 | Oversold ↑ |
| < 30 | Any (below SMA200) | −1 | Oversold (Downtrend) |
| 30–45 | Falling | −2 | Weak & Falling ↓ |
| 30–45 | Flat/Rising | −1 | Weak Momentum ⚠ |
| 45–55 | Rising | +1 | Building ↑ |
| 45–55 | Falling fast (slope < −8) | −2 | Falling Fast ↓ |
| 45–55 | Falling | −1 | Fading ↓ |
| 45–55 | Flat | 0 | Neutral → |
| 55–70 | Rising (RSI < 65) | +1 | Momentum ↑ |
| 55–70 | Rising (RSI > 65) | 0 | Near Overbought ⚠ |
| 55–70 | Falling fast (slope < −8) | −2 | Rolling Over ↓ |
| 55–70 | Falling | −1 | Softening ↓ |
| > 70 | Falling fast (slope < −8) | −3 | Topping Out ⚠ |
| > 70 | Falling | −2 | Fading ↓ |
| > 70 | Flat/Rising | −1 | Overbought |

> **RSI Formula:** Uses Wilder's Smoothing (`ewm(com=13)`) — matches TradingView exactly. Simple rolling average would differ by 5–10 RSI points on NSE stocks.

---

### 2. Fundamental Score (0 to 100)

Measures business quality from financial data. Points are awarded for each metric, with partial credit when yFinance doesn't return a value for NSE stocks.

| Category | Metric | Points |
|----------|--------|--------|
| **Valuation** | PE < 25 (non-financial) or PE < 15 (banks) | +10 |
| | PE 25–35 (non-financial) or PE 15–20 (banks) | +5 |
| | PB < 3 | +5 |
| | PB 3–5 | +3 |
| | PEG < 1 | +10 |
| | PEG 1–2 | +5 |
| | PEG missing | +3 (partial credit) |
| **Profitability** | ROE > 15% | +10 |
| | ROE 10–15% | +5 |
| | ROA > 5% | +5 |
| | ROA 2–5% | +3 |
| | Profit margin > 10% | +10 |
| | Profit margin 5–10% | +5 |
| **Growth** | Revenue growth > 15% | +10 |
| | Revenue growth 10–15% | +7 |
| | Revenue growth 5–10% | +5 |
| | Revenue growth negative | −10 (penalty) |
| | Earnings growth > 15% | +10 |
| | Earnings growth 10–15% | +7 |
| | Earnings growth 5–10% | +5 |
| | Earnings growth negative | −10 (penalty) |
| | *Growth penalty cap* | max −10 total |
| **Balance Sheet** | Debt/Equity < 50 | +15 |
| | Debt/Equity 50–100 | +7 |
| | Current Ratio > 1.5 | +10 |
| | Current Ratio 1.0–1.5 | +5 |
| | Current Ratio missing | +3 (partial credit) |
| | Free Cash Flow positive | +15 |
| **Income** | Dividend yield > 1% | +5 |
| | Beta 0.8–1.2 (steady) | +5 |

> **Sector-adjusted PE:** Banks and Financial Services are evaluated against PE < 15 (not < 25), because banks structurally trade at lower multiples due to capital requirements and NPA provisions.

> **Growth penalty cap:** Cyclical sectors (steel, cement, energy) often show negative quarters temporarily. The total growth penalty is capped at −10 so one bad quarter doesn't destroy an otherwise strong company.

---

### 3. Combined Score (0 to 100)

```
Combined Score = (Tech Score Normalised × Tech Weight) + (Fund Score × Fund Weight)
```

Tech score is normalised from [−6, +6] → [0, 100]:
```
Tech Normalised = ((tech_score + 6) / 12) × 100
```

#### Dynamic Weight System

Weights shift automatically based on how many bearish signals are active:

| Bearish Signals Active | Tech Weight | Fund Weight | Label |
|------------------------|-------------|-------------|-------|
| 0 or 1 | 35% | 65% | Normal |
| 2 or more | 50% | 50% | Downtrend Override |

**Why:** In normal conditions, fundamentals dominate (a good business deserves a higher score). But when 2+ bearish signals fire simultaneously, fundamentals no longer override a clearly deteriorating chart. The weights equalize so both dimensions matter equally.

#### Analyst Consensus Nudge

After the combined score is calculated, a small adjustment is applied based on the Wall Street consensus from yFinance (`recommendationKey`):

| Analyst Rating | Adjustment | Condition |
|----------------|------------|-----------|
| Buy / Strong Buy | +5 points | Only if tech score ≥ 2 |
| Hold | 0 | No effect |
| Sell / Strong Sell | −5 points | Always applies |

> **Important:** yFinance analyst data is US-centric. For most NSE stocks, coverage is sparse and ratings may be stale. The nudge is intentionally small (±5) and the buy bonus requires tech confirmation — it cannot rescue a stock with a weak chart.

---

### 4. Rating Thresholds

| Combined Score | Rating | Recommendation |
|----------------|--------|----------------|
| ≥ 70 | ⭐⭐⭐⭐⭐ STRONG BUY | STRONG BUY |
| ≥ 50 | ⭐⭐⭐⭐ BUY | BUY |
| ≥ 40 | ⭐⭐⭐ HOLD | HOLD |
| ≥ 28 | ⭐⭐ SELL | SELL |
| < 28 | ⭐ STRONG SELL | STRONG SELL |

---

## 🚫 Veto Gate (Trend Override)

Even if a stock scores above 50 (BUY territory) due to strong fundamentals, it is **hard-capped to HOLD** if 3 or more of the following 6 bearish signals are active simultaneously:

| # | Signal | Condition |
|---|--------|-----------|
| 1 | SMA20 declining | SMA20 today < SMA20 five bars ago |
| 2 | Death cross forming | SMA20 < SMA50 |
| 3 | MACD bearish | MACD line < Signal line |
| 4 | Momentum lost | RSI < 50 |
| 5 | Below medium trend | Price < SMA50 |
| 6 | RSI falling fast | RSI direction = Falling AND slope > 8 points |

**Why threshold = 3, not 2:**
In a broad market correction, almost every stock will have 2 bearish signals (typically MACD bearish + RSI < 50 from general selling). A threshold of 2 would veto almost everything in a bear market, leaving only 5–6 stocks in the buy table. Threshold of 3 requires genuinely confirmed multi-signal deterioration — e.g. SMA declining + death cross + MACD bearish = real distribution top.

When the veto fires, the stock is **excluded from the buy table** and shows a `🚫 Trend Veto` badge in the watchlist.

---

## 📊 Buy Table Filters

After rating, stocks go through additional filters before appearing in the Top Buy table. These are applied in order:

1. **Upside > 0%** — Target must be above current price
2. **Target 1 > Price** — Calculated target must be above current price
3. **RSI Safety Gate** — Blocks overbought and topping-out stocks:
   - RSI > 70 → blocked (overbought)
   - RSI > 65 AND Falling → blocked (near overbought and rolling over)
   - RSI > 60 AND slope < −8 → blocked (topping out fast)
4. **Risk:Reward Gate**:
   - Strong Buy: R:R ≥ 1.2×
   - Buy: R:R ≥ 0.6×
5. **Volume Gate** — Vol ratio ≥ 0.6× (thin volume = unreliable move)
6. **Sector Diversity Cap** — Maximum 4 stocks per sector (prevents Financial Services from dominating)
7. **Table size cap** — Maximum 20 stocks shown

---

## 📉 Sell Table Filters

Stocks in SELL / STRONG SELL also go through a safety gate to avoid recommending shorts on stocks that are oversold or recovering:

| Rule | Condition | Why |
|------|-----------|-----|
| Block 1 | RSI < 35 | Oversold — high bounce risk, dangerous to short |
| Block 2 | RSI > 50 AND MACD Bullish | Uptrend — not a valid sell candidate |
| Block 3 | RSI < 45 AND Rising AND slope > 8 | Recovering — don't short a bounce |

---

## 📐 How Targets & Stop Loss Are Calculated

### Stop Loss
Calculated using **ATR (Average True Range)** × a multiplier based on beta:

| Beta | ATR Multiplier | Max Stop % |
|------|---------------|------------|
| < 0.8 (low volatility) | 1.0× | 5% |
| 0.8–1.2 (normal) | 1.5× | 7% |
| > 1.2 (high volatility) | 2.0× | 10% |

Stop is set at: `Price − (ATR × multiplier)`, then capped at the max % to prevent extreme stops on volatile days.

### Targets
Targets are derived from **support/resistance levels** detected algorithmically:
- **Real S/R** — if the script finds genuine resistance levels above the price, targets are set at those levels
- **Beta Cap** — if resistance is too tight (within ATR), a minimum target is enforced
- **ATH Zone** — if price is at all-time highs (no resistance above), targets are projected using ATR × 3

### Risk:Reward Ratio
```
R:R = (Target 1 − Entry Price) ÷ (Entry Price − Stop Loss)
```
- 1.5× and above = excellent
- 1.0–1.5× = good
- 0.8–1.0× = acceptable in bear market
- < 0.8× = blocked from buy table

---

## 📖 How to Read the Report

### Buy Table Columns

| Column | What it shows | How to read it |
|--------|--------------|----------------|
| **Rating / Score** | Stars + combined score 0–100 | Higher is better. 70+ = strong |
| **Upside %** | % from current price to Target 1 | Higher = more room to run |
| **Target (S/R)** | T1 and T2 price levels + stop loss | Your exit and protection levels |
| **ATR** | Daily range in ₹ and % | Higher = more volatile stock |
| **R:R** | Risk:Reward ratio | 1.5× = you risk ₹1 to make ₹1.50 |
| **RSI / Div** | RSI value + direction arrow + slope | ↑ Rising = good entry. ↓ Falling fast = wait |
| **ADX / Vol** | Trend strength + volume vs average | ADX > 25 = strong trend. 2× = high volume |
| **Sup Dist** | % from current price to support | Lower = tighter stop |
| **52W Hi %** | % below 52-week high | −5% = near highs. −30% = far from highs |
| **MACD** | Bullish or Bearish | Bullish = price momentum positive |
| **P/E** | Price to Earnings ratio | Lower = cheaper (vs sector average) |
| **Beta** | Volatility vs market | 1.0 = moves with market. 1.5 = 50% more volatile |
| **Div %** | Dividend yield | Passive income indicator |
| **Quality** | Good / Average / Poor | Based on ROE, margins, FCF combined |
| **Analyst** | Wall St consensus | Buy/Hold/Sell from yFinance |
| **Earnings** | Next earnings date | Avoid holding through unknown earnings |
| **Action** | Final recommendation button | BUY (cyan) / STRONG BUY (green) / HOLD (amber) |

### RSI Direction Arrow Guide

| Display | Meaning | Action |
|---------|---------|--------|
| ↑ +13 (green) | RSI rising strongly | Best entry — momentum building |
| → (grey) | RSI flat | Neutral — watch |
| ↓ −3 (orange) | RSI softening | Caution — wait for stabilisation |
| ↓ −9 (red) | RSI falling fast | Avoid — momentum leaving |
| ↓ −12 (red) | RSI falling very fast | Skip — confirmed distribution |

### VOL/AVG Column

| Value | Meaning |
|-------|---------|
| 2.0× + | Double normal volume — high conviction move |
| 1.5× | 50% above normal — real participation |
| 1.0–1.5× | Normal — ok |
| 0.7–1.0× | Below normal — weak participation |
| < 0.7× | Very thin — price move unreliable |

### Quick 30-Second Checklist Per Stock

```
□ RSI direction → ↑ rising?           YES = proceed  /  ↓ falling fast = SKIP
□ MACD → Bullish?                     YES = proceed  /  Bearish = caution
□ Score → above 65?
□ P/E → reasonable for sector?
□ R:R → above 1.0×?
□ Any red badges (Bear Div)?          YES = SKIP
□ Quality → Good or Average?

5 of 7 pass → BUY
3–4 pass   → WAIT for better entry
< 3 pass   → SKIP this stock
```

### Sector P/E Reference

| Sector | Fair P/E Range |
|--------|---------------|
| Energy (ONGC, Coal India) | 8–12 |
| Metals (Hindalco, Vedanta) | 10–15 |
| Banking / Financial | 10–18 |
| Auto (Bajaj Auto, Maruti) | 20–28 |
| Pharma (Dr Reddy, Sun) | 20–30 |
| IT (TCS, Infosys) | 22–32 |
| FMCG (Marico, Nestle) | 35–55 |
| Industrials (Siemens, ABB) | 30–55 |

---

## ⚙️ Configuration

All key parameters are at the top of the script or in `get_top_recommendations()`:

| Parameter | Default | Description |
|-----------|---------|-------------|
| `VETO_THRESHOLD` | 3 | Bearish signals needed to cap at HOLD |
| `WEIGHT_SHIFT_THRESHOLD` | 2 | Bearish signals to shift weights to 50/50 |
| `STRONG_BUY_THRESHOLD` | 70 | Combined score for STRONG BUY |
| `BUY_THRESHOLD` | 50 | Combined score for BUY |
| `HOLD_THRESHOLD` | 40 | Combined score for HOLD |
| `SELL_THRESHOLD` | 28 | Combined score for SELL |
| `RR_STRONG_BUY` | 1.2× | Min R:R for Strong Buy in buy table |
| `RR_BUY` | 0.6× | Min R:R for Buy in buy table |
| `VOL_MIN` | 0.6× | Min volume ratio for buy table |
| `SECTOR_CAP` | 4 | Max stocks per sector in buy table |
| `MAX_BUY_TABLE` | 20 | Max stocks in buy table |

---

## 🔍 Debug Output

Every run prints a filter trace to the terminal showing exactly why each BUY-rated stock passes or fails the buy table filters:

```
───────────────────────────────────────────────────────────────────────────
  BUY FILTER DEBUG  (14 BUY-rated stocks → target: show all that qualify)
───────────────────────────────────────────────────────────────────────────
  Stock              RSI   Dir      Slp    RR   Vol   Result
  ──────────────────────────────────────────────────────────────────────
  HINDALCO            56   Rising    +3   0.9   1.4   ✅ PASS
  ONGC                59   Falling   -3   1.2   1.8   ✅ PASS
  NTPC                63   Falling   -9   1.0   1.3   ❌ BLOCKED — RSI 63 topping (slope -9)
  BAJAJFINSV          68   Falling  -11   0.7   0.8   ❌ BLOCKED — RSI 68 near-OB + falling
  ...

  Filter summary: 14 BUY rated → 10 after RSI gate → 8 after R:R+Vol gate → sector cap applies
```

---

## 📦 Data Source

All data is fetched from **yFinance** (Yahoo Finance):

| Data | Source field | Used for |
|------|-------------|---------|
| Daily OHLCV | `stock.history(period='1y', auto_adjust=False)` | All technicals |
| PE, PB, PEG | `info['trailingPE']` etc. | Fundamental score |
| ROE, ROA, margins | `info['returnOnEquity']` etc. | Fundamental score |
| D/E, current ratio, FCF | `info['debtToEquity']` etc. | Fundamental score |
| Analyst consensus | `info['recommendationKey']` | Nudge ±5 |
| Analyst target price | `info['targetMeanPrice']` | Target calculation |
| Earnings date | `info['earningsTimestamps']` | Display only |
| Sector | `info['sector']` | PE threshold selection |

> **`auto_adjust=False`** is used intentionally. yFinance's default `auto_adjust=True` adjusts historical closes for dividends and splits, which changes RSI values by 5–10 points vs TradingView. Raw prices match TradingView.

---

## 🔄 Version History

| Version | Key Changes |
|---------|-------------|
| **v5.5** | RSI slope logic (direction-aware scoring), RSI gate on buy table, sell table safety gate, Wilder's RSI formula fix, auto_adjust=False data fix, debug filter trace |
| **v5.4** | Trend Veto Gate (3+ bearish signals → HOLD), SMA200 slope guard, analyst bonus tightened (tech ≥ 2), dynamic weight shift (2+ signals → 50/50) |
| **v5.3** | SMA20 slope penalty (V53-1), death cross penalty (V53-2) — fixes SBIN false BUY |
| **v5.2** | ADX direction-aware (bonus only in uptrend), RSI weak-momentum zone (30–45 = −1), double SMA penalty, sector-adjusted PE for banks |
| **v5.1** | Score threshold calibration (70/50), R:R gate split by rating, 5-day volume average, growth penalty cap |
| **v5.0** | RSI divergence detection, volume gatekeeper, sector diversity cap, FCF weight +15, D/E weight +15 |
| **v4.0** | 65/35 fundamentals/technicals weighting, ADX weak-trend penalty, analyst consensus nudge, 52W high proximity bonus |

---

## ⚠️ Disclaimer

This tool is for **educational and research purposes only**. It is not financial advice. Stock markets carry risk. Always do your own research before investing. Past performance of any indicator does not guarantee future results.

---

## 📄 License

MIT License — free to use, modify, and distribute with attribution.
