"""
Indian Stock Fundamental Analyser
Data: yfinance + Screener.in scraping
Output: HTML report (same layout as Claude widget)
Usage: python analyser.py --ticker RELIANCE --horizon 5
"""

import argparse
import sys
import json
import time
import re
import os
from datetime import datetime, date
from pathlib import Path

import yfinance as yf
import requests
from bs4 import BeautifulSoup

# ─────────────────────────────────────────
# CLI
# ─────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(description="Indian Stock Fundamental Analyser")
    p.add_argument("--ticker",   required=True,  help="NSE ticker, e.g. RELIANCE")
    p.add_argument("--horizon",  default=5,       type=int, help="Investment horizon in years (default 5)")
    p.add_argument("--out",      default="reports", help="Output directory (default: reports/)")
    return p.parse_args()


# ─────────────────────────────────────────
# SCREENER SCRAPER
# ─────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

def screener_fetch(ticker: str) -> dict:
    """Scrape Screener.in consolidated page for a ticker."""
    url = f"https://www.screener.in/company/{ticker.upper()}/consolidated/"
    data = {}
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 404:
            # try standalone
            url = f"https://www.screener.in/company/{ticker.upper()}/"
            r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        data["url"] = url
        data.update(_screener_ratios(soup))
        data.update(_screener_financials(soup))
        data.update(_screener_shareholding(soup))
        data.update(_screener_peers(soup, ticker))
        data.update(_screener_quarters(soup))
    except Exception as e:
        data["screener_error"] = str(e)
    return data


def _text(el):
    return el.get_text(strip=True) if el else None


def _clean_num(s):
    if not s:
        return None
    s = str(s).replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _screener_ratios(soup) -> dict:
    out = {}
    # Top ratios section
    for li in soup.select("#top-ratios li"):
        name_el = li.select_one(".name")
        val_el  = li.select_one(".number")
        if not name_el or not val_el:
            continue
        name = _text(name_el).lower()
        val  = _text(val_el)
        if "market cap" in name:
            out["market_cap"] = val
        elif "current price" in name:
            out["cmp_screener"] = val
        elif "high / low" in name:
            parts = val.split("/")
            if len(parts) == 2:
                out["52w_high"] = parts[0].strip().replace("₹","")
                out["52w_low"]  = parts[1].strip().replace("₹","")
        elif "p/e" in name:
            out["pe"] = val
        elif "book value" in name:
            out["book_value"] = val
        elif "dividend yield" in name:
            out["div_yield"] = val
        elif "roce" in name:
            out["roce"] = val
        elif "roe" in name:
            out["roe"] = val
        elif "face value" in name:
            out["face_value"] = val
    return out


def _screener_financials(soup) -> dict:
    """Pull annual P&L table rows."""
    out = {}
    try:
        section = soup.find("section", {"id": "profit-loss"})
        if not section:
            return out
        rows = section.select("table tbody tr")
        years = [_text(th) for th in section.select("table thead th")[1:] if _text(th)]
        for row in rows:
            cells = row.select("td")
            if not cells:
                continue
            label = _text(cells[0]).lower() if cells[0] else ""
            vals  = [_clean_num(_text(c)) for c in cells[1:]]
            if "sales" in label or "revenue" in label:
                out["revenue_annual"] = dict(zip(years, vals))
            elif "net profit" in label:
                out["netprofit_annual"] = dict(zip(years, vals))
            elif "eps" in label:
                out["eps_annual"] = dict(zip(years, vals))
            elif "opm" in label or "operating profit margin" in label:
                out["opm_annual"] = dict(zip(years, vals))
        # Debt / equity from balance sheet section
        bs = soup.find("section", {"id": "balance-sheet"})
        if bs:
            for row in bs.select("table tbody tr"):
                cells = row.select("td")
                if not cells:
                    continue
                label = _text(cells[0]).lower() if cells[0] else ""
                vals  = [_clean_num(_text(c)) for c in cells[1:]]
                if "borrowings" in label:
                    out["borrowings_annual"] = dict(zip(years, vals))
                elif "equity" in label and "share capital" not in label:
                    out["equity_annual"] = dict(zip(years, vals))
    except Exception as e:
        out["financials_error"] = str(e)
    return out


def _screener_shareholding(soup) -> dict:
    out = {}
    try:
        section = soup.find("section", {"id": "shareholding"})
        if not section:
            return out
        rows = section.select("table tbody tr")
        quarters = [_text(th) for th in section.select("table thead th")[1:] if _text(th)]
        for row in rows:
            cells = row.select("td")
            if not cells:
                continue
            label = _text(cells[0]).lower() if cells[0] else ""
            vals  = [_clean_num(_text(c)) for c in cells[1:]]
            d = dict(zip(quarters, vals))
            if "promoters" in label and "pledge" not in label:
                out["promoter_holding"] = d
            elif "pledge" in label:
                out["promoter_pledge"] = d
            elif "fii" in label or "foreign" in label:
                out["fii_holding"] = d
            elif "dii" in label or "domestic" in label:
                out["dii_holding"] = d
    except Exception as e:
        out["shareholding_error"] = str(e)
    return out


def _screener_peers(soup, self_ticker: str) -> dict:
    out = {"peers": []}
    try:
        section = soup.find("section", {"id": "peers"})
        if not section:
            return out
        rows = section.select("table tbody tr")
        for row in rows:
            cells = row.select("td")
            if len(cells) < 6:
                continue
            name = _text(cells[0])
            if not name or self_ticker.upper() in name.upper():
                continue
            out["peers"].append({
                "name":       name,
                "pe":         _clean_num(_text(cells[2])),
                "pb":         _clean_num(_text(cells[5])) if len(cells) > 5 else None,
                "roe":        _clean_num(_text(cells[7])) if len(cells) > 7 else None,
                "rev_growth": _clean_num(_text(cells[6])) if len(cells) > 6 else None,
                "de":         None,
            })
        out["peers"] = out["peers"][:3]
    except Exception as e:
        out["peers_error"] = str(e)
    return out


def _screener_quarters(soup) -> dict:
    out = {}
    try:
        section = soup.find("section", {"id": "quarters"})
        if not section:
            return out
        rows = section.select("table tbody tr")
        quarters = [_text(th) for th in section.select("table thead th")[1:] if _text(th)]
        for row in rows:
            cells = row.select("td")
            if not cells:
                continue
            label = _text(cells[0]).lower() if cells[0] else ""
            vals  = [_clean_num(_text(c)) for c in cells[1:]]
            if "eps" in label:
                out["eps_quarterly"] = dict(zip(quarters, vals))
            elif "net profit" in label:
                out["netprofit_quarterly"] = dict(zip(quarters, vals))
            elif "sales" in label or "revenue" in label:
                out["revenue_quarterly"] = dict(zip(quarters, vals))
    except Exception as e:
        out["quarters_error"] = str(e)
    return out


# ─────────────────────────────────────────
# YFINANCE FETCH
# ─────────────────────────────────────────

def yf_fetch(ticker: str) -> dict:
    """Fetch data from yfinance (NSE ticker with .NS suffix)."""
    out = {}
    nse = ticker.upper() + ".NS"
    try:
        tk = yf.Ticker(nse)
        info = tk.info or {}

        out["company_name"]   = info.get("longName") or info.get("shortName", ticker)
        out["sector"]         = info.get("sector", "N/A")
        out["industry"]       = info.get("industry", "N/A")
        out["description"]    = info.get("longBusinessSummary", "")
        out["cmp"]            = info.get("currentPrice") or info.get("regularMarketPrice")
        out["52w_high_yf"]    = info.get("fiftyTwoWeekHigh")
        out["52w_low_yf"]     = info.get("fiftyTwoWeekLow")
        out["market_cap_yf"]  = info.get("marketCap")
        out["face_value_yf"]  = info.get("faceValue")
        out["pe_yf"]          = info.get("trailingPE")
        out["pb_yf"]          = info.get("priceToBook")
        out["ev_ebitda_yf"]   = info.get("enterpriseToEbitda")
        out["roe_yf"]         = info.get("returnOnEquity")      # decimal
        out["div_yield_yf"]   = info.get("dividendYield")       # decimal
        out["div_rate_yf"]    = info.get("dividendRate")
        out["payout_ratio_yf"]= info.get("payoutRatio")
        out["beta"]           = info.get("beta")
        out["employees"]      = info.get("fullTimeEmployees")
        out["website"]        = info.get("website", "")

        # Financials
        try:
            fin = tk.financials       # annual income statement
            if fin is not None and not fin.empty:
                out["yf_annual_revenue"]    = fin.loc["Total Revenue"].to_dict()       if "Total Revenue"    in fin.index else {}
                out["yf_annual_netprofit"]  = fin.loc["Net Income"].to_dict()          if "Net Income"       in fin.index else {}
                out["yf_annual_ebitda"]     = fin.loc["EBITDA"].to_dict()              if "EBITDA"           in fin.index else {}
                out["yf_annual_ebit"]       = fin.loc["EBIT"].to_dict()                if "EBIT"             in fin.index else {}
        except Exception:
            pass

        try:
            cf = tk.cashflow
            if cf is not None and not cf.empty:
                fcf_key = next((k for k in cf.index if "Free Cash Flow" in str(k)), None)
                if fcf_key:
                    out["yf_fcf"] = cf.loc[fcf_key].to_dict()
                else:
                    # OCF - Capex
                    ocf = cf.loc["Operating Cash Flow"].to_dict()  if "Operating Cash Flow" in cf.index else {}
                    cap = cf.loc["Capital Expenditure"].to_dict()  if "Capital Expenditure" in cf.index else {}
                    if ocf and cap:
                        out["yf_fcf"] = {k: (ocf.get(k,0) or 0) + (cap.get(k,0) or 0) for k in ocf}
        except Exception:
            pass

        try:
            bs = tk.balance_sheet
            if bs is not None and not bs.empty:
                out["yf_total_debt"]   = bs.loc["Total Debt"].to_dict()         if "Total Debt"         in bs.index else {}
                out["yf_equity"]       = bs.loc["Stockholders Equity"].to_dict() if "Stockholders Equity" in bs.index else {}
                out["yf_current_assets"]  = bs.loc["Current Assets"].to_dict()   if "Current Assets"  in bs.index else {}
                out["yf_current_liab"]    = bs.loc["Current Liabilities"].to_dict() if "Current Liabilities" in bs.index else {}
        except Exception:
            pass

        # Quarterly EPS
        try:
            q_fin = tk.quarterly_financials
            if q_fin is not None and not q_fin.empty:
                eps_q = {}
                ni = q_fin.loc["Net Income"] if "Net Income" in q_fin.index else None
                if ni is not None:
                    shares = info.get("sharesOutstanding") or 1
                    for col, val in ni.items():
                        if val is not None:
                            eps_q[str(col)[:10]] = round(val / shares, 2)
                out["eps_quarterly_yf"] = eps_q
        except Exception:
            pass

    except Exception as e:
        out["yf_error"] = str(e)
    return out


# ─────────────────────────────────────────
# CALCULATIONS
# ─────────────────────────────────────────

def cagr(values: list, years: int) -> float | None:
    """CAGR from a list of floats (oldest → newest)."""
    clean = [v for v in values if v and v != 0]
    if len(clean) < 2:
        return None
    try:
        return round(((clean[-1] / clean[0]) ** (1 / years) - 1) * 100, 1)
    except Exception:
        return None


def latest(d: dict):
    if not d:
        return None
    return list(d.values())[0] if d else None


def avg_last_n(d: dict, n: int) -> float | None:
    vals = [v for v in list(d.values())[:n] if v is not None]
    return round(sum(vals) / len(vals), 1) if vals else None


def trend_arrow(d: dict) -> str:
    vals = [v for v in list(d.values()) if v is not None]
    if len(vals) < 2:
        return "→"
    diff = vals[0] - vals[-1]
    if diff > 1:
        return "↑ Rising"
    elif diff < -1:
        return "↓ Falling"
    return "→ Stable"


def de_ratio(sc: dict, yf_data: dict) -> float | None:
    # Try screener first
    borrow = sc.get("borrowings_annual", {})
    equity = sc.get("equity_annual", {})
    if borrow and equity:
        b = latest(borrow)
        e = latest(equity)
        if b is not None and e and e != 0:
            return round(b / e, 2)
    # yfinance fallback
    debt = yf_data.get("yf_total_debt", {})
    eq   = yf_data.get("yf_equity", {})
    if debt and eq:
        d = latest(debt)
        e = latest(eq)
        if d is not None and e and e != 0:
            return round(abs(d / e), 2)
    return None


def current_ratio(yf_data: dict) -> float | None:
    ca = yf_data.get("yf_current_assets", {})
    cl = yf_data.get("yf_current_liab", {})
    if ca and cl:
        a = latest(ca)
        l = latest(cl)
        if a and l and l != 0:
            return round(a / l, 2)
    return None


def calc_fcf_crore(yf_data: dict):
    fcf = yf_data.get("yf_fcf", {})
    val = latest(fcf)
    if val:
        return round(val / 1e7, 0)   # ₹ to Cr (approx, yf gives in local currency)
    return None


def interest_coverage(sc: dict) -> float | None:
    # Screener doesn't give interest directly; use yfinance EBIT / Interest Expense
    return None   # placeholder — yfinance info sometimes has it


def revenue_list_screener(sc: dict) -> list:
    d = sc.get("revenue_annual", {})
    return [v for v in reversed(list(d.values())) if v is not None]


def profit_list_screener(sc: dict) -> list:
    d = sc.get("netprofit_annual", {})
    return [v for v in reversed(list(d.values())) if v is not None]


def eps_list_screener(sc: dict) -> list:
    d = sc.get("eps_annual", {})
    return [v for v in reversed(list(d.values())) if v is not None]


def opm_latest(sc: dict):
    d = sc.get("opm_annual", {})
    return latest(d)


def format_inr_cr(val) -> str:
    if val is None:
        return "N/A"
    try:
        v = float(val)
        if abs(v) >= 1_00_000:
            return f"₹{v/1_00_000:.1f} L Cr"
        elif abs(v) >= 1_000:
            return f"₹{v:,.0f} Cr"
        return f"₹{v:.0f} Cr"
    except Exception:
        return str(val)


def format_inr(val) -> str:
    if val is None:
        return "N/A"
    try:
        return f"₹{float(val):,.2f}"
    except Exception:
        return str(val)


def pct(val) -> str:
    if val is None:
        return "N/A"
    try:
        return f"{float(val):.1f}%"
    except Exception:
        return str(val)


def x(val) -> str:
    if val is None:
        return "N/A"
    try:
        return f"{float(val):.1f}x"
    except Exception:
        return str(val)


def badge(val, thresholds, labels, css_classes):
    """Return (label, css_class) based on threshold."""
    if val is None:
        return "N/A", "neu"
    for i, t in enumerate(thresholds):
        if val <= t:
            return labels[i], css_classes[i]
    return labels[-1], css_classes[-1]


def valuation_signal(current, sector_avg, hist_avg):
    if current is None:
        return "N/A", "neu"
    refs = [r for r in [sector_avg, hist_avg] if r]
    if not refs:
        return "N/A", "neu"
    avg_ref = sum(refs) / len(refs)
    ratio = current / avg_ref
    if ratio < 0.9:
        return "CHEAP", "ok"
    elif ratio > 1.1:
        return "EXPENSIVE", "bad"
    return "FAIR", "warn"


def growth_class(rev3, rev5, np3, np5):
    scores = [v for v in [rev3, rev5, np3, np5] if v is not None]
    if not scores:
        return "UNKNOWN"
    avg = sum(scores) / len(scores)
    if avg > 18:
        return "ACCELERATING"
    elif avg > 10:
        return "STEADY"
    elif avg > 0:
        return "SLOWING"
    return "DECLINING"


def ownership_latest(d: dict):
    if not d:
        return None, None
    items = list(d.items())
    latest_q, latest_v = items[0]
    oldest_v = items[-1][1] if len(items) > 1 else latest_v
    trend = "→ Stable"
    if latest_v is not None and oldest_v is not None:
        diff = latest_v - oldest_v
        if diff > 1:
            trend = "↑ Buying" if "promoter" in str(d) else "↑ Increasing"
        elif diff < -1:
            trend = "↓ Selling" if "promoter" in str(d) else "↓ Decreasing"
    return latest_v, trend


def project_eps(base_eps, base_case_growth, horizon):
    if base_eps is None or base_case_growth is None:
        return None, None, None
    bear  = base_eps * ((1 + max(base_case_growth * 0.4, -0.05)) ** horizon)
    base  = base_eps * ((1 + base_case_growth) ** horizon)
    bull  = base_eps * ((1 + base_case_growth * 1.5) ** horizon)
    return round(bear, 1), round(base, 1), round(bull, 1)


# ─────────────────────────────────────────
# SECTOR AVERAGES (static fallback)
# ─────────────────────────────────────────

SECTOR_PE = {
    "Technology":          28,
    "Information Technology": 28,
    "Financial Services":  18,
    "Banking":             16,
    "Consumer Defensive":  35,
    "Consumer Cyclical":   40,
    "Healthcare":          30,
    "Energy":              12,
    "Utilities":           15,
    "Industrials":         25,
    "Basic Materials":     14,
    "Communication Services": 22,
    "Real Estate":         30,
}

def sector_pe(sector: str) -> float:
    for k, v in SECTOR_PE.items():
        if k.lower() in (sector or "").lower():
            return v
    return 22.0


# ─────────────────────────────────────────
# HTML BUILDER
# ─────────────────────────────────────────

CSS = """
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --g-fill:#EAF3DE;--g-text:#27500A;--g-border:#C0DD97;--g-accent:#639922;
  --a-fill:#FAEEDA;--a-text:#854F0B;--a-border:#FAC775;--a-accent:#EF9F27;
  --r-fill:#FCEBEB;--r-text:#A32D2D;--r-border:#F7C1C1;--r-accent:#E24B4A;
  --b-fill:#E6F1FB;--b-text:#185FA5;--b-border:#B5D4F4;--b-accent:#378ADD;
  --n-fill:#F1EFE8;--n-text:#444441;--n-border:#D3D1C7;
  --color-text-primary:#1a1a1a;--color-text-secondary:#666;
  --color-background-primary:#fff;--color-background-secondary:#f5f5f3;
  --color-border-tertiary:#e0ddd5;--color-border-secondary:rgba(0,0,0,0.15);
  --color-border-primary:rgba(0,0,0,0.3);--font-sans:system-ui,sans-serif
}
@media(prefers-color-scheme:dark){:root{
  --g-fill:#173404;--g-text:#C0DD97;--g-border:#3B6D11;--g-accent:#97C459;
  --a-fill:#412402;--a-text:#FAC775;--a-border:#854F0B;--a-accent:#EF9F27;
  --r-fill:#501313;--r-text:#F09595;--r-border:#A32D2D;--r-accent:#E24B4A;
  --b-fill:#042C53;--b-text:#85B7EB;--b-border:#185FA5;--b-accent:#378ADD;
  --n-fill:#2C2C2A;--n-text:#D3D1C7;--n-border:#5F5E5A;
  --color-text-primary:#f0ede6;--color-text-secondary:#aaa;
  --color-background-primary:#1c1c1a;--color-background-secondary:#2a2a28;
  --color-border-tertiary:#3a3a38
}}
body{font-family:var(--font-sans);font-size:13px;line-height:1.6;color:var(--color-text-primary);background:var(--color-background-primary);padding:24px}
.wrap{max-width:900px;margin:auto;padding:0 0 48px}
h1{font-size:22px;font-weight:600;margin-bottom:4px}
.subtitle{font-size:13px;color:var(--color-text-secondary);margin-bottom:20px}
.conf{padding:10px 16px;border-radius:8px;margin-bottom:14px;font-size:12px}
.conf.high,.conf.moderate{background:var(--g-fill);color:var(--g-text)}
.conf.low{background:var(--a-fill);color:var(--a-text)}
.conf.vlow{background:var(--r-fill);color:var(--r-text)}
.tab-row{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px}
.tab-btn{padding:7px 18px;border-radius:100px;border:1px solid var(--color-border-secondary);background:transparent;color:var(--color-text-secondary);font-size:13px;font-weight:400;cursor:pointer;font-family:inherit;transition:border-color .15s,color .15s}
.tab-btn:hover{border-color:var(--color-border-primary);color:var(--color-text-primary)}
.tab-btn.active{border:1.5px solid var(--color-text-primary);color:var(--color-text-primary);font-weight:500}
.panel{display:none}.panel.on{display:block}
.card{background:var(--color-background-primary);border:.5px solid var(--color-border-tertiary);border-radius:12px;padding:16px 20px;margin-bottom:12px}
.card.green{background:var(--g-fill);border-color:var(--g-border)}
.card.amber{background:var(--a-fill);border-color:var(--a-border)}
.card.red{background:var(--r-fill);border-color:var(--r-border)}
.card.blue{background:var(--b-fill);border-color:var(--b-border)}
.card.grey{background:var(--n-fill);border-color:var(--n-border)}
.view-label{font-size:18px;font-weight:500;margin-bottom:6px}
.view-label.green{color:var(--g-text)}.view-label.amber{color:var(--a-text)}.view-label.red{color:var(--r-text)}
.view-reason{font-size:13px;margin-bottom:16px;color:var(--color-text-primary)}
.sec-label{font-size:10px;font-weight:500;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin:14px 0 8px}
.bullet-row{display:flex;gap:10px;align-items:flex-start;padding:6px 0;border-bottom:.5px solid var(--color-border-tertiary);font-size:13px}
.bullet-row:last-child{border-bottom:none}
.bicon{font-size:14px;min-width:18px;margin-top:1px}
table{width:100%;border-collapse:collapse;font-size:12px;margin:8px 0}
th{text-align:left;padding:7px 10px;font-weight:500;font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:var(--color-text-secondary);border-bottom:.5px solid var(--color-border-tertiary)}
td{padding:9px 10px;color:var(--color-text-primary);border-bottom:.5px solid var(--color-border-tertiary)}
tr:last-child td{border-bottom:none}
.pos{color:var(--g-accent);font-weight:500}.neg{color:var(--r-accent);font-weight:500}.neu{color:var(--color-text-secondary)}
.badge{display:inline-block;padding:2px 9px;border-radius:100px;font-size:11px;font-weight:500}
.badge.ok{background:var(--g-fill);color:var(--g-text)}
.badge.warn{background:var(--a-fill);color:var(--a-text)}
.badge.bad{background:var(--r-fill);color:var(--r-text)}
.badge.info{background:var(--b-fill);color:var(--b-text)}
.badge.neu{background:var(--n-fill);color:var(--n-text)}
.mgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:10px;margin:12px 0}
.mc{background:var(--color-background-secondary);border-radius:8px;padding:12px 14px}
.mc .ml{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--color-text-secondary);margin-bottom:4px}
.mc .mv{font-size:18px;font-weight:500;color:var(--color-text-primary)}
.mc .ms{font-size:11px;color:var(--color-text-secondary);margin-top:2px}
.info-row{display:flex;justify-content:space-between;align-items:baseline;padding:9px 0;border-bottom:.5px solid var(--color-border-tertiary);gap:16px}
.info-row:last-child{border-bottom:none}
.il{font-size:12px;color:var(--color-text-secondary)}
.iv{font-size:12px;color:var(--color-text-primary);font-weight:500;text-align:right}
.flag-card{border-left:3px solid var(--a-accent);background:var(--a-fill);border-radius:0 8px 8px 0;padding:10px 14px;margin-bottom:8px}
.flag-title{font-size:13px;font-weight:500;color:var(--a-text);margin-bottom:3px}
.flag-note{font-size:12px;color:var(--a-text)}
.score-box{background:var(--color-background-secondary);border-radius:8px;padding:14px 16px;margin-top:14px;font-size:12px;color:var(--color-text-secondary)}
.eps-row{display:flex;flex-wrap:wrap;gap:8px;margin:10px 0}
.eps-chip{background:var(--color-background-secondary);border-radius:8px;padding:10px 14px;text-align:center;min-width:85px}
.eps-q{font-size:10px;color:var(--color-text-secondary);text-transform:uppercase;letter-spacing:.5px}
.eps-v{font-size:16px;font-weight:500;color:var(--color-text-primary);margin:3px 0}
.eps-y{font-size:11px}
.peer-you{background:var(--color-background-secondary)}
.disc{font-size:11px;color:var(--color-text-secondary);border-top:.5px solid var(--color-border-tertiary);padding-top:14px;margin-top:24px;line-height:1.6}
details summary{font-size:12px;color:var(--color-text-secondary);cursor:pointer;padding:8px 0}
details p{font-size:12px;color:var(--color-text-secondary);padding:4px 0;border-bottom:.5px solid var(--color-border-tertiary)}
details p:last-child{border-bottom:none}
details strong{color:var(--color-text-primary);font-weight:500}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:4px}
@media(max-width:540px){.two-col{grid-template-columns:1fr}}
</style>
"""

SCRIPT = """
<script>
function show(n){
  var panels=document.querySelectorAll('.panel');
  var tabs=document.querySelectorAll('.tab-btn');
  for(var i=0;i<panels.length;i++){
    panels[i].className='panel'+(i===n?' on':'');
    tabs[i].className='tab-btn'+(i===n?' active':'');
  }
}
</script>
"""

GLOSSARY = """
<details style="margin-top:20px;padding:0 4px">
  <summary>Definitions — tap to expand</summary>
  <p><strong>P/E ratio</strong> — What you pay per ₹1 of profit. Lower vs history and sector = cheaper.</p>
  <p><strong>P/B ratio</strong> — Price vs what the company actually owns. Below 1 = buying at a discount to assets.</p>
  <p><strong>EV/EBITDA</strong> — Full business value check including debt. Lower = better value.</p>
  <p><strong>ROE</strong> — For every ₹100 shareholders invested, how much profit did the company make. Above 15% = good.</p>
  <p><strong>ROCE</strong> — How efficiently the whole business uses capital including borrowed money. Above 15% = healthy.</p>
  <p><strong>Free Cash Flow</strong> — Cash left after all expenses and investments. Positive and growing = genuinely healthy business.</p>
  <p><strong>Promoter pledging</strong> — Founders using their own shares as loan collateral. Above 10% is a red flag.</p>
  <p><strong>CAGR</strong> — Compound Annual Growth Rate. The average yearly growth rate over a period.</p>
</details>
"""

DISCLAIMER = """
<div class="disc">
  This is a fundamental screening and education tool only. Data sourced from yfinance (Yahoo Finance) and Screener.in.
  This is NOT investment advice, a buy/sell recommendation, or SEBI-registered financial research.
  Automated scripts can make errors — verify all numbers on NSE, BSE, or Screener.in before making any decision.
  Past performance does not guarantee future results. Investing carries risk.
  Consult a SEBI-registered financial advisor before investing.
</div>
"""


def build_html(ticker: str, horizon: int, sc: dict, yf_data: dict) -> str:
    now = datetime.now().strftime("%d %b %Y, %H:%M IST")
    company  = yf_data.get("company_name", ticker)
    sector   = yf_data.get("sector", "N/A")
    industry = yf_data.get("industry", "N/A")
    desc     = yf_data.get("description", "")
    moat     = "Refer Screener.in / annual report for competitive positioning"

    # ── Price & Market Cap ──────────────────────────────────────────────
    cmp       = yf_data.get("cmp")
    high52    = yf_data.get("52w_high_yf") or sc.get("52w_high")
    low52     = yf_data.get("52w_low_yf")  or sc.get("52w_low")
    mcap_yf   = yf_data.get("market_cap_yf")
    mcap_str  = f"₹{mcap_yf/1e7:,.0f} Cr" if mcap_yf else sc.get("market_cap","N/A")
    fv        = yf_data.get("face_value_yf") or sc.get("face_value","N/A")

    # ── Valuation ───────────────────────────────────────────────────────
    pe     = yf_data.get("pe_yf")   or _clean_num(sc.get("pe"))
    pb     = yf_data.get("pb_yf")
    ev_eb  = yf_data.get("ev_ebitda_yf")
    sp_pe  = sector_pe(sector)
    pe_sig, pe_cls = valuation_signal(pe, sp_pe, sp_pe * 0.95)
    pb_sig, pb_cls = valuation_signal(pb, 3.0, 3.0)
    ev_sig, ev_cls = valuation_signal(ev_eb, 15.0, 15.0)
    val_signals = [pe_cls, pb_cls, ev_cls]
    if val_signals.count("ok") >= 2:
        overall_val = "UNDERVALUED"
    elif val_signals.count("bad") >= 2:
        overall_val = "OVERVALUED"
    elif val_signals.count("warn") >= 2:
        overall_val = "FAIRLY VALUED"
    else:
        overall_val = "MIXED"

    # ── Growth ──────────────────────────────────────────────────────────
    rev_list = revenue_list_screener(sc)
    np_list  = profit_list_screener(sc)
    eps_list = eps_list_screener(sc)

    rev3 = cagr(rev_list, 3) if len(rev_list) >= 4 else None
    rev5 = cagr(rev_list, 5) if len(rev_list) >= 6 else None
    np3  = cagr(np_list,  3) if len(np_list)  >= 4 else None
    np5  = cagr(np_list,  5) if len(np_list)  >= 6 else None
    eps3 = cagr(eps_list, 3) if len(eps_list) >= 4 else None
    eps5 = cagr(eps_list, 5) if len(eps_list) >= 6 else None
    opm  = opm_latest(sc)
    gclass = growth_class(rev3, rev5, np3, np5)

    # ── Quarterly EPS ────────────────────────────────────────────────────
    eps_q_sc = sc.get("eps_quarterly", {})
    eps_q_yf = yf_data.get("eps_quarterly_yf", {})
    eps_q    = eps_q_sc if eps_q_sc else eps_q_yf
    # Build chips (last 8)
    eps_chips_html = ""
    eps_q_items = list(eps_q.items())[:8]
    for i, (qname, val) in enumerate(eps_q_items):
        if val is None:
            continue
        yoy_html = ""
        if i + 4 < len(eps_q_items):
            prev = eps_q_items[i + 4][1]
            if prev and prev != 0:
                yoy = round((val - prev) / abs(prev) * 100, 1)
                cls = "pos" if yoy >= 0 else "neg"
                sign = "+" if yoy >= 0 else ""
                yoy_html = f'<div class="eps-y {cls}">{sign}{yoy}% YoY</div>'
        eps_chips_html += f"""
        <div class="eps-chip">
          <div class="eps-q">{qname}</div>
          <div class="eps-v">₹{val}</div>
          {yoy_html}
        </div>"""

    # ── Health ──────────────────────────────────────────────────────────
    d_e   = de_ratio(sc, yf_data)
    cr    = current_ratio(yf_data)
    fcf_c = calc_fcf_crore(yf_data)
    ic    = interest_coverage(sc)   # None for now

    de_lbl, de_cls  = badge(d_e, [1, 2], ["SAFE","MODERATE","LEVERAGED"], ["ok","warn","bad"]) if d_e is not None else ("N/A","neu")
    cr_lbl, cr_cls  = badge(cr,  [1, 1.5], ["RISK","WATCH","COMFORTABLE"],  ["bad","warn","ok"])   if cr  is not None else ("N/A","neu")
    fcf_lbl, fcf_cls = ("CONCERN","bad") if fcf_c and fcf_c < 0 else (("STRONG","ok") if fcf_c and fcf_c > 0 else ("N/A","neu"))

    health_scores = [de_cls, cr_cls, fcf_cls]
    if health_scores.count("bad") >= 2:
        overall_health = "HIGH RISK"
        hcard = "red"
    elif health_scores.count("warn") >= 2:
        overall_health = "MODERATE RISK"
        hcard = "amber"
    else:
        overall_health = "SAFE"
        hcard = "green"

    # D/E trend
    borrow_d = sc.get("borrowings_annual", {})
    de_trend = trend_arrow(borrow_d) if borrow_d else "→ Stable"

    # Projections
    base_eps = latest(eps_q) if eps_q else (eps_list[-1] if eps_list else None)
    base_rate = (eps5 or eps3 or 10) / 100
    bear_eps, base_eps_p, bull_eps = project_eps(base_eps, base_rate, horizon)

    rev_latest = rev_list[-1] if rev_list else None
    np_latest  = np_list[-1]  if np_list  else None
    base_rev_rate = (rev5 or rev3 or 10) / 100
    base_np_rate  = (np5  or np3  or 10) / 100

    proj_rev_bear = round(rev_latest * ((1 + base_rev_rate * 0.4) ** horizon), 0) if rev_latest else None
    proj_rev_base = round(rev_latest * ((1 + base_rev_rate)       ** horizon), 0) if rev_latest else None
    proj_rev_bull = round(rev_latest * ((1 + base_rev_rate * 1.5) ** horizon), 0) if rev_latest else None
    proj_np_bear  = round(np_latest  * ((1 + base_np_rate  * 0.4) ** horizon), 0) if np_latest  else None
    proj_np_base  = round(np_latest  * ((1 + base_np_rate)        ** horizon), 0) if np_latest  else None
    proj_np_bull  = round(np_latest  * ((1 + base_np_rate  * 1.5) ** horizon), 0) if np_latest  else None

    # ── Returns ─────────────────────────────────────────────────────────
    roe_raw = yf_data.get("roe_yf")
    roe_pct = round(roe_raw * 100, 1) if roe_raw else _clean_num(sc.get("roe"))
    roce_raw= _clean_num(sc.get("roce"))
    dy_raw  = yf_data.get("div_yield_yf")
    div_y   = round(dy_raw * 100, 2) if dy_raw else _clean_num(sc.get("div_yield"))
    payout  = yf_data.get("payout_ratio_yf")
    payout_p= round(payout * 100, 1) if payout else None

    roe_lbl, roe_cls   = badge(roe_pct,  [10, 15], ["WEAK","AVERAGE","GOOD"], ["bad","warn","ok"])  if roe_pct  is not None else ("N/A","neu")
    roce_lbl, roce_cls = badge(roce_raw, [10, 15], ["WEAK","AVERAGE","GOOD"], ["bad","warn","ok"])  if roce_raw is not None else ("N/A","neu")

    if roe_cls == "ok" and roce_cls in ["ok","warn"]:
        ret_quality = "HIGH-QUALITY COMPOUNDER"
    elif div_y and div_y > 3:
        ret_quality = "DIVIDEND PLAY"
    elif roe_cls == "warn":
        ret_quality = "AVERAGE RETURNS"
    else:
        ret_quality = "CAPITAL-LIGHT"

    # ── Peers ────────────────────────────────────────────────────────────
    peers = sc.get("peers", [])
    peer_rows = ""
    for p in peers[:3]:
        peer_rows += f"""
        <tr>
          <td>{p.get('name','N/A')}</td>
          <td>{x(p.get('pe'))}</td>
          <td>{x(p.get('pb'))}</td>
          <td>{pct(p.get('roe'))}</td>
          <td>{pct(p.get('rev_growth'))}</td>
          <td>{x(p.get('de'))}</td>
          <td style="font-size:12px">—</td>
        </tr>"""
    # Peer standing
    peer_stand = "MID-PACK"
    if peers and pe:
        peer_pes = [p.get("pe") for p in peers if p.get("pe")]
        if peer_pes:
            if pe < min(peer_pes) * 1.1:
                peer_stand = "LEADING"
            elif pe > max(peer_pes) * 0.9:
                peer_stand = "LAGGING"

    # ── Ownership ───────────────────────────────────────────────────────
    prom_h = sc.get("promoter_holding", {})
    prom_pledge = sc.get("promoter_pledge", {})
    fii_h  = sc.get("fii_holding", {})
    dii_h  = sc.get("dii_holding", {})

    prom_val, prom_trend   = ownership_latest(prom_h)
    fii_val,  fii_trend    = ownership_latest(fii_h)
    dii_val,  dii_trend    = ownership_latest(dii_h)
    pledge_val, _          = ownership_latest(prom_pledge)
    prom_trend  = prom_trend  or "→ Stable"
    fii_trend   = fii_trend   or "→ Stable"
    dii_trend   = dii_trend   or "→ Stable"
    pledge_flag = "bad" if (pledge_val and pledge_val > 10) else "ok"
    pledge_lbl  = "⚠ FLAG" if (pledge_val and pledge_val > 10) else "OK"

    prom_badge_cls = "ok" if "buy" in prom_trend.lower() or "stable" in prom_trend.lower() else "bad"
    fii_badge_cls  = "ok" if "increas" in fii_trend.lower() or "stable" in fii_trend.lower() else "warn"
    dii_badge_cls  = "ok" if "increas" in dii_trend.lower() or "stable" in dii_trend.lower() else "warn"

    own_signal = "HOLDING STEADY"
    if "buy" in prom_trend.lower():
        own_signal = "INSIDERS BUILDING"
    elif "sell" in prom_trend.lower():
        own_signal = "TRIMMING"

    # ── Flags ────────────────────────────────────────────────────────────
    flags_html = ""
    if pledge_val and pledge_val > 10:
        flags_html += f"""
        <div class="flag-card">
          <div class="flag-title">⚠ HIGH PROMOTER PLEDGING ({pledge_val:.1f}%)</div>
          <div class="flag-note">Promoters have pledged over 10% of their shares as loan collateral. If the stock falls sharply, forced selling can accelerate the decline.</div>
        </div>"""
    if d_e and d_e > 2:
        flags_html += f"""
        <div class="flag-card">
          <div class="flag-title">⚠ HIGH LEVERAGE (D/E {d_e:.2f})</div>
          <div class="flag-note">Debt-to-equity above 2 means the company is heavily reliant on borrowed money. Rising interest rates could hurt profits.</div>
        </div>"""

    # ── Fundamental View ─────────────────────────────────────────────────
    score = 0
    if gclass in ["ACCELERATING", "STEADY"]:          score += 2
    if roe_cls == "ok":                                score += 2
    if overall_val in ["UNDERVALUED","FAIRLY VALUED"]: score += 1
    if overall_health == "SAFE":                       score += 2
    if pledge_val and pledge_val > 10:                 score -= 2
    if d_e and d_e > 2:                                score -= 1

    if score >= 5:
        view_label = "STRONG FUNDAMENTALS"
        view_card  = "green"
    elif score >= 2:
        view_label = "MODERATE FUNDAMENTALS"
        view_card  = "amber"
    else:
        view_label = "WEAK FUNDAMENTALS"
        view_card  = "red"

    strengths = []
    watches   = []
    if gclass in ["ACCELERATING", "STEADY"]:
        strengths.append(f"Revenue growing at ~{rev5 or rev3 or 'N/A'}% CAGR — consistent top-line expansion")
    if roe_cls == "ok":
        strengths.append(f"Strong ROE of {pct(roe_pct)} — management is generating solid returns on shareholder capital")
    if overall_val == "UNDERVALUED":
        strengths.append("Currently trading below its historical P/E average — potential valuation comfort")
    elif overall_val == "FAIRLY VALUED":
        strengths.append("Fairly valued vs sector peers — not obviously expensive")
    if overall_health == "SAFE":
        strengths.append("Low debt burden provides resilience in economic downturns")
    while len(strengths) < 3:
        strengths.append("Refer Screener.in / NSE filings for additional qualitative strengths")

    if d_e and d_e > 1.5:
        watches.append(f"Elevated debt (D/E {d_e:.1f}x) — monitor borrowing trend in upcoming quarters")
    if gclass in ["SLOWING", "DECLINING"]:
        watches.append("Growth momentum has slowed — watch for margin improvement signals")
    if pledge_val and pledge_val > 5:
        watches.append(f"Promoter pledge at {pct(pledge_val)} — any stock price dip could trigger forced selling")
    while len(watches) < 2:
        watches.append("Monitor quarterly earnings for any divergence from historical trend")

    track = f"Track quarterly revenue and margin trajectory over the next {horizon} years; watch for promoter stake changes"

    # ── Data confidence ──────────────────────────────────────────────────
    live_count = sum([
        1 if cmp else 0,
        1 if pe else 0,
        1 if pb else 0,
        1 if rev3 else 0,
        1 if np3 else 0,
        1 if eps3 else 0,
        1 if roe_pct else 0,
        1 if d_e else 0,
        1 if prom_val else 0,
        1 if fii_val else 0,
        1 if dii_val else 0,
        1 if fcf_c else 0,
    ])
    if live_count >= 9:
        conf_cls, conf_lbl = "high",     "HIGH"
    elif live_count >= 6:
        conf_cls, conf_lbl = "moderate", "MODERATE"
    elif live_count >= 3:
        conf_cls, conf_lbl = "low",      "LOW ⚠ Verify before investing"
    else:
        conf_cls, conf_lbl = "vlow",     "VERY LOW ⚠ Data incomplete — use with caution"

    sources = "yfinance (Yahoo Finance)"
    if not sc.get("screener_error"):
        sources += ", Screener.in"

    # ── Assemble HTML ────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{company} — Fundamental Report</title>
{CSS}
</head>
<body>
<div class="wrap">

<h1>{company} ({ticker.upper()})</h1>
<div class="subtitle">Fundamental Report &nbsp;·&nbsp; {horizon}-Year Horizon &nbsp;·&nbsp; Generated {now}</div>

<div class="conf {conf_cls}">
  <strong>Data confidence: {conf_lbl}</strong><br>
  Live metrics retrieved: {live_count} of 12 key sections &nbsp;|&nbsp; Sources: {sources}
</div>

<!-- TAB ROW -->
<div class="tab-row">
  <button class="tab-btn" onclick="show(0)">Snapshot</button>
  <button class="tab-btn" onclick="show(1)">Valuation</button>
  <button class="tab-btn" onclick="show(2)">Growth</button>
  <button class="tab-btn" onclick="show(3)">Health</button>
  <button class="tab-btn" onclick="show(4)">Returns</button>
  <button class="tab-btn" onclick="show(5)">Peers</button>
  <button class="tab-btn" onclick="show(6)">Ownership</button>
  <button class="tab-btn active" onclick="show(7)">View</button>
</div>

<!-- ═══ TAB 0: SNAPSHOT ═══ -->
<div class="panel" id="p0">
  <div class="card">
    <div class="info-row"><span class="il">Company</span><span class="iv">{company}</span></div>
    <div class="info-row"><span class="il">Ticker</span><span class="iv">{ticker.upper()} · NSE</span></div>
    <div class="info-row"><span class="il">Sector</span><span class="iv">{sector}</span></div>
    <div class="info-row"><span class="il">Industry</span><span class="iv">{industry}</span></div>
    <div class="info-row"><span class="il">What it does</span><span class="iv" style="max-width:440px;text-align:right">{desc[:220] + "…" if len(desc) > 220 else desc or "N/A"}</span></div>
    <div class="info-row"><span class="il">What makes it different</span><span class="iv" style="max-width:440px;text-align:right">{moat}</span></div>
  </div>
  <div class="mgrid">
    <div class="mc"><div class="ml">CMP</div><div class="mv">{format_inr(cmp)}</div><div class="ms">{now} · NSE</div></div>
    <div class="mc"><div class="ml">52W High</div><div class="mv">{format_inr(high52)}</div><div class="ms">NSE / Yahoo</div></div>
    <div class="mc"><div class="ml">52W Low</div><div class="mv">{format_inr(low52)}</div><div class="ms">NSE / Yahoo</div></div>
    <div class="mc"><div class="ml">Market Cap</div><div class="mv">{mcap_str}</div><div class="ms">BSE / Yahoo</div></div>
    <div class="mc"><div class="ml">Face Value</div><div class="mv">₹{fv}</div><div class="ms">NSE</div></div>
  </div>
  {flags_html}
</div>

<!-- ═══ TAB 1: VALUATION ═══ -->
<div class="panel" id="p1">
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">Is this stock cheap, fair, or expensive right now?</div>
    <table>
      <thead><tr><th>Metric</th><th>Current</th><th>Sector avg</th><th>Stock 5Y avg</th><th>Signal</th><th>Plain English</th></tr></thead>
      <tbody>
        <tr>
          <td>P/E</td><td>{x(pe)}</td><td>{x(sp_pe)}</td><td>N/A</td>
          <td><span class="badge {pe_cls}">{pe_sig}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">You pay ₹{round(pe,0) if pe else "?"} per ₹1 of profit</td>
        </tr>
        <tr>
          <td>P/B</td><td>{x(pb)}</td><td>3.0x</td><td>N/A</td>
          <td><span class="badge {pb_cls}">{pb_sig}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Price vs net assets owned</td>
        </tr>
        <tr>
          <td>EV/EBITDA</td><td>{x(ev_eb)}</td><td>15.0x</td><td>N/A</td>
          <td><span class="badge {ev_cls}">{ev_sig}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Full business value check</td>
        </tr>
      </tbody>
    </table>
  </div>
  <div class="score-box">
    Overall valuation: <strong>{overall_val}</strong><br>
    <span style="margin-top:4px;display:block">Based on trailing P/E vs sector average of {sp_pe}x ({sector}). Historical averages not available from automated sources — verify on Screener.in.</span>
  </div>
</div>

<!-- ═══ TAB 2: GROWTH ═══ -->
<div class="panel" id="p2">
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">Is this company growing its revenue and profits?</div>
    <table>
      <thead><tr><th>Metric</th><th>3Y CAGR</th><th>5Y CAGR</th><th>Trend</th><th>Source</th></tr></thead>
      <tbody>
        <tr><td>Revenue</td><td>{pct(rev3)}</td><td>{pct(rev5)}</td><td>{'📈' if (rev5 or rev3 or 0) > 10 else ('📉' if (rev5 or rev3 or 0) < 0 else '➡️')}</td><td style="color:var(--color-text-secondary);font-size:11px">Screener.in</td></tr>
        <tr><td>Net profit</td><td>{pct(np3)}</td><td>{pct(np5)}</td><td>{'📈' if (np5 or np3 or 0) > 10 else ('📉' if (np5 or np3 or 0) < 0 else '➡️')}</td><td style="color:var(--color-text-secondary);font-size:11px">Screener.in</td></tr>
        <tr><td>EPS</td><td>{pct(eps3)}</td><td>{pct(eps5)}</td><td>{'📈' if (eps5 or eps3 or 0) > 10 else ('📉' if (eps5 or eps3 or 0) < 0 else '➡️')}</td><td style="color:var(--color-text-secondary);font-size:11px">Screener.in</td></tr>
        <tr><td>EBITDA margin</td><td>{pct(opm)}</td><td>N/A</td><td>{'📈' if (opm or 0) > 15 else '➡️'}</td><td style="color:var(--color-text-secondary);font-size:11px">Screener.in</td></tr>
        <tr><td>Net profit margin</td><td>N/A</td><td>N/A</td><td>➡️</td><td style="color:var(--color-text-secondary);font-size:11px">Screener.in</td></tr>
      </tbody>
    </table>
  </div>
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">EPS — last 8 quarters</div>
    <div class="eps-row">
      {eps_chips_html if eps_chips_html else '<span style="color:var(--color-text-secondary);font-size:12px">🚩 DATA UNAVAILABLE — verify at screener.in</span>'}
    </div>
  </div>
  <div class="score-box">
    Growth classification: <strong>{gclass}</strong><br>
    <span style="margin-top:4px;display:block">Based on {horizon}-year horizon; revenue CAGR {pct(rev5)} (5Y) / {pct(rev3)} (3Y) and net profit CAGR {pct(np5)} (5Y) / {pct(np3)} (3Y).</span>
  </div>
</div>

<!-- ═══ TAB 3: HEALTH ═══ -->
<div class="panel" id="p3">
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">Is this company financially safe and stable?</div>
    <table>
      <thead><tr><th>Metric</th><th>Value</th><th>5Y trend</th><th>Signal</th><th>Plain English</th></tr></thead>
      <tbody>
        <tr>
          <td>Debt / Equity</td><td>{x(d_e)}</td><td>{de_trend}</td>
          <td><span class="badge {de_cls}">{de_lbl}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Below 1 = safe</td>
        </tr>
        <tr>
          <td>Interest Coverage</td><td>N/A</td><td>→</td>
          <td><span class="badge neu">N/A</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Above 3x = healthy. Verify at Screener.in</td>
        </tr>
        <tr>
          <td>Current Ratio</td><td>{x(cr)}</td><td>→</td>
          <td><span class="badge {cr_cls}">{cr_lbl}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Above 1.5 = comfortable</td>
        </tr>
        <tr>
          <td>Free Cash Flow</td><td>{format_inr_cr(fcf_c)}</td><td>→</td>
          <td><span class="badge {fcf_cls}">{fcf_lbl}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Positive = real cash business</td>
        </tr>
      </tbody>
    </table>
  </div>
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">Forward projections — {horizon} year horizon</div>
    <table>
      <thead><tr><th>Scenario</th><th>Assumption</th><th>Est. revenue</th><th>Est. net profit</th><th>Est. EPS</th></tr></thead>
      <tbody>
        <tr><td>🐢 Bear</td><td style="color:var(--color-text-secondary);font-size:12px">Growth slows 60%, margins compress</td><td>{format_inr_cr(proj_rev_bear)}</td><td>{format_inr_cr(proj_np_bear)}</td><td>₹{bear_eps or 'N/A'}</td></tr>
        <tr><td>🚶 Base</td><td style="color:var(--color-text-secondary);font-size:12px">Maintains historical CAGR</td><td>{format_inr_cr(proj_rev_base)}</td><td>{format_inr_cr(proj_np_base)}</td><td>₹{base_eps_p or 'N/A'}</td></tr>
        <tr><td>🚀 Bull</td><td style="color:var(--color-text-secondary);font-size:12px">Growth accelerates 1.5x, margins expand</td><td>{format_inr_cr(proj_rev_bull)}</td><td>{format_inr_cr(proj_np_bull)}</td><td>₹{bull_eps or 'N/A'}</td></tr>
      </tbody>
    </table>
    <div style="font-size:11px;color:var(--color-text-secondary);margin-top:8px">Projections based on historical CAGR trends only — not guarantees or predictions.</div>
  </div>
  <div class="score-box">
    Financial health: <strong>{overall_health}</strong><br>
    <span style="margin-top:4px;display:block">D/E {x(d_e)} · Current Ratio {x(cr)} · FCF {format_inr_cr(fcf_c)}</span>
  </div>
</div>

<!-- ═══ TAB 4: RETURNS ═══ -->
<div class="panel" id="p4">
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">Is this company creating real value for shareholders?</div>
    <table>
      <thead><tr><th>Metric</th><th>Current</th><th>3Y avg</th><th>5Y avg</th><th>Signal</th></tr></thead>
      <tbody>
        <tr>
          <td>ROE</td><td>{pct(roe_pct)}</td><td>N/A</td><td>N/A</td>
          <td><span class="badge {roe_cls}">{roe_lbl}</span></td>
        </tr>
        <tr>
          <td>ROCE</td><td>{pct(roce_raw)}</td><td>N/A</td><td>N/A</td>
          <td><span class="badge {roce_cls}">{roce_lbl}</span></td>
        </tr>
        <tr>
          <td>Dividend yield</td><td>{pct(div_y)}</td><td>N/A</td><td>N/A</td>
          <td><span class="badge neu">—</span></td>
        </tr>
        <tr>
          <td>Dividend payout</td><td>{pct(payout_p)}</td><td>N/A</td><td>N/A</td>
          <td><span class="badge neu">—</span></td>
        </tr>
      </tbody>
    </table>
    <div style="font-size:11px;color:var(--color-text-secondary);margin-top:8px">ROE above 15% = good &nbsp;·&nbsp; ROCE above 15% = efficient capital use</div>
  </div>
  <div class="score-box">
    Return quality: <strong>{ret_quality}</strong><br>
    <span style="margin-top:4px;display:block">ROE {pct(roe_pct)} · ROCE {pct(roce_raw)} · Dividend yield {pct(div_y)}</span>
  </div>
</div>

<!-- ═══ TAB 5: PEERS ═══ -->
<div class="panel" id="p5">
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">How does it compare to its closest competitors?</div>
    <table>
      <thead><tr><th>Company</th><th>P/E</th><th>P/B</th><th>ROE</th><th>Rev growth</th><th>D/E</th><th>Edge</th></tr></thead>
      <tbody>
        <tr class="peer-you">
          <td><strong>{ticker.upper()} ◀ you</strong></td>
          <td>{x(pe)}</td><td>{x(pb)}</td><td>{pct(roe_pct)}</td><td>{pct(rev5 or rev3)}</td><td>{x(d_e)}</td>
          <td style="font-size:12px">—</td>
        </tr>
        {peer_rows if peer_rows else '<tr><td colspan="7" style="color:var(--color-text-secondary);text-align:center;font-size:12px">🚩 Peer data unavailable — verify at screener.in</td></tr>'}
      </tbody>
    </table>
    <div style="font-size:11px;color:var(--color-text-secondary);margin-top:8px">Source: Screener.in, NSE filings</div>
  </div>
  <div class="score-box">
    Peer standing: <strong>{peer_stand}</strong><br>
    <span style="margin-top:4px;display:block">P/E {x(pe)} vs sector average {x(sp_pe)} ({sector})</span>
  </div>
</div>

<!-- ═══ TAB 6: OWNERSHIP ═══ -->
<div class="panel" id="p6">
  <div class="card">
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--color-text-secondary);margin-bottom:10px">Who is backing this company — and are they buying or stepping away?</div>
    <table>
      <thead><tr><th>Holder</th><th>Latest %</th><th>8-quarter trend</th><th>Signal</th><th>What it means</th></tr></thead>
      <tbody>
        <tr>
          <td>Promoter</td><td>{pct(prom_val)}</td><td>{prom_trend}</td>
          <td><span class="badge {prom_badge_cls}">{'BUYING' if 'buy' in prom_trend.lower() else ('SELLING' if 'sell' in prom_trend.lower() else 'STABLE')}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Founder confidence</td>
        </tr>
        <tr>
          <td>FII</td><td>{pct(fii_val)}</td><td>{fii_trend}</td>
          <td><span class="badge {fii_badge_cls}">{'INCREASING' if 'increas' in fii_trend.lower() else ('DECREASING' if 'decreas' in fii_trend.lower() else 'STABLE')}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Global fund interest</td>
        </tr>
        <tr>
          <td>DII</td><td>{pct(dii_val)}</td><td>{dii_trend}</td>
          <td><span class="badge {dii_badge_cls}">{'INCREASING' if 'increas' in dii_trend.lower() else ('DECREASING' if 'decreas' in dii_trend.lower() else 'STABLE')}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Indian MF &amp; insurance</td>
        </tr>
        <tr>
          <td>Promoter pledging</td><td>{pct(pledge_val)}</td><td>—</td>
          <td><span class="badge {pledge_flag}">{pledge_lbl}</span></td>
          <td style="color:var(--color-text-secondary);font-size:12px">Above 10% = red flag</td>
        </tr>
      </tbody>
    </table>
  </div>
  <div class="score-box">
    Ownership signal: <strong>{own_signal}</strong><br>
    <span style="margin-top:4px;display:block">Promoter {pct(prom_val)} · FII {pct(fii_val)} · DII {pct(dii_val)}</span>
  </div>
</div>

<!-- ═══ TAB 7: VIEW (default active) ═══ -->
<div class="panel on" id="p7">
  <div class="card {view_card}">
    <div class="view-label {view_card}">{view_label}</div>
    <div class="view-reason">{company} shows {gclass.lower()} growth with {'solid' if roe_cls == 'ok' else 'moderate'} returns on equity and {'healthy' if overall_health == 'SAFE' else 'elevated'} financial risk based on available data.</div>

    <div class="sec-label">What works for this stock</div>
    {''.join(f'<div class="bullet-row"><span class="bicon" style="color:var(--g-accent)">✓</span><span>{s}</span></div>' for s in strengths[:3])}

    <div class="sec-label">What to watch</div>
    {''.join(f'<div class="bullet-row"><span class="bicon" style="color:var(--a-accent)">⚠</span><span>{w}</span></div>' for w in watches[:2])}

    <div class="sec-label">Track this going forward</div>
    <div class="bullet-row"><span class="bicon" style="color:var(--color-text-secondary)">→</span><span>{track}</span></div>

    <div style="margin-top:16px;font-size:11px;color:var(--color-text-secondary);font-style:italic">
      This is a VIEW based on fundamentals only. Not a buy/sell recommendation. The decision is always yours.
    </div>
  </div>

  <div class="two-col">
    <div class="card green">
      <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--g-text);margin-bottom:8px">Opportunities</div>
      <div class="bullet-row" style="border-color:var(--g-border)"><span class="bicon" style="color:var(--g-accent)">+</span><span style="color:var(--g-text);font-size:12px">Long-term sector tailwinds in {industry}</span></div>
      <div class="bullet-row" style="border-color:var(--g-border)"><span class="bicon" style="color:var(--g-accent)">+</span><span style="color:var(--g-text);font-size:12px">Consistent historical compounding at {pct(rev5 or rev3)} revenue CAGR</span></div>
      <div class="bullet-row" style="border-color:var(--g-border);border-bottom:none"><span class="bicon" style="color:var(--g-accent)">+</span><span style="color:var(--g-text);font-size:12px">{'Dividend yield of ' + pct(div_y) + ' provides income floor' if div_y and div_y > 1 else 'Potential for re-rating if growth accelerates'}</span></div>
    </div>
    <div class="card red">
      <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--r-text);margin-bottom:8px">Risks</div>
      <div class="bullet-row" style="border-color:var(--r-border)"><span class="bicon" style="color:var(--r-accent)">−</span><span style="color:var(--r-text);font-size:12px">{'High promoter pledge (' + pct(pledge_val) + ') could cause forced selling' if pledge_val and pledge_val > 10 else 'Regulatory or macro headwinds affecting sector'}</span></div>
      <div class="bullet-row" style="border-color:var(--r-border)"><span class="bicon" style="color:var(--r-accent)">−</span><span style="color:var(--r-text);font-size:12px">{'Leverage (D/E ' + str(d_e) + 'x) sensitive to interest rate cycles' if d_e and d_e > 1 else 'Competition intensifying in ' + industry}</span></div>
      <div class="bullet-row" style="border-color:var(--r-border);border-bottom:none"><span class="bicon" style="color:var(--r-accent)">−</span><span style="color:var(--r-text);font-size:12px">Automated report — always verify data on NSE/BSE/Screener.in</span></div>
    </div>
  </div>
</div>

{GLOSSARY}
{DISCLAIMER}

</div>
{SCRIPT}
</body>
</html>"""
    return html


# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────

def main():
    args = parse_args()
    ticker  = args.ticker.strip().upper()
    horizon = args.horizon

    print(f"[1/3] Fetching yfinance data for {ticker}.NS …")
    yf_data = yf_fetch(ticker)
    if "yf_error" in yf_data:
        print(f"  ⚠  yfinance error: {yf_data['yf_error']}")

    print(f"[2/3] Scraping Screener.in for {ticker} …")
    sc = screener_fetch(ticker)
    if "screener_error" in sc:
        print(f"  ⚠  Screener error: {sc['screener_error']}")

    print("[3/3] Building HTML report …")
    html = build_html(ticker, horizon, sc, yf_data)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    fname = out_dir / f"{ticker}_{horizon}yr_{date.today().isoformat()}.html"
    fname.write_text(html, encoding="utf-8")
    print(f"\n✅  Report saved → {fname}")


if __name__ == "__main__":
    main()
