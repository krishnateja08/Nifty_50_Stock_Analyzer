"""
Indian Stock Fundamental Analyser — Nifty 50 Edition
Data  : yfinance (Yahoo Finance) + Screener.in scraping
Output: Single combined HTML — dropdown/search to switch stocks
Usage : python analyser.py [--out reports]
"""

import argparse
import sys
import time
import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date
from pathlib import Path

import yfinance as yf
import requests
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────────────────
# NIFTY 50 TICKERS
# ─────────────────────────────────────────────────────────────────

NIFTY50 = [
    "ADANIENT", "ADANIPORTS", "APOLLOHOSP", "ASIANPAINT", "AXISBANK",
    "BAJAJ-AUTO", "BAJFINANCE", "BAJAJFINSV", "BPCL", "BHARTIARTL",
    "BRITANNIA", "CIPLA", "COALINDIA", "DIVISLAB", "DRREDDY",
    "EICHERMOT", "GRASIM", "HCLTECH", "HDFCBANK", "HDFCLIFE",
    "HEROMOTOCO", "HINDALCO", "HINDUNILVR", "ICICIBANK", "ITC",
    "INDUSINDBK", "INFY", "JSWSTEEL", "KOTAKBANK", "LTIM",
    "LT", "M&M", "MARUTI", "NTPC", "NESTLEIND",
    "ONGC", "POWERGRID", "RELIANCE", "SBILIFE", "SHRIRAMFIN",
    "SBIN", "SUNPHARMA", "TCS", "TATACONSUM", "TATAMOTORS",
    "TATASTEEL", "TECHM", "TITAN", "ULTRACEMCO", "WIPRO",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# ─────────────────────────────────────────────────────────────────
# SCREENER SCRAPER
# ─────────────────────────────────────────────────────────────────

def _text(el):
    return el.get_text(strip=True) if el else None

def _clean_num(s):
    if not s:
        return None
    s = str(s).replace(",", "").replace("%", "").replace("\u20b9", "").strip()
    try:
        return float(s)
    except ValueError:
        return None

def screener_fetch(ticker: str) -> dict:
    data = {}
    sc_ticker = ticker.replace("&", "%26")
    for suffix in ["/consolidated/", "/"]:
        url = f"https://www.screener.in/company/{sc_ticker}{suffix}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=25)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "lxml")
                data["url"] = url
                data.update(_screener_ratios(soup))
                data.update(_screener_financials(soup))
                data.update(_screener_shareholding(soup))
                data.update(_screener_peers(soup, ticker))
                data.update(_screener_quarters(soup))
                return data
        except Exception as e:
            data["screener_error"] = str(e)
    return data

def _screener_ratios(soup) -> dict:
    out = {}
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
                out["52w_high"] = parts[0].strip().replace("\u20b9", "")
                out["52w_low"]  = parts[1].strip().replace("\u20b9", "")
        elif "p/e" in name:
            out["pe"] = val
        elif "book value" in name:
            out["book_value"] = val
        elif "dividend yield" in name:
            out["div_yield"] = val
        elif name.strip() == "roce":
            out["roce"] = val
        elif name.strip() == "roe":
            out["roe"] = val
        elif "face value" in name:
            out["face_value"] = val
    return out

def _screener_financials(soup) -> dict:
    out = {}
    try:
        section = soup.find("section", {"id": "profit-loss"})
        if not section:
            return out
        years = [_text(th) for th in section.select("table thead th")[1:] if _text(th)]
        for row in section.select("table tbody tr"):
            cells = row.select("td")
            if not cells:
                continue
            label = _text(cells[0]).lower() if cells[0] else ""
            vals  = [_clean_num(_text(c)) for c in cells[1:]]
            if "sales" in label or ("revenue" in label and "growth" not in label):
                out["revenue_annual"] = dict(zip(years, vals))
            elif "net profit" in label and "%" not in label:
                out["netprofit_annual"] = dict(zip(years, vals))
            elif label.strip() == "eps":
                out["eps_annual"] = dict(zip(years, vals))
            elif "opm" in label:
                out["opm_annual"] = dict(zip(years, vals))
        bs = soup.find("section", {"id": "balance-sheet"})
        if bs:
            bs_years = [_text(th) for th in bs.select("table thead th")[1:] if _text(th)]
            for row in bs.select("table tbody tr"):
                cells = row.select("td")
                if not cells:
                    continue
                label = _text(cells[0]).lower() if cells[0] else ""
                vals  = [_clean_num(_text(c)) for c in cells[1:]]
                if "borrowings" in label:
                    out["borrowings_annual"] = dict(zip(bs_years, vals))
                elif "equity" in label and "share capital" not in label:
                    out["equity_annual"] = dict(zip(bs_years, vals))
    except Exception as e:
        out["financials_error"] = str(e)
    return out

def _screener_shareholding(soup) -> dict:
    out = {}
    try:
        section = soup.find("section", {"id": "shareholding"})
        if not section:
            return out
        quarters = [_text(th) for th in section.select("table thead th")[1:] if _text(th)]
        for row in section.select("table tbody tr"):
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
        for row in section.select("table tbody tr"):
            cells = row.select("td")
            if len(cells) < 6:
                continue
            name = _text(cells[0])
            if not name or self_ticker.upper().replace("-", "") in name.upper().replace("-", ""):
                continue
            out["peers"].append({
                "name":       name,
                "pe":         _clean_num(_text(cells[2])),
                "pb":         _clean_num(_text(cells[5])) if len(cells) > 5 else None,
                "roe":        _clean_num(_text(cells[7])) if len(cells) > 7 else None,
                "rev_growth": _clean_num(_text(cells[6])) if len(cells) > 6 else None,
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
        quarters = [_text(th) for th in section.select("table thead th")[1:] if _text(th)]
        for row in section.select("table tbody tr"):
            cells = row.select("td")
            if not cells:
                continue
            label = _text(cells[0]).lower() if cells[0] else ""
            vals  = [_clean_num(_text(c)) for c in cells[1:]]
            if label.strip() == "eps":
                out["eps_quarterly"] = dict(zip(quarters, vals))
            elif "net profit" in label and "%" not in label:
                out["netprofit_quarterly"] = dict(zip(quarters, vals))
            elif "sales" in label or ("revenue" in label and "%" not in label):
                out["revenue_quarterly"] = dict(zip(quarters, vals))
    except Exception as e:
        out["quarters_error"] = str(e)
    return out

# ─────────────────────────────────────────────────────────────────
# YFINANCE FETCH
# ─────────────────────────────────────────────────────────────────

def yf_fetch(ticker: str) -> dict:
    out = {}
    nse = ticker + ".NS"
    try:
        tk   = yf.Ticker(nse)
        info = tk.info or {}
        out["company_name"]    = info.get("longName") or info.get("shortName", ticker)
        out["sector"]          = info.get("sector", "N/A")
        out["industry"]        = info.get("industry", "N/A")
        out["description"]     = info.get("longBusinessSummary", "")
        out["cmp"]             = info.get("currentPrice") or info.get("regularMarketPrice")
        out["52w_high_yf"]     = info.get("fiftyTwoWeekHigh")
        out["52w_low_yf"]      = info.get("fiftyTwoWeekLow")
        out["market_cap_yf"]   = info.get("marketCap")
        out["pe_yf"]           = info.get("trailingPE")
        out["pb_yf"]           = info.get("priceToBook")
        out["ev_ebitda_yf"]    = info.get("enterpriseToEbitda")
        out["roe_yf"]          = info.get("returnOnEquity")
        out["div_yield_yf"]    = info.get("dividendYield")
        out["payout_ratio_yf"] = info.get("payoutRatio")
        out["face_value_yf"]   = info.get("faceValue")
        try:
            fin = tk.financials
            if fin is not None and not fin.empty:
                if "Total Revenue" in fin.index:
                    out["yf_annual_revenue"] = fin.loc["Total Revenue"].to_dict()
                if "Net Income" in fin.index:
                    out["yf_annual_netprofit"] = fin.loc["Net Income"].to_dict()
                if "EBITDA" in fin.index:
                    out["yf_annual_ebitda"] = fin.loc["EBITDA"].to_dict()
        except Exception:
            pass
        try:
            cf = tk.cashflow
            if cf is not None and not cf.empty:
                fcf_key = next((k for k in cf.index if "Free Cash Flow" in str(k)), None)
                if fcf_key:
                    out["yf_fcf"] = cf.loc[fcf_key].to_dict()
                else:
                    ocf = cf.loc["Operating Cash Flow"].to_dict() if "Operating Cash Flow" in cf.index else {}
                    cap = cf.loc["Capital Expenditure"].to_dict() if "Capital Expenditure" in cf.index else {}
                    if ocf and cap:
                        out["yf_fcf"] = {k: (ocf.get(k, 0) or 0) + (cap.get(k, 0) or 0) for k in ocf}
        except Exception:
            pass
        try:
            bs = tk.balance_sheet
            if bs is not None and not bs.empty:
                if "Total Debt" in bs.index:
                    out["yf_total_debt"] = bs.loc["Total Debt"].to_dict()
                if "Stockholders Equity" in bs.index:
                    out["yf_equity"] = bs.loc["Stockholders Equity"].to_dict()
                if "Current Assets" in bs.index:
                    out["yf_current_assets"] = bs.loc["Current Assets"].to_dict()
                if "Current Liabilities" in bs.index:
                    out["yf_current_liab"] = bs.loc["Current Liabilities"].to_dict()
        except Exception:
            pass
        try:
            q_fin = tk.quarterly_financials
            if q_fin is not None and not q_fin.empty and "Net Income" in q_fin.index:
                ni     = q_fin.loc["Net Income"]
                shares = info.get("sharesOutstanding") or 1
                out["eps_quarterly_yf"] = {
                    str(col)[:10]: round(val / shares, 2)
                    for col, val in ni.items() if val is not None
                }
        except Exception:
            pass
    except Exception as e:
        out["yf_error"] = str(e)
    return out

# ─────────────────────────────────────────────────────────────────
# CALCULATIONS
# ─────────────────────────────────────────────────────────────────

def latest(d: dict):
    if not d:
        return None
    return list(d.values())[0]

def cagr(values: list, years: int):
    clean = [v for v in values if v is not None and v != 0]
    if len(clean) < 2:
        return None
    try:
        return round(((clean[-1] / clean[0]) ** (1 / years) - 1) * 100, 1)
    except Exception:
        return None

def trend_arrow(d: dict) -> str:
    vals = [v for v in list(d.values()) if v is not None]
    if len(vals) < 2:
        return "\u2192 Stable"
    diff = vals[0] - vals[-1]
    if diff > 1:
        return "\u2191 Rising"
    if diff < -1:
        return "\u2193 Falling"
    return "\u2192 Stable"

def de_ratio(sc, yfd):
    b = latest(sc.get("borrowings_annual", {}))
    e = latest(sc.get("equity_annual", {}))
    if b is not None and e and e != 0:
        return round(b / e, 2)
    d = latest(yfd.get("yf_total_debt", {}))
    q = latest(yfd.get("yf_equity", {}))
    if d is not None and q and q != 0:
        return round(abs(d / q), 2)
    return None

def current_ratio(yfd):
    a = latest(yfd.get("yf_current_assets", {}))
    l = latest(yfd.get("yf_current_liab", {}))
    if a and l and l != 0:
        return round(a / l, 2)
    return None

def fcf_crore(yfd):
    val = latest(yfd.get("yf_fcf", {}))
    if val:
        return round(val / 1e7, 0)
    return None

def rev_list(sc):
    return [v for v in reversed(list(sc.get("revenue_annual", {}).values())) if v is not None]

def np_list(sc):
    return [v for v in reversed(list(sc.get("netprofit_annual", {}).values())) if v is not None]

def eps_list(sc):
    return [v for v in reversed(list(sc.get("eps_annual", {}).values())) if v is not None]

def ownership_latest(d: dict):
    if not d:
        return None, "\u2192 Stable"
    items = list(d.items())
    lv = items[0][1]
    ov = items[-1][1] if len(items) > 1 else lv
    trend = "\u2192 Stable"
    if lv is not None and ov is not None:
        diff = lv - ov
        if diff > 1:
            trend = "\u2191 Increasing"
        elif diff < -1:
            trend = "\u2193 Decreasing"
    return lv, trend

SECTOR_PE = {
    "Technology": 28, "Information Technology": 28,
    "Financial Services": 18, "Banking": 16,
    "Consumer Defensive": 35, "Consumer Cyclical": 40,
    "Healthcare": 30, "Energy": 12, "Utilities": 15,
    "Industrials": 25, "Basic Materials": 14,
    "Communication Services": 22, "Real Estate": 30,
}

def sector_pe(sector: str) -> float:
    for k, v in SECTOR_PE.items():
        if k.lower() in (sector or "").lower():
            return v
    return 22.0

def fmt_inr_cr(val) -> str:
    if val is None:
        return "N/A"
    try:
        v = float(val)
        if abs(v) >= 1_00_000:
            return f"\u20b9{v/1_00_000:.1f}L Cr"
        if abs(v) >= 1_000:
            return f"\u20b9{v:,.0f} Cr"
        return f"\u20b9{v:.0f} Cr"
    except Exception:
        return str(val)

def fmt_inr(val) -> str:
    if val is None:
        return "N/A"
    try:
        return f"\u20b9{float(val):,.2f}"
    except Exception:
        return str(val)

def pct(val) -> str:
    if val is None:
        return "N/A"
    try:
        return f"{float(val):.1f}%"
    except Exception:
        return str(val)

def xfmt(val) -> str:
    if val is None:
        return "N/A"
    try:
        return f"{float(val):.1f}x"
    except Exception:
        return str(val)

def badge(val, thresholds, labels, classes):
    if val is None:
        return "N/A", "neu"
    for i, t in enumerate(thresholds):
        if val <= t:
            return labels[i], classes[i]
    return labels[-1], classes[-1]

def val_signal(current, sector_avg):
    if current is None or not sector_avg:
        return "N/A", "neu"
    r = current / sector_avg
    if r < 0.9:
        return "CHEAP",     "ok"
    if r > 1.1:
        return "EXPENSIVE", "bad"
    return "FAIR", "warn"

def growth_class(rev3, rev5, np3, np5):
    scores = [v for v in [rev3, rev5, np3, np5] if v is not None]
    if not scores:
        return "UNKNOWN"
    avg = sum(scores) / len(scores)
    if avg > 18: return "ACCELERATING"
    if avg > 10: return "STEADY"
    if avg > 0:  return "SLOWING"
    return "DECLINING"

# ─────────────────────────────────────────────────────────────────
# PER-STOCK ANALYSIS PIPELINE
# ─────────────────────────────────────────────────────────────────

def analyse(ticker: str) -> dict:
    sc  = screener_fetch(ticker)
    yfd = yf_fetch(ticker)
    time.sleep(1.5)

    company  = yfd.get("company_name", ticker)
    sector   = yfd.get("sector", "N/A")
    industry = yfd.get("industry", "N/A")
    desc     = yfd.get("description", "")

    cmp      = yfd.get("cmp")
    high52   = yfd.get("52w_high_yf") or _clean_num(sc.get("52w_high"))
    low52    = yfd.get("52w_low_yf")  or _clean_num(sc.get("52w_low"))
    mcap_yf  = yfd.get("market_cap_yf")
    mcap_str = f"\u20b9{mcap_yf/1e7:,.0f} Cr" if mcap_yf else sc.get("market_cap", "N/A")
    fv       = yfd.get("face_value_yf") or sc.get("face_value", "N/A")

    pe    = yfd.get("pe_yf")   or _clean_num(sc.get("pe"))
    pb    = yfd.get("pb_yf")
    ev_eb = yfd.get("ev_ebitda_yf")
    sp    = sector_pe(sector)
    pe_sig, pe_cls = val_signal(pe, sp)
    pb_sig, pb_cls = val_signal(pb, 3.0)
    ev_sig, ev_cls = val_signal(ev_eb, 15.0)
    v = [pe_cls, pb_cls, ev_cls]
    if v.count("ok") >= 2:     overall_val = "UNDERVALUED"
    elif v.count("bad") >= 2:  overall_val = "OVERVALUED"
    elif v.count("warn") >= 2: overall_val = "FAIRLY VALUED"
    else:                       overall_val = "MIXED"

    rl   = rev_list(sc)
    nl   = np_list(sc)
    el   = eps_list(sc)
    rev3 = cagr(rl, 3) if len(rl) >= 4 else None
    rev5 = cagr(rl, 5) if len(rl) >= 6 else None
    np3  = cagr(nl, 3) if len(nl) >= 4 else None
    np5  = cagr(nl, 5) if len(nl) >= 6 else None
    eps3 = cagr(el, 3) if len(el) >= 4 else None
    eps5 = cagr(el, 5) if len(el) >= 6 else None
    opm  = latest(sc.get("opm_annual", {}))
    gclass = growth_class(rev3, rev5, np3, np5)

    eps_q    = sc.get("eps_quarterly") or yfd.get("eps_quarterly_yf") or {}
    eq_items = list(eps_q.items())[:8]
    eps_chips = []
    for i, (qname, val) in enumerate(eq_items):
        if val is None:
            continue
        yoy = None
        if i + 4 < len(eq_items):
            prev = eq_items[i + 4][1]
            if prev and prev != 0:
                yoy = round((val - prev) / abs(prev) * 100, 1)
        eps_chips.append({"q": qname, "v": val, "yoy": yoy})

    d_e  = de_ratio(sc, yfd)
    cr   = current_ratio(yfd)
    fcfc = fcf_crore(yfd)
    de_lbl,  de_cls  = badge(d_e,  [1, 2],   ["SAFE","MODERATE","LEVERAGED"], ["ok","warn","bad"]) if d_e  is not None else ("N/A","neu")
    cr_lbl,  cr_cls  = badge(cr,   [1, 1.5], ["RISK","WATCH","COMFORTABLE"],  ["bad","warn","ok"])  if cr   is not None else ("N/A","neu")
    fcf_lbl, fcf_cls = ("CONCERN","bad") if fcfc and fcfc < 0 else (("STRONG","ok") if fcfc else ("N/A","neu"))
    hs = [de_cls, cr_cls, fcf_cls]
    if hs.count("bad") >= 2:    overall_health, hcard = "HIGH RISK",     "red"
    elif hs.count("warn") >= 2: overall_health, hcard = "MODERATE RISK", "amber"
    else:                        overall_health, hcard = "SAFE",          "green"
    de_trend = trend_arrow(sc.get("borrowings_annual", {}))

    HORIZON = 5
    br  = (rev5 or rev3 or 10) / 100
    bn  = (np5  or np3  or 10) / 100
    be  = (eps5 or eps3 or 10) / 100
    rla = rl[-1] if rl else None
    nla = nl[-1] if nl else None
    bev = latest(eps_q) if eps_q else (el[-1] if el else None)

    def proj(base, rate, mult):
        return round(base * ((1 + rate * mult) ** HORIZON), 0) if base else None

    proj_rev = [proj(rla, br, m) for m in [0.4, 1.0, 1.5]]
    proj_np  = [proj(nla, bn, m) for m in [0.4, 1.0, 1.5]]
    proj_eps = [round(bev * ((1 + be * m) ** HORIZON), 1) if bev else None for m in [0.4, 1.0, 1.5]]

    roe_raw  = yfd.get("roe_yf")
    roe_pct  = round(roe_raw * 100, 1) if roe_raw else _clean_num(sc.get("roe"))
    roce_raw = _clean_num(sc.get("roce"))
    dy_raw   = yfd.get("div_yield_yf")
    div_y    = round(dy_raw * 100, 2) if dy_raw else _clean_num(sc.get("div_yield"))
    payout_r = yfd.get("payout_ratio_yf")
    payout_p = round(payout_r * 100, 1) if payout_r else None
    roe_lbl,  roe_cls  = badge(roe_pct,  [10, 15], ["WEAK","AVERAGE","GOOD"], ["bad","warn","ok"]) if roe_pct  is not None else ("N/A","neu")
    roce_lbl, roce_cls = badge(roce_raw, [10, 15], ["WEAK","AVERAGE","GOOD"], ["bad","warn","ok"]) if roce_raw is not None else ("N/A","neu")
    if roe_cls == "ok":        ret_q = "HIGH-QUALITY COMPOUNDER"
    elif div_y and div_y > 3:  ret_q = "DIVIDEND PLAY"
    elif roe_cls == "warn":    ret_q = "AVERAGE RETURNS"
    else:                       ret_q = "CAPITAL-LIGHT"

    peers = sc.get("peers", [])
    peer_stand = "MID-PACK"
    if peers and pe:
        peer_pes = [p.get("pe") for p in peers if p.get("pe")]
        if peer_pes:
            if pe < min(peer_pes) * 1.1:  peer_stand = "LEADING"
            elif pe > max(peer_pes) * 0.9: peer_stand = "LAGGING"

    prom_h   = sc.get("promoter_holding", {})
    pledge_d = sc.get("promoter_pledge", {})
    fii_h    = sc.get("fii_holding", {})
    dii_h    = sc.get("dii_holding", {})
    prom_val,   prom_trend = ownership_latest(prom_h)
    fii_val,    fii_trend  = ownership_latest(fii_h)
    dii_val,    dii_trend  = ownership_latest(dii_h)
    pledge_val, _          = ownership_latest(pledge_d)
    prom_trend = prom_trend or "\u2192 Stable"
    fii_trend  = fii_trend  or "\u2192 Stable"
    dii_trend  = dii_trend  or "\u2192 Stable"
    pledge_flag = "bad" if (pledge_val and pledge_val > 10) else "ok"
    pledge_lbl  = "\u26a0 FLAG" if (pledge_val and pledge_val > 10) else "OK"
    prom_badge  = "ok" if any(k in prom_trend.lower() for k in ["increas","stable","buy"]) else "bad"
    fii_badge   = "ok" if any(k in fii_trend.lower()  for k in ["increas","stable"])       else "warn"
    dii_badge   = "ok" if any(k in dii_trend.lower()  for k in ["increas","stable"])       else "warn"

    flags = []
    if pledge_val and pledge_val > 10:
        flags.append(("\u26a0 HIGH PROMOTER PLEDGING",
                       f"{pledge_val:.1f}% of promoter shares pledged as collateral. Forced selling risk if stock falls."))
    if d_e and d_e > 2:
        flags.append(("\u26a0 HIGH LEVERAGE",
                       f"D/E of {d_e:.2f} — heavily reliant on debt. Rising rates could hurt profits."))

    score = 0
    if gclass in ["ACCELERATING","STEADY"]:            score += 2
    if roe_cls == "ok":                                 score += 2
    if overall_val in ["UNDERVALUED","FAIRLY VALUED"]:  score += 1
    if overall_health == "SAFE":                        score += 2
    if pledge_val and pledge_val > 10:                  score -= 2
    if d_e and d_e > 2:                                 score -= 1
    if score >= 5:   view_label, view_card = "STRONG FUNDAMENTALS",   "green"
    elif score >= 2: view_label, view_card = "MODERATE FUNDAMENTALS", "amber"
    else:            view_label, view_card = "WEAK FUNDAMENTALS",     "red"

    strengths, watches = [], []
    if gclass in ["ACCELERATING","STEADY"]:
        strengths.append(f"Revenue growing at ~{pct(rev5 or rev3)} CAGR — consistent top-line expansion")
    if roe_cls == "ok":
        strengths.append(f"Strong ROE of {pct(roe_pct)} — solid returns on shareholder capital")
    if overall_val in ["UNDERVALUED","FAIRLY VALUED"]:
        strengths.append(f"{overall_val.title()} vs sector — P/E {xfmt(pe)} vs sector avg {xfmt(sp)}")
    if overall_health == "SAFE":
        strengths.append("Low debt burden provides resilience in economic downturns")
    while len(strengths) < 3:
        strengths.append("Refer Screener.in / NSE filings for additional qualitative strengths")
    if d_e and d_e > 1.5:
        watches.append(f"Elevated debt (D/E {xfmt(d_e)}) — monitor borrowing trend each quarter")
    if gclass in ["SLOWING","DECLINING"]:
        watches.append("Growth momentum has slowed — watch for margin improvement signals")
    if pledge_val and pledge_val > 5:
        watches.append(f"Promoter pledge at {pct(pledge_val)} — sharp dip could trigger forced selling")
    while len(watches) < 2:
        watches.append("Monitor quarterly earnings for divergence from historical trend")

    live = sum([bool(cmp), bool(pe), bool(pb), bool(rev3), bool(np3),
                bool(eps3), bool(roe_pct), bool(d_e), bool(prom_val),
                bool(fii_val), bool(dii_val), bool(fcfc)])
    if live >= 9:   conf_cls, conf_lbl = "high",     "HIGH"
    elif live >= 6: conf_cls, conf_lbl = "moderate", "MODERATE"
    elif live >= 3: conf_cls, conf_lbl = "low",      "LOW \u26a0"
    else:           conf_cls, conf_lbl = "vlow",     "VERY LOW \u26a0"

    own_signal = (
        "INSIDERS BUILDING" if "increas" in prom_trend.lower()
        else ("TRIMMING" if "decreas" in prom_trend.lower() else "HOLDING STEADY")
    )

    return dict(
        ticker=ticker, company=company, sector=sector, industry=industry,
        desc=desc[:220] + ("\u2026" if len(desc) > 220 else "") if desc else "N/A",
        cmp=cmp, high52=high52, low52=low52, mcap_str=mcap_str, fv=fv,
        pe=pe, pb=pb, ev_eb=ev_eb, sp=sp,
        pe_sig=pe_sig, pe_cls=pe_cls, pb_sig=pb_sig, pb_cls=pb_cls,
        ev_sig=ev_sig, ev_cls=ev_cls, overall_val=overall_val,
        rev3=rev3, rev5=rev5, np3=np3, np5=np5, eps3=eps3, eps5=eps5,
        opm=opm, gclass=gclass, eps_chips=eps_chips,
        d_e=d_e, cr=cr, fcfc=fcfc,
        de_lbl=de_lbl, de_cls=de_cls, cr_lbl=cr_lbl, cr_cls=cr_cls,
        fcf_lbl=fcf_lbl, fcf_cls=fcf_cls,
        overall_health=overall_health, hcard=hcard, de_trend=de_trend,
        proj_rev=proj_rev, proj_np=proj_np, proj_eps=proj_eps,
        roe_pct=roe_pct, roce_raw=roce_raw, div_y=div_y, payout_p=payout_p,
        roe_lbl=roe_lbl, roe_cls=roe_cls, roce_lbl=roce_lbl, roce_cls=roce_cls,
        ret_q=ret_q, peers=peers, peer_stand=peer_stand,
        prom_val=prom_val, prom_trend=prom_trend, prom_badge=prom_badge,
        fii_val=fii_val, fii_trend=fii_trend, fii_badge=fii_badge,
        dii_val=dii_val, dii_trend=dii_trend, dii_badge=dii_badge,
        pledge_val=pledge_val, pledge_flag=pledge_flag, pledge_lbl=pledge_lbl,
        flags=flags, view_label=view_label, view_card=view_card,
        strengths=strengths[:3], watches=watches[:2],
        track="Watch quarterly revenue and margin trend; monitor promoter stake and D/E ratio",
        conf_cls=conf_cls, conf_lbl=conf_lbl, live=live,
        own_signal=own_signal,
        screener_ok=("screener_error" not in sc),
    )

# ─────────────────────────────────────────────────────────────────
# HTML TEMPLATES
# ─────────────────────────────────────────────────────────────────

CSS = """<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --g-fill:#EAF3DE;--g-text:#27500A;--g-border:#C0DD97;--g-accent:#639922;
  --a-fill:#FAEEDA;--a-text:#854F0B;--a-border:#FAC775;--a-accent:#EF9F27;
  --r-fill:#FCEBEB;--r-text:#A32D2D;--r-border:#F7C1C1;--r-accent:#E24B4A;
  --b-fill:#E6F1FB;--b-text:#185FA5;--b-border:#B5D4F4;--b-accent:#378ADD;
  --n-fill:#F1EFE8;--n-text:#444441;--n-border:#D3D1C7;
  --bg:#fff;--bg2:#f5f5f3;--txt:#1a1a1a;--txt2:#666;
  --bdr:#e0ddd5;--bdr2:rgba(0,0,0,0.15);--bdr3:rgba(0,0,0,0.3)
}
@media(prefers-color-scheme:dark){:root{
  --g-fill:#173404;--g-text:#C0DD97;--g-border:#3B6D11;--g-accent:#97C459;
  --a-fill:#412402;--a-text:#FAC775;--a-border:#854F0B;--a-accent:#EF9F27;
  --r-fill:#501313;--r-text:#F09595;--r-border:#A32D2D;--r-accent:#E24B4A;
  --b-fill:#042C53;--b-text:#85B7EB;--b-border:#185FA5;--b-accent:#378ADD;
  --n-fill:#2C2C2A;--n-text:#D3D1C7;--n-border:#5F5E5A;
  --bg:#1c1c1a;--bg2:#2a2a28;--txt:#f0ede6;--txt2:#aaa;
  --bdr:#3a3a38;--bdr2:rgba(255,255,255,0.12);--bdr3:rgba(255,255,255,0.25)
}}
body{font-family:system-ui,sans-serif;font-size:13px;line-height:1.6;color:var(--txt);background:var(--bg);margin:0;padding:0}
.topbar{position:sticky;top:0;z-index:100;background:var(--bg);border-bottom:1px solid var(--bdr);
  padding:10px 20px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.topbar h1{font-size:15px;font-weight:600;white-space:nowrap}
.sw{position:relative;flex:1;min-width:180px;max-width:320px}
.sw input{width:100%;padding:7px 10px 7px 30px;border-radius:8px;border:1px solid var(--bdr2);
  background:var(--bg2);color:var(--txt);font-size:13px;outline:none;font-family:inherit}
.sw input:focus{border-color:var(--bdr3)}
.si{position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--txt2);font-size:13px;pointer-events:none}
.stock-select{padding:7px 10px;border-radius:8px;border:1px solid var(--bdr2);
  background:var(--bg2);color:var(--txt);font-size:13px;font-family:inherit;cursor:pointer;min-width:200px}
.gen-time{font-size:11px;color:var(--txt2);white-space:nowrap;margin-left:auto}
.landing{max-width:1100px;margin:0 auto;padding:20px}
.landing h2{font-size:14px;font-weight:500;margin-bottom:14px;color:var(--txt2)}
.sgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:9px}
.sgc{background:var(--bg2);border:.5px solid var(--bdr);border-radius:10px;
  padding:13px 15px;cursor:pointer;transition:border-color .15s,transform .1s}
.sgc:hover{border-color:var(--bdr3);transform:translateY(-1px)}
.sgc .sn{font-size:13px;font-weight:500;margin-bottom:3px}
.sgc .st{font-size:10px;color:var(--txt2);text-transform:uppercase;letter-spacing:.8px;margin-bottom:7px}
.sgc .sb{display:flex;gap:4px;flex-wrap:wrap}
.sgc .sc{font-size:16px;font-weight:600;margin-top:6px}
.stock-panel{display:none;max-width:900px;margin:0 auto;padding:18px 20px 48px}
.stock-panel.active{display:block}
.ph{margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.ph h2{font-size:19px;font-weight:600}
.ph .sub{font-size:12px;color:var(--txt2)}
.back-btn{background:none;border:1px solid var(--bdr2);border-radius:8px;padding:5px 11px;
  font-size:12px;cursor:pointer;color:var(--txt2);font-family:inherit}
.back-btn:hover{border-color:var(--bdr3);color:var(--txt)}
.conf{padding:8px 13px;border-radius:8px;margin-bottom:13px;font-size:12px}
.conf.high,.conf.moderate{background:var(--g-fill);color:var(--g-text)}
.conf.low{background:var(--a-fill);color:var(--a-text)}
.conf.vlow{background:var(--r-fill);color:var(--r-text)}
.tab-row{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:16px}
.tab-btn{padding:5px 14px;border-radius:100px;border:1px solid var(--bdr2);
  background:transparent;color:var(--txt2);font-size:12px;cursor:pointer;font-family:inherit;transition:.15s}
.tab-btn:hover{border-color:var(--bdr3);color:var(--txt)}
.tab-btn.active{border:1.5px solid var(--txt);color:var(--txt);font-weight:500}
.tpanel{display:none}.tpanel.on{display:block}
.card{background:var(--bg);border:.5px solid var(--bdr);border-radius:12px;padding:15px 18px;margin-bottom:11px}
.card.green{background:var(--g-fill);border-color:var(--g-border)}
.card.amber{background:var(--a-fill);border-color:var(--a-border)}
.card.red{background:var(--r-fill);border-color:var(--r-border)}
.mgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(138px,1fr));gap:8px;margin:11px 0}
.mc{background:var(--bg2);border-radius:8px;padding:10px 12px}
.mc .ml{font-size:10px;text-transform:uppercase;letter-spacing:.7px;color:var(--txt2);margin-bottom:2px}
.mc .mv{font-size:16px;font-weight:500}
.mc .ms{font-size:10px;color:var(--txt2);margin-top:1px}
.info-row{display:flex;justify-content:space-between;align-items:baseline;
  padding:7px 0;border-bottom:.5px solid var(--bdr);gap:12px}
.info-row:last-child{border-bottom:none}
.il{font-size:12px;color:var(--txt2)}
.iv{font-size:12px;font-weight:500;text-align:right;max-width:380px}
table{width:100%;border-collapse:collapse;font-size:12px;margin:7px 0}
th{text-align:left;padding:6px 9px;font-weight:500;font-size:11px;text-transform:uppercase;
  letter-spacing:.6px;color:var(--txt2);border-bottom:.5px solid var(--bdr)}
td{padding:8px 9px;border-bottom:.5px solid var(--bdr)}
tr:last-child td{border-bottom:none}
.peer-you{background:var(--bg2)}
.badge{display:inline-block;padding:2px 7px;border-radius:100px;font-size:11px;font-weight:500}
.badge.ok{background:var(--g-fill);color:var(--g-text)}
.badge.warn{background:var(--a-fill);color:var(--a-text)}
.badge.bad{background:var(--r-fill);color:var(--r-text)}
.badge.neu{background:var(--n-fill);color:var(--n-text)}
.sec-label{font-size:10px;font-weight:500;text-transform:uppercase;letter-spacing:1px;
  color:var(--txt2);margin:13px 0 7px}
.bullet-row{display:flex;gap:8px;align-items:flex-start;padding:5px 0;
  border-bottom:.5px solid var(--bdr);font-size:13px}
.bullet-row:last-child{border-bottom:none}
.bicon{font-size:13px;min-width:16px;margin-top:1px}
.score-box{background:var(--bg2);border-radius:8px;padding:12px 14px;margin-top:11px;font-size:12px;color:var(--txt2)}
.eps-row{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0}
.eps-chip{background:var(--bg2);border-radius:8px;padding:8px 11px;text-align:center;min-width:78px}
.eps-q{font-size:10px;color:var(--txt2);text-transform:uppercase;letter-spacing:.4px}
.eps-v{font-size:14px;font-weight:500;margin:2px 0}
.eps-y{font-size:11px}
.pos{color:var(--g-accent);font-weight:500}.neg{color:var(--r-accent);font-weight:500}
.flag-card{border-left:3px solid var(--a-accent);background:var(--a-fill);
  border-radius:0 8px 8px 0;padding:8px 12px;margin-bottom:7px}
.flag-title{font-size:13px;font-weight:500;color:var(--a-text);margin-bottom:2px}
.flag-note{font-size:12px;color:var(--a-text)}
.view-label{font-size:16px;font-weight:500;margin-bottom:5px}
.view-label.green{color:var(--g-text)}.view-label.amber{color:var(--a-text)}.view-label.red{color:var(--r-text)}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:4px}
@media(max-width:520px){.two-col{grid-template-columns:1fr}}
.cap-note{font-size:11px;color:var(--txt2);margin-top:6px}
.disc{font-size:11px;color:var(--txt2);border-top:.5px solid var(--bdr);padding-top:11px;margin-top:18px;line-height:1.6}
details summary{font-size:12px;color:var(--txt2);cursor:pointer;padding:6px 0}
details p{font-size:12px;color:var(--txt2);padding:4px 0;border-bottom:.5px solid var(--bdr)}
details p:last-child{border-bottom:none}
details strong{color:var(--txt);font-weight:500}
</style>"""

GLOSSARY = """<details style="margin-top:16px;padding:0 2px">
  <summary>Definitions</summary>
  <p><strong>P/E</strong> — What you pay per Rs 1 of profit. Lower vs sector = cheaper.</p>
  <p><strong>P/B</strong> — Price vs net assets. Below 1 = buying at a discount to book value.</p>
  <p><strong>EV/EBITDA</strong> — Full business value check including debt. Lower = better value.</p>
  <p><strong>ROE</strong> — Profit generated per Rs 100 of shareholder equity. Above 15% = good.</p>
  <p><strong>ROCE</strong> — How efficiently the whole business uses all capital. Above 15% = healthy.</p>
  <p><strong>Free Cash Flow</strong> — Cash left after all expenses and capex. Positive and growing = genuinely healthy.</p>
  <p><strong>Promoter pledging</strong> — Founders borrowing against their shares. Above 10% = red flag.</p>
  <p><strong>CAGR</strong> — Compound Annual Growth Rate. Average yearly growth over a period.</p>
</details>"""

DISCLAIMER = """<div class="disc">
  Fundamental screening tool only. Data: yfinance (Yahoo Finance) + Screener.in.
  NOT investment advice or SEBI-registered research. Verify all numbers on NSE/BSE/Screener.in before any decision.
  Past performance does not guarantee future results. Consult a SEBI-registered financial advisor before investing.
</div>"""

SCRIPT = """<script>
var allStocks=[];
function init(s){allStocks=s;renderGrid(s);}
function renderGrid(stocks){
  var g=document.getElementById('sgrid');
  g.innerHTML='';
  stocks.forEach(function(s){
    var vc=s.view_card==='green'?'ok':(s.view_card==='red'?'bad':'warn');
    var el=document.createElement('div');
    el.className='sgc';
    el.onclick=function(){openStock(s.ticker);};
    el.innerHTML='<div class="sn">'+s.company+'</div>'+
      '<div class="st">'+s.ticker+'</div>'+
      '<div class="sb">'+
        '<span class="badge '+vc+'">'+s.view_label.replace(' FUNDAMENTALS','')+'</span>'+
        '<span class="badge neu">'+s.overall_val+'</span>'+
      '</div>'+
      '<div class="sc">'+(s.cmp?'\u20b9'+parseFloat(s.cmp).toFixed(2):'N/A')+'</div>';
    g.appendChild(el);
  });
}
function openStock(ticker){
  document.getElementById('landing').style.display='none';
  document.querySelectorAll('.stock-panel').forEach(function(p){p.classList.remove('active');});
  var panel=document.getElementById('sp-'+ticker);
  if(panel){panel.classList.add('active');showTab(ticker,7);window.scrollTo(0,0);}
  document.getElementById('stock-select').value=ticker;
  document.getElementById('search-input').value='';
}
function showLanding(){
  document.querySelectorAll('.stock-panel').forEach(function(p){p.classList.remove('active');});
  document.getElementById('landing').style.display='block';
  document.getElementById('stock-select').value='';
  renderGrid(allStocks);
  window.scrollTo(0,0);
}
function showTab(ticker,n){
  var base='#sp-'+ticker+' ';
  document.querySelectorAll(base+'.tpanel').forEach(function(p,i){p.className='tpanel'+(i===n?' on':'');});
  document.querySelectorAll(base+'.tab-btn').forEach(function(b,i){b.className='tab-btn'+(i===n?' active':'');});
}
function onSelect(val){if(!val){showLanding();}else{openStock(val);}}
function onSearch(val){
  val=val.toLowerCase().trim();
  if(!val){document.getElementById('landing').style.display='block';renderGrid(allStocks);return;}
  var f=allStocks.filter(function(s){
    return s.ticker.toLowerCase().includes(val)||s.company.toLowerCase().includes(val)||
           (s.sector||'').toLowerCase().includes(val)||(s.industry||'').toLowerCase().includes(val);
  });
  document.getElementById('landing').style.display='block';
  document.querySelectorAll('.stock-panel').forEach(function(p){p.classList.remove('active');});
  renderGrid(f);
}
</script>"""


def render_panel(d: dict) -> str:
    t = d["ticker"]
    esc = t.replace("-","_").replace("&","_")   # safe JS id

    flags_html = "".join(
        f'<div class="flag-card"><div class="flag-title">{f[0]}</div>'
        f'<div class="flag-note">{f[1]}</div></div>'
        for f in d["flags"]
    )

    chips_html = ""
    for c in d["eps_chips"]:
        yoy_html = ""
        if c["yoy"] is not None:
            cls  = "pos" if c["yoy"] >= 0 else "neg"
            sign = "+" if c["yoy"] >= 0 else ""
            yoy_html = f'<div class="eps-y {cls}">{sign}{c["yoy"]}% YoY</div>'
        chips_html += (f'<div class="eps-chip"><div class="eps-q">{c["q"]}</div>'
                       f'<div class="eps-v">\u20b9{c["v"]}</div>{yoy_html}</div>')
    if not chips_html:
        chips_html = '<span style="color:var(--txt2);font-size:12px">Data unavailable — verify at screener.in</span>'

    peer_rows = ""
    for p in d["peers"][:3]:
        peer_rows += (f'<tr><td>{p.get("name","N/A")}</td><td>{xfmt(p.get("pe"))}</td>'
                      f'<td>{xfmt(p.get("pb"))}</td><td>{pct(p.get("roe"))}</td>'
                      f'<td>{pct(p.get("rev_growth"))}</td><td>N/A</td></tr>')
    if not peer_rows:
        peer_rows = '<tr><td colspan="6" style="color:var(--txt2);text-align:center">Verify at screener.in</td></tr>'

    str_html = "".join(
        f'<div class="bullet-row"><span class="bicon" style="color:var(--g-accent)">&#10003;</span><span>{s}</span></div>'
        for s in d["strengths"]
    )
    wat_html = "".join(
        f'<div class="bullet-row"><span class="bicon" style="color:var(--a-accent)">&#9888;</span><span>{w}</span></div>'
        for w in d["watches"]
    )

    pr = d["proj_rev"]; pn = d["proj_np"]; pe_ = d["proj_eps"]

    def prom_lbl(trend):
        tl = trend.lower()
        if "increas" in tl or "buy" in tl: return "BUYING"
        if "decreas" in tl or "sell" in tl: return "SELLING"
        return "STABLE"

    def fi_lbl(trend):
        tl = trend.lower()
        if "increas" in tl: return "INCREASING"
        if "decreas" in tl: return "DECREASING"
        return "STABLE"

    rev_trend = "\U0001f4c8" if (d["rev5"] or d["rev3"] or 0) > 10 else ("\U0001f4c9" if (d["rev5"] or d["rev3"] or 0) < 0 else "\u27a1\ufe0f")
    np_trend  = "\U0001f4c8" if (d["np5"]  or d["np3"]  or 0) > 10 else ("\U0001f4c9" if (d["np5"]  or d["np3"]  or 0) < 0 else "\u27a1\ufe0f")
    eps_trend = "\U0001f4c8" if (d["eps5"] or d["eps3"] or 0) > 10 else ("\U0001f4c9" if (d["eps5"] or d["eps3"] or 0) < 0 else "\u27a1\ufe0f")

    quality_word = "solid" if d["roe_cls"] == "ok" else "moderate"
    health_word  = "healthy" if d["overall_health"] == "SAFE" else "elevated"
    opp3 = (f'Dividend yield {pct(d["div_y"])} provides income floor'
            if d["div_y"] and d["div_y"] > 1 else "Potential re-rating if growth accelerates")
    risk1 = (f'Promoter pledge {pct(d["pledge_val"])} — forced selling risk'
             if d["pledge_val"] and d["pledge_val"] > 10 else "Regulatory or macro headwinds in sector")
    risk2 = (f'High leverage D/E {xfmt(d["d_e"])} — sensitive to rate cycles'
             if d["d_e"] and d["d_e"] > 1.5 else f'Competition intensifying in {d["industry"]}')

    return f"""
<div class="stock-panel" id="sp-{t}">
  <div class="ph">
    <button class="back-btn" onclick="showLanding()">&larr; All Stocks</button>
    <div>
      <h2>{d["company"]} <span style="font-size:13px;font-weight:400;color:var(--txt2)">({t})</span></h2>
      <div class="sub">{d["sector"]} &middot; {d["industry"]}</div>
    </div>
  </div>
  <div class="conf {d["conf_cls"]}">
    <strong>Data confidence: {d["conf_lbl"]}</strong> &nbsp;&middot;&nbsp;
    {d["live"]}/12 metrics live &nbsp;&middot;&nbsp;
    Sources: yfinance{"+ Screener.in" if d["screener_ok"] else " only"}
  </div>
  <div class="tab-row">
    <button class="tab-btn" onclick="showTab('{t}',0)">Snapshot</button>
    <button class="tab-btn" onclick="showTab('{t}',1)">Valuation</button>
    <button class="tab-btn" onclick="showTab('{t}',2)">Growth</button>
    <button class="tab-btn" onclick="showTab('{t}',3)">Health</button>
    <button class="tab-btn" onclick="showTab('{t}',4)">Returns</button>
    <button class="tab-btn" onclick="showTab('{t}',5)">Peers</button>
    <button class="tab-btn" onclick="showTab('{t}',6)">Ownership</button>
    <button class="tab-btn active" onclick="showTab('{t}',7)">View</button>
  </div>

  <div class="tpanel" id="tp-{esc}-0">
    <div class="card">
      <div class="info-row"><span class="il">Company</span><span class="iv">{d["company"]}</span></div>
      <div class="info-row"><span class="il">Ticker</span><span class="iv">{t} &middot; NSE</span></div>
      <div class="info-row"><span class="il">Sector</span><span class="iv">{d["sector"]}</span></div>
      <div class="info-row"><span class="il">Industry</span><span class="iv">{d["industry"]}</span></div>
      <div class="info-row"><span class="il">Business</span><span class="iv" style="max-width:380px">{d["desc"]}</span></div>
    </div>
    <div class="mgrid">
      <div class="mc"><div class="ml">CMP</div><div class="mv">{fmt_inr(d["cmp"])}</div><div class="ms">NSE / Yahoo</div></div>
      <div class="mc"><div class="ml">52W High</div><div class="mv">{fmt_inr(d["high52"])}</div><div class="ms">Yahoo</div></div>
      <div class="mc"><div class="ml">52W Low</div><div class="mv">{fmt_inr(d["low52"])}</div><div class="ms">Yahoo</div></div>
      <div class="mc"><div class="ml">Market Cap</div><div class="mv">{d["mcap_str"]}</div><div class="ms">Yahoo</div></div>
      <div class="mc"><div class="ml">Face Value</div><div class="mv">\u20b9{d["fv"]}</div><div class="ms">NSE</div></div>
    </div>
    {flags_html}
  </div>

  <div class="tpanel" id="tp-{esc}-1">
    <div class="card">
      <table>
        <thead><tr><th>Metric</th><th>Current</th><th>Sector avg</th><th>Signal</th><th>Plain English</th></tr></thead>
        <tbody>
          <tr><td>P/E</td><td>{xfmt(d["pe"])}</td><td>{xfmt(d["sp"])}</td>
            <td><span class="badge {d["pe_cls"]}">{d["pe_sig"]}</span></td>
            <td style="color:var(--txt2)">Pay \u20b9{round(d["pe"],0) if d["pe"] else "?"} per \u20b91 profit</td></tr>
          <tr><td>P/B</td><td>{xfmt(d["pb"])}</td><td>3.0x</td>
            <td><span class="badge {d["pb_cls"]}">{d["pb_sig"]}</span></td>
            <td style="color:var(--txt2)">Price vs net assets</td></tr>
          <tr><td>EV/EBITDA</td><td>{xfmt(d["ev_eb"])}</td><td>15.0x</td>
            <td><span class="badge {d["ev_cls"]}">{d["ev_sig"]}</span></td>
            <td style="color:var(--txt2)">Full business value</td></tr>
        </tbody>
      </table>
    </div>
    <div class="score-box">Overall: <strong>{d["overall_val"]}</strong> &nbsp;&middot;&nbsp;
      Trailing P/E {xfmt(d["pe"])} vs {d["sector"]} sector avg {xfmt(d["sp"])}
    </div>
  </div>

  <div class="tpanel" id="tp-{esc}-2">
    <div class="card">
      <table>
        <thead><tr><th>Metric</th><th>3Y CAGR</th><th>5Y CAGR</th><th>Trend</th><th>Source</th></tr></thead>
        <tbody>
          <tr><td>Revenue</td><td>{pct(d["rev3"])}</td><td>{pct(d["rev5"])}</td><td>{rev_trend}</td><td style="color:var(--txt2)">Screener.in</td></tr>
          <tr><td>Net profit</td><td>{pct(d["np3"])}</td><td>{pct(d["np5"])}</td><td>{np_trend}</td><td style="color:var(--txt2)">Screener.in</td></tr>
          <tr><td>EPS</td><td>{pct(d["eps3"])}</td><td>{pct(d["eps5"])}</td><td>{eps_trend}</td><td style="color:var(--txt2)">Screener.in</td></tr>
          <tr><td>EBITDA margin</td><td>{pct(d["opm"])}</td><td>N/A</td><td>&rarr;</td><td style="color:var(--txt2)">Screener.in</td></tr>
        </tbody>
      </table>
    </div>
    <div class="card">
      <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--txt2);margin-bottom:8px">EPS &mdash; last 8 quarters</div>
      <div class="eps-row">{chips_html}</div>
    </div>
    <div class="score-box">Growth: <strong>{d["gclass"]}</strong> &nbsp;&middot;&nbsp;
      Rev CAGR {pct(d["rev5"])} (5Y) / {pct(d["rev3"])} (3Y) &middot; Net profit CAGR {pct(d["np5"])} (5Y) / {pct(d["np3"])} (3Y)
    </div>
  </div>

  <div class="tpanel" id="tp-{esc}-3">
    <div class="card">
      <table>
        <thead><tr><th>Metric</th><th>Value</th><th>Trend</th><th>Signal</th><th>Benchmark</th></tr></thead>
        <tbody>
          <tr><td>Debt / Equity</td><td>{xfmt(d["d_e"])}</td><td>{d["de_trend"]}</td>
            <td><span class="badge {d["de_cls"]}">{d["de_lbl"]}</span></td><td style="color:var(--txt2)">Below 1 = safe</td></tr>
          <tr><td>Interest Coverage</td><td>N/A</td><td>&rarr;</td>
            <td><span class="badge neu">N/A</span></td><td style="color:var(--txt2)">Above 3x = healthy</td></tr>
          <tr><td>Current Ratio</td><td>{xfmt(d["cr"])}</td><td>&rarr;</td>
            <td><span class="badge {d["cr_cls"]}">{d["cr_lbl"]}</span></td><td style="color:var(--txt2)">Above 1.5 = comfortable</td></tr>
          <tr><td>Free Cash Flow</td><td>{fmt_inr_cr(d["fcfc"])}</td><td>&rarr;</td>
            <td><span class="badge {d["fcf_cls"]}">{d["fcf_lbl"]}</span></td><td style="color:var(--txt2)">Positive = real cash business</td></tr>
        </tbody>
      </table>
    </div>
    <div class="card">
      <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--txt2);margin-bottom:8px">Forward projections &mdash; 5 year horizon (CAGR-based)</div>
      <table>
        <thead><tr><th>Scenario</th><th>Assumption</th><th>Est. revenue</th><th>Est. net profit</th><th>Est. EPS</th></tr></thead>
        <tbody>
          <tr><td>Bear</td><td style="color:var(--txt2)">Growth slows 60%, margins compress</td><td>{fmt_inr_cr(pr[0])}</td><td>{fmt_inr_cr(pn[0])}</td><td>\u20b9{pe_[0] or "N/A"}</td></tr>
          <tr><td>Base</td><td style="color:var(--txt2)">Maintains historical CAGR</td><td>{fmt_inr_cr(pr[1])}</td><td>{fmt_inr_cr(pn[1])}</td><td>\u20b9{pe_[1] or "N/A"}</td></tr>
          <tr><td>Bull</td><td style="color:var(--txt2)">Growth 1.5x, margins expand</td><td>{fmt_inr_cr(pr[2])}</td><td>{fmt_inr_cr(pn[2])}</td><td>\u20b9{pe_[2] or "N/A"}</td></tr>
        </tbody>
      </table>
      <div class="cap-note">Projections based on historical CAGR only &mdash; not guarantees or predictions.</div>
    </div>
    <div class="score-box">Financial health: <strong>{d["overall_health"]}</strong> &nbsp;&middot;&nbsp;
      D/E {xfmt(d["d_e"])} &middot; Current Ratio {xfmt(d["cr"])} &middot; FCF {fmt_inr_cr(d["fcfc"])}
    </div>
  </div>

  <div class="tpanel" id="tp-{esc}-4">
    <div class="card">
      <table>
        <thead><tr><th>Metric</th><th>Current</th><th>Signal</th></tr></thead>
        <tbody>
          <tr><td>ROE</td><td>{pct(d["roe_pct"])}</td><td><span class="badge {d["roe_cls"]}">{d["roe_lbl"]}</span></td></tr>
          <tr><td>ROCE</td><td>{pct(d["roce_raw"])}</td><td><span class="badge {d["roce_cls"]}">{d["roce_lbl"]}</span></td></tr>
          <tr><td>Dividend yield</td><td>{pct(d["div_y"])}</td><td><span class="badge neu">&mdash;</span></td></tr>
          <tr><td>Dividend payout</td><td>{pct(d["payout_p"])}</td><td><span class="badge neu">&mdash;</span></td></tr>
        </tbody>
      </table>
      <div class="cap-note">ROE above 15% = good &middot; ROCE above 15% = efficient capital use</div>
    </div>
    <div class="score-box">Return quality: <strong>{d["ret_q"]}</strong> &nbsp;&middot;&nbsp;
      ROE {pct(d["roe_pct"])} &middot; ROCE {pct(d["roce_raw"])} &middot; Div yield {pct(d["div_y"])}
    </div>
  </div>

  <div class="tpanel" id="tp-{esc}-5">
    <div class="card">
      <table>
        <thead><tr><th>Company</th><th>P/E</th><th>P/B</th><th>ROE</th><th>Rev growth</th><th>D/E</th></tr></thead>
        <tbody>
          <tr class="peer-you">
            <td><strong>{t} &laquo; you</strong></td>
            <td>{xfmt(d["pe"])}</td><td>{xfmt(d["pb"])}</td>
            <td>{pct(d["roe_pct"])}</td><td>{pct(d["rev5"] or d["rev3"])}</td><td>{xfmt(d["d_e"])}</td>
          </tr>
          {peer_rows}
        </tbody>
      </table>
      <div class="cap-note">Source: Screener.in</div>
    </div>
    <div class="score-box">Peer standing: <strong>{d["peer_stand"]}</strong> &nbsp;&middot;&nbsp;
      P/E {xfmt(d["pe"])} vs sector avg {xfmt(d["sp"])} ({d["sector"]})
    </div>
  </div>

  <div class="tpanel" id="tp-{esc}-6">
    <div class="card">
      <table>
        <thead><tr><th>Holder</th><th>Latest %</th><th>Trend</th><th>Signal</th><th>What it means</th></tr></thead>
        <tbody>
          <tr><td>Promoter</td><td>{pct(d["prom_val"])}</td><td>{d["prom_trend"]}</td>
            <td><span class="badge {d["prom_badge"]}">{prom_lbl(d["prom_trend"])}</span></td>
            <td style="color:var(--txt2)">Founder confidence</td></tr>
          <tr><td>FII</td><td>{pct(d["fii_val"])}</td><td>{d["fii_trend"]}</td>
            <td><span class="badge {d["fii_badge"]}">{fi_lbl(d["fii_trend"])}</span></td>
            <td style="color:var(--txt2)">Global fund interest</td></tr>
          <tr><td>DII</td><td>{pct(d["dii_val"])}</td><td>{d["dii_trend"]}</td>
            <td><span class="badge {d["dii_badge"]}">{fi_lbl(d["dii_trend"])}</span></td>
            <td style="color:var(--txt2)">Indian MF &amp; insurance</td></tr>
          <tr><td>Promoter pledge</td><td>{pct(d["pledge_val"])}</td><td>&mdash;</td>
            <td><span class="badge {d["pledge_flag"]}">{d["pledge_lbl"]}</span></td>
            <td style="color:var(--txt2)">Above 10% = red flag</td></tr>
        </tbody>
      </table>
    </div>
    <div class="score-box">Ownership: <strong>{d["own_signal"]}</strong> &nbsp;&middot;&nbsp;
      Promoter {pct(d["prom_val"])} &middot; FII {pct(d["fii_val"])} &middot; DII {pct(d["dii_val"])}
    </div>
  </div>

  <div class="tpanel on" id="tp-{esc}-7">
    <div class="card {d["view_card"]}">
      <div class="view-label {d["view_card"]}">{d["view_label"]}</div>
      <div style="font-size:13px;margin-bottom:13px">
        {d["company"]} shows {d["gclass"].lower()} growth with {quality_word} returns on equity
        and {health_word} financial risk based on available data.
      </div>
      <div class="sec-label">What works for this stock</div>
      {str_html}
      <div class="sec-label">What to watch</div>
      {wat_html}
      <div class="sec-label">Track going forward</div>
      <div class="bullet-row"><span class="bicon" style="color:var(--txt2)">&rarr;</span><span>{d["track"]}</span></div>
      <div style="margin-top:13px;font-size:11px;color:var(--txt2);font-style:italic">
        This is a VIEW based on fundamentals only. Not a buy/sell recommendation. The decision is always yours.
      </div>
    </div>
    <div class="two-col">
      <div class="card green">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--g-text);margin-bottom:7px">Opportunities</div>
        <div class="bullet-row" style="border-color:var(--g-border)"><span class="bicon" style="color:var(--g-accent)">+</span><span style="color:var(--g-text);font-size:12px">Long-term tailwinds in {d["industry"]}</span></div>
        <div class="bullet-row" style="border-color:var(--g-border)"><span class="bicon" style="color:var(--g-accent)">+</span><span style="color:var(--g-text);font-size:12px">Revenue CAGR {pct(d["rev5"] or d["rev3"])} historical compounding</span></div>
        <div class="bullet-row" style="border-color:var(--g-border);border-bottom:none"><span class="bicon" style="color:var(--g-accent)">+</span><span style="color:var(--g-text);font-size:12px">{opp3}</span></div>
      </div>
      <div class="card red">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--r-text);margin-bottom:7px">Risks</div>
        <div class="bullet-row" style="border-color:var(--r-border)"><span class="bicon" style="color:var(--r-accent)">&minus;</span><span style="color:var(--r-text);font-size:12px">{risk1}</span></div>
        <div class="bullet-row" style="border-color:var(--r-border)"><span class="bicon" style="color:var(--r-accent)">&minus;</span><span style="color:var(--r-text);font-size:12px">{risk2}</span></div>
        <div class="bullet-row" style="border-color:var(--r-border);border-bottom:none"><span class="bicon" style="color:var(--r-accent)">&minus;</span><span style="color:var(--r-text);font-size:12px">Automated report &mdash; verify data on NSE/BSE/Screener.in</span></div>
      </div>
    </div>
  </div>

  {GLOSSARY}
  {DISCLAIMER}
</div>"""


def build_html(all_data: list, generated_at: str) -> str:
    opts = '<option value="">— Select a stock —</option>\n'
    for d in all_data:
        opts += f'<option value="{d["ticker"]}">{d["ticker"]} — {d["company"]}</option>\n'

    panels = "\n".join(render_panel(d) for d in all_data)

    js_data = json.dumps([{
        "ticker":     d["ticker"],
        "company":    d["company"],
        "sector":     d["sector"],
        "industry":   d["industry"],
        "view_label": d["view_label"],
        "view_card":  d["view_card"],
        "overall_val":d["overall_val"],
        "cmp":        d["cmp"],
    } for d in all_data], ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Nifty 50 Fundamental Report</title>
{CSS}
</head>
<body>
<div class="topbar">
  <h1>&#128202; Nifty 50 Fundamentals</h1>
  <div class="sw">
    <span class="si">&#128269;</span>
    <input id="search-input" type="text" placeholder="Search company, sector&hellip;"
      oninput="onSearch(this.value)">
  </div>
  <select id="stock-select" class="stock-select" onchange="onSelect(this.value)">
    {opts}
  </select>
  <span class="gen-time">Generated {generated_at}</span>
</div>
<div class="landing" id="landing">
  <h2>All {len(all_data)} stocks &mdash; click any card to open full report</h2>
  <div class="sgrid" id="sgrid"></div>
</div>
{panels}
{SCRIPT}
<script>init({js_data});</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(description="Nifty 50 Fundamental Analyser")
    p.add_argument("--out",     default="reports", help="Output directory (default: reports/)")
    p.add_argument("--tickers", nargs="*",          help="Override ticker list. Default: all Nifty 50")
    return p.parse_args()


def main():
    args    = parse_args()
    tickers = args.tickers if args.tickers else NIFTY50
    total   = len(tickers)

    print(f"\n[Nifty 50 Analyser] Processing {total} stocks in parallel (10 workers) ...\n")

    results   = []
    completed = 0
    lock      = threading.Lock()

    with ThreadPoolExecutor(max_workers=10) as executor:
        future_to_ticker = {executor.submit(analyse, t): t for t in tickers}
        for future in as_completed(future_to_ticker):
            ticker = future_to_ticker[future]
            with lock:
                completed += 1
                idx = completed
            try:
                d = future.result()
                with lock:
                    results.append(d)
                print(f"  [{idx:02d}/{total}] {ticker:15s} conf={d['conf_lbl']:<12} growth={d['gclass']:<14} val={d['overall_val']}")
            except Exception as e:
                print(f"  [{idx:02d}/{total}] {ticker:15s} ERROR: {e}")

    if not results:
        print("\nNo data retrieved. Check network connection.")
        sys.exit(1)

    # Restore original NIFTY50 order for consistent HTML output
    order = {t: i for i, t in enumerate(tickers)}
    results.sort(key=lambda d: order.get(d["ticker"], 9999))

    generated_at = datetime.now().strftime("%d %b %Y, %H:%M IST")
    html  = build_html(results, generated_at)
    out_d = Path(args.out)
    out_d.mkdir(parents=True, exist_ok=True)
    fname = out_d / f"nifty50_{date.today().isoformat()}.html"
    fname.write_text(html, encoding="utf-8")

    print(f"\n  Report ({len(results)} stocks) -> {fname}")
    print("  Open in any browser. No internet needed after generation.\n")


if __name__ == "__main__":
    main()
