"""
NIFTY 100 COMPLETE STOCK ANALYZER — AURORA GLASS THEME v2
Technical + Fundamental Analysis with Email Delivery + GitHub Pages

UPGRADES from v1:
  - 12-Month Swing-High/Low S/R detection (replacing 60-day quantile)
  - Round number S/R levels (context-aware step sizing)
  - ATR-based stop losses anchored near real S/R zones
  - Beta-adjusted ATR multiplier + maximum SL cap
  - ADX (trend strength indicator)
  - Volume/Avg ratio (20-day)
  - Support Distance % column
  - Dynamic Target Promotion (ATR floor on targets)
  - Earnings Date column
  - Analyst Consensus label
  - Sector column
  - Live index strip: SENSEX / NIFTY 50 / BANK NIFTY
  - Live IST clock
  - Full-width mobile-responsive layout
  - ATH Zone / Partial S/R / Real S/R target badges
  - Stop type badge (ATR Stop vs Beta Cap)

Requirements:
    pip install yfinance pandas numpy pytz
"""

import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timezone
import pytz
import warnings
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

warnings.filterwarnings('ignore')


class Nifty100CompleteAnalyzer:
    def __init__(self):
        self.nifty100_stocks = {
            # ── NIFTY 50 ──────────────────────────────────────────
            'RELIANCE.NS':    'Reliance Industries',
            'TCS.NS':         'TCS',
            'HDFCBANK.NS':    'HDFC Bank',
            'INFY.NS':        'Infosys',
            'ICICIBANK.NS':   'ICICI Bank',
            'HINDUNILVR.NS':  'Hindustan Unilever',
            'BHARTIARTL.NS':  'Bharti Airtel',
            'ITC.NS':         'ITC',
            'SBIN.NS':        'State Bank of India',
            'LT.NS':          'L&T',
            'BAJFINANCE.NS':  'Bajaj Finance',
            'KOTAKBANK.NS':   'Kotak Mahindra Bank',
            'AXISBANK.NS':    'Axis Bank',
            'ASIANPAINT.NS':  'Asian Paints',
            'MARUTI.NS':      'Maruti Suzuki',
            'TITAN.NS':       'Titan Company',
            'SUNPHARMA.NS':   'Sun Pharma',
            'ULTRACEMCO.NS':  'UltraTech Cement',
            'NESTLEIND.NS':   'Nestle India',
            'WIPRO.NS':       'Wipro',
            'HCLTECH.NS':     'HCL Tech',
            'BAJAJFINSV.NS':  'Bajaj Finserv',
            'POWERGRID.NS':   'Power Grid',
            'NTPC.NS':        'NTPC',
            'ONGC.NS':        'ONGC',
            'TECHM.NS':       'Tech Mahindra',
            'M&M.NS':         'M&M',
            'TMCV.NS':        'Tata Motors Commerical',
            'TMPV.NS':        'Tata Motors Passenger',
            'TATASTEEL.NS':   'Tata Steel',
            'INDUSINDBK.NS':  'IndusInd Bank',
            'ADANIPORTS.NS':  'Adani Ports',
            'COALINDIA.NS':   'Coal India',
            'JSWSTEEL.NS':    'JSW Steel',
            'HINDALCO.NS':    'Hindalco',
            'CIPLA.NS':       'Cipla',
            'DRREDDY.NS':     'Dr Reddy',
            'GRASIM.NS':      'Grasim',
            'DIVISLAB.NS':    "Divi's Lab",
            'HEROMOTOCO.NS':  'Hero MotoCorp',
            'EICHERMOT.NS':   'Eicher Motors',
            'BRITANNIA.NS':   'Britannia',
            'APOLLOHOSP.NS':  'Apollo Hospital',
            'BAJAJ-AUTO.NS':  'Bajaj Auto',
            'SHRIRAMFIN.NS':  'Shriram Finance',
            'TATACONSUM.NS':  'Tata Consumer',
            'SBILIFE.NS':     'SBI Life',
            'BPCL.NS':        'BPCL',
            'HDFCLIFE.NS':    'HDFC Life',
            'LTIM.NS':        'LTIMindtree',
            'ADANIENT.NS':    'Adani Enterprises',
            'SIEMENS.NS':     'Siemens India',
            # ── NIFTY NEXT 50 ─────────────────────────────────────
            'HAVELLS.NS':     'Havells India',
            'PIDILITIND.NS':  'Pidilite Industries',
            'DABUR.NS':       'Dabur India',
            'MARICO.NS':      'Marico',
            'GODREJCP.NS':    'Godrej Consumer Products',
            'COLPAL.NS':      'Colgate-Palmolive India',
            'BERGEPAINT.NS':  'Berger Paints',
            'MUTHOOTFIN.NS':  'Muthoot Finance',
            'CHOLAFIN.NS':    'Cholamandalam Investment',
            'BAJAJHLDNG.NS':  'Bajaj Holdings',
            'SBICARD.NS':     'SBI Cards',
            'ICICIPRULI.NS':  'ICICI Prudential Life',
            'ICICIGI.NS':     'ICICI Lombard General Insurance',
            'HDFCAMC.NS':     'HDFC AMC',
            'NAUKRI.NS':      'Info Edge (Naukri)',
            'MCDOWELL-N.NS':  'United Spirits',
            'TATAELXSI.NS':   'Tata Elxsi',
            'COFORGE.NS':     'Coforge',
            'PERSISTENT.NS':  'Persistent Systems',
            'OFSS.NS':        'Oracle Financial Services',
            'LTTS.NS':        'L&T Technology Services',
            'PAGEIND.NS':     'Page Industries',
            'VOLTAS.NS':      'Voltas',
            'AMBUJACEM.NS':   'Ambuja Cements',
            'ACC.NS':         'ACC',
            'INDIGO.NS':      'IndiGo (InterGlobe Aviation)',
            'DMART.NS':       'Avenue Supermarts (DMart)',
            'VEDL.NS':        'Vedanta',
            'SAIL.NS':        'Steel Authority of India',
            'NMDC.NS':        'NMDC',
            'RECLTD.NS':      'REC Limited',
            'PFC.NS':         'Power Finance Corporation',
            'IRCTC.NS':       'IRCTC',
            'CONCOR.NS':      'Container Corporation of India',
            'JINDALSTEL.NS':  'Jindal Steel & Power',
            'MOTHERSON.NS':   'Samvardhana Motherson',
            'BALKRISIND.NS':  'Balkrishna Industries',
            'TORNTPHARM.NS':  'Torrent Pharmaceuticals',
            'LUPIN.NS':       'Lupin',
            'AUROPHARMA.NS':  'Aurobindo Pharma',
            'ALKEM.NS':       'Alkem Laboratories',
            'MAXHEALTH.NS':   'Max Healthcare',
            'FORTIS.NS':      'Fortis Healthcare',
            'ZOMATO.NS':      'Zomato',
            'POLICYBZR.NS':   'PB Fintech (PolicyBazaar)',
            'NYKAA.NS':       'FSN E-Commerce (Nykaa)',
            'PAYTM.NS':       'One97 Communications (Paytm)',
            'RVNL.NS':        'Rail Vikas Nigam',
            'ADANIGREEN.NS':  'Adani Green Energy',
        }
        self.results = []

    # =========================================================================
    #  UTILITY
    # =========================================================================
    def get_ist_time(self):
        return datetime.now(pytz.timezone('Asia/Kolkata'))

    def calculate_rsi(self, prices, period=14):
        delta = prices.diff()
        gain  = (delta.where(delta > 0, 0)).rolling(window=period).mean()
        loss  = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
        rs    = gain / loss
        return (100 - (100 / (1 + rs))).iloc[-1]

    def calculate_macd(self, prices):
        ema12  = prices.ewm(span=12, adjust=False).mean()
        ema26  = prices.ewm(span=26, adjust=False).mean()
        macd   = ema12 - ema26
        signal = macd.ewm(span=9, adjust=False).mean()
        return macd.iloc[-1], signal.iloc[-1]

    def calculate_atr(self, df, period=14):
        high  = df['High']
        low   = df['Low']
        close = df['Close']
        tr = pd.concat([
            high - low,
            abs(high - close.shift(1)),
            abs(low  - close.shift(1))
        ], axis=1).max(axis=1)
        return round(tr.ewm(alpha=1 / period, adjust=False).mean().iloc[-1], 2)

    def calculate_adx(self, df, period=14):
        high  = df['High']
        low   = df['Low']
        close = df['Close']
        plus_dm  = high.diff()
        minus_dm = low.diff().abs()
        plus_dm[plus_dm < 0]   = 0
        minus_dm[minus_dm < 0] = 0
        plus_dm[plus_dm < minus_dm]  = 0
        minus_dm[minus_dm < plus_dm] = 0
        tr = pd.concat([
            high - low,
            abs(high - close.shift(1)),
            abs(low  - close.shift(1))
        ], axis=1).max(axis=1)
        atr14    = tr.ewm(alpha=1/period, adjust=False).mean()
        plus_di  = 100 * (plus_dm.ewm(alpha=1/period, adjust=False).mean() / atr14)
        minus_di = 100 * (minus_dm.ewm(alpha=1/period, adjust=False).mean() / atr14)
        dx       = 100 * abs(plus_di - minus_di) / (plus_di + minus_di)
        adx      = dx.ewm(alpha=1/period, adjust=False).mean()
        return round(adx.iloc[-1], 1)

    def calculate_volume_ratio(self, df):
        avg_vol = df['Volume'].tail(20).mean()
        if avg_vol == 0:
            return 1.0
        return round(df['Volume'].iloc[-1] / avg_vol, 2)

    def get_earnings_date(self, info):
        try:
            ts = (info.get('earningsTimestamp') or
                  info.get('earningsTimestampStart') or
                  info.get('earningsDate'))
            if ts:
                if isinstance(ts, (list, tuple)):
                    ts = ts[0]
                if isinstance(ts, (int, float)):
                    dt = datetime.fromtimestamp(ts, tz=timezone.utc)
                    return dt.strftime('%d %b %Y')
                if hasattr(ts, 'strftime'):
                    return ts.strftime('%d %b %Y')
        except Exception:
            pass
        return "N/A"

    def fetch_index_data(self):
        """Fetch SENSEX / NIFTY 50 / BANK NIFTY at report time."""
        indices = {
            'SENSEX':    '^BSESN',
            'NIFTY 50':  '^NSEI',
            'BANK NIFTY':'^NSEBANK',
        }
        result = {}
        for label, sym in indices.items():
            try:
                d     = yf.Ticker(sym).history(period='2d')
                price = d['Close'].iloc[-1]
                prev  = d['Close'].iloc[-2]
                chg   = price - prev
                pct   = chg / prev * 100
                arrow = '▲' if chg >= 0 else '▼'
                cls   = 'up' if chg >= 0 else 'dn'
                sign  = '+' if chg >= 0 else ''
                result[label] = {
                    'price': f"{price:,.2f}",
                    'chg':   f"{arrow} {sign}{pct:.2f}%",
                    'cls':   cls,
                }
            except Exception:
                result[label] = {'price': 'N/A', 'chg': '—', 'cls': ''}
        return result

    # =========================================================================
    #  RESISTANCE & SUPPORT  (12-month swing detection + round numbers + 52W)
    # =========================================================================
    def find_resistance_levels(self, df, current_price, num_levels=5):
        window      = 5
        swing_highs = []
        for src_days in [180, 252]:
            highs = df.tail(src_days)['High'].values
            for i in range(window, len(highs) - window):
                if (highs[i] > max(highs[i-window:i]) and
                        highs[i] > max(highs[i+1:i+window+1])):
                    swing_highs.append(highs[i])
        # 52-week high
        high_52w = df['High'].tail(252).max()
        if high_52w > current_price * 1.005:
            swing_highs.append(high_52w)
        # Round-number levels above price
        magnitude = 10 ** (len(str(int(current_price))) - 2)
        step      = magnitude * 5
        level     = current_price
        for _ in range(20):
            level += step
            if level <= current_price * 1.30:
                swing_highs.append(level)
        if not swing_highs:
            return []
        swing_highs = sorted(set([round(h, 2) for h in swing_highs]))
        clusters, cluster = [], [swing_highs[0]]
        for lv in swing_highs[1:]:
            if (lv - cluster[-1]) / cluster[-1] < 0.015:
                cluster.append(lv)
            else:
                clusters.append(cluster)
                cluster = [lv]
        clusters.append(cluster)
        res = [{'level': round(sum(c)/len(c), 2), 'strength': len(c)}
               for c in clusters
               if sum(c)/len(c) > current_price * 1.005]
        return sorted(res, key=lambda x: x['level'])[:num_levels]

    def find_support_levels(self, df, current_price, num_levels=5):
        window     = 5
        swing_lows = []
        for src_days in [180, 252]:
            lows = df.tail(src_days)['Low'].values
            for i in range(window, len(lows) - window):
                if (lows[i] < min(lows[i-window:i]) and
                        lows[i] < min(lows[i+1:i+window+1])):
                    swing_lows.append(lows[i])
        # 52-week low
        low_52w = df['Low'].tail(252).min()
        if low_52w < current_price * 0.995:
            swing_lows.append(low_52w)
        # Round-number levels below price
        magnitude = 10 ** (len(str(int(current_price))) - 2)
        step      = magnitude * 5
        level     = current_price
        for _ in range(20):
            level -= step
            if level >= current_price * 0.70 and level > 0:
                swing_lows.append(level)
        if not swing_lows:
            return []
        swing_lows = sorted(set([round(l, 2) for l in swing_lows]))
        clusters, cluster = [], [swing_lows[0]]
        for lv in swing_lows[1:]:
            if (lv - cluster[-1]) / cluster[-1] < 0.015:
                cluster.append(lv)
            else:
                clusters.append(cluster)
                cluster = [lv]
        clusters.append(cluster)
        sup = [{'level': round(sum(c)/len(c), 2), 'strength': len(c)}
               for c in clusters
               if sum(c)/len(c) < current_price * 0.995]
        return sorted(sup, key=lambda x: x['level'], reverse=True)[:num_levels]

    # =========================================================================
    #  DYNAMIC TARGETS  (ATR floor, promotion logic)
    # =========================================================================
    def calculate_dynamic_targets(self, current_price, resistance_levels,
                                   support_levels, target_price, atr):
        valid      = [r['level'] for r in resistance_levels
                      if r['level'] > current_price * 1.005]
        min_target = current_price + (atr * 2)
        if len(valid) >= 2:
            t1, t2        = valid[0], valid[1]
            target_status = "Real S/R Levels"
        elif len(valid) == 1:
            t1 = valid[0]
            t2 = (round(target_price, 2)
                  if target_price and target_price > t1 * 1.01
                  else round(t1 * 1.04, 2))
            target_status = "Partial Real Levels"
        else:
            t1 = (round(target_price, 2)
                  if target_price and target_price > current_price * 1.005
                  else round(current_price * 1.03, 2))
            t2            = round(t1 * 1.04, 2)
            target_status = "ATH Zone — Projected"
        if t1 < min_target:
            t1            = round(min_target, 2)
            t2            = round(t1 * 1.04, 2)
            target_status += " (ATR Adj)"
        return round(t1, 2), round(t2, 2), 0, target_status

    # =========================================================================
    #  FUNDAMENTAL SCORE
    # =========================================================================
    def get_fundamental_score(self, info):
        score = 0
        pe  = info.get('trailingPE', info.get('forwardPE', 0))
        pb  = info.get('priceToBook', 0)
        peg = info.get('pegRatio', 0)
        if pe  and 0 < pe  < 25:     score += 10
        elif pe  and 25 <= pe  < 35: score += 5
        if pb  and 0 < pb  < 3:      score += 5
        elif pb  and 3 <= pb  < 5:   score += 3
        if peg and 0 < peg < 1:      score += 10
        elif peg and 1 <= peg < 2:   score += 5
        roe = info.get('returnOnEquity', 0)
        roa = info.get('returnOnAssets', 0)
        pm  = info.get('profitMargins', 0)
        if roe and roe > 0.15:   score += 10
        elif roe and roe > 0.10: score += 5
        if roa and roa > 0.05:   score += 5
        elif roa and roa > 0.02: score += 3
        if pm  and pm  > 0.10:   score += 10
        elif pm  and pm  > 0.05: score += 5
        rg = info.get('revenueGrowth', 0)
        eg = info.get('earningsGrowth', 0)
        if rg and rg > 0.15:   score += 10
        elif rg and rg > 0.10: score += 7
        elif rg and rg > 0.05: score += 5
        if eg and eg > 0.15:   score += 10
        elif eg and eg > 0.10: score += 7
        elif eg and eg > 0.05: score += 5
        de = info.get('debtToEquity', 0)
        cr = info.get('currentRatio', 0)
        fc = info.get('freeCashflow', 0)
        if de is not None:
            if de < 50:    score += 10
            elif de < 100: score += 5
        else:
            score += 5
        if cr and cr > 1.5:   score += 10
        elif cr and cr > 1.0: score += 5
        if fc and fc > 0:     score += 5
        return min(score, 100)

    # =========================================================================
    #  MAIN ANALYSIS
    # =========================================================================
    def analyze_stock(self, symbol, name):
        try:
            stock = yf.Ticker(symbol)
            df    = stock.history(period='1y')
            info  = stock.info
            if df.empty or len(df) < 200:
                return None

            current_price = df['Close'].iloc[-1]
            sma_20  = df['Close'].rolling(20).mean().iloc[-1]
            sma_50  = df['Close'].rolling(50).mean().iloc[-1]
            sma_200 = df['Close'].rolling(200).mean().iloc[-1]

            rsi          = self.calculate_rsi(df['Close'])
            macd, signal = self.calculate_macd(df['Close'])
            atr          = self.calculate_atr(df)
            atr_pct      = round((atr / current_price) * 100, 2)
            adx          = self.calculate_adx(df)
            vol_ratio    = self.calculate_volume_ratio(df)

            high_52w = df['High'].tail(252).max()
            low_52w  = df['Low'].tail(252).min()

            # ── Advanced S/R ──────────────────────────────────────
            resistance_levels = self.find_resistance_levels(df, current_price)
            support_levels    = self.find_support_levels(df, current_price)

            nearest_resistance = (resistance_levels[0]['level']
                                  if resistance_levels
                                  else df.tail(60)['High'].quantile(0.90))
            nearest_support    = (support_levels[0]['level']
                                  if support_levels
                                  else df.tail(60)['Low'].quantile(0.10))

            support_dist_pct = round(
                ((current_price - nearest_support) / current_price) * 100, 2)

            # ── Technical Score ───────────────────────────────────
            tech_score = 0
            tech_score += 1 if current_price > sma_20  else -1
            tech_score += 1 if current_price > sma_50  else -1
            tech_score += 2 if current_price > sma_200 else -2
            if rsi < 30:
                tech_score += 2;  rsi_signal = "Oversold"
            elif rsi > 70:
                tech_score -= 2;  rsi_signal = "Overbought"
            else:
                rsi_signal = "Neutral"
            if macd > signal:
                tech_score += 1;  macd_signal = "Bullish"
            else:
                tech_score -= 1;  macd_signal = "Bearish"
            if adx > 25:
                tech_score = min(tech_score + 1, 6)

            # ── Fundamental Data ──────────────────────────────────
            pe_ratio         = info.get('trailingPE', info.get('forwardPE', 0))
            pb_ratio         = info.get('priceToBook', 0)
            peg_ratio        = info.get('pegRatio', 0)
            market_cap       = info.get('marketCap', 0)
            dividend_yield   = info.get('dividendYield', 0)
            roe              = info.get('returnOnEquity', 0)
            roa              = info.get('returnOnAssets', 0)
            profit_margin    = info.get('profitMargins', 0)
            operating_margin = info.get('operatingMargins', 0)
            eps              = info.get('trailingEps', 0)
            revenue_growth   = info.get('revenueGrowth', 0)
            earnings_growth  = info.get('earningsGrowth', 0)
            debt_to_equity   = info.get('debtToEquity', 0)
            current_ratio    = info.get('currentRatio', 0)
            beta             = info.get('beta', 1.0)
            target_price     = info.get('targetMeanPrice', None)
            sector           = info.get('sector', 'N/A')

            analyst_key   = info.get('recommendationKey', 'N/A')
            analyst_map   = {
                'strongBuy': 'Strong Buy', 'buy': 'Buy',
                'hold': 'Hold', 'sell': 'Sell', 'strongSell': 'Strong Sell'
            }
            analyst_label = analyst_map.get(
                analyst_key,
                analyst_key.title() if analyst_key else 'N/A')
            earnings_date = self.get_earnings_date(info)

            fund_score = self.get_fundamental_score(info)

            tech_score_normalized = ((tech_score + 6) / 12) * 100
            combined_score        = (tech_score_normalized * 0.5) + (fund_score * 0.5)

            if combined_score >= 75:
                rating = "⭐⭐⭐⭐⭐ STRONG BUY";  recommendation = "STRONG BUY"
            elif combined_score >= 55:
                rating = "⭐⭐⭐⭐ BUY";           recommendation = "BUY"
            elif combined_score >= 45:
                rating = "⭐⭐⭐ HOLD";            recommendation = "HOLD"
            elif combined_score >= 30:
                rating = "⭐⭐ SELL";              recommendation = "SELL"
            else:
                rating = "⭐ STRONG SELL";         recommendation = "STRONG SELL"

            # ── Beta-adjusted ATR stop ────────────────────────────
            stock_beta = beta if beta else 1.0
            if stock_beta < 0.8:
                atr_multiplier = 1.0;  max_sl_pct = 5.0
            elif stock_beta < 1.2:
                atr_multiplier = 1.2;  max_sl_pct = 7.0
            elif stock_beta < 1.8:
                atr_multiplier = 1.5;  max_sl_pct = 10.0
            else:
                atr_multiplier = 2.0;  max_sl_pct = 12.0

            if recommendation in ["STRONG BUY", "BUY"]:
                atr_stop       = nearest_support - (atr * atr_multiplier)
                min_allowed_sl = current_price * (1 - max_sl_pct / 100)
                stop_loss      = max(atr_stop, min_allowed_sl)
                sl_percentage  = ((current_price - stop_loss) / current_price) * 100
                stop_type      = "ATR Stop" if atr_stop >= min_allowed_sl else "Beta Cap"

                target_1, target_2, targets_hit, target_status = \
                    self.calculate_dynamic_targets(
                        current_price, resistance_levels,
                        support_levels, target_price, atr)
                if target_1 <= current_price * 1.005:
                    recommendation = "HOLD"; rating = "⭐⭐⭐ HOLD"
                upside = ((target_1 - current_price) / current_price) * 100
            else:
                atr_stop       = nearest_resistance + (atr * atr_multiplier)
                max_allowed_sl = current_price * (1 + max_sl_pct / 100)
                stop_loss      = min(atr_stop, max_allowed_sl)
                sl_percentage  = ((stop_loss - current_price) / current_price) * 100
                stop_type      = "ATR Stop" if atr_stop <= max_allowed_sl else "Beta Cap"

                valid_sups = [s['level'] for s in support_levels
                              if s['level'] < current_price * 0.995]
                if len(valid_sups) >= 2:
                    target_1, target_2 = valid_sups[0], valid_sups[1]
                    target_status = "Real S/R Levels"
                elif len(valid_sups) == 1:
                    target_1 = valid_sups[0]
                    target_2 = round(target_1 * 0.96, 2)
                    target_status = "Partial Real Levels"
                else:
                    target_1 = round(current_price * 0.96, 2)
                    target_2 = round(current_price * 0.92, 2)
                    target_status = "Projected"
                targets_hit = 0
                upside      = ((current_price - target_1) / current_price) * 100

            risk        = abs(current_price - stop_loss)
            reward      = abs(target_1 - current_price)
            risk_reward = round(reward / risk, 2) if risk > 0 else 0

            if fund_score >= 80:   quality = "Excellent"
            elif fund_score >= 60: quality = "Good"
            elif fund_score >= 40: quality = "Average"
            else:                  quality = "Poor"

            return {
                'Symbol':          symbol.replace('.NS', ''),
                'Name':            name,
                'Price':           round(current_price, 2),
                'Sector':          sector,
                'RSI':             round(rsi, 2),
                'RSI_Signal':      rsi_signal,
                'MACD':            macd_signal,
                'ADX':             adx,
                'Vol_Ratio':       vol_ratio,
                'SMA_20':          round(sma_20, 2),
                'SMA_50':          round(sma_50, 2),
                'SMA_200':         round(sma_200, 2),
                'Support':         round(nearest_support, 2),
                'Resistance':      round(nearest_resistance, 2),
                'Support_Dist_Pct':support_dist_pct,
                '52W_High':        round(high_52w, 2),
                '52W_Low':         round(low_52w, 2),
                'Tech_Score':      tech_score,
                'Tech_Score_Norm': round(tech_score_normalized, 1),
                'ATR':             atr,
                'ATR_Pct':         atr_pct,
                'ATR_Multiplier':  atr_multiplier,
                'Stop_Type':       stop_type,
                'PE_Ratio':        round(pe_ratio, 2)           if pe_ratio else 0,
                'PB_Ratio':        round(pb_ratio, 2)           if pb_ratio else 0,
                'PEG_Ratio':       round(peg_ratio, 2)          if peg_ratio else 0,
                'ROE':             round(roe * 100, 2)          if roe else 0,
                'ROA':             round(roa * 100, 2)          if roa else 0,
                'Profit_Margin':   round(profit_margin * 100, 2)    if profit_margin else 0,
                'Operating_Margin':round(operating_margin * 100, 2) if operating_margin else 0,
                'EPS':             round(eps, 2)                if eps else 0,
                'Dividend_Yield':  round(dividend_yield * 100, 2)   if dividend_yield else 0,
                'Revenue_Growth':  round(revenue_growth * 100, 2)   if revenue_growth else 0,
                'Earnings_Growth': round(earnings_growth * 100, 2)  if earnings_growth else 0,
                'Debt_to_Equity':  round(debt_to_equity, 2)    if debt_to_equity else 0,
                'Current_Ratio':   round(current_ratio, 2)     if current_ratio else 0,
                'Market_Cap':      round(market_cap / 1e12, 2) if market_cap else 0,
                'Beta':            round(beta, 2)               if beta else 1.0,
                'Fund_Score':      round(fund_score, 1),
                'Quality':         quality,
                'Combined_Score':  round(combined_score, 1),
                'Rating':          rating,
                'Recommendation':  recommendation,
                'Stop_Loss':       round(stop_loss, 2),
                'SL_Percentage':   round(sl_percentage, 2),
                'Target_1':        round(target_1, 2),
                'Target_2':        round(target_2, 2),
                'Target_Price':    round(target_price, 2) if target_price else 0,
                'Upside':          round(upside, 2),
                'Risk_Reward':     risk_reward,
                'Targets_Hit':     targets_hit,
                'Target_Status':   target_status,
                'Analyst':         analyst_label,
                'Earnings_Date':   earnings_date,
            }
        except Exception:
            return None

    # =========================================================================
    #  ANALYZE ALL
    # =========================================================================
    def analyze_all_stocks(self):
        print(f"🔍 Analyzing {len(self.nifty100_stocks)} NIFTY 100 stocks...")
        print("⏳ ~3-4 minutes...\n")
        for idx, (symbol, name) in enumerate(self.nifty100_stocks.items(), 1):
            result = self.analyze_stock(symbol, name)
            if result:
                self.results.append(result)
            if idx % 10 == 0:
                print(f"  [{idx}/{len(self.nifty100_stocks)}] processed")
        print(f"\n✅ {len(self.results)} stocks analyzed\n")

    # =========================================================================
    #  TOP RECOMMENDATIONS  (same quality filter as US script)
    # =========================================================================
    def get_top_recommendations(self):
        df = pd.DataFrame(self.results)
        all_buys = df[df['Recommendation'].isin(['STRONG BUY', 'BUY'])]
        f1 = all_buys[all_buys['Upside'] > 0.5]
        f2 = f1[f1['Risk_Reward'] >= 0.5]
        f3 = f2[f2['Target_1'] > f2['Price']]
        top_buys = f3.nlargest(20, 'Combined_Score')

        all_sells = df[df['Recommendation'].isin(['STRONG SELL', 'SELL'])]
        s1 = all_sells[all_sells['Upside'] > 0.5]
        s2 = s1[s1['Risk_Reward'] >= 0.5]
        s3 = s2[s2['Target_1'] < s2['Price']]
        top_sells = s3.nsmallest(20, 'Combined_Score')
        return top_buys, top_sells

    # =========================================================================
    #  HTML — Aurora Glass v2 (full-width, mobile-responsive, index strip)
    # =========================================================================
    def generate_html(self):
        df = pd.DataFrame(self.results)
        top_buys, top_sells = self.get_top_recommendations()

        now         = self.get_ist_time()
        idx_data    = self.fetch_index_data()
        time_of_day = "Morning" if now.hour < 12 else "Evening"
        next_update = "4:30 PM" if now.hour < 12 else "9:30 AM (Next Day)"

        strong_buy_count  = len(df[df['Recommendation'] == 'STRONG BUY'])
        buy_count         = len(df[df['Recommendation'] == 'BUY'])
        hold_count        = len(df[df['Recommendation'] == 'HOLD'])
        sell_count        = len(df[df['Recommendation'] == 'SELL'])
        strong_sell_count = len(df[df['Recommendation'] == 'STRONG SELL'])

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>NIFTY 100 Market Influencers — {time_of_day} Report</title>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg:#050f0e; --bg2:#071412; --card:#0a1c1a; --card2:#0d2220;
    --accent:#4dd0c4; --accent2:#2dd4bf;
    --green:#22c55e; --red:#ef4444; --blue:#60a5fa;
    --gold:#f59e0b; --teal:#4dd0c4; --purple:#a78bfa;
    --text:#b2dfdb; --text2:#e0f2f1;
    --sym:#80cbc4; --t2c:#4dd0c4; --muted:#4a7a78;
    --border:#0d3330; --border2:#134540;
  }}
  *, *::before, *::after {{ margin:0; padding:0; box-sizing:border-box; }}

  body {{
    background:var(--bg); color:var(--text);
    font-family:'Plus Jakarta Sans',sans-serif;
    font-size:13px; line-height:1.4;
    background-image:
      radial-gradient(ellipse at 0% 0%,   rgba(77,208,196,0.07) 0%,transparent 50%),
      radial-gradient(ellipse at 100% 100%,rgba(45,212,191,0.04) 0%,transparent 40%);
  }}

  /* ── HEADER ── */
  header {{ background:linear-gradient(180deg,#071412,var(--bg2)); border-bottom:2px solid var(--accent); }}
  .h-top {{
    width:100%; display:flex; align-items:center;
    justify-content:space-between; padding:12px 16px;
    gap:12px; flex-wrap:wrap;
  }}
  .brand {{ display:flex; align-items:center; gap:10px; min-width:0; }}
  .brand-icon {{
    width:36px; height:36px; flex-shrink:0;
    background:linear-gradient(135deg,var(--accent),var(--gold));
    border-radius:8px; display:flex; align-items:center;
    justify-content:center; font-size:17px;
  }}
  .brand-t {{ font-size:clamp(12px,1.8vw,17px); font-weight:800; color:var(--text2); white-space:nowrap; }}
  .brand-s {{ font-size:9px; color:var(--muted); letter-spacing:1px; text-transform:uppercase; margin-top:2px; }}
  .h-meta {{ display:flex; flex-wrap:wrap; gap:0; align-items:center; }}
  .hm {{ padding:6px 14px; border-left:1px solid var(--border2); text-align:right; }}
  .hm-l {{ font-size:8px; color:var(--muted); letter-spacing:2px; text-transform:uppercase; }}
  .hm-v {{ font-family:'JetBrains Mono',monospace; font-size:11px; font-weight:600; margin-top:1px; }}

  /* ── INDEX STRIP ── */
  .idx-strip {{
    display:flex; align-items:center;
    background:rgba(0,0,0,0.35); border:1px solid var(--border2);
    border-radius:8px; padding:4px 0; margin:0 8px;
  }}
  .idx-item {{ display:flex; align-items:center; gap:8px; padding:6px 16px; }}
  .idx-name  {{ font-size:9px; font-weight:800; letter-spacing:2px; color:var(--muted); text-transform:uppercase; }}
  .idx-price {{ font-family:'JetBrains Mono',monospace; font-size:13px; font-weight:700; color:var(--text2); }}
  .idx-chg   {{ font-family:'JetBrains Mono',monospace; font-size:11px; font-weight:700; color:var(--muted); }}
  .idx-chg.up {{ color:var(--green); }}
  .idx-chg.dn {{ color:var(--red); }}
  .idx-sep   {{ width:1px; height:30px; background:var(--border2); }}

  /* ── LIVE CLOCK ── */
  .live-clock-wrap {{
    display:flex; flex-direction:column; align-items:center;
    padding:6px 16px; border-left:1px solid var(--border2); min-width:130px;
  }}
  .lc-label {{ font-size:8px; color:var(--muted); letter-spacing:2px; text-transform:uppercase; }}
  .lc-time  {{ font-family:'JetBrains Mono',monospace; font-size:16px; font-weight:700; color:var(--green); letter-spacing:2px; margin-top:2px; }}
  .lc-date  {{ font-family:'JetBrains Mono',monospace; font-size:9px; color:var(--muted); margin-top:1px; }}
  .lc-last  {{ font-size:8px; color:var(--accent2); margin-top:3px; letter-spacing:0.3px; white-space:nowrap; }}

  /* ── TICKER TAPE ── */
  .ticker {{ background:#020808; border-bottom:1px solid var(--border); }}
  .ticker-inner {{ display:flex; padding:0 16px; overflow-x:auto; scrollbar-width:none; }}
  .ticker-inner::-webkit-scrollbar {{ display:none; }}
  .ti {{ display:flex; gap:5px; align-items:center; padding:5px 10px; border-right:1px solid var(--border); font-family:'JetBrains Mono',monospace; font-size:10px; white-space:nowrap; }}
  .ti-s {{ color:var(--accent2); font-weight:700; }}
  .ti-p {{ color:var(--text2); }}
  .ti-u {{ color:var(--green); }}
  .ti-d {{ color:var(--red); }}

  /* ── KPI BAND ── */
  .kpi-band {{ background:var(--card); border-bottom:1px solid var(--border2); }}
  .kpi-inner {{ display:grid; grid-template-columns:repeat(5,1fr); width:100%; }}
  .kc {{ padding:12px 10px; border-right:1px solid var(--border); text-align:center; }}
  .kc:last-child {{ border-right:none; }}
  .kn {{ font-size:clamp(20px,4vw,30px); font-weight:800; line-height:1; }}
  .kl {{ font-size:8px; letter-spacing:1.5px; text-transform:uppercase; color:var(--muted); margin-top:3px; }}
  .kbar {{ height:2px; border-radius:1px; margin:3px auto 0; width:32px; }}

  /* ── MAIN ── */
  .main {{ width:100%; padding:12px 16px; }}

  /* ── SECTION HEADER ── */
  .sh {{ display:flex; align-items:center; gap:10px; margin-bottom:10px; flex-wrap:wrap; }}
  .sh-icon {{ width:28px; height:28px; border-radius:6px; display:flex; align-items:center; justify-content:center; font-size:13px; flex-shrink:0; }}
  .shi-buy  {{ background:rgba(34,197,94,0.15); }}
  .shi-sell {{ background:rgba(239,68,68,0.15); }}
  .sh-title {{ font-size:15px; font-weight:800; color:var(--text2); }}
  .sh-divider {{ flex:1; height:1px; background:var(--border); min-width:10px; }}
  .sh-count {{ font-size:9px; color:var(--muted); white-space:nowrap; }}

  /* ── TABLE ── */
  .tbl-wrap {{
    width:100%; overflow-x:auto;
    border:1px solid var(--border2); border-radius:8px;
    margin-bottom:20px; background:var(--card);
    box-shadow:0 4px 24px rgba(0,0,0,0.3);
    -webkit-overflow-scrolling:touch;
  }}
  table {{ width:100%; border-collapse:collapse; min-width:1200px; }}
  th {{
    font-size:8px; font-weight:700; letter-spacing:1.5px;
    text-transform:uppercase; color:var(--teal);
    padding:8px 9px; background:var(--card2);
    border-bottom:1px solid var(--border2); text-align:left; white-space:nowrap;
  }}
  td {{ padding:8px 9px; border-bottom:1px solid var(--border); vertical-align:middle; white-space:nowrap; }}
  tr:hover td {{ background:rgba(77,208,196,0.05); }}
  tr:nth-child(even) td {{ background:rgba(0,0,0,0.15); }}
  tr:last-child td {{ border-bottom:none; }}

  /* ── CELL COMPONENTS ── */
  .sn  {{ font-size:13px; font-weight:700; color:var(--text2); }}
  .ss  {{ font-family:'JetBrains Mono',monospace; font-size:9px; font-weight:600; color:var(--sym); letter-spacing:1px; margin-top:2px; }}
  .sec {{ font-size:8px; color:var(--muted); margin-top:2px; max-width:120px; overflow:hidden; text-overflow:ellipsis; }}
  .pv  {{ font-family:'JetBrains Mono',monospace; font-size:13px; font-weight:600; color:var(--gold); }}
  .rt  {{ display:inline-block; font-size:8px; font-weight:700; padding:3px 7px; border-radius:3px; white-space:nowrap; letter-spacing:0.5px; }}
  .rt-sb {{ background:rgba(34,197,94,0.15);  color:#4ade80; border:1px solid rgba(34,197,94,0.3); }}
  .rt-b  {{ background:rgba(77,208,196,0.15); color:#4dd0c4; border:1px solid rgba(77,208,196,0.3); }}
  .rt-s  {{ background:rgba(239,68,68,0.15);  color:#f87171; border:1px solid rgba(239,68,68,0.3); }}
  .rt-ss {{ background:rgba(239,68,68,0.22);  color:#fca5a5; border:1px solid rgba(239,68,68,0.4); }}
  .scn {{ font-size:20px; font-weight:800; }}
  .scb {{ height:3px; border-radius:2px; margin-top:3px; width:36px; }}
  .up {{ color:#4ade80; font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; }}
  .dn {{ color:#f87171; font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; }}
  .t1 {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; color:var(--text2); }}
  .t2 {{ font-size:9px; color:var(--t2c); margin-top:1px; }}
  .sl1 {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; color:#f87171; }}
  .sl2 {{ font-size:9px; color:var(--muted); margin-top:1px; }}
  .rv  {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; }}
  .rsb {{ font-size:8px; color:var(--muted); }}
  .rrv {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; }}
  .qb    {{ font-size:8px; font-weight:700; padding:2px 6px; border-radius:3px; }}
  .qb-ex {{ background:rgba(34,197,94,0.15);  color:#4ade80; }}
  .qb-gd {{ background:rgba(77,208,196,0.15); color:#4dd0c4; }}
  .qb-av {{ background:rgba(245,158,11,0.15); color:#fbbf24; }}
  .qb-po {{ background:rgba(239,68,68,0.15);  color:#f87171; }}
  .ts {{ font-size:7px; font-weight:700; padding:2px 5px; border-radius:3px; letter-spacing:0.5px; display:inline-block; margin-bottom:2px; }}
  .ts-real    {{ background:rgba(34,197,94,0.15);  color:#4ade80; }}
  .ts-partial {{ background:rgba(245,158,11,0.15); color:#fbbf24; }}
  .ts-ath     {{ background:rgba(77,208,196,0.15); color:#4dd0c4; }}
  .sb {{ font-size:7px; font-weight:700; padding:2px 5px; border-radius:3px; display:inline-block; margin-top:2px; }}
  .sb-atr  {{ background:rgba(34,197,94,0.15);  color:#4ade80; }}
  .sb-beta {{ background:rgba(245,158,11,0.15); color:#fbbf24; }}
  .ab {{ font-size:8px; font-weight:700; padding:2px 6px; border-radius:3px; white-space:nowrap; }}
  .ab-sb {{ background:rgba(34,197,94,0.15);  color:#4ade80; }}
  .ab-b  {{ background:rgba(77,208,196,0.15); color:#4dd0c4; }}
  .ab-h  {{ background:rgba(74,122,120,0.2);  color:#80cbc4; }}
  .ab-s  {{ background:rgba(239,68,68,0.15);  color:#f87171; }}
  .adx-strong {{ color:#4ade80; font-weight:700; }}
  .adx-mid    {{ color:#fbbf24; font-weight:600; }}
  .adx-weak   {{ color:#4a7a78; }}
  .vol-high {{ color:#4ade80; font-weight:700; }}
  .vol-norm {{ color:var(--text); }}
  .vol-low  {{ color:#4a7a78; }}
  .earn {{ font-size:9px; color:var(--teal); font-family:'JetBrains Mono',monospace; }}
  .sdist-close {{ color:#4ade80;  font-size:11px; font-weight:600; font-family:'JetBrains Mono',monospace; }}
  .sdist-mid   {{ color:#fbbf24;  font-size:11px; font-weight:600; font-family:'JetBrains Mono',monospace; }}
  .sdist-far   {{ color:#f87171;  font-size:11px; font-weight:600; font-family:'JetBrains Mono',monospace; }}

  /* ── DISCLAIMER ── */
  .disc {{
    background:var(--card); border:1px solid var(--border2);
    border-left:3px solid var(--accent); padding:12px 16px;
    margin:16px 0; font-size:11px; color:var(--muted); line-height:1.7;
  }}
  .disc strong {{ color:#f87171; }}

  /* ── FOOTER ── */
  footer {{
    background:linear-gradient(90deg,var(--bg2),#071412,var(--bg2));
    border-top:2px solid var(--accent); text-align:center;
    padding:14px; font-size:10px; color:var(--muted); letter-spacing:1px;
  }}
  footer strong {{ color:var(--accent2); }}

  /* ── MOBILE ── */
  @media(max-width:900px) {{
    .kpi-inner {{ grid-template-columns:repeat(3,1fr); }}
    .hm:nth-child(n+4) {{ display:none; }}
    .idx-strip {{ display:none; }}
  }}
  @media(max-width:600px) {{
    .kpi-inner {{ grid-template-columns:repeat(2,1fr); }}
    .brand-t {{ font-size:12px; }}
    .hm:nth-child(n+3) {{ display:none; }}
    .main {{ padding:8px; }}
    .h-top {{ padding:10px 10px; }}
    th {{ font-size:7px; padding:6px 7px; letter-spacing:0.5px; }}
    td {{ padding:7px 7px; }}
    .sn {{ font-size:12px; }}
    .kn {{ font-size:18px; }}
    .kl {{ font-size:7px; }}
    .live-clock-wrap {{ display:none; }}
  }}
</style>
</head>
<body>

<!-- HEADER -->
<header>
  <div class="h-top">
    <div class="brand">
      <div class="brand-icon">💎</div>
      <div>
        <div class="brand-t">NIFTY 100 Market Influencers · NSE &amp; BSE</div>
        <div class="brand-s">12M S/R · ATR Stops · Tech &amp; Fundamental v2</div>
      </div>
    </div>

    <!-- ── INDEX STRIP ── -->
    <div class="idx-strip">
      <div class="idx-item">
        <span class="idx-name">SENSEX</span>
        <span class="idx-price">{idx_data['SENSEX']['price']}</span>
        <span class="idx-chg {idx_data['SENSEX']['cls']}">{idx_data['SENSEX']['chg']}</span>
      </div>
      <div class="idx-sep"></div>
      <div class="idx-item">
        <span class="idx-name">NIFTY</span>
        <span class="idx-price">{idx_data['NIFTY 50']['price']}</span>
        <span class="idx-chg {idx_data['NIFTY 50']['cls']}">{idx_data['NIFTY 50']['chg']}</span>
      </div>
      <div class="idx-sep"></div>
      <div class="idx-item">
        <span class="idx-name">BANK NIFTY</span>
        <span class="idx-price">{idx_data['BANK NIFTY']['price']}</span>
        <span class="idx-chg {idx_data['BANK NIFTY']['cls']}">{idx_data['BANK NIFTY']['chg']}</span>
      </div>
    </div>

    <div class="h-meta">
      <div class="hm"><div class="hm-l">Date</div><div class="hm-v" style="color:var(--gold)">{now.strftime('%d %b %Y')}</div></div>
      <div class="live-clock-wrap">
        <div class="lc-label">TIME</div>
        <div class="lc-time" id="liveClock">--:-- --</div>
        <div class="lc-date" id="liveDate">{now.strftime('%d %b %Y')}</div>
        <div class="lc-last">Report: {now.strftime('%d %b %Y %I:%M %p')} IST</div>
      </div>
      <div class="hm"><div class="hm-l">Session</div><div class="hm-v" style="color:var(--green)">▲ {time_of_day.upper()}</div></div>
      <div class="hm"><div class="hm-l">Next Update</div><div class="hm-v" style="color:var(--accent2)">{next_update}</div></div>
    </div>
  </div>

  <!-- TICKER TAPE -->
  <div class="ticker"><div class="ticker-inner">
"""
        for t in self.results[:8]:
            pct  = ((t['Price'] - t['SMA_20']) / t['SMA_20']) * 100
            cls  = "ti-u" if pct >= 0 else "ti-d"
            sign = "+" if pct >= 0 else ""
            html += (f'<div class="ti">'
                     f'<span class="ti-s">{t["Symbol"]}</span>'
                     f'<span class="ti-p">₹{t["Price"]:,.2f}</span>'
                     f'<span class="{cls}">{sign}{pct:.1f}%</span>'
                     f'</div>')

        html += f"""  </div></div>
</header>

<!-- KPI BAND -->
<div class="kpi-band">
  <div class="kpi-inner">
    <div class="kc"><div class="kn" style="color:var(--accent2)">{len(self.results)}</div><div class="kl">Analyzed</div><div class="kbar" style="background:var(--accent)"></div></div>
    <div class="kc"><div class="kn" style="color:var(--green)">{strong_buy_count}</div><div class="kl">Strong Buy</div><div class="kbar" style="background:var(--green)"></div></div>
    <div class="kc"><div class="kn" style="color:var(--teal)">{buy_count}</div><div class="kl">Buy</div><div class="kbar" style="background:var(--teal)"></div></div>
    <div class="kc"><div class="kn" style="color:var(--red)">{sell_count + strong_sell_count}</div><div class="kl">Sell</div><div class="kbar" style="background:var(--red)"></div></div>
    <div class="kc"><div class="kn" style="color:var(--blue)">{hold_count}</div><div class="kl">Hold</div><div class="kbar" style="background:var(--blue)"></div></div>
  </div>
</div>

<!-- MAIN -->
<div class="main">
"""

        # ── helpers ──────────────────────────────────────────────────────────
        def analyst_badge(label):
            m = {'Strong Buy': 'ab-sb', 'Buy': 'ab-b',
                 'Hold': 'ab-h', 'Sell': 'ab-s', 'Strong Sell': 'ab-s'}
            return f'<span class="ab {m.get(label, "ab-h")}">{label}</span>'

        def adx_cell(v):
            if v >= 30:   cls, lbl = "adx-strong", "Strong"
            elif v >= 20: cls, lbl = "adx-mid",    "Moderate"
            else:         cls, lbl = "adx-weak",   "Weak"
            return f'<div class="rv {cls}">{v:.0f}</div><div class="rsb">{lbl}</div>'

        def vol_cell(v):
            cls = "vol-high" if v >= 1.5 else ("vol-low" if v < 0.7 else "vol-norm")
            lbl = "High Vol" if v >= 1.5 else ("Low Vol" if v < 0.7 else "Avg Vol")
            return f'<div class="rv {cls}">{v:.1f}×</div><div class="rsb">{lbl}</div>'

        def sdist_cell(v):
            cls = "sdist-close" if v <= 3 else ("sdist-mid" if v <= 8 else "sdist-far")
            return f'<span class="{cls}">{v:.1f}%</span>'

        def target_badge(ts, th):
            if 'ATH' in ts:      return 'ts-ath',     '🚀 ATH Zone'
            elif 'Partial' in ts: return 'ts-partial', '⚡ Partial S/R'
            else:                 return 'ts-real',    '📍 Real S/R'

        # ── BUY TABLE ─────────────────────────────────────────────────────────
        if not top_buys.empty:
            html += """  <div class="sh">
    <div class="sh-icon shi-buy">▲</div>
    <span class="sh-title">Top 20 Buy Recommendations</span>
    <div class="sh-divider"></div>
    <span class="sh-count">12M S/R · ATR Stop · Sector · Vol · ADX · Earnings</span>
  </div>
  <div class="tbl-wrap"><table>
    <thead><tr>
      <th>#</th><th>Stock / Sector</th><th>Price</th><th>Rating</th>
      <th>Score</th><th>Upside</th><th>Target (S/R)</th><th>Stop Loss</th>
      <th>ATR</th><th>Sup Dist</th><th>RSI</th><th>ADX</th><th>Vol/Avg</th>
      <th>R:R</th><th>52W Hi%</th><th>Beta</th><th>P/E</th><th>Div%</th>
      <th>Analyst</th><th>Earnings</th><th>Quality</th>
    </tr></thead><tbody>
"""
            for i, (_, row) in enumerate(top_buys.iterrows(), 1):
                rtag  = "rt-sb" if row['Recommendation'] == "STRONG BUY" else "rt-b"
                sc_c  = "#4ade80" if row['Combined_Score'] >= 75 else ("#4dd0c4" if row['Combined_Score'] >= 55 else "#fbbf24")
                sc_b  = "#22c55e" if row['Combined_Score'] >= 75 else ("#14b8a6" if row['Combined_Score'] >= 55 else "#f59e0b")
                upcls = "up" if row['Upside'] >= 0 else "dn"
                rsic  = "#f87171" if row['RSI'] > 70 else ("#4ade80" if row['RSI'] < 30 else "#60a5fa")
                w52   = ((row['Price'] - row['52W_High']) / row['52W_High']) * 100
                w52c  = "#f87171" if w52 >= -5 else ("#d4a85a" if w52 >= -20 else "#4ade80")
                betac = "#f87171" if row['Beta'] > 1.5 else ("#fbbf24" if row['Beta'] > 1.0 else "#4ade80")
                rr    = row['Risk_Reward']
                rrc   = "#4ade80" if rr >= 2 else ("#4dd0c4" if rr >= 1 else "#f87171")
                pe    = f"{row['PE_Ratio']:.1f}" if row['PE_Ratio'] > 0 else "N/A"
                pec   = "#4a7a78" if row['PE_Ratio'] <= 0 else ("#4ade80" if row['PE_Ratio'] < 25 else ("#fbbf24" if row['PE_Ratio'] < 40 else "#f87171"))
                div   = f"{row['Dividend_Yield']:.2f}%" if row['Dividend_Yield'] > 0 else "—"
                divc  = "#4ade80" if row['Dividend_Yield'] > 0 else "#4a7a78"
                qcls  = {"Excellent": "qb-ex", "Good": "qb-gd", "Average": "qb-av", "Poor": "qb-po"}.get(row['Quality'], "qb-av")
                tbcls, tbtxt = target_badge(row.get('Target_Status', ''), row.get('Targets_Hit', 0))
                st    = row.get('Stop_Type', 'ATR Stop')
                scls  = "sb-atr" if st == "ATR Stop" else "sb-beta"
                slbl  = f"{'📐' if st == 'ATR Stop' else '🔒'} {st}"
                sec   = row.get('Sector', 'N/A')
                ed    = row.get('Earnings_Date', 'N/A')
                html += f"""      <tr>
        <td style="color:#4a7a78;font-size:11px">{i}</td>
        <td>
          <div class="sn">{row['Name']}</div>
          <div class="ss">{row['Symbol']}</div>
          <div class="sec">{sec}</div>
        </td>
        <td><div class="pv">₹{row['Price']:,.2f}</div></td>
        <td><span class="rt {rtag}">{row['Rating']}</span></td>
        <td>
          <div class="scn" style="color:{sc_c}">{row['Combined_Score']:.0f}</div>
          <div class="scb" style="background:{sc_b}"></div>
        </td>
        <td class="{upcls}">{row['Upside']:+.1f}%</td>
        <td>
          <span class="ts {tbcls}">{tbtxt}</span>
          <div class="t1">₹{row['Target_1']:,.2f}</div>
          <div class="t2">T2: ₹{row['Target_2']:,.2f}</div>
        </td>
        <td>
          <div class="sl1">₹{row['Stop_Loss']:,.2f}</div>
          <div class="sl2">-{row['SL_Percentage']:.1f}%</div>
          <span class="sb {scls}">{slbl}</span>
        </td>
        <td>
          <div style="font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:600;color:var(--teal)">₹{row['ATR']:,.2f}</div>
          <div style="font-size:8px;color:var(--muted)">{row['ATR_Pct']:.1f}% · {row['ATR_Multiplier']}×</div>
        </td>
        <td>{sdist_cell(row.get('Support_Dist_Pct', 0))}</td>
        <td>
          <div class="rv" style="color:{rsic}">{row['RSI']:.0f}</div>
          <div class="rsb">{row['RSI_Signal']}</div>
        </td>
        <td>{adx_cell(row.get('ADX', 0))}</td>
        <td>{vol_cell(row.get('Vol_Ratio', 1.0))}</td>
        <td class="rrv" style="color:{rrc}">{rr:.1f}×</td>
        <td style="color:{w52c};font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:600">{w52:+.1f}%</td>
        <td style="color:{betac};font-size:11px">{row['Beta']:.2f}</td>
        <td style="color:{pec};font-size:11px">{pe}</td>
        <td style="color:{divc};font-size:11px">{div}</td>
        <td>{analyst_badge(row.get('Analyst', 'N/A'))}</td>
        <td><div class="earn">{ed}</div></td>
        <td><span class="qb {qcls}">{row['Quality']}</span></td>
      </tr>
"""
            html += "    </tbody></table></div>\n"

        # ── SELL TABLE ────────────────────────────────────────────────────────
        if not top_sells.empty:
            html += """  <div class="sh">
    <div class="sh-icon shi-sell">▼</div>
    <span class="sh-title">Top 20 Sell Recommendations</span>
    <div class="sh-divider"></div>
    <span class="sh-count">12M S/R · ATR Stop · Sector · Vol · ADX · Earnings</span>
  </div>
  <div class="tbl-wrap"><table>
    <thead><tr>
      <th>#</th><th>Stock / Sector</th><th>Price</th><th>Rating</th>
      <th>Score</th><th>RSI</th><th>MACD</th><th>ADX</th>
      <th>Downside</th><th>Target (S/R)</th><th>Stop Loss</th>
      <th>ATR</th><th>Vol/Avg</th><th>R:R</th>
      <th>Beta</th><th>P/E</th><th>Analyst</th><th>Earnings</th><th>Quality</th>
    </tr></thead><tbody>
"""
            for i, (_, row) in enumerate(top_sells.iterrows(), 1):
                rtag  = "rt-ss" if row['Recommendation'] == "STRONG SELL" else "rt-s"
                rsic  = "#f87171" if row['RSI'] > 70 else ("#4ade80" if row['RSI'] < 30 else "#fbbf24")
                mcdcl = "#f87171" if row['MACD'] == "Bearish" else "#4ade80"
                dncls = "dn" if row['Upside'] >= 0 else "up"
                rr    = row['Risk_Reward']
                rrc   = "#4ade80" if rr >= 2 else ("#fbbf24" if rr >= 1 else "#f87171")
                betac = "#f87171" if row['Beta'] > 1.5 else ("#fbbf24" if row['Beta'] > 1.0 else "#4ade80")
                pe    = f"{row['PE_Ratio']:.1f}" if row['PE_Ratio'] > 0 else "N/A"
                pec   = "#4a7a78" if row['PE_Ratio'] <= 0 else ("#f87171" if row['PE_Ratio'] > 40 else ("#fbbf24" if row['PE_Ratio'] > 25 else "#4ade80"))
                qcls  = {"Excellent": "qb-ex", "Good": "qb-gd", "Average": "qb-av", "Poor": "qb-po"}.get(row['Quality'], "qb-av")
                tbcls, tbtxt = target_badge(row.get('Target_Status', ''), 0)
                st    = row.get('Stop_Type', 'ATR Stop')
                scls  = "sb-atr" if st == "ATR Stop" else "sb-beta"
                slbl  = f"{'📐' if st == 'ATR Stop' else '🔒'} {st}"
                sec   = row.get('Sector', 'N/A')
                ed    = row.get('Earnings_Date', 'N/A')
                html += f"""      <tr>
        <td style="color:#4a7a78;font-size:11px">{i}</td>
        <td>
          <div class="sn">{row['Name']}</div>
          <div class="ss">{row['Symbol']}</div>
          <div class="sec">{sec}</div>
        </td>
        <td><div class="pv">₹{row['Price']:,.2f}</div></td>
        <td><span class="rt {rtag}">{row['Rating']}</span></td>
        <td><div class="scn" style="color:#f87171">{row['Combined_Score']:.0f}</div><div class="scb" style="background:#ef4444"></div></td>
        <td><div class="rv" style="color:{rsic}">{row['RSI']:.0f}</div><div class="rsb">{row['RSI_Signal']}</div></td>
        <td style="color:{mcdcl};font-weight:600;font-size:11px">{row['MACD']}</td>
        <td>{adx_cell(row.get('ADX', 0))}</td>
        <td class="{dncls}">{row['Upside']:+.1f}%</td>
        <td>
          <span class="ts {tbcls}">{tbtxt}</span>
          <div class="t1">₹{row['Target_1']:,.2f}</div>
          <div class="t2">T2: ₹{row['Target_2']:,.2f}</div>
        </td>
        <td>
          <div style="font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:600;color:#fbbf24">₹{row['Stop_Loss']:,.2f}</div>
          <div class="sl2">+{row['SL_Percentage']:.1f}%</div>
          <span class="sb {scls}">{slbl}</span>
        </td>
        <td>
          <div style="font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:600;color:var(--teal)">₹{row['ATR']:,.2f}</div>
          <div style="font-size:8px;color:var(--muted)">{row['ATR_Pct']:.1f}% · {row['ATR_Multiplier']}×</div>
        </td>
        <td>{vol_cell(row.get('Vol_Ratio', 1.0))}</td>
        <td class="rrv" style="color:{rrc}">{rr:.1f}×</td>
        <td style="color:{betac};font-size:11px">{row['Beta']:.2f}</td>
        <td style="color:{pec};font-size:11px">{pe}</td>
        <td>{analyst_badge(row.get('Analyst', 'N/A'))}</td>
        <td><div class="earn">{ed}</div></td>
        <td><span class="qb {qcls}">{row['Quality']}</span></td>
      </tr>
"""
            html += "    </tbody></table></div>\n"

        html += f"""  <div class="disc">
    <strong>⚠ DISCLAIMER:</strong> For <strong>EDUCATIONAL PURPOSES ONLY</strong>. Not financial advice.
    Stop losses are ATR-based near real 12-month S/R zones. Targets derived from swing highs/lows,
    52-week extremes and round-number levels. Earnings dates are estimates.
    Always conduct your own research, consult a SEBI-registered financial advisor,
    and never invest more than you can afford to lose.
  </div>
</div>

<footer>
  <strong>NIFTY 100 Market Influencers: NSE &amp; BSE</strong>
  · 12M S/R · ATR Stops · Sector · ADX · Vol · Earnings v2
  · Next Update: <strong>{next_update} IST</strong> · {now.strftime('%d %b %Y')}
</footer>

<script>
function updateClock() {{
  var now  = new Date();
  var ist  = new Date(now.toLocaleString('en-US', {{timeZone: 'Asia/Kolkata'}}));
  var h    = ist.getHours(), m = ist.getMinutes();
  var ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  var pad = function(n) {{ return String(n).padStart(2, '0'); }};
  document.getElementById('liveClock').textContent = pad(h) + ':' + pad(m) + ' ' + ampm + ' IST';
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  document.getElementById('liveDate').textContent =
    pad(ist.getDate()) + ' ' + months[ist.getMonth()] + ' ' + ist.getFullYear();
}}
updateClock();
setInterval(updateClock, 1000);
</script>

</body></html>"""
        return html

    # =========================================================================
    #  SAVE HTML
    # =========================================================================
    def save_html(self, output_file='index.html'):
        html = self.generate_html()
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"✅ HTML report saved: {output_file}")

    # =========================================================================
    #  EMAIL
    # =========================================================================
    def send_email(self, to_email):
        try:
            from_email = os.environ.get('GMAIL_USER')
            password   = os.environ.get('GMAIL_APP_PASSWORD')
            if not from_email or not password:
                print("❌ Set GMAIL_USER and GMAIL_APP_PASSWORD"); return False
            now = self.get_ist_time()
            tod = "Morning" if now.hour < 12 else "Evening"
            msg = MIMEMultipart('alternative')
            msg['From']    = from_email
            msg['To']      = to_email
            msg['Subject'] = f"💎 NIFTY 100 Report v2 — {tod} {now.strftime('%d %b %Y')}"
            msg.attach(MIMEText(self.generate_html(), 'html'))
            srv = smtplib.SMTP('smtp.gmail.com', 587)
            srv.starttls(); srv.login(from_email, password)
            srv.send_message(msg); srv.quit()
            print(f"✅ Email sent to {to_email}"); return True
        except Exception as e:
            print(f"❌ Email error: {e}"); return False

    # =========================================================================
    #  ENTRY
    # =========================================================================
    def generate_complete_report(self, send_email_flag=True, recipient_email=None,
                                  output_file='index.html'):
        now = self.get_ist_time()
        print("=" * 70)
        print("💎 NIFTY 100 ANALYZER v2 — ATR Stops + 12M S/R + ADX + Vol")
        print(f"   {now.strftime('%d %b %Y, %I:%M %p IST')}")
        print("=" * 70)
        self.analyze_all_stocks()
        self.save_html(output_file)
        if send_email_flag and recipient_email:
            self.send_email(recipient_email)
        print("=" * 70); print("✅ DONE"); print("=" * 70)


# =============================================================================
#  RUN
# =============================================================================
if __name__ == "__main__":
    analyzer  = Nifty100CompleteAnalyzer()
    recipient = os.environ.get('RECIPIENT_EMAIL')
    analyzer.generate_complete_report(
        send_email_flag=bool(recipient),
        recipient_email=recipient,
        output_file='index.html'
    )
