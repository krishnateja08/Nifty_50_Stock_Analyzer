"""
NIFTY 100 COMPLETE STOCK ANALYZER — REDESIGNED UI v3
Technical + Fundamental Analysis with Email Delivery + GitHub Pages

UPGRADES from v2:
  - Grouped column headers: STOCK INFO · TRADE SETUP · TECHNICALS · FUNDAMENTALS · META
  - Color-coded column group bands with top-border accents
  - Score displayed as number + animated fill bar
  - Stop type badge anchored directly under SL price
  - Target S/R badge above target price for instant context
  - Auto-scrolling ticker tape (CSS animation, no JS dependency)
  - Live clock with seconds
  - Space Grotesk + IBM Plex Mono typography
  - Vertical group separators for clean scanning
  - All original analysis logic preserved

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
            'TMCV.NS':        'Tata Motors Commercial',
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
        indices = {
            'SENSEX':     '^BSESN',
            'NIFTY 50':   '^NSEI',
            'BANK NIFTY': '^NSEBANK',
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
    #  RESISTANCE & SUPPORT
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
        high_52w = df['High'].tail(252).max()
        if high_52w > current_price * 1.005:
            swing_highs.append(high_52w)
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
        low_52w = df['Low'].tail(252).min()
        if low_52w < current_price * 0.995:
            swing_lows.append(low_52w)
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
    #  DYNAMIC TARGETS
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
                'Symbol':           symbol.replace('.NS', ''),
                'Name':             name,
                'Price':            round(current_price, 2),
                'Sector':           sector,
                'RSI':              round(rsi, 2),
                'RSI_Signal':       rsi_signal,
                'MACD':             macd_signal,
                'ADX':              adx,
                'Vol_Ratio':        vol_ratio,
                'SMA_20':           round(sma_20, 2),
                'SMA_50':           round(sma_50, 2),
                'SMA_200':          round(sma_200, 2),
                'Support':          round(nearest_support, 2),
                'Resistance':       round(nearest_resistance, 2),
                'Support_Dist_Pct': support_dist_pct,
                '52W_High':         round(high_52w, 2),
                '52W_Low':          round(low_52w, 2),
                'Tech_Score':       tech_score,
                'Tech_Score_Norm':  round(tech_score_normalized, 1),
                'ATR':              atr,
                'ATR_Pct':          atr_pct,
                'ATR_Multiplier':   atr_multiplier,
                'Stop_Type':        stop_type,
                'PE_Ratio':         round(pe_ratio, 2)            if pe_ratio else 0,
                'PB_Ratio':         round(pb_ratio, 2)            if pb_ratio else 0,
                'PEG_Ratio':        round(peg_ratio, 2)           if peg_ratio else 0,
                'ROE':              round(roe * 100, 2)           if roe else 0,
                'ROA':              round(roa * 100, 2)           if roa else 0,
                'Profit_Margin':    round(profit_margin * 100, 2)     if profit_margin else 0,
                'Operating_Margin': round(operating_margin * 100, 2)  if operating_margin else 0,
                'EPS':              round(eps, 2)                 if eps else 0,
                'Dividend_Yield':   round(dividend_yield * 100, 2)    if dividend_yield else 0,
                'Revenue_Growth':   round(revenue_growth * 100, 2)    if revenue_growth else 0,
                'Earnings_Growth':  round(earnings_growth * 100, 2)   if earnings_growth else 0,
                'Debt_to_Equity':   round(debt_to_equity, 2)     if debt_to_equity else 0,
                'Current_Ratio':    round(current_ratio, 2)      if current_ratio else 0,
                'Market_Cap':       round(market_cap / 1e12, 2)  if market_cap else 0,
                'Beta':             round(beta, 2)                if beta else 1.0,
                'Fund_Score':       round(fund_score, 1),
                'Quality':          quality,
                'Combined_Score':   round(combined_score, 1),
                'Rating':           rating,
                'Recommendation':   recommendation,
                'Stop_Loss':        round(stop_loss, 2),
                'SL_Percentage':    round(sl_percentage, 2),
                'Target_1':         round(target_1, 2),
                'Target_2':         round(target_2, 2),
                'Target_Price':     round(target_price, 2) if target_price else 0,
                'Upside':           round(upside, 2),
                'Risk_Reward':      risk_reward,
                'Targets_Hit':      targets_hit,
                'Target_Status':    target_status,
                'Analyst':          analyst_label,
                'Earnings_Date':    earnings_date,
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
    #  TOP RECOMMENDATIONS
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
    #  HTML — Redesigned v3: Grouped Columns + IBM Plex Mono + Score Bars
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

        # ── ticker tape items ──────────────────────────────────────────────
        ticker_items = ""
        for t in self.results[:12]:
            pct  = ((t['Price'] - t['SMA_20']) / t['SMA_20']) * 100
            cls  = "tick-up" if pct >= 0 else "tick-dn"
            sign = "+" if pct >= 0 else ""
            ticker_items += (
                f'<span class="tick">'
                f'<span class="tick-sym">{t["Symbol"]}</span>'
                f'<span class="tick-px">₹{t["Price"]:,.2f}</span>'
                f'<span class="{cls}">{sign}{pct:.1f}%</span>'
                f'</span>'
            )
        # duplicate for seamless loop
        ticker_html = ticker_items + ticker_items

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>NIFTY 100 Market Influencers — {time_of_day} Report · {now.strftime('%d %b %Y')}</title>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Syne:wght@700;800&display=swap" rel="stylesheet">
<style>
:root {{
  --bg:       #04080f;
  --bg2:      #060d18;
  --surface:  #0a1628;
  --surface2: #0e1d35;
  --border:   #132240;
  --border2:  #1a2e50;
  --accent:   #00d4ff;
  --accent2:  #0099cc;
  --green:    #00e676;
  --green2:   #00c853;
  --red:      #ff3d57;
  --gold:     #ffab00;
  --purple:   #7c4dff;
  --teal:     #00bcd4;
  --text:     #cdd6f4;
  --text2:    #e8eeff;
  --muted:    #7a95b8;
  --muted2:   #a0bcd8;
}}
*, *::before, *::after {{ margin:0; padding:0; box-sizing:border-box; }}

body {{
  background: var(--bg);
  color: var(--text);
  font-family: 'Space Grotesk', sans-serif;
  font-size: 13px;
  min-height: 100vh;
  background-image:
    radial-gradient(ellipse 80% 40% at 10% 0%, rgba(0,212,255,0.06) 0%, transparent 60%),
    radial-gradient(ellipse 60% 30% at 90% 100%, rgba(124,77,255,0.05) 0%, transparent 50%);
}}

/* ── HEADER ── */
header {{
  background: linear-gradient(180deg, #060d18 0%, #04080f 100%);
  border-bottom: 1px solid var(--border2);
  position: sticky; top: 0; z-index: 100;
  box-shadow: 0 4px 24px rgba(0,0,0,0.6);
}}
.h-top {{
  display: flex; align-items: center;
  justify-content: space-between;
  padding: 10px 20px; gap: 16px; flex-wrap: wrap;
}}
.brand {{ display: flex; align-items: center; gap: 12px; }}
.brand-gem {{
  width: 40px; height: 40px;
  background: linear-gradient(135deg, #00d4ff, #7c4dff);
  border-radius: 10px; display: flex;
  align-items: center; justify-content: center;
  font-size: 20px; flex-shrink: 0;
  box-shadow: 0 0 20px rgba(0,212,255,0.3);
}}
.brand-name {{
  font-family: 'Syne', sans-serif;
  font-size: 17px; font-weight: 800;
  color: var(--text2); letter-spacing: -0.5px;
}}
.brand-sub {{
  font-size: 10px; color: #8aabcc;
  letter-spacing: 2px; text-transform: uppercase; margin-top: 2px; font-weight: 600;
}}

/* INDEX STRIP */
.idx-strip {{
  display: flex; align-items: center;
  background: rgba(0,0,0,0.4);
  border: 1px solid var(--border2);
  border-radius: 10px; overflow: hidden;
}}
.idx-item {{
  display: flex; flex-direction: column;
  align-items: center; padding: 6px 20px;
  border-right: 1px solid var(--border); gap: 2px;
}}
.idx-item:last-child {{ border-right: none; }}
.idx-name  {{ font-size: 9px; font-weight: 700; letter-spacing: 2px; color: #a0bcd8; text-transform: uppercase; }}
.idx-price {{ font-family: 'IBM Plex Mono', monospace; font-size: 14px; font-weight: 600; color: var(--text2); }}
.idx-chg   {{ font-family: 'IBM Plex Mono', monospace; font-size: 10px; font-weight: 600; }}
.idx-chg.up {{ color: var(--green); }}
.idx-chg.dn {{ color: var(--red); }}

/* CLOCK */
.clock-box {{
  display: flex; flex-direction: column; align-items: flex-end; gap: 2px;
}}
.clock-time {{
  font-family: 'IBM Plex Mono', monospace;
  font-size: 18px; font-weight: 600; color: var(--green);
  text-shadow: 0 0 12px rgba(0,230,118,0.4);
}}
.clock-meta {{ font-size: 10px; color: #a0bcd8; letter-spacing: 1px; font-weight: 600; }}
.clock-next {{ font-size: 9px; color: #8aabcc; margin-top: 2px; font-weight: 500; }}

/* TICKER TAPE */
.ticker {{
  background: rgba(0,0,0,0.6);
  border-top: 1px solid var(--border); overflow: hidden;
}}
.ticker-track {{ display: flex; white-space: nowrap; }}
.ticker-inner {{
  display: flex; white-space: nowrap;
  animation: ticker-scroll 50s linear infinite;
  padding: 5px 0;
}}
@keyframes ticker-scroll {{
  0%   {{ transform: translateX(0); }}
  100% {{ transform: translateX(-50%); }}
}}
.tick {{
  display: inline-flex; align-items: center; gap: 6px;
  padding: 0 18px; border-right: 1px solid var(--border);
  font-family: 'IBM Plex Mono', monospace; font-size: 10px;
}}
.tick-sym  {{ color: var(--accent); font-weight: 600; }}
.tick-px   {{ color: var(--text2); }}
.tick-up   {{ color: var(--green); }}
.tick-dn   {{ color: var(--red); }}

/* KPI BAND */
.kpi-band {{
  display: flex; align-items: center;
  background: var(--surface);
  border-bottom: 1px solid var(--border2);
}}
.kpi-item {{
  display: flex; flex-direction: column; align-items: center;
  padding: 14px 24px; border-right: 1px solid var(--border); flex: 1;
}}
.kpi-item:last-child {{ border-right: none; }}
.kpi-num   {{ font-family: 'Syne', sans-serif; font-size: 32px; font-weight: 800; line-height: 1; }}
.kpi-label {{ font-size: 10px; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; color: #a0bcd8; margin-top: 4px; }}
.kpi-bar   {{ height: 2px; width: 40px; border-radius: 1px; margin-top: 6px; }}

/* MAIN */
.main {{ padding: 20px; }}

/* SECTION HEADER */
.section-hdr {{
  display: flex; align-items: center; gap: 12px; margin-bottom: 14px;
}}
.section-pill {{
  display: flex; align-items: center; gap: 8px;
  padding: 6px 16px; border-radius: 100px;
  font-size: 12px; font-weight: 700; letter-spacing: 0.5px;
}}
.pill-buy  {{ background: rgba(0,230,118,0.12); color: var(--green); border: 1px solid rgba(0,230,118,0.25); }}
.pill-sell {{ background: rgba(255,61,87,0.12);  color: var(--red);   border: 1px solid rgba(255,61,87,0.25); }}
.section-line {{ flex: 1; height: 1px; background: var(--border); }}
.section-note {{ font-size: 10px; color: #8aabcc; letter-spacing: 1px; white-space: nowrap; font-weight: 600; }}

/* TABLE WRAPPER */
.tbl-wrap {{
  width: 100%; overflow-x: auto;
  border: 1px solid var(--border2); border-radius: 12px;
  margin-bottom: 28px;
  box-shadow: 0 8px 40px rgba(0,0,0,0.5);
  -webkit-overflow-scrolling: touch;
}}
table {{ width: 100%; border-collapse: collapse; min-width: 1500px; }}

/* GROUP HEADER ROW */
.grp-row th {{
  font-size: 8px; font-weight: 700; letter-spacing: 2px; text-transform: uppercase;
  padding: 7px 10px; text-align: center; border-bottom: 1px solid var(--border);
  white-space: nowrap;
}}
.grp-stock {{ background: rgba(0,188,212,0.18); color: #4de8f8; font-size:9px; }}
.grp-trade {{ background: rgba(0,230,118,0.15); color: #40ff90; font-size:9px; }}
.grp-tech  {{ background: rgba(0,212,255,0.15); color: #40ddff; font-size:9px; }}
.grp-fund  {{ background: rgba(255,171,0,0.15); color: #ffc040; font-size:9px; }}
.grp-meta  {{ background: rgba(124,77,255,0.18); color: #b090ff; font-size:9px; }}

/* COLUMN HEADER ROW */
.col-row th {{
  font-size: 10px; font-weight: 700; letter-spacing: 0.8px; text-transform: uppercase;
  padding: 8px 10px; color: #c8ddf0;
  background: var(--surface2);
  border-bottom: 2px solid var(--border2);
  white-space: nowrap; text-align: left;
}}
.ch-stock {{ border-top: 2px solid #4de8f8; }}
.ch-trade {{ border-top: 2px solid #40ff90; }}
.ch-tech  {{ border-top: 2px solid #40ddff; }}
.ch-fund  {{ border-top: 2px solid #ffc040; }}
.ch-meta  {{ border-top: 2px solid #b090ff; }}

/* Group separator */
.gsep {{ border-left: 2px solid var(--border2) !important; }}

/* DATA ROWS */
td {{
  padding: 10px 10px; border-bottom: 1px solid var(--border);
  vertical-align: middle; white-space: nowrap;
}}
tr:last-child td {{ border-bottom: none; }}
tr:nth-child(even) td {{ background: rgba(255,255,255,0.015); }}
tr:hover td {{ background: rgba(0,212,255,0.04); transition: background 0.15s; }}

/* ── CELL STYLES ── */
.stock-name {{ font-size: 13px; font-weight: 600; color: var(--text2); }}
.stock-sym  {{ font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: #40ddff; font-weight: 700; letter-spacing: 1px; margin-top: 2px; }}
.stock-sec  {{ font-size: 9px; color: #8aabcc; margin-top: 2px; max-width: 130px; overflow: hidden; text-overflow: ellipsis; font-weight: 500; }}
.price-val  {{ font-family: 'IBM Plex Mono', monospace; font-size: 14px; font-weight: 600; color: var(--gold); }}

/* Rating badge */
.badge {{
  display: inline-flex; align-items: center; gap: 4px;
  font-size: 9px; font-weight: 700; padding: 4px 9px;
  border-radius: 6px; letter-spacing: 0.5px; white-space: nowrap;
}}
.badge-sb {{ background: rgba(0,230,118,0.15); color: #00e676; border: 1px solid rgba(0,230,118,0.3); }}
.badge-b  {{ background: rgba(0,212,255,0.15); color: #00d4ff; border: 1px solid rgba(0,212,255,0.3); }}
.badge-h  {{ background: rgba(74,96,128,0.25); color: #8aa0c0; border: 1px solid rgba(74,96,128,0.3); }}
.badge-s  {{ background: rgba(255,61,87,0.15);  color: #ff3d57; border: 1px solid rgba(255,61,87,0.3); }}
.badge-ss {{ background: rgba(255,61,87,0.22);  color: #ff6b7a; border: 1px solid rgba(255,61,87,0.4); }}

/* Score */
.score-wrap {{ display: flex; flex-direction: column; align-items: center; gap: 4px; margin-top: 6px; }}
.score-num  {{ font-family: 'Syne', sans-serif; font-size: 22px; font-weight: 800; line-height: 1; }}
.score-track {{ width: 44px; height: 3px; background: var(--border2); border-radius: 2px; }}
.score-fill  {{ height: 100%; border-radius: 2px; transition: width 0.5s ease; }}

/* Target / Stop */
.target-badge {{
  font-size: 7px; font-weight: 700; padding: 2px 6px;
  border-radius: 4px; letter-spacing: 0.5px;
  display: block; margin-bottom: 3px;
}}
.tb-real    {{ background: rgba(0,230,118,0.15); color: #40ff90; border: 1px solid rgba(0,230,118,0.3); }}
.tb-partial {{ background: rgba(255,171,0,0.15);  color: #ffc040; border: 1px solid rgba(255,171,0,0.3); }}
.tb-ath     {{ background: rgba(0,212,255,0.15);  color: #40ddff; border: 1px solid rgba(0,212,255,0.3); }}
.t1-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 13px; font-weight: 600; color: var(--text2); }}
.t2-val {{ font-size: 10px; color: #8aabcc; margin-top: 2px; font-weight: 600; }}
.sl-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 13px; font-weight: 600; color: var(--red); }}
.sl-pct {{ font-size: 10px; color: #8aabcc; margin-top: 2px; font-weight: 600; }}
.sl-type {{ font-size: 8px; font-weight: 700; padding: 2px 6px; border-radius: 4px; margin-top: 3px; display: inline-block; }}
.slt-atr  {{ background: rgba(0,230,118,0.15); color: #40ff90; }}
.slt-beta {{ background: rgba(255,171,0,0.15);  color: #ffc040; }}

.upside-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 15px; font-weight: 700; }}
.upside-val.up {{ color: var(--green); }}
.upside-val.dn {{ color: var(--red); }}
.rr-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 14px; font-weight: 700; }}
.atr-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 12px; font-weight: 600; color: #4de8f8; }}
.atr-sub {{ font-size: 9px; color: #8aabcc; margin-top: 1px; font-weight: 600; }}

/* Technicals */
.rsi-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 13px; font-weight: 600; }}
.rsi-sig {{ font-size: 9px; color: #8aabcc; margin-top: 1px; font-weight: 600; }}
.adx-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 13px; font-weight: 600; }}
.adx-lbl {{ font-size: 9px; color: #8aabcc; margin-top: 1px; font-weight: 600; }}
.adx-strong {{ color: #40ff90; }}
.adx-mod    {{ color: #ffc040; }}
.adx-weak   {{ color: #a0bcd8; }}
.vol-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 13px; font-weight: 600; }}
.vol-lbl {{ font-size: 9px; color: #8aabcc; margin-top: 1px; font-weight: 600; }}
.vol-high {{ color: #40ff90; }}
.vol-norm {{ color: var(--text); }}
.vol-low  {{ color: #a0bcd8; }}
.sdist-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 13px; font-weight: 600; }}
.sdist-close {{ color: var(--green); }}
.sdist-mid   {{ color: var(--gold); }}
.sdist-far   {{ color: var(--red); }}

/* Fundamentals */
.mono-sm {{ font-family: 'IBM Plex Mono', monospace; font-size: 12px; font-weight: 600; }}

/* Badges */
.qbadge {{ font-size: 8px; font-weight: 700; padding: 3px 8px; border-radius: 5px; }}
.qb-ex {{ background: rgba(0,230,118,0.12); color: #00e676; }}
.qb-gd {{ background: rgba(0,188,212,0.12); color: #00bcd4; }}
.qb-av {{ background: rgba(255,171,0,0.12);  color: #ffab00; }}
.qb-po {{ background: rgba(255,61,87,0.12);  color: #ff3d57; }}

.analyst-badge {{ font-size: 8px; font-weight: 700; padding: 3px 8px; border-radius: 5px; white-space: nowrap; }}
.ab-sb {{ background: rgba(0,230,118,0.12); color: #00e676; }}
.ab-b  {{ background: rgba(0,212,255,0.12); color: #00d4ff; }}
.ab-h  {{ background: rgba(74,96,128,0.20); color: #8aa0c0; }}
.ab-s  {{ background: rgba(255,61,87,0.12); color: #ff3d57; }}

.earn-date {{ font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: #4de8f8; font-weight: 600; }}
.rnum {{ font-size: 11px; color: #a0bcd8; font-weight: 600; }}
.macd-bull {{ color: var(--green); font-weight: 600; font-size: 11px; }}
.macd-bear {{ color: var(--red);   font-weight: 600; font-size: 11px; }}

/* DISCLAIMER */
.disc {{
  background: var(--surface); border: 1px solid var(--border2);
  border-left: 3px solid var(--red);
  padding: 12px 16px; border-radius: 8px;
  font-size: 11px; color: #8aabcc; line-height: 1.8;
  margin: 16px 0;
}}

/* FOOTER */
footer {{
  text-align: center; padding: 16px;
  background: var(--surface); border-top: 1px solid var(--border2);
  font-size: 11px; color: #8aabcc; letter-spacing: 1px;
}}
footer strong {{ color: var(--accent); }}

/* MOBILE */
@media(max-width: 900px) {{
  .idx-strip {{ display: none; }}
  .kpi-item  {{ padding: 10px 12px; }}
  .kpi-num   {{ font-size: 24px; }}
}}
@media(max-width: 600px) {{
  .h-top  {{ padding: 8px 12px; }}
  .main   {{ padding: 10px; }}
  .kpi-band {{ flex-wrap: wrap; }}
  .kpi-item {{ flex: 0 0 50%; border-bottom: 1px solid var(--border); }}
}}
</style>
</head>
<body>

<!-- HEADER -->
<header>
  <div class="h-top">
    <div class="brand">
      <div class="brand-gem">💎</div>
      <div>
        <div class="brand-name">NIFTY 100 Market Influencers · NSE &amp; BSE</div>
        <div class="brand-sub">12M S/R · ATR Stops · Tech &amp; Fundamental v3</div>
      </div>
    </div>

    <div class="idx-strip">
      <div class="idx-item">
        <span class="idx-name">SENSEX</span>
        <span class="idx-price">{idx_data['SENSEX']['price']}</span>
        <span class="idx-chg {idx_data['SENSEX']['cls']}">{idx_data['SENSEX']['chg']}</span>
      </div>
      <div class="idx-item">
        <span class="idx-name">NIFTY 50</span>
        <span class="idx-price">{idx_data['NIFTY 50']['price']}</span>
        <span class="idx-chg {idx_data['NIFTY 50']['cls']}">{idx_data['NIFTY 50']['chg']}</span>
      </div>
      <div class="idx-item">
        <span class="idx-name">BANK NIFTY</span>
        <span class="idx-price">{idx_data['BANK NIFTY']['price']}</span>
        <span class="idx-chg {idx_data['BANK NIFTY']['cls']}">{idx_data['BANK NIFTY']['chg']}</span>
      </div>
    </div>

    <div class="clock-box">
      <div class="clock-time" id="liveClock">--:-- --</div>
      <div class="clock-meta" id="liveDate">{now.strftime('%d %b %Y')} · IST</div>
      <div class="clock-next">Report: {now.strftime('%d %b %Y %I:%M %p')} IST</div>
      <div class="clock-next">Next Update: <strong style="color:var(--accent2)">{next_update}</strong></div>
    </div>
  </div>

  <div class="ticker">
    <div class="ticker-inner">{ticker_html}</div>
  </div>
</header>

<!-- KPI BAND -->
<div class="kpi-band">
  <div class="kpi-item"><div class="kpi-num" style="color:var(--accent)">{len(self.results)}</div><div class="kpi-label">Analyzed</div><div class="kpi-bar" style="background:var(--accent)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:var(--green)">{strong_buy_count}</div><div class="kpi-label">Strong Buy</div><div class="kpi-bar" style="background:var(--green)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:var(--teal)">{buy_count}</div><div class="kpi-label">Buy</div><div class="kpi-bar" style="background:var(--teal)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:var(--red)">{sell_count + strong_sell_count}</div><div class="kpi-label">Sell / Strong Sell</div><div class="kpi-bar" style="background:var(--red)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:#60a5fa">{hold_count}</div><div class="kpi-label">Hold</div><div class="kpi-bar" style="background:#60a5fa"></div></div>
</div>

<div class="main">
"""

        # ── helper functions ──────────────────────────────────────────────────
        def rating_badge(rec, rating_text):
            cls_map = {
                'STRONG BUY':  'badge-sb',
                'BUY':         'badge-b',
                'HOLD':        'badge-h',
                'SELL':        'badge-s',
                'STRONG SELL': 'badge-ss',
            }
            cls = cls_map.get(rec, 'badge-h')
            return f'<span class="badge {cls}">{rating_text}</span>'

        def score_cell(val, color, bar_color):
            pct = min(int(val), 100)
            return (f'<div class="score-wrap">'
                    f'<div class="score-num" style="color:{color}">{val:.0f}</div>'
                    f'<div class="score-track">'
                    f'<div class="score-fill" style="width:{pct}%;background:{bar_color}"></div>'
                    f'</div></div>')

        def target_badge_html(ts):
            if 'ATH' in ts:       return 'tb-ath',     '🚀 ATH Zone'
            elif 'Partial' in ts: return 'tb-partial',  '⚡ Partial S/R'
            else:                  return 'tb-real',    '📍 Real S/R'

        def adx_cell(v):
            if v >= 30:   cls, lbl = 'adx-strong', 'Strong'
            elif v >= 20: cls, lbl = 'adx-mod',    'Moderate'
            else:         cls, lbl = 'adx-weak',   'Weak'
            return (f'<div class="adx-val {cls}">{v:.0f}</div>'
                    f'<div class="adx-lbl">{lbl}</div>')

        def vol_cell(v):
            cls = 'vol-high' if v >= 1.5 else ('vol-low' if v < 0.7 else 'vol-norm')
            lbl = 'High Vol' if v >= 1.5 else ('Low Vol' if v < 0.7 else 'Avg Vol')
            return (f'<div class="vol-val {cls}">{v:.1f}×</div>'
                    f'<div class="vol-lbl">{lbl}</div>')

        def sdist_cell(v):
            cls = 'sdist-close' if v <= 3 else ('sdist-mid' if v <= 8 else 'sdist-far')
            return f'<span class="sdist-val {cls}">{v:.1f}%</span>'

        def analyst_badge(label):
            m = {'Strong Buy': 'ab-sb', 'Buy': 'ab-b',
                 'Hold': 'ab-h', 'Sell': 'ab-s', 'Strong Sell': 'ab-s'}
            return f'<span class="analyst-badge {m.get(label, "ab-h")}">{label}</span>'

        def quality_badge(q):
            m = {'Excellent': 'qb-ex', 'Good': 'qb-gd', 'Average': 'qb-av', 'Poor': 'qb-po'}
            return f'<span class="qbadge {m.get(q, "qb-av")}">{q}</span>'

        def rr_color(v):
            return '#00e676' if v >= 2 else ('#00d4ff' if v >= 1 else '#ff3d57')

        def pe_color(v, direction='buy'):
            if v <= 0: return '#4a6080'
            if direction == 'buy':
                return '#00e676' if v < 25 else ('#ffab00' if v < 40 else '#ff3d57')
            else:
                return '#ff3d57' if v > 40 else ('#ffab00' if v > 25 else '#00e676')

        def w52_color(pct):
            return '#ff3d57' if pct >= -5 else ('#ffab00' if pct >= -20 else '#00e676')

        def beta_color(v):
            return '#ff3d57' if v > 1.5 else ('#ffab00' if v > 1.0 else '#00e676')

        # ── BUY TABLE ─────────────────────────────────────────────────────────
        if not top_buys.empty:
            html += """
  <div class="section-hdr">
    <div class="section-pill pill-buy">▲ Top 20 Buy Recommendations</div>
    <div class="section-line"></div>
    <div class="section-note">STOCK INFO · TRADE SETUP · TECHNICALS · FUNDAMENTALS · META</div>
  </div>
  <div class="tbl-wrap"><table>
    <thead>
      <tr class="grp-row">
        <th class="grp-stock" colspan="3">STOCK INFO</th>
        <th class="grp-trade gsep" colspan="6">TRADE SETUP</th>
        <th class="grp-tech gsep"  colspan="5">TECHNICALS</th>
        <th class="grp-fund gsep"  colspan="4">FUNDAMENTALS</th>
        <th class="grp-meta gsep"  colspan="3">META</th>
      </tr>
      <tr class="col-row">
        <th class="ch-stock" style="width:26px">#</th>
        <th class="ch-stock">Stock / Sector</th>
        <th class="ch-stock">Price</th>
        <th class="ch-trade gsep">Rating / Score</th>
        <th class="ch-trade">Upside</th>
        <th class="ch-trade">Target (S/R)</th>
        <th class="ch-trade">Stop Loss</th>
        <th class="ch-trade">ATR</th>
        <th class="ch-trade">R : R</th>
        <th class="ch-tech gsep">RSI</th>
        <th class="ch-tech">ADX</th>
        <th class="ch-tech">Vol / Avg</th>
        <th class="ch-tech">Sup Dist</th>
        <th class="ch-tech">52W Hi %</th>
        <th class="ch-fund gsep">P/E</th>
        <th class="ch-fund">Beta</th>
        <th class="ch-fund">Div %</th>
        <th class="ch-fund">Quality</th>
        <th class="ch-meta gsep">Analyst</th>
        <th class="ch-meta">Earnings</th>
        <th class="ch-meta">Action</th>
      </tr>
    </thead>
    <tbody>
"""
            for i, (_, row) in enumerate(top_buys.iterrows(), 1):
                rec      = row['Recommendation']
                sc_color = '#00e676' if row['Combined_Score'] >= 75 else ('#00d4ff' if row['Combined_Score'] >= 55 else '#ffab00')
                sc_bar   = '#00c853' if row['Combined_Score'] >= 75 else ('#0099cc' if row['Combined_Score'] >= 55 else '#f59e0b')
                upcls    = 'up' if row['Upside'] >= 0 else 'dn'
                rsic     = '#ff3d57' if row['RSI'] > 70 else ('#00e676' if row['RSI'] < 30 else '#60a5fa')
                w52      = ((row['Price'] - row['52W_High']) / row['52W_High']) * 100
                tbcls, tbtxt = target_badge_html(row.get('Target_Status', ''))
                st       = row.get('Stop_Type', 'ATR Stop')
                scls     = 'slt-atr' if st == 'ATR Stop' else 'slt-beta'
                slbl     = ('📐 ATR Stop' if st == 'ATR Stop' else '🔒 Beta Cap')
                div      = f"{row['Dividend_Yield']:.2f}%" if row['Dividend_Yield'] > 0 else '—'
                divc     = '#00e676' if row['Dividend_Yield'] > 0 else '#4a6080'
                rr       = row['Risk_Reward']

                html += f"""      <tr>
        <td><span class="rnum">{i}</span></td>
        <td>
          <div class="stock-name">{row['Name']}</div>
          <div class="stock-sym">{row['Symbol']}</div>
          <div class="stock-sec">{row.get('Sector','N/A')}</div>
        </td>
        <td><div class="price-val">₹{row['Price']:,.2f}</div></td>
        <td class="gsep">
          {rating_badge(rec, row['Rating'])}
          {score_cell(row['Combined_Score'], sc_color, sc_bar)}
        </td>
        <td><span class="upside-val {upcls}">{row['Upside']:+.1f}%</span></td>
        <td>
          <span class="target-badge {tbcls}">{tbtxt}</span>
          <div class="t1-val">₹{row['Target_1']:,.2f}</div>
          <div class="t2-val">T2: ₹{row['Target_2']:,.2f}</div>
        </td>
        <td>
          <div class="sl-val">₹{row['Stop_Loss']:,.2f}</div>
          <div class="sl-pct">-{row['SL_Percentage']:.1f}%</div>
          <span class="sl-type {scls}">{slbl}</span>
        </td>
        <td>
          <div class="atr-val">₹{row['ATR']:,.2f}</div>
          <div class="atr-sub">{row['ATR_Pct']:.1f}% · {row['ATR_Multiplier']}×</div>
        </td>
        <td><span class="rr-val" style="color:{rr_color(rr)}">{rr:.1f}×</span></td>
        <td class="gsep">
          <div class="rsi-val" style="color:{rsic}">{row['RSI']:.0f}</div>
          <div class="rsi-sig">{row['RSI_Signal']}</div>
        </td>
        <td>{adx_cell(row.get('ADX', 0))}</td>
        <td>{vol_cell(row.get('Vol_Ratio', 1.0))}</td>
        <td class="gsep">{sdist_cell(row.get('Support_Dist_Pct', 0))}</td>
        <td><span class="mono-sm" style="color:{w52_color(w52)}">{w52:+.1f}%</span></td>
        <td class="gsep"><span class="mono-sm" style="color:{pe_color(row['PE_Ratio'],'buy')}">{f"{row['PE_Ratio']:.1f}" if row['PE_Ratio']>0 else 'N/A'}</span></td>
        <td><span class="mono-sm" style="color:{beta_color(row['Beta'])}">{row['Beta']:.2f}</span></td>
        <td><span class="mono-sm" style="color:{divc}">{div}</span></td>
        <td>{quality_badge(row['Quality'])}</td>
        <td class="gsep">{analyst_badge(row.get('Analyst','N/A'))}</td>
        <td><div class="earn-date">{row.get('Earnings_Date','N/A')}</div></td>
        <td>{rating_badge(rec, 'BUY' if rec=='BUY' else 'STRONG BUY')}</td>
      </tr>
"""
            html += "    </tbody></table></div>\n"

        # ── SELL TABLE ────────────────────────────────────────────────────────
        if not top_sells.empty:
            html += """
  <div class="section-hdr">
    <div class="section-pill pill-sell">▼ Top 20 Sell Recommendations</div>
    <div class="section-line"></div>
    <div class="section-note">STOCK INFO · TRADE SETUP · TECHNICALS · FUNDAMENTALS · META</div>
  </div>
  <div class="tbl-wrap"><table>
    <thead>
      <tr class="grp-row">
        <th class="grp-stock" colspan="3">STOCK INFO</th>
        <th class="grp-trade gsep" colspan="6">TRADE SETUP</th>
        <th class="grp-tech gsep"  colspan="5">TECHNICALS</th>
        <th class="grp-fund gsep"  colspan="4">FUNDAMENTALS</th>
        <th class="grp-meta gsep"  colspan="3">META</th>
      </tr>
      <tr class="col-row">
        <th class="ch-stock" style="width:26px">#</th>
        <th class="ch-stock">Stock / Sector</th>
        <th class="ch-stock">Price</th>
        <th class="ch-trade gsep">Rating / Score</th>
        <th class="ch-trade">Downside</th>
        <th class="ch-trade">Target (S/R)</th>
        <th class="ch-trade">Stop Loss</th>
        <th class="ch-trade">ATR</th>
        <th class="ch-trade">R : R</th>
        <th class="ch-tech gsep">RSI</th>
        <th class="ch-tech">MACD</th>
        <th class="ch-tech">ADX</th>
        <th class="ch-tech">Vol / Avg</th>
        <th class="ch-tech">52W Hi %</th>
        <th class="ch-fund gsep">P/E</th>
        <th class="ch-fund">Beta</th>
        <th class="ch-fund">Div %</th>
        <th class="ch-fund">Quality</th>
        <th class="ch-meta gsep">Analyst</th>
        <th class="ch-meta">Earnings</th>
        <th class="ch-meta">Action</th>
      </tr>
    </thead>
    <tbody>
"""
            for i, (_, row) in enumerate(top_sells.iterrows(), 1):
                rec      = row['Recommendation']
                dncls    = 'dn' if row['Upside'] >= 0 else 'up'
                rsic     = '#ff3d57' if row['RSI'] > 70 else ('#00e676' if row['RSI'] < 30 else '#ffab00')
                mcdcls   = 'macd-bear' if row['MACD'] == 'Bearish' else 'macd-bull'
                w52      = ((row['Price'] - row['52W_High']) / row['52W_High']) * 100
                tbcls, tbtxt = target_badge_html(row.get('Target_Status', ''))
                st       = row.get('Stop_Type', 'ATR Stop')
                scls     = 'slt-atr' if st == 'ATR Stop' else 'slt-beta'
                slbl     = ('📐 ATR Stop' if st == 'ATR Stop' else '🔒 Beta Cap')
                div      = f"{row['Dividend_Yield']:.2f}%" if row['Dividend_Yield'] > 0 else '—'
                divc     = '#00e676' if row['Dividend_Yield'] > 0 else '#4a6080'
                rr       = row['Risk_Reward']

                html += f"""      <tr>
        <td><span class="rnum">{i}</span></td>
        <td>
          <div class="stock-name">{row['Name']}</div>
          <div class="stock-sym">{row['Symbol']}</div>
          <div class="stock-sec">{row.get('Sector','N/A')}</div>
        </td>
        <td><div class="price-val">₹{row['Price']:,.2f}</div></td>
        <td class="gsep">
          {rating_badge(rec, row['Rating'])}
          {score_cell(row['Combined_Score'], '#ff3d57', '#c62828')}
        </td>
        <td><span class="upside-val {dncls}">{row['Upside']:+.1f}%</span></td>
        <td>
          <span class="target-badge {tbcls}">{tbtxt}</span>
          <div class="t1-val">₹{row['Target_1']:,.2f}</div>
          <div class="t2-val">T2: ₹{row['Target_2']:,.2f}</div>
        </td>
        <td>
          <div style="font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;color:#ffab00">₹{row['Stop_Loss']:,.2f}</div>
          <div class="sl-pct">+{row['SL_Percentage']:.1f}%</div>
          <span class="sl-type {scls}">{slbl}</span>
        </td>
        <td>
          <div class="atr-val">₹{row['ATR']:,.2f}</div>
          <div class="atr-sub">{row['ATR_Pct']:.1f}% · {row['ATR_Multiplier']}×</div>
        </td>
        <td><span class="rr-val" style="color:{rr_color(rr)}">{rr:.1f}×</span></td>
        <td class="gsep">
          <div class="rsi-val" style="color:{rsic}">{row['RSI']:.0f}</div>
          <div class="rsi-sig">{row['RSI_Signal']}</div>
        </td>
        <td><span class="{mcdcls}">{row['MACD']}</span></td>
        <td>{adx_cell(row.get('ADX', 0))}</td>
        <td>{vol_cell(row.get('Vol_Ratio', 1.0))}</td>
        <td><span class="mono-sm" style="color:{w52_color(w52)}">{w52:+.1f}%</span></td>
        <td class="gsep"><span class="mono-sm" style="color:{pe_color(row['PE_Ratio'],'sell')}">{f"{row['PE_Ratio']:.1f}" if row['PE_Ratio']>0 else 'N/A'}</span></td>
        <td><span class="mono-sm" style="color:{beta_color(row['Beta'])}">{row['Beta']:.2f}</span></td>
        <td><span class="mono-sm" style="color:{divc}">{div}</span></td>
        <td>{quality_badge(row['Quality'])}</td>
        <td class="gsep">{analyst_badge(row.get('Analyst','N/A'))}</td>
        <td><div class="earn-date">{row.get('Earnings_Date','N/A')}</div></td>
        <td>{rating_badge(rec, 'SELL' if rec=='SELL' else 'STRONG SELL')}</td>
      </tr>
"""
            html += "    </tbody></table></div>\n"

        html += f"""
  <div class="disc">
    <strong style="color:var(--red)">⚠ DISCLAIMER:</strong>
    For <strong>EDUCATIONAL PURPOSES ONLY</strong>. Not financial advice.
    Stop losses are ATR-based near real 12-month S/R zones. Targets derived from
    swing highs/lows, 52-week extremes and round-number levels. Earnings dates are estimates.
    Always conduct your own research, consult a SEBI-registered financial advisor,
    and never invest more than you can afford to lose.
  </div>
</div>

<footer>
  <strong>NIFTY 100 Market Influencers · NSE &amp; BSE</strong>
  · 12M S/R · ATR Stops · Grouped Columns · ADX · Vol · Earnings v3
  · Next Update: <strong>{next_update} IST</strong> · {now.strftime('%d %b %Y')}
</footer>

<script>
function updateClock() {{
  var now  = new Date();
  var ist  = new Date(now.toLocaleString('en-US', {{timeZone: 'Asia/Kolkata'}}));
  var h = ist.getHours(), m = ist.getMinutes(), s = ist.getSeconds();
  var ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  var p = function(n) {{ return String(n).padStart(2,'0'); }};
  document.getElementById('liveClock').textContent =
    p(h) + ':' + p(m) + ':' + p(s) + ' ' + ampm + ' IST';
  var mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  document.getElementById('liveDate').textContent =
    p(ist.getDate()) + ' ' + mo[ist.getMonth()] + ' ' + ist.getFullYear() + ' · IST';
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
                print("❌ Set GMAIL_USER and GMAIL_APP_PASSWORD env vars")
                return False
            now = self.get_ist_time()
            tod = "Morning" if now.hour < 12 else "Evening"
            msg = MIMEMultipart('alternative')
            msg['From']    = from_email
            msg['To']      = to_email
            msg['Subject'] = f"💎 NIFTY 100 Report v3 — {tod} {now.strftime('%d %b %Y')}"
            msg.attach(MIMEText(self.generate_html(), 'html'))
            srv = smtplib.SMTP('smtp.gmail.com', 587)
            srv.starttls()
            srv.login(from_email, password)
            srv.send_message(msg)
            srv.quit()
            print(f"✅ Email sent to {to_email}")
            return True
        except Exception as e:
            print(f"❌ Email error: {e}")
            return False

    # =========================================================================
    #  ENTRY POINT
    # =========================================================================
    def generate_complete_report(self, send_email_flag=True, recipient_email=None,
                                  output_file='index.html'):
        now = self.get_ist_time()
        print("=" * 70)
        print("💎 NIFTY 100 ANALYZER v3 — Grouped UI + ATR Stops + 12M S/R")
        print(f"   {now.strftime('%d %b %Y, %I:%M %p IST')}")
        print("=" * 70)
        self.analyze_all_stocks()
        self.save_html(output_file)
        if send_email_flag and recipient_email:
            self.send_email(recipient_email)
        print("=" * 70)
        print("✅ DONE")
        print("=" * 70)


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
