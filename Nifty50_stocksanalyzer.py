"""
NIFTY 100 COMPLETE STOCK ANALYZER - REDESIGNED UI v5.6
Technical + Fundamental Analysis with Email Delivery + GitHub Pages

=======================================================================
ACCURACY FIXES in v5.6 (resolves Lupin/post-overbought stocks showing BUY
despite MACD bearish + RSI fading from 80):

  V56-1  R:R FLOOR FOR BUY - A BUY rating now requires R:R ≥ 1.0.
           Previously only STRONG BUY had an R:R gate (1.5x). BUY had
           no floor — a trade with 0.9x R:R (reward < risk) was still
           flagged as actionable. Now: R:R < 1.0 on a BUY → capped to
           HOLD (R:R < 1.0). Lupin had R:R 0.9x → correctly becomes HOLD.

  V56-2  RSI POST-OVERBOUGHT PULLBACK SIGNAL - Added as a 7th bearish
           veto signal. Detects stocks where RSI peaked above 70 within
           the last 15 bars but is now falling. This is the classic
           distribution top pattern: overbought → profit-taking → unwinding.
           Combined with MACD bearish + RSI falling fast, Lupin now hits
           3 veto signals → capped to HOLD via V54-1 Trend Veto Gate.
           Prevents fundamentally strong stocks from being flagged BUY
           during active pullbacks from overbought extremes.

=======================================================================
ACCURACY FIXES in v5.4 (resolves fundamentals overriding clear
downtrend signals - root cause of SBIN still showing BUY in v5.3):
  Diagnosed: SBIN had good fundamentals (fund_score ~70) which at 65%
  weight mathematically overwhelmed a negative tech score. Even with
  SMA20 declining + death cross + MACD bearish, the combined score
  stayed above 50 → BUY. Four targeted fixes close this gap:

  V54-1  TREND VETO GATE - Hard block before rating is assigned.
           If 3+ of these bearish signals fire simultaneously:
             · SMA20 declining
             · Death cross forming (SMA20 < SMA50)
             · MACD < Signal (bearish crossover)
             · RSI < 50 (momentum lost)
             · Price < SMA50
           → Maximum rating is capped at HOLD, regardless of combined
           score. Fundamentals cannot rescue a stock in confirmed
           technical downtrend. SBIN had all 5 signals → HOLD cap.

  V54-2  SMA200 SLOPE GUARD - SMA200 bonus now requires SMA200
           itself to be rising (today > 10 bars ago). A rising price
           above a flat/falling SMA200 (common after a run-up) no
           longer earns the +2 bonus. SMA200 must confirm uptrend.
           Flat/falling SMA200 = 0 points instead of +2.

  V54-3  ANALYST BONUS TIGHTENED - Analyst buy bonus now requires
           tech_score >= 2 (was > 0). A tech score of +1 is borderline
           and should not unlock the analyst +5 rescue. Requires clear
           net positive technicals before analyst buy counts.

  V54-4  DYNAMIC WEIGHT SHIFT - When 3+ bearish tech signals fire,
           weight shifts from 65/35 (fund/tech) to 50/50.
           Prevents a great balance sheet from hiding an active
           distribution top. Weight reverts to 65/35 normally.

ACCURACY FIXES in v5.3 (resolves SMA50 lag - stock can be in 6-week
downtrend and still be "above SMA50" from the prior run-up):
  Diagnosed: SBIN peaked at ₹1,200 Jan 2026. SMA50 still ~₹1,110.
  Price 1143 > SMA50 1110 -> V52-1 and V52-3 from v5.2 both miss.
  The fix must not rely on SMA50 position alone.

  V53-1  SMA20 slope penalty - if SMA20 today < SMA20 five bars ago,
           the short-term trend is actively declining right now.
           This is real-time and does not wait for SMA50 to lag down.
           SMA20 declining -> -1 tech score, flagged "SMA Declining".
           This directly catches a stock rolling over from a peak.

  V53-2  Death-cross forming penalty - if SMA20 < SMA50, the short-
           term average has crossed below the medium-term average.
           This is the early warning stage of a bearish crossover
           and adds -1 additional tech score, regardless of where
           price sits relative to SMA50.
           Combined with V53-1, SBIN gets -2 more tech points,
           dropping it cleanly below the BUY threshold.

ACCURACY FIXES in v5.2 (downtrending stocks appearing as BUY):
  V52-1  ADX direction-aware (bonus only when price > SMA50)
  V52-2  RSI weak-momentum zone 30-45 = -1 tech score
  V52-3  Double SMA penalty (price < SMA20 AND < SMA50 = extra -1)
  V52-4  Sector-adjusted PE for Financial sector (banks: PE < 15)

CALIBRATION FIXES in v5.1 (resolves "only 2 stocks" issue):
  CAL-1  Score thresholds: STRONG BUY 75->70, BUY 55->50
  CAL-2  R:R gate split: STRONG BUY 1.5x, BUY 1.2x
  CAL-3  Volume ratio: 5-day avg instead of single last-bar
  CAL-4  Growth penalty capped at -10 total
  CAL-5  Partial credit for missing yFinance fields

RETAINED FROM v5 (7 improvements):
  NEW-1  RSI Divergence detection
  NEW-2  Volume hard gatekeeper
  NEW-3  Sector diversity cap (max 3 per sector)
  NEW-4  yFinance data sanity check
  NEW-6  FCF weight +15
  NEW-7  D/E weight +15

RETAINED FROM v4 (8 improvements):
  FIX-1  Fundamentals 65% / technicals 35% (dynamic in v5.4)
  FIX-2  ADX weak-trend penalty
  FIX-3  RSI context-aware (oversold in downtrend)
  FIX-4  STRONG BUY requires R:R ≥ 1.5
  FIX-5  Negative growth penalises fund score
  FIX-6  Volume ratio influences tech score
  FIX-7  Analyst consensus +/-5 (tightened gate in v5.4)
  FIX-8  52W high proximity bonus
=======================================================================

Requirements:
    pip install yfinance pandas numpy pytz
"""

import yfinance as yf
import pandas as pd
from datetime import datetime, timezone
import pytz
import warnings
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

warnings.filterwarnings('ignore')

# == Sector diversity cap: max picks per sector in Top 20 Buy table ==
MAX_PICKS_PER_SECTOR = 3


class Nifty100CompleteAnalyzer:
    def __init__(self):
        self.nifty100_stocks = {
            # == NIFTY 50 ==========================================
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
            # == NIFTY NEXT 50 =====================================
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

        # ── FIX-SECTOR: Fallback sector map for NSE stocks ────────────────────
        # yFinance frequently returns None for sector on NSE-listed stocks.
        # Without this, every stock whose sector is None lands in the 'N/A'
        # bucket, which holds up to 4 stocks in the top-20 — crowding out
        # genuine sector diversity and causing the same names to repeat daily.
        # This map is the single source of truth; it is used in analyze_stock()
        # to replace any None/empty sector returned by yFinance.
        self.sector_fallback = {
            # Financials
            'HDFCBANK':   'Financial Services', 'ICICIBANK':  'Financial Services',
            'KOTAKBANK':  'Financial Services', 'AXISBANK':   'Financial Services',
            'SBIN':       'Financial Services', 'INDUSINDBK': 'Financial Services',
            'BAJFINANCE': 'Financial Services', 'BAJAJFINSV': 'Financial Services',
            'SHRIRAMFIN': 'Financial Services', 'MUTHOOTFIN': 'Financial Services',
            'CHOLAFIN':   'Financial Services', 'BAJAJHLDNG': 'Financial Services',
            'SBICARD':    'Financial Services', 'ICICIPRULI': 'Financial Services',
            'ICICIGI':    'Financial Services', 'HDFCAMC':    'Financial Services',
            'SBILIFE':    'Financial Services', 'HDFCLIFE':   'Financial Services',
            'POLICYBZR':  'Financial Services',
            # IT
            'TCS':        'Information Technology', 'INFY':    'Information Technology',
            'WIPRO':      'Information Technology', 'HCLTECH': 'Information Technology',
            'TECHM':      'Information Technology', 'LTIM':    'Information Technology',
            'TATAELXSI':  'Information Technology', 'COFORGE': 'Information Technology',
            'PERSISTENT': 'Information Technology', 'OFSS':    'Information Technology',
            'LTTS':       'Information Technology', 'NAUKRI':  'Information Technology',
            # Healthcare / Pharma
            'SUNPHARMA':  'Healthcare', 'CIPLA':      'Healthcare',
            'DRREDDY':    'Healthcare', 'DIVISLAB':   'Healthcare',
            'APOLLOHOSP': 'Healthcare', 'TORNTPHARM': 'Healthcare',
            'LUPIN':      'Healthcare', 'AUROPHARMA': 'Healthcare',
            'ALKEM':      'Healthcare', 'MAXHEALTH':  'Healthcare',
            'FORTIS':     'Healthcare',
            # FMCG / Consumer
            'HINDUNILVR': 'FMCG', 'ITC':       'FMCG', 'NESTLEIND':  'FMCG',
            'BRITANNIA':  'FMCG', 'DABUR':     'FMCG', 'MARICO':     'FMCG',
            'GODREJCP':   'FMCG', 'COLPAL':    'FMCG', 'TATACONSUM': 'FMCG',
            'MCDOWELL-N': 'FMCG', 'PAGEIND':   'FMCG',
            # Auto
            'MARUTI':     'Automobile', 'M&M':        'Automobile',
            'TMCV':       'Automobile', 'TMPV':       'Automobile',
            'HEROMOTOCO': 'Automobile', 'EICHERMOT':  'Automobile',
            'BAJAJ-AUTO': 'Automobile', 'MOTHERSON':  'Automobile',
            'BALKRISIND': 'Automobile',
            # Energy / Oil & Gas
            'RELIANCE':   'Energy', 'ONGC':       'Energy',
            'BPCL':       'Energy', 'POWERGRID':  'Energy',
            'NTPC':       'Energy', 'COALINDIA':  'Energy',
            'ADANIGREEN': 'Energy',
            # Metals & Mining
            'TATASTEEL':  'Metals', 'JSWSTEEL':   'Metals',
            'HINDALCO':   'Metals', 'VEDL':       'Metals',
            'SAIL':       'Metals', 'NMDC':       'Metals',
            'JINDALSTEL': 'Metals',
            # Cement & Construction
            'ULTRACEMCO': 'Cement', 'GRASIM':     'Cement',
            'AMBUJACEM':  'Cement', 'ACC':        'Cement',
            # Telecom
            'BHARTIARTL': 'Telecom',
            # Capital Goods / Industrials
            'LT':         'Capital Goods', 'SIEMENS':   'Capital Goods',
            'HAVELLS':    'Capital Goods', 'VOLTAS':    'Capital Goods',
            'CONCOR':     'Capital Goods', 'RVNL':      'Capital Goods',
            'ADANIPORTS': 'Capital Goods',
            # Paints / Chemicals
            'ASIANPAINT': 'Chemicals', 'BERGEPAINT': 'Chemicals',
            'PIDILITIND': 'Chemicals',
            # Consumer Durables / Retail
            'TITAN':      'Consumer Durables', 'DMART':     'Consumer Durables',
            # New-age / Internet
            'ZOMATO':     'Internet', 'NYKAA':     'Internet',
            'PAYTM':      'Internet',
            # Diversified
            'ADANIENT':   'Diversified',
            # Aviation
            'INDIGO':     'Aviation',
            # Finance (PSU)
            'RECLTD':     'Financial Services', 'PFC':       'Financial Services',
            'IRCTC':      'Consumer Services',
        }

    # =========================================================================
    #  UTILITY
    # =========================================================================
    def get_ist_time(self):
        return datetime.now(pytz.timezone('Asia/Kolkata'))

    def calculate_rsi(self, prices, period=14):
        # V55: Use Wilder's smoothing (EMA with alpha=1/period) to match TradingView.
        # TradingView RSI uses Wilder's method — NOT simple rolling mean.
        # Simple rolling mean overstates RSI by 5-10 points vs TradingView.
        delta = prices.diff()
        gain  = delta.where(delta > 0, 0)
        loss  = (-delta.where(delta < 0, 0))
        # Wilder's smoothing = EMA with com=period-1 (equivalent to alpha=1/period)
        avg_gain = gain.ewm(com=period - 1, min_periods=period).mean()
        avg_loss = loss.ewm(com=period - 1, min_periods=period).mean()
        rs = avg_gain / avg_loss
        return (100 - (100 / (1 + rs))).iloc[-1]

    def calculate_rsi_slope(self, prices, period=14, lookback=5):
        """
        Returns the RSI slope direction and strength.

        Instead of just asking "what is RSI now?", this asks:
        "Which direction is RSI MOVING?" — because:
          · RSI rising from 35 → 50 = momentum building    → GOOD
          · RSI falling from 65 → 46 = momentum fading     → BAD (Power Grid case)
          · RSI flat at 50           = no signal            → NEUTRAL

        Parameters:
          lookback = how many bars back to compare RSI (default 5 = 1 trading week)

        Returns dict:
          'slope'     : float  (RSI today minus RSI 5 bars ago)
          'direction' : 'Rising' | 'Falling' | 'Flat'
          'strong'    : bool   (|slope| > 8 = strong move)
          'rsi_5bar'  : float  (RSI value 5 bars ago, for display)
        """
        try:
            delta    = prices.diff()
            gain     = delta.where(delta > 0, 0)
            loss     = (-delta.where(delta < 0, 0))
            avg_gain = gain.ewm(com=period - 1, min_periods=period).mean()
            avg_loss = loss.ewm(com=period - 1, min_periods=period).mean()
            rsi_ser  = 100 - (100 / (1 + avg_gain / avg_loss))
            rsi_ser  = rsi_ser.dropna()

            if len(rsi_ser) < lookback + 2:
                return {'slope': 0, 'direction': 'Flat', 'strong': False, 'rsi_5bar': rsi_ser.iloc[-1], 'rsi_15bar': rsi_ser.iloc[-1], 'rsi_peak': rsi_ser.iloc[-1]}

            rsi_now    = rsi_ser.iloc[-1]
            rsi_prev   = rsi_ser.iloc[-(lookback + 1)]      # 5 bars ago
            rsi_15bar  = rsi_ser.iloc[-16] if len(rsi_ser) >= 16 else rsi_prev  # 15 bars ago (~3 weeks)
            rsi_peak   = rsi_ser.iloc[-16:].max() if len(rsi_ser) >= 16 else rsi_now  # highest RSI in last 15 bars
            slope      = round(rsi_now - rsi_prev, 2)

            if slope > 3:
                direction = 'Rising'
            elif slope < -3:
                direction = 'Falling'
            else:
                direction = 'Flat'

            strong = abs(slope) > 8   # sharp move — e.g. 71 → 46 in Power Grid

            return {
                'slope':      slope,
                'direction':  direction,
                'strong':     strong,
                'rsi_5bar':   round(rsi_prev, 1),
                'rsi_15bar':  round(rsi_15bar, 1),   # RSI 3 weeks ago
                'rsi_peak':   round(rsi_peak, 1),    # highest RSI in last 15 bars
            }
        except Exception:
            return {'slope': 0, 'direction': 'Flat', 'strong': False, 'rsi_5bar': 50, 'rsi_15bar': 50, 'rsi_peak': 50}

    # == NEW-1: RSI Divergence helper =========================================
    def detect_rsi_divergence(self, prices, window=14):
        """
        Bearish divergence: price makes a HIGHER high in last 20 bars,
        but RSI makes a LOWER high over the same window.
        Returns: 'Bearish Divergence', 'Bullish Divergence', or 'None'
        """
        try:
            # Calculate RSI series using Wilder's smoothing (matches TradingView)
            delta    = prices.diff()
            gain     = delta.where(delta > 0, 0)
            loss     = (-delta.where(delta < 0, 0))
            avg_gain = gain.ewm(com=window - 1, min_periods=window).mean()
            avg_loss = loss.ewm(com=window - 1, min_periods=window).mean()
            rsi_ser  = 100 - (100 / (1 + avg_gain / avg_loss))
            rsi_ser  = rsi_ser.dropna()

            lookback  = 20   # bars to compare
            if len(prices) < lookback + window or len(rsi_ser) < lookback:
                return 'None'

            # Recent window vs prior window
            recent_price = prices.iloc[-lookback:]
            prior_price  = prices.iloc[-(lookback * 2):-lookback]
            recent_rsi   = rsi_ser.iloc[-lookback:]
            prior_rsi    = rsi_ser.iloc[-(lookback * 2):-lookback]

            recent_price_high = recent_price.max()
            prior_price_high  = prior_price.max()
            recent_rsi_high   = recent_rsi.max()
            prior_rsi_high    = prior_rsi.max()

            recent_price_low  = recent_price.min()
            prior_price_low   = prior_price.min()
            recent_rsi_low    = recent_rsi.min()
            prior_rsi_low     = prior_rsi.min()

            # Bearish: price higher high + RSI lower high
            if (recent_price_high > prior_price_high * 1.005 and
                    recent_rsi_high < prior_rsi_high * 0.97):
                return 'Bearish Divergence'

            # Bullish: price lower low + RSI higher low
            if (recent_price_low < prior_price_low * 0.995 and
                    recent_rsi_low > prior_rsi_low * 1.03):
                return 'Bullish Divergence'

            return 'None'
        except Exception:
            return 'None'
    # =========================================================================

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
        plus_dm[plus_dm < 0]        = 0
        minus_dm[minus_dm < 0]      = 0
        plus_dm[plus_dm < minus_dm] = 0
        minus_dm[minus_dm < plus_dm]= 0
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
        # CAL-3: Use 5-day average instead of single last-bar snapshot.
        # Single-bar volume is too noisy - a great stock with a quiet
        # day before the report runs gets blocked despite a strong trend.
        avg_vol    = df['Volume'].tail(20).mean()
        if avg_vol == 0:
            return 1.0
        recent_vol = df['Volume'].tail(5).mean()   # 5-day average
        return round(recent_vol / avg_vol, 2)

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
                result[label] = {'price': 'N/A', 'chg': '-', 'cls': ''}
        return result

    # =========================================================================
    #  NEW-4: yFINANCE DATA SANITY CHECK
    # Detects >20% single-day price move as likely bad/unadjusted data
    # =========================================================================
    def is_data_clean(self, df):
        """
        Returns (True, '') if data looks valid.
        Returns (False, reason) if a suspicious spike is detected.
        Checks daily close-to-close % change - any move >20% in a single
        bar without a corresponding volume surge is flagged as dirty data.
        """
        try:
            close      = df['Close']
            volume     = df['Volume']
            pct_change = close.pct_change().abs()

            # Find days with >20% price move
            spike_days = pct_change[pct_change > 0.20]
            if spike_days.empty:
                return True, ''

            for spike_date, spike_val in spike_days.items():
                # Check if volume on that day was at least 3x average
                # (corporate actions like splits usually come with high volume)
                avg_vol = volume.mean()
                spike_vol = volume.loc[spike_date] if spike_date in volume.index else 0
                if avg_vol > 0 and spike_vol < avg_vol * 3:
                    # Large price move with normal volume = likely data error
                    return False, f"Suspicious {spike_val*100:.0f}% move on {spike_date.date()} - possible bad data"

            return True, ''
        except Exception:
            return True, ''   # if check fails, don't block the stock

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
            target_status = "ATH Zone - Projected"
        if t1 < min_target:
            t1            = round(min_target, 2)
            t2            = round(t1 * 1.04, 2)
            target_status += " (ATR Adj)"
        return round(t1, 2), round(t2, 2), 0, target_status

    # =========================================================================
    #  FUNDAMENTAL SCORE
    #  v5.2: Sector-adjusted PE thresholds for Financial sector (V52-4)
    #  v5.1: CAL-4 growth penalty cap, CAL-5 partial credit for missing fields
    #  v5:   FCF weight +5->+15, D/E weight +10->+15 (NEW-6, NEW-7)
    #  v4:   negative growth penalised (FIX-5)
    # =========================================================================
    def get_fundamental_score(self, info, sector=''):
        score = 0

        # V52-4: Sector-adjusted PE thresholds.
        # Banks and Financial Services always trade at structurally low PE
        # (8-15) due to capital intensity and NPA provisioning requirements.
        # Using the same PE < 25 threshold as IT/FMCG stocks rewards PSU
        # banks for being "cheap" when they are merely sector-typical.
        # Financial sector: PE < 15 = +10, 15-20 = +5, > 20 = 0.
        # All other sectors: PE < 25 = +10, 25-35 = +5 (unchanged).
        pe  = info.get('trailingPE', info.get('forwardPE', 0))
        pb  = info.get('priceToBook', 0)
        peg = info.get('pegRatio', 0)

        is_financial = sector in ('Financial Services', 'Banks', 'Banking',
                                  'Insurance', 'Financial')
        if is_financial:
            if pe and 0 < pe < 15:      score += 10   # genuinely cheap bank
            elif pe and 15 <= pe < 20:  score += 5    # fair value for a bank
            # PE > 20 for a bank = expensive for its sector -> 0 points
        else:
            if pe  and 0 < pe  < 25:    score += 10
            elif pe  and 25 <= pe < 35: score += 5

        if pb  and 0 < pb  < 3:      score += 5
        elif pb  and 3 <= pb  < 5:   score += 3
        if peg and 0 < peg < 1:      score += 10
        elif peg and 1 <= peg < 2:   score += 5
        else:                        score += 3   # CAL-5: partial credit if PEG missing

        # Profitability
        roe = info.get('returnOnEquity', 0)
        roa = info.get('returnOnAssets', 0)
        pm  = info.get('profitMargins', 0)
        if roe and roe > 0.15:   score += 10
        elif roe and roe > 0.10: score += 5
        if roa and roa > 0.05:   score += 5
        elif roa and roa > 0.02: score += 3
        else:                    score += 2   # CAL-5: partial credit if ROA missing
        if pm  and pm  > 0.10:   score += 10
        elif pm  and pm  > 0.05: score += 5

        # Growth (FIX-5: penalise negative growth; CAL-4: cap total penalty at -10)
        rg = info.get('revenueGrowth', 0)
        eg = info.get('earningsGrowth', 0)
        growth_penalty = 0
        if rg and rg > 0.15:    score += 10
        elif rg and rg > 0.10:  score += 7
        elif rg and rg > 0.05:  score += 5
        elif rg and rg < 0:     growth_penalty += 10  # track separately for cap

        if eg and eg > 0.15:    score += 10
        elif eg and eg > 0.10:  score += 7
        elif eg and eg > 0.05:  score += 5
        elif eg and eg < 0:     growth_penalty += 10  # track separately for cap

        # CAL-4: Cap combined growth penalty at 10 - prevents cyclicals
        # (steel, cement, energy) with temporary negative quarters from
        # being wiped out entirely. They score 0 growth points, not -20.
        score -= min(growth_penalty, 10)

        # Balance sheet health
        de = info.get('debtToEquity', 0)
        cr = info.get('currentRatio', 0)
        fc = info.get('freeCashflow', 0)

        # NEW-7: D/E raised from +10 to +15 - high-debt firms collapse in
        # Indian market volatility (IL&FS, YES Bank, DHFL lessons)
        if de is not None:
            if de < 50:    score += 15   # was +10
            elif de < 100: score += 7    # was +5
        else:
            score += 5

        if cr and cr > 1.5:   score += 10
        elif cr and cr > 1.0: score += 5
        else:                  score += 3   # CAL-5: partial credit if CR missing/NA

        # NEW-6: FCF raised from +5 to +15 - Cash is King in Indian markets
        # Companies with strong FCF survive rate hikes & FII outflows
        if fc and fc > 0:     score += 15   # was +5

        return min(max(score, 0), 100)   # clamp 0-100

    # =========================================================================
    #  MAIN ANALYSIS
    # =========================================================================
    def analyze_stock(self, symbol, name):
        try:
            stock = yf.Ticker(symbol)

            # V55-DATA: Fetch with explicit end=today to force yFinance to include
            # the most recent candle. Without this, yFinance period='1y' can lag
            # by 1-3 days, causing RSI/MACD to be calculated on stale data.
            # auto_adjust=False: TradingView uses unadjusted prices — we match that
            # so RSI values align with what you see on TradingView charts.
            today = datetime.now(pytz.timezone('Asia/Kolkata')).date()
            df    = stock.history(period='1y', auto_adjust=False)

            if df.empty or len(df) < 200:
                return None

            # V55-FRESHNESS: Confirm the last candle is within 3 trading days of today.
            # If yFinance is lagging (common on weekends or public holidays),
            # the last row may be several days old — warn but continue.
            last_candle_date = df.index[-1].date() if hasattr(df.index[-1], 'date') else df.index[-1]
            days_lag = (today - last_candle_date).days
            if days_lag > 5:
                print(f"  ⚠ {symbol}: Data lag {days_lag} days (last candle: {last_candle_date})")

            # auto_adjust=False gives columns: Open, High, Low, Close, Adj Close, Volume
            # We use 'Close' (unadjusted) to match TradingView's default RSI calculation.
            # Rename for safety in case column names vary across yFinance versions.
            if 'Close' not in df.columns and 'Adj Close' in df.columns:
                df = df.rename(columns={'Adj Close': 'Close'})

            info  = stock.info

            # == NEW-4: Data sanity check - skip stocks with bad yFinance data ==
            data_ok, data_warn = self.is_data_clean(df)
            if not data_ok:
                print(f"  ⚠ Skipping {symbol}: {data_warn}")
                return None
            # ==================================================================

            current_price = df['Close'].iloc[-1]
            sma_20  = df['Close'].rolling(20).mean().iloc[-1]
            sma_50  = df['Close'].rolling(50).mean().iloc[-1]
            sma_200 = df['Close'].rolling(200).mean().iloc[-1]

            # V53-1: SMA20 slope - compare today's SMA20 vs 5 bars ago.
            # A declining SMA20 means short-term momentum is falling NOW,
            # not waiting for SMA50 to catch up over weeks.
            sma_20_series   = df['Close'].rolling(20).mean()
            sma_20_5bar_ago = sma_20_series.iloc[-6] if len(sma_20_series) >= 6 else sma_20
            sma_20_declining = sma_20 < sma_20_5bar_ago

            # V53-2: Death-cross forming - SMA20 has crossed below SMA50.
            # Early warning of sustained bearish momentum shift.
            death_cross_forming = sma_20 < sma_50

            rsi          = self.calculate_rsi(df['Close'])
            macd, signal = self.calculate_macd(df['Close'])
            atr          = self.calculate_atr(df)
            atr_pct      = round((atr / current_price) * 100, 2)
            adx          = self.calculate_adx(df)
            vol_ratio    = self.calculate_volume_ratio(df)

            # NEW-1: RSI Divergence detection
            rsi_divergence = self.detect_rsi_divergence(df['Close'])

            # V55-RSI-SLOPE: RSI direction — rising vs falling
            # Power Grid case: RSI was 71, now 46 = sharply falling = bearish
            # This is separate from RSI value — a falling RSI at 55 is worse
            # than a rising RSI at 45. Direction matters more than level.
            rsi_slope_data   = self.calculate_rsi_slope(df['Close'])
            rsi_slope        = rsi_slope_data['slope']
            rsi_direction    = rsi_slope_data['direction']    # 'Rising'|'Falling'|'Flat'
            rsi_slope_strong = rsi_slope_data['strong']       # True if |slope| > 8
            rsi_5bar         = rsi_slope_data['rsi_5bar']     # RSI 5 bars ago
            rsi_15bar        = rsi_slope_data['rsi_15bar']    # RSI 15 bars ago (~3 weeks)
            rsi_peak_15      = rsi_slope_data['rsi_peak']     # highest RSI in last 15 bars

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

            # V54-2: SMA200 slope guard - SMA200 must itself be rising to earn
            # the +2 uptrend bonus. A rising price above a flat or falling
            # SMA200 (common after a big run-up that has since peaked) no longer
            # earns the full bonus. Flat/falling SMA200 = 0 pts, not +2.
            sma_200_series   = df['Close'].rolling(200).mean()
            sma_200_10bar_ago = sma_200_series.iloc[-11] if len(sma_200_series) >= 11 else sma_200
            sma_200_rising   = sma_200 > sma_200_10bar_ago

            # == TECHNICAL SCORE ===============================================
            tech_score = 0
            tech_score += 1 if current_price > sma_20  else -1
            tech_score += 1 if current_price > sma_50  else -1
            # V54-2: +2 only when price above SMA200 AND SMA200 itself rising
            if current_price > sma_200 and sma_200_rising:
                tech_score += 2   # confirmed long-term uptrend
            elif current_price > sma_200 and not sma_200_rising:
                tech_score += 0   # above SMA200 but trend flattening - no bonus
            else:
                tech_score -= 2   # below SMA200 - penalise

            # V52-3: Double SMA penalty - price below BOTH SMA20 and SMA50
            # simultaneously confirms active short-term downtrend.
            if current_price < sma_20 and current_price < sma_50:
                tech_score -= 1   # confirmed short-term downtrend

            # V53-1: SMA20 slope penalty - SMA20 actively declining right now.
            # Catches stocks rolling over from a peak BEFORE SMA50 catches up.
            # This is the fix that correctly identifies SBIN's situation:
            # price still above SMA50 (lag) but SMA20 has been falling for weeks.
            if sma_20_declining:
                tech_score -= 1

            # V53-2: Death-cross forming - SMA20 has crossed below SMA50.
            # Early confirmed signal of sustained medium-term trend reversal.
            # Works alongside V53-1: if both fire, that's -2 together.
            if death_cross_forming:
                tech_score -= 1

            # FIX-3: RSI context-aware + V52-2: weak-momentum zone
            # V55: RSI slope now integrated — direction matters as much as value
            if rsi < 30:
                if current_price > sma_200:
                    tech_score += 2
                    rsi_signal = "Oversold ↑" if rsi_direction == 'Rising' else "Oversold"
                else:
                    tech_score -= 1
                    rsi_signal = "Oversold (Downtrend)"
            elif rsi > 70:
                if rsi_divergence == 'Bearish Divergence':
                    tech_score -= 3
                    rsi_signal = "Bearish Divergence ⚠"
                elif rsi_direction == 'Falling' and rsi_slope_strong:
                    # Power Grid case: RSI was >70, now sharply falling
                    # This is the most dangerous signal — overbought AND rolling over
                    tech_score -= 3
                    rsi_signal = f"Topping Out ⚠ ({rsi_5bar:.0f}→{rsi:.0f})"
                elif rsi_direction == 'Falling':
                    # RSI falling from overbought — early warning
                    tech_score -= 2
                    rsi_signal = f"Fading ↓ ({rsi_5bar:.0f}→{rsi:.0f})"
                else:
                    tech_score -= 1
                    rsi_signal = "Overbought"
            elif 30 <= rsi <= 45:
                if rsi_direction == 'Falling':
                    # Weak zone AND still falling — double trouble
                    tech_score -= 2
                    rsi_signal = f"Weak & Falling ↓ ({rsi_5bar:.0f}→{rsi:.0f})"
                else:
                    tech_score -= 1
                    rsi_signal = "Weak Momentum ⚠"
            elif 45 < rsi <= 55:
                # Neutral zone — direction is the only signal here
                if rsi_direction == 'Rising':
                    tech_score += 1
                    rsi_signal = f"Building ↑ ({rsi_5bar:.0f}→{rsi:.0f})"
                elif rsi_direction == 'Falling' and rsi_slope_strong:
                    # RSI falling sharply through neutral — Power Grid mid-fall
                    tech_score -= 2
                    rsi_signal = f"Falling Fast ↓ ({rsi_5bar:.0f}→{rsi:.0f})"
                elif rsi_direction == 'Falling':
                    tech_score -= 1
                    rsi_signal = f"Fading ↓ ({rsi_5bar:.0f}→{rsi:.0f})"
                else:
                    rsi_signal = "Neutral →"
            else:
                # RSI 55-70: healthy zone — reward rising, penalise falling
                if rsi_direction == 'Rising' and rsi > 65:
                    # V55-OB: RSI rising AND already above 65 = approaching overbought
                    # Don't reward further — it's closer to danger than opportunity
                    # Example: Coal India RSI 55→72 = overbought soon, not a buy signal
                    rsi_signal = f"Near Overbought ⚠ ({rsi:.0f}↑)"
                elif rsi_direction == 'Rising':
                    tech_score = min(tech_score + 1, 6)
                    rsi_signal = f"Momentum ↑ ({rsi_5bar:.0f}→{rsi:.0f})"
                elif rsi_direction == 'Falling' and rsi_slope_strong:
                    tech_score -= 2
                    rsi_signal = f"Rolling Over ↓ ({rsi_5bar:.0f}→{rsi:.0f})"
                elif rsi_direction == 'Falling':
                    tech_score -= 1
                    rsi_signal = f"Softening ↓ ({rsi_5bar:.0f}→{rsi:.0f})"
                elif rsi_divergence == 'Bullish Divergence':
                    tech_score = min(tech_score + 1, 6)
                    rsi_signal = "Bullish Divergence ✅"
                else:
                    rsi_signal = f"Healthy ({rsi:.0f})"

            if macd > signal:
                tech_score += 1;  macd_signal = "Bullish"
            else:
                tech_score -= 1;  macd_signal = "Bearish"

            # FIX-2 + V52-1: ADX trend strength - now direction-aware.
            # ADX bonus only when price > SMA50 (confirms uptrend direction).
            # A strong downtrend has high ADX but should NOT be rewarded.
            if adx > 25:
                if current_price > sma_50:
                    tech_score = min(tech_score + 1, 6)   # strong uptrend ✅
                else:
                    tech_score -= 1   # strong downtrend - penalise, not reward
            elif adx < 20:
                tech_score -= 1   # FIX-2: weak/no trend penalty retained

            # FIX-6: Volume influences tech score
            if vol_ratio > 1.5 and current_price > sma_20:
                tech_score = min(tech_score + 1, 6)
            elif vol_ratio < 0.7:
                tech_score -= 1

            # FIX-8: Near 52W high in uptrend
            pct_from_52w_high = ((current_price - high_52w) / high_52w) * 100
            if pct_from_52w_high >= -5 and current_price > sma_200:
                tech_score = min(tech_score + 1, 6)
            # =================================================================

            # =================================================================
            # V58: TECH BUY SCORE — 100% pure technical (0-100).
            # Zero fundamentals. Scored for ALL 100 stocks.
            # RSI (25) + MACD (20) + SMA (20) + Volume (15) + ADX (10) + Div (10)
            # =================================================================
            tech_buy_score = 0

            # RSI zone + direction (0-25)
            if rsi < 30 and rsi_direction == 'Rising':
                tech_buy_score += 25
            elif rsi < 30:
                tech_buy_score += 18
            elif rsi < 40 and rsi_direction == 'Rising':
                tech_buy_score += 22
            elif rsi < 40:
                tech_buy_score += 12
            elif 40 <= rsi <= 55 and rsi_direction == 'Rising':
                tech_buy_score += 16
            elif 40 <= rsi <= 55:
                tech_buy_score += 6
            elif rsi <= 65 and rsi_direction == 'Rising':
                tech_buy_score += 10
            elif rsi <= 65:
                tech_buy_score += 4

            # MACD crossover (0-20)
            if macd > signal:
                tech_buy_score += 20 if (macd - signal) > 0.5 * atr else 14
            elif abs(macd - signal) < 0.2 * atr:
                tech_buy_score += 5

            # SMA structure (0-20)
            if current_price > sma_200: tech_buy_score += 8
            if current_price > sma_50:  tech_buy_score += 6
            if current_price > sma_20:  tech_buy_score += 6

            # Volume confirmation (0-15)
            if vol_ratio >= 1.5 and current_price > sma_20:
                tech_buy_score += 15
            elif vol_ratio >= 1.5:
                tech_buy_score += 10
            elif vol_ratio >= 1.2:
                tech_buy_score += 7
            elif vol_ratio >= 0.8:
                tech_buy_score += 3

            # ADX trend strength (0-10)
            if adx > 30 and current_price > sma_50:
                tech_buy_score += 10
            elif adx > 25 and current_price > sma_50:
                tech_buy_score += 7
            elif adx > 20:
                tech_buy_score += 4
            elif adx > 15:
                tech_buy_score += 2

            # RSI divergence (0-10)
            if rsi_divergence == 'Bullish Divergence':
                tech_buy_score += 10
            elif rsi_divergence == 'Bearish Divergence':
                tech_buy_score -= 5

            tech_buy_score = min(max(tech_buy_score, 0), 100)
            # =================================================================

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
            # FIX-SECTOR: Use yFinance sector if available; fall back to our
            # hardcoded map if yFinance returns None or empty string.
            # This prevents all 'None-sector' stocks from sharing one 'N/A'
            # bucket and bypassing the per-sector diversity cap.
            _yf_sector = info.get('sector', None)
            sym_key    = symbol.replace('.NS', '')
            sector     = (_yf_sector if _yf_sector
                          else self.sector_fallback.get(sym_key, 'N/A'))

            analyst_key   = info.get('recommendationKey', 'N/A')
            analyst_map   = {
                'strongBuy': 'Strong Buy', 'buy': 'Buy',
                'hold': 'Hold', 'sell': 'Sell', 'strongSell': 'Strong Sell'
            }
            analyst_label = analyst_map.get(
                analyst_key,
                analyst_key.title() if analyst_key else 'N/A')
            earnings_date = self.get_earnings_date(info)

            fund_score = self.get_fundamental_score(info, sector)

            # V54-4: Dynamic weight shift.
            # Count how many confirmed bearish technical signals are active RIGHT NOW.

            # V56-2: RSI POST-OVERBOUGHT PULLBACK signal.
            # Detects the Lupin pattern: RSI peaked above 70 recently (within 15 bars)
            # and has now started falling. This is a classic distribution/exit signal.
            # A stock that was overbought and is now fading carries more downside risk
            # than a stock that was never overbought — the unwinding is not yet complete.
            try:
                delta_s    = df['Close'].diff()
                gain_s     = delta_s.where(delta_s > 0, 0)
                loss_s     = (-delta_s.where(delta_s < 0, 0))
                ag_s       = gain_s.ewm(com=13, min_periods=14).mean()
                al_s       = loss_s.ewm(com=13, min_periods=14).mean()
                rsi_full   = 100 - (100 / (1 + ag_s / al_s))
                rsi_full   = rsi_full.dropna()
                # Was RSI above 70 at any point in the last 15 bars?
                rsi_recent_peak = rsi_full.iloc[-16:-1].max() if len(rsi_full) >= 16 else rsi_full.max()
                rsi_post_ob_pullback = (rsi_recent_peak > 70 and rsi_direction == 'Falling')
            except Exception:
                rsi_post_ob_pullback = False

            bearish_signal_count = sum([
                bool(sma_20_declining),                          # SMA20 rolling over
                bool(death_cross_forming),                       # SMA20 < SMA50
                bool(macd < signal),                             # MACD bearish crossover
                bool(rsi < 50),                                  # momentum lost
                bool(current_price < sma_50),                   # below medium trend
                bool(rsi_direction == 'Falling' and rsi_slope_strong),  # V55: RSI falling fast (Power Grid case)
                bool(rsi_post_ob_pullback),                      # V56-2: RSI came from overbought and falling (Lupin case)
            ])
            # V54-4: Dynamic weight shift at 2+ signals (lower than veto threshold of 3).
            # This means the score already shifts to 50/50 before the veto fires,
            # correctly penalising weakening stocks in the score WITHOUT removing them.
            # Only at 3+ signals does the veto hard-cap them to HOLD.
            if bearish_signal_count >= 2:
                tech_weight  = 0.50
                fund_weight  = 0.50
                weight_label = "50/50 (Downtrend Override)"
            else:
                tech_weight  = 0.35
                fund_weight  = 0.65
                weight_label = "35/65 (Normal)"

            # FIX-1: Combined score with dynamic weights
            tech_score_normalized = ((tech_score + 6) / 12) * 100
            combined_score        = (tech_score_normalized * tech_weight) + (fund_score * fund_weight)

            # FIX-7 + V53-3 + V54-3: Analyst consensus +/-5.
            # V54-3 tightens the buy gate: tech_score must be >= 2 (was > 0).
            # A score of +1 is borderline and should NOT rescue a weak chart.
            # Sell/strongSell penalty always applies unconditionally.
            if analyst_key in ('strongBuy', 'buy'):
                if tech_score >= 2:   # V54-3: tightened from > 0 to >= 2
                    combined_score = min(combined_score + 5, 100)
                # tech_score < 2: analyst buy silently ignored - chart disagrees
            elif analyst_key in ('sell', 'strongSell'):
                combined_score = max(combined_score - 5, 0)

            # FIX-MOMENTUM: Daily momentum delta component.
            # Fund score and most tech signals are stable for days/weeks.
            # Vol_Ratio and RSI_Slope are the two signals that genuinely
            # change EVERY day — adding a small component based on these
            # means stocks with improving momentum get a daily boost while
            # stocks with deteriorating momentum get penalised.
            # Effect is ±5 points max — big enough to reorder borderline
            # stocks, small enough not to override fundamental quality.
            #
            # Vol_Ratio > 1.5 in uptrend  → +2.5 (real buying interest today)
            # Vol_Ratio > 1.2 in uptrend  → +1.0 (mild interest)
            # Vol_Ratio < 0.7             → -1.5 (thin, no conviction)
            # RSI rising fast (slope>5)   → +2.5 (momentum building)
            # RSI falling fast (slope<-5) → -2.5 (momentum fading)
            momentum_delta = 0.0
            if vol_ratio > 1.5 and current_price > sma_20:
                momentum_delta += 2.5
            elif vol_ratio > 1.2 and current_price > sma_20:
                momentum_delta += 1.0
            elif vol_ratio < 0.7:
                momentum_delta -= 1.5
            if rsi_slope > 5:
                momentum_delta += 2.5
            elif rsi_slope < -5:
                momentum_delta -= 2.5
            # momentum_delta intentionally NOT applied to combined_score here.
            # It is applied only to ranking_score so it reorders within a tier
            # but never demotes a stock from BUY to HOLD on a weak-volume day.
            combined_score  = round(combined_score, 1)
            ranking_score   = round(min(max(combined_score + momentum_delta, 0), 100), 1)

            # CAL-1: Thresholds relaxed to account for missing yFinance fields
            # on NSE stocks (PEG/ROA/CR often return None, silently scoring 0).
            # STRONG BUY: 75->70  |  BUY: 55->50
            if combined_score >= 70:
                rating = "⭐⭐⭐⭐⭐ STRONG BUY";  recommendation = "STRONG BUY"
            elif combined_score >= 50:
                rating = "⭐⭐⭐⭐ BUY";           recommendation = "BUY"
            elif combined_score >= 40:
                rating = "⭐⭐⭐ HOLD";            recommendation = "HOLD"
            elif combined_score >= 28:
                rating = "⭐⭐ SELL";              recommendation = "SELL"
            else:
                rating = "⭐ STRONG SELL";         recommendation = "STRONG SELL"

            # V54-1: TREND VETO GATE — hard cap BEFORE stop/target calculation.
            # Threshold: 3+ bearish signals → max rating is HOLD.
            #
            # WHY 3 NOT 2:
            # In a broad market correction (Nifty down 1-2%), almost EVERY stock
            # will have at least 2 bearish signals (usually MACD bear + RSI<50).
            # A threshold of 2 was vetoing ALL BUY stocks in bear market conditions,
            # leaving only 5-6 in the table. That defeats the purpose.
            #
            # Threshold of 3 means we need CONFIRMED multi-signal deterioration:
            # e.g. SMA20 declining + death cross + MACD bearish = real distribution top
            # e.g. RSI falling fast + MACD bearish + price < SMA50 = confirmed downtrend
            # Just having RSI<50 + MACD bearish on a down day is NOT enough to veto.
            #
            # The 6-signal set (including RSI slope) means a stock can still get
            # vetoed with 3 of: SMA declining, death cross, MACD bear, RSI<50,
            # price<SMA50, RSI falling fast — which is a genuine distribution pattern.
            veto_fired = bearish_signal_count >= 3 and recommendation in ("STRONG BUY", "BUY")
            if veto_fired:
                recommendation = "HOLD"
                rating         = "⭐⭐⭐ HOLD (Veto)"

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
                # ── LONG setup: stop below support, targets at resistance ──
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

            elif recommendation == "HOLD":
                # ── HOLD setup: neutral — nearest support as soft floor,
                #    nearest resistance as soft ceiling. No directional bias.
                #    Stop = ATR below nearest support (risk management).
                #    Target = nearest resistance above price.
                #    Upside shown as % to resistance so user can judge entry.
                atr_stop      = nearest_support - (atr * atr_multiplier)
                min_allowed_sl = current_price * (1 - max_sl_pct / 100)
                stop_loss     = max(atr_stop, min_allowed_sl)
                sl_percentage = ((current_price - stop_loss) / current_price) * 100
                stop_type     = "ATR Stop" if atr_stop >= min_allowed_sl else "Beta Cap"

                # Target: nearest resistance above price (or project +3%)
                valid_res = [r['level'] for r in resistance_levels
                             if r['level'] > current_price * 1.005]
                if len(valid_res) >= 2:
                    target_1, target_2 = valid_res[0], valid_res[1]
                    target_status = "Hold S/R Levels"
                elif len(valid_res) == 1:
                    target_1      = valid_res[0]
                    target_2      = round(target_1 * 1.03, 2)
                    target_status = "Hold Partial S/R"
                else:
                    target_1      = round(current_price * 1.03, 2)
                    target_2      = round(current_price * 1.06, 2)
                    target_status = "Hold Projected"
                targets_hit   = 0
                upside        = ((target_1 - current_price) / current_price) * 100

            else:
                # ── SHORT/SELL setup: stop above resistance, targets at support ──
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

            # FIX-4: STRONG BUY needs R:R ≥ 1.5 (only applies to non-vetoed buys)
            if recommendation == "STRONG BUY" and risk_reward < 1.5:
                recommendation = "BUY"
                rating         = "⭐⭐⭐⭐ BUY"

            # V56-1: R:R FLOOR — BUY needs R:R ≥ 1.0.
            # If reward < risk, the trade setup is unfavourable regardless of score.
            # Lupin case: R:R 0.9x with MACD bearish + RSI falling = bad entry point.
            # This is the cleanest filter: a sub-1.0 R:R means stop loss > potential gain.
            # Cap to HOLD so the stock stays on watchlist but is not flagged as actionable.
            if recommendation == "BUY" and risk_reward < 1.0:
                recommendation = "HOLD"
                rating         = "⭐⭐⭐ HOLD (R:R < 1.0)"

            if fund_score >= 80:   quality = "Excellent"
            elif fund_score >= 60: quality = "Good"
            elif fund_score >= 40: quality = "Average"
            else:                  quality = "Poor"

            return {
                'Symbol':            symbol.replace('.NS', ''),
                'Name':              name,
                'Price':             round(current_price, 2),
                'Sector':            sector,
                'RSI':               round(rsi, 2),
                'RSI_Signal':        rsi_signal,
                'RSI_Direction':     rsi_direction,            # 'Rising'|'Falling'|'Flat'
                'RSI_Slope':         rsi_slope,                # numeric: e.g. -25 for Power Grid
                'RSI_5Bar':          rsi_5bar,                 # RSI 5 bars ago
                'RSI_15Bar':         rsi_15bar,                # RSI 15 bars ago
                'RSI_Peak_15':       rsi_peak_15,              # highest RSI in last 15 bars — used for post-OB gate
                'RSI_Divergence':    rsi_divergence,          # shown in buy table + watchlist
                'MACD':              macd_signal,
                'ADX':               adx,
                'Vol_Ratio':         vol_ratio,
                'SMA_20':            round(sma_20, 2),        # used by watchlist sma_trend
                'SMA_50':            round(sma_50, 2),        # used by watchlist sma_trend
                'SMA_200':           round(sma_200, 2),       # used by watchlist sma_trend
                'SMA_20_Declining':  sma_20_declining,        # used by watchlist sma_trend
                'Death_Cross':       death_cross_forming,     # used by watchlist sma_trend
                'SMA_200_Rising':    sma_200_rising,          # used by watchlist sma_trend
                'Bearish_Signals':   bearish_signal_count,
                'Weight_Mode':       weight_label,
                'Veto_Fired':        veto_fired,              # shown as badge in watchlist
                'Support':           round(nearest_support, 2),
                'Resistance':        round(nearest_resistance, 2),
                'Support_Dist_Pct':  support_dist_pct,        # shown in buy table
                '52W_High':          round(high_52w, 2),
                'Pct_From_52W_High': round(pct_from_52w_high, 2),  # shown in sell table
                'Tech_Score':        tech_score,              # shown in watchlist
                'ATR':               atr,
                'ATR_Pct':           atr_pct,
                'ATR_Multiplier':    atr_multiplier,
                'Stop_Type':         stop_type,
                'PE_Ratio':          round(pe_ratio, 2)           if pe_ratio else 0,
                'Profit_Margin':     round(profit_margin * 100, 2) if profit_margin else 0,
                'Dividend_Yield':    round(dividend_yield * 100, 2) if dividend_yield else 0,
                'Beta':              round(beta, 2)               if beta else 1.0,
                'Fund_Score':        round(fund_score, 1),        # shown in watchlist
                'Quality':           quality,
                'Combined_Score':    round(combined_score, 1),
                'Ranking_Score':     round(ranking_score, 1),   # combined_score + momentum_delta (for sorting only)
                'Momentum_Delta':    round(momentum_delta, 1),  # FIX-MOMENTUM: daily change component
                'Tech_Buy_Score':    tech_buy_score,            # V58: Pure technical buy score
                'Rating':            rating,
                'Recommendation':    recommendation,
                'Stop_Loss':         round(stop_loss, 2),
                'SL_Percentage':     round(sl_percentage, 2),
                'Target_1':          round(target_1, 2),
                'Target_2':          round(target_2, 2),
                'Upside':            round(upside, 2),
                'Risk_Reward':       risk_reward,
                'Target_Status':     target_status,
                'Analyst':           analyst_label,
                'Earnings_Date':     earnings_date,
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
    #  v5: NEW-2 volume hard gate, NEW-3 sector cap, NEW-5 R:R raised to 1.5
    # =========================================================================
    def get_top_recommendations(self):
        df = pd.DataFrame(self.results)

        # == BUY side ==========================================================
        # FIX-ADAPTIVE-POOL: When the market is in a broad correction, many
        # fundamentally sound stocks score 45-49 (just below the 50pt BUY line)
        # purely because their RSI is temporarily weak. On those days the BUY
        # table collapses to 3-4 names. We detect this by checking what % of
        # all analyzed stocks are rated BUY/STRONG BUY right now:
        #   >= 25% rated BUY  → normal market, use strict 50pt threshold
        #   15-24% rated BUY  → mild correction, include HOLD stocks scoring 46+
        #   < 15%  rated BUY  → broad selloff, include HOLD stocks scoring 43+
        # This means in bear conditions we pull in near-BUY HOLDs as
        # "Watchlist Buys" so traders still have 8-12 names to review.
        total_analyzed = len(df)
        raw_buys = df[df['Recommendation'].isin(['STRONG BUY', 'BUY'])]
        buy_pct = len(raw_buys) / max(total_analyzed, 1)

        if buy_pct >= 0.25:
            # Normal market — strict BUY only
            all_buys = raw_buys
            print(f"  Market mode: NORMAL ({buy_pct*100:.0f}% BUY-rated) — strict threshold")
        elif buy_pct >= 0.15:
            # Mild correction — include near-BUY HOLDs (score 46+)
            near_buys = df[(df['Recommendation'] == 'HOLD') & (df['Combined_Score'] >= 46)]
            all_buys  = pd.concat([raw_buys, near_buys]).drop_duplicates()
            print(f"  Market mode: MILD CORRECTION ({buy_pct*100:.0f}% BUY-rated) — including HOLD 46+ as candidates")
        else:
            # Broad selloff — include near-BUY HOLDs (score 40+)
            near_buys = df[(df['Recommendation'] == 'HOLD') & (df['Combined_Score'] >= 40)]
            all_buys  = pd.concat([raw_buys, near_buys]).drop_duplicates()
            print(f"  Market mode: BROAD SELLOFF ({buy_pct*100:.0f}% BUY-rated) — including HOLD 40+ as candidates")

        f1 = all_buys[all_buys['Upside'] > 0]
        f2 = f1[f1['Target_1'] > f1['Price']]

        # ── DEBUG: show exactly why each BUY stock passes or fails each filter ──
        print(f"\n{'─'*75}")
        print(f"  BUY FILTER DEBUG  ({len(all_buys)} BUY-rated stocks → target: show all that qualify)")
        print(f"{'─'*75}")
        print(f"  {'Stock':<18} {'RSI':>5} {'Dir':<8} {'Slp':>5} {'RR':>5} {'Vol':>5}  {'Result'}")
        print(f"  {'─'*70}")
        for _, row in all_buys.sort_values('Combined_Score', ascending=False).iterrows():
            rsi_val = row.get('RSI', 50)
            rsi_dir = row.get('RSI_Direction', 'Flat')
            rsi_slp = row.get('RSI_Slope', 0)
            rr      = row.get('Risk_Reward', 0)
            vol     = row.get('Vol_Ratio', 0)
            upside  = row.get('Upside', 0)
            t1      = row.get('Target_1', 0)
            price   = row.get('Price', 0)

            fail = None
            if upside <= 0:               fail = f"Upside ≤0 ({upside:.1f}%)"
            elif t1 <= price:             fail = f"T1≤Price (T1={t1:.0f} P={price:.0f})"
            elif rsi_val > 70:            fail = f"RSI overbought ({rsi_val:.0f})"
            elif rsi_val > 65 and rsi_dir == 'Falling': fail = f"RSI {rsi_val:.0f} near-OB + falling"
            elif rsi_val > 60 and rsi_slp < -8: fail = f"RSI {rsi_val:.0f} topping (slope {rsi_slp:.0f})"
            elif rsi_val < 48 and rsi_dir == 'Falling' and row.get('RSI_Peak_15', row.get('RSI_5Bar',50)) > 65:
                fail = f"Post-OB collapse (peak15={row.get('RSI_Peak_15', row.get('RSI_5Bar',50)):.0f}→now {rsi_val:.0f} Falling)"
            elif rr < 0.8:                fail = f"R:R too low ({rr:.2f}x)"
            elif vol < 0.7:               fail = f"Vol too low ({vol:.1f}x)"
            else:                         fail = None

            status = f"❌ BLOCKED — {fail}" if fail else "✅ PASS"
            print(f"  {row['Symbol']:<18} {rsi_val:>5.0f} {rsi_dir:<8} {rsi_slp:>+5.0f} {rr:>5.2f} {vol:>5.1f}  {status}")
        print(f"{'─'*75}\n")
        # ── END DEBUG ──────────────────────────────────────────────────────────

        # V55-RSI-GATE + V57-POST-OB-COLLAPSE
        def rsi_is_safe_to_buy(row):
            rsi_val  = row.get('RSI', 50)
            rsi_dir  = row.get('RSI_Direction', 'Flat')
            rsi_slp  = row.get('RSI_Slope', 0)
            rsi_5bar = row.get('RSI_5Bar', 50)

            # Gate 1: Currently overbought
            if rsi_val > 70:
                return False

            # Gate 2: Near overbought and rolling over
            if rsi_val > 65 and rsi_dir == 'Falling':
                return False

            # Gate 3: Sharply falling from elevated RSI
            if rsi_val > 60 and rsi_slp < -8:
                return False

            # Gate 4 — V57: Post-overbought collapse gate (the Torrent pattern).
            # Uses RSI_Peak_15 = the HIGHEST RSI seen in the last 15 bars (3 weeks).
            # If that peak was above 60 but RSI is now below 50 and still falling,
            # the stock is in active distribution from its recent high.
            # WHY peak not 5-bar: Torrent's RSI peaked ~75 weeks ago; 5 bars ago
            # RSI was already 50 → the 5-bar check missed it completely.
            # RSI_Peak_15 catches the actual top regardless of when exactly it occurred.
            # Only block genuine overbought collapses:
            # peak must have been 65+ (clearly overbought, not just "elevated")
            # RSI must now be 48- (not just touching 50 on a minor dip)
            # This stops blocking stocks like RSI 62→49 which is a normal pullback
            rsi_peak = row.get('RSI_Peak_15', row.get('RSI_5Bar', 50))
            if rsi_val < 48 and rsi_dir == 'Falling' and rsi_peak > 65:
                return False

            return True

        f2 = f2[f2.apply(rsi_is_safe_to_buy, axis=1)]

        # FIX-FRESHNESS: Momentum freshness gate.
        # Stocks whose RSI slope and vol_ratio have barely moved today are
        # structurally identical to yesterday's list. By preferring stocks
        # with fresh momentum signals (RSI slope changed meaningfully, or
        # volume picked up), we naturally rotate the list as market action
        # shifts — without removing any "always good" stock from the pool.
        #
        # How it works: we split the qualified buy pool into two tiers:
        #   Tier 1 (fresh)  — RSI slope > +2 OR vol_ratio > 1.2 today
        #   Tier 2 (stable) — everything else that still qualifies
        # We fill the top-20 from Tier 1 first, then backfill with Tier 2.
        # This means a stock with fresh momentum is always preferred over
        # an equally-scored stock with flat signals.
        def has_fresh_momentum(row):
            return (row.get('RSI_Slope', 0) > 2 or
                    row.get('Vol_Ratio', 0) > 1.2)

        f2_fresh  = f2[f2.apply(has_fresh_momentum, axis=1)]
        f2_stable = f2[~f2.index.isin(f2_fresh.index)]

        # R:R gate applied to both tiers independently.
        # In correction markets vol is structurally lower — 0.5x is still
        # liquid enough to trade. R:R floor of 0.5 = reward at least half the risk.
        def apply_rr_vol_gate(pool):
            sb = pool[(pool['Recommendation'] == 'STRONG BUY') & (pool['Risk_Reward'] >= 1.2)]
            pb = pool[(pool['Recommendation'] == 'BUY')        & (pool['Risk_Reward'] >= 0.5)]
            sb = sb[sb['Vol_Ratio'] >= 0.5]
            pb = pb[pb['Vol_Ratio'] >= 0.5]
            return pd.concat([sb, pb]).drop_duplicates()

        fresh_buys  = apply_rr_vol_gate(f2_fresh)
        stable_buys = apply_rr_vol_gate(f2_stable)

        # Sort each tier by Ranking_Score (combined + momentum delta) descending
        fresh_buys  = fresh_buys.sort_values('Ranking_Score', ascending=False)
        stable_buys = stable_buys.sort_values('Ranking_Score', ascending=False)

        # Merge: Tier 1 first, then backfill with Tier 2, no duplicates
        filtered_buys = pd.concat([fresh_buys, stable_buys]).drop_duplicates()
        sorted_buys   = filtered_buys  # already ordered by tier then score

        print(f"  Filter summary: {len(all_buys)} BUY rated → "
              f"{len(f2_fresh)+len(f2_stable)} after RSI gate → "
              f"fresh tier {len(fresh_buys)} | stable tier {len(stable_buys)} → "
              f"sector cap applies next\n")

        # Sector diversity cap: max 4 per sector
        top_buys_rows = []
        sector_counts = {}
        for _, row in sorted_buys.iterrows():
            sec = row.get('Sector', 'N/A')
            sector_counts[sec] = sector_counts.get(sec, 0)
            if sector_counts[sec] < 4:
                top_buys_rows.append(row)
                sector_counts[sec] += 1
            if len(top_buys_rows) >= 20:
                break
        top_buys = pd.DataFrame(top_buys_rows)

        # MINIMUM GUARANTEE: Always show at least 10 stocks.
        # If strict filters leave fewer than 10, backfill with the next-best
        # HOLDs by ranking_score, skipping any already in the table.
        # These are labelled differently in the HTML (amber instead of green)
        # so the user knows they are "watch, not buy" candidates.
        if len(top_buys) < 10:
            already_shown = set(top_buys['Symbol'].tolist()) if len(top_buys) > 0 else set()
            backfill_pool = df[
                (~df['Symbol'].isin(already_shown)) &
                (df['Combined_Score'] >= 38) &
                (df['Recommendation'].isin(['HOLD', 'BUY', 'STRONG BUY'])) &
                (df['Upside'] > 0) &
                (df['RSI'] >= 35) &
                (df['RSI'] <= 68)
            ].sort_values('Ranking_Score', ascending=False)
            needed = 10 - len(top_buys)
            backfill_rows = []
            bf_sector_counts = dict(sector_counts)
            for _, row in backfill_pool.iterrows():
                sec = row.get('Sector', 'N/A')
                bf_sector_counts[sec] = bf_sector_counts.get(sec, 0)
                if bf_sector_counts[sec] < 4:
                    backfill_rows.append(row)
                    bf_sector_counts[sec] += 1
                if len(backfill_rows) >= needed:
                    break
            if backfill_rows:
                backfill_df = pd.DataFrame(backfill_rows)
                backfill_df = backfill_df.copy()
                backfill_df['Rating']         = backfill_df['Rating'].apply(
                    lambda r: r + ' [WATCH]' if '[WATCH]' not in str(r) else r)
                backfill_df['Recommendation'] = 'WATCH'
                top_buys = pd.concat([top_buys, backfill_df], ignore_index=True)
                print(f"  Minimum guarantee: added {len(backfill_rows)} WATCH stocks to reach 10 total")

        # == SELL side =========================================================
        all_sells = df[df['Recommendation'].isin(['STRONG SELL', 'SELL'])]
        s1 = all_sells[all_sells['Upside'] > 0.5]
        s2 = s1[s1['Risk_Reward'] >= 1.2]
        s3 = s2[s2['Target_1'] < s2['Price']]

        # V55-SELL-GATE: Mirror of RSI buy gate — remove false sells.
        #
        # Problem 1 — OVERSOLD stocks in sell table (TCS RSI 31, Persistent RSI 28):
        #   RSI < 35 = stock is already heavily sold. Recommending SELL here means
        #   you're shorting at the bottom = maximum bounce risk. Remove these.
        #   The stock may still be fundamentally weak (low score) but technically
        #   it's not a safe short entry right now. Wait for RSI to recover to 45+
        #   before a fresh short setup forms.
        #
        # Problem 2 — Bullish MACD with healthy RSI (Divi's Lab RSI 62, MACD Bullish):
        #   If RSI is above 50 AND MACD is Bullish, the chart is saying the stock
        #   still has upward momentum. A low combined score means poor fundamentals,
        #   but that alone is not a sell signal while price action is rising.
        #   These belong in HOLD, not SELL table.
        #
        # Problem 3 — RSI rising fast from oversold:
        #   If RSI slope is strongly positive (rising > 8 points in 5 bars),
        #   the stock is recovering. Don't short a recovering stock.

        def sell_is_valid(row):
            rsi_val = row.get('RSI', 50)
            rsi_dir = row.get('RSI_Direction', 'Flat')
            rsi_slp = row.get('RSI_Slope', 0)
            macd    = row.get('MACD', 'Bearish')

            # Block 1: Oversold — bounce risk, not a short entry
            if rsi_val < 35:
                return False

            # Block 2: RSI healthy (>50) AND MACD Bullish — chart says uptrend
            if rsi_val > 50 and macd == 'Bullish':
                return False

            # Block 3: RSI rising fast from a low base — recovery in progress
            if rsi_val < 45 and rsi_dir == 'Rising' and rsi_slp > 8:
                return False

            return True

        top_sells = s3[s3.apply(sell_is_valid, axis=1)].nsmallest(20, 'Combined_Score')

        return top_buys, top_sells

    # =========================================================================
    #  ACCUMULATE ON DIP  — V57
    #  Stocks scoring 42-49: just below BUY threshold but fundamentally sound.
    #  Criteria:
    #    · Combined score 42-49  (just missed BUY cutoff of 50)
    #    · Fund score >= 50       (decent fundamentals — not junk)
    #    · RSI between 35-65     (not overbought, not in freefall)
    #    · Quality != 'Poor'
    #    · Max 15 stocks, max 3 per sector
    # =========================================================================
    def get_accumulate_watchlist(self):
        df = pd.DataFrame(self.results)
        acc = df[
            (df['Combined_Score'] >= 42) &
            (df['Combined_Score'] < 50) &
            (df['Fund_Score'] >= 50) &
            (~df['Recommendation'].isin(['STRONG BUY', 'BUY'])) &
            (df['RSI'] >= 35) &
            (df['RSI'] <= 65) &
            (df['Quality'] != 'Poor')
        ].copy()
        acc = acc.sort_values(['Fund_Score', 'Combined_Score'], ascending=False)
        rows, sector_counts = [], {}
        for _, row in acc.iterrows():
            sec = row.get('Sector', 'N/A')
            sector_counts[sec] = sector_counts.get(sec, 0)
            if sector_counts[sec] < 3:
                rows.append(row)
                sector_counts[sec] += 1
            if len(rows) >= 15:
                break
        return pd.DataFrame(rows)

    # =========================================================================
    #  V58: TOP 10 TECHNICAL-ONLY BUY PICKS
    #  Reads ALL 100 stocks. Ranks purely by Tech_Buy_Score (0-100).
    #  100% technical — RSI, MACD, SMA, ADX, Volume. Zero fundamentals.
    #  NO veto gate, NO R:R gate, NO RSI safe-to-buy gate.
    #
    #  CRITICAL: The existing Target_1/Stop_Loss/Upside in self.results are
    #  DIRECTIONAL — they depend on the TechnoFunc recommendation.
    #  A stock rated SELL has short-side targets (Target_1 < Price).
    #  We must RECALCULATE buy-side targets from Support/Resistance.
    # =========================================================================
    def get_top_technical_recommendations(self):
        df = pd.DataFrame(self.results)
        if df.empty:
            return pd.DataFrame()

        # Only filter: RSI not extremely overbought
        pool = df[
            (df['Tech_Buy_Score'] > 15) &
            (df['RSI'] <= 78)
        ].copy()

        # ── Recalculate BUY-SIDE trade setup for every stock ──────────────
        # Target = Resistance above price (where price can rally TO)
        # Stop   = Support - ATR * multiplier (where to cut loss)
        # This is independent of TechnoFunc recommendation direction.
        buy_targets  = []
        buy_targets2 = []
        buy_stops    = []
        buy_sl_pcts  = []
        buy_upsides  = []
        buy_rrs      = []

        for _, row in pool.iterrows():
            price      = row['Price']
            resistance = row.get('Resistance', price * 1.03)
            support    = row.get('Support', price * 0.95)
            atr_val    = row.get('ATR', price * 0.02)
            atr_mult   = row.get('ATR_Multiplier', 1.2)

            # Buy target = resistance (above price)
            t1 = resistance if resistance > price * 1.005 else round(price * 1.03, 2)
            t2 = round(t1 * 1.04, 2)

            # Buy stop = below support
            sl = max(support - (atr_val * atr_mult), price * 0.90)

            sl_pct  = round(((price - sl) / price) * 100, 2)
            upside  = round(((t1 - price) / price) * 100, 2)
            risk    = abs(price - sl)
            reward  = abs(t1 - price)
            rr      = round(reward / risk, 2) if risk > 0 else 0

            buy_targets.append(round(t1, 2))
            buy_targets2.append(round(t2, 2))
            buy_stops.append(round(sl, 2))
            buy_sl_pcts.append(sl_pct)
            buy_upsides.append(upside)
            buy_rrs.append(rr)

        pool['Buy_Target_1']    = buy_targets
        pool['Buy_Target_2']    = buy_targets2
        pool['Buy_Stop_Loss']   = buy_stops
        pool['Buy_SL_Pct']      = buy_sl_pcts
        pool['Buy_Upside']      = buy_upsides
        pool['Buy_RR']          = buy_rrs

        # Filter: buy upside must be positive (resistance above price)
        pool = pool[pool['Buy_Upside'] > 0]

        # Sort ALL stocks by Tech_Buy_Score descending
        pool = pool.sort_values('Tech_Buy_Score', ascending=False)

        # Sector diversity cap: max 3 per sector
        top_rows = []
        sector_counts = {}
        for _, row in pool.iterrows():
            sec = row.get('Sector', 'N/A')
            sector_counts[sec] = sector_counts.get(sec, 0)
            if sector_counts[sec] < 3:
                top_rows.append(row)
                sector_counts[sec] += 1
            if len(top_rows) >= 10:
                break

        result = pd.DataFrame(top_rows)

        # Debug
        print(f"\n{'─'*80}")
        print(f"  V58 TECHNICAL-ONLY TOP 10 — All 100 stocks evaluated, pure chart ranking")
        print(f"  Scoring: RSI(25) + MACD(20) + SMA(20) + Vol(15) + ADX(10) + Div(10) = 100")
        print(f"{'─'*80}")
        if not result.empty:
            for i, (_, r) in enumerate(result.iterrows(), 1):
                tf_rec = r.get('Recommendation', '?')
                print(f"  {i:>2}. {r['Symbol']:<16} TechScore={r['Tech_Buy_Score']:<4}"
                      f"  RSI={r['RSI']:.0f} {r.get('RSI_Direction','?'):<8}"
                      f"  MACD={r['MACD']:<8}  ADX={r.get('ADX',0):.0f}"
                      f"  Vol={r.get('Vol_Ratio',0):.1f}x"
                      f"  BuyTarget=₹{r['Buy_Target_1']:,.0f}"
                      f"  [TF: {tf_rec} @ {r['Combined_Score']:.0f}]")
        print(f"{'─'*80}\n")

        return result

    # =========================================================================
    #  HTML - v5: Divergence column added to Buy table
    # =========================================================================
    def generate_html(self):
        df = pd.DataFrame(self.results)
        top_buys, top_sells = self.get_top_recommendations()
        accumulate_df        = self.get_accumulate_watchlist()
        tech_buys            = self.get_top_technical_recommendations()   # V58

        now         = self.get_ist_time()
        idx_data    = self.fetch_index_data()
        time_of_day = "Morning" if now.hour < 12 else "Evening"
        next_update = "4:30 PM" if now.hour < 12 else "9:30 AM (Next Day)"

        strong_buy_count  = len(df[df['Recommendation'] == 'STRONG BUY'])
        buy_count         = len(df[df['Recommendation'] == 'BUY'])
        hold_count        = len(df[df['Recommendation'] == 'HOLD'])
        sell_count        = len(df[df['Recommendation'] == 'SELL'])
        strong_sell_count = len(df[df['Recommendation'] == 'STRONG SELL'])

        # Sector distribution summary for KPI band
        if not top_buys.empty:
            sector_summary = top_buys['Sector'].value_counts().head(4)
            sector_kpi = ' · '.join([f"{s}: {c}" for s, c in sector_summary.items()])
        else:
            sector_kpi = 'No buys'

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
        ticker_html = ticker_items + ticker_items

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>NIFTY 100 Market Influencers - {time_of_day} Report · {now.strftime('%d %b %Y')}</title>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Syne:wght@700;800&display=swap" rel="stylesheet">
<style>
:root {{
  --bg:       #04080f;
  --bg2:      #060d18;
  --surface:  #0a1628;
  --surface2: #0c1a2e;
  --border:   #1e3a5a;
  --border2:  #2a4a6a;
  --accent:   #00f5ff;
  --accent2:  #00ccee;
  --green:    #00ff88;
  --green2:   #00cc66;
  --red:      #ff4466;
  --gold:     #ffcc00;
  --purple:   #cc99ff;
  --teal:     #00f5ff;
  --text:     #ddeeff;
  --text2:    #ffffff;
  --muted:    #aaccee;
  --muted2:   #ccddff;
}}
*, *::before, *::after {{ margin:0; padding:0; box-sizing:border-box; }}
body {{
  background: #04080f; color: #ddeeff;
  font-family: 'Space Grotesk', sans-serif;
  font-size: 15px; min-height: 100vh;
}}
header {{
  background: #060d18; border-bottom: 2px solid #00f5ff;
  position: sticky; top: 0; z-index: 100;
  box-shadow: 0 4px 24px rgba(0,0,0,0.8);
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
  font-family: 'Syne', sans-serif; font-size: 20px;
  font-weight: 800; color: #ffffff; letter-spacing: -0.5px;
}}
.brand-sub {{
  font-size: 12px; color: #aaddff; letter-spacing: 2px;
  text-transform: uppercase; margin-top: 2px; font-weight: 700;
}}
.idx-strip {{
  display: flex; align-items: center;
  background: rgba(0,0,0,0.4);
  border: 1px solid var(--border2);
  border-radius: 10px; overflow: hidden;
}}
.idx-item {{
  display: flex; flex-direction: column; align-items: center;
  padding: 6px 20px; border-right: 1px solid var(--border); gap: 2px;
}}
.idx-item:last-child {{ border-right: none; }}
.idx-name  {{ font-size: 12px; font-weight: 800; letter-spacing: 2px; color: #aaddff; text-transform: uppercase; }}
.idx-price {{ font-family: 'IBM Plex Mono', monospace; font-size: 18px; font-weight: 800; color: #ffffff; }}
.idx-chg   {{ font-family: 'IBM Plex Mono', monospace; font-size: 14px; font-weight: 800; }}
.idx-chg.up {{ color: #00ff88; text-shadow: 0 0 8px rgba(0,255,136,0.6); }}
.idx-chg.dn {{ color: #ff4466; text-shadow: 0 0 8px rgba(255,68,102,0.6); }}
.clock-box {{ display: flex; flex-direction: column; align-items: flex-end; gap: 3px; }}
.clock-time {{
  font-family: 'IBM Plex Mono', monospace; font-size: 24px;
  font-weight: 800; color: #00ff88;
  text-shadow: 0 0 16px rgba(0,255,136,0.8); letter-spacing: 1px;
}}
.clock-meta {{ font-size: 14px; color: #ffffff; letter-spacing: 1px; font-weight: 700; }}
.clock-next {{ font-size: 13px; color: #aaddff; margin-top: 2px; font-weight: 700; }}
.ticker {{
  background: rgba(0,0,0,0.6); border-top: 1px solid var(--border); overflow: hidden;
}}
.ticker-inner {{
  display: flex; white-space: nowrap;
  animation: ticker-scroll 50s linear infinite; padding: 5px 0;
}}
@keyframes ticker-scroll {{
  0%   {{ transform: translateX(0); }}
  100% {{ transform: translateX(-50%); }}
}}
.tick {{
  display: inline-flex; align-items: center; gap: 6px;
  padding: 0 18px; border-right: 1px solid #1e3a5a;
  font-family: 'IBM Plex Mono', monospace; font-size: 14px;
}}
.tick-sym {{ color: #00f5ff; font-weight: 800; }}
.tick-px  {{ color: #ffffff; font-weight: 600; }}
.tick-up  {{ color: #00ff88; font-weight: 700; }}
.tick-dn  {{ color: #ff4466; font-weight: 700; }}
.kpi-band {{
  display: flex; align-items: center;
  background: #080f1e; border-bottom: 2px solid #1e3a5a;
}}
.kpi-item {{
  display: flex; flex-direction: column; align-items: center;
  padding: 16px 24px; border-right: 1px solid #1e3a5a; flex: 1;
}}
.kpi-item:last-child {{ border-right: none; }}
.kpi-num   {{ font-family: 'Syne', sans-serif; font-size: 40px; font-weight: 800; line-height: 1; }}
.kpi-label {{ font-size: 14px; font-weight: 800; letter-spacing: 2px; text-transform: uppercase; color: #aaddff; margin-top: 5px; }}
.kpi-bar   {{ height: 3px; width: 50px; border-radius: 2px; margin-top: 8px; }}
.kpi-sub   {{ font-size: 12px; color: #6699bb; margin-top: 4px; font-weight: 600; letter-spacing: 0.5px; }}
.main {{ padding: 20px; }}
.section-hdr {{ display: flex; align-items: center; gap: 12px; margin-bottom: 14px; }}
.section-pill {{
  display: flex; align-items: center; gap: 8px;
  padding: 8px 20px; border-radius: 100px;
  font-size: 15px; font-weight: 800; letter-spacing: 0.5px;
}}
.pill-buy  {{ background: #004d25; color: #00ff88; border: 2px solid #00cc66; }}
.pill-sell {{ background: #4d0010; color: #ff4466; border: 2px solid #cc0033; }}
.section-line {{ flex: 1; height: 1px; background: #1e3a5a; }}
.section-note {{ font-size: 13px; color: #88aacc; letter-spacing: 1.5px; white-space: nowrap; font-weight: 800; text-transform: uppercase; }}
.tbl-wrap {{
  width: 100%; overflow-x: auto;
  border: 1px solid #1e3a5a; border-radius: 12px;
  margin-bottom: 28px; box-shadow: 0 8px 40px rgba(0,0,0,0.6);
  -webkit-overflow-scrolling: touch; background: #080f1e;
}}
table {{ width: 100%; border-collapse: collapse; min-width: 1600px; }}
.grp-row th {{
  font-size: 12px; font-weight: 800; letter-spacing: 3px; text-transform: uppercase;
  padding: 9px 10px; text-align: center;
  border-bottom: 1px solid rgba(255,255,255,0.1); white-space: nowrap;
}}
.grp-stock {{ background: #0d3a42; color: #00f5ff; text-shadow: 0 0 8px rgba(0,245,255,0.6); }}
.grp-trade {{ background: #0a3320; color: #00ff88; text-shadow: 0 0 8px rgba(0,255,136,0.6); }}
.grp-tech  {{ background: #0a2a40; color: #40c8ff; text-shadow: 0 0 8px rgba(64,200,255,0.6); }}
.grp-fund  {{ background: #3a2a00; color: #ffcc00; text-shadow: 0 0 8px rgba(255,204,0,0.6); }}
.grp-meta  {{ background: #28124a; color: #cc99ff; text-shadow: 0 0 8px rgba(204,153,255,0.6); }}
.col-row th {{
  font-size: 13px; font-weight: 800; letter-spacing: 1px; text-transform: uppercase;
  padding: 9px 10px; color: #ffffff; background: #0c1a2e;
  border-bottom: 3px solid #1e3a5a; white-space: nowrap; text-align: left;
}}
.ch-stock {{ border-top: 3px solid #00f5ff; color: #b0f0ff; }}
.ch-trade {{ border-top: 3px solid #00ff88; color: #b0ffe0; }}
.ch-tech  {{ border-top: 3px solid #40c8ff; color: #c0e8ff; }}
.ch-fund  {{ border-top: 3px solid #ffcc00; color: #fff0a0; }}
.ch-meta  {{ border-top: 3px solid #cc99ff; color: #e8d0ff; }}
.gsep {{ border-left: 2px solid rgba(255,255,255,0.12) !important; }}
td {{
  padding: 11px 10px; border-bottom: 1px solid #0e2040;
  vertical-align: middle; white-space: nowrap;
}}
tr:last-child td {{ border-bottom: none; }}
tr:nth-child(even) td {{ background: rgba(255,255,255,0.025); }}
tr:hover td {{ background: rgba(0,212,255,0.07); transition: background 0.15s; }}
.stock-name {{ font-size: 15px; font-weight: 700; color: #ffffff; }}
.stock-sym  {{ font-family: 'IBM Plex Mono', monospace; font-size: 12px; color: #00f5ff; font-weight: 700; letter-spacing: 1px; margin-top: 2px; }}
.stock-sec  {{ font-size: 12px; color: #88bbdd; margin-top: 2px; max-width: 130px; overflow: hidden; text-overflow: ellipsis; font-weight: 600; }}
.price-val  {{ font-family: 'IBM Plex Mono', monospace; font-size: 16px; font-weight: 700; color: #ffcc00; }}
.badge {{
  display: inline-flex; align-items: center; gap: 4px;
  font-size: 12px; font-weight: 800; padding: 5px 10px;
  border-radius: 6px; letter-spacing: 0.5px; white-space: nowrap;
}}
.badge-sb {{ background: #004d25; color: #00ff88; border: 1px solid #00ff88; }}
.badge-b  {{ background: #003a4d; color: #00f5ff; border: 1px solid #00f5ff; }}
.badge-h  {{ background: #1a2a3a; color: #aabbcc; border: 1px solid #445566; }}
.badge-s  {{ background: #4d0010; color: #ff4466; border: 1px solid #ff4466; }}
.badge-ss {{ background: #5a0015; color: #ff7788; border: 1px solid #ff7788; }}
.score-wrap {{ display: flex; flex-direction: column; align-items: center; gap: 4px; margin-top: 6px; }}
.score-num  {{ font-family: 'Syne', sans-serif; font-size: 27px; font-weight: 800; line-height: 1; }}
.score-track {{ width: 44px; height: 4px; background: #1a2a3a; border-radius: 2px; }}
.score-fill  {{ height: 100%; border-radius: 2px; transition: width 0.5s ease; }}
.target-badge {{
  font-size: 11px; font-weight: 800; padding: 3px 7px;
  border-radius: 4px; letter-spacing: 0.5px; display: block; margin-bottom: 4px;
}}
.tb-real    {{ background: #004d25; color: #00ff88; border: 1px solid #00ff88; }}
.tb-partial {{ background: #4d3300; color: #ffcc00; border: 1px solid #ffcc00; }}
.tb-ath     {{ background: #003a4d; color: #00f5ff; border: 1px solid #00f5ff; }}
.t1-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 15px; font-weight: 700; color: #ffffff; }}
.t2-val {{ font-size: 13px; color: #88bbdd; margin-top: 2px; font-weight: 700; }}
.sl-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 15px; font-weight: 700; color: #ff4466; }}
.sl-pct {{ font-size: 13px; color: #88bbdd; margin-top: 2px; font-weight: 700; }}
.sl-type {{ font-size: 11px; font-weight: 800; padding: 3px 7px; border-radius: 4px; margin-top: 4px; display: inline-block; }}
.slt-atr  {{ background: #004d25; color: #00ff88; border: 1px solid #00cc66; }}
.slt-beta {{ background: #4d3300; color: #ffcc00; border: 1px solid #cc9900; }}
.upside-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 18px; font-weight: 800; }}
.upside-val.up {{ color: #00ff88; }}
.upside-val.dn {{ color: #ff4466; }}
.rr-val  {{ font-family: 'IBM Plex Mono', monospace; font-size: 16px; font-weight: 800; }}
.atr-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 14px; font-weight: 700; color: #00f5ff; }}
.atr-sub {{ font-size: 12px; color: #88bbdd; margin-top: 2px; font-weight: 700; }}
.rsi-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 16px; font-weight: 700; }}
.rsi-sig {{ font-size: 12px; color: #88bbdd; margin-top: 2px; font-weight: 700; }}
.div-badge {{ font-size: 11px; font-weight: 800; padding: 3px 8px; border-radius: 4px; white-space: nowrap; }}
.div-bear {{ background: #4d0010; color: #ff4466; border: 1px solid #ff4466; }}
.div-bull {{ background: #004d25; color: #00ff88; border: 1px solid #00ff88; }}
.div-none {{ background: #1a2a3a; color: #88aacc; border: 1px solid #334455; }}
.adx-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 16px; font-weight: 700; }}
.adx-lbl {{ font-size: 12px; color: #88bbdd; margin-top: 2px; font-weight: 700; }}
.adx-strong {{ color: #00ff88; }}
.adx-mod    {{ color: #ffcc00; }}
.adx-weak   {{ color: #aabbcc; }}
.vol-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 16px; font-weight: 700; }}
.vol-lbl {{ font-size: 12px; color: #88bbdd; margin-top: 2px; font-weight: 700; }}
.vol-high {{ color: #00ff88; }}
.vol-norm {{ color: #ddeeff; }}
.vol-low  {{ color: #aabbcc; }}
.sdist-val {{ font-family: 'IBM Plex Mono', monospace; font-size: 16px; font-weight: 700; }}
.sdist-close {{ color: #00ff88; }}
.sdist-mid   {{ color: #ffcc00; }}
.sdist-far   {{ color: #ff4466; }}
.mono-sm {{ font-family: 'IBM Plex Mono', monospace; font-size: 15px; font-weight: 700; }}
.qbadge {{ font-size: 11px; font-weight: 800; padding: 4px 9px; border-radius: 5px; }}
.qb-ex {{ background: #004d25; color: #00ff88; border: 1px solid #00cc66; }}
.qb-gd {{ background: #003a4d; color: #00f5ff; border: 1px solid #0099bb; }}
.qb-av {{ background: #4d3300; color: #ffcc00; border: 1px solid #cc9900; }}
.qb-po {{ background: #4d0010; color: #ff4466; border: 1px solid #cc0033; }}
.analyst-badge {{ font-size: 11px; font-weight: 800; padding: 4px 9px; border-radius: 5px; white-space: nowrap; }}
.ab-sb {{ background: #004d25; color: #00ff88; border: 1px solid #00cc66; }}
.ab-b  {{ background: #003a4d; color: #00f5ff; border: 1px solid #0099bb; }}
.ab-h  {{ background: #1a2a3a; color: #aabbcc; border: 1px solid #334455; }}
.ab-s  {{ background: #4d0010; color: #ff4466; border: 1px solid #cc0033; }}
.earn-date {{ font-family: 'IBM Plex Mono', monospace; font-size: 13px; color: #00f5ff; font-weight: 700; }}
.rnum {{ font-size: 14px; color: #88bbdd; font-weight: 700; }}
.macd-bull {{ color: #00ff88; font-weight: 800; font-size: 14px; }}
.macd-bear {{ color: #ff4466; font-weight: 800; font-size: 14px; }}
.disc {{
  background: #0c1a2e; border: 1px solid #1e3a5a;
  border-left: 4px solid #ff4466;
  padding: 14px 18px; border-radius: 8px;
  font-size: 14px; color: #aaccee; line-height: 1.9; margin: 16px 0;
}}
footer {{
  text-align: center; padding: 16px;
  background: #080f1e; border-top: 1px solid #1e3a5a;
  font-size: 14px; color: #88aacc; letter-spacing: 1px;
}}
footer strong {{ color: #00f5ff; }}
@media(max-width: 900px) {{
  .idx-strip {{ display: none; }}
  .kpi-item  {{ padding: 10px 12px; }}
  .kpi-num   {{ font-size: 27px; }}
}}
@media(max-width: 600px) {{
  .h-top  {{ padding: 8px 12px; }}
  .main   {{ padding: 10px; }}
  .kpi-band {{ flex-wrap: wrap; }}
  .kpi-item {{ flex: 0 0 50%; border-bottom: 1px solid var(--border); }}
}}
.mode-toggle {{
  display: flex; align-items: center; gap: 0;
  background: #0a1628; border: 2px solid #1e3a5a; border-radius: 100px;
  padding: 3px;
}}
.mode-btn {{
  padding: 9px 24px; border-radius: 100px; border: none; cursor: pointer;
  font-size: 14px; font-family: 'Space Grotesk', sans-serif; font-weight: 700;
  letter-spacing: 0.5px; transition: all .25s ease; white-space: nowrap;
  background: transparent; color: #556688;
}}
.mode-btn:hover {{ color: #aaccee; }}
.mode-btn.active.tf-active {{
  background: linear-gradient(135deg, #00f5ff, #00cc66);
  color: #000; box-shadow: 0 0 20px rgba(0,245,255,0.3);
}}
.mode-btn.active.t-active {{
  background: linear-gradient(135deg, #cc99ff, #7c4dff);
  color: #fff; box-shadow: 0 0 20px rgba(124,77,255,0.3);
}}
.buy-section {{ display: block; }}
.buy-section.hidden {{ display: none; }}
</style>
</head>
<body>
<header>
  <div class="h-top">
    <div class="brand">
      <div class="brand-gem">💎</div>
      <div>
        <div class="brand-name">NIFTY 100 Market Influencers · NSE &amp; BSE</div>
        <div class="brand-sub">12M S/R · ATR Stops · Trend Veto · Dynamic Weights · SMA200 Slope · v5.4</div>
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

<div class="kpi-band">
  <div class="kpi-item"><div class="kpi-num" style="color:var(--accent)">{len(self.results)}</div><div class="kpi-label">Analyzed</div><div class="kpi-bar" style="background:var(--accent)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:var(--green)">{strong_buy_count}</div><div class="kpi-label">Strong Buy</div><div class="kpi-bar" style="background:var(--green)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:var(--teal)">{buy_count}</div><div class="kpi-label">Buy</div><div class="kpi-bar" style="background:var(--teal)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:var(--red)">{sell_count + strong_sell_count}</div><div class="kpi-label">Sell / Strong Sell</div><div class="kpi-bar" style="background:var(--red)"></div></div>
  <div class="kpi-item"><div class="kpi-num" style="color:#60a5fa">{hold_count}</div><div class="kpi-label">Hold</div><div class="kpi-bar" style="background:#60a5fa"></div><div class="kpi-sub">Sectors: {sector_kpi}</div></div>
</div>

<div class="main">
"""

        # == helper functions ==================================================
        def rating_badge(rec, rating_text):
            cls_map = {
                'STRONG BUY':  'badge-sb', 'BUY': 'badge-b',
                'HOLD':        'badge-h',  'SELL': 'badge-s',
                'STRONG SELL': 'badge-ss',
            }
            return f'<span class="badge {cls_map.get(rec, "badge-h")}">{rating_text}</span>'

        def action_button(rec):
            """Renders the final Action column button.
            Uses the actual recommendation value — never hardcoded.
            Each state has its own colour so the eye immediately sees
            BUY (cyan) / STRONG BUY (green) / HOLD (amber) /
            SELL (red) / STRONG SELL (deep red).
            """
            styles = {
                'STRONG BUY':  ('background:#004d25;color:#00ff88;border:1px solid #00ff88;', '⭐⭐ STRONG BUY'),
                'BUY':         ('background:#003a4d;color:#00f5ff;border:1px solid #00f5ff;', '▲ BUY'),
                'HOLD':        ('background:#2a2200;color:#ffab00;border:1px solid #ffab00;', '◆ HOLD'),
                'SELL':        ('background:#4d0010;color:#ff4466;border:1px solid #ff4466;', '▼ SELL'),
                'STRONG SELL': ('background:#5a0015;color:#ff7788;border:1px solid #ff7788;', '⚠ STRONG SELL'),
            }
            style, label = styles.get(rec, styles['HOLD'])
            return (f'<span style="display:inline-block;padding:4px 10px;border-radius:5px;'
                    f'font-size:13px;font-weight:700;letter-spacing:.4px;{style}">{label}</span>')

        def veto_badge(bearish_signals, weight_mode):
            """Shows a small warning pill when the Trend Veto Gate fired,
            so the user can instantly see WHY a score looks high but the
            action is HOLD — fundamentals were good but trend vetoed it."""
            if bearish_signals >= 3:
                return (f'<div style="margin-top:3px;display:inline-block;padding:2px 6px;'
                        f'border-radius:3px;background:#2a1500;color:#ff8c00;'
                        f'border:1px solid #ff8c00;font-size:12px;font-weight:700;">'
                        f'🚫 Trend Veto ({bearish_signals}/5)</div>')
            return ''

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

        def divergence_badge(div):
            if div == 'Bearish Divergence':
                return '<span class="div-badge div-bear">⚠ Bear Div</span>'
            elif div == 'Bullish Divergence':
                return '<span class="div-badge div-bull">✅ Bull Div</span>'
            else:
                return '<span class="div-badge div-none">-</span>'

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
            return '#00e676' if v >= 2 else ('#00d4ff' if v >= 1.5 else ('#ffab00' if v >= 1 else '#ff3d57'))

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

        # == V58 MODE TOGGLE ===================================================
        html += """
  <div style="display:flex;align-items:center;gap:16px;margin:0 0 20px 0;flex-wrap:wrap;">
    <span style="font-size:13px;color:#88aacc;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;">Analysis Mode</span>
    <div class="mode-toggle">
      <button class="mode-btn tf-active active" id="btn-tf" onclick="switchMode('technofunc')">📊 TechnoFunc</button>
      <button class="mode-btn t-active" id="btn-t" onclick="switchMode('technical')">📈 Technical</button>
    </div>
    <span id="mode-desc" style="font-size:13px;color:#00f5ff;font-weight:600;">
      Fundamentals 65% + Technicals 35% · Trend Veto · R:R Gate
    </span>
  </div>
"""
        html += '<div id="sec-technofunc" class="buy-section">\n'

        # == BUY TABLE =========================================================
        if not top_buys.empty:
            html += """
  <div class="section-hdr">
    <div class="section-pill pill-buy">▲ Top Buy Recommendations - Sector Diversified</div>
    <div class="section-line"></div>
    <div class="section-note">STOCK INFO · TRADE SETUP · TECHNICALS · FUNDAMENTALS · META</div>
  </div>
  <div class="tbl-wrap"><table>
    <thead>
      <tr class="grp-row">
        <th class="grp-stock" colspan="3">STOCK INFO</th>
        <th class="grp-trade gsep" colspan="6">TRADE SETUP</th>
        <th class="grp-tech gsep"  colspan="6">TECHNICALS</th>
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
        <th class="ch-tech gsep">RSI / Div</th>
        <th class="ch-tech">ADX</th>
        <th class="ch-tech">Vol / Avg</th>
        <th class="ch-tech">Sup Dist</th>
        <th class="ch-tech">52W Hi %</th>
        <th class="ch-tech">MACD</th>
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
                div      = f"{row['Dividend_Yield']:.2f}%" if row['Dividend_Yield'] > 0 else '-'
                divc     = '#00e676' if row['Dividend_Yield'] > 0 else '#4a6080'
                rr       = row['Risk_Reward']
                mcdcls   = 'macd-bull' if row['MACD'] == 'Bullish' else 'macd-bear'
                bs       = row.get('Bearish_Signals', 0)
                wm       = row.get('Weight_Mode', '')
                rsi_dir_b = row.get('RSI_Direction', 'Flat')
                rsi_slp_b = row.get('RSI_Slope', 0)
                if rsi_dir_b == 'Rising':
                    rsi_slope_html = f'<span style="color:#00e676;font-size:12px">↑ +{rsi_slp_b:.0f}</span>'
                elif rsi_dir_b == 'Falling':
                    slp_clr = '#ff4466' if abs(rsi_slp_b) > 8 else '#ffab00'
                    rsi_slope_html = f'<span style="color:{slp_clr};font-size:12px">↓ {rsi_slp_b:.0f}</span>'
                else:
                    rsi_slope_html = '<span style="color:#4a6080;font-size:12px">→</span>'

                html += f"""      <tr>
        <td><span class="rnum">{i}</span></td>
        <td>
          <div class="stock-name">{row['Name']}</div>
          <div class="stock-sym">{row['Symbol']}</div>
          <div class="stock-sec">{row.get('Sector','N/A')}</div>
        </td>
        <td><div class="price-val">₹{row['Price']:,.2f}</div></td>
        <td class="gsep">
          {rating_badge(rec, {'STRONG BUY':'⭐⭐⭐⭐⭐ STRONG BUY','BUY':'⭐⭐⭐⭐ BUY','HOLD':'⭐⭐⭐ HOLD','SELL':'⭐⭐ SELL','STRONG SELL':'⭐ STRONG SELL'}.get(rec, rec))}
          {score_cell(row['Combined_Score'], sc_color, sc_bar)}
          {veto_badge(bs, wm)}
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
          {divergence_badge(row.get('RSI_Divergence','None'))}
          {rsi_slope_html}
        </td>
        <td>{adx_cell(row.get('ADX', 0))}</td>
        <td>{vol_cell(row.get('Vol_Ratio', 1.0))}</td>
        <td>{sdist_cell(row.get('Support_Dist_Pct', 0))}</td>
        <td><span class="mono-sm" style="color:{w52_color(w52)}">{w52:+.1f}%</span></td>
        <td><span class="{mcdcls}">{row['MACD']}</span></td>
        <td class="gsep"><span class="mono-sm" style="color:{pe_color(row['PE_Ratio'],'buy')}">{f"{row['PE_Ratio']:.1f}" if row['PE_Ratio']>0 else 'N/A'}</span></td>
        <td><span class="mono-sm" style="color:{beta_color(row['Beta'])}">{row['Beta']:.2f}</span></td>
        <td><span class="mono-sm" style="color:{divc}">{div}</span></td>
        <td>{quality_badge(row['Quality'])}</td>
        <td class="gsep">{analyst_badge(row.get('Analyst','N/A'))}</td>
        <td><div class="earn-date">{row.get('Earnings_Date','N/A')}</div></td>
        <td>{action_button(rec)}</td>
      </tr>
"""
            html += "    </tbody></table></div>\n"

        html += '</div><!-- /sec-technofunc -->\n'

        # == V58: TECHNICAL-ONLY BUY TABLE =====================================
        # 100% technical scoring. ALL 100 stocks read & ranked.
        # Buy-side targets RECALCULATED from Support/Resistance.
        # Columns: ONLY technical — NO PE, Beta, Quality, Analyst, Earnings.
        # =================================================================
        html += '<div id="sec-technical" class="buy-section hidden">\n'
        if not tech_buys.empty:
            html += """
  <div class="section-hdr">
    <div class="section-pill" style="background:#28124a;color:#cc99ff;border:2px solid #7c4dff;">
      📈 Top 10 Technical Buy Picks — 100%% Chart Analysis
    </div>
    <div class="section-line"></div>
    <div class="section-note">ALL 100 STOCKS EVALUATED · RSI + MACD + SMA + ADX + VOLUME · ZERO FUNDAMENTALS</div>
  </div>
  <div class="tbl-wrap"><table>
    <thead>
      <tr class="grp-row">
        <th class="grp-stock" colspan="3">STOCK INFO</th>
        <th style="background:#28124a;color:#cc99ff;text-shadow:0 0 8px rgba(204,153,255,0.6);font-size:12px;font-weight:800;letter-spacing:3px;text-transform:uppercase;padding:9px 10px;border-bottom:1px solid rgba(255,255,255,0.1);white-space:nowrap;" class="gsep" colspan="2">TECH SCORE</th>
        <th class="grp-tech gsep" colspan="7">TECHNICAL SIGNALS</th>
        <th class="grp-trade gsep" colspan="5">TRADE SETUP (BUY SIDE)</th>
      </tr>
      <tr class="col-row">
        <th class="ch-stock" style="width:26px">#</th>
        <th class="ch-stock">Stock / Sector</th>
        <th class="ch-stock">Price</th>
        <th style="border-top:3px solid #cc99ff;color:#e8d0ff;font-size:13px;font-weight:800;letter-spacing:1px;text-transform:uppercase;padding:9px 10px;background:#0c1a2e;border-bottom:3px solid #1e3a5a;white-space:nowrap;text-align:left;" class="gsep">Tech Score</th>
        <th style="border-top:3px solid #cc99ff;color:#e8d0ff;font-size:13px;font-weight:800;letter-spacing:1px;text-transform:uppercase;padding:9px 10px;background:#0c1a2e;border-bottom:3px solid #1e3a5a;white-space:nowrap;text-align:left;">Signal</th>
        <th class="ch-tech gsep">RSI / Slope</th>
        <th class="ch-tech">MACD</th>
        <th class="ch-tech">SMA 20</th>
        <th class="ch-tech">SMA 50</th>
        <th class="ch-tech">SMA 200</th>
        <th class="ch-tech">ADX</th>
        <th class="ch-tech">Vol / Avg</th>
        <th class="ch-trade gsep">Target 1 / T2</th>
        <th class="ch-trade">Stop Loss</th>
        <th class="ch-trade">R : R</th>
        <th class="ch-trade">Upside</th>
        <th class="ch-trade">ATR</th>
      </tr>
    </thead>
    <tbody>
"""
            for i, (_, row) in enumerate(tech_buys.iterrows(), 1):
                tbs      = row['Tech_Buy_Score']
                rsic     = '#ff3d57' if row['RSI'] > 70 else ('#00e676' if row['RSI'] < 30 else '#60a5fa')
                mcdcls   = 'macd-bull' if row['MACD'] == 'Bullish' else 'macd-bear'
                rsi_dir  = row.get('RSI_Direction', 'Flat')
                rsi_slp  = row.get('RSI_Slope', 0)

                # Use RECALCULATED buy-side fields
                bt1     = row.get('Buy_Target_1', row['Target_1'])
                bt2     = row.get('Buy_Target_2', row['Target_2'])
                bsl     = row.get('Buy_Stop_Loss', row['Stop_Loss'])
                bsl_pct = row.get('Buy_SL_Pct', row['SL_Percentage'])
                bup     = row.get('Buy_Upside', row['Upside'])
                brr     = row.get('Buy_RR', row['Risk_Reward'])

                upcls   = 'up' if bup >= 0 else 'dn'

                # Score color
                if tbs >= 70:   tbs_color, tbs_bar = '#00ff88', '#00c853'
                elif tbs >= 50: tbs_color, tbs_bar = '#cc99ff', '#7c4dff'
                elif tbs >= 35: tbs_color, tbs_bar = '#00f5ff', '#0099cc'
                else:           tbs_color, tbs_bar = '#ffab00', '#f59e0b'

                # RSI slope
                if rsi_dir == 'Rising':
                    rsi_slope_html = f'<span style="color:#00e676;font-size:12px;font-weight:700">↑ +{rsi_slp:.0f}</span>'
                elif rsi_dir == 'Falling':
                    slp_clr = '#ff4466' if abs(rsi_slp) > 8 else '#ffab00'
                    rsi_slope_html = f'<span style="color:{slp_clr};font-size:12px;font-weight:700">↓ {rsi_slp:.0f}</span>'
                else:
                    rsi_slope_html = '<span style="color:#4a6080;font-size:12px">→ flat</span>'

                # Signal label
                if row['RSI'] < 30 and rsi_dir == 'Rising':
                    signal_text = '<span style="color:#00ff88;font-weight:700">🔥 Oversold Reversal</span>'
                elif row['RSI'] < 30:
                    signal_text = '<span style="color:#00e676;font-weight:700">⬇ Deep Oversold</span>'
                elif row['RSI'] < 40 and rsi_dir == 'Rising':
                    signal_text = '<span style="color:#00e676;font-weight:700">↗ Recovery Setup</span>'
                elif row['RSI'] < 50 and row['MACD'] == 'Bullish':
                    signal_text = '<span style="color:#00d4ff;font-weight:700">📊 MACD Bullish Cross</span>'
                elif row.get('RSI_Divergence') == 'Bullish Divergence':
                    signal_text = '<span style="color:#cc99ff;font-weight:700">🔀 Bullish Divergence</span>'
                elif rsi_dir == 'Rising' and row['MACD'] == 'Bullish':
                    signal_text = '<span style="color:#00d4ff;font-weight:700">📈 Momentum + MACD</span>'
                elif rsi_dir == 'Rising':
                    signal_text = '<span style="color:#00d4ff;font-weight:700">📈 Momentum Build</span>'
                else:
                    signal_text = '<span style="color:#aaccee;font-weight:700">📋 Technical Setup</span>'

                # SMA % from price
                price  = row['Price']
                sma20  = row.get('SMA_20', 0)
                sma50  = row.get('SMA_50', 0)
                sma200 = row.get('SMA_200', 0)

                def sma_cell_fn(sma_val):
                    if sma_val <= 0: return '<span style="color:#4a6080">N/A</span>'
                    pct = ((price - sma_val) / sma_val) * 100
                    if pct >= 0:
                        return f'<span style="color:#00e676;font-weight:700;font-size:14px">▲ +{pct:.1f}%</span>'
                    else:
                        return f'<span style="color:#ff4466;font-weight:700;font-size:14px">▼ {pct:.1f}%</span>'

                # TechnoFunc ref (tiny, for context)
                tf_rec   = row.get('Recommendation', '?')
                tf_score = row.get('Combined_Score', 0)

                html += f"""      <tr>
        <td><span class="rnum">{i}</span></td>
        <td>
          <div class="stock-name">{row['Name']}</div>
          <div class="stock-sym">{row['Symbol']}</div>
          <div class="stock-sec">{row.get('Sector','N/A')}</div>
          <div style="font-size:10px;color:#556677;margin-top:2px;">TF: {tf_rec} ({tf_score:.0f})</div>
        </td>
        <td><div class="price-val">₹{price:,.2f}</div></td>
        <td class="gsep">
          {score_cell(tbs, tbs_color, tbs_bar)}
        </td>
        <td>{signal_text}</td>
        <td class="gsep">
          <div class="rsi-val" style="color:{rsic}">{row['RSI']:.0f}</div>
          <div class="rsi-sig">{row['RSI_Signal']}</div>
          {rsi_slope_html}
        </td>
        <td><span class="{mcdcls}">{row['MACD']}</span></td>
        <td>{sma_cell_fn(sma20)}</td>
        <td>{sma_cell_fn(sma50)}</td>
        <td>{sma_cell_fn(sma200)}</td>
        <td>{adx_cell(row.get('ADX', 0))}</td>
        <td>{vol_cell(row.get('Vol_Ratio', 1.0))}</td>
        <td class="gsep">
          <div class="t1-val">₹{bt1:,.2f}</div>
          <div class="t2-val">T2: ₹{bt2:,.2f}</div>
        </td>
        <td>
          <div class="sl-val">₹{bsl:,.2f}</div>
          <div class="sl-pct">-{bsl_pct:.1f}%</div>
        </td>
        <td><span class="rr-val" style="color:{rr_color(brr)}">{brr:.1f}×</span></td>
        <td><span class="upside-val {upcls}">{bup:+.1f}%</span></td>
        <td>
          <div class="atr-val">₹{row['ATR']:,.2f}</div>
          <div class="atr-sub">{row['ATR_Pct']:.1f}%</div>
        </td>
      </tr>
"""
            html += "    </tbody></table></div>\n"
        else:
            html += '<p style="color:#cc99ff;padding:16px;font-size:15px;">No stocks currently meet technical buy criteria.</p>\n'

        html += '</div><!-- /sec-technical -->\n'

        # == SELL TABLE ========================================================
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
                div      = f"{row['Dividend_Yield']:.2f}%" if row['Dividend_Yield'] > 0 else '-'
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
          {rating_badge(rec, {'STRONG BUY':'⭐⭐⭐⭐⭐ STRONG BUY','BUY':'⭐⭐⭐⭐ BUY','HOLD':'⭐⭐⭐ HOLD','SELL':'⭐⭐ SELL','STRONG SELL':'⭐ STRONG SELL'}.get(rec, rec))}
          {score_cell(row['Combined_Score'], '#ff3d57', '#c62828')}
        </td>
        <td><span class="upside-val {dncls}">{row['Upside']:+.1f}%</span></td>
        <td>
          <span class="target-badge {tbcls}">{tbtxt}</span>
          <div class="t1-val">₹{row['Target_1']:,.2f}</div>
          <div class="t2-val">T2: ₹{row['Target_2']:,.2f}</div>
        </td>
        <td>
          <div style="font-family:'IBM Plex Mono',monospace;font-size:15px;font-weight:600;color:#ffab00">₹{row['Stop_Loss']:,.2f}</div>
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
        <td>{action_button(rec)}</td>
      </tr>
"""
            html += "    </tbody></table></div>\n"

        # == ALL STOCKS WATCHLIST TABLE =========================================
        # Shows every analyzed stock sorted by Combined_Score desc.
        # User can see all signals and decide their own action.
        # Includes a JS filter bar (All / Strong Buy / Buy / Hold / Sell).
        all_sorted = df.sort_values('Combined_Score', ascending=False)

        html += """
  <div class="section-hdr" style="margin-top:32px">
    <div class="section-pill" style="background:linear-gradient(90deg,#1a2a4a,#0c1a2e);color:#aaccee;border:1px solid #2a4a6a;">
      📋 Complete Watchlist — All {count} Stocks · Sorted by Score
    </div>
    <div class="section-line"></div>
    <div class="section-note">Click a filter to show only that recommendation type · You decide the action</div>
  </div>

  <!-- Filter bar -->
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin:0 0 12px 0;padding:0 4px">
    <button onclick="filterWL('ALL')"        class="wl-btn wl-all"  id="wlf-ALL">All ({all_c})</button>
    <button onclick="filterWL('STRONG BUY')" class="wl-btn wl-sb"   id="wlf-STRONG_BUY">⭐⭐ Strong Buy ({sb_c})</button>
    <button onclick="filterWL('BUY')"        class="wl-btn wl-b"    id="wlf-BUY">▲ Buy ({b_c})</button>
    <button onclick="filterWL('HOLD')"       class="wl-btn wl-h"    id="wlf-HOLD">◆ Hold ({h_c})</button>
    <button onclick="filterWL('SELL')"       class="wl-btn wl-s"    id="wlf-SELL">▼ Sell ({s_c})</button>
    <button onclick="filterWL('STRONG SELL')" class="wl-btn wl-ss"  id="wlf-STRONG_SELL">⚠ Strong Sell ({ss_c})</button>
  </div>

  <div class="tbl-wrap"><table id="watchlist-tbl">
    <thead>
      <tr class="col-row" id="wl-sort-row">
        <th style="width:22px">#</th>
        <th onclick="sortWL('name')"     class="sortable">Stock <span class="sort-icon" id="si-name"></span></th>
        <th onclick="sortWL('sector')"   class="sortable">Sector <span class="sort-icon" id="si-sector"></span></th>
        <th onclick="sortWL('price')"    class="sortable">Price <span class="sort-icon" id="si-price"></span></th>
        <th onclick="sortWL('score')"    class="sortable">Score <span class="sort-icon" id="si-score">▼</span></th>
        <th>F/T Split</th>
        <th onclick="sortWL('action')"   class="sortable">Action <span class="sort-icon" id="si-action"></span></th>
        <th onclick="sortWL('rsi')"      class="sortable">RSI <span class="sort-icon" id="si-rsi"></span></th>
        <th>RSI Div</th>
        <th onclick="sortWL('macd')"     class="sortable">MACD <span class="sort-icon" id="si-macd"></span></th>
        <th onclick="sortWL('sma')"      class="sortable">SMA Trend <span class="sort-icon" id="si-sma"></span></th>
        <th onclick="sortWL('vol')"      class="sortable">Vol <span class="sort-icon" id="si-vol"></span></th>
        <th onclick="sortWL('target')"   class="sortable">Target 1 <span class="sort-icon" id="si-target"></span></th>
        <th onclick="sortWL('stoploss')" class="sortable">Stop Loss <span class="sort-icon" id="si-stoploss"></span></th>
        <th onclick="sortWL('rr')"       class="sortable">R:R <span class="sort-icon" id="si-rr"></span></th>
        <th onclick="sortWL('upside')"   class="sortable">Upside <span class="sort-icon" id="si-upside"></span></th>
        <th onclick="sortWL('pe')"       class="sortable">P/E <span class="sort-icon" id="si-pe"></span></th>
        <th onclick="sortWL('quality')"  class="sortable">Quality <span class="sort-icon" id="si-quality"></span></th>
        <th>Analyst</th>
        <th>Signals</th>
      </tr>
    </thead>
    <tbody>
""".format(
            count  = len(all_sorted),
            all_c  = len(all_sorted),
            sb_c   = len(all_sorted[all_sorted['Recommendation'] == 'STRONG BUY']),
            b_c    = len(all_sorted[all_sorted['Recommendation'] == 'BUY']),
            h_c    = len(all_sorted[all_sorted['Recommendation'] == 'HOLD']),
            s_c    = len(all_sorted[all_sorted['Recommendation'] == 'SELL']),
            ss_c   = len(all_sorted[all_sorted['Recommendation'] == 'STRONG SELL']),
        )

        for i, (_, row) in enumerate(all_sorted.iterrows(), 1):
            rec   = row['Recommendation']
            bs    = row.get('Bearish_Signals', 0)
            wm    = row.get('Weight_Mode', '')
            rsic  = '#ff3d57' if row['RSI'] > 70 else ('#00e676' if row['RSI'] < 30 else '#60a5fa')
            mcdcls = 'macd-bull' if row['MACD'] == 'Bullish' else 'macd-bear'
            rr    = row['Risk_Reward']
            upcls = 'up' if row['Upside'] >= 0 else 'dn'

            # SMA trend summary: compact 3-light indicator
            sma_declining = row.get('SMA_20_Declining', False)
            death_cross   = row.get('Death_Cross', False)
            sma200_rising = row.get('SMA_200_Rising', True)
            if sma_declining and death_cross:
                sma_trend = '<span style="color:#ff4466;font-size:13px">↓ Declining</span>'
            elif sma_declining or death_cross:
                sma_trend = '<span style="color:#ffab00;font-size:13px">⚠ Weakening</span>'
            elif sma200_rising:
                sma_trend = '<span style="color:#00e676;font-size:13px">↑ Rising</span>'
            else:
                sma_trend = '<span style="color:#60a5fa;font-size:13px">→ Flat</span>'

            # Bearish signal count pill
            if bs >= 4:
                sig_pill = f'<span style="color:#ff4466;font-size:13px">🔴 {bs}/5 Bear</span>'
            elif bs >= 3:
                sig_pill = f'<span style="color:#ff8c00;font-size:13px">🟠 {bs}/5 Bear</span>'
            elif bs >= 1:
                sig_pill = f'<span style="color:#ffab00;font-size:13px">🟡 {bs}/5</span>'
            else:
                sig_pill = f'<span style="color:#00e676;font-size:13px">🟢 0/5</span>'

            # Score colour
            if row['Combined_Score'] >= 70:   sc = '#00ff88'
            elif row['Combined_Score'] >= 50:  sc = '#00f5ff'
            elif row['Combined_Score'] >= 40:  sc = '#ffab00'
            else:                              sc = '#ff4466'

            # data-rec attribute drives JS filter
            data_rec = rec.replace(' ', '_')

            rsi_div   = row.get('RSI_Divergence', 'None')
            fund_sc   = row.get('Fund_Score', 0)
            tech_sc   = row.get('Tech_Score', 0)
            veto      = row.get('Veto_Fired', False)
            rsi_dir   = row.get('RSI_Direction', 'Flat')
            rsi_slp   = row.get('RSI_Slope', 0)
            rsi_5b    = row.get('RSI_5Bar', row['RSI'])

            # RSI direction arrow + slope for display
            if rsi_dir == 'Rising':
                rsi_dir_html = f'<span style="color:#00e676;font-size:12px">↑ +{rsi_slp:.0f}</span>'
            elif rsi_dir == 'Falling':
                clr = '#ff4466' if abs(rsi_slp) > 8 else '#ffab00'
                rsi_dir_html = f'<span style="color:{clr};font-size:12px">↓ {rsi_slp:.0f}</span>'
            else:
                rsi_dir_html = '<span style="color:#4a6080;font-size:12px">→ flat</span>'

            # F/T split pill: shows fund score / tech score
            tech_col  = '#00e676' if tech_sc >= 3 else ('#ffab00' if tech_sc >= 0 else '#ff4466')
            fund_col  = '#00e676' if fund_sc >= 70 else ('#ffab00' if fund_sc >= 50 else '#ff4466')
            ft_cell   = (f'<span style="font-size:13px;color:{fund_col}">F:{fund_sc:.0f}</span>'
                         f'<span style="color:#4a6080;font-size:12px"> / </span>'
                         f'<span style="font-size:13px;color:{tech_col}">T:{tech_sc:+d}</span>')

            # RSI divergence pill
            if rsi_div == 'Bearish Divergence':
                rsi_div_cell = '<span style="color:#ff4466;font-size:12px">⚠ Bear Div</span>'
            elif rsi_div == 'Bullish Divergence':
                rsi_div_cell = '<span style="color:#00e676;font-size:12px">✅ Bull Div</span>'
            else:
                rsi_div_cell = '<span style="color:#4a6080;font-size:12px">—</span>'

            # Veto label inside action cell
            veto_cell = ('<span style="display:block;font-size:11px;color:#ff8c00;margin-top:2px">🚫 Veto</span>'
                         if veto else '')

            # Sort key for SMA trend (numeric: 3=Rising, 2=Flat, 1=Weakening, 0=Declining)
            sma_sort_key = 3 if sma200_rising and not sma_declining and not death_cross else                            1 if sma_declining or death_cross else                            2 if sma200_rising else 1
            # Quality sort key
            quality_sort = {'Excellent': 4, 'Good': 3, 'Average': 2, 'Poor': 1}.get(row.get('Quality','Average'), 2)
            # MACD sort key
            macd_sort = 1 if row['MACD'] == 'Bullish' else 0
            # Action sort key
            action_sort = {'STRONG BUY': 5, 'BUY': 4, 'WATCH': 3, 'HOLD': 2, 'SELL': 1, 'STRONG SELL': 0}.get(rec, 2)

            html += f"""      <tr data-rec="{data_rec}"
          data-name="{row['Name'].lower()}"
          data-sector="{row.get('Sector','N/A').lower()}"
          data-price="{row['Price']:.2f}"
          data-score="{row['Combined_Score']:.1f}"
          data-rsi="{row['RSI']:.2f}"
          data-vol="{row.get('Vol_Ratio', 1.0):.3f}"
          data-target="{row['Target_1']:.2f}"
          data-stoploss="{row['Stop_Loss']:.2f}"
          data-rr="{row['Risk_Reward']:.2f}"
          data-upside="{row['Upside']:.2f}"
          data-pe="{row['PE_Ratio']:.2f}"
          data-quality="{quality_sort}"
          data-sma="{sma_sort_key}"
          data-macd="{macd_sort}"
          data-action="{action_sort}">
        <td><span class="rnum">{i}</span></td>
        <td>
          <div class="stock-name" style="font-size:14px">{row['Name']}</div>
          <div class="stock-sym">{row['Symbol']}</div>
        </td>
        <td><span style="font-size:13px;color:#8899aa">{row.get('Sector','N/A')}</span></td>
        <td><div class="price-val" style="font-size:15px">₹{row['Price']:,.2f}</div></td>
        <td>
          <span style="font-family:'IBM Plex Mono',monospace;font-size:16px;font-weight:700;color:{sc}">{row['Combined_Score']:.0f}</span>
          {veto_badge(bs, wm)}
        </td>
        <td>{ft_cell}</td>
        <td>
          {action_button(rec)}
          {veto_cell}
        </td>
        <td>
          <div class="rsi-val" style="color:{rsic};font-size:15px">{row['RSI']:.0f}</div>
          <div style="font-size:12px;color:#8899aa">{row['RSI_Signal']}</div>
          {rsi_dir_html}
        </td>
        <td>{rsi_div_cell}</td>
        <td><span class="{mcdcls}" style="font-size:13px">{row['MACD']}</span></td>
        <td>{sma_trend}</td>
        <td>{vol_cell(row.get('Vol_Ratio', 1.0))}</td>
        <td><div style="font-size:14px;color:#00d4ff">₹{row['Target_1']:,.2f}</div></td>
        <td><div style="font-size:14px;color:#ffab00">₹{row['Stop_Loss']:,.2f}</div></td>
        <td><span class="rr-val" style="color:{rr_color(rr)};font-size:14px">{rr:.1f}×</span></td>
        <td><span class="upside-val {upcls}" style="font-size:14px">{row['Upside']:+.1f}%</span></td>
        <td><span style="font-size:14px;color:{pe_color(row['PE_Ratio'],'buy')}">{f"{row['PE_Ratio']:.1f}" if row['PE_Ratio']>0 else 'N/A'}</span></td>
        <td>{quality_badge(row['Quality'])}</td>
        <td>{analyst_badge(row.get('Analyst','N/A'))}</td>
        <td>{sig_pill}</td>
      </tr>
"""
        html += "    </tbody></table></div>\n"

        # =====================================================================
        #  ACCUMULATE ON DIP SECTION — V57
        # =====================================================================
        html += """
  <div style="margin:32px 0 0 0;">
    <div style="display:flex;align-items:center;gap:14px;margin-bottom:14px;">
      <span style="font-size:22px;font-weight:800;color:#ffcc00;font-family:'Syne',sans-serif;letter-spacing:1px;">
        ⏳ Accumulate on Dip
      </span>
      <span style="font-size:13px;color:#aaccee;background:#1a1a00;border:1px solid #ffcc00;
                   border-radius:6px;padding:3px 10px;">
        Score 42–49 · Good Fundamentals · Wait for Signal
      </span>
    </div>
    <p style="font-size:13px;color:#8899aa;margin-bottom:12px;">
      These stocks just missed the BUY threshold but have solid fundamentals.
      Watch for RSI to stabilise above 50 or MACD to flip bullish — then they become actionable.
    </p>
"""
        if accumulate_df.empty:
            html += '<p style="color:#aaccee;padding:12px;">No stocks currently in the accumulate zone.</p>\n'
        else:
            html += """
    <div style="overflow-x:auto;">
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead>
        <tr style="background:#1a1500;color:#ffcc00;text-align:left;">
          <th style="padding:10px 8px;">#</th>
          <th style="padding:10px 8px;">Stock / Sector</th>
          <th style="padding:10px 8px;">Price</th>
          <th style="padding:10px 8px;">Score</th>
          <th style="padding:10px 8px;">F / T</th>
          <th style="padding:10px 8px;">RSI</th>
          <th style="padding:10px 8px;">MACD</th>
          <th style="padding:10px 8px;">Quality</th>
          <th style="padding:10px 8px;">What to Watch For</th>
        </tr>
      </thead>
      <tbody>
"""
            for i, (_, row) in enumerate(accumulate_df.iterrows(), 1):
                rsi_val  = row.get('RSI', 50)
                rsi_dir  = row.get('RSI_Direction', 'Flat')
                rsi_slp  = row.get('RSI_Slope', 0)
                macd_sig = row.get('MACD', 'Bearish')
                fs       = row.get('Fund_Score', 0)
                ts       = row.get('Tech_Score', 0)
                cs       = row.get('Combined_Score', 0)
                qual     = row.get('Quality', 'Average')
                price    = row.get('Price', 0)

                # Colour coding
                rsi_col  = '#ff4466' if rsi_val > 65 else ('#00e676' if rsi_val < 40 else '#60a5fa')
                macd_col = '#00e676' if macd_sig == 'Bullish' else '#ff6680'
                qual_col = {'Excellent': '#00ff88', 'Good': '#00d4ff',
                            'Average': '#ffab00', 'Poor': '#ff4466'}.get(qual, '#aaccee')
                ts_col   = '#00e676' if ts >= 0 else '#ff4466'
                fs_col   = '#00e676' if fs >= 70 else ('#ffab00' if fs >= 50 else '#ff4466')

                # Direction arrow
                if rsi_dir == 'Rising':
                    dir_html = f'<span style="color:#00e676;font-size:11px">↑+{rsi_slp:.0f}</span>'
                elif rsi_dir == 'Falling':
                    dir_html = f'<span style="color:#ffab00;font-size:11px">↓{rsi_slp:.0f}</span>'
                else:
                    dir_html = '<span style="color:#4a6080;font-size:11px">→flat</span>'

                # What to watch for — context-aware trigger hint
                triggers = []
                if macd_sig == 'Bearish':
                    triggers.append('MACD flip bullish')
                if rsi_val < 50:
                    triggers.append('RSI cross above 50')
                if rsi_dir == 'Falling':
                    triggers.append('RSI stabilise')
                if not triggers:
                    triggers.append('Volume confirmation')
                watch_str = ' · '.join(triggers)

                bg = '#0c1005' if i % 2 == 0 else '#090d02'
                html += f"""        <tr style="background:{bg};border-bottom:1px solid #1e2a10;">
          <td style="padding:9px 8px;color:#6a8a4a;">{i}</td>
          <td style="padding:9px 8px;">
            <div style="font-weight:600;color:#ddeeff;font-size:13px;">{row['Name']}</div>
            <div style="font-size:11px;color:#6a8a6a;">{row['Symbol']} · {row.get('Sector','N/A')}</div>
          </td>
          <td style="padding:9px 8px;color:#aaccee;font-family:'IBM Plex Mono',monospace;">₹{price:,.2f}</td>
          <td style="padding:9px 8px;">
            <span style="font-family:'IBM Plex Mono',monospace;font-size:15px;font-weight:700;color:#ffcc00;">{cs:.0f}</span>
          </td>
          <td style="padding:9px 8px;">
            <span style="font-size:12px;color:{fs_col}">F:{fs:.0f}</span>
            <span style="color:#4a6080;font-size:11px"> / </span>
            <span style="font-size:12px;color:{ts_col}">T:{ts:+d}</span>
          </td>
          <td style="padding:9px 8px;">
            <span style="color:{rsi_col};font-size:13px;font-weight:600;">{rsi_val:.0f}</span>
            <span style="margin-left:4px;">{dir_html}</span>
          </td>
          <td style="padding:9px 8px;color:{macd_col};font-size:12px;">{macd_sig}</td>
          <td style="padding:9px 8px;color:{qual_col};font-size:12px;">{qual}</td>
          <td style="padding:9px 8px;color:#ccaa44;font-size:12px;">👁 {watch_str}</td>
        </tr>
"""
            html += "      </tbody></table></div>\n"
        html += "  </div>\n"

        html += f"""
  <div class="disc">
    <strong style="color:var(--red)">⚠ DISCLAIMER:</strong>
    For <strong>EDUCATIONAL PURPOSES ONLY</strong>. Not financial advice.
    Stop losses are ATR-based near real 12-month S/R zones. Targets derived from
    swing highs/lows, 52-week extremes and round-number levels. Earnings dates are estimates.
    RSI divergence signals are algorithmic and not a guarantee of reversal.
    Always conduct your own research, consult a SEBI-registered financial advisor,
    and never invest more than you can afford to lose.
  </div>
</div>

<footer>
  <strong>NIFTY 100 Market Influencers · NSE &amp; BSE</strong>
  · 12M S/R · Trend Veto · Dynamic Weights · SMA200 Slope · Sector PE · v5.7
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

// ── Watchlist filter ────────────────────────────────────────────────────────
var wlCurrentFilter = 'ALL';
function filterWL(rec) {{
  wlCurrentFilter = rec;
  var rows = document.querySelectorAll('#watchlist-tbl tbody tr');
  var btns = document.querySelectorAll('.wl-btn');
  btns.forEach(function(b) {{ b.style.opacity = '0.45'; b.style.fontWeight = '500'; }});
  var activeBtn = document.getElementById('wlf-' + rec.replace(/ /g,'_'));
  if (activeBtn) {{ activeBtn.style.opacity = '1'; activeBtn.style.fontWeight = '800'; }}
  rows.forEach(function(r) {{
    r.style.display = (rec === 'ALL' || r.getAttribute('data-rec') === rec.replace(/ /g,'_')) ? '' : 'none';
  }});
  renumberVisible();
}}

// ── Watchlist column sort ────────────────────────────────────────────────────
var wlSortCol = 'score';
var wlSortAsc = false;
var wlSortIcons = {{}};

function sortWL(col) {{
  if (wlSortCol === col) {{ wlSortAsc = !wlSortAsc; }}
  else {{ wlSortCol = col; wlSortAsc = (col === 'name' || col === 'sector'); }}

  // Update header icons
  var allIcons = document.querySelectorAll('.sort-icon');
  allIcons.forEach(function(ic) {{ ic.textContent = ''; }});
  var activeIcon = document.getElementById('si-' + col);
  if (activeIcon) {{ activeIcon.textContent = wlSortAsc ? ' ▲' : ' ▼'; }}

  var tbody = document.querySelector('#watchlist-tbl tbody');
  var rows  = Array.from(tbody.querySelectorAll('tr'));

  rows.sort(function(a, b) {{
    var av = a.getAttribute('data-' + col) || '';
    var bv = b.getAttribute('data-' + col) || '';

    // String columns
    if (col === 'name' || col === 'sector') {{
      return wlSortAsc ? av.localeCompare(bv) : bv.localeCompare(av);
    }}
    // Numeric columns
    var an = parseFloat(av), bn = parseFloat(bv);
    if (isNaN(an)) an = -Infinity;
    if (isNaN(bn)) bn = -Infinity;
    return wlSortAsc ? an - bn : bn - an;
  }});

  rows.forEach(function(r) {{ tbody.appendChild(r); }});
  filterWL(wlCurrentFilter);
}}

function renumberVisible() {{
  var rows = document.querySelectorAll('#watchlist-tbl tbody tr');
  var n = 1;
  rows.forEach(function(r) {{
    if (r.style.display !== 'none') {{
      var rnum = r.querySelector('.rnum');
      if (rnum) rnum.textContent = n++;
    }}
  }});
}}

window.onload = function() {{ filterWL('ALL'); }};

function switchMode(mode) {{
  var sTF = document.getElementById('sec-technofunc');
  var sT  = document.getElementById('sec-technical');
  var bTF = document.getElementById('btn-tf');
  var bT  = document.getElementById('btn-t');
  var d   = document.getElementById('mode-desc');
  if (mode === 'technical') {{
    sTF.classList.add('hidden');    sT.classList.remove('hidden');
    bTF.classList.remove('active'); bT.classList.add('active');
    d.innerHTML = '100% Technical · RSI + MACD + SMA + ADX + Volume · All 100 Stocks · Top 10';
    d.style.color = '#cc99ff';
  }} else {{
    sT.classList.add('hidden');     sTF.classList.remove('hidden');
    bT.classList.remove('active');  bTF.classList.add('active');
    d.innerHTML = 'Fundamentals 65% + Technicals 35% · Trend Veto · R:R Gate';
    d.style.color = '#00f5ff';
  }}
}}
</script>
<style>
.wl-btn {{
  padding: 6px 14px; border-radius: 6px; border: none; cursor: pointer;
  font-size: 14px; font-family: 'Space Grotesk', sans-serif; font-weight: 500;
  transition: all .2s; letter-spacing: .3px;
}}
th.sortable {{
  cursor: pointer; user-select: none; white-space: nowrap;
  transition: color .15s;
}}
th.sortable:hover {{ color: #00f5ff !important; }}
.sort-icon {{ font-size: 10px; opacity: 0.8; color: #00f5ff; }}
.wl-all {{ background:#1a2a3a; color:#aaccee; border:1px solid #2a4a6a; }}
.wl-sb  {{ background:#004d25; color:#00ff88; border:1px solid #00ff88; }}
.wl-b   {{ background:#003a4d; color:#00f5ff; border:1px solid #00f5ff; }}
.wl-h   {{ background:#2a2200; color:#ffab00; border:1px solid #ffab00; }}
.wl-s   {{ background:#4d0010; color:#ff4466; border:1px solid #ff4466; }}
.wl-ss  {{ background:#5a0015; color:#ff7788; border:1px solid #ff7788; }}
.wl-btn:hover {{ opacity: 1 !important; transform: translateY(-1px); }}
</style>
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
            msg['Subject'] = f"💎 NIFTY 100 Report - {tod} {now.strftime('%d %b %Y')}"
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
        print("💎 NIFTY 100 ANALYZER v5.4 - Trend Veto · Dynamic Weights · SMA200 Slope")
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
        output_file=os.environ.get('OUTPUT_FILE', 'index.html')
    )
