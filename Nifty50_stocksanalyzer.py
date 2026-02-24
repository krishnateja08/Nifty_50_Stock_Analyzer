"""
NIFTY 100 COMPLETE STOCK ANALYZER - AURORA GLASS THEME
Technical + Fundamental Analysis with Email Delivery + GitHub Pages

Requirements:
pip install yfinance pandas numpy openpyxl pytz
"""

import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
import pytz
import warnings
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

warnings.filterwarnings('ignore')


class Nifty50CompleteAnalyzer:
    def __init__(self):
        # Nifty 50 stock symbols
        self.nifty50_stocks = {
            'RELIANCE.NS': 'Reliance Industries',
            'TCS.NS': 'TCS',
            'HDFCBANK.NS': 'HDFC Bank',
            'INFY.NS': 'Infosys',
            'ICICIBANK.NS': 'ICICI Bank',
            'HINDUNILVR.NS': 'Hindustan Unilever',
            'BHARTIARTL.NS': 'Bharti Airtel',
            'ITC.NS': 'ITC',
            'SBIN.NS': 'State Bank of India',
            'LT.NS': 'L&T',
            'BAJFINANCE.NS': 'Bajaj Finance',
            'KOTAKBANK.NS': 'Kotak Mahindra Bank',
            'AXISBANK.NS': 'Axis Bank',
            'ASIANPAINT.NS': 'Asian Paints',
            'MARUTI.NS': 'Maruti Suzuki',
            'TITAN.NS': 'Titan Company',
            'SUNPHARMA.NS': 'Sun Pharma',
            'ULTRACEMCO.NS': 'UltraTech Cement',
            'NESTLEIND.NS': 'Nestle India',
            'WIPRO.NS': 'Wipro',
            'HCLTECH.NS': 'HCL Tech',
            'BAJAJFINSV.NS': 'Bajaj Finserv',
            'POWERGRID.NS': 'Power Grid',
            'NTPC.NS': 'NTPC',
            'ONGC.NS': 'ONGC',
            'TECHM.NS': 'Tech Mahindra',
            'M&M.NS': 'M&M',
            'TATAMOTORS.NS': 'Tata Motors',
            'TATASTEEL.NS': 'Tata Steel',
            'INDUSINDBK.NS': 'IndusInd Bank',
            'ADANIPORTS.NS': 'Adani Ports',
            'COALINDIA.NS': 'Coal India',
            'JSWSTEEL.NS': 'JSW Steel',
            'HINDALCO.NS': 'Hindalco',
            'CIPLA.NS': 'Cipla',
            'DRREDDY.NS': 'Dr Reddy',
            'GRASIM.NS': 'Grasim',
            'DIVISLAB.NS': "Divi's Lab",
            'HEROMOTOCO.NS': 'Hero MotoCorp',
            'EICHERMOT.NS': 'Eicher Motors',
            'BRITANNIA.NS': 'Britannia',
            'APOLLOHOSP.NS': 'Apollo Hospital',
            'BAJAJ-AUTO.NS': 'Bajaj Auto',
            'SHRIRAMFIN.NS': 'Shriram Finance',
            'TATACONSUM.NS': 'Tata Consumer',
            'SBILIFE.NS': 'SBI Life',
            'BPCL.NS': 'BPCL',
            'HDFCLIFE.NS': 'HDFC Life',
            'LTIM.NS': 'LTIMindtree',
            'ADANIENT.NS': 'Adani Enterprises',
            'SIEMENS.NS':       'Siemens India',
            # ── NIFTY NEXT 50 (additional 50) ─────────────────────
            'HAVELLS.NS':       'Havells India',
            'PIDILITIND.NS':    'Pidilite Industries',
            'DABUR.NS':         'Dabur India',
            'MARICO.NS':        'Marico',
            'GODREJCP.NS':      'Godrej Consumer Products',
            'COLPAL.NS':        'Colgate-Palmolive India',
            'BERGEPAINT.NS':    'Berger Paints',
            'MUTHOOTFIN.NS':    'Muthoot Finance',
            'CHOLAFIN.NS':      'Cholamandalam Investment',
            'BAJAJHLDNG.NS':    'Bajaj Holdings',
            'SBICARD.NS':       'SBI Cards',
            'ICICIPRULI.NS':    'ICICI Prudential Life',
            'ICICIGI.NS':       'ICICI Lombard General Insurance',
            'HDFCAMC.NS':       'HDFC AMC',
            'NAUKRI.NS':        'Info Edge (Naukri)',
            'MCDOWELL-N.NS':    'United Spirits',
            'TATAELXSI.NS':     'Tata Elxsi',
            'COFORGE.NS':       'Coforge',
            'PERSISTENT.NS':    'Persistent Systems',
            'OFSS.NS':          'Oracle Financial Services',
            'LTTS.NS':          'L&T Technology Services',
            'PAGEIND.NS':       'Page Industries',
            'VOLTAS.NS':        'Voltas',
            'AMBUJACEM.NS':     'Ambuja Cements',
            'ACC.NS':           'ACC',
            'INDIGO.NS':        'IndiGo (InterGlobe Aviation)',
            'DMART.NS':         'Avenue Supermarts (DMart)',
            'VEDL.NS':          'Vedanta',
            'SAIL.NS':          'Steel Authority of India',
            'NMDC.NS':          'NMDC',
            'RECLTD.NS':        'REC Limited',
            'PFC.NS':           'Power Finance Corporation',
            'IRCTC.NS':         'IRCTC',
            'CONCOR.NS':        'Container Corporation of India',
            'JINDALSTEL.NS':    'Jindal Steel & Power',
            'MOTHERSON.NS':     'Samvardhana Motherson',
            'BALKRISIND.NS':    'Balkrishna Industries',
            'TORNTPHARM.NS':    'Torrent Pharmaceuticals',
            'LUPIN.NS':         'Lupin',
            'AUROPHARMA.NS':    'Aurobindo Pharma',
            'ALKEM.NS':         'Alkem Laboratories',
            'MAXHEALTH.NS':     'Max Healthcare',
            'FORTIS.NS':        'Fortis Healthcare',
            'ZOMATO.NS':        'Zomato',
            'POLICYBZR.NS':     'PB Fintech (PolicyBazaar)',
            'NYKAA.NS':         'FSN E-Commerce (Nykaa)',
            'PAYTM.NS':         'One97 Communications (Paytm)',
            'RVNL.NS':          'Rail Vikas Nigam',
            'ADANIGREEN.NS':    'Adani Green Energy'
        }

        self.results = []

    def get_ist_time(self):
        """Get current time in IST timezone"""
        ist = pytz.timezone('Asia/Kolkata')
        return datetime.now(ist)

    def calculate_rsi(self, prices, period=14):
        """Calculate RSI"""
        delta = prices.diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=period).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
        rs = gain / loss
        rsi = 100 - (100 / (1 + rs))
        return rsi.iloc[-1]

    def calculate_macd(self, prices):
        """Calculate MACD"""
        ema12 = prices.ewm(span=12, adjust=False).mean()
        ema26 = prices.ewm(span=26, adjust=False).mean()
        macd = ema12 - ema26
        signal = macd.ewm(span=9, adjust=False).mean()
        return macd.iloc[-1], signal.iloc[-1]

    def get_fundamental_score(self, info):
        """Calculate fundamental score (0-100)"""
        score = 0

        # Valuation Score (25 points)
        pe = info.get('trailingPE', info.get('forwardPE', 0))
        pb = info.get('priceToBook', 0)
        peg = info.get('pegRatio', 0)

        if pe and 0 < pe < 25:
            score += 10
        elif pe and 25 <= pe < 35:
            score += 5

        if pb and 0 < pb < 3:
            score += 5
        elif pb and 3 <= pb < 5:
            score += 3

        if peg and 0 < peg < 1:
            score += 10
        elif peg and 1 <= peg < 2:
            score += 5

        # Profitability Score (25 points)
        roe = info.get('returnOnEquity', 0)
        roa = info.get('returnOnAssets', 0)
        profit_margin = info.get('profitMargins', 0)

        if roe and roe > 0.15:
            score += 10
        elif roe and roe > 0.10:
            score += 5

        if roa and roa > 0.05:
            score += 5
        elif roa and roa > 0.02:
            score += 3

        if profit_margin and profit_margin > 0.10:
            score += 10
        elif profit_margin and profit_margin > 0.05:
            score += 5

        # Growth Score (25 points)
        revenue_growth = info.get('revenueGrowth', 0)
        earnings_growth = info.get('earningsGrowth', 0)

        if revenue_growth and revenue_growth > 0.15:
            score += 10
        elif revenue_growth and revenue_growth > 0.10:
            score += 7
        elif revenue_growth and revenue_growth > 0.05:
            score += 5

        if earnings_growth and earnings_growth > 0.15:
            score += 10
        elif earnings_growth and earnings_growth > 0.10:
            score += 7
        elif earnings_growth and earnings_growth > 0.05:
            score += 5

        # Financial Health Score (25 points)
        debt_to_equity = info.get('debtToEquity', 0)
        current_ratio = info.get('currentRatio', 0)

        if debt_to_equity is not None:
            if debt_to_equity < 50:
                score += 10
            elif debt_to_equity < 100:
                score += 5
        else:
            score += 5

        if current_ratio and current_ratio > 1.5:
            score += 10
        elif current_ratio and current_ratio > 1.0:
            score += 5

        # Free cash flow
        free_cashflow = info.get('freeCashflow', 0)
        if free_cashflow and free_cashflow > 0:
            score += 5

        return min(score, 100)

    def analyze_stock(self, symbol, name):
        """Analyze individual stock - Technical + Fundamental"""
        try:
            stock = yf.Ticker(symbol)
            df = stock.history(period='1y')
            info = stock.info

            if df.empty or len(df) < 200:
                return None

            # ========== TECHNICAL ANALYSIS ==========
            current_price = df['Close'].iloc[-1]

            # Moving Averages
            sma_20 = df['Close'].rolling(window=20).mean().iloc[-1]
            sma_50 = df['Close'].rolling(window=50).mean().iloc[-1]
            sma_200 = df['Close'].rolling(window=200).mean().iloc[-1]

            # Indicators
            rsi = self.calculate_rsi(df['Close'])
            macd, signal = self.calculate_macd(df['Close'])

            # Support/Resistance
            recent_60 = df.tail(60)
            resistance = recent_60['High'].quantile(0.90)
            support = recent_60['Low'].quantile(0.10)

            # 52-week
            high_52w = df['High'].tail(252).max()
            low_52w = df['Low'].tail(252).min()

            # Technical Score (-6 to +6)
            tech_score = 0

            if current_price > sma_20:
                tech_score += 1
            else:
                tech_score -= 1

            if current_price > sma_50:
                tech_score += 1
            else:
                tech_score -= 1

            if current_price > sma_200:
                tech_score += 2
            else:
                tech_score -= 2

            if rsi < 30:
                tech_score += 2
                rsi_signal = "Oversold"
            elif rsi > 70:
                tech_score -= 2
                rsi_signal = "Overbought"
            else:
                rsi_signal = "Neutral"

            if macd > signal:
                tech_score += 1
                macd_signal = "Bullish"
            else:
                tech_score -= 1
                macd_signal = "Bearish"

            # ========== FUNDAMENTAL ANALYSIS ==========
            pe_ratio = info.get('trailingPE', info.get('forwardPE', 0))
            pb_ratio = info.get('priceToBook', 0)
            peg_ratio = info.get('pegRatio', 0)
            market_cap = info.get('marketCap', 0)
            dividend_yield = info.get('dividendYield', 0)

            roe = info.get('returnOnEquity', 0)
            roa = info.get('returnOnAssets', 0)
            profit_margin = info.get('profitMargins', 0)
            operating_margin = info.get('operatingMargins', 0)
            eps = info.get('trailingEps', 0)

            revenue_growth = info.get('revenueGrowth', 0)
            earnings_growth = info.get('earningsGrowth', 0)

            debt_to_equity = info.get('debtToEquity', 0)
            current_ratio = info.get('currentRatio', 0)
            quick_ratio = info.get('quickRatio', 0)

            beta = info.get('beta', 1.0)
            analyst_recommendation = info.get('recommendationKey', 'hold')
            target_price = info.get('targetMeanPrice', current_price)

            fund_score = self.get_fundamental_score(info)

            # ========== COMBINED SCORING ==========
            tech_score_normalized = ((tech_score + 6) / 12) * 100
            combined_score = (tech_score_normalized * 0.5) + (fund_score * 0.5)

            if combined_score >= 75:
                rating = "⭐⭐⭐⭐⭐ STRONG BUY"
                recommendation = "STRONG BUY"
            elif combined_score >= 55:
                rating = "⭐⭐⭐⭐ BUY"
                recommendation = "BUY"
            elif combined_score >= 45:
                rating = "⭐⭐⭐ HOLD"
                recommendation = "HOLD"
            elif combined_score >= 30:
                rating = "⭐⭐ SELL"
                recommendation = "SELL"
            else:
                rating = "⭐ STRONG SELL"
                recommendation = "STRONG SELL"

            if recommendation in ["STRONG BUY", "BUY"]:
                stop_loss = support * 0.97
                sl_percentage = ((current_price - stop_loss) / current_price) * 100
                target_1 = resistance
                target_2 = min(target_price, resistance * 1.05) if target_price > current_price else resistance * 1.05
                upside = ((target_1 - current_price) / current_price) * 100
            else:
                stop_loss = resistance * 1.03
                sl_percentage = ((stop_loss - current_price) / current_price) * 100
                target_1 = support
                target_2 = support * 0.95
                upside = ((current_price - target_1) / current_price) * 100

            risk = abs(current_price - stop_loss)
            reward = abs(target_1 - current_price)
            risk_reward = reward / risk if risk > 0 else 0

            if fund_score >= 80:
                quality = "Excellent"
            elif fund_score >= 60:
                quality = "Good"
            elif fund_score >= 40:
                quality = "Average"
            else:
                quality = "Poor"

            result = {
                'Symbol': symbol.replace('.NS', ''),
                'Name': name,
                'Price': round(current_price, 2),
                'RSI': round(rsi, 2),
                'RSI_Signal': rsi_signal,
                'MACD': macd_signal,
                'SMA_20': round(sma_20, 2),
                'SMA_50': round(sma_50, 2),
                'SMA_200': round(sma_200, 2),
                'Support': round(support, 2),
                'Resistance': round(resistance, 2),
                '52W_High': round(high_52w, 2),
                '52W_Low': round(low_52w, 2),
                'Tech_Score': tech_score,
                'Tech_Score_Norm': round(tech_score_normalized, 1),
                'PE_Ratio': round(pe_ratio, 2) if pe_ratio else 0,
                'PB_Ratio': round(pb_ratio, 2) if pb_ratio else 0,
                'PEG_Ratio': round(peg_ratio, 2) if peg_ratio else 0,
                'ROE': round(roe * 100, 2) if roe else 0,
                'ROA': round(roa * 100, 2) if roa else 0,
                'Profit_Margin': round(profit_margin * 100, 2) if profit_margin else 0,
                'Operating_Margin': round(operating_margin * 100, 2) if operating_margin else 0,
                'EPS': round(eps, 2) if eps else 0,
                'Dividend_Yield': round(dividend_yield * 100, 2) if dividend_yield else 0,
                'Revenue_Growth': round(revenue_growth * 100, 2) if revenue_growth else 0,
                'Earnings_Growth': round(earnings_growth * 100, 2) if earnings_growth else 0,
                'Debt_to_Equity': round(debt_to_equity, 2) if debt_to_equity else 0,
                'Current_Ratio': round(current_ratio, 2) if current_ratio else 0,
                'Market_Cap': round(market_cap / 1e12, 2) if market_cap else 0,
                'Beta': round(beta, 2) if beta else 1.0,
                'Fund_Score': round(fund_score, 1),
                'Quality': quality,
                'Combined_Score': round(combined_score, 1),
                'Rating': rating,
                'Recommendation': recommendation,
                'Stop_Loss': round(stop_loss, 2),
                'SL_Percentage': round(sl_percentage, 2),
                'Target_1': round(target_1, 2),
                'Target_2': round(target_2, 2),
                'Target_Price': round(target_price, 2) if target_price else 0,
                'Upside': round(upside, 2),
                'Risk_Reward': round(risk_reward, 2),
            }

            return result

        except Exception as e:
            return None

    def analyze_all_stocks(self):
        """Analyze all Nifty 50 stocks"""
        print(f"🔍 Analyzing {len(self.nifty50_stocks)} NIFTY 50 stocks...")

        for idx, (symbol, name) in enumerate(self.nifty50_stocks.items(), 1):
            result = self.analyze_stock(symbol, name)
            if result:
                self.results.append(result)
            print(f"  [{idx}/{len(self.nifty50_stocks)}] {name}")

        print(f"✅ Analysis complete: {len(self.results)} stocks analyzed\n")

    def get_top_recommendations(self):
        """Get top 10 buy and sell recommendations"""
        df = pd.DataFrame(self.results)
        top_buys = df[df['Recommendation'].isin(['STRONG BUY', 'BUY'])].nlargest(10, 'Combined_Score')
        top_sells = df[df['Recommendation'].isin(['STRONG SELL', 'SELL'])].nsmallest(10, 'Combined_Score')
        return top_buys, top_sells

    # =========================================================
    #   AURORA GLASS THEME — SHARED CSS
    # =========================================================
    def _aurora_css(self):
        return """
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&display=swap');

        *, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: 'Outfit', sans-serif;
            background: linear-gradient(135deg, #0d1f1e 0%, #0a1a2e 50%, #0f1e2a 100%);
            min-height: 100vh;
            padding: 24px;
            color: #b2dfdb;
        }

        /* ---- Ambient glow orbs ---- */
        body::before {
            content: '';
            position: fixed; top: -120px; right: -120px;
            width: 500px; height: 500px;
            background: radial-gradient(circle, rgba(32,178,170,.18) 0%, transparent 65%);
            border-radius: 50%; pointer-events: none; z-index: 0;
        }
        body::after {
            content: '';
            position: fixed; bottom: -100px; left: -80px;
            width: 400px; height: 400px;
            background: radial-gradient(circle, rgba(0,150,136,.14) 0%, transparent 65%);
            border-radius: 50%; pointer-events: none; z-index: 0;
        }

        .wrapper {
            max-width: 1300px;
            margin: 0 auto;
            position: relative; z-index: 1;
        }

        /* ---- HEADER ---- */
        .header {
            background: rgba(255,255,255,.05);
            border: 1px solid rgba(77,208,196,.2);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
            border-radius: 20px;
            padding: 36px 40px;
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 20px;
            flex-wrap: wrap;
        }
        .header-left {}
        .header-eyebrow {
            font-size: 11px;
            letter-spacing: 3px;
            color: #4dd0c4;
            text-transform: uppercase;
            margin-bottom: 8px;
            opacity: .8;
        }
        .header-title {
            font-size: 32px;
            font-weight: 700;
            color: #e0f2f1;
            line-height: 1.1;
        }
        .header-title span { color: #4dd0c4; }
        .header-sub {
            font-size: 14px;
            color: #80cbc4;
            margin-top: 6px;
            font-weight: 300;
        }
        .header-badge {
            background: rgba(32,178,170,.15);
            border: 1px solid rgba(77,208,196,.35);
            color: #4dd0c4;
            padding: 10px 22px;
            border-radius: 30px;
            font-size: 13px;
            font-weight: 600;
            backdrop-filter: blur(8px);
            white-space: nowrap;
        }
        .header-badge::before { content: '● '; font-size: 9px; }

        /* ---- LIVE CLOCK ---- */
        #live-clock {
            font-size: 14px;
            color: #4dd0c4;
            font-weight: 500;
            margin-top: 8px;
            letter-spacing: 0.5px;
        }

        /* ---- STAT CARDS ---- */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 16px;
            margin-bottom: 28px;
        }
        .stat-card {
            background: rgba(255,255,255,.06);
            border: 1px solid rgba(77,208,196,.15);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border-radius: 16px;
            padding: 22px 18px;
            text-align: center;
            transition: transform .25s ease, border-color .25s ease, box-shadow .25s ease;
            position: relative; overflow: hidden;
        }
        .stat-card::before {
            content: '';
            position: absolute; top: 0; left: 0; right: 0; height: 2px;
            background: linear-gradient(90deg, transparent, rgba(77,208,196,.6), transparent);
        }
        .stat-card:hover {
            transform: translateY(-4px);
            border-color: rgba(77,208,196,.35);
            box-shadow: 0 8px 28px rgba(32,178,170,.15);
        }
        .stat-card .num {
            font-size: 42px;
            font-weight: 700;
            color: #4dd0c4;
            line-height: 1;
            text-shadow: 0 0 24px rgba(77,208,196,.4);
        }
        .stat-card .lbl {
            font-size: 11px;
            color: #80cbc4;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            margin-top: 8px;
            font-weight: 400;
            opacity: .8;
        }

        /* ---- SECTION ---- */
        .section { margin-bottom: 32px; }
        .section-header {
            display: flex; align-items: center; gap: 12px;
            margin-bottom: 16px;
        }
        .section-dot {
            width: 10px; height: 10px; border-radius: 50%;
            box-shadow: 0 0 8px currentColor;
        }
        .section-dot.buy { background: #4dd0c4; color: #4dd0c4; }
        .section-dot.sell { background: #ef9a9a; color: #ef9a9a; }
        .section-title {
            font-size: 18px;
            font-weight: 600;
            color: #e0f2f1;
            letter-spacing: .5px;
        }
        .section-title.sell { color: #ef9a9a; }
        .section-count {
            margin-left: auto;
            background: rgba(77,208,196,.12);
            border: 1px solid rgba(77,208,196,.25);
            color: #4dd0c4;
            font-size: 11px;
            padding: 4px 12px;
            border-radius: 20px;
        }
        .section-count.sell {
            background: rgba(239,154,154,.1);
            border-color: rgba(239,154,154,.25);
            color: #ef9a9a;
        }

        /* ---- TABLE ---- */
        .table-wrap {
            background: rgba(255,255,255,.04);
            border: 1px solid rgba(77,208,196,.14);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-radius: 16px;
            overflow: hidden;
        }
        table { width: 100%; border-collapse: collapse; }
        thead tr { background: rgba(32,178,170,.14); }
        th {
            padding: 14px 16px;
            text-align: left;
            font-size: 11px;
            font-weight: 600;
            color: #4dd0c4;
            letter-spacing: 1.2px;
            text-transform: uppercase;
            white-space: nowrap;
        }
        td {
            padding: 13px 16px;
            border-bottom: 1px solid rgba(77,208,196,.07);
            color: #b2dfdb;
            font-size: 13px;
            vertical-align: middle;
        }
        tbody tr:last-child td { border-bottom: none; }
        tbody tr { transition: background .15s ease; }
        tbody tr:hover { background: rgba(77,208,196,.06); }

        .stock-name { font-weight: 600; color: #e0f2f1; }
        .price { font-weight: 500; }

        /* ---- SCORE PILL ---- */
        .score-pill {
            display: inline-block;
            background: rgba(32,178,170,.18);
            border: 1px solid rgba(77,208,196,.28);
            color: #4dd0c4;
            padding: 4px 11px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }

        /* ---- SIGNAL COLORS ---- */
        .buy  { color: #4dd0c4; font-weight: 700; }
        .sell { color: #ef9a9a; font-weight: 700; }
        .hold { color: #ffd54f; font-weight: 600; }
        .neutral { color: #90a4ae; }

        /* ---- QUALITY BADGE ---- */
        .badge {
            display: inline-block;
            padding: 4px 11px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-excellent { background: rgba(77,208,196,.2);  border: 1px solid rgba(77,208,196,.35);  color: #4dd0c4; }
        .badge-good      { background: rgba(129,199,132,.15); border: 1px solid rgba(129,199,132,.3);  color: #81c784; }
        .badge-average   { background: rgba(255,213,79,.12);  border: 1px solid rgba(255,213,79,.28);  color: #ffd54f; }
        .badge-poor      { background: rgba(239,154,154,.14); border: 1px solid rgba(239,154,154,.28); color: #ef9a9a; }

        /* ---- RSI / MACD badges ---- */
        .rsi-oversold  { color: #4dd0c4; font-weight: 700; }
        .rsi-overbought{ color: #ef9a9a; font-weight: 700; }
        .rsi-neutral   { color: #ffd54f; font-weight: 600; }

        /* ---- DISCLAIMER ---- */
        .disclaimer {
            background: rgba(255,213,79,.05);
            border: 1px solid rgba(255,213,79,.2);
            border-radius: 14px;
            padding: 22px 26px;
            margin-top: 32px;
        }
        .disclaimer h3 { color: #ef9a9a; font-size: 14px; margin-bottom: 10px; letter-spacing: 1px; }
        .disclaimer p  { color: #80cbc4; font-size: 13px; line-height: 1.7; }
        .disclaimer ul { margin-left: 20px; margin-top: 8px; }
        .disclaimer li { color: #80cbc4; font-size: 13px; line-height: 1.9; }

        /* ---- FOOTER ---- */
        .footer {
            text-align: center;
            padding: 28px 0 10px;
            color: #4dd0c4;
            font-size: 12px;
            opacity: .65;
            letter-spacing: 1px;
        }

        /* ---- RESPONSIVE ---- */
        @media (max-width: 900px) {
            body { padding: 14px; }
            .header { padding: 24px 20px; }
            .header-title { font-size: 24px; }
            th, td { padding: 11px 10px; font-size: 12px; }
        }
        @media (max-width: 600px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .header { flex-direction: column; align-items: flex-start; gap: 14px; }
        }
        """

    # =========================================================
    #   GITHUB PAGES HTML  —  AURORA GLASS
    # =========================================================
    def generate_github_pages_html(self, output_file='index.html'):
        """Generate Aurora Glass HTML for GitHub Pages"""

        df = pd.DataFrame(self.results)
        top_buys, top_sells = self.get_top_recommendations()

        now = self.get_ist_time()
        next_update = "4:30 PM" if now.hour < 12 else "9:30 AM (Next Day)"

        strong_buy_count = len(df[df['Recommendation'] == 'STRONG BUY'])
        buy_count        = len(df[df['Recommendation'] == 'BUY'])
        hold_count       = len(df[df['Recommendation'] == 'HOLD'])
        sell_count       = len(df[df['Recommendation'] == 'SELL'])
        strong_sell_count= len(df[df['Recommendation'] == 'STRONG SELL'])

        # Static generation timestamp (used only as fallback text)
        generated_on = now.strftime('%d %B %Y, %I:%M %p IST')

        css = self._aurora_css()

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NIFTY 50 Analysis — Aurora Glass</title>
<style>{css}
html {{ scroll-behavior: smooth; }}
</style>
</head>
<body>
<div class="wrapper">

  <!-- HEADER -->
  <header class="header">
    <div class="header-left">
      <div class="header-eyebrow">NIFTY 50 · Market Intelligence</div>
      <div class="header-title">💎 Stock <span>Analysis</span> Report</div>
      <div class="header-sub">
        Report Generated on {generated_on}
      </div>
      <!-- Live IST Clock -->
      <div id="live-clock">🕐 Current IST Time: Loading...</div>
    </div>
    <div class="header-badge">Live Market Data</div>
  </header>

  <!-- STAT CARDS -->
  <div class="stats-grid">
    <div class="stat-card">
      <div class="num">{len(self.results)}</div>
      <div class="lbl">Stocks Analyzed</div>
    </div>
    <div class="stat-card">
      <div class="num">{strong_buy_count}</div>
      <div class="lbl">Strong Buy</div>
    </div>
    <div class="stat-card">
      <div class="num">{buy_count}</div>
      <div class="lbl">Buy</div>
    </div>
    <div class="stat-card">
      <div class="num">{hold_count}</div>
      <div class="lbl">Hold</div>
    </div>
    <div class="stat-card">
      <div class="num">{sell_count}</div>
      <div class="lbl">Sell</div>
    </div>
    <div class="stat-card">
      <div class="num">{strong_sell_count}</div>
      <div class="lbl">Strong Sell</div>
    </div>
  </div>
"""

        # ---- TOP 10 BUY ----
        if not top_buys.empty:
            html += f"""
  <!-- BUY SECTION -->
  <div class="section">
    <div class="section-header">
      <div class="section-dot buy"></div>
      <div class="section-title">Top 10 Buy Recommendations</div>
      <div class="section-count">{len(top_buys)} stocks</div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Stock</th>
            <th>Price</th>
            <th>Rating</th>
            <th>Score</th>
            <th>Upside %</th>
            <th>Target 1</th>
            <th>Stop Loss</th>
            <th>R:R Ratio</th>
            <th>Quality</th>
          </tr>
        </thead>
        <tbody>
"""
            for rank, (_, row) in enumerate(top_buys.iterrows(), 1):
                upside_cls = "buy" if row['Upside'] >= 0 else "sell"
                q = row['Quality'].lower()
                badge_cls = f"badge-{q}"
                html += f"""
          <tr>
            <td class="neutral">{rank}</td>
            <td class="stock-name">{row['Name']}</td>
            <td class="price">₹{row['Price']:,.2f}</td>
            <td class="buy" style="font-size:12px">{row['Rating']}</td>
            <td><span class="score-pill">{row['Combined_Score']:.0f}</span></td>
            <td class="{upside_cls}">{row['Upside']:+.1f}%</td>
            <td>₹{row['Target_1']:,.2f}</td>
            <td>₹{row['Stop_Loss']:,.2f}</td>
            <td class="neutral">{row['Risk_Reward']:.2f}x</td>
            <td><span class="badge {badge_cls}">{row['Quality']}</span></td>
          </tr>
"""
            html += """
        </tbody>
      </table>
    </div>
  </div>
"""

        # ---- TOP 10 SELL ----
        if not top_sells.empty:
            html += f"""
  <!-- SELL SECTION -->
  <div class="section">
    <div class="section-header">
      <div class="section-dot sell"></div>
      <div class="section-title sell">Top 10 Sell Recommendations</div>
      <div class="section-count sell">{len(top_sells)} stocks</div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Stock</th>
            <th>Price</th>
            <th>Rating</th>
            <th>Score</th>
            <th>RSI</th>
            <th>RSI Signal</th>
            <th>MACD</th>
            <th>Quality</th>
          </tr>
        </thead>
        <tbody>
"""
            for rank, (_, row) in enumerate(top_sells.iterrows(), 1):
                rsi_val = row['RSI']
                if rsi_val > 70:
                    rsi_cls = "rsi-overbought"
                elif rsi_val < 30:
                    rsi_cls = "rsi-oversold"
                else:
                    rsi_cls = "rsi-neutral"
                macd_cls = "buy" if row['MACD'] == "Bullish" else "sell"
                q = row['Quality'].lower()
                badge_cls = f"badge-{q}"
                html += f"""
          <tr>
            <td class="neutral">{rank}</td>
            <td class="stock-name">{row['Name']}</td>
            <td class="price">₹{row['Price']:,.2f}</td>
            <td class="sell" style="font-size:12px">{row['Rating']}</td>
            <td><span class="score-pill">{row['Combined_Score']:.0f}</span></td>
            <td class="{rsi_cls}">{row['RSI']:.1f}</td>
            <td class="{rsi_cls}">{row['RSI_Signal']}</td>
            <td class="{macd_cls}">{row['MACD']}</td>
            <td><span class="badge {badge_cls}">{row['Quality']}</span></td>
          </tr>
"""
            html += """
        </tbody>
      </table>
    </div>
  </div>
"""

        # ---- FULL TABLE ----
        all_df = pd.DataFrame(self.results).sort_values('Combined_Score', ascending=False)
        html += """
  <!-- ALL STOCKS SECTION -->
  <div class="section">
    <div class="section-header">
      <div class="section-dot buy"></div>
      <div class="section-title">All NIFTY 50 Stocks — Complete Analysis</div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Stock</th>
            <th>Price</th>
            <th>Score</th>
            <th>PE</th>
            <th>PB</th>
            <th>ROE %</th>
            <th>RSI</th>
            <th>MACD</th>
            <th>Rating</th>
            <th>Quality</th>
          </tr>
        </thead>
        <tbody>
"""
        for rank, (_, row) in enumerate(all_df.iterrows(), 1):
            rec = row['Recommendation']
            if rec in ["STRONG BUY", "BUY"]:
                rec_cls = "buy"
            elif rec in ["STRONG SELL", "SELL"]:
                rec_cls = "sell"
            else:
                rec_cls = "hold"
            rsi_val = row['RSI']
            if rsi_val > 70:
                rsi_cls = "rsi-overbought"
            elif rsi_val < 30:
                rsi_cls = "rsi-oversold"
            else:
                rsi_cls = "rsi-neutral"
            q = row['Quality'].lower()
            html += f"""
          <tr>
            <td class="neutral">{rank}</td>
            <td class="stock-name">{row['Name']}</td>
            <td class="price">₹{row['Price']:,.2f}</td>
            <td><span class="score-pill">{row['Combined_Score']:.0f}</span></td>
            <td class="neutral">{row['PE_Ratio']:.1f}</td>
            <td class="neutral">{row['PB_Ratio']:.1f}</td>
            <td class="neutral">{row['ROE']:.1f}%</td>
            <td class="{rsi_cls}">{row['RSI']:.1f}</td>
            <td class="{'buy' if row['MACD']=='Bullish' else 'sell'}">{row['MACD']}</td>
            <td class="{rec_cls}" style="font-size:12px">{row['Rating']}</td>
            <td><span class="badge badge-{q}">{row['Quality']}</span></td>
          </tr>
"""
        html += """
        </tbody>
      </table>
    </div>
  </div>
"""

        html += f"""
  <!-- DISCLAIMER -->
  <div class="disclaimer">
    <h3>⚠ DISCLAIMER</h3>
    <p>This analysis is for <strong>educational purposes only</strong> and does <strong>NOT</strong> constitute financial advice.</p>
    <ul>
      <li>Always conduct your own research before investing.</li>
      <li>Consult a SEBI-registered financial advisor for personalised guidance.</li>
      <li>Use proper risk management and honour stop-loss levels.</li>
      <li>Never invest more than you can afford to lose.</li>
    </ul>
  </div>

  <!-- FOOTER -->
  <div class="footer">
    © 2025 NIFTY 50 Analyzer &nbsp;|&nbsp; Next Update: {next_update} IST
  </div>

</div><!-- /wrapper -->

<!-- LIVE IST CLOCK SCRIPT -->
<script>
  function updateISTClock() {{
    const now = new Date();
    const options = {{
      timeZone: 'Asia/Kolkata',
      day: '2-digit',
      month: 'long',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true
    }};
    const istTime = now.toLocaleString('en-IN', options);
    const el = document.getElementById('live-clock');
    if (el) {{
      el.textContent = '🕐 Current IST Time: ' + istTime;
    }}
  }}
  updateISTClock();
  setInterval(updateISTClock, 1000);
</script>

</body>
</html>
"""

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)

        print(f"✅ Aurora Glass HTML generated: {output_file}\n")
        return output_file

    # =========================================================
    #   EMAIL HTML  —  AURORA GLASS  (inline styles for email)
    # =========================================================
    def generate_email_html(self):
        """Generate Aurora Glass HTML email (table-based, inline styles)"""

        df = pd.DataFrame(self.results)
        top_buys, top_sells = self.get_top_recommendations()

        now = self.get_ist_time()
        next_update = "4:30 PM" if now.hour < 12 else "9:30 AM (Next Day)"

        strong_buy_count  = len(df[df['Recommendation'] == 'STRONG BUY'])
        buy_count         = len(df[df['Recommendation'] == 'BUY'])
        hold_count        = len(df[df['Recommendation'] == 'HOLD'])
        sell_count        = len(df[df['Recommendation'] == 'SELL'])
        strong_sell_count = len(df[df['Recommendation'] == 'STRONG SELL'])

        # Inline style helpers
        bg_outer   = "#0d1f1e"
        bg_card    = "#0f2020"
        teal       = "#4dd0c4"
        teal_dim   = "#80cbc4"
        teal_bg    = "rgba(32,178,170,0.14)"
        teal_border= "rgba(77,208,196,0.2)"
        text_main  = "#e0f2f1"
        text_body  = "#b2dfdb"
        green      = "#4dd0c4"
        red        = "#ef9a9a"
        yellow     = "#ffd54f"
        row_border = "rgba(77,208,196,0.07)"
        divider    = "rgba(77,208,196,0.14)"

        generated_on = now.strftime('%d %B %Y, %I:%M %p IST')

        def quality_color(q):
            return {"Excellent": teal, "Good": "#81c784", "Average": yellow, "Poor": red}.get(q, text_body)

        html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:{bg_outer};font-family:'Segoe UI',Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="{bg_outer}">
<tr><td align="center" style="padding:24px 16px;">

  <!-- OUTER CARD -->
  <table width="680" cellpadding="0" cellspacing="0" border="0"
         style="background:{bg_card};border:1px solid {teal_border};border-radius:20px;overflow:hidden;">

    <!-- HEADER -->
    <tr>
      <td style="background:linear-gradient(135deg,rgba(13,31,30,.9),rgba(10,26,46,.9));
                 border-bottom:1px solid {teal_border};padding:32px 36px;">
        <p style="font-size:11px;letter-spacing:3px;color:{teal};text-transform:uppercase;
                  margin:0 0 8px">NIFTY 50 · MARKET INTELLIGENCE</p>
        <p style="font-size:26px;font-weight:700;color:{text_main};margin:0 0 6px">
          💎 Stock Analysis Report</p>
        <p style="font-size:13px;color:{teal_dim};margin:0;font-weight:300">
          Report Generated on {generated_on}</p>
      </td>
    </tr>

    <!-- STAT ROW -->
    <tr>
      <td style="padding:24px 36px;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
"""
        stats = [
            (len(self.results), "Analyzed"),
            (strong_buy_count,  "Strong Buy"),
            (buy_count,         "Buy"),
            (hold_count,        "Hold"),
            (sell_count,        "Sell"),
            (strong_sell_count, "Strong Sell"),
        ]
        for num, lbl in stats:
            html += f"""
            <td align="center" style="padding:0 4px;">
              <table width="100%" cellpadding="12" cellspacing="0" border="0"
                     style="background:{teal_bg};border:1px solid {teal_border};border-radius:12px;">
                <tr>
                  <td align="center">
                    <p style="font-size:30px;font-weight:700;color:{teal};margin:0;
                               text-shadow:0 0 16px rgba(77,208,196,.3)">{num}</p>
                    <p style="font-size:10px;color:{teal_dim};margin:4px 0 0;
                               text-transform:uppercase;letter-spacing:1px">{lbl}</p>
                  </td>
                </tr>
              </table>
            </td>
"""
        html += """
          </tr>
        </table>
      </td>
    </tr>
"""

        # ---- BUY TABLE ----
        if not top_buys.empty:
            html += f"""
    <!-- BUY SECTION -->
    <tr>
      <td style="padding:4px 36px 20px;">
        <p style="font-size:13px;font-weight:600;color:{teal};
                  letter-spacing:1px;margin:0 0 12px;text-transform:uppercase">
          ● Top 10 Buy Recommendations</p>
        <table width="100%" cellpadding="0" cellspacing="0" border="0"
               style="border:1px solid {divider};border-radius:12px;overflow:hidden;">
          <tr style="background:{teal_bg}">
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{teal};letter-spacing:1px;font-weight:600">STOCK</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{teal};letter-spacing:1px;font-weight:600">PRICE</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{teal};letter-spacing:1px;font-weight:600">SCORE</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{teal};letter-spacing:1px;font-weight:600">UPSIDE</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{teal};letter-spacing:1px;font-weight:600">TARGET</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{teal};letter-spacing:1px;font-weight:600">STOP LOSS</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{teal};letter-spacing:1px;font-weight:600">QUALITY</th>
          </tr>
"""
            for i, (_, row) in enumerate(top_buys.iterrows()):
                row_bg = "rgba(255,255,255,0.02)" if i % 2 == 0 else "rgba(32,178,170,0.04)"
                upside_color = green if row['Upside'] >= 0 else red
                q_color = quality_color(row['Quality'])
                html += f"""
          <tr style="background:{row_bg};border-top:1px solid {row_border}">
            <td style="padding:12px 10px;color:{text_main};font-weight:600;font-size:13px">{row['Name']}</td>
            <td style="padding:12px 10px;color:{text_body};font-size:13px">₹{row['Price']:,.2f}</td>
            <td style="padding:12px 10px">
              <span style="background:rgba(77,208,196,.18);border:1px solid rgba(77,208,196,.3);
                           color:{teal};padding:3px 9px;border-radius:20px;font-size:11px;font-weight:600">
                {row['Combined_Score']:.0f}
              </span>
            </td>
            <td style="padding:12px 10px;color:{upside_color};font-weight:700;font-size:14px">
              {row['Upside']:+.1f}%</td>
            <td style="padding:12px 10px;color:{text_body};font-size:13px">₹{row['Target_1']:,.2f}</td>
            <td style="padding:12px 10px;color:{text_body};font-size:13px">₹{row['Stop_Loss']:,.2f}</td>
            <td style="padding:12px 10px">
              <span style="color:{q_color};font-size:11px;font-weight:600">{row['Quality']}</span>
            </td>
          </tr>
"""
            html += """
        </table>
      </td>
    </tr>
"""

        # ---- SELL TABLE ----
        if not top_sells.empty:
            html += f"""
    <!-- SELL SECTION -->
    <tr>
      <td style="padding:4px 36px 20px;">
        <p style="font-size:13px;font-weight:600;color:{red};
                  letter-spacing:1px;margin:0 0 12px;text-transform:uppercase">
          ● Top 10 Sell Recommendations</p>
        <table width="100%" cellpadding="0" cellspacing="0" border="0"
               style="border:1px solid rgba(239,154,154,.18);border-radius:12px;overflow:hidden;">
          <tr style="background:rgba(239,154,154,.1)">
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{red};letter-spacing:1px;font-weight:600">STOCK</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{red};letter-spacing:1px;font-weight:600">PRICE</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{red};letter-spacing:1px;font-weight:600">SCORE</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{red};letter-spacing:1px;font-weight:600">RSI</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{red};letter-spacing:1px;font-weight:600">RSI SIGNAL</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{red};letter-spacing:1px;font-weight:600">MACD</th>
            <th style="padding:12px 10px;text-align:left;font-size:10px;
                        color:{red};letter-spacing:1px;font-weight:600">QUALITY</th>
          </tr>
"""
            for i, (_, row) in enumerate(top_sells.iterrows()):
                row_bg = "rgba(255,255,255,0.02)" if i % 2 == 0 else "rgba(239,154,154,0.04)"
                rsi_val = row['RSI']
                if rsi_val > 70:
                    rsi_color = red
                elif rsi_val < 30:
                    rsi_color = green
                else:
                    rsi_color = yellow
                macd_color = green if row['MACD'] == "Bullish" else red
                q_color = quality_color(row['Quality'])
                html += f"""
          <tr style="background:{row_bg};border-top:1px solid rgba(239,154,154,0.08)">
            <td style="padding:12px 10px;color:{text_main};font-weight:600;font-size:13px">{row['Name']}</td>
            <td style="padding:12px 10px;color:{text_body};font-size:13px">₹{row['Price']:,.2f}</td>
            <td style="padding:12px 10px">
              <span style="background:rgba(239,154,154,.15);border:1px solid rgba(239,154,154,.3);
                           color:{red};padding:3px 9px;border-radius:20px;font-size:11px;font-weight:600">
                {row['Combined_Score']:.0f}
              </span>
            </td>
            <td style="padding:12px 10px;color:{rsi_color};font-weight:700;font-size:14px">
              {row['RSI']:.1f}</td>
            <td style="padding:12px 10px;color:{rsi_color};font-size:13px;font-weight:600">
              {row['RSI_Signal']}</td>
            <td style="padding:12px 10px;color:{macd_color};font-size:13px;font-weight:600">
              {row['MACD']}</td>
            <td style="padding:12px 10px">
              <span style="color:{q_color};font-size:11px;font-weight:600">{row['Quality']}</span>
            </td>
          </tr>
"""
            html += """
        </table>
      </td>
    </tr>
"""

        html += f"""
    <!-- DISCLAIMER -->
    <tr>
      <td style="padding:4px 36px 28px;">
        <table width="100%" cellpadding="16" cellspacing="0" border="0"
               style="background:rgba(255,213,79,.05);border:1px solid rgba(255,213,79,.2);border-radius:12px;">
          <tr>
            <td>
              <p style="color:{red};font-size:12px;font-weight:700;
                         letter-spacing:1px;margin:0 0 8px">⚠ DISCLAIMER</p>
              <p style="color:{text_body};font-size:12px;line-height:1.7;margin:0">
                This analysis is for <strong>educational purposes only</strong> and does
                <strong>NOT</strong> constitute financial advice. Always do your own research,
                consult a SEBI-registered advisor, use stop-loss levels, and never invest
                more than you can afford to lose.
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- FOOTER -->
    <tr>
      <td align="center"
          style="padding:16px 36px 24px;border-top:1px solid {divider}">
        <p style="color:{teal};font-size:11px;margin:0;opacity:.7;letter-spacing:1px">
          © 2025 NIFTY 50 Analyzer
          &nbsp;|&nbsp; Next Update: {next_update} IST
        </p>
      </td>
    </tr>

  </table><!-- /OUTER CARD -->

</td></tr>
</table><!-- /body table -->

</body>
</html>
"""
        return html

    # =========================================================
    #   SEND EMAIL
    # =========================================================
    def send_email(self, to_email):
        """Send Aurora Glass email report"""
        try:
            from_email = os.environ.get('GMAIL_USER')
            password   = os.environ.get('GMAIL_APP_PASSWORD')

            if not from_email or not password:
                print("❌ Gmail credentials not found in environment variables")
                print("   Set GMAIL_USER and GMAIL_APP_PASSWORD")
                return False

            now = self.get_ist_time()
            generated_on = now.strftime('%d %b %Y, %I:%M %p IST')

            msg = MIMEMultipart('alternative')
            msg['From']    = from_email
            msg['To']      = to_email
            msg['Subject'] = (
                f"💎 NIFTY 50 Stock Analysis — Report Generated on {generated_on}"
            )

            html_body = self.generate_email_html()
            msg.attach(MIMEText(html_body, 'html'))

            print(f"📧 Sending email to {to_email}...")
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
            server.quit()

            print("✅ Email sent successfully!\n")
            return True

        except Exception as e:
            print(f"❌ Error sending email: {e}\n")
            return False

    # =========================================================
    #   MAIN RUNNER
    # =========================================================
    def generate_complete_report(self, send_email_flag=True,
                                 recipient_email=None,
                                 generate_github_pages=True):
        """Generate complete analysis report"""
        ist_time = self.get_ist_time()

        print("=" * 70)
        print("💎 NIFTY 50 STOCK ANALYZER")
        print(f"   Report Generated on: {ist_time.strftime('%d %b %Y, %I:%M %p IST')}")
        print("=" * 70)
        print()

        self.analyze_all_stocks()

        if generate_github_pages:
            self.generate_github_pages_html('index.html')

        if send_email_flag and recipient_email:
            self.send_email(recipient_email)

        print("=" * 70)
        print("✅ ANALYSIS COMPLETE — Report Ready!")
        print("=" * 70)


# =========================================================
#   ENTRY POINT
# =========================================================
def main():
    analyzer = Nifty50CompleteAnalyzer()

    recipient = os.environ.get('RECIPIENT_EMAIL')
    if not recipient:
        print("⚠️  RECIPIENT_EMAIL environment variable not set.")
        print("   Reports will still be generated locally.\n")
        recipient = None

    analyzer.generate_complete_report(
        send_email_flag=True,
        recipient_email=recipient,
        generate_github_pages=True
    )


if __name__ == "__main__":
    main()
