"""
NIFTY 50 COMPLETE STOCK ANALYZER - ROYAL SAPPHIRE THEME
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
            'DIVISLAB.NS': 'Divi\'s Lab',
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
            'ADANIENT.NS': 'Adani Enterprises'
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
            
            # Valuation
            pe_ratio = info.get('trailingPE', info.get('forwardPE', 0))
            pb_ratio = info.get('priceToBook', 0)
            peg_ratio = info.get('pegRatio', 0)
            market_cap = info.get('marketCap', 0)
            dividend_yield = info.get('dividendYield', 0)
            
            # Profitability
            roe = info.get('returnOnEquity', 0)
            roa = info.get('returnOnAssets', 0)
            profit_margin = info.get('profitMargins', 0)
            operating_margin = info.get('operatingMargins', 0)
            eps = info.get('trailingEps', 0)
            
            # Growth
            revenue_growth = info.get('revenueGrowth', 0)
            earnings_growth = info.get('earningsGrowth', 0)
            
            # Financial Health
            debt_to_equity = info.get('debtToEquity', 0)
            current_ratio = info.get('currentRatio', 0)
            quick_ratio = info.get('quickRatio', 0)
            
            # Other
            beta = info.get('beta', 1.0)
            analyst_recommendation = info.get('recommendationKey', 'hold')
            target_price = info.get('targetMeanPrice', current_price)
            
            # Fundamental Score (0-100)
            fund_score = self.get_fundamental_score(info)
            
            # ========== COMBINED SCORING ==========
            
            # Normalize technical score to 0-100 scale
            tech_score_normalized = ((tech_score + 6) / 12) * 100
            
            # Combined score (50% technical + 50% fundamental)
            combined_score = (tech_score_normalized * 0.5) + (fund_score * 0.5)
            
            # Rating - ADJUSTED THRESHOLDS FOR MORE RECOMMENDATIONS
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
            
            # Stop Loss & Targets
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
            
            # Risk-Reward
            risk = abs(current_price - stop_loss)
            reward = abs(target_1 - current_price)
            risk_reward = reward / risk if risk > 0 else 0
            
            # Quality Assessment
            if fund_score >= 80:
                quality = "Excellent"
            elif fund_score >= 60:
                quality = "Good"
            elif fund_score >= 40:
                quality = "Average"
            else:
                quality = "Poor"
            
            result = {
                # Basic Info
                'Symbol': symbol.replace('.NS', ''),
                'Name': name,
                'Price': round(current_price, 2),
                
                # Technical
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
                
                # Fundamental
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
                
                # Combined
                'Combined_Score': round(combined_score, 1),
                'Rating': rating,
                'Recommendation': recommendation,
                
                # Trading
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
        
        # Top 10 Buy (highest combined scores from BUY + STRONG BUY)
        top_buys = df[df['Recommendation'].isin(['STRONG BUY', 'BUY'])].nlargest(10, 'Combined_Score')
        
        # Top 10 Sell (lowest combined scores from SELL + STRONG SELL)
        top_sells = df[df['Recommendation'].isin(['STRONG SELL', 'SELL'])].nsmallest(10, 'Combined_Score')
        
        return top_buys, top_sells
    
    def generate_github_pages_html(self, output_file='index.html'):
        """Generate beautiful HTML for GitHub Pages - ROYAL SAPPHIRE THEME"""
        df = pd.DataFrame(self.results)
        top_buys, top_sells = self.get_top_recommendations()
        
        now = self.get_ist_time()
        time_of_day = "Morning" if now.hour < 12 else "Evening"
        
        # Count recommendations
        strong_buy_count = len(df[df['Recommendation'] == 'STRONG BUY'])
        buy_count = len(df[df['Recommendation'] == 'BUY'])
        hold_count = len(df[df['Recommendation'] == 'HOLD'])
        sell_count = len(df[df['Recommendation'] == 'SELL'])
        strong_sell_count = len(df[df['Recommendation'] == 'STRONG SELL'])
        
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NIFTY 50 Stock Analysis - Royal Sapphire</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Georgia', serif;
            background: #0a0e27;
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: #1a237e;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(63,81,181,0.4);
            overflow: hidden;
            border: 3px solid #3f51b5;
        }}
        
        .header {{
            background: linear-gradient(135deg, #3f51b5 0%, #5c6bc0 100%);
            color: #ffffff;
            padding: 40px;
            text-align: center;
            position: relative;
        }}
        
        .header::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><circle cx="50" cy="50" r="2" fill="rgba(255,255,255,0.1)"/></svg>');
            background-size: 20px 20px;
            opacity: 0.3;
        }}
        
        .header h1 {{
            font-size: 42px;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
            font-weight: 800;
            position: relative;
            z-index: 1;
        }}
        
        .header p {{
            font-size: 18px;
            opacity: 0.95;
            font-weight: 600;
            position: relative;
            z-index: 1;
        }}
        
        .last-updated {{
            background: rgba(0,0,0,0.3);
            padding: 10px 20px;
            border-radius: 25px;
            display: inline-block;
            margin-top: 15px;
            font-size: 14px;
            font-weight: 600;
            border: 1px solid rgba(255,255,255,0.3);
            position: relative;
            z-index: 1;
        }}
        
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 40px;
            background: #0f1535;
        }}
        
        .summary-card {{
            background: linear-gradient(145deg, #283593, #1a237e);
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 4px 15px rgba(63,81,181,0.3);
            text-align: center;
            transition: transform 0.3s ease;
            border: 2px solid #3f51b5;
        }}
        
        .summary-card:hover {{
            transform: translateY(-5px) scale(1.02);
            box-shadow: 0 8px 25px rgba(63,81,181,0.5);
            border-color: #5c6bc0;
        }}
        
        .summary-card .number {{
            font-size: 48px;
            font-weight: bold;
            color: #9fa8da;
            margin-bottom: 10px;
            text-shadow: 0 0 20px rgba(159,168,218,0.5);
        }}
        
        .summary-card .label {{
            font-size: 14px;
            color: #c5cae9;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .content {{
            padding: 40px;
            background: #0f1535;
        }}
        
        .section {{
            margin-bottom: 50px;
        }}
        
        .section-title {{
            font-size: 32px;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 4px solid #3f51b5;
            color: #7986cb;
            text-shadow: 0 0 10px rgba(121,134,203,0.3);
        }}
        
        .section-title.sell {{
            border-bottom-color: #f44336;
            color: #ef5350;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            background: #1a237e;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.4);
        }}
        
        thead {{
            background: linear-gradient(135deg, #3f51b5 0%, #5c6bc0 100%);
            color: #ffffff;
        }}
        
        thead.sell {{
            background: linear-gradient(135deg, #f44336 0%, #ef5350 100%);
        }}
        
        th {{
            padding: 18px 15px;
            text-align: left;
            font-size: 13px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        
        td {{
            padding: 16px 15px;
            border-bottom: 1px solid #283593;
            color: #e8eaf6;
        }}
        
        tr:hover {{
            background-color: #283593;
        }}
        
        .stock-name {{
            font-weight: 600;
            color: #9fa8da;
        }}
        
        .rating {{
            font-weight: bold;
            font-size: 12px;
        }}
        
        .upside-positive {{
            color: #66bb6a;
            font-weight: bold;
            font-size: 16px;
        }}
        
        .upside-negative {{
            color: #ef5350;
            font-weight: bold;
            font-size: 16px;
        }}
        
        .rsi-overbought {{
            color: #ef5350;
            font-weight: bold;
            font-size: 16px;
        }}
        
        .rsi-oversold {{
            color: #66bb6a;
            font-weight: bold;
            font-size: 16px;
        }}
        
        .rsi-neutral {{
            color: #ffa726;
            font-weight: bold;
            font-size: 16px;
        }}
        
        .quality-badge {{
            padding: 6px 14px;
            border-radius: 20px;
            color: white;
            font-size: 11px;
            font-weight: bold;
            display: inline-block;
        }}
        
        .quality-excellent {{ background: #5c6bc0; }}
        .quality-good {{ background: #7986cb; }}
        .quality-average {{ background: #9fa8da; }}
        .quality-poor {{ background: #ef5350; }}
        
        .disclaimer {{
            background: #1a237e;
            border: 3px solid #ffa726;
            border-radius: 15px;
            padding: 30px;
            margin: 40px 0;
        }}
        
        .disclaimer h3 {{
            color: #ef5350;
            margin-bottom: 15px;
            font-size: 20px;
        }}
        
        .disclaimer p {{
            color: #e8eaf6;
        }}
        
        .disclaimer ul {{
            margin-left: 25px;
            margin-top: 15px;
            line-height: 1.8;
            color: #e8eaf6;
        }}
        
        .footer {{
            background: #0a0e27;
            color: #9fa8da;
            text-align: center;
            padding: 30px;
        }}
        
        .footer p {{
            margin: 5px 0;
        }}
        
        @media (max-width: 768px) {{
            .header h1 {{ font-size: 28px; }}
            .summary-grid {{ grid-template-columns: repeat(2, 1fr); }}
            table {{ font-size: 12px; }}
            th, td {{ padding: 10px 8px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>💎 NIFTY 50 Stock Analysis</h1>
            <p>{time_of_day} Market Report - Royal Sapphire Edition</p>
            <div class="last-updated">
                Last Updated: {now.strftime('%d %b %Y, %I:%M %p')} IST
            </div>
        </div>
        
        <!-- Summary Cards -->
        <div class="summary-grid">
            <div class="summary-card">
                <div class="number">{len(self.results)}</div>
                <div class="label">Stocks Analyzed</div>
            </div>
            <div class="summary-card">
                <div class="number">{strong_buy_count}</div>
                <div class="label">Strong Buy</div>
            </div>
            <div class="summary-card">
                <div class="number">{buy_count}</div>
                <div class="label">Buy</div>
            </div>
            <div class="summary-card">
                <div class="number">{hold_count}</div>
                <div class="label">Hold</div>
            </div>
        </div>
        
        <!-- Content -->
        <div class="content">
"""
        
        # Top 10 Buy Recommendations
        if not top_buys.empty:
            html += """
            <div class="section">
                <h2 class="section-title">🟢 TOP 10 BUY RECOMMENDATIONS</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Stock</th>
                            <th>Price</th>
                            <th>Rating</th>
                            <th>Score</th>
                            <th>Upside %</th>
                            <th>Target</th>
                            <th>Stop Loss</th>
                            <th>Quality</th>
                        </tr>
                    </thead>
                    <tbody>
"""
            for idx, row in top_buys.iterrows():
                upside_class = "upside-positive" if row['Upside'] > 0 else "upside-negative"
                
                quality_class = {
                    'Excellent': 'quality-excellent',
                    'Good': 'quality-good',
                    'Average': 'quality-average',
                    'Poor': 'quality-poor'
                }.get(row['Quality'], 'quality-average')
                
                html += f"""
                        <tr>
                            <td class="stock-name">{row['Name']}</td>
                            <td>₹{row['Price']:,.0f}</td>
                            <td class="rating">{row['Rating']}</td>
                            <td><strong>{row['Combined_Score']:.0f}</strong></td>
                            <td class="{upside_class}">{row['Upside']:+.1f}%</td>
                            <td>₹{row['Target_1']:,.0f}</td>
                            <td>₹{row['Stop_Loss']:,.0f}</td>
                            <td><span class="quality-badge {quality_class}">{row['Quality']}</span></td>
                        </tr>
"""
            html += """
                    </tbody>
                </table>
            </div>
"""
        
        # Top 10 Sell Recommendations
        if not top_sells.empty:
            html += """
            <div class="section">
                <h2 class="section-title sell">🔴 TOP 10 SELL RECOMMENDATIONS</h2>
                <table>
                    <thead class="sell">
                        <tr>
                            <th>Stock</th>
                            <th>Price</th>
                            <th>Rating</th>
                            <th>Score</th>
                            <th>RSI</th>
                            <th>MACD</th>
                            <th>Quality</th>
                        </tr>
                    </thead>
                    <tbody>
"""
            for idx, row in top_sells.iterrows():
                if row['RSI'] > 70:
                    rsi_class = "rsi-overbought"
                elif row['RSI'] < 30:
                    rsi_class = "rsi-oversold"
                else:
                    rsi_class = "rsi-neutral"
                
                quality_class = {
                    'Excellent': 'quality-excellent',
                    'Good': 'quality-good',
                    'Average': 'quality-average',
                    'Poor': 'quality-poor'
                }.get(row['Quality'], 'quality-average')
                
                html += f"""
                        <tr>
                            <td class="stock-name">{row['Name']}</td>
                            <td>₹{row['Price']:,.0f}</td>
                            <td class="rating">{row['Rating']}</td>
                            <td><strong>{row['Combined_Score']:.0f}</strong></td>
                            <td class="{rsi_class}">{row['RSI']:.0f}</td>
                            <td>{row['MACD']}</td>
                            <td><span class="quality-badge {quality_class}">{row['Quality']}</span></td>
                        </tr>
"""
            html += """
                    </tbody>
                </table>
            </div>
"""
        
        # Disclaimer
        next_update = "4:30 PM" if now.hour < 12 else "9:30 AM (Next Day)"
        html += f"""
            <div class="disclaimer">
                <h3>⚠️ DISCLAIMER</h3>
                <p>This analysis is for <strong>EDUCATIONAL PURPOSES ONLY</strong>. This is NOT financial advice.</p>
                <ul>
                    <li>Do your own research</li>
                    <li>Consult a SEBI registered financial advisor</li>
                    <li>Use proper risk management and stop losses</li>
                    <li>Never invest more than you can afford to lose</li>
                </ul>
            </div>
        </div>
        
        <!-- Footer -->
        <div class="footer">
            <p><strong>© 2025 NIFTY 50 Analyzer - Royal Sapphire Edition</strong></p>
            <p>Premium Elite Trading Analysis | Next Update: {next_update} IST</p>
        </div>
    </div>
</body>
</html>
"""
        
        # Write to file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)
        
        print(f"✅ GitHub Pages HTML generated: {output_file}\n")
        return output_file
    
    def generate_email_html(self):
        """Generate beautiful HTML email - ROYAL SAPPHIRE THEME"""
        df = pd.DataFrame(self.results)
        top_buys, top_sells = self.get_top_recommendations()
        
        # Get IST time
        now = self.get_ist_time()
        time_of_day = "Morning" if now.hour < 12 else "Evening"
        
        # Count recommendations
        strong_buy_count = len(df[df['Recommendation'] == 'STRONG BUY'])
        buy_count = len(df[df['Recommendation'] == 'BUY'])
        hold_count = len(df[df['Recommendation'] == 'HOLD'])
        sell_count = len(df[df['Recommendation'] == 'SELL'])
        strong_sell_count = len(df[df['Recommendation'] == 'STRONG SELL'])
        
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>
<body bgcolor="#0a0e27" style="margin:0; padding:0; font-family: Georgia, serif;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#0a0e27">
        <tr>
            <td align="center" style="padding: 20px;">
                <table width="900" cellpadding="0" cellspacing="0" border="0" bgcolor="#1a237e" style="border-radius: 20px; border: 3px solid #3f51b5;">
                    <!-- Header -->
                    <tr>
                        <td bgcolor="#3f51b5" align="center" style="padding: 30px; background: linear-gradient(135deg, #3f51b5 0%, #5c6bc0 100%);">
                            <h1 style="color: #ffffff; margin: 0; font-size: 32px;">💎 NIFTY 50 Stock Analysis</h1>
                            <p style="color: #ffffff; margin: 10px 0 0 0; font-size: 16px;">{time_of_day} Update - Royal Sapphire Edition - {now.strftime('%d %b %Y, %I:%M %p')} IST</p>
                        </td>
                    </tr>
                    
                    <!-- Content -->
                    <tr>
                        <td bgcolor="#0f1535" style="padding: 30px;">
                            
                            <!-- Summary Box -->
                            <table width="100%" cellpadding="15" cellspacing="0" border="0" bgcolor="#283593" style="border-radius: 10px; margin-bottom: 30px; border: 2px solid #3f51b5;">
                                <tr>
                                    <td>
                                        <h2 style="color: #9fa8da; margin: 0 0 15px 0; font-size: 20px;">📈 Market Summary</h2>
                                        <table width="100%" cellpadding="10" cellspacing="10" border="0">
                                            <tr>
                                                <td width="25%" bgcolor="#1a237e" align="center" style="border-radius: 8px; border: 2px solid #3f51b5;">
                                                    <strong style="color: #9fa8da; font-size: 32px; display: block;">{len(self.results)}</strong>
                                                    <span style="color: #c5cae9; font-size: 13px;">STOCKS ANALYZED</span>
                                                </td>
                                                <td width="25%" bgcolor="#1a237e" align="center" style="border-radius: 8px; border: 2px solid #3f51b5;">
                                                    <strong style="color: #9fa8da; font-size: 32px; display: block;">{strong_buy_count}</strong>
                                                    <span style="color: #c5cae9; font-size: 13px;">STRONG BUY</span>
                                                </td>
                                                <td width="25%" bgcolor="#1a237e" align="center" style="border-radius: 8px; border: 2px solid #3f51b5;">
                                                    <strong style="color: #9fa8da; font-size: 32px; display: block;">{buy_count}</strong>
                                                    <span style="color: #c5cae9; font-size: 13px;">BUY</span>
                                                </td>
                                                <td width="25%" bgcolor="#1a237e" align="center" style="border-radius: 8px; border: 2px solid #3f51b5;">
                                                    <strong style="color: #9fa8da; font-size: 32px; display: block;">{hold_count}</strong>
                                                    <span style="color: #c5cae9; font-size: 13px;">HOLD</span>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
"""
        
        # Top 10 Buy Recommendations
        if not top_buys.empty:
            html += """
                            <!-- BUY Section -->
                            <h2 style="color: #7986cb; border-bottom: 3px solid #3f51b5; padding-bottom: 10px; margin-top: 40px;">🟢 TOP 10 BUY RECOMMENDATIONS</h2>
                            <table width="100%" cellpadding="12" cellspacing="0" border="1" bordercolor="#283593" style="border-collapse: collapse; margin: 20px 0; background: #1a237e;">
                                <tr bgcolor="#3f51b5">
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">STOCK</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">PRICE</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">RATING</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">SCORE</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">UPSIDE %</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">TARGET</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">STOP LOSS</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">QUALITY</th>
                                </tr>
"""
            row_num = 0
            for idx, row in top_buys.iterrows():
                row_num += 1
                row_bg = "#1a237e" if row_num % 2 == 1 else "#283593"
                
                # Upside color
                if row['Upside'] > 0:
                    upside_color = "#66bb6a"
                elif row['Upside'] < 0:
                    upside_color = "#ef5350"
                else:
                    upside_color = "#e8eaf6"
                
                # Quality badge color
                if row['Quality'] == 'Excellent':
                    badge_color = "#5c6bc0"
                elif row['Quality'] == 'Good':
                    badge_color = "#7986cb"
                elif row['Quality'] == 'Average':
                    badge_color = "#9fa8da"
                else:
                    badge_color = "#ef5350"
                
                html += f"""
                                <tr bgcolor="{row_bg}">
                                    <td style="color: #9fa8da; font-weight: 600; padding: 14px 12px; border: 1px solid #283593;">{row['Name']}</td>
                                    <td style="color: #e8eaf6; padding: 14px 12px; border: 1px solid #283593;">₹{row['Price']:,.0f}</td>
                                    <td style="color: #e8eaf6; padding: 14px 12px; border: 1px solid #283593; font-size: 12px; font-weight: bold;">{row['Rating']}</td>
                                    <td style="color: #e8eaf6; font-weight: bold; padding: 14px 12px; border: 1px solid #283593;">{row['Combined_Score']:.0f}</td>
                                    <td style="color: {upside_color}; font-weight: bold; padding: 14px 12px; border: 1px solid #283593; font-size: 16px;">{row['Upside']:+.1f}%</td>
                                    <td style="color: #e8eaf6; padding: 14px 12px; border: 1px solid #283593;">₹{row['Target_1']:,.0f}</td>
                                    <td style="color: #e8eaf6; padding: 14px 12px; border: 1px solid #283593;">₹{row['Stop_Loss']:,.0f}</td>
                                    <td style="padding: 14px 12px; border: 1px solid #283593;"><span style="background-color: {badge_color}; color: #ffffff; padding: 5px 10px; border-radius: 5px; font-size: 11px; font-weight: bold;">{row['Quality']}</span></td>
                                </tr>
"""
            html += """
                            </table>
"""
        
        # Top 10 Sell Recommendations
        if not top_sells.empty:
            html += """
                            <!-- SELL Section -->
                            <h2 style="color: #ef5350; border-bottom: 3px solid #f44336; padding-bottom: 10px; margin-top: 40px;">🔴 TOP 10 SELL RECOMMENDATIONS</h2>
                            <table width="100%" cellpadding="12" cellspacing="0" border="1" bordercolor="#283593" style="border-collapse: collapse; margin: 20px 0; background: #1a237e;">
                                <tr bgcolor="#f44336">
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">STOCK</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">PRICE</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">RATING</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">SCORE</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">RSI</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">MACD</th>
                                    <th style="color: #ffffff; text-align: left; padding: 16px 12px; font-size: 13px;">QUALITY</th>
                                </tr>
"""
            row_num = 0
            for idx, row in top_sells.iterrows():
                row_num += 1
                row_bg = "#1a237e" if row_num % 2 == 1 else "#283593"
                
                # RSI color
                if row['RSI'] > 70:
                    rsi_color = "#ef5350"
                elif row['RSI'] < 30:
                    rsi_color = "#66bb6a"
                else:
                    rsi_color = "#ffa726"
                
                # Quality badge color
                if row['Quality'] == 'Excellent':
                    badge_color = "#5c6bc0"
                elif row['Quality'] == 'Good':
                    badge_color = "#7986cb"
                elif row['Quality'] == 'Average':
                    badge_color = "#9fa8da"
                else:
                    badge_color = "#ef5350"
                
                html += f"""
                                <tr bgcolor="{row_bg}">
                                    <td style="color: #9fa8da; font-weight: 600; padding: 14px 12px; border: 1px solid #283593;">{row['Name']}</td>
                                    <td style="color: #e8eaf6; padding: 14px 12px; border: 1px solid #283593;">₹{row['Price']:,.0f}</td>
                                    <td style="color: #e8eaf6; padding: 14px 12px; border: 1px solid #283593; font-size: 12px; font-weight: bold;">{row['Rating']}</td>
                                    <td style="color: #e8eaf6; font-weight: bold; padding: 14px 12px; border: 1px solid #283593;">{row['Combined_Score']:.0f}</td>
                                    <td style="color: {rsi_color}; font-weight: bold; padding: 14px 12px; border: 1px solid #283593; font-size: 16px;">{row['RSI']:.0f}</td>
                                    <td style="color: #e8eaf6; padding: 14px 12px; border: 1px solid #283593;">{row['MACD']}</td>
                                    <td style="padding: 14px 12px; border: 1px solid #283593;"><span style="background-color: {badge_color}; color: #ffffff; padding: 5px 10px; border-radius: 5px; font-size: 11px; font-weight: bold;">{row['Quality']}</span></td>
                                </tr>
"""
            html += """
                            </table>
"""
        
        # Disclaimer and Footer
        next_update = "4:30 PM" if now.hour < 12 else "9:30 AM (Next Day)"
        html += f"""
                            <!-- Disclaimer -->
                            <table width="100%" cellpadding="20" cellspacing="0" border="2" bordercolor="#ffa726" bgcolor="#1a237e" style="margin: 30px 0;">
                                <tr>
                                    <td>
                                        <p style="color: #e8eaf6; margin: 0 0 10px 0;"><strong style="color: #ef5350;">⚠️ DISCLAIMER:</strong> This analysis is for <strong>EDUCATIONAL PURPOSES ONLY</strong>. This is NOT financial advice. Always:</p>
                                        <ul style="color: #e8eaf6; margin: 10px 0; padding-left: 20px;">
                                            <li>Do your own research</li>
                                            <li>Consult a SEBI registered financial advisor</li>
                                            <li>Use proper risk management and stop losses</li>
                                            <li>Never invest more than you can afford to lose</li>
                                        </ul>
                                    </td>
                                </tr>
                            </table>
                            
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td bgcolor="#0a0e27" align="center" style="padding: 25px;">
                            <p style="color: #9fa8da; margin: 0 0 5px 0; font-size: 13px;"><strong>© 2025 NIFTY 50 Analyzer - Royal Sapphire Edition</strong></p>
                            <p style="color: #7986cb; margin: 0; font-size: 13px;">Premium Elite Trading Analysis | Next Update: {next_update} IST</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"""
        
        return html
    
    def send_email(self, to_email):
        """Send email with analysis report"""
        try:
            # Get credentials from environment variables
            from_email = os.environ.get('GMAIL_USER')
            password = os.environ.get('GMAIL_APP_PASSWORD')
            
            if not from_email or not password:
                print("❌ Gmail credentials not found in environment variables")
                print("   Set GMAIL_USER and GMAIL_APP_PASSWORD")
                return False
            
            # Get IST time
            now = self.get_ist_time()
            time_of_day = "Morning" if now.hour < 12 else "Evening"
            
            # Create message
            msg = MIMEMultipart('alternative')
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = f"💎 NIFTY 50 Stock Analysis Report - {time_of_day} Report ({now.strftime('%d %b %Y')})"
            
            # Generate email body
            html_body = self.generate_email_html()
            msg.attach(MIMEText(html_body, 'html'))
            
            # Send email
            print(f"📧 Sending email to {to_email}...")
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
            server.quit()
            
            print(f"✅ Email sent successfully!\n")
            return True
            
        except Exception as e:
            print(f"❌ Error sending email: {e}\n")
            return False
    
    def generate_complete_report(self, send_email_flag=True, recipient_email=None, generate_github_pages=True):
        """Generate complete analysis report"""
        ist_time = self.get_ist_time()
        
        print("=" * 70)
        print("💎 NIFTY 50 STOCK ANALYZER - ROYAL SAPPHIRE EDITION")
        print(f"Started: {ist_time.strftime('%d %b %Y, %I:%M %p IST')}")
        print("=" * 70)
        print()
        
        # Analyze all stocks
        self.analyze_all_stocks()
        
        # Generate GitHub Pages HTML
        if generate_github_pages:
            self.generate_github_pages_html('index.html')
        
        # Send email if requested
        if send_email_flag and recipient_email:
            self.send_email(recipient_email)
        
        print("=" * 70)
        print("✅ ANALYSIS COMPLETE!")
        print("=" * 70)


def main():
    """Main execution"""
    analyzer = Nifty50CompleteAnalyzer()
    
    # Get recipient email from environment variable
    recipient = os.environ.get('RECIPIENT_EMAIL')
    
    if not recipient:
        print("⚠️  RECIPIENT_EMAIL environment variable not set")
        print("   Please set it to receive email reports")
        recipient = None
    
    # Generate report, GitHub Pages HTML, and send email
    analyzer.generate_complete_report(
        send_email_flag=True, 
        recipient_email=recipient,
        generate_github_pages=True
    )


if __name__ == "__main__":
    main()
