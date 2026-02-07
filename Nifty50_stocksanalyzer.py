"""
NIFTY 50 COMPLETE STOCK ANALYZER
Technical + Fundamental Analysis with Email Delivery

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
    
    def generate_email_html(self):
        """Generate beautiful HTML email with BLACK background"""
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
    <style>
        body {{
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #e0e0e0;
            background-color: #000000 !important;
            margin: 0;
            padding: 0;
        }}
        .email-container {{
            max-width: 1000px;
            margin: 20px auto;
            background-color: #1a1a1a !important;
            box-shadow: 0 0 20px rgba(255,255,255,0.1);
            border-radius: 10px;
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 30px 20px;
            text-align: center;
        }}
        .header h1 {{
            margin: 0 0 10px 0;
            font-size: 32px;
            font-weight: 700;
        }}
        .header p {{
            margin: 0;
            font-size: 16px;
            opacity: 0.95;
        }}
        .content {{
            padding: 30px 20px;
            background-color: #1a1a1a !important;
        }}
        .summary-box {{
            background: linear-gradient(135deg, #4a90e2 0%, #357abd 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            margin-bottom: 30px;
        }}
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            margin-top: 15px;
        }}
        .summary-item {{
            background: rgba(255,255,255,0.15);
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.2);
        }}
        .summary-item strong {{
            display: block;
            font-size: 32px;
            margin-bottom: 8px;
            font-weight: 700;
        }}
        .summary-item span {{
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        h2 {{
            color: #4a90e2;
            border-bottom: 3px solid #4a90e2;
            padding-bottom: 10px;
            margin-top: 40px;
            font-size: 24px;
        }}
        .buy-section h2 {{
            color: #10b981;
            border-bottom: 3px solid #10b981;
        }}
        .sell-section h2 {{
            color: #ef4444;
            border-bottom: 3px solid #ef4444;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
            background-color: #2d2d2d !important;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        }}
        th {{
            background: #3a3a3a !important;
            color: #ffffff;
            padding: 16px 12px;
            text-align: left;
            font-weight: 600;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .buy-section th {{
            background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
        }}
        .sell-section th {{
            background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%) !important;
        }}
        td {{
            border-bottom: 1px solid #404040;
            padding: 14px 12px;
            font-size: 13px;
            color: #e0e0e0;
            background-color: #2d2d2d !important;
        }}
        tr:last-child td {{
            border-bottom: none;
        }}
        tr:hover td {{
            background-color: #353535 !important;
        }}
        .stock-name {{
            font-weight: 600;
            color: #ffffff;
        }}
        .rating {{
            font-weight: bold;
            font-size: 12px;
        }}
        .positive {{
            color: #10b981;
            font-weight: bold;
        }}
        .negative {{
            color: #ef4444;
            font-weight: bold;
        }}
        .neutral {{
            color: #f59e0b;
        }}
        .disclaimer {{
            background: linear-gradient(135deg, #78350f 0%, #92400e 100%);
            border-left: 4px solid #f59e0b;
            padding: 20px;
            margin: 30px 0;
            border-radius: 8px;
            color: #fef3c7;
        }}
        .disclaimer strong {{
            color: #fbbf24;
        }}
        .footer {{
            background-color: #0d0d0d !important;
            color: #9ca3af;
            padding: 25px;
            text-align: center;
            font-size: 13px;
        }}
        .footer strong {{
            color: #e0e0e0;
        }}
        .badge {{
            display: inline-block;
            padding: 5px 10px;
            border-radius: 5px;
            font-size: 11px;
            font-weight: bold;
        }}
        .badge-excellent {{
            background-color: #10b981;
            color: white;
        }}
        .badge-good {{
            background-color: #3b82f6;
            color: white;
        }}
        .badge-average {{
            background-color: #f59e0b;
            color: white;
        }}
        .badge-poor {{
            background-color: #ef4444;
            color: white;
        }}
    </style>
</head>
<body style="background-color: #000000 !important;">
    <div class="email-container">
        <div class="header">
            <h1>📊 NIFTY 50 Stock Analysis Report</h1>
            <p>{time_of_day} Update - {now.strftime('%d %b %Y, %I:%M %p')} IST</p>
        </div>
        
        <div class="content">
            <div class="summary-box">
                <h2 style="margin: 0 0 15px 0; color: white; border: none; font-size: 20px;">📈 Market Summary</h2>
                <div class="summary-grid">
                    <div class="summary-item">
                        <strong>{len(self.results)}</strong>
                        <span>Stocks Analyzed</span>
                    </div>
                    <div class="summary-item">
                        <strong>{strong_buy_count}</strong>
                        <span>Strong Buy</span>
                    </div>
                    <div class="summary-item">
                        <strong>{buy_count}</strong>
                        <span>Buy</span>
                    </div>
                    <div class="summary-item">
                        <strong>{hold_count}</strong>
                        <span>Hold</span>
                    </div>
                </div>
            </div>
"""
        
        # Top 10 Buy Recommendations
        if not top_buys.empty:
            html += """
            <div class="buy-section">
                <h2>🟢 TOP 10 BUY RECOMMENDATIONS</h2>
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
                quality_badge = f"badge-{row['Quality'].lower()}"
                upside_color = "positive" if row['Upside'] > 0 else "negative"
                html += f"""
                        <tr>
                            <td class="stock-name">{row['Name']}</td>
                            <td>₹{row['Price']:,.0f}</td>
                            <td class="rating">{row['Rating']}</td>
                            <td><strong>{row['Combined_Score']:.0f}</strong></td>
                            <td class="{upside_color}">{row['Upside']:+.1f}%</td>
                            <td>₹{row['Target_1']:,.0f}</td>
                            <td>₹{row['Stop_Loss']:,.0f}</td>
                            <td><span class="badge {quality_badge}">{row['Quality']}</span></td>
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
            <div class="sell-section">
                <h2>🔴 TOP 10 SELL RECOMMENDATIONS</h2>
                <table>
                    <thead>
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
                quality_badge = f"badge-{row['Quality'].lower()}"
                rsi_color = "negative" if row['RSI'] > 70 else ("positive" if row['RSI'] < 30 else "neutral")
                html += f"""
                        <tr>
                            <td class="stock-name">{row['Name']}</td>
                            <td>₹{row['Price']:,.0f}</td>
                            <td class="rating">{row['Rating']}</td>
                            <td><strong>{row['Combined_Score']:.0f}</strong></td>
                            <td class="{rsi_color}">{row['RSI']:.0f}</td>
                            <td>{row['MACD']}</td>
                            <td><span class="badge {quality_badge}">{row['Quality']}</span></td>
                        </tr>
"""
            html += """
                    </tbody>
                </table>
            </div>
"""
        
        # Disclaimer and Footer
        next_update = "4:30 PM" if now.hour < 12 else "9:30 AM (Next Day)"
        html += f"""
            <div class="disclaimer">
                <p><strong>⚠️ DISCLAIMER:</strong> This analysis is for <strong>EDUCATIONAL PURPOSES ONLY</strong>. This is NOT financial advice. Always:</p>
                <ul style="margin: 10px 0;">
                    <li>Do your own research</li>
                    <li>Consult a SEBI registered financial advisor</li>
                    <li>Use proper risk management and stop losses</li>
                    <li>Never invest more than you can afford to lose</li>
                </ul>
            </div>
        </div>
        
        <div class="footer">
            <p style="margin: 0 0 5px 0;"><strong>© 2025 NIFTY 50 Analyzer</strong></p>
            <p style="margin: 0;">Automated Stock Analysis System | Next Update: {next_update} IST</p>
        </div>
    </div>
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
            msg['Subject'] = f"📊 NIFTY 50 Analysis - {time_of_day} Report ({now.strftime('%d %b %Y')})"
            
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
    
    def generate_complete_report(self, send_email_flag=True, recipient_email=None):
        """Generate complete analysis report"""
        ist_time = self.get_ist_time()
        
        print("=" * 70)
        print("📊 NIFTY 50 STOCK ANALYZER")
        print(f"Started: {ist_time.strftime('%d %b %Y, %I:%M %p IST')}")
        print("=" * 70)
        print()
        
        # Analyze all stocks
        self.analyze_all_stocks()
        
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
    
    # Generate report and send email
    analyzer.generate_complete_report(send_email_flag=True, recipient_email=recipient)


if __name__ == "__main__":
    main()
