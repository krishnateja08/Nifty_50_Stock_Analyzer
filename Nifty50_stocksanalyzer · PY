"""
NIFTY 50 COMPLETE STOCK ANALYZER
Technical + Fundamental Analysis with Email Delivery

Requirements:
pip install yfinance pandas numpy tabulate openpyxl
"""

import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
import warnings
from tabulate import tabulate
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
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
        self.email_body = []
    
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
            
            # Rating
            if combined_score >= 90:
                rating = "⭐⭐⭐⭐⭐ STRONG BUY"
                recommendation = "STRONG BUY"
            elif combined_score >= 70:
                rating = "⭐⭐⭐⭐ BUY"
                recommendation = "BUY"
            elif combined_score >= 50:
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
        print(f"Analyzing {len(self.nifty50_stocks)} stocks...")
        
        for symbol, name in self.nifty50_stocks.items():
            result = self.analyze_stock(symbol, name)
            if result:
                self.results.append(result)
        
        print(f"✅ Analysis complete: {len(self.results)} stocks analyzed")
    
    def get_top_recommendations(self):
        """Get top 10 buy and sell recommendations"""
        df = pd.DataFrame(self.results)
        
        # Top 10 Buy (highest combined scores)
        top_buys = df[df['Recommendation'].isin(['STRONG BUY', 'BUY'])].nlargest(10, 'Combined_Score')
        
        # Top 10 Sell (lowest combined scores)
        top_sells = df[df['Recommendation'].isin(['STRONG SELL', 'SELL'])].nsmallest(10, 'Combined_Score')
        
        return top_buys, top_sells
    
    def create_html_table(self, df, title, columns):
        """Create HTML table for email"""
        html = f"<h2 style='color: #2c3e50;'>{title}</h2>\n"
        html += "<table style='border-collapse: collapse; width: 100%; margin-bottom: 30px;'>\n"
        html += "<tr style='background-color: #3498db; color: white;'>\n"
        
        for col in columns:
            html += f"<th style='border: 1px solid #ddd; padding: 12px; text-align: left;'>{col}</th>\n"
        html += "</tr>\n"
        
        for idx, row in df.iterrows():
            bg_color = "#f2f2f2" if idx % 2 == 0 else "white"
            html += f"<tr style='background-color: {bg_color};'>\n"
            
            for col in columns:
                value = row.get(col, '')
                html += f"<td style='border: 1px solid #ddd; padding: 10px;'>{value}</td>\n"
            html += "</tr>\n"
        
        html += "</table>\n"
        return html
    
    def generate_email_body(self):
        """Generate HTML email body"""
        df = pd.DataFrame(self.results)
        top_buys, top_sells = self.get_top_recommendations()
        
        now = datetime.now()
        time_of_day = "Morning" if now.hour < 12 else "Evening"
        
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                .header {{ background-color: #2c3e50; color: white; padding: 20px; text-align: center; }}
                .content {{ padding: 20px; }}
                .footer {{ background-color: #ecf0f1; padding: 15px; text-align: center; font-size: 12px; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>📊 NIFTY 50 Stock Analysis Report</h1>
                <p>{time_of_day} Update - {now.strftime('%d %b %Y, %I:%M %p IST')}</p>
            </div>
            
            <div class="content">
                <h2 style='color: #27ae60;'>📈 Summary</h2>
                <p><strong>Total Stocks Analyzed:</strong> {len(self.results)}</p>
                <p><strong>Strong Buy Opportunities:</strong> {len(df[df['Recommendation'] == 'STRONG BUY'])}</p>
                <p><strong>Buy Opportunities:</strong> {len(df[df['Recommendation'] == 'BUY'])}</p>
                <p><strong>Hold Recommendations:</strong> {len(df[df['Recommendation'] == 'HOLD'])}</p>
        """
        
        # Top 10 Buy Recommendations
        if not top_buys.empty:
            buy_data = []
            for idx, row in top_buys.iterrows():
                buy_data.append({
                    'Stock': row['Name'],
                    'Price': f"₹{row['Price']:,.0f}",
                    'Rating': row['Rating'],
                    'Score': f"{row['Combined_Score']:.0f}",
                    'Upside': f"{row['Upside']:+.1f}%",
                    'Target': f"₹{row['Target_1']:,.0f}",
                    'Stop Loss': f"₹{row['Stop_Loss']:,.0f}"
                })
            buy_df = pd.DataFrame(buy_data)
            html += self.create_html_table(buy_df, "🟢 TOP 10 BUY RECOMMENDATIONS", 
                                          ['Stock', 'Price', 'Rating', 'Score', 'Upside', 'Target', 'Stop Loss'])
        
        # Top 10 Sell Recommendations
        if not top_sells.empty:
            sell_data = []
            for idx, row in top_sells.iterrows():
                sell_data.append({
                    'Stock': row['Name'],
                    'Price': f"₹{row['Price']:,.0f}",
                    'Rating': row['Rating'],
                    'Score': f"{row['Combined_Score']:.0f}",
                    'RSI': f"{row['RSI']:.0f}",
                    'MACD': row['MACD']
                })
            sell_df = pd.DataFrame(sell_data)
            html += self.create_html_table(sell_df, "🔴 TOP 10 SELL RECOMMENDATIONS", 
                                          ['Stock', 'Price', 'Rating', 'Score', 'RSI', 'MACD'])
        
        html += """
                <div style='background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0;'>
                    <p><strong>⚠️ DISCLAIMER:</strong> This analysis is for EDUCATIONAL purposes only. NOT financial advice. 
                    Always do your own research and consult a SEBI registered advisor before investing.</p>
                </div>
            </div>
            
            <div class="footer">
                <p>© 2025 NIFTY 50 Analyzer | Automated Stock Analysis System</p>
            </div>
        </body>
        </html>
        """
        
        return html
    
    def export_to_excel(self, filename="Nifty50_Analysis.xlsx"):
        """Export all results to Excel"""
        try:
            df = pd.DataFrame(self.results)
            df_sorted = df.sort_values('Combined_Score', ascending=False)
            
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # All stocks
                df_sorted.to_excel(writer, sheet_name='All Stocks', index=False)
                
                # Top Buys
                top_buys = df[df['Recommendation'].isin(['STRONG BUY', 'BUY'])].nlargest(10, 'Combined_Score')
                top_buys.to_excel(writer, sheet_name='Top 10 Buys', index=False)
                
                # Top Sells
                top_sells = df[df['Recommendation'].isin(['STRONG SELL', 'SELL'])].nsmallest(10, 'Combined_Score')
                top_sells.to_excel(writer, sheet_name='Top 10 Sells', index=False)
            
            print(f"✅ Excel report created: {filename}")
            return filename
            
        except Exception as e:
            print(f"❌ Error exporting to Excel: {e}")
            return None
    
    def send_email(self, to_email):
        """Send email with analysis report"""
        try:
            # Get credentials from environment variables
            from_email = os.environ.get('GMAIL_USER')
            password = os.environ.get('GMAIL_APP_PASSWORD')
            
            if not from_email or not password:
                print("❌ Gmail credentials not found in environment variables")
                return False
            
            now = datetime.now()
            time_of_day = "Morning" if now.hour < 12 else "Evening"
            
            # Create message
            msg = MIMEMultipart('alternative')
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = f"📊 NIFTY 50 Analysis - {time_of_day} Report ({now.strftime('%d %b %Y')})"
            
            # Generate email body
            html_body = self.generate_email_body()
            msg.attach(MIMEText(html_body, 'html'))
            
            # Attach Excel file
            excel_file = self.export_to_excel()
            if excel_file and os.path.exists(excel_file):
                with open(excel_file, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={excel_file}')
                    msg.attach(part)
            
            # Send email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
            server.quit()
            
            print(f"✅ Email sent successfully to {to_email}")
            return True
            
        except Exception as e:
            print(f"❌ Error sending email: {e}")
            return False
    
    def generate_complete_report(self, send_email_flag=True, recipient_email=None):
        """Generate complete analysis report"""
        print("="*60)
        print("📊 NIFTY 50 STOCK ANALYZER")
        print(f"Started: {datetime.now().strftime('%d-%b-%Y %H:%M IST')}")
        print("="*60)
        
        # Analyze all stocks
        self.analyze_all_stocks()
        
        # Send email if requested
        if send_email_flag and recipient_email:
            self.send_email(recipient_email)
        
        print("="*60)
        print("✅ Analysis Complete!")
        print("="*60)


def main():
    """Main execution"""
    analyzer = Nifty50CompleteAnalyzer()
    
    # Get recipient email from environment variable
    recipient = os.environ.get('RECIPIENT_EMAIL', 'your-email@gmail.com')
    
    # Generate report and send email
    analyzer.generate_complete_report(send_email_flag=True, recipient_email=recipient)


if __name__ == "__main__":
    main()
