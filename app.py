"""
Amazon Marketing ROI Analyzer - Enhanced Version
Includes: Historical tracking, trend analysis, keyword performance, product recommendations
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from pathlib import Path
import io
import json
import os

# Page configuration
st.set_page_config(
    page_title="Amazon Marketing ROI Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #FF9900;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #FF9900;
    }
    .recommendation-box {
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .expand {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
    }
    .optimize {
        background-color: #cfe2ff;
        border-left: 4px solid #0d6efd;
    }
    .pause {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
    }
    .close {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
    }
    .trending-up {
        color: #28a745;
        font-weight: bold;
    }
    .trending-down {
        color: #dc3545;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# Historical data storage directory
HISTORY_DIR = Path("marketing_history")
HISTORY_DIR.mkdir(exist_ok=True)

class HistoricalDataManager:
    """Manage historical data storage and retrieval"""
    
    def __init__(self):
        self.history_file = HISTORY_DIR / "campaign_history.csv"
        self.keyword_history_file = HISTORY_DIR / "keyword_history.csv"
        self.product_history_file = HISTORY_DIR / "product_history.csv"
    
    def save_campaign_snapshot(self, df, snapshot_date=None):
        """Save campaign performance snapshot"""
        if df is None or len(df) == 0:
            return
        
        if snapshot_date is None:
            snapshot_date = datetime.now().strftime('%Y-%m-%d')
        
        df_snapshot = df.copy()
        df_snapshot['snapshot_date'] = snapshot_date
        
        # Append to history
        if self.history_file.exists():
            history = pd.read_csv(self.history_file)
            # Remove duplicates for same date/campaign
            history = history[history['snapshot_date'] != snapshot_date]
            df_snapshot = pd.concat([history, df_snapshot], ignore_index=True)
        
        df_snapshot.to_csv(self.history_file, index=False)
    
    def save_keyword_snapshot(self, df, snapshot_date=None):
        """Save keyword performance snapshot"""
        if df is None or len(df) == 0:
            return
        
        if snapshot_date is None:
            snapshot_date = datetime.now().strftime('%Y-%m-%d')
        
        df_snapshot = df.copy()
        df_snapshot['snapshot_date'] = snapshot_date
        
        if self.keyword_history_file.exists():
            history = pd.read_csv(self.keyword_history_file)
            history = history[history['snapshot_date'] != snapshot_date]
            df_snapshot = pd.concat([history, df_snapshot], ignore_index=True)
        
        df_snapshot.to_csv(self.keyword_history_file, index=False)
    
    def save_product_snapshot(self, df, snapshot_date=None):
        """Save product performance snapshot"""
        if df is None or len(df) == 0:
            return
        
        if snapshot_date is None:
            snapshot_date = datetime.now().strftime('%Y-%m-%d')
        
        df_snapshot = df.copy()
        df_snapshot['snapshot_date'] = snapshot_date
        
        if self.product_history_file.exists():
            history = pd.read_csv(self.product_history_file)
            history = history[history['snapshot_date'] != snapshot_date]
            df_snapshot = pd.concat([history, df_snapshot], ignore_index=True)
        
        df_snapshot.to_csv(self.product_history_file, index=False)
    
    def get_campaign_history(self, days=30):
        """Get campaign history for last N days"""
        if not self.history_file.exists():
            return None
        
        history = pd.read_csv(self.history_file)
        history['snapshot_date'] = pd.to_datetime(history['snapshot_date'])
        
        cutoff_date = datetime.now() - timedelta(days=days)
        return history[history['snapshot_date'] >= cutoff_date]
    
    def get_keyword_history(self, days=30):
        """Get keyword history for last N days"""
        if not self.keyword_history_file.exists():
            return None
        
        history = pd.read_csv(self.keyword_history_file)
        history['snapshot_date'] = pd.to_datetime(history['snapshot_date'])
        
        cutoff_date = datetime.now() - timedelta(days=days)
        return history[history['snapshot_date'] >= cutoff_date]
    
    def get_product_history(self, days=30):
        """Get product history for last N days"""
        if not self.product_history_file.exists():
            return None
        
        history = pd.read_csv(self.product_history_file)
        history['snapshot_date'] = pd.to_datetime(history['snapshot_date'])
        
        cutoff_date = datetime.now() - timedelta(days=days)
        return history[history['snapshot_date'] >= cutoff_date]


class AmazonROIAnalyzer:
    """Main analyzer class for processing Amazon marketing data"""
    
    def __init__(self):
        self.weekly_sales = None
        self.weekly_campaigns = None
        self.weekly_cpr = None
        self.weekly_repeat = None
        self.daily_campaigns = None
        self.daily_targets = None
        self.daily_inventory = None
        self.combined_data = None
        self.history_manager = HistoricalDataManager()
        
    def parse_currency(self, value):
        """Parse currency string to float"""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        cleaned = str(value).replace('₹', '').replace(',', '').replace('$', '').strip()
        try:
            return float(cleaned)
        except:
            return 0.0
    
    def parse_percentage(self, value):
        """Parse percentage string to float"""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        cleaned = str(value).replace('%', '').strip()
        try:
            return float(cleaned)
        except:
            return 0.0
    
    def load_weekly_sales(self, file):
        """Load weekly sales report"""
        try:
            xl = pd.ExcelFile(file)
            sheets_data = []
            
            for sheet in xl.sheet_names:
                if sheet not in ['WeeklyOnePager', 'Summary']:
                    df = pd.read_excel(file, sheet_name=sheet)
                    df['Week'] = sheet
                    sheets_data.append(df)
            
            if sheets_data:
                self.weekly_sales = pd.concat(sheets_data, ignore_index=True)
                st.success(f"✅ Loaded Weekly Sales: {len(sheets_data)} weeks")
                return True
        except Exception as e:
            st.error(f"Error loading weekly sales: {str(e)}")
            return False
    
    def load_weekly_campaigns(self, file):
        """Load weekly campaign report"""
        try:
            xl = pd.ExcelFile(file)
            sheets_data = []
            
            for sheet in xl.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet)
                df['Week'] = sheet
                sheets_data.append(df)
            
            if sheets_data:
                self.weekly_campaigns = pd.concat(sheets_data, ignore_index=True)
                numeric_cols = ['Impressions', 'Clicks', 'CTR', 'Total cost', 'CPC', 
                               'Purchases', 'Sales', 'ACOS', 'ROAS', 'CVR']
                for col in numeric_cols:
                    if col in self.weekly_campaigns.columns:
                        self.weekly_campaigns[col] = self.weekly_campaigns[col].apply(self.parse_currency)
                
                st.success(f"✅ Loaded Weekly Campaigns: {len(self.weekly_campaigns)} campaigns")
                return True
        except Exception as e:
            st.error(f"Error loading weekly campaigns: {str(e)}")
            return False
    
    def load_weekly_cpr(self, file):
        """Load weekly CPR report"""
        try:
            xl = pd.ExcelFile(file)
            sheets_data = []
            
            for sheet in xl.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet)
                df['Week'] = sheet
                sheets_data.append(df)
            
            if sheets_data:
                self.weekly_cpr = pd.concat(sheets_data, ignore_index=True)
                st.success(f"✅ Loaded Weekly CPR: {len(self.weekly_cpr)} entries")
                
                # Save keyword history
                if 'Keyword' in self.weekly_cpr.columns:
                    self.history_manager.save_keyword_snapshot(self.weekly_cpr)
                
                return True
        except Exception as e:
            st.error(f"Error loading weekly CPR: {str(e)}")
            return False
    
    def load_weekly_repeat(self, file):
        """Load weekly repeat purchase report"""
        try:
            xl = pd.ExcelFile(file)
            sheets_data = []
            
            for sheet in xl.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet)
                df['Week'] = sheet
                sheets_data.append(df)
            
            if sheets_data:
                self.weekly_repeat = pd.concat(sheets_data, ignore_index=True)
                st.success(f"✅ Loaded Weekly Repeat Purchase: {len(self.weekly_repeat)} products")
                
                # Save product history
                self.history_manager.save_product_snapshot(self.weekly_repeat)
                
                return True
        except Exception as e:
            st.error(f"Error loading weekly repeat purchase: {str(e)}")
            return False
    
    def load_daily_campaigns(self, file):
        """Load daily campaign report"""
        try:
            df = pd.read_csv(file)
            
            if 'Campaign start date' in df.columns:
                df['Campaign start date'] = pd.to_datetime(df['Campaign start date'], errors='coerce')
            
            numeric_cols = ['Clicks', 'CTR', 'Total cost (converted)', 'Total cost', 'CPC (converted)', 
                           'Purchases', 'Sales (converted)', 'Sales', 'ROAS']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(self.parse_currency)
            
            if self.daily_campaigns is not None:
                df = pd.concat([self.daily_campaigns, df])
                df = df.drop_duplicates(subset=['Campaign name'], keep='last')
            
            self.daily_campaigns = df
            
            # Save to history
            self.history_manager.save_campaign_snapshot(df)
            
            st.success(f"✅ Loaded Daily Campaigns: {len(df)} campaigns")
            return True
        except Exception as e:
            st.error(f"Error loading daily campaigns: {str(e)}")
            return False
    
    def load_daily_targets(self, file):
        """Load daily targets report"""
        try:
            df = pd.read_csv(file)
            
            numeric_cols = ['ROAS', 'Conversion rate', 'CPC', 'Bid', 'Suggested bid',
                           'Impressions', 'Clicks', 'Spend', 'CTR', 'Orders', 'Sales', 'ACOS']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(self.parse_currency)
            
            if self.daily_targets is not None:
                df = pd.concat([self.daily_targets, df])
                df = df.drop_duplicates(subset=['Target', 'Campaign'], keep='last')
            
            self.daily_targets = df
            
            # Save keyword/target history
            if 'Target' in df.columns:
                self.history_manager.save_keyword_snapshot(df)
            
            st.success(f"✅ Loaded Daily Targets: {len(df)} targets")
            return True
        except Exception as e:
            st.error(f"Error loading daily targets: {str(e)}")
            return False
    
    def load_daily_inventory(self, file):
        """Load daily inventory/search term report"""
        try:
            df = pd.read_csv(file)
            
            numeric_cols = ['Impressions', 'Clicks', 'CTR', 'Total cost', 'Purchases',
                           'Sales', 'ROAS', 'Purchase rate']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(self.parse_currency)
            
            dup_cols = ['Campaign name', 'Search term'] if 'Search term' in df.columns else ['Campaign name']
            if self.daily_inventory is not None:
                df = pd.concat([self.daily_inventory, df])
                df = df.drop_duplicates(subset=dup_cols, keep='last')
            
            self.daily_inventory = df
            st.success(f"✅ Loaded Daily Inventory: {len(df)} entries")
            return True
        except Exception as e:
            st.error(f"Error loading daily inventory: {str(e)}")
            return False
    
    def calculate_roi_metrics(self, df):
        """Calculate ROI and performance metrics"""
        if df is None or len(df) == 0:
            return df
        
        df = df.copy()
        
        # Calculate ROI
        sales_col = 'Sales (converted)' if 'Sales (converted)' in df.columns else 'Sales'
        cost_col = 'Total cost (converted)' if 'Total cost (converted)' in df.columns else 'Total cost'
        
        if sales_col in df.columns and cost_col in df.columns:
            df['ROI'] = ((df[sales_col] - df[cost_col]) / df[cost_col].replace(0, np.nan)) * 100
        
        if 'ROAS' in df.columns:
            df['ROAS_score'] = df['ROAS']
        else:
            df['ROAS_score'] = 0
        
        if 'CTR' in df.columns:
            df['CTR_score'] = df['CTR'] * 100
        else:
            df['CTR_score'] = 0
        
        df['Efficiency_Score'] = (df['ROAS_score'] * 0.6 + df['CTR_score'] * 0.4)
        
        return df
    
    def analyze_keyword_performance(self):
        """Analyze keyword-level performance"""
        results = {
            'top_performers': [],
            'poor_performers': [],
            'trending_up': [],
            'trending_down': []
        }
        
        if self.daily_targets is None:
            return results
        
        df = self.daily_targets.copy()
        
        # Filter out keywords with insufficient data
        df = df[df['Spend'] > 50]  # At least ₹50 spend
        
        # Top performers
        if all(col in df.columns for col in ['Target', 'ROAS', 'Sales', 'Spend']):
            top = df.nlargest(20, 'ROAS')
            for _, row in top.iterrows():
                if row['ROAS'] >= 2.0:  # Profitable keywords
                    results['top_performers'].append({
                        'keyword': row['Target'],
                        'campaign': row.get('Campaign', 'Unknown'),
                        'roas': row['ROAS'],
                        'sales': row['Sales'],
                        'spend': row['Spend'],
                        'action': 'Increase bid' if row['ROAS'] >= 3.0 else 'Maintain'
                    })
            
            # Poor performers
            poor = df.nsmallest(20, 'ROAS')
            for _, row in poor.iterrows():
                if row['ROAS'] < 1.5 and row['Spend'] > 100:
                    results['poor_performers'].append({
                        'keyword': row['Target'],
                        'campaign': row.get('Campaign', 'Unknown'),
                        'roas': row['ROAS'],
                        'sales': row['Sales'],
                        'spend': row['Spend'],
                        'action': 'Pause' if row['ROAS'] < 1.0 else 'Reduce bid'
                    })
        
        # Analyze trends from history
        history = self.history_manager.get_keyword_history(days=30)
        if history is not None and 'Target' in history.columns:
            for keyword in df['Target'].unique()[:50]:  # Top 50 keywords
                kw_history = history[history['Target'] == keyword].sort_values('snapshot_date')
                if len(kw_history) >= 2:
                    recent_roas = kw_history['ROAS'].iloc[-3:].mean() if len(kw_history) >= 3 else kw_history['ROAS'].iloc[-1]
                    old_roas = kw_history['ROAS'].iloc[:3].mean() if len(kw_history) >= 6 else kw_history['ROAS'].iloc[0]
                    
                    change = ((recent_roas - old_roas) / max(old_roas, 0.1)) * 100
                    
                    if change > 20:  # 20% improvement
                        results['trending_up'].append({
                            'keyword': keyword,
                            'change_pct': change,
                            'old_roas': old_roas,
                            'new_roas': recent_roas
                        })
                    elif change < -20:  # 20% decline
                        results['trending_down'].append({
                            'keyword': keyword,
                            'change_pct': change,
                            'old_roas': old_roas,
                            'new_roas': recent_roas
                        })
        
        return results
    
    def analyze_product_performance(self):
        """Analyze product-level performance and recommendations"""
        results = {
            'scale_up': [],
            'maintain': [],
            'reduce': [],
            'discontinue': []
        }
        
        # Combine data from repeat purchase and campaign data
        if self.weekly_repeat is None:
            return results
        
        df = self.weekly_repeat.copy()
        
        # Group by product
        if 'Product Title' in df.columns:
            for product in df['Product Title'].unique():
                product_data = df[df['Product Title'] == product]
                
                # Calculate metrics
                total_sales = product_data['Repeat Ordered Product Sales: Sales'].sum() if 'Repeat Ordered Product Sales: Sales' in product_data.columns else 0
                repeat_rate = product_data['Repeat Customer Share: % Share of Total Customers'].mean() if 'Repeat Customer Share: % Share of Total Customers' in product_data.columns else 0
                
                # Decision logic
                if repeat_rate > 5 and total_sales > 10000:
                    results['scale_up'].append({
                        'product': product,
                        'repeat_rate': repeat_rate,
                        'sales': total_sales,
                        'reason': 'High repeat purchase rate and strong sales',
                        'action': 'Increase ad spend by 30-50%'
                    })
                elif repeat_rate > 2 and total_sales > 5000:
                    results['maintain'].append({
                        'product': product,
                        'repeat_rate': repeat_rate,
                        'sales': total_sales,
                        'reason': 'Stable performance',
                        'action': 'Maintain current spend'
                    })
                elif repeat_rate < 1 and total_sales < 3000:
                    results['discontinue'].append({
                        'product': product,
                        'repeat_rate': repeat_rate,
                        'sales': total_sales,
                        'reason': 'Low repeat rate and weak sales',
                        'action': 'Consider discontinuing or major changes'
                    })
                else:
                    results['reduce'].append({
                        'product': product,
                        'repeat_rate': repeat_rate,
                        'sales': total_sales,
                        'reason': 'Below target performance',
                        'action': 'Reduce ad spend by 20-30%'
                    })
        
        return results
    
    def generate_campaign_trends(self):
        """Generate trend data for campaigns"""
        history = self.history_manager.get_campaign_history(days=30)
        
        if history is None or 'Campaign name' not in history.columns:
            return None
        
        trends = []
        
        for campaign in history['Campaign name'].unique()[:20]:  # Top 20 campaigns
            camp_history = history[history['Campaign name'] == campaign].sort_values('snapshot_date')
            
            if len(camp_history) >= 2:
                trend_data = {
                    'campaign': campaign,
                    'dates': camp_history['snapshot_date'].tolist(),
                    'roas': camp_history['ROAS'].tolist() if 'ROAS' in camp_history.columns else [],
                    'sales': camp_history['Sales (converted)'].tolist() if 'Sales (converted)' in camp_history.columns else camp_history['Sales'].tolist() if 'Sales' in camp_history.columns else [],
                    'spend': camp_history['Total cost (converted)'].tolist() if 'Total cost (converted)' in camp_history.columns else camp_history['Total cost'].tolist() if 'Total cost' in camp_history.columns else []
                }
                
                # Calculate trend direction
                if len(trend_data['roas']) >= 2:
                    recent_avg = np.mean(trend_data['roas'][-3:]) if len(trend_data['roas']) >= 3 else trend_data['roas'][-1]
                    old_avg = np.mean(trend_data['roas'][:3]) if len(trend_data['roas']) >= 6 else trend_data['roas'][0]
                    trend_data['direction'] = 'up' if recent_avg > old_avg else 'down'
                    trend_data['change_pct'] = ((recent_avg - old_avg) / max(old_avg, 0.1)) * 100
                
                trends.append(trend_data)
        
        return trends
    
    def generate_recommendations(self):
        """Generate actionable recommendations"""
        recommendations = {
            'expand': [],
            'optimize': [],
            'pause': [],
            'close': []
        }
        
        if self.daily_campaigns is None or len(self.daily_campaigns) == 0:
            return recommendations
        
        df = self.calculate_roi_metrics(self.daily_campaigns)
        
        # Get historical data for trend analysis
        history = self.history_manager.get_campaign_history(days=14)
        
        ROAS_EXCELLENT = 3.0
        ROAS_GOOD = 2.0
        ROAS_POOR = 1.0
        MIN_SPEND = 100
        
        for idx, row in df.iterrows():
            campaign = row.get('Campaign name', 'Unknown')
            roas = row.get('ROAS', 0)
            spend = row.get('Total cost (converted)', row.get('Total cost', 0))
            sales = row.get('Sales (converted)', row.get('Sales', 0))
            purchases = row.get('Purchases', 0)
            
            if spend < MIN_SPEND:
                continue
            
            # Calculate trend if history available
            trend = ''
            if history is not None and campaign in history['Campaign name'].values:
                camp_hist = history[history['Campaign name'] == campaign].sort_values('snapshot_date')
                if len(camp_hist) >= 2:
                    old_roas = camp_hist['ROAS'].iloc[0] if 'ROAS' in camp_hist.columns else 0
                    new_roas = camp_hist['ROAS'].iloc[-1] if 'ROAS' in camp_hist.columns else 0
                    if new_roas > old_roas * 1.1:
                        trend = '📈 Trending up'
                    elif new_roas < old_roas * 0.9:
                        trend = '📉 Trending down'
            
            # EXPAND
            if roas >= ROAS_EXCELLENT and purchases >= 5:
                recommendations['expand'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'trend': trend,
                    'reason': f'Excellent ROAS ({roas:.2f}x) with consistent conversions',
                    'action': f'Increase budget by 30-50%'
                })
            
            # OPTIMIZE
            elif roas >= ROAS_GOOD and roas < ROAS_EXCELLENT:
                recommendations['optimize'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'trend': trend,
                    'reason': f'Good ROAS ({roas:.2f}x) but can be improved',
                    'action': 'Review keywords, adjust bids, test new creatives'
                })
            
            # PAUSE
            elif roas >= ROAS_POOR and roas < ROAS_GOOD:
                recommendations['pause'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'trend': trend,
                    'reason': f'Below target ROAS ({roas:.2f}x)',
                    'action': 'Pause and analyze. Review targeting, keywords, and audience'
                })
            
            # CLOSE
            elif roas < ROAS_POOR and spend > 500:
                recommendations['close'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'trend': trend,
                    'reason': f'Very low ROAS ({roas:.2f}x) with significant spend',
                    'action': 'Close campaign or completely restructure'
                })
        
        # Sort by impact
        for key in recommendations:
            recommendations[key] = sorted(
                recommendations[key],
                key=lambda x: x['spend'] * (x['roas'] if key == 'expand' else 1/max(x['roas'], 0.1)),
                reverse=True
            )
        
        return recommendations
    
    def get_summary_stats(self):
        """Get summary statistics"""
        stats = {}
        
        if self.daily_campaigns is not None and len(self.daily_campaigns) > 0:
            df = self.daily_campaigns
            
            stats['total_campaigns'] = len(df)
            stats['active_campaigns'] = len(df[df['State'] == 'ENABLED']) if 'State' in df.columns else len(df)
            stats['total_spend'] = df['Total cost (converted)'].sum() if 'Total cost (converted)' in df.columns else df['Total cost'].sum() if 'Total cost' in df.columns else 0
            stats['total_sales'] = df['Sales (converted)'].sum() if 'Sales (converted)' in df.columns else df['Sales'].sum() if 'Sales' in df.columns else 0
            stats['avg_roas'] = df['ROAS'].mean() if 'ROAS' in df.columns else 0
            stats['total_purchases'] = df['Purchases'].sum() if 'Purchases' in df.columns else 0
            stats['total_clicks'] = df['Clicks'].sum() if 'Clicks' in df.columns else 0
            stats['avg_ctr'] = df['CTR'].mean() if 'CTR' in df.columns else 0
        
        return stats


def main():
    """Main Streamlit application"""
    
    st.markdown('<p class="main-header">🎯 Amazon Marketing ROI Analyzer</p>', unsafe_allow_html=True)
    st.markdown("**Analyze performance, track trends, optimize ROI with data-driven recommendations**")
    
    # Initialize session state
    if 'analyzer' not in st.session_state:
        st.session_state.analyzer = AmazonROIAnalyzer()
    
    analyzer = st.session_state.analyzer
    
    # Sidebar
    with st.sidebar:
        st.header("📁 Data Upload")
        
        st.subheader("Weekly Reports (Every Friday)")
        weekly_sales = st.file_uploader("Weekly Sales Report (Excel)", type=['xlsx'], key='weekly_sales')
        weekly_campaigns = st.file_uploader("Weekly Campaigns Report (Excel)", type=['xlsx'], key='weekly_campaigns')
        weekly_cpr = st.file_uploader("Weekly CPR Report (Excel)", type=['xlsx'], key='weekly_cpr')
        weekly_repeat = st.file_uploader("Weekly Repeat Purchase (Excel)", type=['xlsx'], key='weekly_repeat')
        
        st.subheader("Daily Reports")
        daily_campaigns = st.file_uploader("Daily Campaign Report (CSV)", type=['csv'], key='daily_campaigns')
        daily_targets = st.file_uploader("Daily Targets Report (CSV)", type=['csv'], key='daily_targets')
        daily_inventory = st.file_uploader("Daily Inventory Report (CSV)", type=['csv'], key='daily_inventory')
        
        if st.button("🔄 Load All Data", type="primary"):
            with st.spinner("Loading data..."):
                if weekly_sales:
                    analyzer.load_weekly_sales(weekly_sales)
                if weekly_campaigns:
                    analyzer.load_weekly_campaigns(weekly_campaigns)
                if weekly_cpr:
                    analyzer.load_weekly_cpr(weekly_cpr)
                if weekly_repeat:
                    analyzer.load_weekly_repeat(weekly_repeat)
                if daily_campaigns:
                    analyzer.load_daily_campaigns(daily_campaigns)
                if daily_targets:
                    analyzer.load_daily_targets(daily_targets)
                if daily_inventory:
                    analyzer.load_daily_inventory(daily_inventory)
                
                st.success("✅ All data loaded successfully!")
                st.info("📊 Historical data saved for trend analysis")
    
    # Main tabs
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📊 Dashboard",
        "🎯 Campaign Recommendations",
        "🔑 Keyword Performance",
        "📦 Product Analysis",
        "📈 Trends & History",
        "🔍 Detailed Data",
        "💡 Insights"
    ])
    
    # Tab 1: Dashboard
    with tab1:
        st.header("Performance Dashboard")
        
        stats = analyzer.get_summary_stats()
        
        if stats:
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Campaigns", f"{stats.get('total_campaigns', 0)}", 
                         f"{stats.get('active_campaigns', 0)} active")
            with col2:
                st.metric("Total Spend", f"₹{stats.get('total_spend', 0):,.0f}")
            with col3:
                roi_pct = ((stats.get('total_sales', 0) / max(stats.get('total_spend', 1), 1) - 1) * 100)
                st.metric("Total Sales", f"₹{stats.get('total_sales', 0):,.0f}", 
                         f"{roi_pct:.1f}% ROI")
            with col4:
                st.metric("Avg ROAS", f"{stats.get('avg_roas', 0):.2f}x")
            
            st.markdown("---")
            
            col5, col6, col7 = st.columns(3)
            with col5:
                st.metric("Total Purchases", f"{stats.get('total_purchases', 0):,.0f}")
            with col6:
                st.metric("Total Clicks", f"{stats.get('total_clicks', 0):,.0f}")
            with col7:
                st.metric("Avg CTR", f"{stats.get('avg_ctr', 0):.2f}%")
            
            # Visualizations
            if analyzer.daily_campaigns is not None and len(analyzer.daily_campaigns) > 0:
                st.markdown("---")
                st.subheader("Campaign Performance Overview")
                
                df_viz = analyzer.calculate_roi_metrics(analyzer.daily_campaigns)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**Top 10 Campaigns by ROAS**")
                    
                    if 'ROAS' in df_viz.columns and 'Campaign name' in df_viz.columns:
                        sales_col = 'Sales (converted)' if 'Sales (converted)' in df_viz.columns else 'Sales'
                        
                        cols_to_show = ['Campaign name', 'ROAS']
                        if sales_col in df_viz.columns:
                            cols_to_show.append(sales_col)
                        
                        top_roas = df_viz.nlargest(10, 'ROAS')[cols_to_show]
                        
                        fig = px.bar(top_roas, x='ROAS', y='Campaign name', orientation='h',
                                    color='ROAS', color_continuous_scale='RdYlGn',
                                    title='Top Performers')
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.warning("Insufficient data for ROAS chart")
                
                with col2:
                    st.markdown("**Spend vs Sales**")
                    spend_col = 'Total cost (converted)' if 'Total cost (converted)' in df_viz.columns else 'Total cost'
                    sales_col = 'Sales (converted)' if 'Sales (converted)' in df_viz.columns else 'Sales'
                    
                    if all(col in df_viz.columns for col in [spend_col, sales_col, 'ROAS', 'Campaign name']):
                        fig = px.scatter(df_viz, x=spend_col, y=sales_col, size='ROAS', color='ROAS',
                                        hover_data=['Campaign name'], color_continuous_scale='RdYlGn',
                                        title='ROI Efficiency')
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.warning("Insufficient data for Spend vs Sales chart")
        else:
            st.info("👈 Please upload data files from the sidebar")
    
    # Tab 2: Campaign Recommendations
    with tab2:
        st.header("🎯 Campaign Recommendations")
        
        if analyzer.daily_campaigns is not None:
            recommendations = analyzer.generate_recommendations()
            
            # EXPAND
            if recommendations['expand']:
                st.markdown("### 🚀 EXPAND - Scale These Winners")
                st.markdown("High performers with excellent ROAS. Increase budget to maximize returns.")
                
                for rec in recommendations['expand'][:10]:
                    trend_indicator = f"<br><span class='trending-{rec['trend'].split()[1] if rec['trend'] else 'neutral'}'>{rec['trend']}</span>" if rec['trend'] else ""
                    st.markdown(f"""
                    <div class="recommendation-box expand">
                        <strong>{rec['campaign']}</strong>{trend_indicator}<br>
                        📈 ROAS: {rec['roas']:.2f}x | 💰 Sales: ₹{rec['sales']:,.0f} | 💵 Spend: ₹{rec['spend']:,.0f}<br>
                        <em>{rec['reason']}</em><br>
                        <strong>Action:</strong> {rec['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # OPTIMIZE
            if recommendations['optimize']:
                st.markdown("### ⚙️ OPTIMIZE - Improve Performance")
                st.markdown("Good campaigns with potential for optimization.")
                
                for rec in recommendations['optimize'][:10]:
                    trend_indicator = f"<br><span class='trending-{rec['trend'].split()[1] if rec['trend'] else 'neutral'}'>{rec['trend']}</span>" if rec['trend'] else ""
                    st.markdown(f"""
                    <div class="recommendation-box optimize">
                        <strong>{rec['campaign']}</strong>{trend_indicator}<br>
                        📈 ROAS: {rec['roas']:.2f}x | 💰 Sales: ₹{rec['sales']:,.0f} | 💵 Spend: ₹{rec['spend']:,.0f}<br>
                        <em>{rec['reason']}</em><br>
                        <strong>Action:</strong> {rec['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # PAUSE
            if recommendations['pause']:
                st.markdown("### ⏸️ PAUSE - Needs Review")
                st.markdown("Underperforming campaigns that should be paused for analysis.")
                
                for rec in recommendations['pause'][:10]:
                    trend_indicator = f"<br><span class='trending-{rec['trend'].split()[1] if rec['trend'] else 'neutral'}'>{rec['trend']}</span>" if rec['trend'] else ""
                    st.markdown(f"""
                    <div class="recommendation-box pause">
                        <strong>{rec['campaign']}</strong>{trend_indicator}<br>
                        📉 ROAS: {rec['roas']:.2f}x | 💰 Sales: ₹{rec['sales']:,.0f} | 💵 Spend: ₹{rec['spend']:,.0f}<br>
                        <em>{rec['reason']}</em><br>
                        <strong>Action:</strong> {rec['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # CLOSE
            if recommendations['close']:
                st.markdown("### ❌ CLOSE - Stop the Bleed")
                st.markdown("Money-losing campaigns that should be closed immediately.")
                
                for rec in recommendations['close'][:10]:
                    trend_indicator = f"<br><span class='trending-{rec['trend'].split()[1] if rec['trend'] else 'neutral'}'>{rec['trend']}</span>" if rec['trend'] else ""
                    st.markdown(f"""
                    <div class="recommendation-box close">
                        <strong>{rec['campaign']}</strong>{trend_indicator}<br>
                        ❌ ROAS: {rec['roas']:.2f}x | 💰 Sales: ₹{rec['sales']:,.0f} | 💵 Spend: ₹{rec['spend']:,.0f}<br>
                        <em>{rec['reason']}</em><br>
                        <strong>Action:</strong> {rec['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # Summary
            st.markdown("---")
            st.subheader("Summary")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Expand", len(recommendations['expand']))
            with col2:
                st.metric("Optimize", len(recommendations['optimize']))
            with col3:
                st.metric("Pause", len(recommendations['pause']))
            with col4:
                st.metric("Close", len(recommendations['close']))
        else:
            st.info("👈 Please upload campaign data")
    
    # Tab 3: Keyword Performance
    with tab3:
        st.header("🔑 Keyword Performance Analysis")
        
        keyword_analysis = analyzer.analyze_keyword_performance()
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("✅ Top Performing Keywords")
            if keyword_analysis['top_performers']:
                for kw in keyword_analysis['top_performers'][:15]:
                    st.markdown(f"""
                    <div class="recommendation-box expand">
                        <strong>{kw['keyword']}</strong><br>
                        Campaign: {kw['campaign']}<br>
                        ROAS: {kw['roas']:.2f}x | Sales: ₹{kw['sales']:,.0f} | Spend: ₹{kw['spend']:,.0f}<br>
                        <strong>Action:</strong> {kw['action']}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No keyword data available. Upload Daily Targets report.")
        
        with col2:
            st.subheader("❌ Poor Performing Keywords")
            if keyword_analysis['poor_performers']:
                for kw in keyword_analysis['poor_performers'][:15]:
                    st.markdown(f"""
                    <div class="recommendation-box close">
                        <strong>{kw['keyword']}</strong><br>
                        Campaign: {kw['campaign']}<br>
                        ROAS: {kw['roas']:.2f}x | Sales: ₹{kw['sales']:,.0f} | Spend: ₹{kw['spend']:,.0f}<br>
                        <strong>Action:</strong> {kw['action']}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No underperforming keywords detected.")
        
        st.markdown("---")
        
        col3, col4 = st.columns(2)
        
        with col3:
            st.subheader("📈 Trending Up")
            if keyword_analysis['trending_up']:
                for kw in keyword_analysis['trending_up'][:10]:
                    st.success(f"**{kw['keyword']}**: {kw['old_roas']:.2f}x → {kw['new_roas']:.2f}x ({kw['change_pct']:+.1f}%)")
            else:
                st.info("Need more historical data to detect trends")
        
        with col4:
            st.subheader("📉 Trending Down")
            if keyword_analysis['trending_down']:
                for kw in keyword_analysis['trending_down'][:10]:
                    st.warning(f"**{kw['keyword']}**: {kw['old_roas']:.2f}x → {kw['new_roas']:.2f}x ({kw['change_pct']:+.1f}%)")
            else:
                st.info("No declining trends detected")
    
    # Tab 4: Product Analysis
    with tab4:
        st.header("📦 Product Performance & Recommendations")
        
        product_analysis = analyzer.analyze_product_performance()
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🚀 Products to Scale Up")
            if product_analysis['scale_up']:
                for prod in product_analysis['scale_up']:
                    st.markdown(f"""
                    <div class="recommendation-box expand">
                        <strong>{prod['product']}</strong><br>
                        Repeat Rate: {prod['repeat_rate']:.1f}% | Sales: ₹{prod['sales']:,.0f}<br>
                        <em>{prod['reason']}</em><br>
                        <strong>Action:</strong> {prod['action']}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No high-performing products detected. Upload Weekly Repeat Purchase report.")
            
            st.subheader("✅ Products to Maintain")
            if product_analysis['maintain']:
                for prod in product_analysis['maintain']:
                    st.markdown(f"""
                    <div class="recommendation-box optimize">
                        <strong>{prod['product']}</strong><br>
                        Repeat Rate: {prod['repeat_rate']:.1f}% | Sales: ₹{prod['sales']:,.0f}<br>
                        <em>{prod['reason']}</em><br>
                        <strong>Action:</strong> {prod['action']}
                    </div>
                    """, unsafe_allow_html=True)
        
        with col2:
            st.subheader("⚠️ Products to Reduce Spend")
            if product_analysis['reduce']:
                for prod in product_analysis['reduce']:
                    st.markdown(f"""
                    <div class="recommendation-box pause">
                        <strong>{prod['product']}</strong><br>
                        Repeat Rate: {prod['repeat_rate']:.1f}% | Sales: ₹{prod['sales']:,.0f}<br>
                        <em>{prod['reason']}</em><br>
                        <strong>Action:</strong> {prod['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            st.subheader("❌ Products to Discontinue")
            if product_analysis['discontinue']:
                for prod in product_analysis['discontinue']:
                    st.markdown(f"""
                    <div class="recommendation-box close">
                        <strong>{prod['product']}</strong><br>
                        Repeat Rate: {prod['repeat_rate']:.1f}% | Sales: ₹{prod['sales']:,.0f}<br>
                        <em>{prod['reason']}</em><br>
                        <strong>Action:</strong> {prod['action']}
                    </div>
                    """, unsafe_allow_html=True)
    
    # Tab 5: Trends & History
    with tab5:
        st.header("📈 Performance Trends & Historical Analysis")
        
        trends = analyzer.generate_campaign_trends()
        
        if trends:
            st.subheader("Campaign Performance Over Time")
            
            # Select campaign to view
            campaign_names = [t['campaign'] for t in trends]
            selected_campaign = st.selectbox("Select Campaign to View Trend", campaign_names)
            
            if selected_campaign:
                camp_data = next((t for t in trends if t['campaign'] == selected_campaign), None)
                
                if camp_data and len(camp_data['dates']) > 1:
                    # Create trend chart
                    fig = go.Figure()
                    
                    if camp_data['roas']:
                        fig.add_trace(go.Scatter(
                            x=camp_data['dates'],
                            y=camp_data['roas'],
                            name='ROAS',
                            line=dict(color='#28a745', width=3)
                        ))
                    
                    fig.update_layout(
                        title=f"ROAS Trend: {selected_campaign}",
                        xaxis_title="Date",
                        yaxis_title="ROAS",
                        height=400,
                        hovermode='x unified'
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Show change metrics
                    if 'change_pct' in camp_data:
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Trend Direction", 
                                     "↗️ Improving" if camp_data['direction'] == 'up' else "↘️ Declining")
                        with col2:
                            st.metric("Change", f"{camp_data['change_pct']:+.1f}%")
            
            st.markdown("---")
            st.subheader("All Campaigns Trend Summary")
            
            trend_summary = []
            for t in trends:
                if 'change_pct' in t:
                    trend_summary.append({
                        'Campaign': t['campaign'],
                        'Direction': '↗️ Up' if t['direction'] == 'up' else '↘️ Down',
                        'Change %': f"{t['change_pct']:+.1f}%"
                    })
            
            if trend_summary:
                st.dataframe(pd.DataFrame(trend_summary), use_container_width=True)
        else:
            st.info("📊 Historical data will appear here after multiple data uploads over time.")
            st.markdown("""
            **To enable trend analysis:**
            1. Upload your data regularly (weekly recommended)
            2. Data is automatically saved to track performance over time
            3. After 2+ uploads, trend lines and change metrics will appear
            """)
    
    # Tab 6: Detailed Data
    with tab4:
        st.header("🔍 Detailed Data View")
        
        data_view = st.selectbox(
            "Select Data to View",
            ["Daily Campaigns", "Daily Targets", "Daily Inventory", 
             "Weekly Campaigns", "Weekly CPR", "Weekly Repeat Purchase"]
        )
        
        if data_view == "Daily Campaigns" and analyzer.daily_campaigns is not None:
            st.dataframe(analyzer.daily_campaigns, use_container_width=True)
        elif data_view == "Daily Targets" and analyzer.daily_targets is not None:
            st.dataframe(analyzer.daily_targets, use_container_width=True)
        elif data_view == "Daily Inventory" and analyzer.daily_inventory is not None:
            st.dataframe(analyzer.daily_inventory, use_container_width=True)
        elif data_view == "Weekly Campaigns" and analyzer.weekly_campaigns is not None:
            st.dataframe(analyzer.weekly_campaigns, use_container_width=True)
        elif data_view == "Weekly CPR" and analyzer.weekly_cpr is not None:
            st.dataframe(analyzer.weekly_cpr, use_container_width=True)
        elif data_view == "Weekly Repeat Purchase" and analyzer.weekly_repeat is not None:
            st.dataframe(analyzer.weekly_repeat, use_container_width=True)
        else:
            st.info(f"No data loaded for {data_view}")
    
    # Tab 7: Insights
    with tab7:
        st.header("💡 Key Insights & Action Summary")
        
        if analyzer.daily_campaigns is not None and len(analyzer.daily_campaigns) > 0:
            df = analyzer.calculate_roi_metrics(analyzer.daily_campaigns)
            
            st.subheader("📊 Quick Insights")
            
            spend_col = 'Total cost (converted)' if 'Total cost (converted)' in df.columns else 'Total cost'
            sales_col = 'Sales (converted)' if 'Sales (converted)' in df.columns else 'Sales'
            
            if 'ROAS' in df.columns and len(df) > 0:
                best = df.nlargest(1, 'ROAS').iloc[0]
                st.success(f"🏆 **Best ROAS:** {best['Campaign name']} with {best['ROAS']:.2f}x ROAS")
                
                if spend_col in df.columns:
                    highest_spend = df.nlargest(1, spend_col).iloc[0]
                    st.info(f"💰 **Highest Spend:** {highest_spend['Campaign name']} spent ₹{highest_spend[spend_col]:,.0f}")
                
                if sales_col in df.columns:
                    most_sales = df.nlargest(1, sales_col).iloc[0]
                    st.info(f"📈 **Highest Sales:** {most_sales['Campaign name']} generated ₹{most_sales[sales_col]:,.0f}")
            
            # Performance Matrix
            st.markdown("---")
            st.subheader("Performance Matrix")
            
            if all(col in df.columns for col in ['ROAS', spend_col]):
                median_roas = df['ROAS'].median()
                median_spend = df[spend_col].median()
                
                df['Quadrant'] = df.apply(
                    lambda row: 'Star' if row['ROAS'] >= median_roas and row[spend_col] >= median_spend
                    else 'Question Mark' if row['ROAS'] >= median_roas and row[spend_col] < median_spend
                    else 'Cash Cow' if row['ROAS'] < median_roas and row[spend_col] >= median_spend
                    else 'Dog',
                    axis=1
                )
                
                if sales_col in df.columns:
                    fig = px.scatter(df, x=spend_col, y='ROAS', color='Quadrant',
                                    size=sales_col, hover_data=['Campaign name'],
                                    title='Campaign Performance Matrix',
                                    color_discrete_map={
                                        'Star': '#28a745',
                                        'Question Mark': '#ffc107',
                                        'Cash Cow': '#17a2b8',
                                        'Dog': '#dc3545'
                                    })
                    
                    fig.add_hline(y=median_roas, line_dash="dash", line_color="gray", 
                                 annotation_text="Median ROAS")
                    fig.add_vline(x=median_spend, line_dash="dash", line_color="gray", 
                                 annotation_text="Median Spend")
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("⭐ Stars", len(df[df['Quadrant'] == 'Star']), 
                                 help="High ROAS, High Spend - Scale these!")
                    with col2:
                        st.metric("❓ Question Marks", len(df[df['Quadrant'] == 'Question Mark']),
                                 help="High ROAS, Low Spend - Invest more")
                    with col3:
                        st.metric("🐮 Cash Cows", len(df[df['Quadrant'] == 'Cash Cow']),
                                 help="Low ROAS, High Spend - Optimize or reduce")
                    with col4:
                        st.metric("🐕 Dogs", len(df[df['Quadrant'] == 'Dog']),
                                 help="Low ROAS, Low Spend - Consider closing")
        else:
            st.info("👈 Please upload campaign data to see insights")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "Amazon Marketing ROI Analyzer | Enhanced with Historical Tracking & Trend Analysis"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
