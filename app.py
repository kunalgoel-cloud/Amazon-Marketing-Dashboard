"""
Amazon Marketing ROI Analyzer
A Streamlit tool to analyze weekly and daily Amazon marketing reports
and provide actionable recommendations to increase ROI.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from pathlib import Path
import io

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
    .pause {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
    }
    .close {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
    }
    </style>
""", unsafe_allow_html=True)

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
        
    def parse_currency(self, value):
        """Parse currency string to float"""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        # Remove currency symbols and commas
        cleaned = str(value).replace('₹', '').replace(',', '').strip()
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
            # Read all sheets
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
                # Clean numeric columns
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
        """Load weekly CPR (Customer Purchase Rate) report"""
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
                return True
        except Exception as e:
            st.error(f"Error loading weekly repeat purchase: {str(e)}")
            return False
    
    def load_daily_campaigns(self, file):
        """Load daily campaign report"""
        try:
            df = pd.read_csv(file)
            
            # Parse date if available
            if 'Campaign start date' in df.columns:
                df['Campaign start date'] = pd.to_datetime(df['Campaign start date'], errors='coerce')
            
            # Clean numeric columns
            numeric_cols = ['Clicks', 'CTR', 'Total cost (converted)', 'CPC (converted)', 
                           'Purchases', 'Sales (converted)', 'ROAS']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(self.parse_currency)
            
            # Remove duplicates based on Campaign name
            if self.daily_campaigns is not None:
                # Merge with existing, keeping latest data
                df = pd.concat([self.daily_campaigns, df])
                df = df.drop_duplicates(subset=['Campaign name'], keep='last')
            
            self.daily_campaigns = df
            st.success(f"✅ Loaded Daily Campaigns: {len(df)} campaigns")
            return True
        except Exception as e:
            st.error(f"Error loading daily campaigns: {str(e)}")
            return False
    
    def load_daily_targets(self, file):
        """Load daily targets report"""
        try:
            df = pd.read_csv(file)
            
            # Clean numeric columns
            numeric_cols = ['ROAS', 'Conversion rate', 'CPC', 'Bid', 'Suggested bid',
                           'Impressions', 'Clicks', 'Spend', 'CTR', 'Orders', 'Sales', 'ACOS']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(self.parse_currency)
            
            # Remove duplicates based on Target and Campaign
            if self.daily_targets is not None:
                df = pd.concat([self.daily_targets, df])
                df = df.drop_duplicates(subset=['Target', 'Campaign'], keep='last')
            
            self.daily_targets = df
            st.success(f"✅ Loaded Daily Targets: {len(df)} targets")
            return True
        except Exception as e:
            st.error(f"Error loading daily targets: {str(e)}")
            return False
    
    def load_daily_inventory(self, file):
        """Load daily inventory/search term report"""
        try:
            df = pd.read_csv(file)
            
            # Parse date range
            if 'Date range' in df.columns:
                df['Date range'] = df['Date range'].astype(str)
            
            # Clean numeric columns
            numeric_cols = ['Impressions', 'Clicks', 'CTR', 'Total cost', 'Purchases',
                           'Sales', 'ROAS', 'Purchase rate']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(self.parse_currency)
            
            # Remove duplicates based on Campaign name and Search term
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
    
    def combine_data(self):
        """Combine all data sources for analysis"""
        try:
            # Start with daily campaigns as base
            if self.daily_campaigns is not None:
                combined = self.daily_campaigns.copy()
                
                # Add weekly campaign data if not already present
                if self.weekly_campaigns is not None:
                    weekly_agg = self.weekly_campaigns.groupby('Campaign name').agg({
                        'Impressions': 'sum',
                        'Clicks': 'sum',
                        'Total cost': 'sum',
                        'Purchases': 'sum',
                        'Sales': 'sum',
                        'ROAS': 'mean',
                        'ACOS': 'mean',
                        'CVR': 'mean'
                    }).reset_index()
                    weekly_agg.columns = ['Campaign name'] + [f'Weekly_{col}' for col in weekly_agg.columns[1:]]
                    
                    combined = combined.merge(weekly_agg, on='Campaign name', how='left')
                
                self.combined_data = combined
                return True
            
            return False
        except Exception as e:
            st.error(f"Error combining data: {str(e)}")
            return False
    
    def calculate_roi_metrics(self, df):
        """Calculate ROI and performance metrics"""
        if df is None or len(df) == 0:
            return df
        
        df = df.copy()
        
        # Calculate metrics
        if 'Sales (converted)' in df.columns and 'Total cost (converted)' in df.columns:
            df['ROI'] = ((df['Sales (converted)'] - df['Total cost (converted)']) / 
                        df['Total cost (converted)'].replace(0, np.nan)) * 100
        elif 'Sales' in df.columns and 'Total cost' in df.columns:
            df['ROI'] = ((df['Sales'] - df['Total cost']) / 
                        df['Total cost'].replace(0, np.nan)) * 100
        
        # Calculate efficiency score (combines ROAS, CTR, CVR)
        if 'ROAS' in df.columns:
            df['ROAS_score'] = df['ROAS']
        else:
            df['ROAS_score'] = 0
            
        if 'CTR' in df.columns:
            df['CTR_score'] = df['CTR'] * 100  # Normalize
        else:
            df['CTR_score'] = 0
        
        df['Efficiency_Score'] = (df['ROAS_score'] * 0.6 + df['CTR_score'] * 0.4)
        
        return df
    
    def generate_recommendations(self):
        """Generate actionable recommendations based on data analysis"""
        recommendations = {
            'expand': [],
            'optimize': [],
            'pause': [],
            'close': []
        }
        
        if self.daily_campaigns is None or len(self.daily_campaigns) == 0:
            return recommendations
        
        # Calculate metrics
        df = self.calculate_roi_metrics(self.daily_campaigns)
        
        # Define thresholds
        ROAS_EXCELLENT = 3.0
        ROAS_GOOD = 2.0
        ROAS_POOR = 1.0
        ACOS_EXCELLENT = 20
        ACOS_GOOD = 35
        ACOS_POOR = 50
        MIN_SPEND = 100
        
        for idx, row in df.iterrows():
            campaign = row.get('Campaign name', 'Unknown')
            roas = row.get('ROAS', 0)
            spend = row.get('Total cost (converted)', row.get('Total cost', 0))
            sales = row.get('Sales (converted)', row.get('Sales', 0))
            clicks = row.get('Clicks', 0)
            purchases = row.get('Purchases', 0)
            
            # Skip if no significant data
            if spend < MIN_SPEND:
                continue
            
            # EXPAND: High performers
            if roas >= ROAS_EXCELLENT and purchases >= 5:
                recommendations['expand'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'reason': f'Excellent ROAS ({roas:.2f}x) with consistent conversions',
                    'action': f'Increase budget by 30-50%'
                })
            
            # OPTIMIZE: Good performers with room to improve
            elif roas >= ROAS_GOOD and roas < ROAS_EXCELLENT:
                recommendations['optimize'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'reason': f'Good ROAS ({roas:.2f}x) but can be improved',
                    'action': 'Review keywords, adjust bids, test new creatives'
                })
            
            # PAUSE: Poor performers with some potential
            elif roas >= ROAS_POOR and roas < ROAS_GOOD:
                recommendations['pause'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'reason': f'Below target ROAS ({roas:.2f}x)',
                    'action': 'Pause and analyze. Review targeting, keywords, and audience'
                })
            
            # CLOSE: Very poor performers
            elif roas < ROAS_POOR and spend > 500:
                recommendations['close'].append({
                    'campaign': campaign,
                    'roas': roas,
                    'sales': sales,
                    'spend': spend,
                    'reason': f'Very low ROAS ({roas:.2f}x) with significant spend',
                    'action': 'Close campaign or completely restructure'
                })
        
        # Sort recommendations by impact (spend * roas potential)
        for key in recommendations:
            recommendations[key] = sorted(
                recommendations[key], 
                key=lambda x: x['spend'] * (x['roas'] if key == 'expand' else 1/max(x['roas'], 0.1)),
                reverse=True
            )
        
        return recommendations
    
    def get_summary_stats(self):
        """Get summary statistics across all data"""
        stats = {}
        
        if self.daily_campaigns is not None and len(self.daily_campaigns) > 0:
            df = self.daily_campaigns
            
            stats['total_campaigns'] = len(df)
            stats['active_campaigns'] = len(df[df['State'] == 'ENABLED']) if 'State' in df.columns else len(df)
            stats['total_spend'] = df['Total cost (converted)'].sum() if 'Total cost (converted)' in df.columns else 0
            stats['total_sales'] = df['Sales (converted)'].sum() if 'Sales (converted)' in df.columns else 0
            stats['avg_roas'] = df['ROAS'].mean() if 'ROAS' in df.columns else 0
            stats['total_purchases'] = df['Purchases'].sum() if 'Purchases' in df.columns else 0
            stats['total_clicks'] = df['Clicks'].sum() if 'Clicks' in df.columns else 0
            stats['avg_ctr'] = df['CTR'].mean() if 'CTR' in df.columns else 0
        
        return stats


def main():
    """Main Streamlit application"""
    
    st.markdown('<p class="main-header">🎯 Amazon Marketing ROI Analyzer</p>', unsafe_allow_html=True)
    st.markdown("**Analyze weekly and daily reports to optimize your Amazon advertising ROI**")
    
    # Initialize session state
    if 'analyzer' not in st.session_state:
        st.session_state.analyzer = AmazonROIAnalyzer()
    
    analyzer = st.session_state.analyzer
    
    # Sidebar for file uploads
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
                # Load weekly reports
                if weekly_sales:
                    analyzer.load_weekly_sales(weekly_sales)
                if weekly_campaigns:
                    analyzer.load_weekly_campaigns(weekly_campaigns)
                if weekly_cpr:
                    analyzer.load_weekly_cpr(weekly_cpr)
                if weekly_repeat:
                    analyzer.load_weekly_repeat(weekly_repeat)
                
                # Load daily reports
                if daily_campaigns:
                    analyzer.load_daily_campaigns(daily_campaigns)
                if daily_targets:
                    analyzer.load_daily_targets(daily_targets)
                if daily_inventory:
                    analyzer.load_daily_inventory(daily_inventory)
                
                # Combine data
                analyzer.combine_data()
                
                st.success("✅ All data loaded successfully!")
    
    # Main content tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Dashboard", 
        "🎯 Recommendations", 
        "📈 Campaign Analysis",
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
                st.metric(
                    "Total Campaigns",
                    f"{stats.get('total_campaigns', 0)}",
                    f"{stats.get('active_campaigns', 0)} active"
                )
            
            with col2:
                st.metric(
                    "Total Spend",
                    f"₹{stats.get('total_spend', 0):,.0f}"
                )
            
            with col3:
                st.metric(
                    "Total Sales",
                    f"₹{stats.get('total_sales', 0):,.0f}",
                    f"{((stats.get('total_sales', 0) / max(stats.get('total_spend', 1), 1) - 1) * 100):.1f}% ROI"
                )
            
            with col4:
                st.metric(
                    "Avg ROAS",
                    f"{stats.get('avg_roas', 0):.2f}x"
                )
            
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
                
                # Top campaigns by ROAS
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**Top 10 Campaigns by ROAS**")
                    top_roas = df_viz.nlargest(10, 'ROAS')[['Campaign name', 'ROAS', 'Sales (converted)']]
                    if 'Sales (converted)' not in top_roas.columns:
                        top_roas = df_viz.nlargest(10, 'ROAS')[['Campaign name', 'ROAS', 'Sales']]
                    
                    fig = px.bar(
                        top_roas,
                        x='ROAS',
                        y='Campaign name',
                        orientation='h',
                        color='ROAS',
                        color_continuous_scale='RdYlGn',
                        title='Top Performers'
                    )
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    st.markdown("**Spend vs Sales**")
                    spend_col = 'Total cost (converted)' if 'Total cost (converted)' in df_viz.columns else 'Total cost'
                    sales_col = 'Sales (converted)' if 'Sales (converted)' in df_viz.columns else 'Sales'
                    
                    fig = px.scatter(
                        df_viz,
                        x=spend_col,
                        y=sales_col,
                        size='ROAS',
                        color='ROAS',
                        hover_data=['Campaign name'],
                        color_continuous_scale='RdYlGn',
                        title='ROI Efficiency'
                    )
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
        
        else:
            st.info("👈 Please upload data files from the sidebar to see the dashboard")
    
    # Tab 2: Recommendations
    with tab2:
        st.header("🎯 AI-Powered Recommendations")
        
        if analyzer.daily_campaigns is not None:
            recommendations = analyzer.generate_recommendations()
            
            # Expand recommendations
            if recommendations['expand']:
                st.markdown("### 🚀 EXPAND - Scale These Winners")
                st.markdown("These campaigns are performing exceptionally well. Increase budget to maximize returns.")
                
                for rec in recommendations['expand'][:10]:  # Top 10
                    st.markdown(f"""
                    <div class="recommendation-box expand">
                        <strong>{rec['campaign']}</strong><br>
                        📈 ROAS: {rec['roas']:.2f}x | 💰 Sales: ₹{rec['sales']:,.0f} | 💵 Spend: ₹{rec['spend']:,.0f}<br>
                        <em>{rec['reason']}</em><br>
                        <strong>Action:</strong> {rec['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # Optimize recommendations
            if recommendations['optimize']:
                st.markdown("### ⚙️ OPTIMIZE - Improve Performance")
                st.markdown("These campaigns have potential but need optimization.")
                
                for rec in recommendations['optimize'][:10]:
                    st.markdown(f"""
                    <div class="recommendation-box pause">
                        <strong>{rec['campaign']}</strong><br>
                        📈 ROAS: {rec['roas']:.2f}x | 💰 Sales: ₹{rec['sales']:,.0f} | 💵 Spend: ₹{rec['spend']:,.0f}<br>
                        <em>{rec['reason']}</em><br>
                        <strong>Action:</strong> {rec['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # Pause recommendations
            if recommendations['pause']:
                st.markdown("### ⏸️ PAUSE - Needs Review")
                st.markdown("These campaigns are underperforming and should be paused for analysis.")
                
                for rec in recommendations['pause'][:10]:
                    st.markdown(f"""
                    <div class="recommendation-box pause">
                        <strong>{rec['campaign']}</strong><br>
                        📉 ROAS: {rec['roas']:.2f}x | 💰 Sales: ₹{rec['sales']:,.0f} | 💵 Spend: ₹{rec['spend']:,.0f}<br>
                        <em>{rec['reason']}</em><br>
                        <strong>Action:</strong> {rec['action']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # Close recommendations
            if recommendations['close']:
                st.markdown("### ❌ CLOSE - Stop the Bleed")
                st.markdown("These campaigns are losing money and should be closed immediately.")
                
                for rec in recommendations['close'][:10]:
                    st.markdown(f"""
                    <div class="recommendation-box close">
                        <strong>{rec['campaign']}</strong><br>
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
            st.info("👈 Please upload campaign data to see recommendations")
    
    # Tab 3: Campaign Analysis
    with tab3:
        st.header("📈 Campaign Performance Analysis")
        
        if analyzer.daily_campaigns is not None and len(analyzer.daily_campaigns) > 0:
            df = analyzer.calculate_roi_metrics(analyzer.daily_campaigns)
            
            # Filters
            col1, col2 = st.columns(2)
            
            with col1:
                if 'Type' in df.columns:
                    campaign_types = ['All'] + list(df['Type'].unique())
                    selected_type = st.selectbox("Campaign Type", campaign_types)
                    if selected_type != 'All':
                        df = df[df['Type'] == selected_type]
            
            with col2:
                if 'State' in df.columns:
                    states = ['All'] + list(df['State'].unique())
                    selected_state = st.selectbox("Status", states)
                    if selected_state != 'All':
                        df = df[df['State'] == selected_state]
            
            # Key metrics
            st.subheader("Performance Metrics")
            
            spend_col = 'Total cost (converted)' if 'Total cost (converted)' in df.columns else 'Total cost'
            sales_col = 'Sales (converted)' if 'Sales (converted)' in df.columns else 'Sales'
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                fig = px.histogram(
                    df,
                    x='ROAS',
                    nbins=30,
                    title='ROAS Distribution',
                    color_discrete_sequence=['#FF9900']
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                fig = px.box(
                    df,
                    y=spend_col,
                    title='Spend Distribution',
                    color_discrete_sequence=['#146EB4']
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with col3:
                fig = px.box(
                    df,
                    y=sales_col,
                    title='Sales Distribution',
                    color_discrete_sequence=['#28a745']
                )
                st.plotly_chart(fig, use_container_width=True)
            
            # Campaign comparison
            st.subheader("Campaign Comparison")
            
            comparison_df = df[['Campaign name', 'ROAS', spend_col, sales_col, 'Clicks', 'Purchases']].sort_values('ROAS', ascending=False).head(20)
            st.dataframe(comparison_df, use_container_width=True)
            
        else:
            st.info("👈 Please upload campaign data to see analysis")
    
    # Tab 4: Detailed Data
    with tab4:
        st.header("🔍 Detailed Data View")
        
        data_view = st.selectbox(
            "Select Data to View",
            ["Daily Campaigns", "Daily Targets", "Daily Inventory", "Weekly Campaigns", "Weekly CPR", "Weekly Repeat Purchase"]
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
    
    # Tab 5: Insights
    with tab5:
        st.header("💡 Key Insights & Trends")
        
        if analyzer.daily_campaigns is not None and len(analyzer.daily_campaigns) > 0:
            df = analyzer.calculate_roi_metrics(analyzer.daily_campaigns)
            
            # Top insights
            st.subheader("Quick Insights")
            
            spend_col = 'Total cost (converted)' if 'Total cost (converted)' in df.columns else 'Total cost'
            sales_col = 'Sales (converted)' if 'Sales (converted)' in df.columns else 'Sales'
            
            # Best performing
            best_campaign = df.nlargest(1, 'ROAS').iloc[0]
            st.success(f"🏆 **Best ROAS:** {best_campaign['Campaign name']} with {best_campaign['ROAS']:.2f}x ROAS")
            
            # Highest spend
            highest_spend = df.nlargest(1, spend_col).iloc[0]
            st.info(f"💰 **Highest Spend:** {highest_spend['Campaign name']} spent ₹{highest_spend[spend_col]:,.0f}")
            
            # Most sales
            most_sales = df.nlargest(1, sales_col).iloc[0]
            st.info(f"📈 **Highest Sales:** {most_sales['Campaign name']} generated ₹{most_sales[sales_col]:,.0f}")
            
            # Efficiency analysis
            st.markdown("---")
            st.subheader("Efficiency Analysis")
            
            # ROAS vs Spend quadrant
            median_roas = df['ROAS'].median()
            median_spend = df[spend_col].median()
            
            df['Quadrant'] = df.apply(
                lambda row: 'Star' if row['ROAS'] >= median_roas and row[spend_col] >= median_spend
                else 'Question Mark' if row['ROAS'] >= median_roas and row[spend_col] < median_spend
                else 'Cash Cow' if row['ROAS'] < median_roas and row[spend_col] >= median_spend
                else 'Dog',
                axis=1
            )
            
            fig = px.scatter(
                df,
                x=spend_col,
                y='ROAS',
                color='Quadrant',
                size=sales_col,
                hover_data=['Campaign name'],
                title='Campaign Performance Matrix',
                color_discrete_map={
                    'Star': '#28a745',
                    'Question Mark': '#ffc107',
                    'Cash Cow': '#17a2b8',
                    'Dog': '#dc3545'
                }
            )
            
            fig.add_hline(y=median_roas, line_dash="dash", line_color="gray", annotation_text="Median ROAS")
            fig.add_vline(x=median_spend, line_dash="dash", line_color="gray", annotation_text="Median Spend")
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Quadrant breakdown
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                stars = len(df[df['Quadrant'] == 'Star'])
                st.metric("⭐ Stars", stars, help="High ROAS, High Spend - Scale these!")
            
            with col2:
                qm = len(df[df['Quadrant'] == 'Question Mark'])
                st.metric("❓ Question Marks", qm, help="High ROAS, Low Spend - Invest more")
            
            with col3:
                cc = len(df[df['Quadrant'] == 'Cash Cow'])
                st.metric("🐮 Cash Cows", cc, help="Low ROAS, High Spend - Optimize or reduce")
            
            with col4:
                dogs = len(df[df['Quadrant'] == 'Dog'])
                st.metric("🐕 Dogs", dogs, help="Low ROAS, Low Spend - Consider closing")
            
        else:
            st.info("👈 Please upload campaign data to see insights")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "Amazon Marketing ROI Analyzer | Built for MamaNourish Performance Analysis"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
