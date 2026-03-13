"""
Amazon Marketing ROI Analyzer - Improved Version
Focus: Actionable insights, better UX, tabular data, product mapping
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from pathlib import Path
import json

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
    .stDataFrame {
        font-size: 0.9rem;
    }
    .action-expand {
        background-color: #d4edda;
        padding: 0.5rem;
        border-radius: 0.3rem;
        font-weight: bold;
        color: #155724;
    }
    .action-optimize {
        background-color: #cfe2ff;
        padding: 0.5rem;
        border-radius: 0.3rem;
        font-weight: bold;
        color: #084298;
    }
    .action-pause {
        background-color: #fff3cd;
        padding: 0.5rem;
        border-radius: 0.3rem;
        font-weight: bold;
        color: #856404;
    }
    .action-close {
        background-color: #f8d7da;
        padding: 0.5rem;
        border-radius: 0.3rem;
        font-weight: bold;
        color: #721c24;
    }
    </style>
""", unsafe_allow_html=True)

# Historical data directory
HISTORY_DIR = Path("marketing_history")
HISTORY_DIR.mkdir(exist_ok=True)

# Product mapping file
PRODUCT_MAPPING_FILE = HISTORY_DIR / "product_asin_mapping.json"


class ProductMappingManager:
    """Manage product to ASIN mapping"""
    
    def __init__(self):
        self.mapping = self.load_mapping()
    
    def load_mapping(self):
        """Load product-ASIN mapping from file"""
        if PRODUCT_MAPPING_FILE.exists():
            with open(PRODUCT_MAPPING_FILE, 'r') as f:
                return json.load(f)
        return {}
    
    def save_mapping(self, mapping):
        """Save product-ASIN mapping to file"""
        with open(PRODUCT_MAPPING_FILE, 'w') as f:
            json.dump(mapping, f, indent=2)
        self.mapping = mapping
    
    def add_mapping(self, product_name, asin):
        """Add a product-ASIN mapping"""
        self.mapping[product_name] = asin
        self.save_mapping(self.mapping)
    
    def get_asin(self, product_name):
        """Get ASIN for a product"""
        return self.mapping.get(product_name, "Unknown")
    
    def get_product(self, asin):
        """Get product name for an ASIN"""
        for product, mapped_asin in self.mapping.items():
            if mapped_asin == asin:
                return product
        return "Unknown"


class HistoricalDataManager:
    """Manage historical data storage"""
    
    def __init__(self):
        self.history_file = HISTORY_DIR / "campaign_history.csv"
        self.keyword_history_file = HISTORY_DIR / "keyword_history.csv"
    
    def save_campaign_snapshot(self, df, snapshot_date=None):
        """Save campaign performance snapshot"""
        if df is None or len(df) == 0:
            return
        
        if snapshot_date is None:
            snapshot_date = datetime.now().strftime('%Y-%m-%d')
        
        df_snapshot = df.copy()
        df_snapshot['snapshot_date'] = snapshot_date
        
        if self.history_file.exists():
            history = pd.read_csv(self.history_file)
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
    
    def get_campaign_history(self):
        """Get all campaign history"""
        if self.history_file.exists():
            return pd.read_csv(self.history_file)
        return None
    
    def get_keyword_history(self):
        """Get all keyword history"""
        if self.keyword_history_file.exists():
            return pd.read_csv(self.keyword_history_file)
        return None


class AmazonROIAnalyzer:
    """Main analyzer class"""
    
    def __init__(self):
        self.weekly_sales = None
        self.weekly_campaigns = None
        self.weekly_cpr = None
        self.weekly_repeat = None
        self.daily_campaigns = None
        self.daily_targets = None
        self.daily_inventory = None
        self.history_manager = HistoricalDataManager()
        self.product_mapper = ProductMappingManager()
        
    def parse_currency(self, value):
        """Parse currency to float"""
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
        """Parse percentage to float"""
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
            
            numeric_cols = ['Clicks', 'CTR', 'Total cost (converted)', 'Total cost', 
                           'CPC (converted)', 'Purchases', 'Sales (converted)', 'Sales', 'ROAS']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(self.parse_currency)
            
            if self.daily_campaigns is not None:
                df = pd.concat([self.daily_campaigns, df])
                df = df.drop_duplicates(subset=['Campaign name'], keep='last')
            
            self.daily_campaigns = df
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
            self.history_manager.save_keyword_snapshot(df)
            
            st.success(f"✅ Loaded Daily Targets: {len(df)} targets")
            return True
        except Exception as e:
            st.error(f"Error loading daily targets: {str(e)}")
            return False
    
    def load_daily_inventory(self, file):
        """Load daily inventory report"""
        try:
            df = pd.read_csv(file)
            
            numeric_cols = ['Impressions', 'Clicks', 'CTR', 'Total cost', 
                           'Purchases', 'Sales', 'ROAS', 'Purchase rate']
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
    
    def get_summary_stats(self):
        """Get summary statistics"""
        stats = {}
        
        if self.daily_campaigns is not None and len(self.daily_campaigns) > 0:
            df = self.daily_campaigns
            
            stats['total_campaigns'] = len(df)
            stats['active_campaigns'] = len(df[df['State'] == 'ENABLED']) if 'State' in df.columns else len(df)
            
            # Handle both column naming conventions
            spend_col = 'Total cost (converted)' if 'Total cost (converted)' in df.columns else 'Total cost'
            sales_col = 'Sales (converted)' if 'Sales (converted)' in df.columns else 'Sales'
            
            stats['total_spend'] = df[spend_col].sum() if spend_col in df.columns else 0
            stats['total_sales'] = df[sales_col].sum() if sales_col in df.columns else 0
            stats['avg_roas'] = df['ROAS'].mean() if 'ROAS' in df.columns else 0
            stats['total_purchases'] = df['Purchases'].sum() if 'Purchases' in df.columns else 0
            stats['total_clicks'] = df['Clicks'].sum() if 'Clicks' in df.columns else 0
            stats['avg_ctr'] = df['CTR'].mean() if 'CTR' in df.columns else 0
        
        return stats
    
    def generate_campaign_recommendations(self):
        """Generate campaign recommendations in tabular format"""
        try:
            if self.daily_campaigns is None or len(self.daily_campaigns) == 0:
                return pd.DataFrame()
            
            df = self.daily_campaigns.copy()
            
            # Get column names
            spend_col = 'Total cost (converted)' if 'Total cost (converted)' in df.columns else 'Total cost'
            sales_col = 'Sales (converted)' if 'Sales (converted)' in df.columns else 'Sales'
            
            # Filter campaigns with minimum spend (only if column exists)
            if spend_col in df.columns:
                df = df[df[spend_col] > 100]
            
            # Check if we have any data left
            if len(df) == 0:
                return pd.DataFrame()
            
            # Calculate metrics
            if 'ROAS' in df.columns and 'Campaign name' in df.columns:
                # Determine action
                def get_action(row):
                    roas = row['ROAS']
                    purchases = row.get('Purchases', 0)
                    
                    if roas >= 3.0 and purchases >= 5:
                        return 'EXPAND', 'Increase budget by 30-50%', 1
                    elif roas >= 2.0:
                        return 'OPTIMIZE', 'Review keywords and adjust bids', 2
                    elif roas >= 1.0:
                        return 'PAUSE', 'Pause and analyze', 3
                    else:
                        return 'CLOSE', 'Close immediately', 4
                
                df[['Action', 'Recommendation', 'Priority']] = df.apply(
                    lambda row: pd.Series(get_action(row)), axis=1
                )
                
                # Select relevant columns including Priority for sorting
                result_cols = ['Campaign name', 'Action', 'ROAS', sales_col, spend_col, 
                              'Purchases', 'Clicks', 'Recommendation', 'Priority']
                
                # Filter to only existing columns
                result_cols = [col for col in result_cols if col in df.columns]
                
                result = df[result_cols].copy()
                
                # Sort by priority and ROAS BEFORE renaming
                sort_cols = []
                if 'Priority' in result.columns:
                    sort_cols.append('Priority')
                if 'ROAS' in result.columns:
                    sort_cols.append('ROAS')
                
                if sort_cols:
                    ascending = [True] * len(sort_cols)
                    ascending[-1] = False  # Last column (ROAS) descending
                    result = result.sort_values(sort_cols, ascending=ascending)
                
                # Rename for display
                result.columns = [col.replace(' (converted)', '').replace('_', ' ').title() 
                                for col in result.columns]
                
                return result
            
            return pd.DataFrame()
            
        except Exception as e:
            # Log error and return empty dataframe
            import traceback
            print(f"Error in generate_campaign_recommendations: {str(e)}")
            print(traceback.format_exc())
            return pd.DataFrame()
    
    def generate_keyword_recommendations(self):
        """Generate keyword recommendations in tabular format"""
        try:
            if self.daily_targets is None or len(self.daily_targets) == 0:
                return pd.DataFrame()
            
            df = self.daily_targets.copy()
            
            # Filter keywords with minimum spend (only if column exists)
            if 'Spend' in df.columns:
                df = df[df['Spend'] > 50]
            
            # Check if we have any data left
            if len(df) == 0:
                return pd.DataFrame()
            
            # Ensure required columns exist
            if 'Target' not in df.columns or 'ROAS' not in df.columns:
                return pd.DataFrame()
            
            # Determine action
            def get_keyword_action(row):
                roas = row.get('ROAS', 0)
                spend = row.get('Spend', 0)
                
                if roas >= 3.0:
                    return 'INCREASE BID', 'High performer - bid higher', 1
                elif roas >= 2.0:
                    return 'MAINTAIN', 'Stable - keep current bid', 2
                elif roas >= 1.0:
                    return 'REDUCE BID', 'Underperforming - lower bid', 3
                else:
                    return 'PAUSE', 'Money loser - pause keyword', 4
            
            df[['Action', 'Recommendation', 'Priority']] = df.apply(
                lambda row: pd.Series(get_keyword_action(row)), axis=1
            )
            
            # Select relevant columns including Priority for sorting
            result_cols = ['Target', 'Campaign', 'Action', 'ROAS', 'Sales', 'Spend', 
                          'Orders', 'Clicks', 'CTR', 'Recommendation', 'Priority']
            result_cols = [col for col in result_cols if col in df.columns]
            
            result = df[result_cols].copy()
            
            # Sort by priority and ROAS BEFORE renaming
            sort_cols = []
            if 'Priority' in result.columns:
                sort_cols.append('Priority')
            if 'ROAS' in result.columns:
                sort_cols.append('ROAS')
            
            if sort_cols:
                ascending = [True] * len(sort_cols)
                ascending[-1] = False  # Last column (ROAS) descending
                result = result.sort_values(sort_cols, ascending=ascending)
            
            # Rename for display
            result.columns = [col.replace('_', ' ').title() for col in result.columns]
            
            return result
            
        except Exception as e:
            # Log error and return empty dataframe
            import traceback
            print(f"Error in generate_keyword_recommendations: {str(e)}")
            print(traceback.format_exc())
            return pd.DataFrame()
    
    def generate_product_recommendations(self):
        """Generate product recommendations in tabular format"""
        try:
            results = []
            
            # From repeat purchase data
            if self.weekly_repeat is not None and len(self.weekly_repeat) > 0:
                df = self.weekly_repeat.copy()
                
                if 'Product Title' in df.columns:
                    for product in df['Product Title'].unique():
                        # Skip if product is NaN
                        if pd.isna(product):
                            continue
                            
                        product_data = df[df['Product Title'] == product]
                        
                        # Skip if no data for this product
                        if len(product_data) == 0:
                            continue
                        
                        sales = product_data['Repeat Ordered Product Sales: Sales'].sum() if 'Repeat Ordered Product Sales: Sales' in product_data.columns else 0
                        repeat_rate = product_data['Repeat Customer Share: % Share of Total Customers'].mean() if 'Repeat Customer Share: % Share of Total Customers' in product_data.columns else 0
                        
                        # Get ASIN safely
                        if 'ASIN' in product_data.columns and len(product_data['ASIN']) > 0:
                            asin = product_data['ASIN'].iloc[0]
                        else:
                            asin = self.product_mapper.get_asin(product)
                        
                        # Determine action
                        if repeat_rate > 5 and sales > 10000:
                            action = 'SCALE UP'
                            recommendation = 'Increase ad spend 30-50%'
                            priority = 1
                        elif repeat_rate > 2 and sales > 5000:
                            action = 'MAINTAIN'
                            recommendation = 'Keep current spend level'
                            priority = 2
                        elif repeat_rate < 1 and sales < 3000:
                            action = 'DISCONTINUE'
                            recommendation = 'Stop or major restructure'
                            priority = 4
                        else:
                            action = 'REDUCE'
                            recommendation = 'Cut spend by 20-30%'
                            priority = 3
                        
                        results.append({
                            'Product': product,
                            'ASIN': asin,
                            'Action': action,
                            'Repeat Rate %': round(repeat_rate, 2),
                            'Sales': round(sales, 0),
                            'Recommendation': recommendation,
                            'Priority': priority
                        })
            
            if results:
                result_df = pd.DataFrame(results)
                result_df = result_df.sort_values(['Priority', 'Sales'], ascending=[True, False])
                return result_df
            
            return pd.DataFrame()
            
        except Exception as e:
            # Log error and return empty dataframe
            import traceback
            print(f"Error in generate_product_recommendations: {str(e)}")
            print(traceback.format_exc())
            return pd.DataFrame()


def main():
    """Main Streamlit application"""
    
    st.markdown('<p class="main-header">🎯 Amazon Marketing ROI Analyzer</p>', unsafe_allow_html=True)
    
    # Initialize
    if 'analyzer' not in st.session_state:
        st.session_state.analyzer = AmazonROIAnalyzer()
    
    analyzer = st.session_state.analyzer
    
    # Sidebar
    with st.sidebar:
        st.header("📁 Data Upload")
        
        st.subheader("Weekly Reports")
        weekly_sales = st.file_uploader("Weekly Sales (Excel)", type=['xlsx'], key='ws')
        weekly_campaigns = st.file_uploader("Weekly Campaigns (Excel)", type=['xlsx'], key='wc')
        weekly_cpr = st.file_uploader("Weekly CPR (Excel)", type=['xlsx'], key='wcpr')
        weekly_repeat = st.file_uploader("Weekly Repeat Purchase (Excel)", type=['xlsx'], key='wr')
        
        st.subheader("Daily Reports")
        daily_campaigns = st.file_uploader("Daily Campaigns (CSV)", type=['csv'], key='dc')
        daily_targets = st.file_uploader("Daily Targets (CSV)", type=['csv'], key='dt')
        daily_inventory = st.file_uploader("Daily Inventory (CSV)", type=['csv'], key='di')
        
        if st.button("🔄 Load All Data", type="primary"):
            with st.spinner("Loading..."):
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
                
                st.success("✅ Data loaded!")
        
        st.markdown("---")
        
        # Date filter
        st.subheader("⏰ Time Filter")
        date_range = st.selectbox(
            "Focus on campaigns from:",
            ["All Time", "Last 7 Days", "Last 14 Days", "Last 30 Days", "Custom Range"]
        )
        
        if date_range == "Custom Range":
            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input("Start Date")
            with col2:
                end_date = st.date_input("End Date")
    
    # Main tabs - SIMPLIFIED AND ACTIONABLE
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Dashboard",
        "🎯 Campaign Actions",
        "🔑 Keyword Actions", 
        "📦 Product Actions",
        "⚙️ Settings"
    ])
    
    # Tab 1: Dashboard
    with tab1:
        st.header("Performance Summary")
        
        stats = analyzer.get_summary_stats()
        
        if stats:
            # Top metrics
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.metric("Total Campaigns", f"{stats.get('total_campaigns', 0)}")
            with col2:
                st.metric("Total Spend", f"₹{stats.get('total_spend', 0):,.0f}")
            with col3:
                st.metric("Total Sales", f"₹{stats.get('total_sales', 0):,.0f}")
            with col4:
                roi = ((stats.get('total_sales', 0) / max(stats.get('total_spend', 1), 1) - 1) * 100)
                st.metric("ROI", f"{roi:.1f}%")
            with col5:
                st.metric("Avg ROAS", f"{stats.get('avg_roas', 0):.2f}x")
            
            st.markdown("---")
            
            # Quick action summary
            recs = analyzer.generate_campaign_recommendations()
            if not recs.empty and 'Action' in recs.columns:
                st.subheader("📋 Action Summary")
                
                action_summary = recs['Action'].value_counts()
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    expand_count = action_summary.get('EXPAND', 0)
                    st.metric("🚀 Campaigns to EXPAND", expand_count)
                    if expand_count > 0:
                        st.caption("Scale up budget")
                
                with col2:
                    optimize_count = action_summary.get('OPTIMIZE', 0)
                    st.metric("⚙️ Campaigns to OPTIMIZE", optimize_count)
                    if optimize_count > 0:
                        st.caption("Review & improve")
                
                with col3:
                    pause_count = action_summary.get('PAUSE', 0)
                    st.metric("⏸️ Campaigns to PAUSE", pause_count)
                    if pause_count > 0:
                        st.caption("Stop & analyze")
                
                with col4:
                    close_count = action_summary.get('CLOSE', 0)
                    st.metric("❌ Campaigns to CLOSE", close_count)
                    if close_count > 0:
                        st.caption("Shut down now")
        else:
            st.info("👈 Upload data files to see dashboard")
    
    # Tab 2: Campaign Actions
    with tab2:
        st.header("🎯 Campaign Recommendations")
        
        recs = analyzer.generate_campaign_recommendations()
        
        if not recs.empty:
            # Summary at top
            st.subheader("📊 Summary")
            action_counts = recs['Action'].value_counts()
            
            summary_cols = st.columns(4)
            
            actions_config = [
                ('EXPAND', '🚀', '#d4edda'),
                ('OPTIMIZE', '⚙️', '#cfe2ff'),
                ('PAUSE', '⏸️', '#fff3cd'),
                ('CLOSE', '❌', '#f8d7da')
            ]
            
            for idx, (action, emoji, color) in enumerate(actions_config):
                with summary_cols[idx]:
                    count = action_counts.get(action, 0)
                    st.markdown(f"""
                    <div style="background-color: {color}; padding: 1rem; border-radius: 0.5rem; text-align: center;">
                        <h3>{emoji} {action}</h3>
                        <h1>{count}</h1>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Filters
            col1, col2, col3 = st.columns(3)
            
            with col1:
                action_filter = st.multiselect(
                    "Filter by Action",
                    options=['EXPAND', 'OPTIMIZE', 'PAUSE', 'CLOSE'],
                    default=['EXPAND', 'OPTIMIZE', 'PAUSE', 'CLOSE']
                )
            
            with col2:
                min_roas = st.number_input("Min ROAS", min_value=0.0, value=0.0, step=0.1)
            
            with col3:
                min_spend = st.number_input("Min Spend (₹)", min_value=0, value=0, step=100)
            
            # Filter data
            filtered_recs = recs.copy()
            if action_filter:
                filtered_recs = filtered_recs[filtered_recs['Action'].isin(action_filter)]
            if 'Roas' in filtered_recs.columns:
                filtered_recs = filtered_recs[filtered_recs['Roas'] >= min_roas]
            if 'Total Cost' in filtered_recs.columns:
                filtered_recs = filtered_recs[filtered_recs['Total Cost'] >= min_spend]
            elif 'Spend' in filtered_recs.columns:
                filtered_recs = filtered_recs[filtered_recs['Spend'] >= min_spend]
            
            st.subheader(f"📋 All Recommendations ({len(filtered_recs)} campaigns)")
            
            # Display table with formatting
            st.dataframe(
                filtered_recs.drop('Priority', axis=1, errors='ignore'),
                use_container_width=True,
                height=600
            )
            
            # Export option
            csv = filtered_recs.to_csv(index=False)
            st.download_button(
                "📥 Download Recommendations (CSV)",
                csv,
                "campaign_recommendations.csv",
                "text/csv"
            )
        else:
            st.info("No campaign data available. Upload Daily Campaigns report.")
    
    # Tab 3: Keyword Actions
    with tab3:
        st.header("🔑 Keyword Recommendations")
        
        keyword_recs = analyzer.generate_keyword_recommendations()
        
        if not keyword_recs.empty:
            # Summary
            st.subheader("📊 Summary")
            action_counts = keyword_recs['Action'].value_counts()
            
            summary_cols = st.columns(4)
            
            keyword_actions = [
                ('INCREASE BID', '🔼', '#d4edda'),
                ('MAINTAIN', '➡️', '#cfe2ff'),
                ('REDUCE BID', '🔽', '#fff3cd'),
                ('PAUSE', '⏸️', '#f8d7da')
            ]
            
            for idx, (action, emoji, color) in enumerate(keyword_actions):
                with summary_cols[idx]:
                    count = action_counts.get(action, 0)
                    st.markdown(f"""
                    <div style="background-color: {color}; padding: 1rem; border-radius: 0.5rem; text-align: center;">
                        <h4>{emoji} {action}</h4>
                        <h2>{count}</h2>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Filters
            col1, col2, col3 = st.columns(3)
            
            with col1:
                keyword_action_filter = st.multiselect(
                    "Filter by Action",
                    options=['INCREASE BID', 'MAINTAIN', 'REDUCE BID', 'PAUSE'],
                    default=['INCREASE BID', 'MAINTAIN', 'REDUCE BID', 'PAUSE']
                )
            
            with col2:
                keyword_min_roas = st.number_input("Min ROAS", min_value=0.0, value=0.0, step=0.1, key='kw_roas')
            
            with col3:
                keyword_min_spend = st.number_input("Min Spend (₹)", min_value=0, value=0, step=50, key='kw_spend')
            
            # Filter
            filtered_keywords = keyword_recs.copy()
            if keyword_action_filter:
                filtered_keywords = filtered_keywords[filtered_keywords['Action'].isin(keyword_action_filter)]
            if 'Roas' in filtered_keywords.columns:
                filtered_keywords = filtered_keywords[filtered_keywords['Roas'] >= keyword_min_roas]
            if 'Spend' in filtered_keywords.columns:
                filtered_keywords = filtered_keywords[filtered_keywords['Spend'] >= keyword_min_spend]
            
            st.subheader(f"📋 All Keywords ({len(filtered_keywords)} keywords)")
            
            # Display table
            st.dataframe(
                filtered_keywords.drop('Priority', axis=1, errors='ignore'),
                use_container_width=True,
                height=600
            )
            
            # Export
            csv = filtered_keywords.to_csv(index=False)
            st.download_button(
                "📥 Download Keyword Recommendations (CSV)",
                csv,
                "keyword_recommendations.csv",
                "text/csv"
            )
        else:
            st.info("No keyword data available. Upload Daily Targets report.")
    
    # Tab 4: Product Actions
    with tab4:
        st.header("📦 Product Recommendations")
        
        product_recs = analyzer.generate_product_recommendations()
        
        if not product_recs.empty:
            # Summary
            st.subheader("📊 Summary")
            action_counts = product_recs['Action'].value_counts()
            
            summary_cols = st.columns(4)
            
            product_actions = [
                ('SCALE UP', '🚀', '#d4edda'),
                ('MAINTAIN', '✅', '#cfe2ff'),
                ('REDUCE', '⚠️', '#fff3cd'),
                ('DISCONTINUE', '❌', '#f8d7da')
            ]
            
            for idx, (action, emoji, color) in enumerate(product_actions):
                with summary_cols[idx]:
                    count = action_counts.get(action, 0)
                    st.markdown(f"""
                    <div style="background-color: {color}; padding: 1rem; border-radius: 0.5rem; text-align: center;">
                        <h4>{emoji} {action}</h4>
                        <h2>{count}</h2>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Filters
            col1, col2, col3 = st.columns(3)
            
            with col1:
                product_action_filter = st.multiselect(
                    "Filter by Action",
                    options=['SCALE UP', 'MAINTAIN', 'REDUCE', 'DISCONTINUE'],
                    default=['SCALE UP', 'MAINTAIN', 'REDUCE', 'DISCONTINUE']
                )
            
            with col2:
                sort_by = st.selectbox(
                    "Sort By",
                    options=['Sales', 'Repeat Rate %', 'Action'],
                    index=0
                )
            
            with col3:
                sort_order = st.selectbox(
                    "Order",
                    options=['Descending', 'Ascending'],
                    index=0
                )
            
            # Filter and sort
            filtered_products = product_recs.copy()
            if product_action_filter:
                filtered_products = filtered_products[filtered_products['Action'].isin(product_action_filter)]
            
            if sort_by:
                ascending = sort_order == 'Ascending'
                filtered_products = filtered_products.sort_values(sort_by, ascending=ascending)
            
            st.subheader(f"📋 All Products ({len(filtered_products)} products)")
            
            # Display table
            st.dataframe(
                filtered_products.drop('Priority', axis=1, errors='ignore'),
                use_container_width=True,
                height=600
            )
            
            # Export
            csv = filtered_products.to_csv(index=False)
            st.download_button(
                "📥 Download Product Recommendations (CSV)",
                csv,
                "product_recommendations.csv",
                "text/csv"
            )
            
            # Show detailed view with campaign data
            st.markdown("---")
            st.subheader("🔍 Product Campaign Details")
            
            if analyzer.daily_campaigns is not None and not filtered_products.empty:
                selected_product = st.selectbox(
                    "Select Product to View Campaign Details",
                    options=filtered_products['Product'].tolist()
                )
                
                if selected_product:
                    # Find campaigns for this product (simplified - would need better mapping)
                    st.info(f"Showing campaign data for: **{selected_product}**")
                    st.caption("Note: Full product-campaign mapping requires ASIN linking. Configure in Settings tab.")
        else:
            st.info("No product data available. Upload Weekly Repeat Purchase report.")
    
    # Tab 5: Settings
    with tab5:
        st.header("⚙️ Settings & Configuration")
        
        st.subheader("🔗 Product-ASIN Mapping")
        st.markdown("Map your product names to ASINs for better cross-report analysis.")
        
        # Show current mappings
        current_mappings = analyzer.product_mapper.mapping
        
        if current_mappings:
            st.markdown("**Current Mappings:**")
            mapping_df = pd.DataFrame([
                {'Product Name': k, 'ASIN': v}
                for k, v in current_mappings.items()
            ])
            st.dataframe(mapping_df, use_container_width=True)
        
        st.markdown("---")
        st.markdown("**Add New Mapping:**")
        
        col1, col2, col3 = st.columns([2, 2, 1])
        
        with col1:
            new_product = st.text_input("Product Name", key='new_product')
        
        with col2:
            new_asin = st.text_input("ASIN", key='new_asin')
        
        with col3:
            st.write("")  # Spacing
            st.write("")
            if st.button("➕ Add Mapping"):
                if new_product and new_asin:
                    analyzer.product_mapper.add_mapping(new_product, new_asin)
                    st.success(f"✅ Mapped '{new_product}' to '{new_asin}'")
                    st.rerun()
                else:
                    st.error("Please enter both product name and ASIN")
        
        st.markdown("---")
        
        # Import mappings from uploaded data
        st.subheader("📥 Auto-Import Mappings")
        
        if analyzer.weekly_repeat is not None:
            if st.button("Import from Weekly Repeat Purchase Data"):
                if 'ASIN' in analyzer.weekly_repeat.columns and 'Product Title' in analyzer.weekly_repeat.columns:
                    count = 0
                    for _, row in analyzer.weekly_repeat.iterrows():
                        if pd.notna(row['Product Title']) and pd.notna(row['ASIN']):
                            analyzer.product_mapper.add_mapping(row['Product Title'], row['ASIN'])
                            count += 1
                    st.success(f"✅ Imported {count} product-ASIN mappings!")
                    st.rerun()
                else:
                    st.error("Required columns not found in data")
        else:
            st.info("Upload Weekly Repeat Purchase report first to auto-import mappings")
        
        st.markdown("---")
        
        # Threshold settings
        st.subheader("🎯 Recommendation Thresholds")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Campaign Thresholds:**")
            expand_roas = st.number_input("EXPAND ROAS threshold", value=3.0, step=0.1)
            optimize_roas = st.number_input("OPTIMIZE ROAS threshold", value=2.0, step=0.1)
            pause_roas = st.number_input("PAUSE ROAS threshold", value=1.0, step=0.1)
        
        with col2:
            st.markdown("**Minimum Spend Filters:**")
            min_campaign_spend = st.number_input("Min campaign spend (₹)", value=100, step=50)
            min_keyword_spend = st.number_input("Min keyword spend (₹)", value=50, step=25)
        
        st.markdown("---")
        
        # Data management
        st.subheader("🗂️ Data Management")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📊 View Historical Data Stats"):
                campaign_history = analyzer.history_manager.get_campaign_history()
                keyword_history = analyzer.history_manager.get_keyword_history()
                
                if campaign_history is not None:
                    st.metric("Campaign Snapshots", len(campaign_history['snapshot_date'].unique() if 'snapshot_date' in campaign_history.columns else []))
                
                if keyword_history is not None:
                    st.metric("Keyword Snapshots", len(keyword_history['snapshot_date'].unique() if 'snapshot_date' in keyword_history.columns else []))
        
        with col2:
            if st.button("🗑️ Clear Historical Data", type="secondary"):
                if st.checkbox("Confirm deletion"):
                    # Would implement clear logic here
                    st.warning("This would clear all historical data")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "Amazon Marketing ROI Analyzer v2.0 | Actionable Insights Edition"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
