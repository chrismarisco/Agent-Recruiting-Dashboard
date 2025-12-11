import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import re
from io import StringIO

# Configure page
st.set_page_config(
    page_title="Bluesky Realty - Agent Recruiting Dashboard",
    page_icon="🏠",
    layout="wide"
)

#---------------------------
# Helper Functions
#---------------------------
def dollar_to_numeric(x):
    """Convert dollar strings to numeric values"""
    if pd.isna(x):
        return np.nan
    if isinstance(x, str):
        return float(re.sub(r'[$,]', '', x))
    return float(x)

def clean_agent_data(raw_df):
    """Clean and prepare agent data"""
    df = raw_df.copy()
    
    # Convert dollar columns
    for col in ['TotalVolume', 'ListSideVolume', 'SellSideVolume']:
        if col in df.columns:
            df[col] = df[col].apply(dollar_to_numeric)
    
    # Convert numeric columns
    for col in ['TotalCount', 'DaysOnMarket']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Convert dates and calculate recency
    if 'MovedDate' in df.columns:
        df['MovedDate'] = pd.to_datetime(df['MovedDate'], format='%m/%d/%Y', errors='coerce')
        df['RecencyDays'] = (pd.Timestamp.now() - df['MovedDate']).dt.days
    
    # Clean names
    df['FirstName'] = df['FirstName'].fillna('Unnamed').replace('', 'Unnamed')
    df['LastName'] = df['LastName'].fillna('Agent').replace('', 'Agent')
    
    # Clean city names
    if 'City' in df.columns:
        df['City_clean'] = df['City'].str.strip().str.title()
    
    return df

def safe_rescale(series):
    """Safely rescale a series to 0-1 range"""
    if series.empty or series.isna().all():
        return pd.Series([0.5] * len(series), index=series.index)
    
    min_val = series.min()
    max_val = series.max()
    
    if pd.isna(min_val) or pd.isna(max_val) or min_val == max_val:
        return pd.Series([0.5] * len(series), index=series.index)
    
    return (series - min_val) / (max_val - min_val)

def format_currency(x):
    """Format currency values"""
    if pd.isna(x):
        return "—"
    return f"${x:,.0f}"

def format_integer(x):
    """Format integer values"""
    if pd.isna(x):
        return "—"
    return f"{int(x):,}"

@st.cache_data
def load_and_prepare_data():
    """Load and prepare all data"""
    try:
        # Load main data
        raw_data = pd.read_csv("CA Dashboard.csv")
        agents = clean_agent_data(raw_data)
        
        # Load city to county mapping
        try:
            city2county = pd.read_csv("CA_City_to_County_Mapping.csv")
            city2county['City_clean'] = city2county['City'].str.strip().str.title()
            city2county['County'] = city2county['County'].str.strip().str.title()
            city2county = city2county[['City_clean', 'County']].drop_duplicates('City_clean')
        except FileNotFoundError:
            st.warning("City to County mapping file not found. All cities will be marked as 'Unknown'.")
            city2county = pd.DataFrame(columns=['City_clean', 'County'])
        
        # Find unmapped cities
        if 'City_clean' in agents.columns:
            unmapped_cities = agents[['City_clean']].drop_duplicates()
            unmapped_cities = unmapped_cities[~unmapped_cities['City_clean'].isin(city2county['City_clean'])]
            unmapped_cities['County'] = 'Unknown'
            
            # Combine mappings
            city2county_full = pd.concat([city2county, unmapped_cities], ignore_index=True)
            city2county_full = city2county_full.drop_duplicates('City_clean')
            
            # Add county information to agents
            agents = agents.merge(city2county_full, on='City_clean', how='left')
            agents['County'] = agents['County'].fillna('Unknown').replace('', 'Unknown')
        else:
            agents['County'] = 'Unknown'
        
        return agents, city2county_full
        
    except FileNotFoundError:
        st.error("Please upload 'CA Dashboard.csv' file to proceed.")
        return None, None

#---------------------------
# Main Application
#---------------------------
def main():
    st.title("🏠 Bluesky Realty – Agent Recruiting Dashboard")
    
    st.markdown("""
    This interactive tool helps brokerage teams identify promising real estate agents for recruitment.
    Rows with missing fields are kept and scored conservatively so you see more people.
    """)
    
    with st.expander("🛈 How to Use This Tool"):
        st.markdown("""
        - Adjust weights for Volume, DOM, Recency, and Office Changes.
        - Filters allow missing values; missing DOM/Recency are treated as worst allowed for scoring.
        - Use the County filter (includes 'Unknown') to focus markets.
        """)
    
    # Load data
    agents, city2county_full = load_and_prepare_data()
    
    if agents is None:
        st.stop()
    
    # Sidebar controls
    with st.sidebar:
        st.header("Scoring Weights")
        volume_weight = st.slider("Weight: Volume", 0.0, 1.0, 0.30, 0.05)
        dom_weight = st.slider("Weight: DOM (Lower is Better)", 0.0, 1.0, 0.35, 0.05)
        recency_weight = st.slider("Weight: Recency (Recent Better)", 0.0, 1.0, 0.20, 0.05)
        change_weight = st.slider("Weight: Office Changes", 0.0, 1.0, 0.20, 0.05)
        
        st.divider()
        
        st.header("Filters")
        st.caption("Rows with missing fields are kept")
        
        # County filter
        counties = sorted(agents['County'].dropna().unique())
        county_filter = st.multiselect(
            "County (optional):",
            options=counties,
            default=None,
            help="Leave empty to include all counties"
        )
        
        min_volume = st.number_input(
            "Min Total Volume ($):",
            min_value=0,
            value=100000,
            step=100000,
            format="%d"
        )
        
        max_dom = st.number_input(
            "Max Days on Market:",
            min_value=1,
            value=180,
            step=10
        )
        
        max_recency = st.number_input(
            "Max Days Since Move:",
            min_value=1,
            value=999,
            step=50
        )
        
        rescore_button = st.button("🔄 Rescore & Filter Agents", type="primary")
    
    # Process data (runs automatically in Streamlit)
    def process_agents():
        df0 = agents.copy()
        total_rows = len(df0)
        
        # County filter
        if county_filter:
            df1 = df0[df0['County'].isin(county_filter)]
        else:
            df1 = df0.copy()
        
        # Apply filters (keep rows with NAs)
        df2 = df1.copy()
        if 'TotalVolume' in df2.columns:
            df2 = df2[(df2['TotalVolume'].isna()) | (df2['TotalVolume'] >= min_volume)]
        if 'DaysOnMarket' in df2.columns:
            df2 = df2[(df2['DaysOnMarket'].isna()) | (df2['DaysOnMarket'] <= max_dom)]
        if 'RecencyDays' in df2.columns:
            df2 = df2[(df2['RecencyDays'].isna()) | (df2['RecencyDays'] <= max_recency)]
        
        # Impute worst-allowed values for scoring
        df3 = df2.copy()
        df3['vol_for_score'] = df3.get('TotalVolume', 0).fillna(0)
        df3['dom_for_score'] = df3.get('DaysOnMarket', max_dom).fillna(max_dom)
        df3['recency_for_score'] = df3.get('RecencyDays', max_recency).fillna(max_recency)
        
        # Calculate component scores
        df3['Volume_Score'] = safe_rescale(df3['vol_for_score'])
        df3['DOM_Score'] = safe_rescale(-df3['dom_for_score'])
        df3['Recency_Score'] = safe_rescale(-df3['recency_for_score'])
        
        # Office change score
        if 'PreviousOfficeName' in df3.columns:
            has_previous = (~df3['PreviousOfficeName'].isna()) & (df3['PreviousOfficeName'] != '')
            df3['Change_Score'] = safe_rescale(has_previous.astype(int))
        else:
            df3['Change_Score'] = 0
        
        # Calculate final score
        df3['Final_Score'] = (
            volume_weight * df3['Volume_Score'] +
            dom_weight * df3['DOM_Score'] +
            recency_weight * df3['Recency_Score'] +
            change_weight * df3['Change_Score']
        )
        
        # Sort and rank
        df_final = df3.sort_values('Final_Score', ascending=False).reset_index(drop=True)
        df_final['Rank'] = df_final.index + 1
        
        # Select display columns
        display_cols = ['Rank', 'FirstName', 'LastName', 'OfficeName', 'County', 
                       'TotalVolume', 'DaysOnMarket', 'RecencyDays', 'Final_Score']
        display_cols = [col for col in display_cols if col in df_final.columns]
        df_display = df_final[display_cols].copy()
        
        return df_display, df_final, total_rows, len(df1), len(df2)
    
    # Process the data
    df_display, df_full, total_rows, after_county, after_filters = process_agents()
    
    # Row audit
    st.subheader("Row Audit")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Rows in CSV", f"{total_rows:,}")
    with col2:
        st.metric("After County Filter", f"{after_county:,}")
    with col3:
        st.metric("After Thresholds", f"{after_filters:,}")
    with col4:
        st.metric("Rows Shown", f"{len(df_display):,}")
    
    if len(df_display) == 0:
        st.warning("No agents match the current filters.")
        return
    
    # Scatter plot
    st.subheader("Score vs. Total Volume")
    
    if 'TotalVolume' in df_full.columns:
        fig = px.scatter(
            df_full,
            x='TotalVolume',
            y='Final_Score',
            color='Final_Score',
            color_continuous_scale='Viridis',
            hover_data={
                'Rank': True,
                'FirstName': True,
                'LastName': True,
                'County': True,
                'OfficeName': True,
                'TotalVolume': ':$,.0f',
                'DaysOnMarket': ':.0f',
                'RecencyDays': ':.0f',
                'Final_Score': ':.3f'
            },
            title="Recruiting Score vs Total Volume"
        )
        
        fig.update_layout(
            xaxis_title="Total Volume ($)",
            yaxis_title="Recruiting Score",
            height=500
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Agent table
    st.subheader("Ranked Agent Table")
    
    # Format the display dataframe
    df_formatted = df_display.copy()
    if 'TotalVolume' in df_formatted.columns:
        df_formatted['TotalVolume'] = df_formatted['TotalVolume'].apply(format_currency)
    if 'DaysOnMarket' in df_formatted.columns:
        df_formatted['DaysOnMarket'] = df_formatted['DaysOnMarket'].apply(format_integer)
    if 'RecencyDays' in df_formatted.columns:
        df_formatted['RecencyDays'] = df_formatted['RecencyDays'].apply(format_integer)
    if 'Final_Score' in df_formatted.columns:
        df_formatted['Final_Score'] = df_formatted['Final_Score'].round(4)
    
    # Display the table with search and sort functionality
    st.dataframe(
        df_formatted,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Rank": st.column_config.NumberColumn("Rank", format="%d"),
            "FirstName": "First Name",
            "LastName": "Last Name",
            "OfficeName": "Office Name",
            "County": "County",
            "TotalVolume": "Total Volume",
            "DaysOnMarket": "Days on Market", 
            "RecencyDays": "Days Since Move",
            "Final_Score": st.column_config.NumberColumn("Final Score", format="%.4f")
        }
    )
    
    # Download button
    if st.button("📥 Download CSV"):
        csv_buffer = StringIO()
        df_display.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        
        st.download_button(
            label="Download agent scores CSV",
            data=csv_data,
            file_name=f"agent_scores_{date.today().strftime('%Y-%m-%d')}.csv",
            mime="text/csv"
        )

if __name__ == "__main__":
    main()