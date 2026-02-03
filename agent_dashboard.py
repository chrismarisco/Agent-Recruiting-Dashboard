import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import re
from io import StringIO, BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Configure page
st.set_page_config(
    page_title="RealtyMetric Solutions - Agent Recruiting Dashboard",
    page_icon="🏠",
    layout="wide"
)

#---------------------------
# Custom Branding & Styling
#---------------------------
st.markdown("""
<style>
    .main {
        padding-top: 10px;
    }
    .branding-header {
        background: linear-gradient(135deg, #1B2A4A, #2E5090);
        color: white;
        padding: 20px 30px;
        border-radius: 10px;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .branding-header h1 {
        margin: 0;
        font-size: 28px;
        color: white;
    }
    .branding-header p {
        margin: 5px 0 0 0;
        color: #A8C4E0;
        font-size: 14px;
    }
    .branding-logo {
        font-size: 42px;
        font-weight: bold;
        color: white;
        letter-spacing: 1px;
    }
    .branding-logo span {
        color: #5CACEE;
    }
    .footer {
        margin-top: 40px;
        padding: 15px 30px;
        border-top: 1px solid #E0E0E0;
        color: #888;
        font-size: 12px;
        text-align: center;
    }
    .stButton > button {
        background-color: #2E5090;
        color: white;
    }
    .stButton > button:hover {
        background-color: #1B2A4A;
    }
</style>
""", unsafe_allow_html=True)

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
    
    # Remove duplicate agents (keep the row with highest TotalVolume)
    if 'TotalVolume' in df.columns:
        df = df.sort_values('TotalVolume', ascending=False)
    df = df.drop_duplicates(subset=['FirstName', 'LastName', 'OfficeName'], keep='first')
    
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
            city2county_full = pd.DataFrame(columns=['City_clean', 'County'])
        
        return agents, city2county_full
        
    except FileNotFoundError:
        st.error("Please upload 'CA Dashboard.csv' file to proceed.")
        return None, None

#---------------------------
# Authentication
#---------------------------
def check_password():
    """Returns `True` if the user had the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.write("*Please contact your administrator for access.*")
        return False
    elif not st.session_state["password_correct"]:
        # Password incorrect, show input + error
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct
        return True

#---------------------------
# Excel Export
#---------------------------
def export_to_excel(df):
    """Create a professionally formatted Excel export"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Ranked Agents', startrow=1)
        ws = writer.sheets['Ranked Agents']

        # Styles
        title_font = Font(bold=True, size=14, color="FFFFFF")
        title_fill = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2E5090", end_color="2E5090", fill_type="solid")
        header_align = Alignment(horizontal="center")
        alt_row_fill = PatternFill(start_color="F0F7FF", end_color="F0F7FF", fill_type="solid")

        # Title row
        ws['A1'] = 'RealtyMetric Solutions – Agent Recruiting Report'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))

        # Style header row
        for col in range(1, len(df.columns) + 1):
            cell = ws.cell(row=2, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align

        # Auto-size columns
        for i, col_name in enumerate(df.columns, 1):
            max_len = max(len(str(col_name)), df[col_name].astype(str).str.len().max() or 0)
            col_letter = openpyxl.utils.get_column_letter(i + 0)
            ws.column_dimensions[col_letter].width = min(max_len + 4, 30)

        # Format currency column
        if 'TotalVolume' in df.columns:
            tv_col = list(df.columns).index('TotalVolume') + 1
            for row in range(3, len(df) + 3):
                ws.cell(row=row, column=tv_col).number_format = '$#,##0'

        # Format score column
        if 'Final_Score' in df.columns:
            sc_col = list(df.columns).index('Final_Score') + 1
            for row in range(3, len(df) + 3):
                ws.cell(row=row, column=sc_col).number_format = '0.0000'

        # Alternate row colors
        for row in range(3, len(df) + 3, 2):
            for col in range(1, len(df.columns) + 1):
                ws.cell(row=row, column=col).fill = alt_row_fill

    output.seek(0)
    return output

#---------------------------
# Main Application
#---------------------------
def main():
    if not check_password():
        st.stop()

    # --- Branding Header ---
    st.markdown("""
    <div class="branding-header">
        <div>
            <h1>🏠 RealtyMetric Solutions</h1>
            <p>Agent Recruiting Dashboard – California</p>
        </div>
        <div class="branding-logo">RM<span>.</span></div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    This interactive tool helps brokerage teams identify promising real estate agents for recruitment.
    Rows with missing fields are kept and scored conservatively so you see more people.
    """)
    
    with st.expander("🛈 How to Use This Tool"):
        st.markdown("""
        - Adjust weights for Volume, DOM, Recency, and Office Changes.
        - Filters allow missing values; missing DOM/Recency are treated as worst allowed for scoring.
        - Use the County filter (includes 'Unknown') to focus markets.
        - Use the search bar to find specific agents by name.
        """)
    
    # Load data
    agents, city2county_full = load_and_prepare_data()
    
    if agents is None:
        st.stop()
    
    # Sidebar controls
    with st.sidebar:
        st.header("⚖️ Scoring Weights")
        volume_weight = st.slider("Weight: Volume", 0.0, 1.0, 0.30, 0.05)
        dom_weight = st.slider("Weight: DOM (Lower is Better)", 0.0, 1.0, 0.35, 0.05)
        recency_weight = st.slider("Weight: Recency (Recent Better)", 0.0, 1.0, 0.20, 0.05)
        change_weight = st.slider("Weight: Office Changes", 0.0, 1.0, 0.20, 0.05)
        
        st.divider()
        
        st.header("🔍 Filters")
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

        st.divider()

        # Agent Search
        st.header("🔎 Agent Search")
        search_query = st.text_input("Search by agent name...", placeholder="e.g. John Smith")
    
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

        # Agent search filter
        if search_query.strip():
            q = search_query.strip().lower()
            mask = (
                df_display['FirstName'].str.lower().str.contains(q, na=False) |
                df_display['LastName'].str.lower().str.contains(q, na=False) |
                (df_display['FirstName'].str.lower() + ' ' + df_display['LastName'].str.lower()).str.contains(q, na=False)
            )
            df_display = df_display[mask].reset_index(drop=True)
            df_display['Rank'] = df_display.index + 1
        
        return df_display, df_final, total_rows, len(df1), len(df2)
    
    # Process the data
    df_display, df_full, total_rows, after_county, after_filters = process_agents()

    # --- Summary Stats Cards ---
    st.subheader("📊 Summary Overview")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("Total Agents", f"{total_rows:,}")
    with c2:
        st.metric("After Filters", f"{len(df_display):,}")
    with c3:
        avg_vol = df_display['TotalVolume'].mean() if 'TotalVolume' in df_display.columns else 0
        st.metric("Avg Volume", f"${avg_vol:,.0f}" if not pd.isna(avg_vol) else "—")
    with c4:
        med_dom = df_display['DaysOnMarket'].median() if 'DaysOnMarket' in df_display.columns else 0
        st.metric("Median DOM", f"{int(med_dom)}" if not pd.isna(med_dom) else "—")
    with c5:
        top_score = df_display['Final_Score'].max() if not df_display.empty else 0
        st.metric("Top Score", f"{top_score:.4f}" if not pd.isna(top_score) else "—")
    
    if len(df_display) == 0:
        st.warning("No agents match the current filters.")
        return
    
    # Row audit
    st.subheader("🔍 Row Audit")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Rows in CSV", f"{total_rows:,}")
    with col2:
        st.metric("After County Filter", f"{after_county:,}")
    with col3:
        st.metric("After Thresholds", f"{after_filters:,}")
    with col4:
        st.metric("Rows Shown", f"{len(df_display):,}")

    # --- Top 10 Bar Chart ---
    st.subheader("🏆 Top 10 Agents by Recruiting Score")
    top10 = df_display.head(10).copy()
    top10['Agent Name'] = top10['FirstName'] + ' ' + top10['LastName']
    fig_bar = px.bar(
        top10,
        x='Final_Score',
        y='Agent Name',
        orientation='h',
        color='Final_Score',
        color_continuous_scale='Blues',
        text='Final_Score',
        hover_data={
            'OfficeName': True,
            'County': True,
            'TotalVolume': ':$,.0f',
            'Final_Score': ':.4f'
        }
    )
    fig_bar.update_layout(
        yaxis=dict(categoryorder='total ascending', title=None),
        xaxis_title="Recruiting Score",
        height=420,
        showlegend=False,
        coloraxis_showscale=False
    )
    fig_bar.update_traces(texttemplate='%{text:.4f}', textposition='inside')
    st.plotly_chart(fig_bar, use_container_width=True)
    
    # Scatter plot
    st.subheader("📈 Score vs. Total Volume")
    
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
    st.subheader("📋 Ranked Agent Table")
    
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
    
    # --- Export Buttons ---
    st.subheader("📥 Export Data")
    col_csv, col_xlsx = st.columns(2)

    with col_csv:
        csv_buffer = StringIO()
        df_display.to_csv(csv_buffer, index=False)
        st.download_button(
            label="📄 Download CSV",
            data=csv_buffer.getvalue(),
            file_name=f"RealtyMetric_CA_Agents_{date.today().strftime('%Y-%m-%d')}.csv",
            mime="text/csv"
        )

    with col_xlsx:
        xlsx_buffer = export_to_excel(df_display)
        st.download_button(
            label="📊 Download Excel",
            data=xlsx_buffer,
            file_name=f"RealtyMetric_CA_Agents_{date.today().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- Footer ---
    st.markdown("""
    <div class="footer">
        <p>© 2026 RealtyMetric Solutions | Agent Recruiting Dashboard | Confidential</p>
        <p>For support, contact: support@realtymetricsolutions.com | Version 1.1</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
