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
    for col in ['TotalVolume', 'ListSideVolume', 'SellSideVolume']:
        if col in df.columns:
            df[col] = df[col].apply(dollar_to_numeric)
    for col in ['TotalCount', 'DaysOnMarket']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    if 'MovedDate' in df.columns:
        df['MovedDate'] = pd.to_datetime(df['MovedDate'], format='%m/%d/%Y', errors='coerce')
        df['RecencyDays'] = (pd.Timestamp.now() - df['MovedDate']).dt.days
    df['FirstName'] = df['FirstName'].fillna('Unnamed').replace('', 'Unnamed')
    df['LastName'] = df['LastName'].fillna('Agent').replace('', 'Agent')
    if 'City' in df.columns:
        df['City_clean'] = df['City'].str.strip().str.title()
    # Remove duplicate agents (keep highest TotalVolume)
    if 'TotalVolume' in df.columns:
        df = df.sort_values('TotalVolume', ascending=False)
    df = df.drop_duplicates(subset=['FirstName', 'LastName', 'OfficeName'], keep='first')
    return df

def safe_rescale(series):
    """Safely rescale a series to 0-1 range"""
    if series.empty or series.isna().all():
        return pd.Series([0.5] * len(series), index=series.index)
    min_val, max_val = series.min(), series.max()
    if pd.isna(min_val) or pd.isna(max_val) or min_val == max_val:
        return pd.Series([0.5] * len(series), index=series.index)
    return (series - min_val) / (max_val - min_val)

def format_currency(x):
    """Format currency values"""
    if pd.isna(x): return "—"
    return f"${x:,.0f}"

def format_integer(x):
    """Format integer values"""
    if pd.isna(x): return "—"
    return f"{int(x):,}"

def has_county_data(df):
    """Check if real county data exists (not all Unknown)"""
    counties = df['County'].dropna().unique()
    return not (len(counties) == 0 or (len(counties) == 1 and counties[0] == 'Unknown'))

@st.cache_data
def load_and_prepare_data():
    """Load and prepare all data"""
    try:
        raw_data = pd.read_csv("CA Dashboard.csv")
        agents = clean_agent_data(raw_data)
        try:
            city2county = pd.read_csv("CA_City_to_County_Mapping.csv")
            city2county['City_clean'] = city2county['City'].str.strip().str.title()
            city2county['County'] = city2county['County'].str.strip().str.title()
            city2county = city2county[['City_clean', 'County']].drop_duplicates('City_clean')
        except FileNotFoundError:
            st.warning("City to County mapping file not found. All cities will be marked as 'Unknown'.")
            city2county = pd.DataFrame(columns=['City_clean', 'County'])
        if 'City_clean' in agents.columns:
            unmapped = agents[['City_clean']].drop_duplicates()
            unmapped = unmapped[~unmapped['City_clean'].isin(city2county['City_clean'])]
            unmapped['County'] = 'Unknown'
            city2county_full = pd.concat([city2county, unmapped], ignore_index=True).drop_duplicates('City_clean')
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
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.write("*Please contact your administrator for access.*")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        return False
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
        title_font = Font(bold=True, size=14, color="FFFFFF")
        title_fill = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2E5090", end_color="2E5090", fill_type="solid")
        header_align = Alignment(horizontal="center")
        alt_row_fill = PatternFill(start_color="F0F7FF", end_color="F0F7FF", fill_type="solid")

        ws['A1'] = 'RealtyMetric Solutions – Agent Recruiting Report'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))

        for col in range(1, len(df.columns) + 1):
            cell = ws.cell(row=2, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align

        # Auto-size columns
        for i, col_name in enumerate(df.columns, 1):
            max_len = max(len(str(col_name)), df[col_name].astype(str).str.len().max() or 0)
            col_letter = openpyxl.utils.get_column_letter(i)
            ws.column_dimensions[col_letter].width = min(max_len + 4, 30)

        if 'TotalVolume' in df.columns:
            tv_col = list(df.columns).index('TotalVolume') + 1
            for row in range(3, len(df) + 3):
                ws.cell(row=row, column=tv_col).number_format = '$#,##0'

        if 'Final_Score' in df.columns:
            sc_col = list(df.columns).index('Final_Score') + 1
            for row in range(3, len(df) + 3):
                ws.cell(row=row, column=sc_col).number_format = '0.0000'

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

    # --- Branding Header with Logo ---
    st.markdown("""
    <div class="branding-header">
        <div style="display:flex; align-items:center; gap:18px;">
            <svg width="70" height="70" viewBox="0 0 70 70" xmlns="http://www.w3.org/2000/svg">
                <rect width="70" height="70" rx="14" fill="white"/>
                <text x="35" y="40" text-anchor="middle" font-family="Arial Black, Arial" font-size="28" font-weight="900" fill="#1B2A4A">RM</text>
                <polyline points="8,58 20,48 32,53 48,38 62,28" stroke="#5CACEE" stroke-width="3.5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
                <circle cx="62" cy="28" r="3.5" fill="#5CACEE"/>
            </svg>
            <div>
                <h1 style="margin:0; font-size:28px; color:white;">RealtyMetric Solutions</h1>
                <p style="margin:4px 0 0 0; color:#A8C4E0; font-size:14px;">Agent Recruiting Dashboard – California</p>
            </div>
        </div>
        <div style="text-align:right; color:#A8C4E0; font-size:12px;">
            <p style="margin:0;">Version 1.2</p>
            <p style="margin:2px 0 0 0;">© 2026 RealtyMetric Solutions</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    This interactive tool helps brokerage teams identify promising real estate agents for recruitment.
    Rows with missing fields are kept and scored conservatively so you see more people.
    """)

    # --- About This Tool ---
    with st.expander("📖 About This Tool"):
        st.markdown("""
        ### What Is This Tool?
        RealtyMetric Solutions' Agent Recruiting Dashboard is a data-driven platform designed to help 
        real estate brokerages identify, evaluate, and recruit top-performing agents. It analyzes publicly 
        available agent activity data and generates an objective recruiting score for each agent, helping 
        your team prioritize outreach efforts.

        ---

        ### How Does the Scoring Work?
        Each agent is scored on a **0 to 1 scale** based on four weighted components:

        | Factor | Description | What It Means |
        |---|---|---|
        | **Volume** | Agent's total transaction volume | Higher volume = stronger producer |
        | **DOM (Days on Market)** | Average days properties stay listed | Lower DOM = more efficient agent |
        | **Recency** | How recently the agent completed a transaction | More recent = more active agent |
        | **Office Changes** | Whether the agent has switched offices | Recent switchers may be open to moving again |

        ---

        ### How Are Weights Used?
        The sliders in the sidebar control how much each factor contributes to the **Final Score**. 
        For example, if your brokerage values high volume producers, increase the Volume weight. 
        If you prefer agents who close quickly, increase the DOM weight. The weights do **not** need 
        to add up to 1.0 — the system will calculate the score based on whatever combination you set.

        ---

        ### How Should I Interpret the Scores?
        - **0.8 – 1.0** → Highly promising recruits — prioritize outreach
        - **0.6 – 0.8** → Strong candidates — worth investigating further
        - **0.4 – 0.6** → Moderate candidates — consider for long-term pipeline
        - **Below 0.4** → Lower priority at this time

        ---

        ### What About Missing Data?
        Agents with missing fields are **not removed** from the results. Instead, missing values are 
        scored conservatively (treated as the least favorable value), so these agents still appear but 
        rank lower. This ensures you don't accidentally overlook promising agents due to incomplete records.
        """)

    # --- How to Use ---
    with st.expander("🛈 How to Use This Tool"):
        st.markdown("""
        - **Adjust Weights:** Use the sidebar sliders to set how much each factor matters to your recruiting criteria.
        - **Filter by County:** Select one or more counties to narrow your focus to specific markets.
        - **Set Thresholds:** Use Min Volume, Max DOM, and Max Recency to exclude agents who don't meet your minimums.
        - **Search Agents:** Use the search bar in the sidebar to quickly find a specific agent by name.
        - **Read the Charts:** The Top 10 chart shows your best candidates at a glance. The scatter plot shows score vs. volume.
        - **Export Data:** Download results as CSV or Excel for further analysis or sharing with your team.
        - **Missing Values:** Filters allow missing values — agents with missing DOM or Recency are kept but scored conservatively.
        """)

    # --- Disclaimer / Terms of Use ---
    with st.expander("⚖️ Disclaimer & Terms of Use"):
        st.markdown("""
        ### Disclaimer
        RealtyMetric Solutions provides this dashboard for **informational purposes only**. The data and 
        scores presented are based on publicly available information and proprietary algorithms, but are 
        not guaranteed to be accurate, complete, or up to date. RealtyMetric Solutions makes no warranties, 
        express or implied, regarding the quality or fitness of this tool for any particular purpose.

        ---

        ### Terms of Use
        By using this dashboard, you agree to the following:

        1. **Confidentiality:** All data displayed in this dashboard is confidential. You may not share, 
           distribute, or publish any information obtained from this tool without the prior written consent 
           of RealtyMetric Solutions.

        2. **No Redistribution:** You may not copy, reproduce, or distribute the content of this dashboard 
           or any exports without authorization.

        3. **Authorized Use Only:** This tool is intended solely for the authorized user or organization. 
           Sharing login credentials or access with unauthorized parties is strictly prohibited.

        4. **No Guarantee of Results:** Recruiting outcomes based on this tool are not guaranteed. 
           RealtyMetric Solutions is not liable for any business decisions made based on the data or 
           scores provided.

        5. **Data Accuracy:** While we strive to provide accurate data, RealtyMetric Solutions is not 
           responsible for any errors or omissions in the data. Users should independently verify 
           information before making critical business decisions.

        6. **Intellectual Property:** The scoring methodology, algorithms, and dashboard design are the 
           intellectual property of RealtyMetric Solutions and are protected by law.

        7. **Changes:** RealtyMetric Solutions reserves the right to update, modify, or discontinue 
           this service at any time without prior notice.

        ---

        *By accessing this dashboard, you acknowledge and agree to these terms. For questions regarding 
        these terms, please contact support@realtymetricsolutions.com*
        """)

    # Load data
    agents, city2county_full = load_and_prepare_data()
    if agents is None:
        st.stop()

    # --- Sidebar ---
    with st.sidebar:
        st.header("⚖️ Scoring Weights")
        volume_weight   = st.slider("Weight: Volume",                  0.0, 1.0, 0.30, 0.05)
        dom_weight      = st.slider("Weight: DOM (Lower is Better)",   0.0, 1.0, 0.35, 0.05)
        recency_weight  = st.slider("Weight: Recency (Recent Better)", 0.0, 1.0, 0.20, 0.05)
        change_weight   = st.slider("Weight: Office Changes",          0.0, 1.0, 0.20, 0.05)
        st.divider()
        st.header("🔍 Filters")
        st.caption("Rows with missing fields are kept")
        counties = sorted(agents['County'].dropna().unique())
        county_filter = st.multiselect("County (optional):", options=counties, default=None, help="Leave empty to include all counties")
        min_volume  = st.number_input("Min Total Volume ($):", min_value=0, value=100000, step=100000, format="%d")
        max_dom     = st.number_input("Max Days on Market:",   min_value=1, value=180, step=10)
        max_recency = st.number_input("Max Days Since Move:",  min_value=1, value=999, step=50)
        st.divider()
        st.header("🔎 Agent Search")
        search_query = st.text_input("Search by agent name...", placeholder="e.g. John Smith")

    # --- Process Data ---
    def process_agents():
        df0 = agents.copy()
        total_rows = len(df0)
        df1 = df0[df0['County'].isin(county_filter)] if county_filter else df0.copy()
        df2 = df1.copy()
        if 'TotalVolume'  in df2.columns: df2 = df2[(df2['TotalVolume'].isna())  | (df2['TotalVolume']  >= min_volume)]
        if 'DaysOnMarket' in df2.columns: df2 = df2[(df2['DaysOnMarket'].isna()) | (df2['DaysOnMarket'] <= max_dom)]
        if 'RecencyDays'  in df2.columns: df2 = df2[(df2['RecencyDays'].isna())  | (df2['RecencyDays']  <= max_recency)]

        df3 = df2.copy()
        df3['vol_for_score']     = df3['TotalVolume'].fillna(0) if 'TotalVolume' in df3.columns else 0
        df3['dom_for_score']     = df3['DaysOnMarket'].fillna(max_dom) if 'DaysOnMarket' in df3.columns else max_dom
        df3['recency_for_score'] = df3['RecencyDays'].fillna(max_recency) if 'RecencyDays' in df3.columns else max_recency

        df3['Volume_Score']  = safe_rescale(df3['vol_for_score'])
        df3['DOM_Score']     = safe_rescale(-df3['dom_for_score'])
        df3['Recency_Score'] = safe_rescale(-df3['recency_for_score'])

        if 'PreviousOfficeName' in df3.columns:
            has_prev = (~df3['PreviousOfficeName'].isna()) & (df3['PreviousOfficeName'] != '')
            df3['Change_Score'] = safe_rescale(has_prev.astype(int))
        else:
            df3['Change_Score'] = 0

        df3['Final_Score'] = (
            volume_weight  * df3['Volume_Score']  +
            dom_weight     * df3['DOM_Score']     +
            recency_weight * df3['Recency_Score'] +
            change_weight  * df3['Change_Score']
        )

        df_final = df3.sort_values('Final_Score', ascending=False).reset_index(drop=True)
        df_final['Rank'] = df_final.index + 1

        cols = ['Rank','FirstName','LastName','OfficeName','County','TotalVolume','DaysOnMarket','RecencyDays','Final_Score']
        cols = [c for c in cols if c in df_final.columns]
        df_display = df_final[cols].copy()

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

    # --- Row Audit ---
    st.subheader("🔍 Row Audit")
    a1, a2, a3, a4 = st.columns(4)
    with a1: st.metric("Rows in CSV",            f"{total_rows:,}")
    with a2: st.metric("After County Filter",    f"{after_county:,}")
    with a3: st.metric("After Thresholds",       f"{after_filters:,}")
    with a4: st.metric("Rows Shown",             f"{len(df_display):,}")

    # --- Top 10 Bar Chart ---
    st.subheader("🏆 Top 10 Agents by Recruiting Score")
    top10 = df_display.head(10).copy()
    top10['Agent Name'] = top10['FirstName'] + ' ' + top10['LastName']
    fig_bar = px.bar(
        top10, x='Final_Score', y='Agent Name', orientation='h',
        color='Final_Score', color_continuous_scale='Blues', text='Final_Score',
        hover_data={'OfficeName': True, 'County': True, 'TotalVolume': ':$,.0f', 'Final_Score': ':.4f'}
    )
    fig_bar.update_layout(
        yaxis=dict(categoryorder='total ascending', title=None),
        xaxis_title="Recruiting Score", height=420,
        showlegend=False, coloraxis_showscale=False
    )
    fig_bar.update_traces(texttemplate='%{text:.4f}', textposition='inside')
    st.plotly_chart(fig_bar, use_container_width=True)

    # --- County Charts (only shown if county data exists) ---
    show_counties = has_county_data(df_display)

    if show_counties:
        st.subheader("🗺️ County Insights")
        col_pie, col_vol = st.columns(2)

        # County Distribution Pie Chart
        with col_pie:
            county_counts = df_display.groupby('County').size().reset_index(name='Agent Count')
            county_counts = county_counts.sort_values('Agent Count', ascending=False)
            fig_pie = px.pie(
                county_counts, values='Agent Count', names='County',
                color_discrete_sequence=px.colors.qualitative.Set2, hole=0.4,
                title="Agent Distribution by County"
            )
            fig_pie.update_layout(height=420, showlegend=True, legend=dict(orientation="v", x=1.02, y=0.5))
            fig_pie.update_traces(textinfo='percent', textposition='inside')
            st.plotly_chart(fig_pie, use_container_width=True)

        # Volume by County Bar Chart
        with col_vol:
            county_vol = df_display.dropna(subset=['TotalVolume']).groupby('County')['TotalVolume'].sum().reset_index()
            county_vol = county_vol.sort_values('TotalVolume', ascending=True)
            fig_vol = px.bar(
                county_vol, x='TotalVolume', y='County', orientation='h',
                color='TotalVolume', color_continuous_scale='Blues',
                title="Total Volume by County"
            )
            fig_vol.update_layout(
                height=420, xaxis_title="Total Volume ($)", yaxis_title=None,
                showlegend=False, coloraxis_showscale=False,
                xaxis=dict(tickprefix="$", tickformat=",")
            )
            fig_vol.update_traces(texttemplate='$%{x:,.0f}', textposition='inside')
            st.plotly_chart(fig_vol, use_container_width=True)

    # --- Scatter Plot (colored by County if available, otherwise by Score) ---
    st.subheader("📈 Score vs. Total Volume")
    if 'TotalVolume' in df_display.columns:
        scatter_df = df_display.dropna(subset=['TotalVolume']).copy()

        if show_counties:
            fig_scatter = px.scatter(
                scatter_df, x='TotalVolume', y='Final_Score', color='County',
                color_discrete_sequence=px.colors.qualitative.Set2,
                hover_data={
                    'Rank': True, 'FirstName': True, 'LastName': True,
                    'County': True, 'OfficeName': True,
                    'TotalVolume': ':$,.0f', 'DaysOnMarket': ':.0f',
                    'RecencyDays': ':.0f', 'Final_Score': ':.3f'
                },
                title="Recruiting Score vs Total Volume"
            )
        else:
            fig_scatter = px.scatter(
                scatter_df, x='TotalVolume', y='Final_Score', color='Final_Score',
                color_continuous_scale='Viridis',
                hover_data={
                    'Rank': True, 'FirstName': True, 'LastName': True,
                    'County': True, 'OfficeName': True,
                    'TotalVolume': ':$,.0f', 'DaysOnMarket': ':.0f',
                    'RecencyDays': ':.0f', 'Final_Score': ':.3f'
                },
                title="Recruiting Score vs Total Volume"
            )

        fig_scatter.update_layout(
            xaxis_title="Total Volume ($)", yaxis_title="Recruiting Score",
            height=500, xaxis=dict(tickprefix="$", tickformat=",")
        )
        fig_scatter.update_traces(marker=dict(size=9))
        st.plotly_chart(fig_scatter, use_container_width=True)

    # --- Office Performance Comparison ---
    st.subheader("🏢 Office Performance Comparison")
    office_df = df_display.dropna(subset=['OfficeName']).copy()
    office_df = office_df[office_df['OfficeName'].str.strip() != '']

    if len(office_df) > 0:
        office_stats = office_df.groupby('OfficeName').agg(
            Agent_Count=('FirstName', 'count'),
            Avg_Score=('Final_Score', 'mean'),
            Total_Volume=('TotalVolume', 'sum'),
            Avg_Volume=('TotalVolume', 'mean')
        ).reset_index()
        office_stats = office_stats.sort_values('Avg_Score', ascending=False).head(15)
        # Truncate long office names for readability
        office_stats['Display_Name'] = office_stats['OfficeName'].str[:45]

        fig_office = px.bar(
            office_stats, x='Avg_Score', y='Display_Name', orientation='h',
            color='Agent_Count', color_continuous_scale='Viridis',
            hover_data={
                'OfficeName': True,
                'Total_Volume': ':$,.0f',
                'Avg_Volume': ':$,.0f',
                'Agent_Count': True,
                'Avg_Score': ':.4f',
                'Display_Name': False
            },
            text='Agent_Count',
            title="Top 15 Offices by Average Recruiting Score"
        )
        fig_office.update_layout(
            height=520, xaxis_title="Average Recruiting Score",
            yaxis_title=None, coloraxis_title="# Agents",
            yaxis=dict(categoryorder='total ascending')
        )
        fig_office.update_traces(texttemplate='%{text} agents', textposition='outside')
        st.plotly_chart(fig_office, use_container_width=True)
    else:
        st.info("No office data available to display.")

    # --- Ranked Agent Table ---
    st.subheader("📋 Ranked Agent Table")
    df_formatted = df_display.copy()
    if 'TotalVolume'  in df_formatted.columns: df_formatted['TotalVolume']  = df_formatted['TotalVolume'].apply(format_currency)
    if 'DaysOnMarket' in df_formatted.columns: df_formatted['DaysOnMarket'] = df_formatted['DaysOnMarket'].apply(format_integer)
    if 'RecencyDays'  in df_formatted.columns: df_formatted['RecencyDays']  = df_formatted['RecencyDays'].apply(format_integer)
    if 'Final_Score'  in df_formatted.columns: df_formatted['Final_Score']  = df_formatted['Final_Score'].round(4)

    st.dataframe(
        df_formatted, use_container_width=True, hide_index=True,
        column_config={
            "Rank":          st.column_config.NumberColumn("Rank", format="%d"),
            "FirstName":     "First Name",
            "LastName":      "Last Name",
            "OfficeName":    "Office Name",
            "County":        "County",
            "TotalVolume":   "Total Volume",
            "DaysOnMarket":  "Days on Market",
            "RecencyDays":   "Days Since Move",
            "Final_Score":   st.column_config.NumberColumn("Final Score", format="%.4f")
        }
    )

    # --- Export Buttons ---
    st.subheader("📥 Export Data")
    col_csv, col_xlsx = st.columns(2)
    with col_csv:
        csv_buffer = StringIO()
        df_display.to_csv(csv_buffer, index=False)
        st.download_button(
            label="📄 Download CSV", data=csv_buffer.getvalue(),
            file_name=f"RealtyMetric_CA_Agents_{date.today().strftime('%Y-%m-%d')}.csv",
            mime="text/csv"
        )
    with col_xlsx:
        xlsx_buffer = export_to_excel(df_display)
        st.download_button(
            label="📊 Download Excel", data=xlsx_buffer,
            file_name=f"RealtyMetric_CA_Agents_{date.today().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- Footer ---
    st.markdown("""
    <div class="footer">
        <p>© 2026 RealtyMetric Solutions | Agent Recruiting Dashboard | Confidential</p>
        <p>For support, contact: support@realtymetricsolutions.com | Version 1.2</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
