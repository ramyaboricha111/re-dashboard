import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os, shutil, tempfile

st.set_page_config(page_title="Silver Star RE · Analytics", page_icon="🏢",
                   layout="wide", initial_sidebar_state="collapsed")

# ── PASSWORD PROTECTION ───────────────────────────────────────────────────────
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    st.markdown("""
    <div style='display:flex;flex-direction:column;align-items:center;justify-content:center;
                height:80vh;'>
        <h1 style='color:#f1f5f9;font-size:28px;margin-bottom:8px;'>🏢 Silver Star RE</h1>
        <p style='color:#64748b;margin-bottom:32px;'>Analytics Dashboard</p>
    </div>
    """, unsafe_allow_html=True)
    col1, col2, col3 = st.columns([2,1,2])
    with col2:
        pwd = st.text_input("Password", type="password", placeholder="Enter password")
        if st.button("Login", use_container_width=True):
            if pwd == "SilverStar2024!":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password")
    return False

if not check_password():
    st.stop()

# ── DARK THEME CSS ─────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"], .stApp {
  font-family: 'Inter', sans-serif;
  background-color: #0d1117 !important;
  color: #e2e8f0 !important;
}
.stApp { background-color: #0d1117 !important; }
.block-container { background-color: #0d1117 !important; padding-top: 1.2rem; padding-bottom: 2rem; }

/* Sidebar */
[data-testid="stSidebar"] { background-color: #0d1117 !important; border-right: 1px solid #1e293b; }
[data-testid="stSidebar"] * { color: #64748b !important; }

/* Metrics */
div[data-testid="metric-container"] {
  background: #161b22 !important; border: 1px solid #21262d !important;
  border-radius: 12px; padding: 18px 22px;
}
div[data-testid="metric-container"] label { font-size:11px !important; color:#4b5563 !important; font-weight:500; text-transform:uppercase; letter-spacing:.08em; }
div[data-testid="metric-container"] [data-testid="stMetricValue"] { font-size:28px !important; color:#f1f5f9 !important; font-weight:700; }
div[data-testid="metric-container"] [data-testid="stMetricDelta"] { font-size:12px !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { background: transparent; gap:4px; border-bottom: 1px solid #21262d; }
.stTabs [data-baseweb="tab"] { font-size:13px; font-weight:500; color:#64748b; padding:8px 18px; border-radius:8px 8px 0 0; background:transparent; }
.stTabs [aria-selected="true"] { color:#60a5fa !important; border-bottom:2px solid #60a5fa !important; background:#1a2233 !important; }

/* Inputs / selects */
div[data-baseweb="select"] > div { background-color:#161b22 !important; border-color:#21262d !important; color:#e2e8f0 !important; border-radius:8px; }
div[data-baseweb="select"] span { color:#e2e8f0 !important; }
div[data-baseweb="popover"] { background:#161b22 !important; border:1px solid #21262d !important; }
div[data-baseweb="menu"] { background:#161b22 !important; }
div[data-baseweb="option"] { background:#161b22 !important; color:#e2e8f0 !important; }
div[data-baseweb="option"]:hover { background:#1e293b !important; }
input[type="text"], input[type="date"] { background:#161b22 !important; color:#e2e8f0 !important; border-color:#21262d !important; border-radius:8px; }
.stMultiSelect [data-baseweb="tag"] { background:#1e3a5f !important; color:#60a5fa !important; }
label { color:#94a3b8 !important; font-size:11px !important; text-transform:uppercase; letter-spacing:.06em; }

/* Dataframes */
[data-testid="stDataFrame"] { border-radius:10px; overflow:hidden; border:1px solid #21262d !important; }
[data-testid="stDataFrame"] iframe { background:#161b22 !important; }

/* Headings */
h1 { color:#f1f5f9 !important; font-weight:700 !important; letter-spacing:-.03em; font-size:24px !important; }
h2,h3,h4 { color:#cbd5e1 !important; font-weight:600 !important; }
p, span { color:#94a3b8; }
hr { border:none; border-top:1px solid #21262d; margin:1.2rem 0; }

/* Buttons */
.stButton > button { background:#1e3a5f; color:#60a5fa; border:1px solid #2d5a8e; border-radius:8px; font-weight:500; }
.stButton > button:hover { background:#2d5a8e; }

/* Download button */
.stDownloadButton > button { background:#16213e; color:#818cf8; border:1px solid #3730a3; border-radius:8px; }

/* Section cards */
.section-card { background:#161b22; border:1px solid #21262d; border-radius:12px; padding:16px 20px; margin-bottom:16px; }

footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

# ── DARK PLOT DEFAULTS ────────────────────────────────────────────────────────
PLOT_BG    = '#161b22'
PAPER_BG   = '#0d1117'
GRID_COLOR = '#21262d'
TEXT_COLOR = '#94a3b8'
TICK_COLOR = '#64748b'

PLOT_DEFAULTS = dict(
    plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG,
    font=dict(family='Inter, sans-serif', size=12, color=TEXT_COLOR),
)

PALETTE = {
    'blue':   '#60a5fa', 'teal':   '#2dd4bf', 'amber':  '#fbbf24',
    'red':    '#f87171', 'green':  '#4ade80', 'purple': '#a78bfa',
    'slate':  '#94a3b8', 'orange': '#fb923c',
}

_FOLDER = os.path.dirname(os.path.abspath(__file__))
FILE_PATHS = {
    "vacancy": os.path.join(_FOLDER, "Daily_-_Vacancy.xlsx"),
    "master":  os.path.join(_FOLDER, "Master_File.xlsx"),
    "monthly": os.path.join(_FOLDER, "Monthly_Numbers.xlsx"),
    "pnl":     os.path.join(_FOLDER, "PnL_-_Ramya_Format.xlsx"),
}

PNL_DESC_ORDER = [
    'Gross Scheduled Rent','Vacancy Loss','Promotion','Rental Assistance','Eviction',
    'Uncollected This Month','Past Due Collections','Net Collected Rent',
    'Late Fee/NSF Charge','Parking','Misc Income','RUBS','Insurance-Tenant',
    'Total Actual Income','Total Repairs & Maintenance','Total General Maint/Business Exp',
    'Total Rent Ready','Total Utilities','Total Payroll','Total Controllable Expenses',
    'Total Noncontrollable Expenses','Total Operating Expenses',
    'Net Operating Income (NOI)','Net Income',
]

def load_excel(key):
    import time, uuid
    src = FILE_PATHS[key]
    # Copy to a uniquely named temp file to avoid Windows file lock issues
    tmp_path = os.path.join(tempfile.gettempdir(), f"re_dash_{key}_{uuid.uuid4().hex}.xlsx")
    for attempt in range(5):
        try:
            shutil.copy2(src, tmp_path)
            xl = pd.ExcelFile(tmp_path)
            return xl
        except (PermissionError, OSError):
            time.sleep(0.5)
    # Last attempt — read directly into bytes buffer
    with open(src, 'rb') as f:
        import io
        return pd.ExcelFile(io.BytesIO(f.read()))

def read_wide_sheet(xl, sheet_name, value_col):
    df = xl.parse(sheet_name, header=0)
    if str(df.columns[0]).strip() in ['#','Unnamed: 0']:
        df = df.drop(columns=df.columns[0])
    df = df.rename(columns={df.columns[0]: 'Property'})
    df['Property'] = df['Property'].astype(str).str.strip().str.upper()
    df = df[df['Property'].notna() & ~df['Property'].isin(['NAN',''])]
    date_cols = [c for c in df.columns if c != 'Property' and isinstance(c, datetime)]
    rows = []
    for _, row in df.iterrows():
        for dc in date_cols:
            rows.append({'Property': row['Property'], 'Date': dc.date(), value_col: row[dc]})
    result = pd.DataFrame(rows)
    result[value_col] = pd.to_numeric(result[value_col], errors='coerce').fillna(0)
    return result

@st.cache_data(ttl=1800, show_spinner=False)
def load_all_data():
    xl_master = load_excel("master")
    master = xl_master.parse("RE Listing")
    master = master.drop(columns=['#'], errors='ignore')
    master['Property'] = master['Property'].astype(str).str.strip().str.upper()
    master = master.rename(columns={'Apartment Name':'Apt_Name','Regional Manager':'RM',
                                    'Property Manager':'PM','Units':'Total_Units'})
    master = master[['Property','Apt_Name','Total_Units','RM','PM','City','State','Entity Name']]

    xl_vac    = load_excel("vacancy")
    daily = read_wide_sheet(xl_vac,'Ready','Units_Ready').merge(
            read_wide_sheet(xl_vac,'Not Ready','Units_Not_Ready'),on=['Property','Date'],how='outer').merge(
            read_wide_sheet(xl_vac,'WO','WO_Count'),on=['Property','Date'],how='outer').merge(
            read_wide_sheet(xl_vac,'Delinquency','Delinquency_Amt'),on=['Property','Date'],how='outer').fillna(0)
    # Compute month-start delinquency as the target (last column = oldest date)
    delinq_raw = xl_vac.parse('Delinquency', header=0)
    if str(delinq_raw.columns[0]).strip() in ['#','Unnamed: 0']:
        delinq_raw = delinq_raw.drop(columns=delinq_raw.columns[0])
    delinq_raw = delinq_raw.rename(columns={delinq_raw.columns[0]: 'Property'})
    delinq_raw['Property'] = delinq_raw['Property'].astype(str).str.strip().str.upper()
    # Date cols sorted descending — find first date of current month as target
    date_cols = sorted([c for c in delinq_raw.columns
                        if isinstance(c, datetime) or isinstance(c, pd.Timestamp)],
                       reverse=True)
    if date_cols:
        latest_dt  = pd.Timestamp(date_cols[0])
        cur_month  = latest_dt.to_period('M')
        month_cols = [c for c in date_cols if pd.Timestamp(c).to_period('M') == cur_month]
        # First day of month = last entry in month_cols (they're descending)
        target_col = month_cols[-1] if month_cols else date_cols[-1]
        dq_goals   = delinq_raw[['Property', target_col]].rename(columns={target_col: 'DQ_Goal'})
        dq_goals['DQ_Goal'] = pd.to_numeric(dq_goals['DQ_Goal'], errors='coerce').fillna(0)
    else:
        dq_goals = pd.DataFrame(columns=['Property','DQ_Goal'])
    daily['Date'] = pd.to_datetime(daily['Date'])
    daily = daily.merge(master, on='Property', how='left')
    daily['Vacancy_Total'] = daily['Units_Ready'] + daily['Units_Not_Ready']
    daily['Vacancy_Rate']  = daily.apply(
        lambda r: (r['Vacancy_Total']/r['Total_Units']*100) if r['Total_Units']>0 else 0, axis=1)

    xl_pnl  = load_excel("pnl")
    pnl_raw = xl_pnl.parse("PnL", header=0)
    pnl_raw = pnl_raw.dropna(subset=['Property','Type'])
    pnl_raw = pnl_raw[pnl_raw['Type'].isin(['Income','Expense'])]
    pnl_raw['Description'] = pnl_raw['Description'].astype(str).str.strip()
    pnl_raw['Property']    = pnl_raw['Property'].astype(str).str.strip().str.upper()
    month_cols = [c for c in pnl_raw.columns if c not in ['Property','Type','GL Code','Description']]
    pnl = pnl_raw.melt(id_vars=['Property','Type','GL Code','Description'],
                       value_vars=month_cols, var_name='Month_Str', value_name='Amount')
    pnl['Amount'] = pd.to_numeric(pnl['Amount'], errors='coerce').fillna(0)
    pnl['Month']  = pd.to_datetime(pnl['Month_Str'].str.strip(), format='%b %y', errors='coerce')
    pnl = pnl.dropna(subset=['Month'])
    pnl = pnl.merge(master[['Property','Apt_Name','RM','State','City','Total_Units']],
                    on='Property', how='left')

    xl_mo   = load_excel("monthly")
    apps    = xl_mo.parse("Application", header=0)
    movein  = xl_mo.parse("MoveIn",      header=0)
    renewal = xl_mo.parse("Renewal",     header=0)
    moveout = xl_mo.parse("MoveOut",     header=0)

    # Standardise property column name
    apps['Property']    = apps['Property'].astype(str).str.strip().str.upper()
    moveout['Property'] = moveout['Property'].astype(str).str.strip().str.upper()
    movein['Property']  = movein['Property Code'].astype(str).str.strip().str.upper()
    renewal['Property'] = renewal['Property Code'].astype(str).str.strip().str.upper()

    apps['Date_Applied']      = pd.to_datetime(apps['Date Applied'],       errors='coerce')
    movein['MoveIn_Date']     = pd.to_datetime(movein['Move in Date'],      errors='coerce')
    renewal['Renewal_Date']   = pd.to_datetime(renewal['Current Lease From'], errors='coerce')
    moveout['MoveOut_Date']   = pd.to_datetime(moveout['Move Out Date'],    errors='coerce')

    # Rent comparison columns for MoveIn
    movein['Prev_Actual_Rent']  = pd.to_numeric(movein['Previous Actual Rent'], errors='coerce')
    movein['New_Actual_Rent']   = pd.to_numeric(movein['Actual Rent'],           errors='coerce')
    movein['Rent_Var']          = movein['New_Actual_Rent'] - movein['Prev_Actual_Rent']
    movein['Rent_Var_Pct']      = movein.apply(
        lambda r: (r['Rent_Var']/r['Prev_Actual_Rent']*100) if r['Prev_Actual_Rent'] and r['Prev_Actual_Rent']!=0 else 0, axis=1)

    # Rent comparison columns for Renewal
    renewal['Prev_Actual_Rent'] = pd.to_numeric(renewal['Previous Actual Rent'], errors='coerce')
    renewal['New_Actual_Rent']  = pd.to_numeric(renewal['Current Actual Rent'],  errors='coerce')
    renewal['Rent_Var']         = renewal['New_Actual_Rent'] - renewal['Prev_Actual_Rent']
    renewal['Rent_Var_Pct']     = renewal.apply(
        lambda r: (r['Rent_Var']/r['Prev_Actual_Rent']*100) if r['Prev_Actual_Rent'] and r['Prev_Actual_Rent']!=0 else 0, axis=1)

    for df in [apps, movein, renewal, moveout]:
        df2 = df.merge(master[['Property','Apt_Name','RM','State']], on='Property', how='left')
        df[['Apt_Name','RM','State']] = df2[['Apt_Name','RM','State']]

    return master, daily, pnl, apps, movein, renewal, moveout, dq_goals

with st.spinner("Loading data…"):
    master, daily, pnl, apps, movein, renewal, moveout, dq_goals = load_all_data()

# ── SCOPE OPTIONS (formatted) ─────────────────────────────────────────────────
def make_scope_options(master):
    opts = ['All Properties']
    for rm in sorted(master['RM'].dropna().unique()):
        opts.append(f"Region — {rm}")
    for _, row in master.dropna(subset=['Apt_Name']).sort_values('Apt_Name').iterrows():
        opts.append(f"{row['Property']} — {row['Apt_Name']}")
    return opts

scope_options = make_scope_options(master)

def filter_by_scope(df, scope, prop_col='Property'):
    if scope == 'All Properties':
        return df
    if scope.startswith('Region — '):
        rm = scope.replace('Region — ','')
        props = master[master['RM']==rm]['Property'].tolist()
        return df[df[prop_col].isin(props)]
    prop_code = scope.split(' — ')[0]
    return df[df[prop_col]==prop_code]

def master_by_scope(scope):
    if scope == 'All Properties': return master
    if scope.startswith('Region — '):
        rm = scope.replace('Region — ','')
        return master[master['RM']==rm]
    prop_code = scope.split(' — ')[0]
    return master[master['Property']==prop_code]

# ── HEADER ────────────────────────────────────────────────────────────────────
st.title("🏢 Silver Star Real Estate · Analytics")

tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview & Vacancy", "📋 Leasing", "💰 P&L", "🗂 MIS Table"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW & VACANCY  (filters on page)
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    # Inline filters
    ov_f1, ov_f2 = st.columns([3,2])
    with ov_f1:
        ov_scope = st.selectbox("Scope", scope_options, index=0, key='ov_scope')
    with ov_f2:
        min_date = daily['Date'].min().date()
        max_date = daily['Date'].max().date()
        date_range = st.date_input("Date range", value=(min_date, max_date),
                                   min_value=min_date, max_value=max_date, key='ov_dates')

    d_start  = pd.Timestamp(date_range[0]) if len(date_range)==2 else daily['Date'].min()
    d_end    = pd.Timestamp(date_range[1]) if len(date_range)==2 else daily['Date'].max()
    master_f = master_by_scope(ov_scope)
    sel_props = master_f['Property'].tolist()
    daily_f  = daily[daily['Property'].isin(sel_props) & (daily['Date']>=d_start) & (daily['Date']<=d_end)]

    sel_states = sorted(master_f['State'].dropna().unique().tolist())
    st.caption(f"As of **{daily['Date'].max().strftime('%b %d, %Y')}**  ·  {len(sel_props)} properties  ·  {', '.join(sel_states)}")
    st.markdown("---")

    # KPIs
    latest_date = daily_f['Date'].max()
    latest      = daily_f[daily_f['Date']==latest_date]
    prev_date   = daily_f[daily_f['Date']<latest_date]['Date'].max()
    prev        = daily_f[daily_f['Date']==prev_date] if pd.notna(prev_date) else latest

    total_units    = master_f['Total_Units'].sum()
    total_vacant   = int(latest['Vacancy_Total'].sum())
    total_ready    = int(latest['Units_Ready'].sum())
    total_notready = int(latest['Units_Not_Ready'].sum())
    total_delinq   = latest['Delinquency_Amt'].sum()
    delta_vacant   = total_vacant - int(prev['Vacancy_Total'].sum())
    delta_delinq   = total_delinq - prev['Delinquency_Amt'].sum()

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Units",     f"{int(total_units):,}")
    c2.metric("Vacant Units",    f"{total_vacant}",
              delta=f"{delta_vacant:+d} units vs prev day", delta_color="inverse")
    c3.metric("Units Ready",     f"{total_ready}")
    c4.metric("Units Not Ready", f"{total_notready}")
    c5.metric("Delinquency",     f"${total_delinq:,.0f}",
              delta=f"${abs(delta_delinq):,.0f} {'▼ less' if delta_delinq < 0 else '▲ more'} vs prev",
              delta_color="normal" if delta_delinq < 0 else "inverse")

    st.markdown("---")

    # Line chart + Pie
    vc_l, vc_r = st.columns([3,2])
    with vc_l:
        st.markdown("#### Daily vacancy — units")
        vac_trend = daily_f.groupby('Date').agg(
            Vacant=('Vacancy_Total','sum'),
            Ready=('Units_Ready','sum'),
            Not_Ready=('Units_Not_Ready','sum')
        ).reset_index().sort_values('Date')
        # Fill missing weekday dates with previous day value (forward fill)
        all_dates = pd.date_range(vac_trend['Date'].min(), vac_trend['Date'].max(), freq='D')
        vac_trend = (vac_trend.set_index('Date')
                              .reindex(all_dates)
                              .ffill()
                              .reset_index()
                              .rename(columns={'index':'Date'}))
        # Remove weekends (Sat=5, Sun=6) and zero rows
        vac_trend = vac_trend[vac_trend['Date'].dt.dayofweek < 5]
        vac_trend = vac_trend[vac_trend['Vacant'] > 0]

        last7 = vac_trend['Date'].max() - pd.Timedelta(days=7)
        # Add padding so last point isn't clipped
        x_max = vac_trend['Date'].max() + pd.Timedelta(hours=12)

        fig = go.Figure()
        specs = [('Vacant','Vacant Total',PALETTE['blue']),
                 ('Ready','Ready',PALETTE['teal']),
                 ('Not_Ready','Not Ready',PALETTE['amber'])]
        for col, name, color in specs:
            fig.add_trace(go.Scatter(
                x=vac_trend['Date'], y=vac_trend[col], name=name,
                mode='lines+markers+text',
                text=vac_trend[col].astype(int), textposition='top center',
                textfont=dict(size=9, color=color),
                line=dict(color=color, width=2.5),
                marker=dict(size=6, color=color,
                            line=dict(color=PLOT_BG, width=1.5)),
            ))
        fig.update_layout(
            **PLOT_DEFAULTS, height=300,
            margin=dict(l=0, r=30, t=40, b=0),
            legend=dict(orientation='h', y=1.15, x=0,
                        font=dict(color=TEXT_COLOR), bgcolor='rgba(0,0,0,0)'),
            yaxis=dict(gridcolor=GRID_COLOR, zeroline=False,
                       tickfont=dict(color=TICK_COLOR), title=''),
            xaxis=dict(
                range=[last7, x_max],
                rangeslider=dict(visible=True, thickness=0.04,
                                 bgcolor=PLOT_BG, bordercolor=GRID_COLOR),
                gridcolor=GRID_COLOR, tickfont=dict(color=TICK_COLOR),
            ),
        )
        st.plotly_chart(fig, use_container_width=True)

    with vc_r:
        st.markdown("#### Ready vs Not Ready (today)")
        pie_data = pd.DataFrame({'Status':['Ready','Not Ready'],
                                  'Units':[total_ready, total_notready]})
        fig2 = px.pie(pie_data, names='Status', values='Units', color='Status',
                      color_discrete_map={'Ready':'#6366f1','Not Ready':'#f472b6'},
                      hole=0.55)
        fig2.update_traces(textinfo='label+value',
                           textfont=dict(size=14, color='#ffffff', family='Inter'),
                           marker=dict(line=dict(color='#0d1117', width=3)))
        fig2.update_layout(**PLOT_DEFAULTS, height=300, showlegend=True,
                           legend=dict(font=dict(color='#94a3b8'), orientation='h', y=-0.05),
                           margin=dict(l=10,r=10,t=40,b=10))
        st.plotly_chart(fig2, use_container_width=True)

    # RM bar
    st.markdown("#### Vacancy by RM (today — units)")
    rm_vac = latest.groupby('RM').agg(Vacant=('Vacancy_Total','sum'),
                                       Total=('Total_Units','sum')).reset_index()
    rm_vac['Rate'] = (rm_vac['Vacant']/rm_vac['Total']*100).round(1)
    rm_vac = rm_vac.sort_values('Vacant', ascending=False)
    bar_colors = [PALETTE['red'] if r>=15 else PALETTE['amber'] if r>=5 else PALETTE['teal']
                  for r in rm_vac['Rate']]
    fig3 = go.Figure(go.Bar(
        x=rm_vac['RM'], y=rm_vac['Vacant'],
        text=rm_vac['Rate'].apply(lambda x: f"{x:.1f}%"),
        textposition='outside', textfont=dict(color=TEXT_COLOR),
        marker_color=bar_colors, marker=dict(line=dict(width=0)),
    ))
    fig3.update_layout(**PLOT_DEFAULTS, height=260,
                       margin=dict(l=0,r=0,t=30,b=0),
                       yaxis=dict(gridcolor=GRID_COLOR, zeroline=False,
                                  tickfont=dict(color=TICK_COLOR), title=''),
                       xaxis=dict(tickfont=dict(color=TEXT_COLOR)))
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")
    st.markdown("#### Work orders (last 7 days default)")
    wo_trend = daily_f.groupby('Date')['WO_Count'].sum().reset_index().sort_values('Date')
    wo_trend = wo_trend[wo_trend['WO_Count'] > 0].copy()
    last7_wo = wo_trend[wo_trend['WO_Count']>0]['Date'].max() - pd.Timedelta(days=7)
    x_max_wo = wo_trend['Date'].max() + pd.Timedelta(hours=12)
    fig_wo = go.Figure(go.Bar(
        x=wo_trend['Date'], y=wo_trend['WO_Count'],
        text=wo_trend['WO_Count'].astype(int),
        textposition='outside', textfont=dict(color=PALETTE['blue']),
        marker_color=PALETTE['blue'], marker=dict(line=dict(width=0), opacity=0.85),
    ))
    fig_wo.update_layout(
        **PLOT_DEFAULTS, height=260, margin=dict(l=0,r=30,t=10,b=0),
        yaxis=dict(gridcolor=GRID_COLOR, zeroline=False, tickfont=dict(color=TICK_COLOR)),
        xaxis=dict(
            range=[last7_wo, x_max_wo],
            rangeslider=dict(visible=True, thickness=0.04,
                             bgcolor=PLOT_BG, bordercolor=GRID_COLOR),
            tickfont=dict(color=TEXT_COLOR),
        )
    )
    st.plotly_chart(fig_wo, use_container_width=True)

    st.markdown("---")
    st.markdown("#### Delinquency by RM")
    latest_all = daily[daily['Date']==daily['Date'].max()]
    prev_all   = daily[daily['Date']==(daily[daily['Date']<daily['Date'].max()]['Date'].max())]
    dq_t = latest_all.groupby('RM')['Delinquency_Amt'].sum().reset_index().rename(columns={'Delinquency_Amt':'Today'})
    dq_p = prev_all.groupby('RM')['Delinquency_Amt'].sum().reset_index().rename(columns={'Delinquency_Amt':'Prev_Day'})
    dq   = dq_t.merge(dq_p, on='RM', how='left').fillna(0)
    dq['Var'] = dq['Prev_Day'] - dq['Today']

    # Merge month-start target
    if not dq_goals.empty:
        goals_with_rm = dq_goals.merge(master[['Property','RM']], on='Property', how='left')
        rm_goals = goals_with_rm.groupby('RM')['DQ_Goal'].sum().reset_index()
        dq = dq.merge(rm_goals, on='RM', how='left').fillna(0)
        dq['Collected']   = (dq['DQ_Goal'] - dq['Today']).clip(lower=0)
        dq['Still_Needed']= dq['Today']   # what's still left = today's balance
        dq['Pct_Achieved']= dq.apply(
            lambda r: min(r['Collected']/r['DQ_Goal']*100, 100) if r['DQ_Goal']>0 else 0, axis=1)
        has_goals = True
    else:
        has_goals = False

    totals_row = {'RM':'Total','Today':dq['Today'].sum(),
                  'Prev_Day':dq['Prev_Day'].sum(),'Var':dq['Var'].sum()}
    if has_goals:
        totals_row['DQ_Goal']      = dq['DQ_Goal'].sum()
        totals_row['Collected']    = dq['Collected'].sum()
        totals_row['Still_Needed'] = dq['Still_Needed'].sum()
        totals_row['Pct_Achieved'] = (dq['Collected'].sum()/dq['DQ_Goal'].sum()*100) if dq['DQ_Goal'].sum()>0 else 0
    dq = pd.concat([dq, pd.DataFrame([totals_row])], ignore_index=True)

    dq_disp = dq[['RM','DQ_Goal','Collected','Still_Needed','Today','Prev_Day','Var'] if has_goals else ['RM','Today','Prev_Day','Var']].copy()
    dq_disp['Today']    = dq_disp['Today'].apply(lambda x: f"${x:,.1f}K")
    dq_disp['Prev_Day'] = dq_disp['Prev_Day'].apply(lambda x: f"${x:,.1f}K")
    dq_disp['Var']      = dq_disp['Var'].apply(
        lambda x: f"🟢 +{x:,.1f}K" if x>0 else (f"🔴 {x:,.1f}K" if x<0 else "0"))
    if has_goals:
        dq_disp['DQ_Goal']      = dq_disp['DQ_Goal'].apply(lambda x: f"${x:,.1f}K")
        dq_disp['Collected']    = dq_disp['Collected'].apply(lambda x: f"${x:,.1f}K")
        dq_disp['Still_Needed'] = dq_disp['Still_Needed'].apply(
            lambda x: f"${x:,.1f}K" if x > 0 else "✅ Done")
    rename_map = {'Today': f"Today ({latest_all['Date'].max().strftime('%m/%d')})",
                  'Prev_Day':'Prev Day','DQ_Goal':'Month Start (Target)',
                  'Collected':'Collected','Still_Needed':'Still Outstanding'}
    st.dataframe(dq_disp.rename(columns=rename_map), use_container_width=True, hide_index=True)

    # Horizontal progress bar chart — fixed width (100%), filled = % reduced
    if has_goals:
        st.markdown("#### Delinquency reduction — progress toward target by RM")
        dq_chart = dq[dq['RM'] != 'Total'].copy().sort_values('Pct_Achieved', ascending=True)

        fig_dq = go.Figure()

        # Background bar — full width = 100%
        fig_dq.add_trace(go.Bar(
            name='Remaining', y=dq_chart['RM'], x=[100]*len(dq_chart),
            orientation='h',
            marker_color='#1e293b', marker=dict(line=dict(width=0)),
            showlegend=False
        ))

        # Filled portion = % reduced
        fill_colors = [PALETTE['green'] if p>=100 else PALETTE['amber'] if p>=50 else PALETTE['red']
                       for p in dq_chart['Pct_Achieved']]
        fig_dq.add_trace(go.Bar(
            name='Reduced', y=dq_chart['RM'], x=dq_chart['Pct_Achieved'].clip(upper=100),
            orientation='h',
            marker_color=fill_colors, marker=dict(line=dict(width=0), opacity=0.9),
            text=dq_chart.apply(
                lambda r: f"  Collected ${r['Collected']:,.0f}K of ${r['DQ_Goal']:,.0f}K  ({r['Pct_Achieved']:.0f}%)", axis=1),
            textposition='inside', insidetextanchor='start',
            textfont=dict(color='white', size=11),
            showlegend=False
        ))

        # % label at end of bar
        for _, row in dq_chart.iterrows():
            pct = min(row['Pct_Achieved'], 100)
            color = PALETTE['green'] if pct>=100 else PALETTE['amber'] if pct>=50 else PALETTE['red']
            fig_dq.add_annotation(
                y=row['RM'], x=102,
                text=f"{pct:.0f}%",
                xanchor='left', showarrow=False,
                font=dict(size=12, color=color, family='Inter')
            )

        fig_dq.update_layout(
            **PLOT_DEFAULTS, barmode='overlay', height=max(250, len(dq_chart)*45),
            margin=dict(l=0, r=60, t=10, b=0),
            xaxis=dict(range=[0,115], showticklabels=False, showgrid=False, zeroline=False),
            yaxis=dict(tickfont=dict(color=TEXT_COLOR)),
            showlegend=False
        )
        st.plotly_chart(fig_dq, use_container_width=True)
    else:
        st.info("Add a 'Target' column to the Delinquency sheet to see goal progress chart.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — LEASING
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("#### Leasing Activity")

    lf1, lf2 = st.columns([2,3])
    with lf1:
        lease_scope = st.selectbox("Scope", scope_options, key='lease_scope')
    with lf2:
        valid_apps    = apps[apps['Date_Applied'].dt.year >= 2020]
        valid_movein  = movein[movein['MoveIn_Date'].dt.year >= 2020]
        valid_renewal = renewal[renewal['Renewal_Date'].dt.year >= 2020]
        all_lease_months = sorted(set(
            valid_apps['Date_Applied'].dt.to_period('M').dropna().astype(str).tolist() +
            valid_movein['MoveIn_Date'].dt.to_period('M').dropna().astype(str).tolist() +
            valid_renewal['Renewal_Date'].dt.to_period('M').dropna().astype(str).tolist()
        ), reverse=True)
        default_end = min(12, len(all_lease_months))
        sel_lease_months = st.multiselect("Months (for tables)", all_lease_months,
                                           default=all_lease_months[:default_end],
                                           key='lease_months')

    def scope_date_filter(df, scope, date_col, months=None):
        d = filter_by_scope(df, scope)
        d = d[d[date_col].notna() & (d[date_col].dt.year >= 2020)].copy()
        d['Month'] = d[date_col].dt.to_period('M').astype(str)
        if months is not None:
            d = d[d['Month'].isin(months)]
        return d

    # Chart: last 12 months hardcoded — use union of all date sources
    all_valid_months = sorted(set(
        valid_apps['Date_Applied'].dt.to_period('M').dropna().astype(str).tolist() +
        valid_movein['MoveIn_Date'].dt.to_period('M').dropna().astype(str).tolist() +
        valid_renewal['Renewal_Date'].dt.to_period('M').dropna().astype(str).tolist()
    ))
    chart_months = all_valid_months[-12:] if len(all_valid_months) >= 12 else all_valid_months

    apps_chart    = scope_date_filter(apps,    lease_scope, 'Date_Applied', chart_months)
    movein_chart  = scope_date_filter(movein,  lease_scope, 'MoveIn_Date',  chart_months)
    renewal_chart = scope_date_filter(renewal, lease_scope, 'Renewal_Date', chart_months)
    moveout_chart = scope_date_filter(moveout, lease_scope, 'MoveOut_Date', chart_months)

    # Tables: use month filter
    apps_f    = scope_date_filter(apps,    lease_scope, 'Date_Applied', sel_lease_months)
    movein_f  = scope_date_filter(movein,  lease_scope, 'MoveIn_Date',  sel_lease_months)
    renewal_f = scope_date_filter(renewal, lease_scope, 'Renewal_Date', sel_lease_months)
    moveout_f = scope_date_filter(moveout, lease_scope, 'MoveOut_Date', sel_lease_months)

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Applications",  f"{apps_f.shape[0]:,}")
    k2.metric("Move-ins",      f"{movein_f.shape[0]:,}")
    k3.metric("Renewals",      f"{renewal_f.shape[0]:,}")
    k4.metric("Move-outs",     f"{moveout_f.shape[0]:,}")

    st.markdown("#### Monthly trend (last 12 months)")
    apps_mo    = apps_chart.groupby('Month').size().reset_index(name='Applications')
    movein_mo  = movein_chart.groupby('Month').size().reset_index(name='Move-ins')
    renewal_mo = renewal_chart.groupby('Month').size().reset_index(name='Renewals')
    moveout_mo = moveout_chart.groupby('Month').size().reset_index(name='Move-outs')
    trend = (apps_mo.merge(movein_mo,on='Month',how='outer')
                    .merge(renewal_mo,on='Month',how='outer')
                    .merge(moveout_mo,on='Month',how='outer').fillna(0).sort_values('Month'))

    fig_lease = go.Figure()
    for col, name, color in [('Applications','Applications',PALETTE['blue']),
                               ('Move-ins','Move-ins',PALETTE['teal']),
                               ('Renewals','Renewals',PALETTE['purple']),
                               ('Move-outs','Move-outs',PALETTE['amber'])]:
        fig_lease.add_trace(go.Bar(
            name=name, x=trend['Month'], y=trend[col],
            text=trend[col].astype(int), textposition='outside',
            textfont=dict(size=10, color=color),
            marker_color=color, marker=dict(line=dict(width=0), opacity=0.85)
        ))
    fig_lease.update_layout(
        **PLOT_DEFAULTS, height=320, barmode='group',
        margin=dict(l=0,r=30,t=10,b=0),
        legend=dict(orientation='h', y=1.1, bgcolor='rgba(0,0,0,0)', font=dict(color=TEXT_COLOR)),
        yaxis=dict(gridcolor=GRID_COLOR, zeroline=False, tickfont=dict(color=TICK_COLOR)),
        xaxis=dict(tickangle=-30, tickfont=dict(color=TEXT_COLOR))
    )
    st.plotly_chart(fig_lease, use_container_width=True)

    # Monthly rent analysis table
    st.markdown("#### Monthly rent analysis (selected months)")

    def build_rent_table(df, date_col, prev_col, new_col, label):
        d = df[df[prev_col].notna() & (df[prev_col]>0)].copy()
        if d.empty:
            return pd.DataFrame()
        d['Month'] = d[date_col].dt.to_period('M').dt.to_timestamp()
        grp = d.groupby('Month').agg(
            Prev_Rent=(prev_col,'sum'),
            New_Rent=(new_col,'sum'),
        ).reset_index()
        grp['Var_$'] = grp['New_Rent'] - grp['Prev_Rent']
        grp['Var_%'] = (grp['Var_$'] / grp['Prev_Rent'] * 100).round(1)
        grp = grp.sort_values('Month', ascending=False)
        grp['Month'] = grp['Month'].dt.strftime('%b %Y')
        grp.columns = ['Month', f'Prev Actual Rent ({label})', f'New Actual Rent ({label})',
                       f'$ Var ({label})', f'% Var ({label})']
        # Add total row
        total_prev = grp[f'Prev Actual Rent ({label})'].sum()
        total_new  = grp[f'New Actual Rent ({label})'].sum()
        total_var  = total_new - total_prev
        total_pct  = round(total_var / total_prev * 100, 1) if total_prev else 0
        totals = pd.DataFrame([['Total', total_prev, total_new, total_var, total_pct]],
                               columns=grp.columns)
        return pd.concat([grp, totals], ignore_index=True)

    mi_table = build_rent_table(movein_f, 'MoveIn_Date', 'Prev_Actual_Rent', 'New_Actual_Rent', 'Move-in')
    rn_table = build_rent_table(renewal_f, 'Renewal_Date', 'Prev_Actual_Rent', 'New_Actual_Rent', 'Renewal')

    if not mi_table.empty and not rn_table.empty:
        combined = mi_table.merge(rn_table, on='Month', how='outer').fillna(0)
    elif not mi_table.empty:
        combined = mi_table
    else:
        combined = rn_table

    if not combined.empty:
        # Format numbers first, then style by bracket presence
        for col in combined.columns:
            if '$ Var' in col or 'Rent' in col:
                combined[col] = combined[col].apply(
                    lambda x: f"🔴 ({abs(x):,.0f})" if isinstance(x,(int,float)) and x<0
                    else (f"{x:,.0f}" if isinstance(x,(int,float)) else x))
            elif '% Var' in col:
                combined[col] = combined[col].apply(
                    lambda x: f"🔴 ({abs(x):.1f}%)" if isinstance(x,(int,float)) and x<0
                    else (f"{x:.1f}%" if isinstance(x,(int,float)) else x))

        st.dataframe(combined, use_container_width=True, hide_index=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### By Regional Manager")
        apps_rm    = apps_f.groupby('RM').size().reset_index(name='Applications')
        movein_rm  = movein_f.groupby('RM').size().reset_index(name='Move-ins')
        renewal_rm = renewal_f.groupby('RM').size().reset_index(name='Renewals')
        moveout_rm = moveout_f.groupby('RM').size().reset_index(name='Move-outs')
        rm_lease   = (apps_rm.merge(movein_rm,on='RM',how='outer')
                              .merge(renewal_rm,on='RM',how='outer')
                              .merge(moveout_rm,on='RM',how='outer').fillna(0))
        for c in ['Applications','Move-ins','Renewals','Move-outs']:
            rm_lease[c] = rm_lease[c].astype(int)
        st.dataframe(rm_lease, use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format='%d')
                                    for c in ['Applications','Move-ins','Renewals','Move-outs']})

    with col_b:
        st.markdown("#### By Property")
        apps_prop    = apps_f.groupby(['Property','Apt_Name']).size().reset_index(name='Applications')
        movein_prop  = movein_f.groupby('Property').size().reset_index(name='Move-ins')
        renewal_prop = renewal_f.groupby('Property').size().reset_index(name='Renewals')
        moveout_prop = moveout_f.groupby('Property').size().reset_index(name='Move-outs')
        prop_lease   = (apps_prop.merge(movein_prop,on='Property',how='outer')
                                  .merge(renewal_prop,on='Property',how='outer')
                                  .merge(moveout_prop,on='Property',how='outer').fillna(0))
        for c in ['Applications','Move-ins','Renewals','Move-outs']:
            prop_lease[c] = prop_lease[c].astype(int)
        prop_lease = prop_lease.sort_values('Applications', ascending=False)
        st.dataframe(prop_lease[['Apt_Name','Applications','Move-ins','Renewals','Move-outs']].rename(
            columns={'Apt_Name':'Property'}),
            use_container_width=True, hide_index=True, height=380,
            column_config={'Property': st.column_config.Column(width='medium'),
                           **{c: st.column_config.NumberColumn(c, format='%d')
                              for c in ['Applications','Move-ins','Renewals','Move-outs']}})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — P&L
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("#### P&L Analysis")
    pf1,pf2,pf3 = st.columns([2,2,2])
    with pf1:
        pnl_view = st.selectbox("View", ["Month over Month (MoM)","Year to Date (YTD)"])
    all_pnl_months = sorted(pnl['Month'].dropna().unique())
    all_years = sorted(set(m.year for m in all_pnl_months))
    with pf2:
        if "Month" in pnl_view:
            sel_month_str = st.selectbox("Select Month",
                [m.strftime('%b %Y') for m in all_pnl_months], index=len(all_pnl_months)-1)
            sel_month = pd.to_datetime(sel_month_str, format='%b %Y')
        else:
            sel_year = st.selectbox("Select Year", all_years, index=len(all_years)-1)
    with pf3:
        pnl_scope = st.selectbox("Scope", scope_options, index=0, key='pnl_scope')

    all_descs_raw     = list(dict.fromkeys(pnl['Description'].dropna().str.strip().tolist()))
    ordered_available = [d for d in PNL_DESC_ORDER if d in all_descs_raw]
    rest              = [d for d in all_descs_raw if d not in PNL_DESC_ORDER]
    all_descs_ordered = ordered_available + rest
    default_descs     = [d for d in PNL_DESC_ORDER if d in all_descs_raw]
    sel_descs = st.multiselect("P&L lines to show", all_descs_ordered, default=default_descs)

    pnl_prop = filter_by_scope(pnl, pnl_scope)

    def agg_pnl(df, descs, months):
        return (df[df['Description'].isin(descs) & df['Month'].isin(months)]
                .groupby('Description')['Amount'].sum().reset_index())

    if "Month" in pnl_view:
        comp_month  = sel_month - pd.DateOffset(years=1)
        col_curr, col_prev = sel_month.strftime('%b %Y'), comp_month.strftime('%b %Y')
        curr_months, prev_months = [sel_month], [comp_month]
    else:
        curr_months_all = [m for m in all_pnl_months if m.year==sel_year]
        n = len(curr_months_all)
        curr_months = curr_months_all
        prev_months = [m - pd.DateOffset(years=1) for m in curr_months_all]
        col_curr, col_prev = f"YTD {sel_year} ({n}M)", f"YTD {sel_year-1} ({n}M)"

    curr_agg = agg_pnl(pnl_prop, sel_descs, curr_months).rename(columns={'Amount':col_curr})
    prev_agg = agg_pnl(pnl_prop, sel_descs, prev_months).rename(columns={'Amount':col_prev})
    pnl_comp = curr_agg.merge(prev_agg, on='Description', how='outer').fillna(0)


    def fmt_amt(x): return f"🔴 ({abs(x):,.0f})" if x<0 else f"{x:,.0f}"
    def fmt_pct(x): return f"🔴 ({abs(x):.1f}%)" if x<0 else f"{x:.1f}%"

    # Maintain PnL order
    pnl_comp['_ord'] = pnl_comp['Description'].map({d:i for i,d in enumerate(all_descs_ordered)})
    pnl_comp = pnl_comp.sort_values('_ord').drop(columns='_ord')
    # Merge type info for variance direction logic
    type_map = pnl[['Description','Type']].drop_duplicates().set_index('Description')['Type'].to_dict()
    pnl_comp['Type'] = pnl_comp['Description'].map(type_map).fillna('Income')
    pnl_comp['Raw_Var'] = pnl_comp[col_curr] - pnl_comp[col_prev]
    # For expenses: going up is bad (negative), going down is good (positive)
    pnl_comp['Var ($)'] = pnl_comp.apply(
        lambda r: -r['Raw_Var'] if r['Type']=='Expense' else r['Raw_Var'], axis=1)
    pnl_comp['Var (%)'] = pnl_comp.apply(
        lambda r: (r['Var ($)']/abs(r[col_prev])*100) if r[col_prev]!=0 else 0, axis=1)

    pnl_display = pnl_comp[['Description', col_curr, col_prev, 'Var ($)', 'Var (%)']].copy()
    for col in [col_curr, col_prev]:
        pnl_display[col] = pnl_display[col].apply(fmt_amt)
    pnl_display['Var ($)'] = pnl_display['Var ($)'].apply(fmt_amt)
    pnl_display['Var (%)'] = pnl_display['Var (%)'].apply(fmt_pct)

    st.dataframe(pnl_display.rename(columns={'Description':'P&L Line'}),
                 use_container_width=True, hide_index=True, height=380)

    chart_data = pnl_comp[pnl_comp['Description'].isin(sel_descs)].copy()
    chart_data = chart_data[chart_data[[col_curr,col_prev]].abs().max(axis=1)>0]
    if not chart_data.empty:
        def fmt_k(x):
            if abs(x)>=1_000_000: s=f"${abs(x)/1_000_000:.1f}M"
            elif abs(x)>=1_000:   s=f"${abs(x)/1_000:.0f}K"
            else:                  s=f"${abs(x):.0f}"
            return f"({s})" if x<0 else s

        chart_data['pct_change'] = chart_data.apply(
            lambda r: (r['Var ($)']/abs(r[col_prev])*100) if r[col_prev]!=0 else 0, axis=1)
        chart_data['pct_label'] = chart_data['pct_change'].apply(
            lambda x: f"({abs(x):.1f}%)" if x<0 else f"+{x:.1f}%")

        fig_pnl = go.Figure()
        fig_pnl.add_trace(go.Bar(
            name=col_prev, x=chart_data['Description'], y=chart_data[col_prev],
            text=chart_data[col_prev].apply(fmt_k), textposition='outside',
            textfont=dict(color=PALETTE['slate']),
            marker_color=[PALETTE['red'] if v<0 else '#475569' for v in chart_data[col_prev]],
            marker=dict(line=dict(width=0), opacity=0.8)
        ))
        fig_pnl.add_trace(go.Bar(
            name=col_curr, x=chart_data['Description'], y=chart_data[col_curr],
            text=chart_data[col_curr].apply(fmt_k), textposition='outside',
            textfont=dict(color=PALETTE['blue']),
            marker_color=[PALETTE['red'] if v<0 else PALETTE['blue'] for v in chart_data[col_curr]],
            marker=dict(line=dict(width=0), opacity=0.9)
        ))
        for _, row in chart_data.iterrows():
            max_y = max(abs(row[col_curr]), abs(row[col_prev])) * 1.25
            color = PALETTE['red'] if row['pct_change']<0 else PALETTE['green']
            fig_pnl.add_annotation(
                x=row['Description'], y=max_y, text=row['pct_label'],
                showarrow=False, font=dict(size=10, color=color, family='Inter'),
                bgcolor=PLOT_BG, borderpad=2
            )
        fig_pnl.update_layout(
            **PLOT_DEFAULTS, barmode='group', height=420,
            margin=dict(l=0,r=0,t=50,b=0),
            legend=dict(orientation='h', y=1.08, bgcolor='rgba(0,0,0,0)',
                        font=dict(color=TEXT_COLOR)),
            yaxis=dict(gridcolor=GRID_COLOR, zeroline=False,
                       tickfont=dict(color=TICK_COLOR), tickprefix='$'),
            xaxis=dict(tickangle=-30, tickfont=dict(color=TEXT_COLOR))
        )
        st.plotly_chart(fig_pnl, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — MIS TABLE
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    # Use overview filters if set, else all
    try:
        mis_props = sel_props
        mis_daily = daily_f
        mis_master = master_f
    except:
        mis_props = master['Property'].tolist()
        mis_daily = daily
        mis_master = master

    latest_mis = mis_daily['Date'].max()
    st.markdown(f"#### Property MIS — {latest_mis.strftime('%b %d, %Y')}")

    mis = (mis_daily[mis_daily['Date']==latest_mis]
           .groupby('Property')
           .agg(Apt_Name=('Apt_Name','first'), RM=('RM','first'), State=('State','first'),
                Total_Units=('Total_Units','first'), Units_Ready=('Units_Ready','sum'),
                Units_Not_Ready=('Units_Not_Ready','sum'), WO_Count=('WO_Count','sum'),
                Delinquency_Amt=('Delinquency_Amt','sum')).reset_index())
    mis['Vacancy']      = mis['Units_Ready'] + mis['Units_Not_Ready']
    mis['Vacancy_Rate'] = (mis['Vacancy']/mis['Total_Units']*100).round(1)

    def status(row):
        v = row['Vacancy_Rate']
        if v < 5:   return '🟢 Healthy'
        if v <= 15: return '🟡 Watch'
        return '🔴 At Risk'
    mis['Status'] = mis.apply(status, axis=1)

    display = mis[['Apt_Name','State','RM','Total_Units','Vacancy','Vacancy_Rate',
                   'Units_Ready','Units_Not_Ready','WO_Count','Delinquency_Amt','Status']].copy()
    display = display.rename(columns={
        'Apt_Name':'Property','Total_Units':'Total','Vacancy_Rate':'Vac %',
        'Units_Ready':'Ready','Units_Not_Ready':'Not Ready',
        'WO_Count':'WOs','Delinquency_Amt':'Delinquency ($)'
    }).sort_values('Vac %', ascending=False)
    display['Vac %']           = display['Vac %'].apply(lambda x: f"{x:.1f}%")
    display['Delinquency ($)'] = display['Delinquency ($)'].apply(lambda x: f"${x:,.1f}K")

    st.dataframe(display, use_container_width=True, hide_index=True,
                 column_config={
                     'Property': st.column_config.Column(width='medium'),
                     'RM':       st.column_config.Column(width='medium'),
                     **{c: st.column_config.Column(width='small')
                        for c in display.columns if c not in ['Property','RM']}
                 })

    st.markdown("#### Summary by Regional Manager")
    rm_sum = mis.groupby('RM').agg(
        Properties=('Property','count'), Total=('Total_Units','sum'),
        Vacant=('Vacancy','sum'), Delinquency_Amt=('Delinquency_Amt','sum'),
        WOs=('WO_Count','sum')).reset_index()
    rm_sum['Vac %'] = (rm_sum['Vacant']/rm_sum['Total']*100).round(1).apply(lambda x: f"{x:.1f}%")
    rm_sum['Delinquency ($K)'] = rm_sum['Delinquency_Amt'].apply(lambda x: f"${x:,.1f}K")
    rm_sum = rm_sum.drop(columns=['Delinquency_Amt'])
    st.dataframe(rm_sum, use_container_width=True, hide_index=True)

    csv = mis.to_csv(index=False)
    st.download_button("⬇ Download MIS as CSV", csv,
                       file_name=f"MIS_{latest_mis.strftime('%Y-%m-%d')}.csv", mime="text/csv")
