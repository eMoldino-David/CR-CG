import streamlit as st
import pandas as pd
import numpy as np
import cr_CG_utils as utils
from datetime import datetime, timedelta
import importlib

# Ensure the latest logic from the utility file is loaded
importlib.reload(utils)

# ==============================================================================
# --- PAGE CONFIGURATION ---
# ==============================================================================
st.set_page_config(
    layout="wide", 
    page_title="Supply Assurance Control Tower",
    page_icon="📦"
)

# Custom CSS for a professional dark-themed look
st.markdown("""
    <style>
    .main { background-color: #0E1117; }
    .stMetric { background-color: #1E1E26; padding: 15px; border-radius: 10px; border: 1px solid #41424C; }
    [data-testid="stExpander"] { border: 1px solid #41424C; border-radius: 10px; background-color: #161B22; }
    div[data-testid="stExpander"] p { font-size: 0.85rem; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# --- SIDEBAR: DATA INTAKE & CONFIG ---
# ==============================================================================
st.sidebar.title("📦 Control Tower v1.5")

with st.sidebar.expander("1. Data Sources", expanded=True):
    prod_files = st.file_uploader("Production Shot Data", accept_multiple_files=True, type=['xlsx', 'csv'])
    po_files = st.file_uploader("PO Planning Data", accept_multiple_files=True, type=['xlsx', 'csv'])

# Load and standardize data using the utility engine
df_raw = utils.load_production_data(prod_files) if prod_files else pd.DataFrame()
df_po = utils.load_po_data(po_files) if po_files else pd.DataFrame()

if df_raw.empty:
    st.info("👋 Welcome! Please upload production shot data in the sidebar to initialize the Control Tower.")
    st.stop()

with st.sidebar.expander("2. Supply Scope Filters", expanded=True):
    # Primary Pivot: Purchase Order
    po_list = ["All Orders"] + sorted(df_raw['po_number'].dropna().unique().tolist())
    sel_po = st.selectbox("Purchase Order Context", po_list)
    
    scope_df = df_raw.copy()
    if sel_po != "All Orders":
        scope_df = scope_df[scope_df['po_number'] == sel_po]

    # Project Filter
    projects = sorted(scope_df['project'].dropna().unique().tolist())
    sel_proj = st.multiselect("Project", projects, default=projects)
    scope_df = scope_df[scope_df['project'].isin(sel_proj)]
    
    # Part Filter (Demand Anchor)
    parts = sorted(scope_df['part_id'].dropna().unique().tolist())
    sel_parts = st.multiselect("Part ID (Benchmark)", parts, default=parts)
    scope_df = scope_df[scope_df['part_id'].isin(sel_parts)]
    
    # Tool Filter (Operational Contributors)
    tools = sorted(scope_df['tool_id'].dropna().unique().tolist())
    sel_tools = st.multiselect("Tooling Assets", tools, default=tools)
    scope_df = scope_df[scope_df['tool_id'].isin(sel_tools)]

with st.sidebar.expander("3. Operating Window (Stress Config)"):
    cal_days = st.number_input("Operating Days / Week", 1, 7, 5)
    cal_hours = st.number_input("Operating Hours / Day", 1, 24, 16)
    cal_config = {'days': cal_days, 'hours': cal_hours}
    tolerance = st.slider("Cycle Tolerance Band (%)", 0.01, 0.25, 0.05)

# ==============================================================================
# --- DASHBOARD LAYOUT ---
# ==============================================================================
t_assure, t_risk, t_trends = st.tabs(["📊 Supply Assurance", "🗼 Risk Tower", "📈 Trends"])

# Configuration for the calculation engine
e_config = {'tolerance': tolerance, 'run_interval_hours': 8}

if not scope_df.empty:
    # Run the core calculation engine
    engine = utils.CapacityRiskCalculator(scope_df, e_config)
    res = engine.results
    
    # Handle PO demand targets
    po_target_qty = 0
    po_due_date = datetime.now()
    if not df_po.empty and sel_po != "All Orders":
        po_match = df_po[(df_po['po_number'] == sel_po) & (df_po['part_id'].isin(sel_parts))]
        if not po_match.empty:
            po_target_qty = po_match['total_qty'].sum()
            po_due_date = po_match['due_date'].max()

    # Aggregate metrics for time-series charts
    agg_df = utils.get_supply_metrics(scope_df, e_config, po_target=po_target_qty)

# --- TAB 1: SUPPLY ASSURANCE ---
with t_assure:
    if sel_po == "All Orders":
        st.warning("⚠️ Please select a specific Purchase Order in the sidebar to view detailed Supply Assurance metrics.")
        st.stop()

    st.header(f"Fulfillment Analysis: {sel_po}")
    
    # Top KPI Row
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Actual Production", f"{res['actual_output']:,.0f}")
    k2.metric("PO Target Goal", f"{po_target_qty:,.0f}")
    
    fulfillment = (res['actual_output'] / po_target_qty * 100) if po_target_qty > 0 else 0
    k3.metric("Fulfillment Status", f"{fulfillment:.1f}%")
    
    # Stress Metric calculation
    actual_hrs = res['total_runtime_sec'] / 3600
    planned_hrs = cal_days * cal_hours * (agg_df['period'].nunique())
    stress_pct = (actual_hrs / planned_hrs * 100) if planned_hrs > 0 else 0
    k4.metric("Tool Stress Factor", f"{stress_pct:.1f}%", delta=f"{stress_pct-100:.1f}% vs Plan", delta_color="inverse")

    st.markdown("---")
    
    # Main Performance Chart
    st.plotly_chart(utils.create_hybrid_chart(agg_df, cal_config), use_container_width=True)
    
    # Natural Language Insights
    st.info(f"**Automated Summary:** {utils.generate_capacity_insights(res)}", icon="ℹ️")

    # Mid-Section: Fulfillment and Stability Drivers
    c_left, c_mid, c_right = st.columns([1, 1, 1])
    with c_left:
        st.plotly_chart(utils.create_modern_gauge(fulfillment, "PO Achievement"), use_container_width=True)
    with c_mid:
        st.plotly_chart(utils.create_modern_gauge(res['stability_index'], "Stability Index"), use_container_width=True)
    with c_right:
        st.subheader("Stability Driver")
        st.plotly_chart(utils.create_stability_driver_bar(res['mtbf_min'], res['mttr_min']), use_container_width=True)
        st.caption(f"MTBF: {res['mtbf_min']:.1f}m | MTTR: {res['mttr_min']:.1f}m")

    # Bottom Row: Demand Path and Loss Bridge
    st.markdown("---")
    c_burn, c_bridge = st.columns([2, 1])
    with c_burn:
        st.plotly_chart(utils.create_po_burnup(agg_df, po_target_qty, po_due_date), use_container_width=True)
    with c_bridge:
        st.plotly_chart(utils.plot_waterfall(res), use_container_width=True)

    with st.expander("🔬 Deep Dive: Shot-by-Shot Production Rhythm"):
        st.plotly_chart(utils.plot_shot_analysis(res['processed_df']), use_container_width=True)

# --- TAB 2: RISK TOWER ---
with t_risk:
    st.header("Strategic Asset Risk Tower")
    pivot_choice = st.radio("Group Analysis By:", ["Part ID", "Project", "Tool ID"], horizontal=True)
    
    risk_stats = scope_df.groupby(pivot_choice.lower().replace(" ","_")).agg({
        'shot_time': 'count',
        'actual_ct': 'mean',
        'tool_id': 'nunique',
        'working_cavities': 'mean'
    }).reset_index().rename(columns={
        'shot_time': 'Total Shots',
        'actual_ct': 'Avg Cycle Time (s)',
        'tool_id': 'Unique Assets',
        'working_cavities': 'Avg Cavities'
    })
    
    st.dataframe(risk_stats.style.format({
        'Avg Cycle Time (s)': '{:.2f}', 
        'Avg Cavities': '{:.1f}'
    }), use_container_width=True, hide_index=True)

# --- TAB 3: TRENDS ---
with t_trends:
    st.subheader("Weekly Performance Log")
    st.dataframe(agg_df.style.format({
        'actual_output': '{:,.0f}',
        'runtime_sec': '{:,.0f}',
        'target': '{:,.0f}',
        'stability': '{:.1f}%'
    }), use_container_width=True, hide_index=True)