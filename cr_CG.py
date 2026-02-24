import streamlit as st
import pandas as pd
import numpy as np
import cr_cg_utils as utils
from datetime import datetime, timedelta
import importlib

# Ensure the utility logic is fresh
importlib.reload(utils)

# ==============================================================================
# --- PAGE CONFIG ---
# ==============================================================================
st.set_page_config(layout="wide", page_title="Supply Assurance Control Tower")

# Custom CSS for modern look
st.markdown("""
    <style>
    .main { background-color: #0E1117; }
    .stMetric { background-color: #1E1E26; padding: 15px; border-radius: 10px; border: 1px solid #41424C; }
    [data-testid="stExpander"] { border: 1px solid #41424C; border-radius: 10px; background-color: #161B22; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# --- SIDEBAR: DATA INTAKE ---
# ==============================================================================
st.sidebar.title("📦 Control Tower v1.1")

with st.sidebar.expander("1. Data Sources", expanded=True):
    prod_files = st.file_uploader("Production Shot Data", accept_multiple_files=True, type=['xlsx', 'csv'])
    po_files = st.file_uploader("PO Planning Data", accept_multiple_files=True, type=['xlsx', 'csv'])

# Data Processing
df_raw = utils.load_production_data(prod_files) if prod_files else pd.DataFrame()
df_po = utils.load_po_data(po_files) if po_files else pd.DataFrame()

if df_raw.empty:
    st.info("👋 Welcome! Please upload production shot data in the sidebar to initialize the dashboard."); st.stop()

# ==============================================================================
# --- SIDEBAR: CASCADING HIERARCHY ---
# ==============================================================================
with st.sidebar.expander("2. Supply Scope Filters", expanded=True):
    # PO Selection (Primary Context)
    po_list = ["All Orders"] + sorted(df_raw['po_number'].dropna().unique().tolist())
    sel_po = st.selectbox("Purchase Order Context", po_list)
    
    scope_df = df_raw.copy()
    if sel_po != "All Orders":
        scope_df = scope_df[scope_df['po_number'] == sel_po]

    # Project Filter
    projects = sorted(scope_df['project'].dropna().unique().tolist())
    sel_proj = st.multiselect("Project", projects, default=projects)
    scope_df = scope_df[scope_df['project'].isin(sel_proj)]
    
    # Part Filter (The Target Anchor)
    parts = sorted(scope_df['part_id'].dropna().unique().tolist())
    sel_parts = st.multiselect("Part ID (Benchmark)", parts, default=parts)
    scope_df = scope_df[scope_df['part_id'].isin(sel_parts)]
    
    # Tool Filter (The Contributors)
    tools = sorted(scope_df['tool_id'].dropna().unique().tolist())
    sel_tools = st.multiselect("Tooling Assets", tools, default=tools)
    scope_df = scope_df[scope_df['tool_id'].isin(sel_tools)]

# ==============================================================================
# --- SIDEBAR: CAPACITY CONFIG ---
# ==============================================================================
with st.sidebar.expander("3. Operating Window (Stress Config)"):
    cal_days = st.number_input("Operating Days / Week", 1, 7, 5)
    cal_hours = st.number_input("Operating Hours / Day", 1, 24, 16)
    cal_config = {'days': cal_days, 'hours': cal_hours}
    tolerance = st.slider("Tolerance Band (%)", 0.01, 0.25, 0.05)

# ==============================================================================
# --- CALCULATION ENGINE ---
# ==============================================================================
t_assure, t_risk, t_trends = st.tabs(["📊 Supply Assurance", "🗼 Risk Tower", "📈 Trends"])

e_config = {'tolerance': tolerance, 'run_interval_hours': 8}

if not scope_df.empty:
    # Aggregated Engine Results
    engine = utils.CapacityRiskCalculator(scope_df, e_config)
    res = engine.results
    
    # Weekly Aggregation
    agg_df = utils.get_supply_metrics(scope_df, 'Weekly', e_config)
    
    # PO Alignment Logic
    po_target_qty = 0
    po_due_date = datetime.now()
    if not df_po.empty and sel_po != "All Orders":
        # Match by PO number and selected parts to get demand
        po_match = df_po[(df_po['po_number'] == sel_po) & (df_po['part_id'].isin(sel_parts))]
        if not po_match.empty:
            po_target_qty = po_match['total_qty'].sum()
            po_due_date = po_match['due_date'].max()
            # Distribute PO quantity across active weeks for the hybrid chart
            agg_df['Target'] = po_target_qty / max(1, len(agg_df))

# ==============================================================================
# --- TAB 1: SUPPLY ASSURANCE ---
# ==============================================================================
with t_assure:
    if sel_po == "All Orders":
        st.warning("⚠️ Please select a specific Purchase Order in the sidebar to view fulfillment metrics."); st.stop()

    st.header(f"PO Fulfillment Analysis: {sel_po}")
    
    # KPI Grid
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Actual Production", f"{res['actual_output']:,.0f}")
    k2.metric("PO Goal", f"{po_target_qty:,.0f}")
    
    fulfillment = (res['actual_output'] / po_target_qty * 100) if po_target_qty > 0 else 0
    k3.metric("Fulfillment %", f"{fulfillment:.1f}%")
    
    # Stress Calculation
    actual_hrs = res['total_runtime_sec'] / 3600
    planned_hrs = cal_days * cal_hours * (agg_df['Period'].nunique())
    stress_pct = (actual_hrs / planned_hrs * 100) if planned_hrs > 0 else 0
    k4.metric("Tool Stress Factor", f"{stress_pct:.1f}%", delta=f"{stress_pct-100:.1f}% vs Plan", delta_color="inverse")

    st.markdown("---")
    
    # Main Visualization
    st.plotly_chart(utils.create_hybrid_chart(agg_df, cal_config), use_container_width=True)
    
    # Capacity Insights
    st.info(f"**Operational Analysis:** {utils.generate_capacity_insights(res)}")

    # Secondary Visuals
    c_left, c_right = st.columns([2, 1])
    with c_left:
        st.plotly_chart(utils.create_po_burnup(agg_df, po_target_qty, po_due_date), use_container_width=True)
    with c_right:
        st.plotly_chart(utils.plot_waterfall(res), use_container_width=True)

    st.markdown("---")
    with st.expander("🔬 Technical Detail: Shot-by-Shot Analysis"):
        st.plotly_chart(utils.plot_shot_analysis(res['processed_df']), use_container_width=True)

# ==============================================================================
# --- TAB 2: RISK TOWER ---
# ==============================================================================
with t_risk:
    st.header("Supply Chain Risk Tower")
    st.markdown("Aggregating operational risk based on current filter hierarchy.")
    
    pivot = st.radio("Group Analysis By:", ["Part ID", "Project", "Tool ID"], horizontal=True)
    
    risk_summary = scope_df.groupby(pivot.lower().replace(" ","_")).agg({
        'actual_ct': 'mean',
        'tool_id': 'nunique',
        'working_cavities': 'mean'
    }).reset_index().rename(columns={
        'actual_ct': 'Avg Cycle Time (s)',
        'tool_id': 'Tools in Group',
        'working_cavities': 'Avg Cavities'
    })
    
    st.dataframe(risk_summary.style.format({'Avg Cycle Time (s)': '{:.2f}', 'Avg Cavities': '{:.1f}'}), use_container_width=True, hide_index=True)

# ==============================================================================
# --- TAB 3: TRENDS ---
# ==============================================================================
with t_trends:
    st.subheader("Time-Series Performance Totals")
    st.dataframe(agg_df.style.format({
        'Actual Output': '{:,.0f}', 
        'Optimal Output': '{:,.0f}',
        'Downtime Loss': '{:,.0f}',
        'Slow Loss': '{:,.0f}',
        'Run Time Sec': '{:,.0f}'
    }), use_container_width=True, hide_index=True)