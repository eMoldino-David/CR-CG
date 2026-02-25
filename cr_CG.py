import streamlit as st
import pandas as pd
import numpy as np
import cr_CG_utils as utils
from datetime import datetime, timedelta
import importlib

# Ensure fresh logic from the utility file
importlib.reload(utils)

st.set_page_config(layout="wide", page_title="Supply Control Tower", page_icon="🏭")

# ==============================================================================
# --- SIDEBAR: DATA INTAKE ---
# ==============================================================================
st.sidebar.title("🏭 Asset Physics Tower")

with st.sidebar.expander("📂 1. Data Intake", expanded=True):
    p_files = st.file_uploader("Production Shot Data", accept_multiple_files=True)

df_all = utils.load_production_data(p_files) if p_files else pd.DataFrame()

if df_all.empty:
    st.info("👋 Please upload production shot data in the sidebar to begin."); st.stop()

# ==============================================================================
# --- SIDEBAR: NAVIGATION & FILTERS ---
# ==============================================================================
with st.sidebar.expander("🔍 2. Analysis Scope", expanded=True):
    nav_mode = st.radio("Navigation Mode", ["Single Tool View", "Hierarchical Breakdown"])
    
    scope_df = df_all.copy()
    
    if nav_mode == "Hierarchical Breakdown":
        # Project Filter
        projs = ["All Projects"] + sorted(scope_df['project'].dropna().unique().tolist())
        sel_proj = st.selectbox("Project", projs)
        if sel_proj != "All Projects":
            scope_df = scope_df[scope_df['project'] == sel_proj]
            
        # Component Filter
        if 'component_id' in scope_df.columns:
            comps = ["All Components"] + sorted(scope_df['component_id'].dropna().unique().tolist())
            sel_comp = st.selectbox("Component", comps)
            if sel_comp != "All Components":
                scope_df = scope_df[scope_df['component_id'] == sel_comp]
        
        # Part Filter
        parts = ["All Parts"] + sorted(scope_df['part_id'].dropna().unique().tolist())
        sel_part = st.selectbox("Part Number", parts)
        if sel_part != "All Parts":
            scope_df = scope_df[scope_df['part_id'] == sel_part]

    # Tooling selection (available in both modes, but usually final level in hierarchy)
    tools = sorted(scope_df['tool_id'].unique().tolist())
    if nav_mode == "Single Tool View":
        selected_tool = st.selectbox("Select Tooling ID", tools)
        display_name = selected_tool
        filtered_df = scope_df[scope_df['tool_id'] == selected_tool]
    else:
        selected_tools = st.multiselect("Select Tools (Aggregate)", tools, default=tools)
        display_name = f"Aggregated {len(selected_tools)} Tools"
        filtered_df = scope_df[scope_df['tool_id'].isin(selected_tools)]

# ==============================================================================
# --- SIDEBAR: PHYSICS CONFIG ---
# ==============================================================================
with st.sidebar.expander("⚙️ 3. Physics & Capacity Settings"):
    tolerance = st.slider("CT Tolerance Band", 0.01, 0.50, 0.05, 0.01)
    downtime_gap_tolerance = st.slider("Downtime Gap (sec)", 0.0, 5.0, 2.0, 0.5)
    run_interval_hours = st.slider("Run Interval (hours)", 1, 24, 8, 1)
    target_output_perc = st.slider("Target Output %", 50, 100, 90)
    default_cavities = st.number_input("Default Cavities", 1)

    config = {
        'tolerance': tolerance, 
        'downtime_gap_tolerance': downtime_gap_tolerance, 
        'run_interval_hours': run_interval_hours, 
        'default_cavities': default_cavities,
        'target_output_perc': target_output_perc
    }

# ==============================================================================
# --- MAIN UI DASHBOARD ---
# ==============================================================================
st.title("🏭 Production Asset Health")

if filtered_df.empty:
    st.warning("No data found for the current selection."); st.stop()

# Tabs based on the original app's logic
t_risk, t_opt, t_tgt, t_trend = st.tabs([
    "🗼 Risk Tower", 
    "🛠️ Capacity (Optimal)", 
    "🎯 Capacity (Target)", 
    "📈 Trends"
])

with t_risk:
    utils.render_risk_tower(df_all, config)

with t_opt:
    utils.render_dashboard(filtered_df, display_name, config, mode="Optimal")

with t_tgt:
    # Logic: Target basis uses the target_output_perc from config
    utils.render_dashboard(filtered_df, display_name, config, mode="Target")

with t_trend:
    st.subheader("Asset Performance Trends")
    # Simple weekly aggregation for the trend view
    df_trend = filtered_df.copy()
    df_trend['Week'] = df_trend['shot_time'].dt.to_period('W').apply(lambda r: r.start_time)
    
    trend_stats = []
    for week, subset in df_trend.groupby('Week'):
        calc = utils.CapacityRiskCalculator(subset, **config)
        res = calc.results
        trend_stats.append({
            'Week': week,
            'Actual Output': res['actual_output'],
            'Stability %': res['stability_index'],
            'Efficiency %': (res['normal_shots'] / res['total_shots'] * 100) if res['total_shots'] > 0 else 0
        })
    
    st.dataframe(pd.DataFrame(trend_stats), use_container_width=True)