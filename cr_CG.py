import streamlit as st
import pandas as pd
import numpy as np
import cr_CG_utils as utils
from datetime import datetime, timedelta
import importlib

# reload utils to ensure latest logic is active
importlib.reload(utils)

st.set_page_config(layout="wide", page_title="supply control tower")

# --- sidebar ---
st.sidebar.title("📦 supply control tower")

with st.sidebar.expander("1. data sources", expanded=True):
    prod_files = st.file_uploader("production data", accept_multiple_files=True)
    po_files = st.file_uploader("planning (po) data", accept_multiple_files=True)

df_raw = utils.load_production_data(prod_files) if prod_files else pd.DataFrame()
df_po = utils.load_po_data(po_files) if po_files else pd.DataFrame()

if df_raw.empty:
    st.info("👋 welcome. please upload production data to initialize."); st.stop()

with st.sidebar.expander("2. filter scope", expanded=True):
    # primary context: purchase order
    po_list = ["all orders"] + sorted(df_raw['po_number'].dropna().unique().tolist())
    sel_po = st.selectbox("po context", po_list)
    
    scope_df = df_raw.copy()
    if sel_po != "all orders":
        scope_df = scope_df[scope_df['po_number'] == sel_po]

    # hierarchical tree
    projects = sorted(scope_df['project'].dropna().unique().tolist())
    sel_proj = st.multiselect("projects", projects, default=projects)
    scope_df = scope_df[scope_df['project'].isin(sel_proj)]
    
    parts = sorted(scope_df['part_id'].dropna().unique().tolist())
    sel_parts = st.multiselect("part ids", parts, default=parts)
    scope_df = scope_df[scope_df['part_id'].isin(sel_parts)]
    
    tools = sorted(scope_df['tool_id'].dropna().unique().tolist())
    sel_tools = st.multiselect("tooling assets", tools, default=tools)
    scope_df = scope_df[scope_df['tool_id'].isin(sel_tools)]

with st.sidebar.expander("3. capacity config"):
    cal_days = st.number_input("operating days / week", 1, 7, 5)
    cal_hours = st.number_input("operating hours / day", 1, 24, 16)
    cal_config = {'days': cal_days, 'hours': cal_hours}
    tolerance = st.slider("tolerance band (%)", 0.01, 0.20, 0.05)

# --- dashboard logic ---
t_assure, t_risk, t_trends = st.tabs(["📊 supply assurance", "🗼 risk tower", "📈 trends"])
e_config = {'tolerance': tolerance, 'run_interval_hours': 8}

if not scope_df.empty:
    engine = utils.CapacityRiskCalculator(scope_df, e_config)
    res = engine.results
    agg_df = utils.get_supply_metrics(scope_df, e_config)
    
    # align po targets for burn-up and hybrid charts
    po_qty = 0; po_due = datetime.now()
    if not df_po.empty and sel_po != "all orders":
        po_match = df_po[(df_po['po_number'] == sel_po) & (df_po['part_id'].isin(sel_parts))]
        if not po_match.empty:
            po_qty = po_match['total_qty'].sum()
            po_due = po_match['due_date'].max()
            # distribute demand across active weeks
            agg_df['target'] = po_qty / max(1, len(agg_df))

# --- tab 1: supply assurance ---
with t_assure:
    if sel_po == "all orders":
        st.warning("⚠️ select a specific po in the sidebar to view supply assurance."); st.stop()
    
    st.header(f"po fulfillment: {sel_po}")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("actual parts", f"{res['actual_output']:,.0f}")
    c2.metric("target goal", f"{po_qty:,.0f}")
    fulfillment = (res['actual_output'] / po_qty * 100) if po_qty > 0 else 0
    c3.metric("fulfillment", f"{fulfillment:.1f}%")
    
    actual_hrs = res['total_runtime_sec'] / 3600
    planned_hrs = cal_days * cal_hours * (agg_df['period'].nunique())
    stress = (actual_hrs / planned_hrs * 100) if planned_hrs > 0 else 0
    c4.metric("tool stress", f"{stress:.1f}%", delta=f"{stress-100:.1f}% vs plan", delta_color="inverse")

    st.markdown("---")
    st.plotly_chart(utils.create_hybrid_chart(agg_df, cal_config), use_container_width=True)
    st.info(f"**operational summary:** {utils.generate_capacity_insights(res)}", icon="ℹ️")

    g1, g2, g3 = st.columns(3)
    with g1: st.plotly_chart(utils.create_modern_gauge(fulfillment, "po achievement"), use_container_width=True)
    with g2: st.plotly_chart(utils.plot_waterfall(res), use_container_width=True)
    with g3: 
        eff = (res['normal_shots'] / res['total_shots'] * 100) if res['total_shots'] > 0 else 0
        st.plotly_chart(utils.create_modern_gauge(eff, "process efficiency"), use_container_width=True)

    c_left, c_right = st.columns([2, 1])
    with c_left: st.plotly_chart(utils.create_po_burnup(agg_df, po_qty, po_due), use_container_width=True)
    with c_right: st.plotly_chart(utils.create_time_donut(res['total_runtime_sec'], res['actual_output'] * scope_df['approved_ct'].mean(), res['downtime_sec']), use_container_width=True)

    with st.expander("🔬 detailed shot-by-shot analysis"):
        st.plotly_chart(utils.plot_shot_analysis(res['processed_df']), use_container_width=True)

# --- tab 2: risk tower ---
with t_risk:
    st.header("projected supply risk")
    pivot = st.radio("group analysis by:", ["part_id", "project", "tool_id"], horizontal=True)
    risk_stats = scope_df.groupby(pivot).agg({
        'shot_time': 'count', 
        'actual_ct': 'mean', 
        'tool_id': 'nunique',
        'working_cavities': 'mean'
    }).reset_index()
    st.dataframe(risk_stats, use_container_width=True)

# --- tab 3: trends ---
with t_trends:
    st.subheader("weekly performance log")
    st.dataframe(agg_df, use_container_width=True)