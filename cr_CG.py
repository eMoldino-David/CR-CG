import streamlit as st
import pandas as pd
import numpy as np
import cr_CG_utils as utils
from datetime import datetime, timedelta
import importlib

# Ensure fresh logic
importlib.reload(utils)

st.set_page_config(layout="wide", page_title="Supply Control Tower", page_icon="🏭")

# Professional styling
st.markdown("""
    <style>
    .main { background-color: #0E1117; }
    .stMetric { background-color: #1E1E26; padding: 15px; border-radius: 10px; border: 1px solid #41424C; }
    [data-testid="stExpander"] { border: 1px solid #41424C; border-radius: 10px; background-color: #161B22; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# --- SIDEBAR: DATA & HIERARCHY ---
# ==============================================================================
st.sidebar.title("🏭 Control Tower")

with st.sidebar.expander("📂 1. Data Intake", expanded=True):
    p_files = st.file_uploader("Production Data", accept_multiple_files=True)
    o_files = st.file_uploader("PO Demand Data", accept_multiple_files=True)

df_raw = utils.load_production_data(p_files) if p_files else pd.DataFrame()
df_po = utils.load_po_data(o_files) if o_files else pd.DataFrame()

if df_raw.empty:
    st.info("👋 Please upload production data to initialize the tower."); st.stop()

with st.sidebar.expander("🔍 2. Analysis Hierarchy", expanded=True):
    # PO Filter
    pos = ["All POs"] + sorted(df_raw['po_number'].dropna().unique().tolist())
    sel_po = st.selectbox("Purchase Order", pos)
    f_df = df_raw.copy()
    if sel_po != "All POs": f_df = f_df[f_df['po_number'] == sel_po]
    
    # Project Filter
    projs = ["All Projects"] + sorted(f_df['project'].dropna().unique().tolist())
    sel_proj = st.selectbox("Project", projs)
    if sel_proj != "All Projects": f_df = f_df[f_df['project'] == sel_proj]
    
    # Component Filter (Safeguarded against KeyError)
    if 'component_id' in f_df.columns:
        comps = ["All Components"] + sorted(f_df['component_id'].dropna().unique().tolist())
        sel_comp = st.selectbox("Component", comps)
        if sel_comp != "All Components": f_df = f_df[f_df['component_id'] == sel_comp]
    else:
        st.caption("No Component data found in source.")

    # Part & Tool Filter
    parts = ["All Parts"] + sorted(f_df['part_id'].dropna().unique().tolist())
    sel_part = st.selectbox("Part Number", parts)
    if sel_part != "All Parts": f_df = f_df[f_df['part_id'] == sel_part]
    
    tools = sorted(f_df['tool_id'].unique().tolist())
    sel_tools = st.multiselect("Specific Tool(s)", tools, default=tools)
    scope_df = f_df[f_df['tool_id'].isin(sel_tools)]

with st.sidebar.expander("⚙️ 3. Physical Window"):
    cal_days = st.number_input("Operating Days / Week", 1, 7, 5)
    cal_hours = st.number_input("Operating Hours / Day", 1, 24, 16)
    cal_config = {'days': cal_days, 'hours': cal_hours}
    tolerance = st.slider("CT Tolerance Band", 0.01, 0.25, 0.05)

# ==============================================================================
# --- CALCULATION ENGINE ---
# ==============================================================================
e_config = {'tolerance': tolerance, 'run_interval_hours': 8}
if not scope_df.empty:
    engine = utils.CapacityRiskCalculator(scope_df, e_config)
    res = engine.results
    
    # Supply Targets
    po_qty = 0; po_due = datetime.now()
    if not df_po.empty and sel_po != "All POs":
        po_match = df_po[df_po['po_number'] == sel_po]
        if sel_part != "All Parts": po_match = po_match[po_match['part_id'] == sel_part]
        if not po_match.empty:
            po_qty = po_match['total_qty'].sum()
            po_due = po_match['due_date'].max()

    agg_df = utils.get_supply_metrics(scope_df, e_config)
    if po_qty > 0: agg_df['target'] = po_qty / max(1, len(agg_df))

# ==============================================================================
# --- MAIN UI ---
# ==============================================================================
st.title(f"📊 {sel_po if sel_po != 'All POs' else 'Global Overview'}")
st.markdown(f"**Hierarchy:** {sel_proj} > {sel_part}")

t_physics, t_supply, t_risk = st.tabs(["🛠️ Asset Health", "📈 Supply Assurance", "🗼 Tower Breakdown"])

with t_physics:
    st.subheader("Physical Tool Performance")
    c1, c2, c3 = st.columns(3)
    eff = (res['normal_shots'] / res['total_shots'] * 100) if res['total_shots'] > 0 else 0
    c1.plotly_chart(utils.create_modern_gauge(eff, "Process Efficiency"), use_container_width=True, key="p_eff")
    c2.plotly_chart(utils.create_modern_gauge(res['stability_index'], "Stability Index"), use_container_width=True, key="p_stab")
    
    # Handle potentially missing approved_ct for display
    avg_app_ct = scope_df['approved_ct'].mean() if 'approved_ct' in scope_df.columns else scope_df['actual_ct'].mean()
    c3.plotly_chart(utils.create_time_donut(res['total_runtime_sec'], res['actual_output'] * avg_app_ct, res['downtime_sec']), use_container_width=True)

    st.markdown("---")
    cw, ci = st.columns([2, 1])
    cw.plotly_chart(utils.plot_waterfall(res), use_container_width=True, key="p_bridge")
    ci.metric("Actual Parts", f"{res['actual_output']:,.0f}")
    ci.metric("Downtime Loss", f"{res['loss_dt']:,.0f} units")
    ci.metric("Slow Cycle Loss", f"{res['loss_slow']:,.0f} units")

with t_supply:
    if sel_po == "All POs":
        st.warning("⚠️ Select a specific PO to view fulfillment details.")
    else:
        st.subheader(f"Supply Commitment: {sel_po}")
        k1, k2, k3 = st.columns(3)
        fulfillment = (res['actual_output'] / po_qty * 100) if po_qty > 0 else 0
        k1.metric("Fulfillment Status", f"{fulfillment:.1f}%")
        k2.metric("PO Demand", f"{po_qty:,.0f} units")
        
        act_h = res['total_runtime_sec'] / 3600
        plan_h = cal_days * cal_hours * (agg_df['period'].nunique())
        stress = (act_h / plan_h * 100) if plan_h > 0 else 0
        k3.metric("Asset Stress Factor", f"{stress:.1f}%", delta=f"{stress-100:.1f}% vs Plan", delta_color="inverse")

        st.plotly_chart(utils.create_hybrid_chart(agg_df, cal_config), use_container_width=True, key="s_hybrid")
        st.plotly_chart(utils.create_po_burnup(agg_df, po_qty, po_due), use_container_width=True, key="s_burnup")

with t_risk:
    st.subheader("Hierarchical Risk Analysis")
    drill = "part_id" if sel_part == "All Parts" else "tool_id"
    if sel_proj == "All Projects": drill = "project"
    
    # Defensive aggregation to handle missing columns
    agg_map = {'shot_time': 'count', 'actual_ct': 'mean'}
    if 'working_cavities' in scope_df.columns: agg_map['working_cavities'] = 'mean'
    
    tower_df = scope_df.groupby(drill).agg(agg_map).reset_index()
    st.dataframe(tower_df, use_container_width=True, hide_index=True)