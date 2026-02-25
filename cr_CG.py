import streamlit as st
import pandas as pd
import numpy as np
import cr_CG_utils as utils
from datetime import datetime, timedelta
import importlib

# Ensure fresh logic from the utility file
importlib.reload(utils)

st.set_page_config(layout="wide", page_title="Supply Control Tower", page_icon="🏭")

# Professional styling to match the high-density Control Tower aesthetic
st.markdown("""
    <style>
    .main { background-color: #0E1117; }
    .stMetric { background-color: #1E1E26; padding: 15px; border-radius: 10px; border: 1px solid #41424C; }
    [data-testid="stExpander"] { border: 1px solid #41424C; border-radius: 10px; background-color: #161B22; }
    .status-card { background-color: #1E1E26; padding: 20px; border-radius: 10px; border-left: 5px solid #3498DB; margin-bottom: 20px; }
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

with st.sidebar.expander("🔍 2. Supply Hierarchy", expanded=True):
    # Cascade Level 1: PO Context
    pos = ["All POs"] + sorted(df_raw['po_number'].dropna().unique().tolist())
    sel_po = st.selectbox("Purchase Order", pos)
    f_df = df_raw.copy()
    if sel_po != "All POs": f_df = f_df[f_df['po_number'] == sel_po]
    
    # Cascade Level 2: Project
    projs = ["All Projects"] + sorted(f_df['project'].dropna().unique().tolist())
    sel_proj = st.selectbox("Project", projs)
    if sel_proj != "All Projects": f_df = f_df[f_df['project'] == sel_proj]
    
    # Cascade Level 3: Component (Safeguarded)
    if 'component_id' in f_df.columns:
        comps = ["All Components"] + sorted(f_df['component_id'].dropna().unique().tolist())
        sel_comp = st.selectbox("Component", comps)
        if sel_comp != "All Components": f_df = f_df[f_df['component_id'] == sel_comp]
    else:
        st.caption("No Component data found in source.")

    # Cascade Level 4: Part ID
    parts = ["All Parts"] + sorted(f_df['part_id'].dropna().unique().tolist())
    sel_part = st.selectbox("Part Number", parts)
    if sel_part != "All Parts": f_df = f_df[f_df['part_id'] == sel_part]
    
    # Cascade Level 5: Individual Tooling Asset
    tools = sorted(f_df['tool_id'].unique().tolist())
    sel_tools = st.multiselect("Specific Tool(s)", tools, default=tools)
    scope_df = f_df[f_df['tool_id'].isin(sel_tools)]

with st.sidebar.expander("⚙️ 3. Physical Window"):
    cal_days = st.number_input("Operating Days / Week", 1, 7, 5)
    cal_hours = st.number_input("Operating Hours / Day", 1, 24, 16)
    cal_config = {'days': cal_days, 'hours': cal_hours}
    tolerance = st.slider("CT Tolerance Band", 0.01, 0.25, 0.05)

# ==============================================================================
# --- ENGINE & CALCULATION ---
# ==============================================================================
e_config = {'tolerance': tolerance, 'run_interval_hours': 8}

if not scope_df.empty:
    # Run the core Physics Engine
    engine = utils.CapacityRiskCalculator(scope_df, e_config)
    res = engine.results
    
    # Process Supply Targets
    po_qty = 0; po_due = datetime.now()
    if not df_po.empty and sel_po != "All POs":
        po_match = df_po[df_po['po_number'] == sel_po]
        # Filter PO target by Part if selected
        if sel_part != "All Parts":
            po_match = po_match[po_match['part_id'] == sel_part]
        
        if not po_match.empty:
            po_qty = po_match['total_qty'].sum()
            po_due = po_match['due_date'].max()

    # Time-series aggregation for hybrid charts
    agg_df = utils.get_supply_metrics(scope_df, e_config, po_target=po_qty)

# ==============================================================================
# --- MAIN UI DASHBOARD ---
# ==============================================================================
st.title(f"📊 {sel_po if sel_po != 'All POs' else 'Global Supply Overview'}")
st.markdown(f"**Drill-down Path:** {sel_proj} > {sel_part if sel_part != 'All Parts' else 'All Parts'} ({len(sel_tools)} Tools)")

t_supply, t_physics, t_tower = st.tabs([
    "📈 Supply Assurance", 
    "🛠️ Asset Health (Physics)", 
    "🗼 Tower Breakdown"
])

# --- TAB 1: SUPPLY ASSURANCE (THE PO LAYER) ---
with t_supply:
    if sel_po == "All POs":
        st.warning("⚠️ Select a specific PO in the sidebar to view detailed Fulfillment Analytics."); st.stop()

    # High Level Metrics
    k1, k2, k3, k4 = st.columns(4)
    fulfillment = (res['actual_output'] / po_qty * 100) if po_qty > 0 else 0
    k1.metric("Current Fulfillment", f"{fulfillment:.1f}%")
    k2.metric("PO Demand Goal", f"{po_qty:,.0f} units")
    
    # Tool Stress Metric
    act_h = res['total_runtime_sec'] / 3600
    plan_h = cal_days * cal_hours * (agg_df['period'].nunique())
    stress = (act_h / plan_h * 100) if plan_h > 0 else 0
    k3.metric("Tool Stress Factor", f"{stress:.1f}%", delta=f"{stress-100:.1f}% vs Plan", delta_color="inverse")
    k4.metric("Actual Production", f"{res['actual_output']:,.0f}")

    st.markdown("---")
    
    # Hybrid Volume vs Accomplishment/Stress
    st.plotly_chart(utils.create_hybrid_chart(agg_df, cal_config), use_container_width=True, key="s_hybrid")
    
    # Burn-up Path
    st.plotly_chart(utils.create_po_burnup(agg_df, po_qty, po_due), use_container_width=True, key="s_burnup")

# --- TAB 2: ASSET HEALTH (THE ORIGINAL PHYSICS) ---
with t_physics:
    st.subheader("Physical Tool Performance Deep-Dive")
    
    # The Original 3-Gauge Header
    g1, g2, g3 = st.columns(3)
    eff = (res['normal_shots'] / res['total_shots'] * 100) if res['total_shots'] > 0 else 0
    with g1:
        st.plotly_chart(utils.create_modern_gauge(eff, "Process Efficiency"), use_container_width=True, key="p_eff")
    with g2:
        st.plotly_chart(utils.create_modern_gauge(res['stability_index'], "Stability Index"), use_container_width=True, key="p_stab")
    with g3:
        avg_app_ct = scope_df['approved_ct'].mean() if 'approved_ct' in scope_df.columns else scope_df['actual_ct'].mean()
        st.plotly_chart(utils.create_time_donut(res['total_runtime_sec'], res['actual_output'] * avg_app_ct, res['downtime_sec']), use_container_width=True)

    st.markdown("---")
    
    # Loss Waterfall & Stability Drivers
    c_bridge, c_drivers = st.columns([2, 1])
    with c_bridge:
        st.plotly_chart(utils.plot_waterfall(res), use_container_width=True, key="p_bridge")
    with c_drivers:
        st.subheader("Asset Stability Driver")
        st.plotly_chart(utils.create_stability_driver_bar(res['mtbf_min'], res['mttr_min']), use_container_width=True, key="p_driver")
        st.caption(f"**MTBF:** {res['mtbf_min']:.1f} min | **MTTR:** {res['mttr_min']:.1f} min")
        st.info(f"**Insight:** {utils.generate_capacity_insights(res)}")

    # Original Shot-by-Shot Bar Chart
    with st.expander("🔬 View Shot-by-Shot Production Rhythm", expanded=False):
        st.plotly_chart(utils.plot_shot_analysis(res['processed_df']), use_container_width=True, key="p_shots")

# --- TAB 3: TOWER BREAKDOWN ---
with t_tower:
    st.subheader("Hierarchical Risk Analysis")
    
    # Select Pivot based on current selection
    drill = "part_id" if sel_part == "All Parts" else "tool_id"
    if sel_proj == "All Projects": drill = "project"
    
    st.markdown(f"**Drilling Down by:** {drill.replace('_',' ').title()}")
    
    # Aggregation for the Tower table
    agg_map = {'shot_time': 'count', 'actual_ct': 'mean'}
    if 'working_cavities' in scope_df.columns: agg_map['working_cavities'] = 'mean'
    if 'approved_ct' in scope_df.columns: agg_map['approved_ct'] = 'mean'
    
    tower_df = scope_df.groupby(drill).agg(agg_map).reset_index()
    
    # Formatting
    if 'approved_ct' in tower_df.columns:
        tower_df['Cycle Variance %'] = (tower_df['actual_ct'] / tower_df['approved_ct'] - 1) * 100
    
    st.dataframe(tower_df.style.background_gradient(subset=['actual_ct'], cmap='Blues').format({
        'actual_ct': '{:.2f}s',
        'approved_ct': '{:.2f}s' if 'approved_ct' in tower_df.columns else '{:.2f}s',
        'working_cavities': '{:.1f}' if 'working_cavities' in tower_df.columns else '{:.1f}',
        'Cycle Variance %': '{:.1f}%' if 'Cycle Variance %' in tower_df.columns else '{:.1f}%'
    }), use_container_width=True, hide_index=True)
    
    # Export capability
    st.download_button(
        "📥 Download Performance Report (CSV)",
        tower_df.to_csv(index=False),
        f"supply_report_{datetime.now().strftime('%Y%m%d')}.csv",
        "text/csv"
    )