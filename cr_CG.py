import streamlit as st
import pandas as pd
import numpy as np
import cr_CG_utils as utils
from datetime import datetime, timedelta
import importlib
import io

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

# Custom CSS for a professional dark-themed look and high-density UI
st.markdown("""
    <style>
    .main { background-color: #0E1117; }
    .stMetric { background-color: #1E1E26; padding: 15px; border-radius: 10px; border: 1px solid #41424C; }
    [data-testid="stExpander"] { border: 1px solid #41424C; border-radius: 10px; background-color: #161B22; }
    div[data-testid="stExpander"] p { font-size: 0.85rem; }
    .reportview-container .main .block-container { padding-top: 2rem; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #3498DB; color: white; }
    .status-card { background-color: #1E1E26; padding: 20px; border-radius: 10px; border-left: 5px solid #3498DB; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# --- SIDEBAR: DATA INTAKE & CONFIG ---
# ==============================================================================
st.sidebar.title("📦 Control Tower v2.0")

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
# --- GLOBAL ANALYTICS PRE-PROCESSING ---
# ==============================================================================
e_config = {'tolerance': tolerance, 'run_interval_hours': 8}

if not scope_df.empty:
    # Run the core calculation engine for the filtered scope
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

# ==============================================================================
# --- MAIN INTERFACE: TABS ---
# ==============================================================================
t_assure, t_risk, t_compare, t_sim, t_trends = st.tabs([
    "📊 Supply Assurance", 
    "🗼 Risk Tower", 
    "⚖️ Asset Comparison", 
    "🔮 Capacity Simulator", 
    "📈 Trend Analysis"
])

# --- TAB 1: SUPPLY ASSURANCE ---
with t_assure:
    if sel_po == "All Orders":
        st.warning("⚠️ Select a specific Purchase Order in the sidebar to unlock Fulfillment Analytics.")
        st.stop()

    st.header(f"Fulfillment Analysis: {sel_po}")
    
    # Summary Insight Card
    st.markdown(f"""
    <div class="status-card">
        <h3>Operational Summary</h3>
        <p>{utils.generate_capacity_insights(res)}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Top KPI Row
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Actual Production", f"{res['actual_output']:,.0f}")
    k2.metric("PO Target Goal", f"{po_target_qty:,.0f}")
    
    fulfillment = (res['actual_output'] / po_target_qty * 100) if po_target_qty > 0 else 0
    k3.metric("Fulfillment Status", f"{fulfillment:.1f}%")
    
    actual_hrs = res['total_runtime_sec'] / 3600
    planned_hrs = cal_days * cal_hours * (agg_df['period'].nunique())
    stress_pct = (actual_hrs / planned_hrs * 100) if planned_hrs > 0 else 0
    k4.metric("Tool Stress Factor", f"{stress_pct:.1f}%", delta=f"{stress_pct-100:.1f}% vs Plan", delta_color="inverse")

    st.markdown("---")
    st.plotly_chart(utils.create_hybrid_chart(agg_df, cal_config), use_container_width=True)

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
    st.markdown("Identifying bottlenecks and systemic risks across the supply chain.")
    
    pivot_choice = st.radio("Pivot Risk Analysis By:", ["Part ID", "Project", "Tool ID"], horizontal=True)
    pivot_col = pivot_choice.lower().replace(" ","_")
    
    # Advanced Risk Aggregation
    risk_stats = scope_df.groupby(pivot_col).agg({
        'shot_time': 'count',
        'actual_ct': 'mean',
        'approved_ct': 'mean',
        'tool_id': 'nunique',
        'working_cavities': 'mean'
    }).reset_index()
    
    # Calculate Risk Score: (Actual CT / Approved CT) * (1 / Stability) - simplistic proxy
    risk_stats['Performance Gap'] = (risk_stats['actual_ct'] / risk_stats['approved_ct'] - 1) * 100
    risk_stats = risk_stats.rename(columns={
        'shot_time': 'Total Shots',
        'actual_ct': 'Avg CT (s)',
        'tool_id': 'Unique Assets',
        'working_cavities': 'Avg Cavities'
    })
    
    st.dataframe(risk_stats.style.background_gradient(subset=['Performance Gap'], cmap='Reds').format({
        'Avg CT (s)': '{:.2f}', 
        'Avg Cavities': '{:.1f}',
        'Performance Gap': '{:.1f}%'
    }), use_container_width=True, hide_index=True)

# --- TAB 3: ASSET COMPARISON ---
with t_compare:
    st.header("⚖️ Asset Side-by-Side Comparison")
    if len(sel_tools) < 2:
        st.warning("Please select at least two Tooling Assets in the sidebar to enable comparison.")
    else:
        comp_col1, comp_col2 = st.columns(2)
        
        # Tool A Analysis
        with comp_col1:
            tool_a = st.selectbox("Select Asset A", sel_tools, index=0)
            df_a = scope_df[scope_df['tool_id'] == tool_a]
            res_a = utils.CapacityRiskCalculator(df_a, e_config).results
            st.subheader(f"Asset: {tool_a}")
            st.metric("Avg Cycle Time", f"{res_a['processed_df']['actual_ct'].mean():.2f}s")
            st.plotly_chart(utils.create_modern_gauge(res_a['stability_index'], "Stability"), use_container_width=True)
            st.plotly_chart(utils.plot_waterfall(res_a), use_container_width=True)

        # Tool B Analysis
        with comp_col2:
            tool_b = st.selectbox("Select Asset B", sel_tools, index=1)
            df_b = scope_df[scope_df['tool_id'] == tool_b]
            res_b = utils.CapacityRiskCalculator(df_b, e_config).results
            st.subheader(f"Asset: {tool_b}")
            st.metric("Avg Cycle Time", f"{res_b['processed_df']['actual_ct'].mean():.2f}s")
            st.plotly_chart(utils.create_modern_gauge(res_b['stability_index'], "Stability"), use_container_width=True)
            st.plotly_chart(utils.plot_waterfall(res_b), use_container_width=True)

# --- TAB 4: WHAT-IF SIMULATOR ---
with t_sim:
    st.header("🔮 Capacity Simulator")
    st.markdown("Simulate how changes in tool physics or schedule impact your ability to meet the PO.")
    
    sim_col1, sim_col2 = st.columns([1, 2])
    
    with sim_col1:
        st.subheader("Simulation Parameters")
        sim_ct = st.slider("Simulated Cycle Time (s)", 1.0, 120.0, float(scope_df['approved_ct'].mean()))
        sim_cav = st.number_input("Simulated Cavities", 1, 32, int(scope_df['working_cavities'].max()))
        sim_days = st.slider("Remaining Days to Deadline", 1, 60, 14)
        sim_eff = st.slider("Expected Efficiency (%)", 10, 100, 85)
        
    with sim_col2:
        st.subheader("Projection Results")
        # Calc: (Total Seconds Available) / CT * Cavities * Efficiency
        total_seconds = sim_days * cal_hours * 3600
        potential_output = (total_seconds / sim_ct) * sim_cav * (sim_eff / 100)
        
        needed_for_po = max(0, po_target_qty - res['actual_output'])
        
        st.metric("Potential Future Output", f"{potential_output:,.0f} units")
        st.metric("Remaining PO Gap", f"{needed_for_po:,.0f} units")
        
        if potential_output >= needed_for_po:
            st.success(f"✅ Feasible: You are projected to exceed the gap by {potential_output - needed_for_po:,.0f} units.")
        else:
            st.error(f"🚨 At Risk: You will likely fall short by {needed_for_po - potential_output:,.0f} units.")

# --- TAB 5: TREND ANALYSIS & EXPORT ---
with t_trends:
    st.header("Historical Performance Data")
    
    # Export Section
    csv = agg_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Processed Weekly Data (CSV)",
        data=csv,
        file_name=f"control_tower_export_{datetime.now().strftime('%Y%m%d')}.csv",
        mime='text/csv',
    )
    
    st.markdown("---")
    st.subheader("Raw Aggregated Logs")
    st.dataframe(agg_df.style.format({
        'actual_output': '{:,.0f}',
        'runtime_sec': '{:,.0f}',
        'target': '{:,.0f}',
        'stability': '{:.1f}%'
    }), use_container_width=True, hide_index=True)