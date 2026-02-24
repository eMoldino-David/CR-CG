import streamlit as st
import pandas as pd
import numpy as np
import cr_CG_utils as utils
from datetime import datetime, timedelta

st.set_page_config(layout="wide", page_title="Supply Control Tower")

# ==============================================================================
# --- SIDEBAR: DATA INTAKE ---
# ==============================================================================
st.sidebar.title("📦 Supply Assurance v1.0")

with st.sidebar.expander("1. Upload Source Files", expanded=True):
    prod_files = st.file_uploader("Upload Production Data", accept_multiple_files=True)
    po_files = st.file_uploader("Upload PO Planning Data", accept_multiple_files=True)

df_raw = utils.load_production_data(prod_files) if prod_files else pd.DataFrame()
df_po = utils.load_po_data(po_files) if po_files else pd.DataFrame()

if df_raw.empty:
    st.info("👋 Welcome! Please upload your production shot data in the sidebar to begin."); st.stop()

# ==============================================================================
# --- SIDEBAR: CASCADING HIERARCHY ---
# ==============================================================================
with st.sidebar.expander("2. Filter Scope", expanded=True):
    # Top Level: PO Number
    po_list = ["All POs"] + sorted(df_raw['po_number'].dropna().unique().tolist())
    sel_po = st.selectbox("Select Target PO", po_list)
    
    # Cascade
    scope_df = df_raw.copy()
    if sel_po != "All POs":
        scope_df = scope_df[scope_df['po_number'] == sel_po]

    projects = sorted(scope_df['project'].dropna().unique().tolist())
    sel_proj = st.multiselect("Project", projects, default=projects)
    scope_df = scope_df[scope_df['project'].isin(sel_proj)]
    
    parts = sorted(scope_df['part_id'].dropna().unique().tolist())
    sel_parts = st.multiselect("Part ID", parts, default=parts)
    scope_df = scope_df[scope_df['part_id'].isin(sel_parts)]
    
    tools = sorted(scope_df['tool_id'].dropna().unique().tolist())
    sel_tools = st.multiselect("Tooling Assets", tools, default=tools)
    scope_df = scope_df[scope_df['tool_id'].isin(sel_tools)]

# ==============================================================================
# --- SIDEBAR: PLANNING CONFIG ---
# ==============================================================================
with st.sidebar.expander("3. Working Calendar (Stress Config)"):
    cal_days = st.number_input("Operating Days / Week", 1, 7, 5)
    cal_hours = st.number_input("Operating Hours / Day", 1, 24, 16)
    cal_config = {'days': cal_days, 'hours': cal_hours}

# ==============================================================================
# --- MAIN DASHBOARD ---
# ==============================================================================
t_assurance, t_risk = st.tabs(["📊 Supply Assurance", "🗼 Risk Tower"])

# Global Engine Settings
e_config = {'tolerance': 0.05, 'run_interval_hours': 8}

with t_assurance:
    if sel_po == "All POs":
        st.warning("⚠️ Please select a specific Purchase Order to view Fulfillment analysis."); st.stop()

    # Aggregate Data
    agg_df = utils.get_supply_metrics(scope_df, 'Weekly', e_config)
    
    # Join Demand from PO File
    po_target_qty = 0
    po_due_date = datetime.now()
    if not df_po.empty:
        po_match = df_po[(df_po['po_number'] == sel_po) & (df_po['part_id'].isin(sel_parts))]
        if not po_match.empty:
            po_target_qty = po_match['total_qty'].sum()
            po_due_date = po_match['due_date'].max()
            # Distribute target for the weekly bars
            agg_df['Target'] = po_target_qty / max(1, len(agg_df))

    # Top KPI Row
    c1, c2, c3 = st.columns(3)
    total_output = agg_df['Actual Output'].sum()
    c1.metric("Actual Output", f"{total_output:,.0f} units")
    
    accomplishment = (total_output / po_target_qty * 100) if po_target_qty > 0 else 0
    c2.metric("PO Accomplishment", f"{accomplishment:.1f}%", delta=f"{total_output - po_target_qty:,.0f} vs Goal")
    
    # Stress KPI
    actual_hrs = agg_df['Run Time Sec'].sum() / 3600
    planned_hrs = cal_days * cal_hours * len(agg_df)
    stress_pct = (actual_hrs / planned_hrs * 100) if planned_hrs > 0 else 0
    c3.metric("Tool Stress Factor", f"{stress_pct:.1f}%", delta_color="inverse")


    # Main Visuals
    st.markdown("---")
    st.plotly_chart(utils.create_hybrid_chart(agg_df, cal_config), use_container_width=True)
    
    with st.expander("🔍 View Fulfillment Burn-up"):
        st.plotly_chart(utils.create_burnup(agg_df, po_target_qty, po_due_date), use_container_width=True)

with t_risk:
    st.header("Strategic Risk Tower")
    st.markdown("This view groups risk by the current filter selection.")
    
    # Simplified Risk grouping
    if not scope_df.empty:
        risk_summary = scope_df.groupby('part_id').agg({
            'actual_ct': 'mean',
            'tool_id': 'nunique'
        }).reset_index()
        st.dataframe(risk_summary, use_container_width=True)