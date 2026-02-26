import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import cr_CG_utils as cr_CG_utils
import importlib

# Force reload of utils to ensure latest logic is used
importlib.reload(cr_CG_utils)

# ==============================================================================
# --- 🔒 SECURITY: Initial LOGIN ---
# ==============================================================================
# This stops the app from loading ANY data until the password is correct.

def check_password():
    """Returns `True` if the user had the correct password."""
    if st.session_state.get("password_correct", False):
        return True

    st.header("🔒 Protected Internal Tool")
    password_input = st.text_input("Enter Company Password", type="password")
    
    if password_input:
        if password_input == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            st.rerun()  
        else:
            st.error("😕 Password incorrect")
            
    return False

if not check_password():
    st.stop()  



# ==============================================================================
# --- PAGE CONFIG ---
# ==============================================================================
st.set_page_config(layout="wide", page_title="Capacity Risk Dashboard (v10.6)")

# ==============================================================================
# --- HELPER FUNCTIONS ---
# ==============================================================================

def display_filter_context(ctx, tool_name=None):
    """Displays a clear banner indicating exactly what data is currently filtered and active."""
    if not ctx:
        tool_str = f" | **Tool:** {tool_name}" if tool_name and tool_name != 'Multiple' else ""
        st.info(f"🗂️ **Current Filter Scope:** All Data{tool_str}")
        return
        
    active_filters = [f"**{k}:** {v}" for k, v in ctx.items() if v != "All"]
    if tool_name and tool_name != "Multiple":
        active_filters.append(f"**Tool:** {tool_name}")
        
    if active_filters:
        st.info("🗂️ **Current Filter Scope:** " + " | ".join(active_filters))
    else:
        st.info(f"🗂️ **Current Filter Scope:** All Global Data")

def create_capsule(value, color_logic="neutral", suffix="", inverse=False):
    """
    Generates an HTML string for a styled pill/capsule.
    """
    bg_color = "#262730" 
    text_color = "#ffffff"
    
    if color_logic == "grey":
        bg_color = "#41424C" 
        text_color = "#ffffff"
    elif color_logic == "good_bad":
        if value >= 90: bg_color = cr_CG_utils.PASTEL_COLORS['green']; text_color = "#0E1117"
        elif value >= 75: bg_color = cr_CG_utils.PASTEL_COLORS['orange']; text_color = "#0E1117"
        else: bg_color = cr_CG_utils.PASTEL_COLORS['red']; text_color = "#0E1117"
    elif color_logic == "bad_good":
        if value <= 10: bg_color = cr_CG_utils.PASTEL_COLORS['green']; text_color = "#0E1117"
        elif value <= 20: bg_color = cr_CG_utils.PASTEL_COLORS['orange']; text_color = "#0E1117"
        else: bg_color = cr_CG_utils.PASTEL_COLORS['red']; text_color = "#0E1117"
    elif color_logic == "net":
        if value >= 0: bg_color = cr_CG_utils.PASTEL_COLORS['green']; text_color = "#0E1117"
        else: bg_color = cr_CG_utils.PASTEL_COLORS['red']; text_color = "#0E1117"

    return f'<span style="background-color:{bg_color}; color:{text_color}; padding:2px 8px; border-radius:10px; font-weight:bold; font-size:0.8em;">{value:,.1f}{suffix}</span>'

# ==============================================================================
# --- 1. RENDER FUNCTIONS ---
# ==============================================================================

def render_risk_tower(df_all_tools, config, filter_context):
    """Renders the Risk Tower (Tab 1)."""
    st.title("Capacity Risk Tower")
    display_filter_context(filter_context, "Multiple Tools (Rolled-Up)")
    st.info("This tower identifies tools at risk by analyzing weekly production gaps over the last 4 weeks.")

    with st.expander("ℹ️ How the Risk Tower Works"):
        st.markdown("""
        The Risk Tower evaluates each tool based on its performance over its own most recent 4-week period.
        
        ### 1. Risk Score (0-100)
        Represents the **Average Capacity Achievement %** over the analysis period.
        - **Formula:** `Average(Actual Output / Target Output)` across the active weeks.
        - **Goal:** A score of 95+ indicates the tool is consistently meeting capacity demand.
        
        ### 2. Primary Risk Factor
        Identifies the dominant root cause preventing the tool from hitting 100% capacity.
        - **Downtime:** The majority of lost parts are due to machine stops (Run Rate Downtime).
        - **Cycle Time:** The majority of lost parts are due to slow cycles (Running above Ideal CT).
        - **Stable:** The tool is operating above 95% achievement; no significant risk detected.
        
        ### 3. Achievement Trend
        Displays the weekly progression of Capacity Achievement to highlight stability.
        - Shows: `Week 1 % → Week 2 % → Week 3 % → Week 4 %`
        - Helps identify if performance is improving, degrading, or fluctuating wildly.
        
        ### 4. Details
        Provides specific context on the magnitude of the risk.
        - Displays the total **Net Parts Lost** attributed to the Primary Risk Factor.
        """, unsafe_allow_html=True)

    results = []
    tools = sorted(df_all_tools['tool_id'].unique())
    
    for tool_id in tools:
        tool_df = df_all_tools[df_all_tools['tool_id'] == tool_id]
        
        weekly_df = cr_CG_utils.get_aggregated_data(tool_df, 'Weekly', config)
        
        if weekly_df.empty:
            continue
            
        recent = weekly_df.tail(4).copy()
        
        cols_needed = ['Actual Output', 'Target Output', 'Downtime Loss', 'Slow Loss']
        for c in cols_needed:
            if c not in recent.columns: recent[c] = 0
            
        recent['Target Output'] = recent['Target Output'].replace(0, 1)
        recent['Achieve %'] = (recent['Actual Output'] / recent['Target Output'] * 100).fillna(0)
        
        trend_str = " → ".join([f"{x:.0f}%" for x in recent['Achieve %']])
        
        avg_achieve = recent['Achieve %'].mean()
        risk_score = min(avg_achieve, 100)
        
        total_dt_loss = recent['Downtime Loss'].sum()
        total_slow_loss = recent['Slow Loss'].sum()
        net_gap = recent['Target Output'].sum() - recent['Actual Output'].sum()
        
        risk_factor = "Stable"
        details = f"Running well. Overall achievement is {avg_achieve:.1f}%."
        
        if avg_achieve < 95:
            if total_dt_loss > total_slow_loss:
                risk_factor = "Downtime"
                details = f"Primary loss driver is Downtime ({total_dt_loss:,.0f} parts lost)."
            elif total_slow_loss > 0:
                risk_factor = "Cycle Time"
                details = f"Primary loss driver is Slow Cycles ({total_slow_loss:,.0f} parts lost)."
            else:
                risk_factor = "Unspecified"
                details = f"Output is below target by {net_gap:,.0f} parts."
        
        p_min = recent['Period'].min()
        p_max = recent['Period'].max()
        if isinstance(p_min, pd.Period): p_min = p_min.start_time.date()
        if isinstance(p_max, pd.Period): p_max = p_max.start_time.date() 

        results.append({
            "Tool ID": tool_id,
            "Analysis Period": f"{p_min} to {p_max}",
            "Risk Score": risk_score,
            "Primary Risk Factor": risk_factor,
            "Achievement Trend": trend_str,
            "Details": details
        })

    if not results:
        st.warning("Not enough data to generate the Risk Tower.")
        return

    risk_df = pd.DataFrame(results)

    def style_risk_tower(row):
        score = row['Risk Score']
        styles = [''] * len(row)
        
        if score >= 90: base_color = cr_CG_utils.PASTEL_COLORS['green']
        elif score >= 75: base_color = cr_CG_utils.PASTEL_COLORS['orange']
        else: base_color = cr_CG_utils.PASTEL_COLORS['red']
        
        return [f'background-color: {base_color}; color: black' for _ in row]

    st.dataframe(
        risk_df.style.apply(style_risk_tower, axis=1)
        .format({'Risk Score': '{:.0f}'}),
        use_container_width=True, 
        hide_index=True
    )

def render_trends_tab(df_tool, tool_name, config, filter_context):
    """Renders the Trends Tab."""
    st.header("Historical Performance Trends")
    display_filter_context(filter_context, tool_name)
    st.info("Trends are calculated using 'Run-Based' logic consistent with the Dashboard.")

    col_freq, col_mode, _ = st.columns([1, 1, 2])
    with col_freq:
        trend_freq = st.selectbox("Select Trend Frequency", ["Daily", "Weekly", "Monthly"], key="cr_trend_freq")
    with col_mode:
        trend_mode = st.selectbox("Dashboard Mode", ["Optimal", "Target"], key="cr_trend_mode")

    agg_df = cr_CG_utils.get_aggregated_data(df_tool, trend_freq, config)
    
    if agg_df.empty:
        st.warning("No trend data available.")
        return

    col_map = {
        'Run Time': 'Total Run Duration',
        'Downtime': 'Run Rate Downtime',
        'Production Time': 'Production Time'
    }
    agg_df = agg_df.rename(columns=col_map)

    display_cols = ['Period', 'Total Run Duration', 'Run Rate Downtime']
    
    if trend_mode == "Optimal":
        opt_cols = ['Actual Output', 'Optimal Output', 'Total Loss', 'Downtime Loss', 'Slow Loss', 'Fast Gain']
        display_cols.extend([c for c in opt_cols if c in agg_df.columns])
    else:
        if 'Target Output' in agg_df.columns:
            agg_df['Net Diff (vs Target)'] = agg_df['Actual Output'] - agg_df['Target Output']
            display_cols.extend(['Actual Output', 'Target Output', 'Net Diff (vs Target)'])

    view_df = agg_df[display_cols].copy()

    def style_trends(row):
        styles = [''] * len(row)
        for i, col in enumerate(view_df.columns):
            val = row[col]
            if isinstance(val, (int, float)):
                if 'Loss' in col and 'Net' not in col: 
                    if val > 0: styles[i] = 'color: #ff6961' 
                elif 'Gain' in col:
                    if val > 0: styles[i] = 'color: #77dd77' 
                elif 'Net' in col or 'Diff' in col:
                    if val < 0: styles[i] = 'color: #ff6961' 
                    elif val > 0: styles[i] = 'color: #77dd77' 
        return styles

    st.dataframe(
        view_df.style.apply(style_trends, axis=1).format(precision=1), 
        use_container_width=True, 
        hide_index=True
    )

    st.subheader("Visual Trend")
    metric_to_plot = st.selectbox("Select Metric to Visualize", 
                                  ['Actual Output', 'Optimal Output', 'Target Output', 'Total Loss', 'Total Run Duration', 'Run Rate Downtime'],
                                  key="cr_trend_viz_select")
    
    if metric_to_plot in agg_df.columns:
        fig = px.line(agg_df, x='Period', y=metric_to_plot, markers=True, title=f"{metric_to_plot} Trend")
        
        if "Output" in metric_to_plot:
            if metric_to_plot != "Actual Output":
                 fig.add_trace(go.Scatter(x=agg_df['Period'], y=agg_df['Actual Output'], mode='lines+markers', name='Actual Output', line=dict(dash='dot', color='blue')))
            if metric_to_plot != "Target Output" and "Target Output" in agg_df.columns:
                 fig.add_trace(go.Scatter(x=agg_df['Period'], y=agg_df['Target Output'], mode='lines', name='Target Output', line=dict(color='green', width=1)))
            if metric_to_plot != "Optimal Output" and "Optimal Output" in agg_df.columns:
                 fig.add_trace(go.Scatter(x=agg_df['Period'], y=agg_df['Optimal Output'], mode='lines', name='Optimal Output', line=dict(color='orange', width=1, dash='dash')))
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning(f"Metric {metric_to_plot} not found in data.")

def render_forecast_tab(df_scope, config, df_logistics, working_days_per_week, working_hours_per_day, filter_context):
    """Renders the PO Forecast & Burn-up Tab using dynamic configuration panels."""
    st.header("Logistics Plan Tracking & Forecast")
    display_filter_context(filter_context, "Multiple Tools (Rolled-up)")
    
    # Check if we have PO data mapping
    has_po_in_shots = 'po_number' in df_scope.columns and not df_scope['po_number'].replace("Unknown", pd.NA).isna().all()
    
    if not df_logistics.empty and has_po_in_shots:
        part_pos = df_scope['po_number'].unique()
        avail_pos = df_logistics[df_logistics['po_number'].isin(part_pos)]['po_number'].unique()
        
        if len(avail_pos) == 0:
            st.warning("No matching POs found between Logistics Plan and Production Data for the current filter scope.")
            return
            
        # --- Advanced Tracking Configuration ---
        st.markdown("### ⚙️ Tracking Configuration")
        track_mode = st.radio("Group & Track Progress By:", ["Purchase Order(s)", "Supplier(s)", "Plant(s)"], horizontal=True)
        
        selected_po_list = []
        
        if track_mode == "Purchase Order(s)":
            selected_pos = st.multiselect("Select Purchase Order(s) to Track", avail_pos, default=avail_pos[:1])
            selected_po_list = selected_pos
            
        elif track_mode == "Supplier(s)":
            avail_sups = [s for s in df_scope['supplier_id'].unique() if str(s).lower() not in ['unknown', 'nan']]
            if not avail_sups: st.warning("No identified Supplier data in scope."); return
            selected_sups = st.multiselect("Select Supplier(s) to Track", avail_sups, default=avail_sups)
            linked_pos = df_scope[df_scope['supplier_id'].isin(selected_sups)]['po_number'].unique()
            selected_po_list = [po for po in linked_pos if po in avail_pos]
            
        elif track_mode == "Plant(s)":
            avail_plts = [p for p in df_scope['plant_id'].unique() if str(p).lower() not in ['unknown', 'nan']]
            if not avail_plts: st.warning("No identified Plant data in scope."); return
            selected_plts = st.multiselect("Select Plant(s) to Track", avail_plts, default=avail_plts)
            linked_pos = df_scope[df_scope['plant_id'].isin(selected_plts)]['po_number'].unique()
            selected_po_list = [po for po in linked_pos if po in avail_pos]
            
        if not selected_po_list:
            st.warning(f"No Purchase Orders are associated with your current {track_mode} selection. Please select at least one item.")
            return
            
        # Aggregate the Logistics PO records safely into a composite
        subset_logistics = df_logistics[df_logistics['po_number'].isin(selected_po_list)]
        df_po_shots = df_scope[df_scope['po_number'].isin(selected_po_list)].copy()
        
        total_qty = pd.to_numeric(subset_logistics['total_qty'], errors='coerce').sum()
        min_start = pd.to_datetime(subset_logistics['start_date']).min()
        max_due = pd.to_datetime(subset_logistics['due_date']).max()
        
        po_display_name = ", ".join(selected_po_list) if len(selected_po_list) <= 3 else f"{len(selected_po_list)} POs Selected"
        proj_display = ", ".join(subset_logistics['project_id'].astype(str).unique())
        part_display = ", ".join(subset_logistics['part_id'].astype(str).unique())
        
        composite_po_record = {
            'po_number': po_display_name,
            'total_qty': total_qty,
            'start_date': min_start,
            'due_date': max_due
        }
        
        # --- PO Summary Box ---
        with st.container(border=True):
            st.markdown(f"### 📋 {track_mode} Summary Details")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Tracked POs", po_display_name)
            col2.metric("Project / Part", f"{proj_display[:20]} / {part_display[:20]}")
            col3.metric("Total Target Quantity", f"{total_qty:,.0f}")
            
            involved_tools = df_po_shots['tool_id'].unique()
            tools_str = ", ".join([str(t) for t in involved_tools])
            col4.metric("Assigned Toolings", tools_str if len(tools_str) < 40 else f"{len(involved_tools)} Tools")
            
            c1, c2 = st.columns(2)
            c1.write(f"**Earliest Start Date:** {min_start.strftime('%Y-%m-%d') if pd.notnull(min_start) else 'N/A'}")
            c2.write(f"**Latest Due Date:** {max_due.strftime('%Y-%m-%d') if pd.notnull(max_due) else 'N/A'}")

        st.markdown("---")

        # --- GRAPH 1: Periodic Breakdown vs Demand ---
        st.markdown(f"### 1. Periodic Production vs Estimated Demand")
        col_freq, col_tool = st.columns([1, 2])
        with col_freq:
            bar_freq = st.selectbox("Select Frequency Spread", ["Weekly", "Monthly", "Daily"], index=0)
        
        with col_tool:
            if len(involved_tools) > 1:
                options = ["All Tracked Tools Combined"] + list(involved_tools)
                selected_tool_for_bar = st.radio("Select View Segment", options, horizontal=True)
            else:
                selected_tool_for_bar = "All Tracked Tools Combined"
        
        df_bar_view = df_po_shots if selected_tool_for_bar == "All Tracked Tools Combined" else df_po_shots[df_po_shots['tool_id'] == selected_tool_for_bar]
        
        # Process the raw data to generate 'stop_flag' and other metrics needed for the chart
        calc_for_chart = cr_CG_utils.CapacityRiskCalculator(df_bar_view, **config)
        df_processed_for_chart = calc_for_chart.results.get('processed_df', df_bar_view)
        
        agg_po = cr_CG_utils.generate_po_periodic_data(df_bar_view, composite_po_record, bar_freq, config, working_days_per_week, working_hours_per_day)
        
        if not agg_po.empty:
            # Fixed disconnect: passed processed data and track_mode to match utils requirement
            st.plotly_chart(cr_CG_utils.plot_po_periodic_chart(agg_po, df_processed_for_chart, bar_freq, track_mode), use_container_width=True)
        else:
            st.warning("No periodic data available.")

        st.markdown("---")
        
        # --- GRAPH 2: Burn Up Chart ---
        st.markdown(f"### 2. Global Target Burn-Up ({track_mode})")
        pred_data = cr_CG_utils.generate_po_prediction_data(df_po_shots, composite_po_record, config)
        if pred_data:
            # Fixed disconnect: passed subset_logistics to handle multiple due date annotations
            st.plotly_chart(cr_CG_utils.plot_po_burnup(pred_data, subset_logistics), use_container_width=True)
            
            # --- Forecast Analysis Insights ---
            current_cum = pred_data['current_cum']
            target_qty = pred_data['total_qty']
            avg_rate = pred_data['avg_daily_rate']
            opt_rate = pred_data['opt_daily_rate']
            due_date = pred_data['due_date']
            
            if current_cum >= target_qty:
                st.success(f"🎉 **Target Fulfilled!** Current output ({current_cum:,.0f}) has met or exceeded the aggregated quantity ({target_qty:,.0f}).")
            else:
                remaining = target_qty - current_cum
                days_avg = remaining / avg_rate if avg_rate > 0 else 9999
                days_opt = remaining / opt_rate if opt_rate > 0 else 9999
                
                # Protect against series max
                actual_dates = pred_data.get('actual_dates', [])
                last_act_date = actual_dates[-1] if len(actual_dates) > 0 else min_start.date()
                
                finish_avg = last_act_date + timedelta(days=int(days_avg))
                finish_opt = last_act_date + timedelta(days=int(days_opt))
                
                status_color = "#ff6961" if finish_avg > due_date else "#77dd77"
                status_text = "LATE - AT RISK" if finish_avg > due_date else "ON TRACK"
                
                analysis_html = f"""
                <div style="background-color: #262730; padding: 15px; border-radius: 5px; border: 1px solid #41424C; margin-bottom: 20px;">
                    <h4 style="margin-top:0;">Forecast Analysis</h4>
                    <ul>
                        <li><strong>Status:</strong> <span style="color:{status_color}; font-weight:bold;">{status_text}</span> to meet demand by {due_date.strftime('%Y-%m-%d')}.</li>
                        <li>To meet demand of <strong>{target_qty:,.0f}</strong>, you need <strong>{remaining:,.0f}</strong> more parts.</li>
                        <li>At your current rate ({avg_rate:,.0f}/day), you are projected to finish on <strong>{finish_avg.strftime('%Y-%m-%d')}</strong>.</li>
                        <li>At optimal rate ({opt_rate:,.0f}/day), you could finish by <strong>{finish_opt.strftime('%Y-%m-%d')}</strong>.</li>
                    </ul>
                </div>
                """
                st.markdown(analysis_html, unsafe_allow_html=True)
        else:
            st.warning("Not enough production data to generate burn-up.")
            
        # --- Breakdown per Tooling Table ---
        st.markdown("### Breakdown per Active Tooling")
        tool_summary = []
        for tool_id, tool_df in df_po_shots.groupby('tool_id'):
            calc = cr_CG_utils.CapacityRiskCalculator(tool_df, **config)
            res = calc.results
            if not res: continue
            
            sup = tool_df['supplier_id'].iloc[0] if 'supplier_id' in tool_df.columns else 'Unknown'
            plt_id = tool_df['plant_id'].iloc[0] if 'plant_id' in tool_df.columns else 'Unknown'
            
            tool_summary.append({
                'Tool ID': tool_id,
                'Supplier Name': sup,
                'Plant': plt_id,
                'Total Shots': res['total_shots'],
                'Actual Output': res['actual_output_parts'],
                'Optimal Output': res['optimal_output_parts'],
                'Downtime Loss': res['capacity_loss_downtime_parts'],
                'Slow Loss': res['capacity_loss_slow_parts'],
                'Efficiency (%)': res['efficiency_rate']
            })
        
        if tool_summary:
            st.dataframe(pd.DataFrame(tool_summary).style.format({
                'Total Shots': '{:,.0f}',
                'Actual Output': '{:,.0f}',
                'Optimal Output': '{:,.0f}',
                'Downtime Loss': '{:,.0f}',
                'Slow Loss': '{:,.0f}',
                'Efficiency (%)': '{:.1f}'
            }), use_container_width=True, hide_index=True)
            
    else:
        # Fallback to Generic Projection if no PO data available
        st.warning("Upload a Logistics Plan and ensure Production Data has PO_NUMBER to enable full tracking. Displaying generic forecast below.")
        
        with st.expander("ℹ️ How Prediction Works (Formulas & Logic)", expanded=False):
            st.markdown("""
            This model projects future capacity based on historical daily performance.
            ### 1. 🔵 Blue Line: Forecast (Average Rate)
            Projects future output using your **Average Daily Rate**.
            ### 2. 🟢 Green Line: Best Case (Peak Rate)
            Projects output using your **Peak Daily Rate** (90th percentile).
            ### 3. 🟠 Orange Line: Required Rate
            Shows the daily output required to hit your Demand Goal by the Target Date.
            """, unsafe_allow_html=True)

        agg_daily = cr_CG_utils.get_aggregated_data(df_scope, 'Daily', config)
        if agg_daily.empty:
            st.warning("Not enough daily data to generate a forecast.")
            return

        c_ctrl, c_chart = st.columns([1, 2])
        with c_ctrl:
            with st.container(border=True):
                st.markdown("#### Forecast Settings")
                data_min = pd.to_datetime(agg_daily['Period']).min().date()
                data_max = pd.to_datetime(agg_daily['Period']).max().date()
                hist_start_date = st.date_input("History From Date", data_min, min_value=data_min, max_value=data_max, key="fc_hist_start")
                tgt_date = st.date_input("Target Date", data_max + timedelta(days=30), min_value=data_max, key="fc_date")
                dem_goal = st.number_input("Demand Goal (Total Parts)", 0, step=1000, key="fc_goal")
                
        agg_filtered = agg_daily[pd.to_datetime(agg_daily['Period']).dt.date >= hist_start_date]
        if agg_filtered.empty:
            st.warning("No data available for the selected history range.")
            return

        with c_chart:
            pred = cr_CG_utils.generate_prediction_data(agg_filtered, data_max, tgt_date, dem_goal)
            fig = cr_CG_utils.plot_prediction_chart(pred, dem_goal)
            fig.update_layout(title="Future Capacity Projection")
            st.plotly_chart(fig, use_container_width=True, key="fc_chart")
            
            if dem_goal > 0 and pred:
                current_cum = pred['historic_cum'].iloc[-1]
                remaining = dem_goal - current_cum
                avg_rate = pred['rates']['avg']
                days_needed = remaining / avg_rate if avg_rate > 0 else 9999
                finish_date = data_max + timedelta(days=int(days_needed))
                is_late = finish_date > tgt_date
                status_color = "#ff6961" if is_late else "#77dd77"
                status_text = "LATE - AT RISK" if is_late else "ON TRACK"
                
                st.markdown(f"""
                <div style="background-color: #262730; padding: 15px; border-radius: 5px; border: 1px solid #41424C;">
                    <h4 style="margin-top:0;">Forecast Analysis</h4>
                    <ul>
                        <li><strong>Status:</strong> <span style="color:{status_color}; font-weight:bold;">{status_text}</span> to meet demand by {tgt_date}.</li>
                        <li>To meet demand of <strong>{dem_goal:,.0f}</strong>, you need <strong>{remaining:,.0f}</strong> more parts.</li>
                        <li>At your current rate ({avg_rate:,.0f}/day), you are projected to finish on <strong>{finish_date}</strong>.</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)


def render_dashboard(df_tool, tool_name, config, dashboard_mode="Optimal", filter_context=None):
    """
    Renders the Main Capacity Dashboard.
    """
    st.header(f"Capacity Dashboard ({dashboard_mode})")
    display_filter_context(filter_context, tool_name)
    
    benchmark_mode = "Optimal Output" if dashboard_mode == "Optimal" else "Target Output"
    key_suffix = f"_{dashboard_mode.lower()}"

    # --- Controls ---
    c1, c2 = st.columns([2, 1])
    with c1:
        analysis_level = st.radio(f"Select Analysis Level ({dashboard_mode})",
            options=["Daily (by Run)", "Weekly (by Run)", "Monthly (by Run)", "Custom Period"],
            horizontal=True, key=f"cr_analysis_level{key_suffix}")
    with c2:
        enable_filter = st.toggle("Filter Small Runs", value=False, key=f"cr_filter_runs{key_suffix}")
        min_shots_filter = 1
        if enable_filter: min_shots_filter = st.number_input("Min Shots per Run", 1, 1000, 10, key=f"cr_min_shots{key_suffix}")

    st.markdown("---")

    base_calc = cr_CG_utils.CapacityRiskCalculator(df_tool, **config)
    df_processed = base_calc.results.get('processed_df', pd.DataFrame())
    
    if df_processed.empty: st.error("No data."); return
    if enable_filter:
        run_counts = df_processed.groupby('run_id')['run_id'].transform('count')
        df_processed = df_processed[run_counts >= min_shots_filter]

    # --- Selection ---
    df_view = pd.DataFrame(); sub_header = ""
    if "Daily" in analysis_level:
        dates = sorted(df_processed['shot_time'].dt.date.unique())
        sel_date = st.selectbox("Select Date", dates, index=len(dates)-1, format_func=lambda x: x.strftime('%d %b %Y'), key=f"cr_date_select{key_suffix}")
        df_view = df_processed[df_processed['shot_time'].dt.date == sel_date]
        sub_header = f"Summary for {sel_date.strftime('%d %b %Y')}"
    elif "Weekly" in analysis_level:
        df_processed['week_lbl'] = df_processed['shot_time'].dt.to_period('W')
        weeks = sorted(df_processed['week_lbl'].unique())
        sel_week = st.selectbox("Select Week", weeks, index=len(weeks)-1, key=f"cr_week_select{key_suffix}")
        df_view = df_processed[df_processed['week_lbl'] == sel_week]
        sub_header = f"Summary for {sel_week}"
    elif "Monthly" in analysis_level:
        df_processed['month_lbl'] = df_processed['shot_time'].dt.to_period('M')
        months = sorted(df_processed['month_lbl'].unique())
        sel_month = st.selectbox("Select Month", months, index=len(months)-1, format_func=lambda x: x.strftime('%B %Y'), key=f"cr_month_select{key_suffix}")
        df_view = df_processed[df_processed['month_lbl'] == sel_month]
        sub_header = f"Summary for {sel_month.strftime('%B %Y')}"
    else:
        d_min = df_processed['shot_time'].min().date(); d_max = df_processed['shot_time'].max().date()
        c1, c2 = st.columns(2)
        s_date = c1.date_input("Start Date", d_min, key=f"d1{key_suffix}"); e_date = c2.date_input("End Date", d_max, key=f"d2{key_suffix}")
        if s_date and e_date:
            df_view = df_processed[(df_processed['shot_time'].dt.date >= s_date) & (df_processed['shot_time'].dt.date <= e_date)]
            sub_header = f"Summary for {s_date} to {e_date}"

    if df_view.empty: st.warning("No data found."); return

    # --- Calculations ---
    run_breakdown_df = cr_CG_utils.calculate_run_summaries(df_view, config)
    if run_breakdown_df.empty: st.warning("No runs found."); return

    total_runtime = run_breakdown_df['total_runtime_sec'].sum()
    prod_time = run_breakdown_df['production_time_sec'].sum()
    downtime = run_breakdown_df['downtime_sec'].sum()
    total_cap_loss_sec = run_breakdown_df['total_capacity_loss_sec'].sum()
    total_shots = run_breakdown_df['total_shots'].sum()
    normal_shots = run_breakdown_df['normal_shots'].sum()
    stop_events = run_breakdown_df['stop_events'].sum()
    
    opt_output = run_breakdown_df['optimal_output_parts'].sum()
    tgt_output = run_breakdown_df['target_output_parts'].sum() if 'target_output_parts' in run_breakdown_df.columns else (opt_output * (config['target_output_perc']/100.0))
    act_output = run_breakdown_df['actual_output_parts'].sum()
    
    loss_downtime = run_breakdown_df['capacity_loss_downtime_parts'].sum()
    loss_slow = run_breakdown_df['capacity_loss_slow_parts'].sum()
    gain_fast = run_breakdown_df['capacity_gain_fast_parts'].sum()
    total_loss_parts = run_breakdown_df['total_capacity_loss_parts'].sum()

    eff_rate = (normal_shots / total_shots * 100) if total_shots > 0 else 0
    stab_index = (prod_time / total_runtime * 100) if total_runtime > 0 else 0
    mttr_min = (downtime / 60 / stop_events) if stop_events > 0 else 0
    mtbf_min = (prod_time / 60 / stop_events) if stop_events > 0 else (prod_time / 60)

    if dashboard_mode == "Target":
        benchmark_output = tgt_output
        net_diff = act_output - tgt_output
    else:
        benchmark_output = opt_output
        net_diff = act_output - opt_output

    res = {
        'total_runtime_sec': total_runtime, 'production_time_sec': prod_time, 'downtime_sec': downtime,
        'total_capacity_loss_sec': total_cap_loss_sec, 'efficiency_rate': eff_rate, 'stability_index': stab_index,
        'mttr_min': mttr_min, 'mtbf_min': mtbf_min, 'optimal_output_parts': opt_output,
        'target_output_parts': tgt_output, 'actual_output_parts': act_output, 'total_shots': total_shots,
        'normal_shots': normal_shots, 'stop_events': stop_events, 'capacity_loss_downtime_parts': loss_downtime,
        'capacity_loss_slow_parts': loss_slow, 'capacity_gain_fast_parts': gain_fast,
        'total_capacity_loss_parts': total_loss_parts, 'processed_df': df_view 
    }

    # --- Header & Export ---
    c_head, c_btn = st.columns([3, 1])
    with c_head: st.subheader(sub_header)
    with c_btn:
        st.download_button(
            label="📥 Export Capacity Report",
            data=cr_CG_utils.prepare_and_generate_capacity_excel(df_view, config),
            file_name=f"Capacity_Report_{datetime.now():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_btn_{key_suffix}"
        )

    # ==========================================================================
    # --- VISUAL KPI DASHBOARD (New Section) ---
    # ==========================================================================
    
    st.plotly_chart(
        cr_CG_utils.create_time_breakdown_donut(
            res['total_runtime_sec'], 
            res['production_time_sec'], 
            res['downtime_sec']
        ), 
        use_container_width=True,
        key=f"time_donut_{key_suffix}"
    )
    
    c_g1, c_g2 = st.columns(2)
    with c_g1:
        with st.container(border=True):
             st.plotly_chart(
                cr_CG_utils.create_modern_gauge(res['efficiency_rate'], "Run Rate Efficiency"),
                use_container_width=True,
                key=f"eff_gauge_{key_suffix}"
             )
             st.markdown(f"""
             <div style='text-align: center; font-size: 14px; color: #a0a0a0; margin-bottom: 10px;'>
                {res['normal_shots']:,} / {res['total_shots']:,} (Normal Shots / Total Shots)
             </div>
             """, unsafe_allow_html=True)
             
             st.markdown(f"""
             <div style='text-align: center; font-size: 12px; margin-top: 5px;'>
                <span style='color:{cr_CG_utils.PASTEL_COLORS['red']};'>● 0-50%</span> &nbsp;
                <span style='color:{cr_CG_utils.PASTEL_COLORS['orange']};'>● 50-70%</span> &nbsp;
                <span style='color:{cr_CG_utils.PASTEL_COLORS['green']};'>● 70-100%</span>
             </div>
             """, unsafe_allow_html=True)

    with c_g2:
        with st.container(border=True):
             st.plotly_chart(
                cr_CG_utils.create_modern_gauge(res['stability_index'], "Run Rate Stability Index"),
                use_container_width=True,
                key=f"stab_gauge_{key_suffix}"
             )
             prod_str = cr_CG_utils.format_seconds_to_dhm(res['production_time_sec'])
             total_str = cr_CG_utils.format_seconds_to_dhm(res['total_runtime_sec'])
             st.markdown(f"""
             <div style='text-align: center; font-size: 14px; color: #a0a0a0; margin-bottom: 10px;'>
                {prod_str} / {total_str} (Production Time / Total Run Duration)
             </div>
             """, unsafe_allow_html=True)
             
             st.markdown(f"""
             <div style='text-align: center; font-size: 12px; margin-top: 5px;'>
                <span style='color:{cr_CG_utils.PASTEL_COLORS['red']};'>● 0-50%</span> &nbsp;
                <span style='color:{cr_CG_utils.PASTEL_COLORS['orange']};'>● 50-70%</span> &nbsp;
                <span style='color:{cr_CG_utils.PASTEL_COLORS['green']};'>● 70-100%</span>
             </div>
             """, unsafe_allow_html=True)
    
    with st.container(border=True):
        st.plotly_chart(
            cr_CG_utils.create_stability_driver_bar(res['mtbf_min'], res['mttr_min'], res['stability_index']),
            use_container_width=True,
            key=f"stab_driver_{key_suffix}"
        )
        
        with st.expander("🔍 View Correlation Analysis"):
            st.markdown(cr_CG_utils.generate_mttr_mtbf_analysis(run_breakdown_df), unsafe_allow_html=True)

    # ==========================================================================
    # --- NUMERIC METRICS (KPI Grid 3) ---
    # ==========================================================================
    with st.container(border=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Shots", f"{res['total_shots']:,.0f}")
        
        with c2:
            pct_normal = (res['normal_shots'] / res['total_shots'] * 100) if res['total_shots'] > 0 else 0
            st.metric("Normal Shots", f"{res['normal_shots']:,.0f}")
            st.markdown(create_capsule(pct_normal, "grey", "%"), unsafe_allow_html=True)
        
        c3.metric(f"Total {dashboard_mode} (Parts)", f"{benchmark_output:,.0f}")
        
        with c4:
            pct_achieve = (res['actual_output_parts'] / benchmark_output * 100) if benchmark_output > 0 else 0
            st.metric("Actual Output (Parts)", f"{res['actual_output_parts']:,.0f}")
            st.markdown(create_capsule(pct_achieve, "grey", "%"), unsafe_allow_html=True)
            
        with c5:
            label = "Net Gain (Parts)" if net_diff >= 0 else "Net Loss (Parts)"
            pct_diff = (net_diff / benchmark_output * 100) if benchmark_output > 0 else 0
            st.metric(label, f"{abs(net_diff):,.0f}")
            st.markdown(create_capsule(pct_diff, "net", "%"), unsafe_allow_html=True)

    with st.expander("ℹ️ Metric Definitions"):
        st.markdown(f"""
        ### 1. Run Rate Efficiency
        **Definition:** The ratio of high-quality, validated cycles to the total cycles attempted.
        - **Formula:** `Normal Shots / Total Shots`
        - **Why it matters:** Indicates process capability and adherence to standard cycle times.

        ### 2. Run Rate Stability Index
        **Definition:** The percentage of total run time where the machine was actively producing parts (not stopped).
        - **Formula:** `Total Production Time / Total Run Duration`
        - **Why it matters:** Measures machine availability during scheduled runs.

        ### 3. Run Rate MTTR (Mean Time To Repair)
        **Definition:** The average time taken to resolve a stop and resume production.
        - **Formula:** `Total Downtime / Number of Stop Events`
        - **Why it matters:** Highlights the speed of response and repair efficiency.

        ### 4. Run Rate MTBF (Mean Time Between Failures)
        **Definition:** The average duration of uninterrupted production between two stop events.
        - **Formula:** `Total Production Time / Number of Stop Events`
        - **Why it matters:** Indicates the reliability and consistency of the process.

        ### 5. Capacity Outputs
        - **Total {dashboard_mode}:** The theoretical maximum output based on ideal conditions ({ "Time / Ideal CT" if dashboard_mode == "Optimal" else "Optimal * Target %" }).
        - **Actual Output:** The total number of good parts produced.
        - **Net Loss/Gain:** The difference between Actual Output and the Benchmark ({dashboard_mode}).
        """)

    with st.container(border=True):
        c1, c2 = st.columns([1,3])
        c1.metric("Approved CT", f"{df_view['approved_ct'].mean():.2f} s")
        rc1, rc2, rc3 = c2.columns(3)
        rc1.metric("Lower Limit", f"{run_breakdown_df['mode_lower'].min():.2f}-{run_breakdown_df['mode_lower'].max():.2f} s")
        rc2.metric("Mode CT", f"{run_breakdown_df['mode_ct'].min():.2f}-{run_breakdown_df['mode_ct'].max():.2f} s")
        rc3.metric("Upper Limit", f"{run_breakdown_df['mode_upper'].min():.2f}-{run_breakdown_df['mode_upper'].max():.2f} s")

    st.markdown("---")

    with st.expander("🤖 View Automated Analysis Summary", expanded=True):
        insights = cr_CG_utils.generate_capacity_insights(res, dashboard_mode)
        st.markdown(f"""
        **Overall:** {insights['overall']}  
        **Drivers:** {insights['drivers']}  
        **Recommendation:** {insights['recommendation']}
        """, unsafe_allow_html=True)
    
    with st.expander("View Detailed Run Breakdown Table", expanded=False):
        d_df = run_breakdown_df.copy()
        
        d_df = d_df.sort_values('start_time').reset_index(drop=True)
        d_df['RUN ID'] = d_df.index.map(lambda x: f"Run {x+1:03d}")
        
        d_df["Period"] = d_df.apply(lambda row: f"{row['start_time'].strftime('%Y-%m-%d %H:%M')} to {row['end_time'].strftime('%Y-%m-%d %H:%M')}", axis=1)
        
        d_df['Run Rate MTTR (min)'] = (d_df['downtime_sec'] / 60) / d_df['stop_events'].replace(0, 1)
        d_df.loc[d_df['stop_events'] == 0, 'Run Rate MTTR (min)'] = 0
        
        d_df['Run Rate MTBF (min)'] = (d_df['production_time_sec'] / 60) / d_df['stop_events'].replace(0, 1)
        d_df.loc[d_df['stop_events'] == 0, 'Run Rate MTBF (min)'] = d_df['production_time_sec'] / 60

        d_df["Total Run Duration"] = d_df['total_runtime_sec'].apply(cr_CG_utils.format_seconds_to_dhm)
        d_df["Production Time"] = d_df['production_time_sec'].apply(cr_CG_utils.format_seconds_to_dhm)
        d_df["Run Rate Downtime"] = d_df['downtime_sec'].apply(cr_CG_utils.format_seconds_to_dhm)
        
        d_df = d_df.rename(columns={
            'tool_ids': 'Tool(s)',
            'total_shots': 'Total Shots',
            'normal_shots': 'Normal Shots',
            'stop_events': 'Stop Events',
            'mode_ct': 'Mode CT',
            'optimal_output_parts': 'Optimal Output',
            'actual_output_parts': 'Actual Output',
            'capacity_loss_downtime_parts': 'Loss (RR Downtime)',
            'capacity_loss_slow_parts': 'Loss (Slow Cycles)',
            'total_capacity_loss_parts': 'Total Net Loss'
        })

        cols_to_show = [
            'RUN ID', 'Tool(s)', 'Period', 'Total Shots', 'Normal Shots', 'Stop Events',
            'Mode CT', 'Optimal Output', 'Actual Output',
            'Loss (RR Downtime)', 'Loss (Slow Cycles)', 'Total Net Loss',
            'Total Run Duration', 'Production Time', 'Run Rate Downtime',
            'Run Rate MTTR (min)', 'Run Rate MTBF (min)'
        ]
        
        st.dataframe(d_df[cols_to_show].style.format({
            'Mode CT': '{:.2f}',
            'Optimal Output': '{:,.0f}',
            'Actual Output': '{:,.0f}',
            'Loss (RR Downtime)': '{:,.0f}',
            'Loss (Slow Cycles)': '{:,.0f}',
            'Total Net Loss': '{:,.0f}',
            'Run Rate MTTR (min)': '{:.1f}',
            'Run Rate MTBF (min)': '{:.1f}'
        }), use_container_width=True, hide_index=True)


    st.subheader(f"Capacity Loss Waterfall (vs {dashboard_mode})")
    
    waterfall_mode = "Standard (Net)"
    is_allocated = False
    if dashboard_mode == "Target":
        waterfall_mode = st.selectbox("Waterfall View Mode", ["Standard (Net)", "Allocated Impact"], key=f"wf_mode_{key_suffix}")
        if waterfall_mode == "Allocated Impact":
            is_allocated = True

    c_chart, c_details = st.columns([1.5, 1]) 
    gap_tgt = max(0, tgt_output - act_output)
    
    with c_chart:
        if is_allocated and dashboard_mode == "Target":
             net_loss_optimal = loss_downtime + (loss_slow - gain_fast)
             alloc_dt = 0; alloc_slow = 0; alloc_fast = 0
             if gap_tgt > 0 and net_loss_optimal > 0:
                 scale_factor = gap_tgt / net_loss_optimal
                 alloc_dt = loss_downtime * scale_factor
                 alloc_slow = loss_slow * scale_factor
                 alloc_fast = gain_fast * scale_factor
             
             y_dt = -alloc_dt
             y_slow = -alloc_slow
             y_fast = alloc_fast
             
             fig_wf = go.Figure(go.Waterfall(
                name="Allocated Impact", orientation="v",
                measure=["absolute", "relative", "relative", "relative", "total"],
                x=["Target Output", "Allocated: Downtime", "Allocated: Slow Cycles", "Allocated: Fast Cycles", "Actual Output"],
                y=[tgt_output, y_dt, y_slow, y_fast, act_output],
                text=[f"{tgt_output:,.0f}", f"{abs(y_dt):,.0f}", f"{abs(y_slow):,.0f}", f"+{abs(y_fast):,.0f}", f"{act_output:,.0f}"],
                textposition="outside",
                connector={"line": {"color": "rgb(63, 63, 63)"}},
                decreasing={"marker": {"color": cr_CG_utils.PASTEL_COLORS['red']}},
                increasing={"marker": {"color": cr_CG_utils.PASTEL_COLORS['green']}},
                totals={"marker": {"color": cr_CG_utils.PASTEL_COLORS['blue']}}
             ))
             fig_wf.update_layout(title="Allocated Capacity Loss (Target -> Actual)", showlegend=False, height=450)
             st.plotly_chart(fig_wf, use_container_width=True, key=f"waterfall_chart_{key_suffix}")
        else:
             st.plotly_chart(cr_CG_utils.plot_waterfall(res, benchmark_mode), use_container_width=True, key=f"waterfall_chart_{key_suffix}")
    
    with c_details:
        with st.container(border=True):
            if is_allocated:
                st.markdown(f"**Total Gap to Target**")
                color_hex = "#ff6961" if gap_tgt > 0 else "#77dd77" 
                st.markdown(f"<h2 style='color:{color_hex}; margin:0;'>-{gap_tgt:,.0f} parts</h2>", unsafe_allow_html=True)
                st.caption("Gap allocated by root cause ratios")
            else:
                st.markdown(f"**Total Net Impact (vs {dashboard_mode})**")
                color_hex = "#77dd77" if net_diff >= 0 else "#ff6961"
                st.markdown(f"<h2 style='color:{color_hex}; margin:0;'>{net_diff:+,.0f} parts</h2>", unsafe_allow_html=True)
                if dashboard_mode == "Optimal":
                    st.caption(f"Net Time Lost: {cr_CG_utils.format_seconds_to_dhm(res['total_capacity_loss_sec'])}")
        
        if is_allocated:
            net_loss_optimal = loss_downtime + (loss_slow - gain_fast)
            scale_factor = gap_tgt / net_loss_optimal if (gap_tgt > 0 and net_loss_optimal > 0) else 0
            
            a_dt = loss_downtime * scale_factor
            a_sl = loss_slow * scale_factor
            a_fg = gain_fast * scale_factor
            
            breakdown_data = [
                {"Metric": "Target Output", "Parts": tgt_output},
                {"Metric": "Actual Output", "Parts": act_output},
                {"Metric": "Total Gap", "Parts": gap_tgt},
                {"Metric": "--- Allocation ---", "Parts": 0},
                {"Metric": "Allocated Impact: Downtime", "Parts": a_dt},
                {"Metric": "Allocated Impact: Slow Cycles", "Parts": a_sl},
                {"Metric": "Allocated Impact: Fast Cycles (Gain)", "Parts": a_fg},
            ]
        else:
            net_cycle_loss = res['capacity_loss_slow_parts'] - res['capacity_gain_fast_parts']
            breakdown_data = [
                {"Metric": "Loss (RR Downtime)", "Parts": res['capacity_loss_downtime_parts']},
                {"Metric": "Net Loss (Cycle Time)", "Parts": net_cycle_loss},
                {"Metric": "└ Loss (Slow Cycles)", "Parts": res['capacity_loss_slow_parts']},
                {"Metric": "└ Gain (Fast Cycles)", "Parts": res['capacity_gain_fast_parts']},
            ]

        df_breakdown = pd.DataFrame(breakdown_data)
        
        def style_breakdown(row):
            styles = [''] * len(row)
            if "Allocated" in row['Metric']:
                 styles[0] = 'font-style: italic;'
                 if "Gain" in row['Metric'] or "Fast" in row['Metric']:
                     if row['Parts'] > 0: styles[1] = 'color: #77dd77;' 
                 elif row['Parts'] > 0: 
                     styles[1] = 'color: #ff6961;' 
            
            if row['Metric'] == "Total Gap":
                 styles[1] = 'color: #ff6961; font-weight: bold;'
            
            if row['Metric'] == "Loss (RR Downtime)":
                styles[1] = 'color: #ff6961; font-weight: bold;'
            elif row['Metric'] == "Net Loss (Cycle Time)":
                color = '#ff6961' if row['Parts'] > 0 else '#77dd77'
                styles[1] = f'color: {color}; font-weight: bold;'
            elif "Gain" in row['Metric']:
                styles[1] = 'color: #77dd77;'
            elif "Loss" in row['Metric'] and "Net" not in row['Metric']:
                styles[1] = 'color: #ff6961;'
            return styles

        st.dataframe(
            df_breakdown.style.apply(style_breakdown, axis=1).format({"Parts": "{:,.0f}"}), 
            use_container_width=True, 
            hide_index=True
        )

    st.markdown("---")

    st.subheader(f"Performance Breakdown (Stacked Trend)")
    st.info("View how capacity and losses were distributed over the selected period.")
    
    chart_freq = st.selectbox("Chart Aggregation", ["Daily", "Weekly", "Run"], key=f"chart_agg_{key_suffix}")
    freq_map = {"Daily": "Daily", "Weekly": "Weekly", "Run": "by Run"}
    
    agg_chart_df = cr_CG_utils.get_aggregated_data(df_view, freq_map[chart_freq], config)
    if not agg_chart_df.empty:
        st.plotly_chart(
            cr_CG_utils.plot_performance_breakdown(agg_chart_df, 'Period', benchmark_mode), 
            use_container_width=True,
            key=f"perf_breakdown_{key_suffix}"
        )
    else:
        st.warning("Not enough data to generate breakdown chart.")

    if not agg_chart_df.empty:
        st.subheader(f"Production Totals Report ({chart_freq})")
        
        totals_df = agg_chart_df.copy()
        if 'Production Time Sec' in totals_df and 'Run Time Sec' in totals_df:
            totals_df['Actual Production Time'] = totals_df.apply(
                lambda r: f"{cr_CG_utils.format_seconds_to_dhm(r['Production Time Sec'])} ({r['Production Time Sec']/r['Run Time Sec']:.1%})" if r['Run Time Sec'] > 0 else "0m (0.0%)", 
                axis=1
            )
        else:
            totals_df['Actual Production Time'] = "N/A"

        if 'Normal Shots' in totals_df and 'Total Shots' in totals_df:
            totals_df['Production Shots (Pct)'] = totals_df.apply(
                lambda r: f"{r['Normal Shots']:,.0f} ({r['Normal Shots']/r['Total Shots']:.1%})" if r['Total Shots'] > 0 else "0 (0.0%)", 
                axis=1
            )
        else:
            totals_df['Production Shots (Pct)'] = "N/A"
        
        totals_table = pd.DataFrame()
        totals_table['Period'] = totals_df['Period']
        totals_table['Total Run Duration'] = totals_df['Run Time'] + " (" + totals_df['Run Time Sec'].apply(lambda x: f"{x:.0f}s") + ")"
        totals_table['Actual Production Time'] = totals_df['Actual Production Time']
        totals_table['Total Shots'] = totals_df['Total Shots'].map('{:,.0f}'.format)
        totals_table['Production Shots'] = totals_df['Production Shots (Pct)']
        totals_table['Downtime Shots'] = totals_df['Downtime Shots'].map('{:,.0f}'.format)
        
        st.dataframe(totals_table, use_container_width=True, hide_index=True)

        st.subheader(f"Capacity Loss & Gain Report (vs Optimal) ({chart_freq})")
        
        lg_table_opt = pd.DataFrame()
        lg_table_opt['Period'] = totals_df['Period']
        lg_table_opt['Optimal Output'] = totals_df['Optimal Output'].map('{:,.2f}'.format)
        lg_table_opt['Actual Output'] = totals_df['Actual Output'].map('{:,.2f}'.format)
        
        lg_table_opt['Loss (Downtime)'] = totals_df['Downtime Loss'].map('{:,.2f}'.format)
        lg_table_opt['Loss (Slow Cycles)'] = totals_df['Slow Loss'].map('{:,.2f}'.format)
        lg_table_opt['Gain (Fast Cycles)'] = totals_df['Fast Gain'].map('{:,.2f}'.format)
        lg_table_opt['Total Net Loss'] = totals_df['Total Loss'].map('{:,.2f}'.format)

        def style_loss_gain(col):
            col_name = col.name
            if 'Loss' in col_name: 
                return ['color: #ff6961'] * len(col) 
            if 'Gain' in col_name: 
                return ['color: #77dd77'] * len(col) 
            if col_name == 'Total Net Loss':
                return ['font-weight: bold'] * len(col)
            return [''] * len(col)

        st.dataframe(lg_table_opt.style.apply(style_loss_gain, axis=0), use_container_width=True, hide_index=True)

        if dashboard_mode == "Target" and 'Target Output' in totals_df.columns:
            st.subheader(f"Capacity Loss & Gain Report (vs Target) [Allocated] ({chart_freq})")
            
            tgt_table = pd.DataFrame()
            tgt_table['Period'] = totals_df['Period']
            tgt_table['Target Output'] = totals_df['Target Output'].map('{:,.2f}'.format)
            tgt_table['Actual Output'] = totals_df['Actual Output'].map('{:,.2f}'.format)
            
            def calc_alloc(row):
                gap = max(0, row['Target Output'] - row['Actual Output'])
                net_loss_opt = row['Downtime Loss'] + (row['Slow Loss'] - row['Fast Gain'])
                scale = gap / net_loss_opt if (gap > 0 and net_loss_opt > 0) else 0
                
                dt_alloc = row['Downtime Loss'] * scale
                slow_alloc = row['Slow Loss'] * scale
                fast_alloc = row['Fast Gain'] * scale
                return pd.Series([gap, dt_alloc, slow_alloc, fast_alloc])

            alloc_res = totals_df.apply(calc_alloc, axis=1)
            alloc_res.columns = ['Gap', 'Alloc_DT', 'Alloc_Slow', 'Alloc_Fast']
            
            tgt_table['Gap to Target'] = alloc_res['Gap'].map('{:,.2f}'.format)
            tgt_table['Allocated: Downtime'] = alloc_res['Alloc_DT'].map('{:,.2f}'.format)
            tgt_table['Allocated: Slow Cycles'] = alloc_res['Alloc_Slow'].map('{:,.2f}'.format)
            tgt_table['Allocated: Fast Cycles (Gain)'] = alloc_res['Alloc_Fast'].map('{:,.2f}'.format)
            
            def style_target_alloc(col):
                if 'Gap' in col.name or 'Allocated' in col.name:
                    if 'Gain' in col.name or 'Fast' in col.name:
                        return ['color: #77dd77'] * len(col) 
                    return ['color: #ff6961'] * len(col) 
                return [''] * len(col)

            st.dataframe(tgt_table.style.apply(style_target_alloc, axis=0), use_container_width=True, hide_index=True)

    st.markdown("---")

    st.subheader("Shot Analysis")
    st.plotly_chart(cr_CG_utils.plot_shot_analysis(res['processed_df']), use_container_width=True, key=f"shot_analysis_{key_suffix}")
    
    with st.expander("View Shot Data Table", expanded=False):
            cols_to_show = ['tool_id', 'shot_time', 'actual_ct', 'adj_ct_sec', 'time_diff_sec', 'stop_flag', 'stop_event']
            rename_map = {
                'tool_id': 'Tool ID',
                'shot_time': 'Date / Time',
                'actual_ct': 'Actual CT (sec)',
                'adj_ct_sec': 'Adjusted CT (sec)',
                'time_diff_sec': 'Time Difference (sec)',
                'stop_flag': 'Stop Flag',
                'stop_event': 'Stop Event'
            }
            if 'run_id' in res['processed_df'].columns:
                cols_to_show.append('run_id')
                rename_map['run_id'] = 'Run ID'
                
            df_shot_data = res['processed_df'][cols_to_show].copy()
            df_shot_data.rename(columns=rename_map, inplace=True)
            st.dataframe(df_shot_data)

# ==============================================================================
# --- MAIN ENTRY POINT ---
# ==============================================================================

def main():
    st.sidebar.title("Capacity Risk v10.6")
    
    st.sidebar.markdown("### Data Upload")
    files = st.sidebar.file_uploader("1. Upload Production Data (Excel/CSV)", accept_multiple_files=True, type=['xlsx', 'csv', 'xls'])
    if not files: st.info("👈 Upload production data files."); st.stop()
    
    logistics_file = st.sidebar.file_uploader("2. Upload Logistics Plan (Excel/CSV) [Optional]", accept_multiple_files=False, type=['xlsx', 'csv', 'xls'])
    
    df_all = cr_CG_utils.load_all_data_cr(files)
    if df_all.empty: st.error("No valid production data."); st.stop()
    
    df_logistics = cr_CG_utils.load_logistics_plan(logistics_file) if logistics_file else pd.DataFrame()

    # Detect if data actually contains hierarchical columns
    has_hierarchy = False
    for col in ['project_id', 'component_id', 'part_id', 'supplier_id', 'plant_id']:
        if col in df_all.columns:
            uniques = [u for u in df_all[col].astype(str).unique() if str(u).lower() not in ["unknown", "nan", "none"]]
            if len(uniques) > 0:
                has_hierarchy = True
                break

    if has_hierarchy:
        st.sidebar.markdown("### Hierarchy Filters")
        
        def get_options(df, col):
            if col in df.columns:
                uniques = [x for x in df[col].astype(str).unique() if str(x).lower() not in ["nan", "unknown", "none"]]
                if uniques:
                    return ["All"] + sorted(uniques)
            return ["All"]

        # 1. Project Filter
        sel_proj = st.sidebar.selectbox("Project", get_options(df_all, 'project_id'))
        df_f1 = df_all if sel_proj == "All" else df_all[df_all['project_id'].astype(str) == sel_proj]
        
        # 2. Component Filter (respects Project)
        sel_comp = st.sidebar.selectbox("Component", get_options(df_f1, 'component_id'))
        df_f2 = df_f1 if sel_comp == "All" else df_f1[df_f1['component_id'].astype(str) == sel_comp]
        
        # 3. Part Filter (respects Component)
        sel_part = st.sidebar.selectbox("Part", get_options(df_f2, 'part_id'))
        df_f3 = df_f2 if sel_part == "All" else df_f2[df_f2['part_id'].astype(str) == sel_part]
        
        # 4. Supplier Filter (respects Part)
        sel_sup = st.sidebar.selectbox("Supplier", get_options(df_f3, 'supplier_id'))
        df_f4 = df_f3 if sel_sup == "All" else df_f3[df_f3['supplier_id'].astype(str) == sel_sup]
        
        # 5. Plant Filter (respects Supplier)
        sel_plt = st.sidebar.selectbox("Plant", get_options(df_f4, 'plant_id'))
        df_part = df_f4 if sel_plt == "All" else df_f4[df_f4['plant_id'].astype(str) == sel_plt]
        
        filter_context = {
            "Project": sel_proj,
            "Component": sel_comp,
            "Part": sel_part,
            "Supplier": sel_sup,
            "Plant": sel_plt
        }
    else:
        df_part = df_all
        filter_context = {}

    tool_ids = sorted([str(x) for x in df_part['tool_id'].unique() if str(x).lower() not in ["nan", "unknown", "none"]])
    
    if not tool_ids:
        st.sidebar.warning("No tools found for this selection.")
        st.stop()

    # --- ALIGNED TOOL SELECTION ---
    st.sidebar.markdown("### Tool Selection")
    
    # NEW: Allow selecting 'All Tools' to take advantage of the updated run isolation logic in the Utils file
    tool_options = ["All Tools Combined"] + tool_ids
    selected_tool_option = st.sidebar.selectbox("Select Tool(s) (Dashboards)", tool_options)
    
    if selected_tool_option == "All Tools Combined":
        df_tool = df_part
        tool_name_display = "Multiple Tools (Rolled-Up)"
    else:
        df_tool = df_part[df_part['tool_id'].astype(str) == selected_tool_option]
        tool_name_display = selected_tool_option

    with st.sidebar.expander("Configure Metrics"):
        tolerance = st.slider("Tolerance Band", 0.01, 0.50, 0.05, 0.01)
        downtime_gap_tolerance = st.slider("Downtime Gap (sec)", 0.0, 5.0, 2.0, 0.5)
        run_interval_hours = st.slider("Run Interval (hours)", 1, 24, 8, 1)
        
    with st.sidebar.expander("Logistics & Schedule Config"):
        working_days_per_week = st.slider("Working Days per Week", 1, 7, 5)
        working_hours_per_day = st.slider("Working Hours per Day", 1, 24, 24)
    
    with st.sidebar.expander("Capacity Settings"):
        target_output_perc = st.slider("Target Output %", 50, 100, 90)
        default_cavities = st.number_input("Default Cavities", 1)
        remove_maint = st.checkbox("Remove Maintenance", False)

    config = {'target_output_perc': target_output_perc, 'tolerance': tolerance, 
              'downtime_gap_tolerance': downtime_gap_tolerance, 'run_interval_hours': run_interval_hours, 
              'default_cavities': default_cavities, 'remove_maintenance': remove_maint}
    
    # --- TABS ---
    t_risk, t_opt, t_tgt, t_trend, t_fc = st.tabs(["Risk Tower", "Capacity (Optimal)", "Capacity (Target)", "Trends", "Forecast (PO Tracking)"])
    
    with t_risk: render_risk_tower(df_part, config, filter_context)
    with t_opt: render_dashboard(df_tool, tool_name_display, config, "Optimal", filter_context) if not df_tool.empty else st.warning("No data.")
    with t_tgt: render_dashboard(df_tool, tool_name_display, config, "Target", filter_context) if not df_tool.empty else st.warning("No data.")
    with t_trend: render_trends_tab(df_tool, tool_name_display, config, filter_context) if not df_tool.empty else st.warning("No data.")
    with t_fc: render_forecast_tab(df_part, config, df_logistics, working_days_per_week, working_hours_per_day, filter_context) if not df_part.empty else st.warning("No data.")

if __name__ == "__main__":
    main()