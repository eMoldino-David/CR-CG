import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta

# ==============================================================================
# --- CONSTANTS & SHARED FUNCTIONS ---
# ==============================================================================

PASTEL_COLORS = {
    'red': '#ff6961',
    'orange': '#ffb347',
    'green': '#77dd77',
    'blue': '#3498DB',
    'grey': '#808080',
    'target_line': 'deepskyblue',
    'optimal_line': 'darkblue',
    'purple': '#8A2BE2'
}

def format_seconds_to_dhm(total_seconds):
    """Converts total seconds into a 'Xd Yh Zm' or 'Xs' string."""
    if pd.isna(total_seconds) or total_seconds < 0: return "N/A"
    if total_seconds < 60: return f"{total_seconds:.1f}s"
    total_minutes = int(total_seconds / 60)
    days = total_minutes // (60 * 24)
    remaining_minutes = total_minutes % (60 * 24)
    hours = remaining_minutes // 60
    minutes = remaining_minutes % 60
    parts = []
    if days > 0: parts.append(f"{days}d")
    if hours > 0: parts.append(f"{hours}h")
    if minutes > 0 or not parts: parts.append(f"{minutes}m")
    return " ".join(parts) if parts else "0m"

def load_production_data(files):
    """Loads and standardizes production shot data with hierarchical column support."""
    df_list = []
    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            col_map = {col.strip().upper(): col for col in df.columns}
            
            def get_col(target_list):
                for t in target_list:
                    if t in col_map: return col_map[t]
                return None

            mapping = {
                "tool_id": ["TOOLING ID", "EQUIPMENT CODE", "TOOL_ID", "TOOLING_ID", "TOOL", "MACHINE"],
                "shot_time": ["SHOT TIME", "SHOT_TIME", "TIMESTAMP", "DATE", "TIME", "STAMP"],
                "actual_ct": ["ACTUAL CT", "ACTUAL_CT", "CYCLE TIME", "CYCLE_TIME", "ACTUAL CYCLE"],
                "approved_ct": ["APPROVED CT", "APPROVED_CT", "STD CT", "STD_CT", "IDEAL CT"],
                "working_cavities": ["WORKING CAVITIES", "WORKING_CAVITIES", "CAVITIES", "CAV"],
                "project": ["PROJECT", "PROJECT_NAME", "CUSTOMER"],
                "part_id": ["PART_ID", "PART_NUMBER", "PART NUMBER", "PART", "ITEM"],
                "component_id": ["COMPONENT_ID", "COMPONENT", "SUB-ASSEMBLY"],
                "po_number": ["PO_NUMBER", "PO #", "PURCHASE ORDER", "PO_NUM"]
            }

            for key, targets in mapping.items():
                col = get_col(targets)
                if col: df.rename(columns={col: key}, inplace=True)

            if "shot_time" in df.columns and "actual_ct" in df.columns:
                df["shot_time"] = pd.to_datetime(df["shot_time"], errors="coerce")
                df["actual_ct"] = pd.to_numeric(df["actual_ct"], errors="coerce")
                df.dropna(subset=["shot_time", "actual_ct"], inplace=True)
                df_list.append(df)
        except Exception: continue
    
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

def load_po_data(files):
    """Loads Purchase Order planning data."""
    df_list = []
    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            df.columns = [c.strip().upper() for c in df.columns]
            mapping = {
                "PO_NUMBER": "po_number", "PROJECT": "project", "PART_ID": "part_id",
                "TOTAL_QTY": "total_qty", "START_DATE": "start_date", "DUE_DATE": "due_date"
            }
            df.rename(columns=mapping, inplace=True)
            df['start_date'] = pd.to_datetime(df['start_date'], errors='coerce')
            df['due_date'] = pd.to_datetime(df['due_date'], errors='coerce')
            df_list.append(df)
        except Exception: continue
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

# ==============================================================================
# --- CORE CALCULATION ENGINE ---
# ==============================================================================

class CapacityRiskCalculator:
    def __init__(self, df: pd.DataFrame, tolerance: float, downtime_gap_tolerance: float, 
                 run_interval_hours: float, default_cavities: int = 1, **kwargs):
        self.df_raw = df.copy()
        self.tolerance = tolerance
        self.downtime_gap_tolerance = downtime_gap_tolerance
        self.run_interval_hours = run_interval_hours
        self.default_cavities = default_cavities
        self.target_output_perc = kwargs.get('target_output_perc', 90)
        self.results = self._calculate_metrics()

    def _calculate_metrics(self) -> dict:
        df = self.df_raw.copy()
        if df.empty: return {}

        # Default columns if missing
        if 'approved_ct' not in df.columns: df['approved_ct'] = df['actual_ct'].median() 
        if 'working_cavities' not in df.columns: df['working_cavities'] = self.default_cavities
        
        # Multi-tool sorting
        sort_cols = ['tool_id', 'shot_time'] if 'tool_id' in df.columns else ['shot_time']
        df = df.sort_values(sort_cols).reset_index(drop=True)
        
        # 1. Run Identification (Independent per tool)
        if 'tool_id' in df.columns:
            df['time_diff_sec'] = df.groupby('tool_id')['shot_time'].diff().dt.total_seconds().fillna(0)
        else:
            df['time_diff_sec'] = df['shot_time'].diff().dt.total_seconds().fillna(0)
            
        is_new_run = df['time_diff_sec'] > (self.run_interval_hours * 3600)
        df['run_id'] = is_new_run.cumsum()

        # 2. Physics: Mode CT Detection per Run
        run_modes = df.groupby('run_id')['actual_ct'].transform(lambda x: x.mode().iloc[0] if not x.mode().empty else x.mean())
        df['mode_ct'] = run_modes
        df['mode_lower'] = df['mode_ct'] * (1 - self.tolerance)
        df['mode_upper'] = df['mode_ct'] * (1 + self.tolerance)
        
        # 3. Stop Logic
        df['next_shot_diff'] = df['time_diff_sec'].shift(-1).fillna(0)
        is_gap = df['next_shot_diff'] > (df['actual_ct'] + self.downtime_gap_tolerance)
        is_abnormal = (df['actual_ct'] < df['mode_lower']) | (df['actual_ct'] > df['mode_upper'])
        
        df['stop_flag'] = np.where(is_gap | is_abnormal, 1, 0)
        df.loc[is_new_run, 'stop_flag'] = 0 
        
        # 4. Volumetric Calculations
        prod_df = df[df['stop_flag'] == 0].copy()
        actual_output = prod_df['working_cavities'].sum()

        total_runtime_sec = 0
        for _, run in df.groupby('run_id'):
            if not run.empty:
                total_runtime_sec += (run['shot_time'].max() - run['shot_time'].min()).total_seconds() + run.iloc[-1]['actual_ct']

        downtime_sec = max(0, total_runtime_sec - prod_df['actual_ct'].sum())
        
        # Basis Calculations
        avg_app_ct = df['approved_ct'].mean()
        avg_cav = df['working_cavities'].mean()
        
        optimal_output = (total_runtime_sec / avg_app_ct) * avg_cav if avg_app_ct > 0 else 0
        loss_downtime = (downtime_sec / avg_app_ct) * avg_cav if avg_app_ct > 0 else 0
        
        # 5. Stability (MTTR/MTBF)
        stops = df[df['stop_flag'] == 1]
        num_stops = len(stops)
        mtbf_min = (total_runtime_sec / 60 / num_stops) if num_stops > 0 else (total_runtime_sec / 60)
        mttr_min = (downtime_sec / 60 / num_stops) if num_stops > 0 else 0
        stability_index = (1 - (downtime_sec / total_runtime_sec)) * 100 if total_runtime_sec > 0 else 0

        return {
            "processed_df": df,
            "total_runtime_sec": total_runtime_sec,
            "downtime_sec": downtime_sec,
            "actual_output": actual_output,
            "optimal_output": optimal_output,
            "loss_downtime": loss_downtime,
            "loss_speed": max(0, (optimal_output - loss_downtime) - actual_output),
            "gain_speed": max(0, actual_output - (optimal_output - loss_downtime)),
            "total_shots": len(df),
            "normal_shots": len(prod_df),
            "mtbf_min": mtbf_min,
            "mttr_min": mttr_min,
            "stability_index": stability_index
        }

# ==============================================================================
# --- RENDERING & VISUALIZATION ---
# ==============================================================================

def render_dashboard(df, selected_name, config, mode="Optimal"):
    import streamlit as st
    calc = CapacityRiskCalculator(df, **config)
    res = calc.results
    if not res: return

    st.subheader(f"Physics Analysis: {selected_name} ({mode} Basis)")

    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("Stability Index", f"{res['stability_index']:.1f}%")
    with k2: st.metric("Actual Output", f"{res['actual_output']:,.0f}")
    with k3: st.metric("MTBF", f"{res['mtbf_min']:.1f}m")
    with k4: st.metric("MTTR", f"{res['mttr_min']:.1f}m")

    st.markdown("---")
    g1, g2, g3 = st.columns(3)
    with g1:
        eff = (res['normal_shots'] / res['total_shots'] * 100) if res['total_shots'] > 0 else 0
        st.plotly_chart(create_gauge(eff, "Process Efficiency", PASTEL_COLORS['blue']), use_container_width=True)
    with g2:
        st.plotly_chart(create_gauge(res['stability_index'], "Availability", PASTEL_COLORS['green']), use_container_width=True)
    with g3:
        st.plotly_chart(create_time_donut(res, selected_name), use_container_width=True)

    st.markdown("---")
    bridge_res = res.copy()
    if mode == "Target":
        target_factor = config.get('target_output_perc', 90) / 100
        bridge_res['optimal_output'] *= target_factor
        bridge_res['loss_downtime'] *= target_factor
        bridge_res['loss_speed'] = max(0, (bridge_res['optimal_output'] - bridge_res['loss_downtime']) - bridge_res['actual_output'])
        bridge_res['gain_speed'] = max(0, bridge_res['actual_output'] - (bridge_res['optimal_output'] - bridge_res['loss_downtime']))

    st.plotly_chart(create_waterfall_bridge(bridge_res, selected_name, mode), use_container_width=True)

    with st.expander("Shot-by-Shot Cycle Analysis"):
        st.plotly_chart(plot_shot_analysis(res['processed_df'], selected_name, config), use_container_width=True)

def render_risk_tower(df_all, config):
    import streamlit as st
    st.subheader("Global Asset Physics Risk Tower")
    tower_data = []
    for tool, subset in df_all.groupby('tool_id'):
        calc = CapacityRiskCalculator(subset, **config)
        res = calc.results
        if not res: continue
        tower_data.append({
            'Tooling ID': tool,
            'Stability %': res['stability_index'],
            'Efficiency %': (res['normal_shots'] / res['total_shots'] * 100) if res['total_shots'] > 0 else 0,
            'Actual Output': res['actual_output'],
            'MTBF (min)': res['mtbf_min'],
            'MTTR (min)': res['mttr_min']
        })
    tower_df = pd.DataFrame(tower_data)
    if not tower_df.empty:
        st.dataframe(tower_df.style.background_gradient(subset=['Stability %'], cmap='RdYlGn'), use_container_width=True, hide_index=True)

def create_gauge(value, title, color):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number", value = value,
        title = {'text': title, 'font': {'size': 18}},
        gauge = {'axis': {'range': [0, 100]}, 'bar': {'color': color}, 'bgcolor': "white", 'borderwidth': 2, 'bordercolor': "gray"}
    ))
    fig.update_layout(height=280, margin=dict(l=30, r=30, t=50, b=30), paper_bgcolor='rgba(0,0,0,0)')
    return fig

def create_time_donut(res, selected_tool):
    labels = ['Productive Time', 'Downtime / Gaps']
    values = [res['total_runtime_sec'] - res['downtime_sec'], res['downtime_sec']]
    fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.6, marker=dict(colors=[PASTEL_COLORS['blue'], PASTEL_COLORS['red']]))])
    fig.update_layout(title_text=f"Time Allocation", height=300, showlegend=True, legend=dict(orientation="h", y=-0.1))
    return fig

def create_waterfall_bridge(res, selected_tool, mode="Optimal"):
    fig = go.Figure(go.Waterfall(
        name = "Capacity", orientation = "v",
        measure = ["absolute", "relative", "relative", "relative", "total"],
        x = [f"{mode} Cap", "Downtime Loss", "Speed Loss", "Speed Gain", "Actual Output"],
        y = [res['optimal_output'], -res['loss_downtime'], -res['loss_speed'], res['gain_speed'], 0],
        connector = {"line":{"color":"rgb(63, 63, 63)"}},
        decreasing = {"marker":{"color":PASTEL_COLORS['red']}},
        increasing = {"marker":{"color":PASTEL_COLORS['green']}},
        totals = {"marker":{"color":PASTEL_COLORS['blue']}}
    ))
    fig.update_layout(title = f"Capacity Loss Bridge", showlegend = False, height=450, template="plotly_white")
    return fig

def plot_shot_analysis(df_shots, selected_tool, config):
    fig = go.Figure()
    for r_id, run_df in df_shots.groupby('run_id'):
        lower, upper = run_df['mode_lower'].iloc[0], run_df['mode_upper'].iloc[0]
        start, end = run_df['shot_time'].min(), run_df['shot_time'].max()
        fig.add_shape(type="rect", x0=start, x1=end, y0=lower, y1=upper, fillcolor="grey", opacity=0.1, line_width=0)
    fig.add_trace(go.Scatter(x=df_shots['shot_time'], y=df_shots['actual_ct'], mode='markers', marker=dict(color=PASTEL_COLORS['blue'], size=4, opacity=0.6)))
    avg_ref = df_shots['approved_ct'].mean()
    fig.add_hline(y=avg_ref, line_dash="dash", line_color="green", annotation_text=f"Approved CT ({avg_ref:.1f}s)")
    fig.update_layout(title=f"Shot-by-Shot Cycle: {selected_tool}", xaxis_title="Time", yaxis_title="Cycle Time (s)", height=450, template="plotly_white")
    return fig

def create_stability_driver_bar(mtbf, mttr):
    """Visualizes the balance between failure frequency and repair duration."""
    total = mtbf + mttr
    if total == 0: return go.Figure()
    fig = go.Figure()
    fig.add_trace(go.Bar(y=['Asset Driver'], x=[mtbf], name='MTBF (Frequency)', orientation='h', marker_color=PASTEL_COLORS['blue']))
    fig.add_trace(go.Bar(y=['Asset Driver'], x=[mttr], name='MTTR (Duration)', orientation='h', marker_color=PASTEL_COLORS['red']))
    fig.update_layout(barmode='stack', height=180, template="plotly_white", showlegend=True, margin=dict(t=20, b=20, l=10, r=10))
    return fig