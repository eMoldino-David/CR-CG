import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta

# --- constants & styling ---
colors = {
    'red': '#ff6961',
    'orange': '#ffb347',
    'green': '#77dd77',
    'blue': '#3498db',
    'grey': '#808080',
    'purple': '#8a2be2',
    'stress_line': '#e91e63',
    'optimal_line': '#1f77b4'
}

def format_seconds_to_dhm(total_seconds):
    """Converts total seconds into a readable day, hour, minute format."""
    if pd.isna(total_seconds) or total_seconds < 0: return "n/a"
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

# --- robust mapping engine ---
def find_and_rename_cols(df, mapping):
    """improved fuzzy mapping to handle variations in user headers and prevent crashes."""
    col_map = {str(col).strip().upper(): col for col in df.columns}
    rename_dict = {}
    
    for internal_key, search_terms in mapping.items():
        found = False
        # priority 1: exact matches from list
        for term in search_terms:
            if term.upper() in col_map:
                rename_dict[col_map[term.upper()]] = internal_key
                found = True
                break
        
        # priority 2: partial matches if not found
        if not found:
            for actual_col_upper, actual_col_raw in col_map.items():
                if any(term.upper() in actual_col_upper for term in search_terms):
                    rename_dict[actual_col_raw] = internal_key
                    break
                    
    return df.rename(columns=rename_dict)

# --- data loading ---
def load_production_data(files):
    """standardizes uploaded production shot data with fuzzy matching and type safety."""
    df_list = []
    prod_mapping = {
        "po_number": ["PO_NUMBER", "PO #", "PURCHASE ORDER", "PO_NUM", "ORDER"],
        "project": ["PROJECT", "PROJ", "PROGRAM", "CUSTOMER"],
        "part_id": ["PART_ID", "PART NUMBER", "PART", "ITEM", "SKU"],
        "tool_id": ["TOOLING ID", "EQUIPMENT", "TOOL_ID", "ASSET", "TOOL", "MACHINE"],
        "shot_time": ["SHOT TIME", "TIMESTAMP", "DATE", "TIME", "STAMP", "CLOCK"],
        "actual_ct": ["ACTUAL CT", "ACTUAL_CT", "CYCLE TIME", "ACTUAL_CYCLE", "CT"],
        "approved_ct": ["APPROVED CT", "APPROVED_CT", "STD CT", "IDEAL CT", "TARGET CT"],
        "working_cavities": ["WORKING CAVITIES", "CAVITIES", "CAV", "CAVS", "WORKING_CAV"]
    }

    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            df = find_and_rename_cols(df, prod_mapping)

            if "shot_time" in df.columns:
                df["shot_time"] = pd.to_datetime(df["shot_time"], errors="coerce")
                df = df.dropna(subset=["shot_time"])
                
                # handle numeric safety
                numeric_cols = ["actual_ct", "approved_ct", "working_cavities"]
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                
                # defaults for missing but critical columns
                if "actual_ct" not in df.columns:
                    df = df.sort_values("shot_time")
                    df["actual_ct"] = df["shot_time"].diff().dt.total_seconds().fillna(30).clip(lower=1)
                
                if "working_cavities" not in df.columns or df["working_cavities"].sum() == 0:
                    df["working_cavities"] = 1
                
                df_list.append(df)
        except Exception: continue
    
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

def load_po_data(files):
    """standardizes uploaded planning/po data with fuzzy matching."""
    df_list = []
    po_mapping = {
        "po_number": ["PO_NUMBER", "PO #", "PURCHASE ORDER", "ORDER_NUM"], 
        "project": ["PROJECT", "PROJ", "PROGRAM"], 
        "part_id": ["PART_ID", "PART", "ITEM", "SKU"], 
        "total_qty": ["TOTAL_QTY", "QTY", "TARGET", "ORDER_QTY", "DEMAND"], 
        "due_date": ["DUE_DATE", "DUE", "DEADLINE", "SHIP_DATE"]
    }

    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            df = find_and_rename_cols(df, po_mapping)
            if 'due_date' in df.columns:
                df['due_date'] = pd.to_datetime(df['due_date'], errors='coerce')
            if 'total_qty' in df.columns:
                df['total_qty'] = pd.to_numeric(df['total_qty'], errors='coerce').fillna(0)
            df_list.append(df)
        except Exception: continue
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

# --- calculation engine ---
class CapacityRiskCalculator:
    """Restored original complexity for cycle, downtime, and stability analysis."""
    def __init__(self, df, config):
        self.df = df.copy()
        self.config = config
        self.results = self._calculate()

    def _calculate(self):
        if self.df.empty: return {}
        df = self.df.sort_values("shot_time").reset_index(drop=True)
        
        # 1. Run Identification
        df['time_diff'] = df['shot_time'].diff().dt.total_seconds().fillna(0)
        is_new_run = df['time_diff'] > (self.config.get('run_interval_hours', 8) * 3600)
        df['run_id'] = is_new_run.cumsum()
        
        # 2. Advanced Mode Detection
        run_modes = df.groupby('run_id')['actual_ct'].apply(lambda x: x.mode().iloc[0] if not x.mode().empty else x.median())
        df['mode_ct'] = df['run_id'].map(run_modes)
        df['mode_lower'] = df['mode_ct'] * (1 - self.config.get('tolerance', 0.05))
        df['mode_upper'] = df['mode_ct'] * (1 + self.config.get('tolerance', 0.05))
        
        # 3. Stop/Stability Identification
        df['next_diff'] = df['time_diff'].shift(-1).fillna(0)
        is_gap = df['next_diff'] > (df['actual_ct'] + 2.0)
        is_abnormal = (df['actual_ct'] < df['mode_lower']) | (df['actual_ct'] > df['mode_upper'])
        df['stop_flag'] = np.where(is_gap | is_abnormal, 1, 0)
        df.loc[is_new_run, 'stop_flag'] = 0
        
        # 4. State Classification
        if 'approved_ct' not in df.columns or (df['approved_ct'] == 0).all(): 
            df['approved_ct'] = df['mode_ct']
        
        conditions = [
            df['stop_flag'] == 1,
            df['actual_ct'] > (df['approved_ct'] * 1.10),
            df['actual_ct'] < (df['approved_ct'] * 0.90)
        ]
        choices = ['Downtime (Stop)', 'Slow Cycle', 'Fast Cycle']
        df['shot_type'] = np.select(conditions, choices, default='On Target')

        # 5. Volumetric & Time Calculations
        prod_df = df[df['stop_flag'] == 0]
        actual_output = prod_df['working_cavities'].sum()
        
        total_runtime_sec = 0
        for _, run in df.groupby('run_id'):
            if not run.empty:
                total_runtime_sec += (run['shot_time'].max() - run['shot_time'].min()).total_seconds() + run.iloc[-1]['actual_ct']
        
        downtime_sec = max(0, total_runtime_sec - prod_df['actual_ct'].sum())
        
        # Stability Metrics
        stops_count = df['stop_flag'].sum()
        mtbf_min = (total_runtime_sec / 60 / stops_count) if stops_count > 0 else (total_runtime_sec / 60)
        mttr_min = (downtime_sec / 60 / stops_count) if stops_count > 0 else 0
        stability_index = (1 - (downtime_sec / total_runtime_sec)) * 100 if total_runtime_sec > 0 else 0

        # Loss Bridge Logic
        avg_app_ct = df['approved_ct'].replace(0, np.nan).mean() or 30
        avg_cav = df['working_cavities'].mean() or 1
        
        optimal_output = (total_runtime_sec / avg_app_ct) * avg_cav
        loss_dt = (downtime_sec / avg_app_ct) * avg_cav
        # Speed Loss is the difference between what we should have made while running vs what we actually made
        loss_slow = max(0, (optimal_output - loss_dt) - actual_output)
        gain_fast = max(0, actual_output - (optimal_output - loss_dt))

        return {
            "processed_df": df,
            "total_runtime_sec": total_runtime_sec,
            "downtime_sec": downtime_sec,
            "actual_output": actual_output,
            "optimal_output": optimal_output,
            "loss_dt": loss_dt,
            "loss_slow": loss_slow,
            "gain_fast": gain_fast,
            "mtbf_min": mtbf_min,
            "mttr_min": mttr_min,
            "stability_index": stability_index,
            "total_shots": len(df),
            "normal_shots": len(prod_df)
        }

def get_supply_metrics(df, config, po_target=None):
    """improved aggregation that pro-rates PO targets across history."""
    if df.empty: return pd.DataFrame()
    df = df.copy()
    df['period'] = df['shot_time'].dt.to_period('W').apply(lambda r: r.start_time)
    
    # Calculate pro-rated target if a single PO target exists
    periods = df['period'].unique()
    num_periods = len(periods)
    target_per_period = (po_target / num_periods) if po_target and num_periods > 0 else 0
    
    data = []
    for period, subset in df.groupby('period'):
        calc = CapacityRiskCalculator(subset, config)
        res = calc.results
        data.append({
            'period': period,
            'actual_output': res.get('actual_output', 0),
            'runtime_sec': res.get('total_runtime_sec', 0),
            'target': target_per_period,
            'stability': res.get('stability_index', 0)
        })
    return pd.DataFrame(data)

def generate_capacity_insights(res):
    """Restored natural language logic."""
    if not res: return "no data available."
    gap = res.get('optimal_output', 0) - res.get('actual_output', 0)
    if gap <= 0: return "Tooling is operating efficiently at or above theoretical capacity."
    
    dt_p = (res.get('loss_dt', 0) / gap * 100) if gap > 0 else 0
    sl_p = (res.get('loss_slow', 0) / gap * 100) if gap > 0 else 0
    
    driver = "Unplanned Downtime (Stability)" if dt_p > sl_p else "Slow Cycles (Performance)"
    severity = "Critical" if gap > (res.get('actual_output', 0) * 0.2) else "Moderate"
    
    return f"Status: <b>{severity} Gap</b>. The primary bottleneck is <b>{driver}</b>, which accounts for {max(dt_p, sl_p):.1f}% of missing parts."

# --- visualization ---
def create_modern_gauge(value, title, unit="%"):
    color = colors['green']
    if value <= 60: color = colors['red']
    elif value <= 85: color = colors['orange']
    
    fig = go.Figure(data=[go.Pie(
        values=[value, 100-value if unit=="%" else max(0, 100-value), 100 if unit=="%" else 0], 
        hole=0.75, sort=False, direction='clockwise', rotation=-90,
        marker=dict(colors=[color, '#262730', 'rgba(0,0,0,0)']), hoverinfo='none', textinfo='none'
    )])
    fig.add_annotation(text=f"{value:.1f}{unit}", x=0.5, y=0.15, font=dict(size=38, color='white', weight='bold'), showarrow=False)
    fig.update_layout(title=dict(text=title, x=0.5, font=dict(size=16)), margin=dict(l=10, r=10, t=40, b=0), height=200, showlegend=False, paper_bgcolor='rgba(0,0,0,0)')
    return fig

def create_time_donut(total_sec, prod_sec, down_sec):
    fig = go.Figure(data=[go.Pie(
        labels=['Production', 'Downtime'], values=[prod_sec, down_sec], hole=0.7,
        marker=dict(colors=[colors['green'], colors['red']]), textinfo='none'
    )])
    center_text = f"Total Time<br><b>{format_seconds_to_dhm(total_sec)}</b>"
    fig.add_annotation(text=center_text, x=0.5, y=0.5, font=dict(size=14, color='white'), showarrow=False)
    fig.update_layout(title="Asset Availability", showlegend=True, legend=dict(orientation="h", y=-0.1), height=280, margin=dict(t=40, b=40, l=10, r=10), paper_bgcolor='rgba(0,0,0,0)')
    return fig

def create_hybrid_chart(df_agg, cal_config):
    """Volumetric bars + Schedule Accomplishment % + Stress Factor line."""
    fig = go.Figure()
    
    # 1. Volumetric Bars
    fig.add_trace(go.Bar(
        x=df_agg['period'], y=df_agg['actual_output'], 
        name='Actual Output', marker_color=colors['blue'], yaxis='y1'
    ))
    
    # 2. Accomplishment Line (Red)
    if 'target' in df_agg.columns and df_agg['target'].sum() > 0:
        df_agg['acc_pct'] = (df_agg['actual_output'] / df_agg['target'] * 100).clip(upper=150)
        fig.add_trace(go.Scatter(
            x=df_agg['period'], y=df_agg['acc_pct'], 
            name='Accomplishment %', line=dict(color=colors['red'], width=3), yaxis='y2'
        ))
    
    # 3. Stress Line (Dashed)
    if cal_config:
        planned_hrs_per_week = cal_config['days'] * cal_config['hours']
        df_agg['stress'] = (df_agg['runtime_sec'] / 3600 / planned_hrs_per_week * 100).clip(upper=200)
        fig.add_trace(go.Scatter(
            x=df_agg['period'], y=df_agg['stress'], 
            name='Tool Stress %', line=dict(color=colors['stress_line'], width=2, dash='dot'), yaxis='y2'
        ))

    fig.update_layout(
        template="plotly_dark", hovermode="x unified",
        yaxis=dict(title="Volume (Parts)"),
        yaxis2=dict(title="Percentage (%)", overlaying='y', side='right', range=[0, 150], showgrid=False),
        legend=dict(orientation="h", y=-0.2), height=480, paper_bgcolor='rgba(0,0,0,0)'
    )
    return fig

def plot_waterfall(res):
    """Standardized Waterfall for Capacity Bridge."""
    fig = go.Figure(go.Waterfall(
        name="bridge", orientation="v", measure=["absolute", "relative", "relative", "relative", "total"],
        x=["Optimal", "Downtime", "Slow Cycle", "Fast Cycle", "Actual"],
        y=[res.get('optimal_output', 0), -res.get('loss_dt', 0), -res.get('loss_slow', 0), res.get('gain_fast', 0), res.get('actual_output', 0)],
        decreasing={"marker": {"color": colors['red']}}, 
        increasing={"marker": {"color": colors['green']}},
        totals={"marker": {"color": colors['blue']}}
    ))
    fig.update_layout(title="Supply Gap Waterfall", template="plotly_dark", height=400)
    return fig

def plot_shot_analysis(df_shots):
    """Shot-by-shot distribution chart."""
    fig = go.Figure()
    cmap = {'Slow Cycle': colors['red'], 'Fast Cycle': colors['orange'], 'On Target': colors['blue'], 'Downtime (Stop)': '#808080'}
    for stype, color in cmap.items():
        sub = df_shots[df_shots['shot_type'] == stype]
        if not sub.empty:
            fig.add_trace(go.Bar(x=sub['shot_time'], y=sub['actual_ct'], name=stype, marker_color=color))
    fig.update_layout(template="plotly_dark", title="Production Rhythm (Shot-by-Shot)", yaxis_title="Seconds", height=380)
    return fig

def create_stability_driver_bar(mtbf, mttr):
    """Horizontal stacked bar to see if frequency or duration is the problem."""
    total = mtbf + mttr
    if total == 0: return go.Figure()
    
    fig = go.Figure()
    fig.add_trace(go.Bar(y=['Driver'], x=[mtbf], name='MTBF (Frequency)', orientation='h', marker_color=colors['blue']))
    fig.add_trace(go.Bar(y=['Driver'], x=[mttr], name='MTTR (Duration)', orientation='h', marker_color=colors['red']))
    fig.update_layout(barmode='stack', template="plotly_dark", height=150, showlegend=True, margin=dict(t=30, b=10, l=10, r=10))
    return fig

def create_po_burnup(df_agg, target_qty, target_date):
    if df_agg.empty: return go.Figure()
    df = df_agg.sort_values('period').copy()
    df['cumulative'] = df['actual_output'].cumsum()
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['period'], y=df['cumulative'], name='Actual Output', fill='tozeroy', line_color=colors['blue'], mode='lines+markers'))
    
    if target_qty > 0:
        # Step staircase for PO Target
        fig.add_trace(go.Scatter(
            x=[df['period'].min(), target_date], 
            y=[0, target_qty], 
            mode='lines', name='Demand Path', 
            line=dict(color=colors['purple'], dash='dash', shape='hv')
        ))
    
    fig.update_layout(template="plotly_dark", title="Fulfillment Path (PO Burn-up)", height=450)
    return fig