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
    'stress_line': '#e91e63'
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
    """improved fuzzy mapping to handle variations in user headers."""
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
    """standardizes uploaded production shot data with fuzzy matching."""
    df_list = []
    prod_mapping = {
        "po_number": ["PO_NUMBER", "PO #", "PURCHASE ORDER", "PO_NUM"],
        "project": ["PROJECT", "PROJ", "PROGRAM"],
        "part_id": ["PART_ID", "PART NUMBER", "PART", "ITEM"],
        "tool_id": ["TOOLING ID", "EQUIPMENT", "TOOL_ID", "ASSET", "TOOL"],
        "shot_time": ["SHOT TIME", "TIMESTAMP", "DATE", "TIME", "STAMP"],
        "actual_ct": ["ACTUAL CT", "ACTUAL_CT", "CYCLE TIME", "ACTUAL_CYCLE"],
        "approved_ct": ["APPROVED CT", "APPROVED_CT", "STD CT", "IDEAL CT"],
        "working_cavities": ["WORKING CAVITIES", "CAVITIES", "CAV", "CAVS"]
    }

    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            df = find_and_rename_cols(df, prod_mapping)

            if "shot_time" in df.columns:
                df["shot_time"] = pd.to_datetime(df["shot_time"], errors="coerce")
                
                # handle missing critical numeric columns
                if "actual_ct" not in df.columns:
                    # try to calculate from timestamp diff if missing
                    df = df.sort_values("shot_time")
                    df["actual_ct"] = df["shot_time"].diff().dt.total_seconds().fillna(30)
                else:
                    df["actual_ct"] = pd.to_numeric(df["actual_ct"], errors="coerce")
                
                if "working_cavities" not in df.columns:
                    df["working_cavities"] = 1
                else:
                    df["working_cavities"] = pd.to_numeric(df["working_cavities"], errors="coerce").fillna(1)

                df.dropna(subset=["shot_time"], inplace=True)
                df_list.append(df)
        except Exception: continue
    
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

def load_po_data(files):
    """standardizes uploaded planning/po data with fuzzy matching."""
    df_list = []
    po_mapping = {
        "po_number": ["PO_NUMBER", "PO #", "PURCHASE ORDER"], 
        "project": ["PROJECT", "PROJ"], 
        "part_id": ["PART_ID", "PART", "ITEM"], 
        "total_qty": ["TOTAL_QTY", "QTY", "TARGET", "ORDER_QTY"], 
        "due_date": ["DUE_DATE", "DUE", "DEADLINE"]
    }

    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            df = find_and_rename_cols(df, po_mapping)
            if 'due_date' in df.columns:
                df['due_date'] = pd.to_datetime(df['due_date'], errors='coerce')
            if 'total_qty' in df.columns:
                df['total_qty'] = pd.to_numeric(df['total_qty'], errors='coerce')
            df_list.append(df)
        except Exception: continue
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

# --- calculation engine ---
class CapacityRiskCalculator:
    """core logic to identify production runs, downtime, and cycle efficiency."""
    def __init__(self, df, config):
        self.df = df.copy()
        self.config = config
        self.results = self._calculate()

    def _calculate(self):
        if self.df.empty: return {}
        df = self.df.sort_values("shot_time").reset_index(drop=True)
        
        # identify production runs (default 8hr gap = new run)
        df['time_diff'] = df['shot_time'].diff().dt.total_seconds().fillna(0)
        is_new_run = df['time_diff'] > (self.config.get('run_interval_hours', 8) * 3600)
        df['run_id'] = is_new_run.cumsum()
        
        # calculate mode ct (physics of the tool)
        run_modes = df.groupby('run_id')['actual_ct'].apply(lambda x: x.mode().iloc[0] if not x.mode().empty else x.mean())
        df['mode_ct'] = df['run_id'].map(run_modes)
        df['mode_lower'] = df['mode_ct'] * (1 - self.config.get('tolerance', 0.05))
        df['mode_upper'] = df['mode_ct'] * (1 + self.config.get('tolerance', 0.05))
        
        # identify downtime vs abnormal speed
        df['next_diff'] = df['time_diff'].shift(-1).fillna(0)
        is_gap = df['next_diff'] > (df['actual_ct'] + 2.0)
        is_abnormal = (df['actual_ct'] < df['mode_lower']) | (df['actual_ct'] > df['mode_upper'])
        df['stop_flag'] = np.where(is_gap | is_abnormal, 1, 0)
        df.loc[is_new_run, 'stop_flag'] = 0
        
        # classification
        if 'approved_ct' not in df.columns: 
            df['approved_ct'] = df['mode_ct']
        
        conditions = [
            df['stop_flag'] == 1,
            df['actual_ct'] > (df['approved_ct'] * 1.05),
            df['actual_ct'] < (df['approved_ct'] * 0.95)
        ]
        choices = ['Downtime (Stop)', 'Slow Cycle', 'Fast Cycle']
        df['shot_type'] = np.select(conditions, choices, default='On Target')

        # volume calculations
        prod_df = df[df['stop_flag'] == 0]
        actual_output = prod_df['working_cavities'].sum()
        
        total_runtime_sec = 0
        for _, run in df.groupby('run_id'):
            if not run.empty:
                total_runtime_sec += (run['shot_time'].max() - run['shot_time'].min()).total_seconds() + run.iloc[-1]['actual_ct']
        
        downtime_sec = max(0, total_runtime_sec - prod_df['actual_ct'].sum())
        avg_app_ct = df['approved_ct'].mean()
        avg_cav = df['working_cavities'].mean()
        
        optimal_output = (total_runtime_sec / avg_app_ct) * avg_cav if avg_app_ct > 0 else 0
        loss_dt = (downtime_sec / avg_app_ct) * avg_cav if avg_app_ct > 0 else 0
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
            "total_shots": len(df),
            "normal_shots": len(prod_df)
        }

def get_supply_metrics(df, config):
    if df.empty: return pd.DataFrame()
    df = df.copy()
    df['period'] = df['shot_time'].dt.to_period('W').apply(lambda r: r.start_time)
    
    data = []
    for period, subset in df.groupby('period'):
        calc = CapacityRiskCalculator(subset, config)
        res = calc.results
        data.append({
            'period': period,
            'actual_output': res.get('actual_output', 0),
            'runtime_sec': res.get('total_runtime_sec', 0),
            'total_shots': res.get('total_shots', 0)
        })
    return pd.DataFrame(data)

def generate_capacity_insights(res):
    if not res: return "no data available."
    gap = res.get('optimal_output', 0) - res.get('actual_output', 0)
    if gap <= 0: return "tooling is operating at theoretical capacity."
    dt_p = (res.get('loss_dt', 0) / gap * 100) if gap > 0 else 0
    sl_p = (res.get('loss_slow', 0) / gap * 100) if gap > 0 else 0
    driver = "downtime" if dt_p > sl_p else "slow cycles"
    return f"the primary constraint is <b>{driver}</b>, accounting for {max(dt_p, sl_p):.1f}% of capacity losses."

# --- visualization ---
def create_modern_gauge(value, title):
    color = colors['green']
    if value <= 50: color = colors['red']
    elif value <= 75: color = colors['orange']
    fig = go.Figure(data=[go.Pie(
        values=[value, 100-value, 100], hole=0.7, sort=False, direction='clockwise', rotation=-90,
        marker=dict(colors=[color, '#262730', 'rgba(0,0,0,0)']), hoverinfo='none', textinfo='none'
    )])
    fig.add_annotation(text=f"{value:.1f}%", x=0.5, y=0.15, font=dict(size=42, color='white', weight='bold'), showarrow=False)
    fig.update_layout(title=dict(text=title, x=0.5, font=dict(size=18)), margin=dict(l=20, r=20, t=40, b=0), height=220, showlegend=False, paper_bgcolor='rgba(0,0,0,0)')
    return fig

def create_time_donut(total_sec, prod_sec, down_sec):
    fig = go.Figure(data=[go.Pie(
        labels=['production', 'downtime'], values=[prod_sec, down_sec], hole=0.7,
        marker=dict(colors=[colors['green'], colors['red']]), textinfo='none'
    )])
    center_text = f"total runtime<br><span style='font-size:22px; font-weight:bold;'>{format_seconds_to_dhm(total_sec)}</span>"
    fig.add_annotation(text=center_text, x=0.5, y=0.5, font=dict(size=14, color='white'), showarrow=False)
    fig.update_layout(title="time allocation", showlegend=True, legend=dict(orientation="h", y=-0.1), height=300, margin=dict(t=50, b=50, l=20, r=20), paper_bgcolor='rgba(0,0,0,0)')
    return fig

def create_hybrid_chart(df_agg, cal_config):
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_agg['period'], y=df_agg['actual_output'], name='actual output', marker_color=colors['blue'], yaxis='y1'))
    if 'target' in df_agg.columns:
        df_agg['acc_pct'] = (df_agg['actual_output'] / df_agg['target'] * 100).fillna(0)
        fig.add_trace(go.Scatter(x=df_agg['period'], y=df_agg['acc_pct'], name='accomplishment %', line=dict(color=colors['red'], width=3), yaxis='y2'))
    if cal_config:
        planned = cal_config['days'] * cal_config['hours']
        df_agg['stress'] = (df_agg['runtime_sec'] / 3600 / planned * 100).clip(upper=200)
        fig.add_trace(go.Scatter(x=df_agg['period'], y=df_agg['stress'], name='stress factor %', line=dict(color=colors['stress_line'], width=2, dash='dot'), yaxis='y2'))
    fig.update_layout(template="plotly_dark", hovermode="x unified", yaxis=dict(title="volume"), yaxis2=dict(title="%", overlaying='y', side='right', range=[0, 150], showgrid=False), legend=dict(orientation="h", y=-0.2), height=480, paper_bgcolor='rgba(0,0,0,0)')
    return fig

def plot_waterfall(res):
    fig = go.Figure(go.Waterfall(
        name="bridge", orientation="v", measure=["absolute", "relative", "relative", "relative", "total"],
        x=["optimal", "downtime", "slow cycle", "fast cycle", "actual"],
        y=[res.get('optimal_output', 0), -res.get('loss_dt', 0), -res.get('loss_slow', 0), res.get('gain_fast', 0), res.get('actual_output', 0)],
        decreasing={"marker": {"color": colors['red']}}, increasing={"marker": {"color": colors['green']}},
        totals={"marker": {"color": colors['blue']}}
    ))
    fig.update_layout(title="capacity loss bridge", template="plotly_dark", height=450)
    return fig

def plot_shot_analysis(df_shots):
    fig = go.Figure()
    cmap = {'Slow Cycle': colors['red'], 'Fast Cycle': colors['orange'], 'On Target': colors['blue'], 'Downtime (Stop)': '#808080'}
    for stype, color in cmap.items():
        sub = df_shots[df_shots['shot_type'] == stype]
        if not sub.empty:
            fig.add_trace(go.Bar(x=sub['shot_time'], y=sub['actual_ct'], name=stype, marker_color=color))
    fig.update_layout(template="plotly_dark", title="cycle analysis (shot-by-shot)", yaxis_title="seconds", height=400)
    return fig

def create_po_burnup(df_agg, target_qty, target_date):
    if df_agg.empty: return go.Figure()
    df = df_agg.sort_values('period').copy()
    df['cumulative'] = df['actual_output'].cumsum()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['period'], y=df['cumulative'], name='actual', fill='tozeroy', line_color=colors['blue']))
    if target_qty > 0:
        fig.add_trace(go.Scatter(x=[df['period'].min(), target_date], y=[0, target_qty], mode='lines', name='target', line=dict(color=colors['purple'], dash='dash', shape='hv')))
    fig.update_layout(template="plotly_dark", title="po burn-up staircase", height=450)
    return fig