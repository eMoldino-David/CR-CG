import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter

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
    'purple': '#8A2BE2',
    'stress_line': '#E91E63'
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

# ==============================================================================
# --- DATA LOADING ---
# ==============================================================================

def load_production_data(files):
    """Loads and standardizes shot-by-shot production data."""
    df_list = []
    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            col_map = {col.strip().upper(): col for col in df.columns}
            
            mapping = {
                "po_number": ["PO_NUMBER", "PO #", "PURCHASE ORDER"],
                "project": ["PROJECT", "PROJECT_NAME"],
                "part_id": ["PART_ID", "PART NUMBER", "PART"],
                "tool_id": ["TOOLING ID", "EQUIPMENT CODE", "TOOL_ID"],
                "shot_time": ["SHOT TIME", "TIMESTAMP", "DATE", "TIME"],
                "actual_ct": ["ACTUAL CT", "ACTUAL_CT", "CYCLE TIME"],
                "approved_ct": ["APPROVED CT", "APPROVED_CT", "STD CT"],
                "working_cavities": ["WORKING CAVITIES", "CAVITIES"]
            }

            for key, targets in mapping.items():
                for t in targets:
                    if t in col_map:
                        df.rename(columns={col_map[t]: key}, inplace=True)
                        break

            if "shot_time" in df.columns and "actual_ct" in df.columns:
                df["shot_time"] = pd.to_datetime(df["shot_time"], errors="coerce")
                df["actual_ct"] = pd.to_numeric(df["actual_ct"], errors="coerce")
                df.dropna(subset=["shot_time", "actual_ct"], inplace=True)
                df_list.append(df)
        except Exception: continue
    
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

def load_po_data(files):
    """Loads Purchase Order demand data."""
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
    def __init__(self, df, config):
        self.df = df.copy()
        self.config = config
        self.results = self._calculate()

    def _calculate(self):
        """The primary brain that identifies runs, stops, and efficiency gaps."""
        if self.df.empty: return {}
        df = self.df.sort_values("shot_time").reset_index(drop=True)
        
        # 1. Run Identification
        df['time_diff'] = df['shot_time'].diff().dt.total_seconds().fillna(0)
        is_new_run = df['time_diff'] > (self.config['run_interval_hours'] * 3600)
        df['run_id'] = is_new_run.cumsum()
        
        # 2. Mode CT & Tolerance Bands
        run_modes = df.groupby('run_id')['actual_ct'].apply(
            lambda x: x.mode().iloc[0] if not x.mode().empty else x.mean()
        )
        df['mode_ct'] = df['run_id'].map(run_modes)
        df['mode_lower'] = df['mode_ct'] * (1 - self.config['tolerance'])
        df['mode_upper'] = df['mode_ct'] * (1 + self.config['tolerance'])
        
        # 3. Stop Detection (Downtime vs Abnormal Cycle)
        df['next_diff'] = df['time_diff'].shift(-1).fillna(0)
        is_gap = df['next_diff'] > (df['actual_ct'] + 2.0) # 2s tolerance for timestamp jitter
        is_abnormal = (df['actual_ct'] < df['mode_lower']) | (df['actual_ct'] > df['mode_upper'])
        
        df['stop_flag'] = np.where(is_gap | is_abnormal, 1, 0)
        df.loc[is_new_run, 'stop_flag'] = 0 
        
        # 4. State Classification
        if 'approved_ct' not in df.columns: df['approved_ct'] = df['mode_ct']
        conditions = [
            df['stop_flag'] == 1,
            df['actual_ct'] > (df['approved_ct'] * 1.05),
            df['actual_ct'] < (df['approved_ct'] * 0.95)
        ]
        choices = ['Downtime (Stop)', 'Slow Cycle', 'Fast Cycle']
        df['shot_type'] = np.select(conditions, choices, default='On Target')

        # 5. Volumetric Calculation
        prod_df = df[df['stop_flag'] == 0]
        actual_output = prod_df['working_cavities'].sum() if 'working_cavities' in prod_df.columns else len(prod_df)
        
        total_runtime_sec = 0
        for _, run in df.groupby('run_id'):
            if not run.empty:
                total_runtime_sec += (run['shot_time'].max() - run['shot_time'].min()).total_seconds() + run.iloc[-1]['actual_ct']
        
        downtime_sec = max(0, total_runtime_sec - prod_df['actual_ct'].sum())
        
        # Optimal Capacity Benchmark
        avg_app_ct = df['approved_ct'].mean()
        avg_cav = df['working_cavities'].mean()
        optimal_output = (total_runtime_sec / avg_app_ct) * avg_cav if avg_app_ct > 0 else 0

        # Loss Allocation
        loss_downtime = (downtime_sec / avg_app_ct) * avg_cav if avg_app_ct > 0 else 0
        # Inefficiency is Actual - (Optimal - Loss_Downtime)
        loss_speed = max(0, (optimal_output - loss_downtime) - actual_output)
        gain_speed = max(0, actual_output - (optimal_output - loss_downtime))

        return {
            "processed_df": df,
            "total_runtime_sec": total_runtime_sec,
            "downtime_sec": downtime_sec,
            "actual_output": actual_output,
            "optimal_output": optimal_output,
            "loss_downtime_parts": loss_downtime,
            "loss_slow_parts": loss_speed,
            "gain_fast_parts": gain_speed,
            "total_shots": len(df),
            "normal_shots": len(prod_df),
            "stops": df[df['stop_flag'] == 1].groupby('run_id').size().sum()
        }

# ==============================================================================
# --- AGGREGATION & INSIGHTS ---
# ==============================================================================

def get_supply_metrics(df, freq, engine_config):
    """Aggregates metrics into time periods for charting."""
    if df.empty: return pd.DataFrame()
    
    df = df.copy()
    if freq == 'Daily': df['Period'] = df['shot_time'].dt.date
    elif freq == 'Weekly': df['Period'] = df['shot_time'].dt.to_period('W').apply(lambda r: r.start_time)
    
    data = []
    for period, subset in df.groupby('Period'):
        calc = CapacityRiskCalculator(subset, engine_config)
        res = calc.results
        data.append({
            'Period': period,
            'Actual Output': res['actual_output'],
            'Optimal Output': res['optimal_output'],
            'Downtime Loss': res['loss_downtime_parts'],
            'Speed Loss': res['loss_slow_parts'],
            'Speed Gain': res['gain_fast_parts'],
            'Run Time Sec': res['total_runtime_sec'],
            'Downtime Sec': res['downtime_sec'],
            'Total Shots': res['total_shots'],
            'Normal Shots': res['normal_shots']
        })
    return pd.DataFrame(data)

def generate_capacity_insights(res):
    """Generates natural language breakdown of constraints."""
    if not res: return "No data."
    
    gap = res['optimal_output'] - res['actual_output']
    if gap <= 0: return "Tooling is operating at or above theoretical capacity."
    
    dt_share = (res['loss_downtime_parts'] / gap * 100) if gap > 0 else 0
    sl_share = (res['loss_slow_parts'] / gap * 100) if gap > 0 else 0
    
    primary = "Downtime" if dt_share > sl_share else "Slow Cycle Times"
    
    return f"The primary constraint is <b>{primary}</b>, representing {max(dt_share, sl_share):.1f}% of the uncaptured capacity."

# ==============================================================================
# --- VISUALIZATION ---
# ==============================================================================

def create_modern_gauge(value, title):
    """Half-donut gauge for high-level KPIs."""
    color = PASTEL_COLORS['green']
    if value <= 50: color = PASTEL_COLORS['red']
    elif value <= 75: color = PASTEL_COLORS['orange']
    
    fig = go.Figure(data=[go.Pie(
        values=[value, 100-value, 100], hole=0.7, sort=False, direction='clockwise', rotation=-90,
        marker=dict(colors=[color, '#262730', 'rgba(0,0,0,0)']), hoverinfo='none', textinfo='none'
    )])
    fig.add_annotation(text=f"{value:.1f}%", x=0.5, y=0.15, font=dict(size=42, color='white', weight='bold'), showarrow=False)
    fig.update_layout(title=dict(text=title, x=0.5, font=dict(size=18)), margin=dict(l=20, r=20, t=40, b=0), height=220, showlegend=False, paper_bgcolor='rgba(0,0,0,0)')
    return fig

def create_time_donut(total_sec, prod_sec, down_sec):
    """Donut for production vs downtime duration."""
    fig = go.Figure(data=[go.Pie(
        labels=['Production Time', 'RR Downtime'], values=[prod_sec, down_sec], hole=0.7,
        marker=dict(colors=[PASTEL_COLORS['green'], PASTEL_COLORS['red']]), textinfo='none'
    )])
    center_text = f"Total Runtime<br><span style='font-size:22px; font-weight:bold;'>{format_seconds_to_dhm(total_sec)}</span>"
    fig.add_annotation(text=center_text, x=0.5, y=0.5, font=dict(size=14, color='white'), showarrow=False)
    fig.update_layout(title="Time Allocation Breakdown", showlegend=True, legend=dict(orientation="h", y=-0.1), height=300, margin=dict(t=50, b=50, l=20, r=20), paper_bgcolor='rgba(0,0,0,0)')
    return fig

def create_hybrid_chart(df_agg, cal_config):
    """Volume Bars + Schedule Accomplishment + Stress Factor (Client Style)."""
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_agg['Period'], y=df_agg['Actual Output'], name='Actual Volume', marker_color=PASTEL_COLORS['blue'], yaxis='y1'))
    
    if 'Target' in df_agg.columns:
        df_agg['Acc %'] = (df_agg['Actual Output'] / df_agg['Target'] * 100).fillna(0)
        fig.add_trace(go.Scatter(x=df_agg['Period'], y=df_agg['Acc %'], name='Schedule Accomplishment %', line=dict(color=PASTEL_COLORS['red'], width=3), yaxis='y2'))
    
    if cal_config:
        planned_hrs = cal_config['days'] * cal_config['hours']
        df_agg['Stress'] = (df_agg['Run Time Sec'] / 3600 / planned_hrs * 100).clip(upper=200)
        fig.add_trace(go.Scatter(x=df_agg['Period'], y=df_agg['Stress'], name='Tool Stress Factor %', line=dict(color=PASTEL_COLORS['stress_line'], width=2, dash='dot'), yaxis='y2'))

    fig.update_layout(
        template="plotly_dark", hovermode="x unified",
        yaxis=dict(title="Parts Output"),
        yaxis2=dict(title="Percentage (%)", overlaying='y', side='right', range=[0, 150], showgrid=False),
        legend=dict(orientation="h", y=-0.2), height=480, paper_bgcolor='rgba(0,0,0,0)'
    )
    return fig

def plot_waterfall(res):
    """Capacity Bridge: Optimal -> Losses -> Actual."""
    fig = go.Figure(go.Waterfall(
        name="Capacity Bridge", orientation="v",
        measure=["absolute", "relative", "relative", "relative", "total"],
        x=["Optimal (Theory)", "Loss: Downtime", "Loss: Speed", "Gain: Speed", "Actual Output"],
        y=[res['optimal_output'], -res['loss_downtime_parts'], -res['loss_slow_parts'], res['gain_fast_parts'], res['actual_output']],
        decreasing={"marker": {"color": PASTEL_COLORS['red']}},
        increasing={"marker": {"color": PASTEL_COLORS['green']}},
        totals={"marker": {"color": PASTEL_COLORS['blue']}}
    ))
    fig.update_layout(title="Capacity Loss Bridge", template="plotly_dark", height=450)
    return fig


def plot_shot_analysis(df_shots):
    """Scatter/Bar chart of every single shot cycle time."""
    fig = go.Figure()
    color_map = {'Slow Cycle': PASTEL_COLORS['red'], 'Fast Cycle': PASTEL_COLORS['orange'], 'On Target': PASTEL_COLORS['blue'], 'Downtime (Stop)': '#808080'}
    
    for stype, color in color_map.items():
        sub = df_shots[df_shots['shot_type'] == stype]
        if not sub.empty:
            fig.add_trace(go.Bar(x=sub['shot_time'], y=sub['actual_ct'], name=stype, marker_color=color, hovertemplate='CT: %{y:.2f}s<br>Time: %{x}'))
    
    # Tolerance band rectangle
    for _, run in df_shots.groupby('run_id'):
        fig.add_shape(type="rect", x0=run['shot_time'].min(), x1=run['shot_time'].max(), y0=run['mode_lower'].iloc[0], y1=run['mode_upper'].iloc[0], fillcolor="grey", opacity=0.1, line_width=0)
    
    fig.update_layout(template="plotly_dark", title="Shot-by-Shot Cycle Analysis", yaxis_title="Seconds", height=400, yaxis_range=[0, df_shots['actual_ct'].quantile(0.98)*1.5])
    return fig

def create_stability_driver_bar(mtbf, mttr, stability):
    """Stacked bar showing if frequency (MTBF) or duration (MTTR) is the driver."""
    total = mtbf + mttr
    if total == 0: return go.Figure()
    
    fig = go.Figure()
    fig.add_trace(go.Bar(y=['Downtime Driver'], x=[mtbf], name=f'MTBF: {mtbf:.1f}m', orientation='h', marker_color=PASTEL_COLORS['blue']))
    fig.add_trace(go.Bar(y=['Downtime Driver'], x=[mttr], name=f'MTTR: {mttr:.1f}m', orientation='h', marker_color=PASTEL_COLORS['red']))
    
    fig.update_layout(barmode='stack', template="plotly_dark", height=200, showlegend=True, title="MTTR vs MTBF Analysis", margin=dict(t=40, b=20, l=20, r=20))
    return fig

def create_po_burnup(df_agg, target_qty, target_date):
    """Burn-up chart with cumulative actuals vs PO target staircase."""
    df = df_agg.sort_values('Period').copy()
    df['Cumulative'] = df['Actual Output'].cumsum()
    
    fig = go.Figure()
    # Actuals
    fig.add_trace(go.Scatter(x=df['Period'], y=df['Cumulative'], name='Actual Output', fill='tozeroy', line_color=PASTEL_COLORS['blue'], mode='lines+markers'))
    
    # PO Target Staircase
    if target_qty > 0:
        fig.add_trace(go.Scatter(
            x=[df['Period'].min(), target_date], 
            y=[0, target_qty],
            mode='lines', name='Demand Target Staircase',
            line=dict(color=PASTEL_COLORS['purple'], dash='dash', shape='hv')
        ))

    fig.update_layout(template="plotly_dark", title="Supply Assurance Burn-up", xaxis_title="Time", yaxis_title="Cumulative Parts", height=450)
    return fig