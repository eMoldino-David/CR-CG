import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
from io import BytesIO

# ==============================================================================
# --- CONSTANTS ---
# ==============================================================================

COLORS = {
    'actual_bar': '#3498DB',
    'schedule_line': '#ff6961',
    'stress_line': '#E91E63',
    'target_staircase': '#8A2BE2',
    'grid_grey': '#41424C'
}

# ==============================================================================
# --- DATA LOADING & STANDARDIZATION ---
# ==============================================================================

def load_production_data(files):
    """Loads and standardizes shot-by-shot data."""
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
# --- CALCULATION ENGINE ---
# ==============================================================================

class CapacityRiskCalculator:
    def __init__(self, df, config):
        self.df = df.copy()
        self.config = config
        self.results = self._calculate()

    def _calculate(self):
        if self.df.empty: return {}
        
        df = self.df.sort_values("shot_time").reset_index(drop=True)
        
        # 1. Run ID & Stops (Simplified for aggregate view)
        df['time_diff'] = df['shot_time'].diff().dt.total_seconds().fillna(0)
        is_new_run = df['time_diff'] > (self.config['run_interval_hours'] * 3600)
        df['run_id'] = is_new_run.cumsum()
        
        # Mode detection per run
        run_modes = df.groupby('run_id')['actual_ct'].transform(lambda x: x.mode().iloc[0] if not x.mode().empty else x.mean())
        df['mode_ct'] = run_modes
        
        # Stop identification
        df['stop_flag'] = np.where(df['actual_ct'] > (df['mode_ct'] * (1 + self.config['tolerance'])), 1, 0)
        
        # Output logic
        prod_df = df[df['stop_flag'] == 0]
        actual_output = prod_df['working_cavities'].sum() if 'working_cavities' in prod_df.columns else len(prod_df)
        
        # Total Runtime (Total machine occupancy)
        total_runtime_sec = 0
        for _, run in df.groupby('run_id'):
            if not run.empty:
                total_runtime_sec += (run['shot_time'].max() - run['shot_time'].min()).total_seconds() + run.iloc[-1]['actual_ct']

        return {
            "actual_output": actual_output,
            "runtime_sec": total_runtime_sec,
            "df": df
        }

def get_supply_metrics(df, freq, engine_config):
    """Aggregates metrics into time periods for charting."""
    if df.empty: return pd.DataFrame()
    
    if freq == 'Daily': df['Period'] = df['shot_time'].dt.date
    elif freq == 'Weekly': df['Period'] = df['shot_time'].dt.to_period('W').apply(lambda r: r.start_time)
    
    data = []
    for period, subset in df.groupby('Period'):
        calc = CapacityRiskCalculator(subset, engine_config)
        res = calc.results
        data.append({
            'Period': period,
            'Actual Output': res['actual_output'],
            'Run Time Sec': res['runtime_sec']
        })
    return pd.DataFrame(data)

# ==============================================================================
# --- VISUALIZATION ---
# ==============================================================================

def create_hybrid_chart(df_agg, cal_config):
    """Volume Bars + Schedule Accomplishment + Stress Factor."""
    fig = go.Figure()

    # 1. Bars: Production Volume
    fig.add_trace(go.Bar(
        x=df_agg['Period'], y=df_agg['Actual Output'],
        name='Actual Production', marker_color=COLORS['actual_bar'],
        yaxis='y1'
    ))

    # 2. Line: Schedule Accomplishment (Red)
    if 'Target' in df_agg.columns:
        df_agg['Accomplishment'] = (df_agg['Actual Output'] / df_agg['Target'] * 100).fillna(0)
        fig.add_trace(go.Scatter(
            x=df_agg['Period'], y=df_agg['Accomplishment'],
            name='Schedule Accomplishment %', line=dict(color=COLORS['schedule_line'], width=3),
            yaxis='y2'
        ))

    # 3. Line: Stress Factor (Dashed Pink)
    if cal_config:
        planned_hrs_per_period = cal_config['days'] * cal_config['hours']
        # If weekly, multiply by 1. If daily, divide by 7 (handled by calling logic)
        df_agg['Stress'] = (df_agg['Run Time Sec'] / 3600 / planned_hrs_per_period * 100).clip(upper=200)
        fig.add_trace(go.Scatter(
            x=df_agg['Period'], y=df_agg['Stress'],
            name='Tool Stress Factor %', line=dict(color=COLORS['stress_line'], width=2, dash='dot'),
            yaxis='y2'
        ))

    fig.update_layout(
        template="plotly_dark",
        hovermode="x unified",
        yaxis=dict(title="Volume (Parts)", gridcolor=COLORS['grid_grey']),
        yaxis2=dict(title="Percent (%)", overlaying='y', side='right', range=[0, 150], showgrid=False),
        legend=dict(orientation="h", y=-0.2, x=0.5, xanchor='center'),
        height=500, margin=dict(t=50, b=50, l=50, r=50)
    )
    return fig

def create_burnup(df_agg, target_qty, target_date):
    """Cumulative output vs PO Target Step-line."""
    df = df_agg.sort_values('Period')
    df['Cumulative'] = df['Actual Output'].cumsum()
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['Period'], y=df['Cumulative'], name='Cumulative Actual', fill='tozeroy', line_color=COLORS['actual_bar']))
    
    if target_qty > 0:
        fig.add_trace(go.Scatter(
            x=[df['Period'].min(), target_date], 
            y=[0, target_qty],
            mode='lines', name='Demand Staircase',
            line=dict(color=COLORS['target_staircase'], dash='dash', shape='hv')
        ))

    fig.update_layout(template="plotly_dark", title="Fulfillment Path", yaxis_title="Total Parts")
    return fig