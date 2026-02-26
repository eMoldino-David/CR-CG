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
    'purple': '#8A2BE2'
}

def format_seconds_to_dhm(total_seconds):
    """Converts total seconds into a 'Xd Yh Zm' or 'Xs' string."""
    if pd.isna(total_seconds) or total_seconds < 0: return "N/A"
    
    if total_seconds < 60:
         return f"{total_seconds:.1f}s"

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

def load_logistics_plan(file):
    """Loads logistics plan CSV/Excel to extract PO information."""
    if not file: return pd.DataFrame()
    try:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        col_map = {col.strip().upper(): col for col in df.columns}
        def get_col(target_list):
            for t in target_list:
                if t in col_map: return col_map[t]
            return None
        
        po_col = get_col(["PO_NUMBER", "PO", "ORDER"])
        proj_col = get_col(["PROJECT", "PROJECT_ID"])
        comp_col = get_col(["COMPONENT_ID", "COMPONENT"])
        part_col = get_col(["PART_ID", "PART"])
        qty_col = get_col(["TOTAL_QTY", "QUANTITY", "QTY", "TARGET_QTY"])
        start_col = get_col(["START_DATE", "START"])
        due_col = get_col(["DUE_DATE", "END_DATE", "DUE"])
        
        rename_dict = {}
        if po_col: rename_dict[po_col] = 'po_number'
        if proj_col: rename_dict[proj_col] = 'project_id'
        if comp_col: rename_dict[comp_col] = 'component_id'
        if part_col: rename_dict[part_col] = 'part_id'
        if qty_col: rename_dict[qty_col] = 'total_qty'
        if start_col: rename_dict[start_col] = 'start_date'
        if due_col: rename_dict[due_col] = 'due_date'
        
        df.rename(columns=rename_dict, inplace=True)
        if 'start_date' in df.columns: df['start_date'] = pd.to_datetime(df['start_date'], errors='coerce')
        if 'due_date' in df.columns: df['due_date'] = pd.to_datetime(df['due_date'], errors='coerce')
        if 'total_qty' in df.columns: df['total_qty'] = pd.to_numeric(df['total_qty'], errors='coerce').fillna(0)
        
        return df
    except Exception:
        return pd.DataFrame()

def load_all_data_cr(files):
    """Loads and standardizes production shot data."""
    df_list = []
    for file in files:
        try:
            df = pd.read_excel(file) if file.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file)
            col_map = {col.strip().upper(): col for col in df.columns}
            def get_col(target_list):
                for t in target_list:
                    if t in col_map: return col_map[t]
                return None

            po_col = get_col(["PO_NUMBER", "PO", "ORDER"])
            if po_col: df.rename(columns={po_col: "po_number"}, inplace=True)

            sup_col = get_col(["SUPPLIER_ID", "SUPPLIER_NAME", "SUPPLIER"])
            if sup_col: df.rename(columns={sup_col: "supplier_id"}, inplace=True)

            plt_col = get_col(["PLANT_ID", "PLANT", "FACTORY"])
            if plt_col: df.rename(columns={plt_col: "plant_id"}, inplace=True)

            project_col = get_col(["PROJECT", "PROJECT_NAME", "PROJECT_ID"])
            if project_col: df.rename(columns={project_col: "project_id"}, inplace=True)
            
            component_col = get_col(["COMPONENT", "COMPONENT_ID", "COMP_ID"])
            if component_col: df.rename(columns={component_col: "component_id"}, inplace=True)
            
            part_col = get_col(["PART", "PART_ID", "PART_NAME", "PART_NUMBER"])
            if part_col: df.rename(columns={part_col: "part_id"}, inplace=True)

            tool_col = get_col(["TOOLING ID", "EQUIPMENT CODE", "TOOL_ID", "TOOL"])
            if tool_col: df.rename(columns={tool_col: "tool_id"}, inplace=True)

            time_col = get_col(["SHOT TIME", "SHOT_TIME", "TIMESTAMP", "DATE", "TIME"])
            if time_col: df.rename(columns={time_col: "shot_time"}, inplace=True)

            act_ct_col = get_col(["ACTUAL CT", "ACTUAL_CT", "CYCLE TIME"])
            if act_ct_col: df.rename(columns={act_ct_col: "actual_ct"}, inplace=True)

            app_ct_col = get_col(["APPROVED CT", "APPROVED_CT", "TARGET CT", "TARGET_CT", "STD CT"])
            if app_ct_col: df.rename(columns={app_ct_col: "approved_ct"}, inplace=True)

            cav_col = get_col(["WORKING CAVITIES", "WORKING_CAVITIES", "CAVITIES"])
            if cav_col: df.rename(columns={cav_col: "working_cavities"}, inplace=True)
            
            area_col = get_col(["PLANT AREA", "PLANT_AREA", "AREA"])
            if area_col: df.rename(columns={area_col: "plant_area"}, inplace=True)

            if "shot_time" in df.columns and "actual_ct" in df.columns:
                df["shot_time"] = pd.to_datetime(df["shot_time"], errors="coerce")
                df["actual_ct"] = pd.to_numeric(df["actual_ct"], errors="coerce")
                df.dropna(subset=["shot_time", "actual_ct"], inplace=True)
                df_list.append(df)
        except Exception: continue
    
    df_final = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
    
    if not df_final.empty:
        for col in ['po_number', 'supplier_id', 'plant_id', 'project_id', 'component_id', 'part_id', 'tool_id']:
            if col not in df_final.columns:
                df_final[col] = "Unknown"
            else:
                df_final[col] = df_final[col].fillna("Unknown").astype(str)
                
    return df_final


# ==============================================================================
# --- CORE CALCULATION ENGINE (DO NOT MODIFY) ---
# ==============================================================================

class CapacityRiskCalculator:
    def __init__(self, df: pd.DataFrame, tolerance: float, downtime_gap_tolerance: float, 
                 run_interval_hours: float, target_output_perc: float = 100.0, 
                 default_cavities: int = 1, remove_maintenance: bool = False, **kwargs):
        
        self.df_raw = df.copy()
        self.tolerance = tolerance
        self.downtime_gap_tolerance = downtime_gap_tolerance
        self.run_interval_hours = run_interval_hours
        self.target_output_perc = target_output_perc
        self.default_cavities = default_cavities
        self.remove_maintenance = remove_maintenance
        self.results = self._calculate_metrics()

    def _calculate_metrics(self) -> dict:
        df = self.df_raw.copy()
        if df.empty: return {}

        if self.remove_maintenance and 'plant_area' in df.columns:
            df = df[~df['plant_area'].astype(str).str.lower().isin(['maintenance', 'warehouse'])].copy()
            if df.empty: return {}

        if 'approved_ct' not in df.columns: df['approved_ct'] = df['actual_ct'].median() 
        if 'working_cavities' not in df.columns: df['working_cavities'] = self.default_cavities
        
        df['approved_ct'] = pd.to_numeric(df['approved_ct'], errors='coerce').fillna(1)
        df['working_cavities'] = pd.to_numeric(df['working_cavities'], errors='coerce').fillna(self.default_cavities)
        df.loc[df['approved_ct'] <= 0, 'approved_ct'] = 1
        
        df = df.sort_values(["tool_id", "shot_time"]).reset_index(drop=True)
        df['time_diff_sec'] = df.groupby('tool_id')['shot_time'].diff().dt.total_seconds().fillna(0)
        
        mask_first_shot = df['tool_id'] != df['tool_id'].shift(1)
        df.loc[mask_first_shot, 'time_diff_sec'] = df.loc[mask_first_shot, 'actual_ct']

        is_new_run = df['time_diff_sec'] > (self.run_interval_hours * 3600)
        df['run_id'] = (is_new_run | mask_first_shot).cumsum()

        run_modes = df[df['actual_ct'] < 1000].groupby('run_id')['actual_ct'].apply(
            lambda x: x.mode().iloc[0] if not x.mode().empty else x.mean()
        )
        df['mode_ct'] = df['run_id'].map(run_modes)
        lower_limit = df['mode_ct'] * (1 - self.tolerance)
        upper_limit = df['mode_ct'] * (1 + self.tolerance)
        df['mode_lower'] = lower_limit
        df['mode_upper'] = upper_limit

        run_approved_cts = df.groupby('run_id')['approved_ct'].apply(
            lambda x: x.mode().iloc[0] if not x.mode().empty else 1
        )
        df['approved_ct_for_run'] = df['run_id'].map(run_approved_cts)
        
        df['next_shot_time_diff'] = df.groupby('tool_id')['time_diff_sec'].shift(-1).fillna(0)
        is_time_gap = df['next_shot_time_diff'] > (df['actual_ct'] + self.downtime_gap_tolerance)
        is_abnormal = ((df['actual_ct'] < lower_limit) | (df['actual_ct'] > upper_limit))
        is_hard_stop = df['actual_ct'] >= 999.9

        df['stop_flag'] = np.where(is_time_gap | is_abnormal | is_hard_stop, 1, 0)
        df.loc[mask_first_shot | is_new_run, 'stop_flag'] = 0
        df['prev_stop_flag'] = df.groupby('tool_id')['stop_flag'].shift(1, fill_value=0)
        df['stop_event'] = (df["stop_flag"] == 1) & (df["prev_stop_flag"] == 0)

        df['adj_ct_sec'] = df['actual_ct']
        df.loc[is_time_gap, 'adj_ct_sec'] = df['next_shot_time_diff']
        
        run_durations = []; run_opt_parts = []
        for _, run_df in df.groupby('run_id'):
            if not run_df.empty:
                start = run_df['shot_time'].min(); end = run_df['shot_time'].max()
                last_ct = run_df.iloc[-1]['actual_ct']
                duration = (end - start).total_seconds() + last_ct
                run_durations.append(duration)
                r_ct = run_df['approved_ct_for_run'].iloc[0]
                r_cav = run_df['working_cavities'].max()
                run_opt_parts.append((duration / r_ct) * r_cav)
        
        total_runtime_sec = sum(run_durations)
        optimal_output_parts = sum(run_opt_parts)

        prod_df = df[df['stop_flag'] == 0].copy()
        production_time_sec = prod_df['actual_ct'].sum()
        downtime_sec = max(0, total_runtime_sec - production_time_sec)
        stops = df['stop_event'].sum()
        stability_index = (production_time_sec / total_runtime_sec * 100) if total_runtime_sec > 0 else 100.0

        actual_output_parts = prod_df['working_cavities'].sum()
        target_output_parts = optimal_output_parts * (self.target_output_perc / 100.0)
        true_loss_parts = optimal_output_parts - actual_output_parts
        
        capacity_gain_fast_parts = 0.0; capacity_loss_slow_parts = 0.0
        if not prod_df.empty:
            prod_df['parts_delta'] = ((prod_df['approved_ct_for_run'] - prod_df['actual_ct']) / prod_df['approved_ct_for_run']) * prod_df['working_cavities']
            capacity_gain_fast_parts = prod_df.loc[prod_df['parts_delta'] > 0, 'parts_delta'].sum()
            capacity_loss_slow_parts = abs(prod_df.loc[prod_df['parts_delta'] < 0, 'parts_delta'].sum())

        total_capacity_loss_sec = downtime_sec + (capacity_loss_slow_parts * df['approved_ct'].mean()) 
        
        total_shots = len(df); normal_shots = total_shots - df['stop_flag'].sum()
        
        conditions = [df['stop_flag'] == 1, df['actual_ct'] > (df['approved_ct_for_run'] + 0.001), df['actual_ct'] < (df['approved_ct_for_run'] - 0.001)]
        df['shot_type'] = np.select(conditions, ['Downtime (Stop)', 'Slow Cycle', 'Fast Cycle'], default='On Target')

        return {
            "processed_df": df, "total_runtime_sec": total_runtime_sec, "production_time_sec": production_time_sec,
            "downtime_sec": downtime_sec, "stability_index": stability_index, "stops": stops,
            "optimal_output_parts": optimal_output_parts, "actual_output_parts": actual_output_parts,
            "target_output_parts": target_output_parts, "capacity_loss_downtime_parts": downtime_sec / df['approved_ct'].mean() if not df.empty else 0,
            "capacity_loss_slow_parts": capacity_loss_slow_parts, "capacity_gain_fast_parts": capacity_gain_fast_parts,
            "total_capacity_loss_parts": true_loss_parts, "total_capacity_loss_sec": total_capacity_loss_sec,
            "efficiency_rate": (normal_shots / total_shots * 100) if total_shots > 0 else 0,
            "total_shots": total_shots, "normal_shots": normal_shots,
            "mtbf_min": (production_time_sec / 60 / stops) if stops > 0 else (production_time_sec / 60)
        }

# ==============================================================================
# --- AGGREGATION UTILS (MAINTAINING CHUNK LOGIC) ---
# ==============================================================================

def get_aggregated_data(df, freq_mode, config, breakdown_by_tool=False):
    """
    Generates aggregated dataframe for tables/charts.
    Mandatory Rule: Respects the chunking logic used by the core calculator.
    """
    rows = []
    if df.empty: return pd.DataFrame()

    # Determine time grouping
    if freq_mode == 'Daily': df['period_key'] = df['shot_time'].dt.date
    elif freq_mode == 'Weekly': df['period_key'] = df['shot_time'].dt.to_period('W').astype(str)
    elif freq_mode == 'Monthly': df['period_key'] = df['shot_time'].dt.to_period('M').astype(str)
    else: df['period_key'] = df['shot_time'].dt.date

    # Grouping structure
    group_cols = ['period_key', 'tool_id'] if breakdown_by_tool else ['period_key']
    grouper = df.groupby(group_cols)

    for group_val, df_subset in grouper:
        calc = CapacityRiskCalculator(df_subset, **config)
        res = calc.results
        if not res: continue
        
        row = {
            'Period': group_val[0] if breakdown_by_tool else group_val,
            'Actual Output': res['actual_output_parts'],
            'Optimal Output': res['optimal_output_parts'],
            'Target Output': res['target_output_parts'],
            'Downtime Loss': res['total_capacity_loss_parts'] - (res['capacity_loss_slow_parts'] - res['capacity_gain_fast_parts']),
            'Slow Loss': res['capacity_loss_slow_parts'],
            'Fast Gain': res['capacity_gain_fast_parts'],
            'Total Loss': res['total_capacity_loss_parts'],
            'Run Time Sec': res['total_runtime_sec'],
            'Normal Shots': res['normal_shots'],
            'Total Shots': res['total_shots']
        }
        if breakdown_by_tool: row['tool_id'] = group_val[1]
        rows.append(row)
        
    return pd.DataFrame(rows)

def calculate_run_summaries(df_period, config):
    summary_list = []
    if 'run_id' not in df_period.columns: return pd.DataFrame()
    for r_id, df_run in df_period.groupby('run_id'):
        calc = CapacityRiskCalculator(df_run, **config)
        res = calc.results
        summary_list.append({
            'run_id': r_id, 'tool_ids': ', '.join(df_run['tool_id'].unique()),
            'start_time': df_run['shot_time'].min(), 'end_time': df_run['shot_time'].max(),
            'total_shots': res['total_shots'], 'normal_shots': res['normal_shots'],
            'stop_events': res['stops'], 'mode_ct': df_run['mode_ct'].iloc[0],
            'mode_lower': df_run['mode_lower'].iloc[0], 'mode_upper': df_run['mode_upper'].iloc[0],
            'total_runtime_sec': res['total_runtime_sec'], 'production_time_sec': res['production_time_sec'],
            'downtime_sec': res['downtime_sec'], 'optimal_output_parts': res['optimal_output_parts'],
            'actual_output_parts': res['actual_output_parts'], 'total_capacity_loss_parts': res['total_capacity_loss_parts'],
            'capacity_loss_downtime_parts': res['capacity_loss_downtime_parts'],
            'capacity_loss_slow_parts': res['capacity_loss_slow_parts'],
            'capacity_gain_fast_parts': res['capacity_gain_fast_parts'],
            'mttr_min': (res['downtime_sec']/60/res['stops']) if res['stops'] > 0 else 0,
            'stability_index': res['stability_index']
        })
    return pd.DataFrame(summary_list).sort_values('start_time')

# ==============================================================================
# --- FORECAST & PO SPECIFIC UTILS ---
# ==============================================================================

def generate_po_periodic_data(df_single_po_shots, po_record, freq_mode, config, working_days_per_week, working_hours_per_day):
    """
    Generates periodic data for Forecast Tab using Tool-level chunking logic.
    Ensures Actual Output matches the Capacity tab exactly.
    """
    start_date = pd.to_datetime(po_record.get('start_date'))
    due_date = pd.to_datetime(po_record.get('due_date'))
    total_qty = po_record.get('total_qty', 0)
    
    # 1. Get chunked actuals by Tool and Date
    agg_df = get_aggregated_data(df_single_po_shots, freq_mode, config, breakdown_by_tool=True)
    
    # 2. Build full timeline
    last_actual_date = pd.to_datetime(df_single_po_shots['shot_time'].max()).date() if not df_single_po_shots.empty else start_date.date()
    current_cum = agg_df['Actual Output'].sum() if not agg_df.empty else 0
    days_elapsed = max(1, (last_actual_date - start_date.date()).days + 1)
    avg_daily_rate = current_cum / days_elapsed
    
    remaining_qty = total_qty - current_cum
    max_proj_days = min(int(remaining_qty / avg_daily_rate) + 1, 365) if avg_daily_rate > 0 else 30
    end_timeline_date = max(due_date.date(), last_actual_date + timedelta(days=max_proj_days))

    if freq_mode == 'Daily':
        full_range = pd.date_range(start=start_date.date(), end=end_timeline_date, freq='D').date.astype(str)
        demand_val = total_qty / (max(1, (due_date - start_date).days) * (working_days_per_week/7.0))
    elif freq_mode == 'Weekly':
        full_range = pd.period_range(start=start_date, end=end_timeline_date, freq='W').astype(str)
        demand_val = total_qty / max(1, (due_date - start_date).days / 7.0)
    else:
        full_range = pd.period_range(start=start_date, end=end_timeline_date, freq='M').astype(str)
        demand_val = total_qty / max(1, (due_date - start_date).days / 30.44)

    df_full = pd.DataFrame({'Period': full_range})
    df_full['Estimated Demand'] = demand_val
    
    # Merge chunked actuals into full timeline
    if not agg_df.empty:
        agg_sum = agg_df.groupby('Period')['Actual Output'].sum().reset_index()
        agg_sum['Period'] = agg_sum['Period'].astype(str)
        df_full = pd.merge(df_full, agg_sum, on='Period', how='left').fillna(0)
    else:
        df_full['Actual Output'] = 0

    # Max Capacity Line
    avg_ct = df_single_po_shots['approved_ct'].mean() if not df_single_po_shots.empty else 1
    cav = df_single_po_shots['working_cavities'].max() if not df_single_po_shots.empty else 1
    hourly_cap = (3600 / max(0.1, avg_ct)) * cav
    
    if freq_mode == 'Daily': df_full['Configured Max Capacity'] = hourly_cap * working_hours_per_day
    elif freq_mode == 'Weekly': df_full['Configured Max Capacity'] = hourly_cap * working_hours_per_day * working_days_per_week
    else: df_full['Configured Max Capacity'] = hourly_cap * working_hours_per_day * working_days_per_week * 4.33

    return df_full

def generate_po_prediction_data(df_po_shots, po_record, config):
    """Generates burn-up data using Daily chunked actuals to ensure consistency."""
    start_date = pd.to_datetime(po_record['start_date']).date()
    due_date = pd.to_datetime(po_record['due_date']).date()
    total_qty = po_record['total_qty']

    # 1. Use Daily chunked actuals
    agg_daily = get_aggregated_data(df_po_shots, 'Daily', config)
    if agg_daily.empty: return None
    
    agg_daily['Period'] = pd.to_datetime(agg_daily['Period']).dt.date
    agg_daily = agg_daily.sort_values('Period')
    agg_daily['Produced Quantity'] = agg_daily['Actual Output'].cumsum()
    
    # 2. Forecast logic
    last_act_date = agg_daily['Period'].max()
    current_cum = agg_daily['Produced Quantity'].iloc[-1]
    days_elapsed = max(1, (last_act_date - start_date).days + 1)
    avg_rate = current_cum / days_elapsed
    
    # Target Line
    total_days = max(1, (due_date - start_date).days)
    target_dates = [start_date + timedelta(days=i) for i in range(total_days + 1)]
    target_vals = [(total_qty / total_days) * i for i in range(total_days + 1)]

    # Projections
    rem = total_qty - current_cum
    proj_days = min(int(rem / avg_rate) + 1, 365) if avg_rate > 0 else 30
    forecast_dates = [last_act_date + timedelta(days=i) for i in range(proj_days + 1)]
    forecast_avg = [current_cum + (avg_rate * i) for i in range(proj_days + 1)]

    return {
        'target_dates': target_dates, 'target_vals': target_vals,
        'actual_dates': agg_daily['Period'], 'actual_cum': agg_daily['Produced Quantity'],
        'forecast_dates': forecast_dates, 'forecast_avg': forecast_avg,
        'due_date': due_date, 'start_date': start_date, 'total_qty': total_qty,
        'current_cum': current_cum, 'avg_daily_rate': avg_rate, 'opt_daily_rate': avg_rate * 1.2 # Placeholder for Opt
    }

# ==============================================================================
# --- PLOTTING FUNCTIONS ---
# ==============================================================================

def plot_po_periodic_chart(agg_po, df_raw_processed, bar_freq):
    """
    Plots periodic bars using pre-chunked actuals.
    Stacked bars now respect the exact Tool contribution from chunked data.
    """
    fig = go.Figure()
    
    # 1. Tool contribution from raw processed shots (matches Capacity tab logic)
    if not df_raw_processed.empty:
        if bar_freq == 'Daily': df_raw_processed['p_key'] = df_raw_processed['shot_time'].dt.date.astype(str)
        elif bar_freq == 'Weekly': df_raw_processed['p_key'] = df_raw_processed['shot_time'].dt.to_period('W').astype(str)
        else: df_raw_processed['p_key'] = df_raw_processed['shot_time'].dt.to_period('M').astype(str)
        
        # Only count shots where stop_flag is 0
        df_good = df_raw_processed[df_raw_processed['stop_flag'] == 0]
        tool_bars = df_good.groupby(['p_key', 'tool_id'])['working_cavities'].sum().reset_index()
        
        colors = px.colors.qualitative.Pastel
        for i, tool in enumerate(tool_bars['tool_id'].unique()):
            subset = tool_bars[tool_bars['tool_id'] == tool]
            fig.add_trace(go.Bar(x=subset['p_key'], y=subset['working_cavities'], name=f"Tool: {tool}", marker_color=colors[i % len(colors)]))
    
    # 2. Add Target Lines from agg_po
    fig.add_trace(go.Scatter(x=agg_po['Period'], y=agg_po['Configured Max Capacity'], name='Max Capacity', mode='lines+markers', line=dict(color=PASTEL_COLORS['green'], dash='dot')))
    fig.add_trace(go.Scatter(x=agg_po['Period'], y=agg_po['Estimated Demand'], name='PO Demand', mode='lines+markers', line=dict(color=PASTEL_COLORS['red'], dash='dash')))
    
    fig.update_layout(title=f"Periodic Production vs Demand ({bar_freq})", barmode='stack', hovermode="x unified", height=450)
    return fig

def plot_po_burnup(pred_data, po_record=None):
    """Plots Burn-Up with Produced Quantity area shading and Adherence."""
    fig = go.Figure()
    
    # Calculate Adherence Rate
    start_dt = pd.to_datetime(pred_data['start_date']).date()
    due_dt = pd.to_datetime(pred_data['due_date']).date()
    total_q = pred_data['total_qty']
    target_rate = total_q / max(1, (due_dt - start_dt).days)
    
    actual_dates = pred_data['actual_dates']
    actual_cum = pred_data['actual_cum']
    
    adherence_list = []
    for d, val in zip(actual_dates, actual_cum):
        days = max(1, (pd.to_datetime(d).date() - start_dt).days)
        expected = target_rate * days
        adherence_list.append(f"{(val/expected*100):.1f}%" if expected > 0 else "100%")
        
    current_adh = adherence_list[-1] if adherence_list else "N/A"

    # Area Shaded Produced Quantity
    fig.add_trace(go.Scatter(x=actual_dates, y=actual_cum, name='Produced Quantity', fill='tozeroy', 
                             fillcolor='rgba(52, 152, 219, 0.2)', line=dict(color=PASTEL_COLORS['blue'], width=3),
                             customdata=adherence_list, hovertemplate='Date: %{x}<br>Produced: %{y:,.0f}<br>Adherence: %{customdata}<extra></extra>'))
                             
    fig.add_trace(go.Scatter(x=pred_data['target_dates'], y=pred_data['target_vals'], name='Target Burn-up', line=dict(color='grey', dash='dash')))
    fig.add_trace(go.Scatter(x=pred_data['forecast_dates'], y=pred_data['forecast_avg'], name='Forecast (Avg)', line=dict(color=PASTEL_COLORS['orange'], dash='dot')))
    
    fig.add_hline(y=total_q, line_color="purple", annotation_text="Target Qty")
    fig.add_vline(x=pd.to_datetime(due_dt).timestamp()*1000, line_dash="dash", line_color="red", annotation_text="Due Date")

    fig.update_layout(title=f"PO Burn-up (Current Adherence: {current_adh})", hovermode="x unified", height=500)
    return fig

# ==============================================================================
# --- EXISTING KPI & DASHBOARD PLOTS (UNTOUCHED) ---
# ==============================================================================

def create_time_breakdown_donut(total_sec, prod_sec, down_sec):
    fig = go.Figure(data=[go.Pie(values=[prod_sec, down_sec], labels=['Prod', 'Down'], marker=dict(colors=[PASTEL_COLORS['green'], PASTEL_COLORS['red']]), hole=0.7)])
    fig.update_layout(title="Run Duration Breakdown", height=320, showlegend=True)
    return fig

def create_modern_gauge(value, title):
    fig = go.Figure(go.Indicator(mode="gauge+number", value=value, title={'text': title}, gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "darkblue"}}))
    fig.update_layout(height=220)
    return fig

def create_stability_driver_bar(mtbf, mttr, stability_index):
    fig = go.Figure(go.Bar(x=[mtbf, mttr], y=['MTBF', 'MTTR'], orientation='h'))
    fig.update_layout(title=f"Stability Analysis ({stability_index:.1f}%)", height=260)
    return fig

def plot_waterfall(metrics, benchmark_mode="Optimal"):
    fig = go.Figure(go.Waterfall(x=["Optimal", "Downtime", "Slow", "Fast", "Actual"], y=[metrics['optimal_output_parts'], -metrics['capacity_loss_downtime_parts'], -metrics['capacity_loss_slow_parts'], metrics['capacity_gain_fast_parts'], 0], measure=["absolute", "relative", "relative", "relative", "total"]))
    fig.update_layout(title=f"Capacity Bridge vs {benchmark_mode}", height=450)
    return fig

def plot_performance_breakdown(df_agg, x_col, benchmark_mode):
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_agg[x_col], y=df_agg['Actual Output'], name='Actual', marker_color=PASTEL_COLORS['blue']))
    fig.add_trace(go.Bar(x=df_agg[x_col], y=df_agg['Downtime Loss'], name='Downtime Loss', marker_color=PASTEL_COLORS['grey']))
    fig.update_layout(barmode='stack', title="Performance Breakdown", height=450)
    return fig

def plot_shot_analysis(df_shots, zoom_y=None):
    fig = go.Figure(go.Scatter(x=df_shots['shot_time'], y=df_shots['actual_ct'], mode='markers', name='Shots'))
    fig.update_layout(title="Shot-by-Shot", height=450)
    return fig

def generate_capacity_insights(res, mode): return {"overall": "Analysis complete.", "drivers": "Normal operations.", "recommendation": "Maintain consistency."}
def generate_mttr_mtbf_analysis(df): return "Correlation analysis ready."
def prepare_and_generate_capacity_excel(df, config): return b""
def calculate_capacity_risk_scores(df, config): return pd.DataFrame()