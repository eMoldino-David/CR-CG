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
    
    # Ensure hierarchical columns exist and are string-formatted for clean filtering
    if not df_final.empty:
        for col in ['po_number', 'supplier_id', 'plant_id', 'project_id', 'component_id', 'part_id', 'tool_id']:
            if col not in df_final.columns:
                df_final[col] = "Unknown"
            else:
                df_final[col] = df_final[col].fillna("Unknown").astype(str)
                
    return df_final


# ==============================================================================
# --- CORE CALCULATION ENGINE ---
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
        
        # Ensure positive Approved CT
        df.loc[df['approved_ct'] <= 0, 'approved_ct'] = 1
        
        # Sort fundamentally respects tool_id first so rolled-up logic functions perfectly
        df = df.sort_values(["tool_id", "shot_time"]).reset_index(drop=True)

        # 1. Run Identification (Grouped by tool_id safely)
        df['time_diff_sec'] = df.groupby('tool_id')['shot_time'].diff().dt.total_seconds().fillna(0)
        
        # Initialize the first shot per tool to actual_ct
        mask_first_shot = df['tool_id'] != df['tool_id'].shift(1)
        df.loc[mask_first_shot, 'time_diff_sec'] = df.loc[mask_first_shot, 'actual_ct']

        is_new_run = df['time_diff_sec'] > (self.run_interval_hours * 3600)
        
        # Generates globally unique Run IDs securely isolated by Tool boundaries
        df['run_id'] = (is_new_run | mask_first_shot).cumsum()

        # 2. Mode CT & Limits
        run_modes = df[df['actual_ct'] < 1000].groupby('run_id')['actual_ct'].apply(
            lambda x: x.mode().iloc[0] if not x.mode().empty else x.mean()
        )
        df['mode_ct'] = df['run_id'].map(run_modes)
        lower_limit = df['mode_ct'] * (1 - self.tolerance)
        upper_limit = df['mode_ct'] * (1 + self.tolerance)
        
        df['mode_lower'] = lower_limit
        df['mode_upper'] = upper_limit

        # 3. Approved CT
        run_approved_cts = df.groupby('run_id')['approved_ct'].apply(
            lambda x: x.mode().iloc[0] if not x.mode().empty else 1
        )
        df['approved_ct_for_run'] = df['run_id'].map(run_approved_cts)
        
        # 4. Stop Detection (Isolated explicitly by tool limits)
        df['next_shot_time_diff'] = df.groupby('tool_id')['time_diff_sec'].shift(-1).fillna(0)
        
        is_time_gap = df['next_shot_time_diff'] > (df['actual_ct'] + self.downtime_gap_tolerance)
        is_abnormal = ((df['actual_ct'] < lower_limit) | (df['actual_ct'] > upper_limit))
        is_hard_stop = df['actual_ct'] >= 999.9

        df['stop_flag'] = np.where(is_time_gap | is_abnormal | is_hard_stop, 1, 0)
        
        # Reset stop_flag for first shots and new runs
        df.loc[mask_first_shot | is_new_run, 'stop_flag'] = 0

        # Protect against stop events bleeding over tool boundaries
        df['prev_stop_flag'] = df.groupby('tool_id')['stop_flag'].shift(1, fill_value=0)
        df['stop_event'] = (df["stop_flag"] == 1) & (df["prev_stop_flag"] == 0)

        df['adj_ct_sec'] = df['actual_ct']
        df.loc[is_time_gap, 'adj_ct_sec'] = df['next_shot_time_diff']
        
        # --- Metrics Calculation ---
        run_durations = []
        run_opt_parts = []
        
        for _, run_df in df.groupby('run_id'):
            if not run_df.empty:
                start = run_df['shot_time'].min()
                end = run_df['shot_time'].max()
                last_ct = run_df.iloc[-1]['actual_ct']
                duration = (end - start).total_seconds() + last_ct
                run_durations.append(duration)
                
                # Calculate Optimal Output precisely per run to prevent multi-tool averages from skewing output
                r_ct = run_df['approved_ct_for_run'].iloc[0]
                r_cav = run_df['working_cavities'].max()
                run_opt_parts.append((duration / r_ct) * r_cav)
        
        total_runtime_sec = sum(run_durations)
        optimal_output_parts = sum(run_opt_parts)

        prod_df = df[df['stop_flag'] == 0].copy()
        production_time_sec = prod_df['actual_ct'].sum()
        downtime_sec = max(0, total_runtime_sec - production_time_sec)

        stops = df['stop_event'].sum()
        mttr_min = (downtime_sec / 60 / stops) if stops > 0 else 0
        stability_index = (production_time_sec / total_runtime_sec * 100) if total_runtime_sec > 0 else 100.0

        # --- Capacity Logic ---
        actual_output_parts = prod_df['working_cavities'].sum()
        target_output_parts = optimal_output_parts * (self.target_output_perc / 100.0)

        true_loss_parts = optimal_output_parts - actual_output_parts
        
        # Initialize variables
        capacity_gain_fast_parts = 0.0
        capacity_loss_slow_parts = 0.0
        capacity_gain_fast_sec = 0.0
        capacity_loss_slow_sec = 0.0

        if not prod_df.empty:
            # Inefficiency Calculation
            prod_df['parts_delta'] = ((prod_df['approved_ct_for_run'] - prod_df['actual_ct']) / prod_df['approved_ct_for_run']) * prod_df['working_cavities']
            
            capacity_gain_fast_parts = prod_df.loc[prod_df['parts_delta'] > 0, 'parts_delta'].sum()
            capacity_loss_slow_parts = abs(prod_df.loc[prod_df['parts_delta'] < 0, 'parts_delta'].sum())
            
            # Inefficiency Time Calculation
            prod_df['time_delta'] = prod_df['approved_ct_for_run'] - prod_df['actual_ct']
            capacity_gain_fast_sec = prod_df.loc[prod_df['time_delta'] > 0, 'time_delta'].sum()
            capacity_loss_slow_sec = abs(prod_df.loc[prod_df['time_delta'] < 0, 'time_delta'].sum())

        net_cycle_loss_parts = capacity_loss_slow_parts - capacity_gain_fast_parts
        capacity_loss_downtime_parts = true_loss_parts - net_cycle_loss_parts
        
        net_cycle_loss_sec = capacity_loss_slow_sec - capacity_gain_fast_sec
        
        # Enhanced to use time-exact loss instead of multiplying by average CT
        total_capacity_loss_sec = downtime_sec + capacity_loss_slow_sec
        
        gap_to_target_parts = actual_output_parts - target_output_parts
        capacity_loss_vs_target_parts = max(0, -gap_to_target_parts)
        
        total_shots = len(df)
        stop_count_shots = df['stop_flag'].sum()
        normal_shots = total_shots - stop_count_shots
        
        run_rate_efficiency = (normal_shots / total_shots * 100) if total_shots > 0 else 0
        capacity_efficiency = (actual_output_parts / optimal_output_parts) if optimal_output_parts > 0 else 0

        # Shot Typing
        epsilon = 0.001
        conditions = [
            df['stop_flag'] == 1,
            df['actual_ct'] > (df['approved_ct_for_run'] + epsilon), 
            df['actual_ct'] < (df['approved_ct_for_run'] - epsilon)
        ]
        choices = ['Downtime (Stop)', 'Slow Cycle', 'Fast Cycle']
        df['shot_type'] = np.select(conditions, choices, default='On Target')

        return {
            "processed_df": df,
            "total_runtime_sec": total_runtime_sec,
            "production_time_sec": production_time_sec,
            "downtime_sec": downtime_sec,
            "mttr_min": mttr_min,
            "stability_index": stability_index,
            "stops": stops,
            "optimal_output_parts": optimal_output_parts,
            "actual_output_parts": actual_output_parts,
            "target_output_parts": target_output_parts,
            "capacity_loss_downtime_parts": capacity_loss_downtime_parts,
            "capacity_loss_slow_parts": capacity_loss_slow_parts,
            "capacity_gain_fast_parts": capacity_gain_fast_parts,
            "total_capacity_loss_parts": true_loss_parts,
            "total_capacity_loss_sec": total_capacity_loss_sec,
            "gap_to_target_parts": gap_to_target_parts,
            "capacity_loss_vs_target_parts": capacity_loss_vs_target_parts,
            "efficiency_rate": run_rate_efficiency,
            "capacity_efficiency": capacity_efficiency,
            "total_shots": total_shots,
            "normal_shots": normal_shots,
            "mtbf_min": (production_time_sec / 60 / stops) if stops > 0 else (production_time_sec / 60)
        }

def calculate_run_summaries(df_period, config):
    """Calculates per-run metrics for the breakdown table."""
    summary_list = []
    if 'run_id' not in df_period.columns: return pd.DataFrame()
    for r_id, df_run in df_period.groupby('run_id'):
        if df_run.empty: continue
        calc = CapacityRiskCalculator(df_run, **config)
        res = calc.results
        summary_list.append({
            'run_id': r_id,
            'tool_ids': ', '.join(df_run['tool_id'].astype(str).unique()) if 'tool_id' in df_run.columns else 'Unknown',
            'start_time': df_run['shot_time'].min(),
            'end_time': df_run['shot_time'].max(),
            'total_shots': res['total_shots'],
            'normal_shots': res['normal_shots'],
            'stop_events': res['stops'],
            'stopped_shots': res['total_shots'] - res['normal_shots'],
            'mode_ct': df_run['mode_ct'].iloc[0] if 'mode_ct' in df_run else 0,
            'mode_lower': df_run['mode_lower'].iloc[0] if 'mode_lower' in df_run else 0,
            'mode_upper': df_run['mode_upper'].iloc[0] if 'mode_upper' in df_run else 0,
            'total_runtime_sec': res['total_runtime_sec'],
            'production_time_sec': res['production_time_sec'],
            'downtime_sec': res['downtime_sec'],
            'total_capacity_loss_sec': res['total_capacity_loss_sec'],
            'optimal_output_parts': res['optimal_output_parts'],
            'target_output_parts': res['target_output_parts'],
            'actual_output_parts': res['actual_output_parts'],
            'capacity_loss_downtime_parts': res['capacity_loss_downtime_parts'],
            'capacity_loss_slow_parts': res['capacity_loss_slow_parts'],
            'capacity_gain_fast_parts': res['capacity_gain_fast_parts'],
            'total_capacity_loss_parts': res['total_capacity_loss_parts'],
            'mttr_min': res['mttr_min'], 
            'stability_index': res['stability_index'] 
        })
    df_summary = pd.DataFrame(summary_list).sort_values('start_time')
    df_summary['display_run_id'] = range(1, len(df_summary) + 1)
    return df_summary

# ==============================================================================
# --- AGGREGATION, PREDICTION & RISK LOGIC ---
# ==============================================================================

def get_aggregated_data(df, freq_mode, config):
    """Generates aggregated dataframe for tables/charts."""
    rows = []
    
    if freq_mode == 'Daily': grouper = df.groupby(df['shot_time'].dt.date)
    elif freq_mode == 'Weekly': grouper = df.groupby(df['shot_time'].dt.to_period('W').astype(str))
    elif freq_mode == 'Monthly': grouper = df.groupby(df['shot_time'].dt.to_period('M').astype(str))
    elif freq_mode == 'Hourly': grouper = df.groupby(df['shot_time'].dt.floor('H'))
    elif freq_mode == 'by Run': 
        temp_calc = CapacityRiskCalculator(df, **config)
        grouper = temp_calc.results['processed_df'].groupby('run_id')
    else: return pd.DataFrame()

    for group_name, df_subset in grouper:
        calc = CapacityRiskCalculator(df_subset, **config)
        res = calc.results
        if not res: continue
        
        period_label = group_name
        if freq_mode == 'by Run':
            try:
                period_label = f"Run {int(group_name) + 1}"
            except (ValueError, TypeError):
                pass

        rows.append({
            'Period': period_label,
            'Actual Output': res['actual_output_parts'],
            'Optimal Output': res['optimal_output_parts'],
            'Target Output': res['target_output_parts'],
            'Downtime Loss': res['capacity_loss_downtime_parts'],
            'Slow Loss': res['capacity_loss_slow_parts'],
            'Fast Gain': res['capacity_gain_fast_parts'],
            'Net Cycle Loss': res['capacity_loss_slow_parts'] - res['capacity_gain_fast_parts'],
            'Total Loss': res['total_capacity_loss_parts'],
            'Gap to Target': res['gap_to_target_parts'],
            'Run Time': format_seconds_to_dhm(res['total_runtime_sec']),
            'Downtime': format_seconds_to_dhm(res['downtime_sec']),
            
            # --- Added for detailed Tables ---
            'Run Time Sec': res['total_runtime_sec'],
            'Production Time Sec': res['production_time_sec'],
            'Downtime Sec': res['downtime_sec'],
            'Total Shots': res['total_shots'],
            'Normal Shots': res['normal_shots'],
            'Downtime Shots': res['total_shots'] - res['normal_shots']
        })
        
    return pd.DataFrame(rows)

def generate_po_periodic_data(df_bar_view, po_record, freq_mode, config, working_days_per_week, working_hours_per_day):
    """Generates periodic aggregated data spanning the full PO timeline."""
    start_date = po_record.get('start_date')
    due_date = po_record.get('due_date')
    if pd.isna(start_date) or pd.isna(due_date): return pd.DataFrame()
    
    start_date = pd.to_datetime(start_date)
    due_date = pd.to_datetime(due_date)
    total_qty = po_record.get('total_qty', 0)
    
    # Calculate uniform demand spread
    total_calendar_days = (due_date - start_date).days
    if total_calendar_days <= 0: total_calendar_days = 1
    
    total_weeks = total_calendar_days / 7.0
    total_working_days = total_weeks * working_days_per_week
    
    daily_demand = total_qty / total_working_days if total_working_days > 0 else 0
    weekly_demand = daily_demand * working_days_per_week
    monthly_demand = weekly_demand * 4.33
    
    # Calculate Configured Max Capacity based on optimal cycle times
    avg_ct = df_bar_view['approved_ct'].mean() if (not df_bar_view.empty and 'approved_ct' in df_bar_view.columns) else 1
    if pd.isna(avg_ct) or avg_ct <= 0: avg_ct = 1
    cav = df_bar_view['working_cavities'].max() if (not df_bar_view.empty and 'working_cavities' in df_bar_view.columns) else 1
    
    hourly_cap = (3600 / avg_ct) * cav
    
    # Process Actuals Data
    agg_df = get_aggregated_data(df_bar_view, freq_mode, config)
    
    # Calculate timeline bounds to match the Burn-Up chart scope (Start -> Due or Late Finish)
    current_cum = agg_df['Actual Output'].sum() if not agg_df.empty else 0
    last_actual_date = pd.to_datetime(df_bar_view['shot_time'].max()).date() if not df_bar_view.empty else start_date.date()
    
    days_elapsed = (last_actual_date - start_date.date()).days + 1
    if days_elapsed <= 0: days_elapsed = 1
    avg_daily_rate = current_cum / days_elapsed
    
    remaining_qty = total_qty - current_cum
    max_proj_days = 0
    if remaining_qty > 0 and avg_daily_rate > 0:
        max_proj_days = int(remaining_qty / avg_daily_rate) + 1
        max_proj_days = min(max_proj_days, 365) # cap prediction limits
        
    projected_end_date = last_actual_date + timedelta(days=max_proj_days)
    end_timeline_date = max(due_date.date(), projected_end_date)
    
    # Generate continuous empty timeline
    if freq_mode == 'Daily':
        full_periods = pd.date_range(start=start_date.date(), end=end_timeline_date, freq='D').date
        df_full = pd.DataFrame({'Period': full_periods})
        df_full['Estimated Demand'] = daily_demand
        df_full['Configured Max Capacity'] = hourly_cap * working_hours_per_day
        if not agg_df.empty: agg_df['Period'] = pd.to_datetime(agg_df['Period']).dt.date
    elif freq_mode == 'Weekly':
        full_periods = pd.period_range(start=start_date, end=end_timeline_date, freq='W').astype(str)
        df_full = pd.DataFrame({'Period': full_periods})
        df_full['Estimated Demand'] = weekly_demand
        df_full['Configured Max Capacity'] = hourly_cap * working_hours_per_day * working_days_per_week
    elif freq_mode == 'Monthly':
        full_periods = pd.period_range(start=start_date, end=end_timeline_date, freq='M').astype(str)
        df_full = pd.DataFrame({'Period': full_periods})
        df_full['Estimated Demand'] = monthly_demand
        df_full['Configured Max Capacity'] = hourly_cap * working_hours_per_day * working_days_per_week * 4.33
    else:
        df_full = pd.DataFrame()

    # Merge full timeline with available actuals, filling gaps with zero
    if not df_full.empty:
        df_full['Period'] = df_full['Period'].astype(str)
        if not agg_df.empty:
            agg_df['Period'] = agg_df['Period'].astype(str)
            final_df = pd.merge(df_full, agg_df[['Period', 'Actual Output']], on='Period', how='left')
            final_df['Actual Output'] = final_df['Actual Output'].fillna(0)
        else:
            final_df = df_full
            final_df['Actual Output'] = 0
        return final_df
        
    return agg_df

def generate_po_prediction_data(df_po_shots, po_record, config):
    """Generates time-series data specifically for PO Burn-up charting."""
    if pd.isna(po_record.get('start_date')) or pd.isna(po_record.get('due_date')):
        return None
        
    start_date = po_record['start_date'].date() if isinstance(po_record['start_date'], pd.Timestamp) else pd.to_datetime(po_record['start_date']).date()
    due_date = po_record['due_date'].date() if isinstance(po_record['due_date'], pd.Timestamp) else pd.to_datetime(po_record['due_date']).date()
    total_qty = po_record.get('total_qty', 0)

    # Target Burnup Line (Ideal linear progress)
    total_days = (due_date - start_date).days
    if total_days <= 0: total_days = 1
    
    target_dates = [start_date + timedelta(days=i) for i in range(total_days + 1)]
    target_vals = [(total_qty / total_days) * i for i in range(total_days + 1)]

    # Get daily aggregations for actuals
    agg_daily = get_aggregated_data(df_po_shots, 'Daily', config) if not df_po_shots.empty else pd.DataFrame()
    
    if agg_daily.empty:
        return {
            'target_dates': target_dates, 'target_vals': target_vals,
            'actual_dates': [], 'actual_cum': [],
            'forecast_dates': [], 'forecast_avg': [], 'forecast_opt': [],
            'due_date': due_date, 'start_date': start_date, 'total_qty': total_qty,
            'current_cum': 0, 'avg_daily_rate': 0, 'opt_daily_rate': 0
        }
        
    agg_daily['Period'] = pd.to_datetime(agg_daily['Period']).dt.date
    agg_daily = agg_daily.sort_values('Period')
    agg_daily['Cumulative Actual'] = agg_daily['Actual Output'].cumsum()
    
    last_actual_date = agg_daily['Period'].max()
    current_cum = agg_daily['Cumulative Actual'].max()
    
    # Calculate optimal rate from the whole subset
    calc = CapacityRiskCalculator(df_po_shots, **config)
    res = calc.results
    
    days_elapsed = (last_actual_date - start_date).days + 1
    if days_elapsed <= 0: days_elapsed = 1
    
    avg_daily_rate = current_cum / days_elapsed
    opt_daily_rate = res['optimal_output_parts'] / days_elapsed if res else 0
    
    remaining_qty = total_qty - current_cum
    max_proj_days = 0
    
    if remaining_qty > 0:
        days_to_finish_avg = int(remaining_qty / avg_daily_rate) + 1 if avg_daily_rate > 0 else 30
        days_to_finish_opt = int(remaining_qty / opt_daily_rate) + 1 if opt_daily_rate > 0 else 30
            
        max_proj_days = max(days_to_finish_avg, days_to_finish_opt, (due_date - last_actual_date).days)
        max_proj_days = min(max_proj_days, 365) # cap projection at 1 year max
    else:
        max_proj_days = max(0, (due_date - last_actual_date).days)
    
    forecast_dates = [last_actual_date + timedelta(days=i) for i in range(max_proj_days + 1)]
    forecast_avg = [current_cum + (avg_daily_rate * i) for i in range(max_proj_days + 1)]
    forecast_opt = [current_cum + (opt_daily_rate * i) for i in range(max_proj_days + 1)]

    return {
        'target_dates': target_dates, 'target_vals': target_vals,
        'actual_dates': agg_daily['Period'].tolist(), 'actual_cum': agg_daily['Cumulative Actual'].tolist(),
        'forecast_dates': forecast_dates, 'forecast_avg': forecast_avg, 'forecast_opt': forecast_opt,
        'due_date': due_date, 'start_date': start_date, 'total_qty': total_qty,
        'current_cum': current_cum, 'avg_daily_rate': avg_daily_rate, 'opt_daily_rate': opt_daily_rate
    }

def generate_prediction_data(df_daily_agg, start_date, target_date, demand_target_total=None):
    """Fallback projection chart data generation."""
    if df_daily_agg.empty: return None

    df = df_daily_agg.copy()
    df['Period'] = pd.to_datetime(df['Period'])
    df = df.sort_values('Period')
    
    df['Cumulative Actual'] = df['Actual Output'].cumsum()
    
    last_historic_ts = df['Period'].max()
    last_historic_date = last_historic_ts.date() if hasattr(last_historic_ts, 'date') else last_historic_ts
    current_cumulative = df['Cumulative Actual'].max()
    
    days_with_data = (last_historic_date - df['Period'].min().date()).days + 1
    if days_with_data < 1: days_with_data = 1
    
    avg_daily_rate = df['Actual Output'].sum() / days_with_data
    peak_daily_rate = df['Actual Output'].quantile(0.90) if len(df) > 5 else df['Actual Output'].max()

    if isinstance(target_date, datetime):
        target_date = target_date.date()
        
    projection_days = (target_date - last_historic_date).days
    if projection_days < 1: projection_days = 0
    
    future_dates = [last_historic_date + timedelta(days=i) for i in range(projection_days + 1)]
    
    proj_avg = [current_cumulative + (avg_daily_rate * i) for i in range(len(future_dates))]
    proj_peak = [current_cumulative + (peak_daily_rate * i) for i in range(len(future_dates))]
    
    req_rate = 0
    proj_req = []
    if demand_target_total:
        remaining_qty = max(0, demand_target_total - current_cumulative)
        if projection_days > 0:
            req_rate = remaining_qty / projection_days
            proj_req = [current_cumulative + (req_rate * i) for i in range(len(future_dates))]
    
    return {
        'historic_dates': df['Period'],
        'historic_cum': df['Cumulative Actual'],
        'future_dates': future_dates,
        'proj_avg': proj_avg,
        'proj_peak': proj_peak,
        'proj_req': proj_req,
        'rates': {'avg': avg_daily_rate, 'peak': peak_daily_rate, 'req': req_rate}
    }

def calculate_capacity_risk_scores(df_all, config):
    risk_data = []
    for tool_id, df_tool in df_all.groupby('tool_id'):
        max_date = df_tool['shot_time'].max()
        cutoff_date = max_date - timedelta(weeks=4)
        df_period = df_tool[df_tool['shot_time'] >= cutoff_date].copy()
        
        if df_period.empty: continue
        
        calc = CapacityRiskCalculator(df_period, **config)
        res = calc.results
        if res['target_output_parts'] == 0: continue
        
        ach_perc = (res['actual_output_parts'] / res['target_output_parts']) * 100
        
        midpoint = cutoff_date + (max_date - cutoff_date) / 2
        df_late = df_period[df_period['shot_time'] >= midpoint]
        df_early = df_period[df_period['shot_time'] < midpoint]
        
        trend = "Stable"
        if not df_early.empty and not df_late.empty:
            c_early = CapacityRiskCalculator(df_early, **config).results
            c_late = CapacityRiskCalculator(df_late, **config).results
            early_rate = c_early['actual_output_parts'] / (c_early['total_runtime_sec']/3600) if c_early['total_runtime_sec'] > 0 else 0
            late_rate = c_late['actual_output_parts'] / (c_late['total_runtime_sec']/3600) if c_late['total_runtime_sec'] > 0 else 0
            
            if late_rate < early_rate * 0.95: trend = "Declining"
            elif late_rate > early_rate * 1.05: trend = "Improving"

        base_score = min(ach_perc, 100)
        if trend == "Declining": base_score -= 20
        
        risk_data.append({
            'Tool ID': tool_id,
            'Risk Score': max(0, base_score),
            'Achievement %': ach_perc,
            'Trend': trend,
            'Gap': res['gap_to_target_parts']
        })
    return pd.DataFrame(risk_data).sort_values('Risk Score')

# ==============================================================================
# --- NEW: AUTOMATED INSIGHTS & EXPORT ---
# ==============================================================================

def generate_capacity_insights(res, benchmark_mode):
    """Generates natural language summary of the capacity loss."""
    if not res: return {"overall": "No data available."}
    
    act = res['actual_output_parts']
    tgt = res['target_output_parts'] if benchmark_mode == "Target" else res['optimal_output_parts']
    diff = act - tgt
    
    status = "exceeded" if diff >= 0 else "missed"
    perc = abs(diff / tgt * 100) if tgt > 0 else 0
    
    color = "green" if diff >= 0 else "red"
    overall = f"Production <strong style='color:{color}'>{status}</strong> the {benchmark_mode} goal by <strong>{perc:.1f}%</strong> ({abs(diff):,.0f} parts)."
    
    # Driver Analysis
    drivers = []
    loss_dt = res['capacity_loss_downtime_parts']
    net_slow = res['capacity_loss_slow_parts'] - res['capacity_gain_fast_parts']
    
    total_loss_absolute = loss_dt + max(0, net_slow)
    
    if total_loss_absolute > 0:
        dt_share = (loss_dt / total_loss_absolute) * 100
        slow_share = (max(0, net_slow) / total_loss_absolute) * 100
        
        loss_term = "uncaptured capacity" if status == "exceeded" else "loss"
        driver_intro = "The primary constraint was" if status == "exceeded" else "The primary driver was"
        
        if dt_share > 60:
            drivers.append(f"{driver_intro} <strong>Downtime</strong>, accounting for <strong>{dt_share:.0f}%</strong> of the {loss_term}.")
            rec = "Focus on reducing Stop Events (MTBF) and improving Reaction Time (MTTR)."
        elif slow_share > 60:
            drivers.append(f"{driver_intro} <strong>Slow Cycle Time</strong>, accounting for <strong>{slow_share:.0f}%</strong> of the {loss_term}.")
            rec = "Investigate process parameters causing the machine to run slower than the Approved Cycle Time."
        else:
            drivers.append(f"Constraints were split between <strong>Downtime ({dt_share:.0f}%)</strong> and <strong>Slow Cycles ({slow_share:.0f}%)</strong>.")
            rec = "A balanced approach addressing both uptime and cycle speed is required."
    else:
        drivers.append("No significant capacity losses detected.")
        rec = "Maintain current performance standards."

    if res['capacity_gain_fast_parts'] > (res['actual_output_parts'] * 0.05):
        drivers.append(f"Note: Running faster than standard gained <strong>{res['capacity_gain_fast_parts']:,.0f}</strong> bonus parts.")

    return {"overall": overall, "drivers": " ".join(drivers), "recommendation": rec}

def generate_forecast_insights(pred_data, demand_target):
    """Generates insights for the forecast tab."""
    if not pred_data or demand_target <= 0:
        return "Please set a Demand Goal to generate a completion forecast."
    
    current_cum = pred_data['historic_cum'].iloc[-1]
    if current_cum >= demand_target:
        return f"🎉 <strong>Goal Achieved!</strong> Current output ({current_cum:,.0f}) has already exceeded the demand target ({demand_target:,.0f})."

    remaining = demand_target - current_cum
    rates = pred_data['rates']
    avg_rate = rates['avg']
    peak_rate = rates['peak']
    start_date = pred_data['future_dates'][0]

    def get_finish_date(rate):
        if rate <= 0: return None
        days = remaining / rate
        return start_date + timedelta(days=int(days))

    date_avg = get_finish_date(avg_rate)
    date_peak = get_finish_date(peak_rate)

    insight_html = f"""
    <ul style='margin-bottom:0;'>
        <li>To meet demand of <strong>{demand_target:,.0f}</strong>, you need <strong>{remaining:,.0f}</strong> more parts.</li>
    """

    if date_avg:
        insight_html += f"<li>At your <strong>Current Average Rate</strong> ({avg_rate:,.0f} parts/day), you will meet demand on <strong>{date_avg.strftime('%Y-%m-%d')}</strong>.</li>"
    else:
        insight_html += f"<li>At your <strong>Current Average Rate</strong>, you are not projected to meet demand (rate is 0 or negative).</li>"

    if date_peak:
        insight_html += f"<li>At <strong>Optimal/Peak Performance</strong> ({peak_rate:,.0f} parts/day), you could meet demand as early as <strong>{date_peak.strftime('%Y-%m-%d')}</strong>.</li>"
    
    insight_html += "</ul>"
    return insight_html

def prepare_and_generate_capacity_excel(df_view, config):
    """Generates the Excel export with formatted sheets."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        
        fmt_header = wb.add_format({'bold':True,'bg_color':'#002060','font_color':'white','border':1})
        fmt_num = wb.add_format({'num_format':'#,##0','border':1})
        fmt_dec = wb.add_format({'num_format':'0.00','border':1})
        
        ws_sum = wb.add_worksheet("Management Summary")
        run_data = calculate_run_summaries(df_view, config)
        
        if not run_data.empty:
            ws_sum.write_row('A1', ['Start Time', 'End Time', 'Optimal', 'Actual', 'Loss (Downtime)', 'Loss (Slow)', 'Gain (Fast)'], fmt_header)
            
            for i, row in run_data.iterrows():
                ws_sum.write(i+1, 0, str(row['start_time']))
                ws_sum.write(i+1, 1, str(row['end_time']))
                ws_sum.write(i+1, 2, row['optimal_output_parts'], fmt_num)
                ws_sum.write(i+1, 3, row['actual_output_parts'], fmt_num)
                ws_sum.write(i+1, 4, row['capacity_loss_downtime_parts'], fmt_num)
                ws_sum.write(i+1, 5, row['capacity_loss_slow_parts'], fmt_num)
                ws_sum.write(i+1, 6, row['capacity_gain_fast_parts'], fmt_num)

        df_view.to_excel(writer, sheet_name="Raw Data", index=False)
        
    return output.getvalue()

def generate_mttr_mtbf_analysis(analysis_df):
    """Generates correlation text analysis for MTTR/MTBF drivers."""
    if analysis_df is None or analysis_df.empty or 'stop_events' not in analysis_df.columns:
        return "Not enough data to generate detailed correlation analysis."
        
    analysis_df_clean = analysis_df.dropna(subset=['stop_events', 'stability_index', 'mttr_min'])
    
    if analysis_df_clean.empty or len(analysis_df_clean) < 2:
        return "Insufficient data: At least 2 production runs with stop events are required to perform correlation analysis."

    period_col = 'display_run_id' if 'display_run_id' in analysis_df_clean.columns else 'run_id'
    
    df_calc = analysis_df_clean.rename(columns={
        'stop_events': 'stops', 
        'stability_index': 'stability', 
        'mttr_min': 'mttr',
        period_col: 'period'
    })
    
    stops_stability_corr = df_calc['stops'].corr(df_calc['stability'])
    mttr_stability_corr = df_calc['mttr'].corr(df_calc['stability'])
    
    corr_insight = ""
    primary_driver_is_frequency = False
    primary_driver_is_duration = False
    
    if not pd.isna(stops_stability_corr) and not pd.isna(mttr_stability_corr):
        if abs(stops_stability_corr) > abs(mttr_stability_corr) * 1.5:
            primary_driver = "the **frequency of stops**"
            primary_driver_is_frequency = True
        elif abs(mttr_stability_corr) > abs(stops_stability_corr) * 1.5:
            primary_driver = "the **duration of stops**"
            primary_driver_is_duration = True
        else:
            primary_driver = "both the **frequency and duration of stops**"
        corr_insight = (f"This analysis suggests that <strong>{primary_driver}</strong> has the strongest impact on overall stability.")
    
    example_insight = ""
    if primary_driver_is_frequency:
        highest_stops_row = df_calc.loc[df_calc['stops'].idxmax()]
        example_insight = (f"For example, Run {highest_stops_row['period']} recorded the most interruptions (<strong>{int(highest_stops_row['stops'])} stops</strong>). Prioritizing the root cause of these frequent events is recommended.")
    elif primary_driver_is_duration:
        highest_mttr_row = df_calc.loc[df_calc['mttr'].idxmax()]
        example_insight = (f"Run {highest_mttr_row['period']} experienced prolonged downtimes with an average repair time of <strong>{highest_mttr_row['mttr']:.1f} minutes</strong>. Investigating the cause of these prolonged stops is the top priority.")
    else:
        if not df_calc['mttr'].empty:
            highest_mttr_row = df_calc.loc[df_calc['mttr'].idxmax()]
            example_insight = (f"As an example, Run {highest_mttr_row['period']} experienced prolonged downtimes with an average repair time of <strong>{highest_mttr_row['mttr']:.1f} minutes</strong>, highlighting the impact of long stops.")
            
    return f"<div style='line-height: 1.6;'><p>{corr_insight}</p><p>{example_insight}</p></div>"

# ==============================================================================
# --- PLOTTING FUNCTIONS ---
# ==============================================================================

def plot_po_periodic_chart(agg_po, bar_freq):
    """Plots the periodic bar chart for PO tracking vs Demand & Configured Capacity."""
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=agg_po['Period'], y=agg_po['Actual Output'], 
        name='Actual Output', marker_color=PASTEL_COLORS['blue']
    ))
    
    fig.add_trace(go.Scatter(
        x=agg_po['Period'], y=agg_po['Configured Max Capacity'], 
        name='Configured Max Capacity', mode='lines+markers', 
        line=dict(color=PASTEL_COLORS['green'], dash='dot', width=2)
    ))
    
    fig.add_trace(go.Scatter(
        x=agg_po['Period'], y=agg_po['Estimated Demand'], 
        name='Estimated PO Demand', mode='lines+markers', 
        line=dict(color=PASTEL_COLORS['red'], dash='dash', width=2)
    ))
    
    fig.update_layout(
        title=f"Periodic Production vs Demand ({bar_freq})",
        barmode='group', hovermode="x unified", height=450,
        yaxis_title="Parts Output", xaxis_title="Period"
    )
    return fig

def plot_po_burnup(pred_data):
    """Plots the PO specific Burn-Up tracking chart."""
    if not pred_data: return go.Figure()
    fig = go.Figure()
    
    # Target Burnup (Grey Dashed)
    if pred_data['target_dates']:
        fig.add_trace(go.Scatter(x=pred_data['target_dates'], y=pred_data['target_vals'], 
                                 mode='lines', name='PO Target Burn-up', line=dict(color='grey', dash='dash')))
                             
    # Actual Cumulative (Blue Line)
    if pred_data['actual_dates']:
        fig.add_trace(go.Scatter(x=pred_data['actual_dates'], y=pred_data['actual_cum'], 
                                 mode='lines+markers', name='Actual Accumulated', line=dict(color=PASTEL_COLORS['blue'], width=3)))
                             
    # Forecast Avg (Orange Dot)
    if pred_data['avg_daily_rate'] > 0 and pred_data['forecast_dates']:
        fig.add_trace(go.Scatter(x=pred_data['forecast_dates'], y=pred_data['forecast_avg'], 
                                 mode='lines', name=f"Forecast (Avg: {pred_data['avg_daily_rate']:.0f}/d)", line=dict(color=PASTEL_COLORS['orange'], dash='dot')))
                             
    # Forecast Opt (Green Dot)
    if pred_data['opt_daily_rate'] > 0 and pred_data['forecast_dates']:
        fig.add_trace(go.Scatter(x=pred_data['forecast_dates'], y=pred_data['forecast_opt'], 
                                 mode='lines', name=f"Forecast (Opt: {pred_data['opt_daily_rate']:.0f}/d)", line=dict(color=PASTEL_COLORS['green'], dash='dot')))
                             
    # Annotations - Fix applied using Unix Timestamp mapping for Plotly compatibility
    due_ts = pd.to_datetime(pred_data['due_date']).timestamp() * 1000
    fig.add_vline(x=due_ts, line_width=2, line_dash="dash", line_color="red", annotation_text="PO Due Date")
    fig.add_hline(y=pred_data['total_qty'], line_width=2, line_dash="solid", line_color="purple", annotation_text="PO Total Qty")
    
    # Force the X-axis bounds to represent the exact same context duration as the periodic chart
    start_dt = pd.to_datetime(pred_data['start_date'])
    max_dt_target = pd.to_datetime(pred_data['target_dates'][-1]) if pred_data['target_dates'] else start_dt
    max_dt_forecast = pd.to_datetime(pred_data['forecast_dates'][-1]) if pred_data['forecast_dates'] else max_dt_target
    end_dt = max(max_dt_target, max_dt_forecast)

    fig.update_layout(
        title="PO Target Burn-up vs Reality", 
        hovermode="x unified", 
        height=500, 
        yaxis_title="Accumulated Parts Output", 
        xaxis_title="Date",
        xaxis_range=[start_dt, end_dt] 
    )
    return fig

def create_time_breakdown_donut(total_sec, prod_sec, down_sec):
    c_prod = PASTEL_COLORS['green']
    c_down = PASTEL_COLORS['red']
    
    center_text = f"<span style='font-size:18px; color:#cccccc;'>Total Run Duration</span><br><br><span style='font-size:32px; font-weight:bold; color:white; line-height:1.2'>{format_seconds_to_dhm(total_sec)}</span>"
    
    fig = go.Figure(data=[go.Pie(
        values=[prod_sec, down_sec],
        labels=['Production Time', 'Run Rate Downtime'],
        marker=dict(colors=[c_prod, c_down]),
        hole=0.7, 
        sort=False,
        direction='clockwise',
        textinfo='none',
        hoverinfo='label+percent+value'
    )])
    
    fig.update_layout(
        annotations=[dict(text=center_text, x=0.5, y=0.5, font_size=16, showarrow=False)],
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5, font=dict(size=14)),
        margin=dict(t=30, b=30, l=20, r=20),
        height=320,
        title=dict(text="Total Run Time Breakdown", x=0, font=dict(size=18)),
        paper_bgcolor='rgba(0,0,0,0)', 
        plot_bgcolor='rgba(0,0,0,0)'
    )
    
    fig.update_traces(textinfo='label+percent', textposition='outside', textfont=dict(size=14))
    return fig

def create_modern_gauge(value, title):
    color = PASTEL_COLORS['green']
    if value <= 50: color = PASTEL_COLORS['red']
    elif value <= 70: color = PASTEL_COLORS['orange']
    
    plot_value = max(0, min(value, 100))
    remainder = 100 - plot_value
    visible_total = 100 
    
    values = [plot_value, remainder, visible_total]
    colors = [color, '#41424C', 'rgba(255, 255, 255, 0)']
    
    fig = go.Figure(data=[go.Pie(
        values=values,
        hole=0.65,
        sort=False,
        direction='clockwise',
        rotation=-90, 
        textinfo='none',
        marker=dict(colors=colors), 
        hoverinfo='none'
    )])

    fig.add_annotation(
        text=f"{value:.1f}%",
        x=0.5, y=0.15,
        font=dict(size=48, weight='bold', color='white', family="Arial"),
        showarrow=False
    )
    
    fig.update_layout(
        title=dict(text=title, x=0, xanchor='left', y=0.9, font=dict(size=20)),
        margin=dict(l=20, r=20, t=40, b=0),
        height=220, 
        showlegend=False,
        paper_bgcolor='rgba(0,0,0,0)', 
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig

def create_stability_driver_bar(mtbf, mttr, stability_index):
    total = mtbf + mttr
    if total == 0: return go.Figure()
    
    mtbf_pct = (mtbf / total) * 100
    mttr_pct = 100 - mtbf_pct
    
    downtime_pct = 100 - stability_index
    label_mtbf = f"MTBF: {mtbf:.1f}m ({mtbf_pct:.1f}%)"
    label_mttr = f"MTTR: {mttr:.1f}m ({mttr_pct:.1f}%)"

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=['Cycle'], x=[mtbf], name=label_mtbf, orientation='h',
        marker_color=PASTEL_COLORS['blue'],
        hoverinfo='name' 
    ))
    fig.add_trace(go.Bar(
        y=['Cycle'], x=[mttr], name=label_mttr, orientation='h',
        marker_color=PASTEL_COLORS['red'],
        hoverinfo='name'
    ))
    
    footnote_text = f"Stability Index: {stability_index:.1f}% Stable Production Time vs. {downtime_pct:.1f}% Run Rate Downtime"

    fig.update_layout(
        barmode='stack',
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        margin=dict(t=60, b=120, l=10, r=10), 
        height=260, 
        title=dict(text="MTTR & MTBF Analysis", x=0, font=dict(size=24)), 
        showlegend=True,
        legend=dict(
            orientation="h", 
            yanchor="top", 
            y=-0.1, 
            xanchor="center", 
            x=0.5,
            font=dict(size=18), 
            bgcolor='rgba(0,0,0,0)'
        ),
        paper_bgcolor='rgba(0,0,0,0)', 
        plot_bgcolor='rgba(0,0,0,0)',
        annotations=[
            dict(x=0, y=-0.7, text=footnote_text, showarrow=False, xref='paper', yref='paper', xanchor='left', yanchor='top', font=dict(size=16, color="#cccccc"))
        ]
    )
    return fig

def create_donut_chart(value, title, color_scheme='blue'):
    if color_scheme == 'blue': main_color = PASTEL_COLORS['blue']
    elif color_scheme == 'green': main_color = PASTEL_COLORS['green']
    elif color_scheme == 'dynamic':
        if value < 70: main_color = PASTEL_COLORS['red']
        elif value < 90: main_color = PASTEL_COLORS['orange']
        else: main_color = PASTEL_COLORS['green']
    else: main_color = color_scheme

    plot_val = min(value, 100)
    remainder = 100 - plot_val
    
    fig = go.Figure(data=[go.Pie(
        values=[plot_val, remainder], hole=0.75, sort=False, direction='clockwise',
        textinfo='none', marker=dict(colors=[main_color, '#e6e6e6']), hoverinfo='none'
    )])

    fig.add_annotation(text=f"{value:.1f}%", x=0.5, y=0.5, font=dict(size=24, weight='bold', color=main_color), showarrow=False)
    
    fig.update_layout(
        title=dict(text=title, x=0.5, xanchor='center', y=0.95, font=dict(size=14)),
        margin=dict(l=20, r=20, t=30, b=20), height=180, showlegend=False,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig

def plot_waterfall(metrics, benchmark_mode="Optimal"):
    total_opt = metrics['optimal_output_parts']
    actual = metrics['actual_output_parts']
    
    loss_dt = -metrics['capacity_loss_downtime_parts']
    loss_slow = -metrics['capacity_loss_slow_parts']
    gain_fast = metrics['capacity_gain_fast_parts']
    
    measure = ["absolute", "relative", "relative", "relative", "total"]
    x_label = ["Optimal (Theory)", "Loss: Downtime", "Loss: Speed", "Gain: Speed", "Actual Output"]
    y_val = [total_opt, loss_dt, loss_slow, gain_fast, actual]
    text_val = [f"{total_opt:,.0f}", f"{loss_dt:,.0f}", f"{loss_slow:,.0f}", f"+{gain_fast:,.0f}", f"{actual:,.0f}"]
    
    fig = go.Figure(go.Waterfall(
        name="Capacity Bridge",
        orientation="v",
        measure=measure,
        x=x_label,
        y=y_val,
        text=text_val,
        textposition="outside",
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        decreasing={"marker": {"color": PASTEL_COLORS['red']}},
        increasing={"marker": {"color": PASTEL_COLORS['green']}},
        totals={"marker": {"color": PASTEL_COLORS['blue']}}
    ))
    
    total_target = metrics['target_output_parts']
    if "Target" in str(benchmark_mode):
        fig.add_shape(type="line", x0=-0.5, x1=4.5, y0=total_target, y1=total_target,
                      line=dict(color=PASTEL_COLORS['target_line'], width=2, dash="dash"))
        fig.add_annotation(x=0, y=total_target, text=f"Target: {total_target:,.0f}", showarrow=False, yshift=10)

    fig.update_layout(
        title=f"Capacity Bridge: Where am I now? (vs {benchmark_mode})",
        showlegend=False, 
        height=450,
        yaxis_title="Parts Output"
    )
    return fig

def plot_prediction_chart(pred_data, demand_target_total=None):
    if not pred_data: return go.Figure()

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=pred_data['historic_dates'], 
        y=pred_data['historic_cum'],
        mode='lines+markers',
        name='Actual History',
        line=dict(color=PASTEL_COLORS['blue'], width=3)
    ))

    fig.add_trace(go.Scatter(
        x=pred_data['future_dates'], 
        y=pred_data['proj_avg'],
        mode='lines',
        name=f"Forecast ({pred_data['rates']['avg']:.0f}/day)",
        line=dict(color=PASTEL_COLORS['blue'], width=2, dash='dash')
    ))

    fig.add_trace(go.Scatter(
        x=pred_data['future_dates'], 
        y=pred_data['proj_peak'],
        mode='lines',
        name=f"Best Case ({pred_data['rates']['peak']:.0f}/day)",
        line=dict(color=PASTEL_COLORS['green'], width=1, dash='dot')
    ))
    
    if pred_data['proj_req']:
        fig.add_trace(go.Scatter(
            x=pred_data['future_dates'], 
            y=pred_data['proj_req'],
            mode='lines',
            name=f"Required ({pred_data['rates']['req']:.0f}/day)",
            line=dict(color=PASTEL_COLORS['orange'], width=2, dash='longdash')
        ))
    
    if demand_target_total:
         fig.add_hline(y=demand_target_total, line_dash="solid", line_color=PASTEL_COLORS['purple'], annotation_text=f"Total Demand: {demand_target_total:,.0f}")

    fig.update_layout(
        title="Future Capacity Projection: Where will I be?",
        xaxis_title="Date",
        yaxis_title="Cumulative Output (Parts)",
        hovermode="x unified",
        height=500
    )
    return fig

def plot_performance_breakdown(df_agg, x_col, benchmark_mode):
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_agg[x_col], y=df_agg['Actual Output'], name='Actual Output', marker_color=PASTEL_COLORS['blue']))
    
    cycle_loss_net = df_agg['Slow Loss'] - df_agg['Fast Gain']
    cycle_loss_plot = cycle_loss_net.clip(lower=0)
    
    fig.add_trace(go.Bar(x=df_agg[x_col], y=cycle_loss_plot, name='Net Cycle Loss', marker_color=PASTEL_COLORS['orange']))
    fig.add_trace(go.Bar(x=df_agg[x_col], y=df_agg['Downtime Loss'], name='Downtime Loss', marker_color=PASTEL_COLORS['grey']))
    
    fig.add_trace(go.Scatter(x=df_agg[x_col], y=df_agg['Optimal Output'], name='Optimal Output', mode='lines', line=dict(color=PASTEL_COLORS['optimal_line'], dash='dot')))
    
    if "Target" in str(benchmark_mode):
        fig.add_trace(go.Scatter(x=df_agg[x_col], y=df_agg['Target Output'], name='Target Output', mode='lines', line=dict(color=PASTEL_COLORS['target_line'], dash='dash')))

    fig.update_layout(barmode='stack', title="Periodic Performance Breakdown", hovermode="x unified", height=450)
    return fig

def plot_shot_analysis(df_shots, zoom_y=None):
    if df_shots.empty: return go.Figure()
    fig = go.Figure()
    color_map = {'Slow Cycle': PASTEL_COLORS['red'], 'Fast Cycle': PASTEL_COLORS['orange'], 'On Target': PASTEL_COLORS['blue'], 'Downtime (Stop)': PASTEL_COLORS['grey'], 'Run Break (Excluded)': '#d3d3d3'}
    
    for shot_type, color in color_map.items():
        subset = df_shots[df_shots['shot_type'] == shot_type]
        if not subset.empty:
            fig.add_trace(go.Bar(x=subset['shot_time'], y=subset['actual_ct'], name=shot_type, marker_color=color, hovertemplate='Time: %{x}<br>CT: %{y:.2f}s<extra></extra>'))
            
    if 'run_id' in df_shots.columns:
        run_starts = df_shots.groupby('run_id')['shot_time'].min().sort_values()
        
        for i, start_time in enumerate(run_starts):
            if i > 0: 
                 fig.add_vline(x=start_time.timestamp() * 1000, line_width=2, line_dash="dash", line_color="purple")

    for r_id, run_df in df_shots.groupby('run_id'):
        lower = run_df['mode_lower'].iloc[0]
        upper = run_df['mode_upper'].iloc[0]
        start = run_df['shot_time'].min()
        end = run_df['shot_time'].max()
        
        fig.add_shape(type="rect", x0=start, x1=end, y0=lower, y1=upper, fillcolor="grey", opacity=0.2, line_width=0)
        if r_id == 0: fig.add_annotation(x=start, y=upper, text="Mode Tolerance Band", showarrow=False, yshift=10, font=dict(color="grey", size=10))

    avg_ref = df_shots['approved_ct'].mean()
    fig.add_hline(y=avg_ref, line_dash="dash", line_color="green", annotation_text=f"Avg Approved CT: {avg_ref:.2f}s")
    
    if zoom_y is None and not df_shots.empty:
        cts = df_shots['actual_ct']
        if len(cts) > 0:
            ref_max = df_shots['approved_ct'].max() * 4
            dist_max = cts.quantile(0.95) * 1.5
            zoom_y = max(ref_max, dist_max)

    layout_args = dict(title="Shot-by-Shot Analysis", yaxis_title="Cycle Time (sec)", hovermode="closest")
    if zoom_y: layout_args['yaxis_range'] = [0, zoom_y]
    fig.update_layout(**layout_args)
    return fig