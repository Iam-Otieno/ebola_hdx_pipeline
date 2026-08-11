import os
import sys
import openpyxl
import pandas as pd
from hdx.api.configuration import Configuration
from hdx.data.dataset import Dataset

# ==============================================================================
# CONFIGURATION & PARAMETERS
# ==============================================================================
DATASET_SLUG = "republique-democratique-du-congo-cas-et-deces-d-ebola"
EXCEL_MODEL_PATH = "DRC_Ebola_Epidemiology_Dashboard_and_Analysis.xlsx"
DOWNLOAD_DIR = "./downloads"
USER_AGENT = "Ebola_Live_Tracker_Learning_Project"

# ==============================================================================
# 1. EXTRACT PHASE (HDX API Ingestion)
# ==============================================================================
def extract_hdx_data(slug: str, download_folder: str = DOWNLOAD_DIR) -> str:
    """
    Connects to HDX API, retrieves metadata for the dataset slug,
    and downloads the latest primary CSV resource file.
    """
    print("🚀 Starting HDX Extraction Pipeline...")
    print("🔌 Connecting to HDX API...")
    
    try:
        Configuration.create(
            hdx_site="prod", 
            user_agent=USER_AGENT, 
            hdx_read_only=True
        )
        print("✅ Connected to HDX API.")
    except Exception as e:
        print(f"⚠️ HDX Configuration note: {e}")

    print(f"🔍 Fetching dataset metadata for slug: '{slug}'...")
    dataset = Dataset.read_from_hdx(slug)

    if not dataset or len(dataset.get_resources()) == 0:
        raise ValueError(f"❌ No dataset resources found on HDX for slug: '{slug}'")

    first_resource = dataset.get_resources()[0]
    os.makedirs(download_folder, exist_ok=True)
    
    print(f"⬇️ Downloading resource: {first_resource.get('name')}...")
    url, path = first_resource.download(download_folder)
    print(f"🎉 Extract Success! Raw CSV saved locally to: {path}")
    return path

# ==============================================================================
# 2. TRANSFORM PHASE (Cleaning, Velocity & Data Quality Auditing)
# ==============================================================================
def transform_data(raw_file_path: str) -> pd.DataFrame:
    """
    Cleans raw HDX records, computes daily net changes, 7-day velocity,
    and assigns dynamic data quality audit flags.
    """
    print("\n⚙️ Starting Data Transformation Phase...")
    df = pd.read_csv(raw_file_path)

    # Standardize and parse dates
    df['reference_date'] = pd.to_datetime(df['reference_date'])

    # Define analytical grouping keys
    group_cols = ['location_name', 'measure', 'case_classification']
    df = df.sort_values(by=group_cols + ['reference_date']).reset_index(drop=True)

    # 1. Calculate Previous Value & Daily Net Change
    df['previous_value'] = df.groupby(group_cols)['value'].shift(1)
    df['daily_net_change'] = df['value'] - df['previous_value']

    # 2. Detect Reporting Gaps (Days since last report)
    df['previous_date'] = df.groupby(group_cols)['reference_date'].shift(1)
    df['days_since_last_report'] = (df['reference_date'] - df['previous_date']).dt.days

    # 3. Data Quality Audit Engine
    def evaluate_dq(row):
        flags = []
        if pd.notnull(row['daily_net_change']) and row['daily_net_change'] < 0:
            flags.append("Downward Revision")
        if pd.notnull(row['days_since_last_report']) and row['days_since_last_report'] > 1:
            flags.append("Reporting Gap (>1 Day)")
        return "; ".join(flags) if flags else "Normal"

    df['dq_status'] = df.apply(evaluate_dq, axis=1)

    # 4. Compute 7-Day Rolling Velocity (Daily Average Growth)
    df['velocity_7day'] = (
        df.groupby(group_cols)['daily_net_change']
        .transform(lambda x: x.rolling(window=7, min_periods=1).mean())
    )

    max_date_str = df['reference_date'].max().strftime('%Y-%m-%d')
    print(f"✅ Transform Success! Processed {len(df):,} rows through {max_date_str}.")
    return df

# ==============================================================================
# 3. LOAD PHASE (Excel Tab Syncing & KPI Card Updates)
# ==============================================================================
def load_and_update_excel(df_clean: pd.DataFrame, excel_path: str):
    """
    Overwrites the 'Processed Data' tab in Excel and updates cell references
    on the 'Executive Summary' and 'Dashboard' sheets dynamically.
    """
    print(f"\n💾 Starting Load & Sync Phase for workbook: '{excel_path}'...")

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"❌ Target workbook '{excel_path}' does not exist.")

    # --------------------------------------------------------------------------
    # Step A: Update Processed Data Sheet via OpenPyXL
    # --------------------------------------------------------------------------
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_clean.to_excel(writer, sheet_name="Processed Data", index=False)
        print("  ✓ 'Processed Data' sheet successfully overwritten.")
    except PermissionError:
        print(f"\n❌ PERMISSION ERROR: Please close '{excel_path}' in Microsoft Excel and re-run!")
        sys.exit(1)

    # --------------------------------------------------------------------------
    # Step B: Compute Executive Summary KPIs in Python
    # --------------------------------------------------------------------------
    latest_date = df_clean['reference_date'].max()
    latest_date_str = latest_date.strftime('%Y-%m-%d')

    cases_df = df_clean[(df_clean['measure'] == 'cases') & (df_clean['case_classification'] == 'confirmed')]
    deaths_df = df_clean[(df_clean['measure'] == 'deaths') & (df_clean['case_classification'] == 'confirmed')]

    daily_cases = cases_df.groupby('reference_date')['value'].sum()
    daily_deaths = deaths_df.groupby('reference_date')['value'].sum()

    latest_cases = int(daily_cases.loc[latest_date])
    latest_deaths = int(daily_deaths.loc[latest_date])
    latest_cfr = latest_deaths / latest_cases if latest_cases > 0 else 0.0

    # 7-Day Velocity Calculation
    seven_days_ago_date = latest_date - pd.Timedelta(days=7)
    if seven_days_ago_date in daily_cases.index:
        cases_7d_ago = daily_cases.loc[seven_days_ago_date]
        deaths_7d_ago = daily_deaths.loc[seven_days_ago_date] if seven_days_ago_date in daily_deaths.index else 0
    else:
        cases_7d_ago = daily_cases.iloc[-8] if len(daily_cases) >= 8 else daily_cases.iloc[0]
        deaths_7d_ago = daily_deaths.iloc[-8] if len(daily_deaths) >= 8 else daily_deaths.iloc[0]

    case_growth_7d = int(latest_cases - cases_7d_ago)
    death_growth_7d = int(latest_deaths - deaths_7d_ago)
    velocity_7d = case_growth_7d / 7.0

    # --------------------------------------------------------------------------
    # Step C: Write Updated Values Directly to Excel Sheet Cells
    # --------------------------------------------------------------------------
    wb = openpyxl.load_workbook(excel_path)

    # 1. Update Executive Summary Tab
    ws_exec = wb["Executive Summary"]
    ws_exec["A3"].value = f"Data cutoff: {latest_date_str} | Source series are cumulative snapshots"
    ws_exec["B6"].value = latest_cases
    ws_exec["B9"].value = latest_deaths
    ws_exec["B16"].value = case_growth_7d
    ws_exec["B17"].value = velocity_7d
    ws_exec["D8"].value = (
        f"Acceleration: Confirmed cases increased by {case_growth_7d:,} in the last 7 days "
        f"({velocity_7d:.1f} per day), while confirmed deaths increased by {death_growth_7d:,}."
    )

    # 2. Update Dashboard Tab Subheader String
    ws_dash = wb["Dashboard"]
    ws_dash["A3"].value = f"Situation as of {latest_date_str} | Bundibugyo virus disease | Decision-support view"

    wb.save(excel_path)
    
    print("\n📊 Updated Dashboard Metrics Summary:")
    print(f"  • As of Date:        {latest_date_str}")
    print(f"  • Confirmed Cases:   {latest_cases:,}")
    print(f"  • Confirmed Deaths:  {latest_deaths:,}")
    print(f"  • Confirmed CFR:     {latest_cfr * 100:.2f}%")
    print(f"  • 7-Day Velocity:    {velocity_7d:.2f} cases/day")
    print(f"🎉 Load Success! Dashboard cards and Executive Summary are fully in sync.")

# ==============================================================================
# PIPELINE EXECUTION CONTROLLER
# ==============================================================================
if __name__ == "__main__":
    try:
        # Step 1: Extract
        raw_csv_path = extract_hdx_data(DATASET_SLUG)
        
        # Step 2: Transform
        df_transformed = transform_data(raw_csv_path)
        
        # Step 3: Load
        load_and_update_excel(df_transformed, EXCEL_MODEL_PATH)
        
        print("\n🏆 ETL PIPELINE EXECUTION COMPLETED SUCCESSFULLY!")

    except Exception as e:
        print(f"\n❌ PIPELINE FAILED: {str(e)}")