import streamlit as st
import pandas as pd
import plotly.express as px
import io
import inspect # Import inspect module

# Set page configuration for a wider layout
st.set_page_config(layout="wide", page_title="Capital Project Portfolio Dashboard")

# --- Data Loading and Cleaning ---
@st.cache_data
def load_data(uploaded_file):
    """Loads and preprocesses the CSV data."""
    try:
        df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return pd.DataFrame() # Return empty DataFrame on error

    # 1. Clean Column Names
    # A function to clean individual column names
    def clean_col_name(col_name):
        col_name = str(col_name).strip().replace(' ', '_').replace('+', '_').replace('.', '').replace('-', '_').replace('__', '_')
        col_name = col_name.upper()
        # Specific corrections for common typos/inconsistencies observed
        if 'PROJEC_TID' in col_name:
            col_name = col_name.replace('PROJEC_TID', 'PROJECT_ID')
        if 'INI_MATIVE_PROGRAM' in col_name:
            col_name = col_name.replace('INI_MATIVE_PROGRAM', 'INITIATIVE_PROGRAM')
        if 'ALL_PRIOR_YEARS_A' in col_name:
            col_name = col_name.replace('ALL_PRIOR_YEARS_A', 'ALL_PRIOR_YEARS_ACTUALS')
        if 'C_URRENT_EAC' in col_name: # Fix for 'C URRENT_EAC'
            col_name = 'CURRENT_EAC'
        if 'QE_RUN_RATE' in col_name: # Fix for 'QE Run Rate'
            col_name = 'QE_RUN_RATE'
        return col_name

    df.columns = [clean_col_name(col) for col in df.columns]

    # Handle duplicate column names by making them unique if they exist (e.g., Rate, Rate.1)
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i, i_val in enumerate(cols[cols == dup].index.values)]
    df.columns = cols

    # Identify financial columns that need numeric conversion
    financial_cols = [
        'ALL_PRIOR_YEARS_ACTUALS', 'BUSINESS_ALLOCATION', 'CURRENT_EAC',
        'QE_FORECAST_VS_QE_PLAN', 'FORECAST_VS_BA',
        'YE_RUN', 'RATE', 'QE_RUN', 'RATE_1', 'QE_RUN_RATE_0', # Use RATE_1 or similar if it becomes unique after cleaning
        '2025_01_A', '2025_02_A', '2025_03_A', '2025_04_A', '2025_05_A', '2025_06_A', '2025_07_A',
        '2025_08_A', '2025_09_A', '2025_10_A', '2025_11_A', '2025_12_A',
        '2025_01_F', '2025_02_F', '2025_03_F', '2025_04_F', '2025_05_F', '2025_06_F', '2025_07_F',
        '2025_08_F', '2025_09_F', '2025_10_F', '2025_11_F', '2025_12_F',
        '2025_01_CP', '2025_02_CP', '2025_03_CP', '2025_04_CP', '2025_05_CP', '2025_06_CP', '2025_07_CP',
        '2025_08_CP', '2025_09_CP', '2025_10_CP', '2025_11_CP', '2025_12_CP',
        # Check for cleaned duplicate versions of monthly columns if they exist after cleaning
        '2025_01__A_1', '2025_02__A_1', '2025_03__A_1', '2025_04__A_1', '2025_05__A_1', '2025_06__A_1', '2025_07__A_1',
        '2025_08__A_1', '2025_09_A_1', '2025_10__A_1', '2025_11__A_1', '2025_12__A_1',
        '2025_01_F_1', '2025_02_F_1', '2025_03_F_1', '2025_04_F_1', '2025_05_F_1', '2025_06_F_1', '2025_07_F_1',
        '2025_08_F_1', '2025_09_F_1', '2025_10_F_1', '2025_11_F_1', '2025_12_F_1',
        '2025_01_CP_1', '2025_02_CP_1', '2025_03_CP_1', '2025_04_CP_1', '2025_05_CP_1', '2025_06_CP_1', '2025_07_CP_1',
        '2025_08_CP_1', '2025_09_CP_1', '2025_10_CP_1', '2025_11_CP_1', '2025_12_CP_1'
    ]

    # Filter financial columns to only those actually present in the DataFrame
    financial_cols_present = [col for col in financial_cols if col in df.columns]

    # Convert financial columns to numeric, handling commas and potential errors
    for col in financial_cols_present:
        # Check if the column is of object type (string) before cleaning
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.replace(',', '').str.replace(' ', '').replace('', '0').astype(float)
        else: # If already numeric, ensure it's float to avoid errors with mixed types later
            df[col] = pd.to_numeric(df[col], errors='coerce')


    # Identify monthly columns for 2025 after cleaning
    monthly_actuals_cols = [col for col in df.columns if col.startswith('2025_') and col.endswith('_A')]
    monthly_forecasts_cols = [col for col in df.columns if col.startswith('2025_') and col.endswith('_F')]
    monthly_plan_cols = [col for col in df.columns if col.startswith('2025_') and col.endswith('_CP')]

    # Ensure all columns are numeric
    for col_list in [monthly_actuals_cols, monthly_forecasts_cols, monthly_plan_cols]:
        for col in col_list:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Calculate total 2025 Actuals, Forecasts, and Plans
    df['TOTAL_2025_ACTUALS'] = df[monthly_actuals_cols].sum(axis=1) if monthly_actuals_cols else 0
    df['TOTAL_2025_FORECASTS'] = df[monthly_forecasts_cols].sum(axis=1) if monthly_forecasts_cols else 0
    df['TOTAL_2025_CAPITAL_PLAN'] = df[monthly_plan_cols].sum(axis=1) if monthly_plan_cols else 0

    # Calculate Total Actuals to Date (Prior Years + 2025 Actuals)
    df['TOTAL_ACTUALS_TO_DATE'] = df['ALL_PRIOR_YEARS_ACTUALS'] + df['TOTAL_2025_ACTUALS']

    return df

# --- Streamlit App Layout ---
st.title("💰 Capital Project Portfolio Dashboard")
st.markdown("""
    This dashboard provides an interactive overview of your capital projects, allowing you to track financials,
    monitor trends, and identify variances.
""")

uploaded_file = st.file_uploader("Upload your Capital Project CSV file", type=["csv"])

if uploaded_file is not None:
    df = load_data(uploaded_file)

    if not df.empty:
        # --- Sidebar Filters ---
        st.sidebar.header("Filter Projects")
        # Ensure column names are accessed using their clean format
        all_portfolio_levels = ['All'] + df['PORTFOLIO_OBS_LEVEL'].dropna().unique().tolist()
        selected_portfolio = st.sidebar.selectbox("Select Portfolio Level", all_portfolio_levels)

        all_sub_portfolio_levels = ['All'] + df['SUB_PORTFOLIO_OBS_LEVEL'].dropna().unique().tolist()
        selected_sub_portfolio = st.sidebar.selectbox("Select Sub-Portfolio Level", all_sub_portfolio_levels)

        all_managers = ['All'] + df['PROJECT_MANAGER'].dropna().unique().tolist()
        selected_manager = st.sidebar.selectbox("Select Project Manager", all_managers)

        all_brs_classifications = ['All'] + df['BRS_CLASSIFICATION'].dropna().unique().tolist()
        selected_brs_classification = st.sidebar.selectbox("Select BRS Classification", all_brs_classifications)

        # Apply filters - also ensuring clean column names are used
        filtered_df = df.copy()
        if selected_portfolio != 'All':
            filtered_df = filtered_df[filtered_df['PORTFOLIO_OBS_LEVEL'] == selected_portfolio]
        if selected_sub_portfolio != 'All':
            filtered_df = filtered_df[filtered_df['SUB_PORTFOLIO_OBS_LEVEL'] == selected_sub_portfolio]
        if selected_manager != 'All':
            filtered_df = filtered_df[filtered_df['PROJECT_MANAGER'] == selected_manager]
        if selected_brs_classification != 'All':
            filtered_df = filtered_df[filtered_df['BRS_CLASSIFICATION'] == selected_brs_classification]

        st.subheader("Key Metrics Overview")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            total_business_allocation = filtered_df['BUSINESS_ALLOCATION'].sum()
            st.metric(label="Total Business Allocation", value=f"${total_business_allocation:,.2f}")
        with col2:
            total_current_eac = filtered_df['CURRENT_EAC'].sum()
            st.metric(label="Total Current EAC", value=f"${total_current_eac:,.2f}")
        with col3:
            total_actuals_to_date = filtered_df['TOTAL_ACTUALS_TO_DATE'].sum()
            st.metric(label="Total Actuals To Date", value=f"${total_actuals_to_date:,.2f}")
        with col4:
            total_projects = len(filtered_df)
            st.metric(label="Number of Projects", value=total_projects)

        st.markdown("---")

        # --- Project Table ---
        st.subheader("Project Details")
        project_details_table = filtered_df[[
            'PORTFOLIO_OBS_LEVEL', 'SUB_PORTFOLIO_OBS_LEVEL', 'MASTER_PROJECT_ID',
            'PROJECT_NAME', 'PROJECT_MANAGER', 'BRS_CLASSIFICATION',
            'BUSINESS_ALLOCATION', 'CURRENT_EAC', 'ALL_PRIOR_YEARS_ACTUALS',
            'TOTAL_2025_ACTUALS', 'TOTAL_2025_FORECASTS', 'TOTAL_2025_CAPITAL_PLAN',
            'QE_FORECAST_VS_QE_PLAN', 'FORECAST_VS_BA'
        ]].style.format({
            'BUSINESS_ALLOCATION': "${:,.2f}",
            'CURRENT_EAC': "${:,.2f}",
            'ALL_PRIOR_YEARS_ACTUALS': "${:,.2f}",
            'TOTAL_2025_ACTUALS': "${:,.2f}",
            'TOTAL_2025_FORECASTS': "${:,.2f}",
            'TOTAL_2025_CAPITAL_PLAN': "${:,.2f}",
            'QE_FORECAST_VS_QE_PLAN': "{:,.2f}",
            'FORECAST_VS_BA': "{:,.2f}"
        })
        st.dataframe(project_details_table, use_container_width=True, hide_index=True)

        st.markdown("---")

        # --- Monthly Spend Trends ---
        st.subheader("2025 Monthly Spend Trends")

        monthly_actuals_cols_present = [col for col in df.columns if col.startswith('2025_') and col.endswith('_A')]
        monthly_forecasts_cols_present = [col for col in df.columns if col.startswith('2025_') and col.endswith('_F')]
        monthly_plan_cols_present = [col for col in df.columns if col.startswith('2025_') and col.endswith('_CP')]

        monthly_combined_df = pd.DataFrame()
        if monthly_actuals_cols_present or monthly_forecasts_cols_present or monthly_plan_cols_present:
            monthly_data_actuals = filtered_df[monthly_actuals_cols_present].sum().reset_index()
            monthly_data_actuals.columns = ['Month', 'Amount']
            monthly_data_actuals['Type'] = 'Actuals'

            monthly_data_forecasts = filtered_df[monthly_forecasts_cols_present].sum().reset_index()
            monthly_data_forecasts.columns = ['Month', 'Amount']
            monthly_data_forecasts['Type'] = 'Forecasts'

            monthly_data_plan = filtered_df[monthly_plan_cols_present].sum().reset_index()
            monthly_data_plan.columns = ['Month', 'Amount']
            monthly_data_plan['Type'] = 'Capital Plan'

            monthly_combined_df = pd.concat([monthly_data_actuals, monthly_data_forecasts, monthly_data_plan])

            month_order = [
                '2025_01', '2025_02', '2025_03', '2025_04', '2025_05', '2025_06',
                '2025_07', '2025_08', '2025_09', '2025_10', '2025_11', '2025_12'
            ]
            monthly_combined_df['Month_Sort'] = monthly_combined_df['Month'].str.replace('_A', '').str.replace('_F', '').str.replace('_CP', '')
            monthly_combined_df['Month_Sort'] = pd.Categorical(monthly_combined_df['Month_Sort'], categories=month_order, ordered=True)
            monthly_combined_df = monthly_combined_df.sort_values('Month_Sort')

            fig_monthly_trends = px.line(
                monthly_combined_df,
                x='Month_Sort',
                y='Amount',
                color='Type',
                title='Monthly Capital Trends (Actuals, Forecasts, Plan)',
                labels={'Month_Sort': 'Month', 'Amount': 'Amount ($)'},
                line_shape='linear',
                markers=True
            )
            fig_monthly_trends.update_layout(hovermode="x unified", legend_title_text='Type')
            fig_monthly_trends.update_xaxes(title_text="Month (2025)")
            fig_monthly_trends.update_yaxes(title_text="Amount ($)")
            st.plotly_chart(fig_monthly_trends, use_container_width=True)
        else:
            st.warning("No 2025 monthly actuals, forecasts, or plan data found for trend analysis.")

        st.markdown("---")

        # --- Variance Analysis ---
        st.subheader("Variance Analysis")
        col_var1, col_var2 = st.columns(2)
        fig_qe_variance = None
        fig_ba_variance = None

        if 'QE_FORECAST_VS_QE_PLAN' in filtered_df.columns:
            with col_var1:
                fig_qe_variance = px.bar(
                    filtered_df,
                    x='PROJECT_NAME',
                    y='QE_FORECAST_VS_QE_PLAN',
                    title='QE Forecast vs QE Plan Variance',
                    labels={'QE_FORECAST_VS_QE_PLAN': 'Variance'},
                    height=400
                )
                fig_qe_variance.update_layout(xaxis_title="Project Name", yaxis_title="Variance")
                st.plotly_chart(fig_qe_variance, use_container_width=True)
        else:
            with col_var1:
                st.info("Column 'QE_FORECAST_VS_QE_PLAN' not found for variance analysis.")

        if 'FORECAST_VS_BA' in filtered_df.columns:
            with col_var2:
                fig_ba_variance = px.bar(
                    filtered_df,
                    x='PROJECT_NAME',
                    y='FORECAST_VS_BA',
                    title='Forecast vs Business Allocation Variance',
                    labels={'FORECAST_VS_BA': 'Variance'},
                    height=400
                )
                fig_ba_variance.update_layout(xaxis_title="Project Name", yaxis_title="Variance")
                st.plotly_chart(fig_ba_variance, use_container_width=True)
        else:
            with col_var2:
                st.info("Column 'FORECAST_VS_BA' not found for variance analysis.")

        st.markdown("---")

        # --- Allocation Breakdown ---
        st.subheader("Capital Allocation Breakdown")
        col_alloc1, col_alloc2, col_alloc3 = st.columns(3)
        fig_portfolio_alloc = None
        fig_sub_portfolio_alloc = None
        fig_brs_alloc = None

        if 'BUSINESS_ALLOCATION' in filtered_df.columns:
            with col_alloc1:
                if not filtered_df['PORTFOLIO_OBS_LEVEL'].isnull().all():
                    fig_portfolio_alloc = px.pie(
                        filtered_df,
                        names='PORTFOLIO_OBS_LEVEL',
                        values='BUSINESS_ALLOCATION',
                        title='Allocation by Portfolio Level',
                        hole=0.3
                    )
                    st.plotly_chart(fig_portfolio_alloc, use_container_width=True)
                else:
                    st.info("No 'PORTFOLIO_OBS_LEVEL' data available for allocation.")

            with col_alloc2:
                if not filtered_df['SUB_PORTFOLIO_OBS_LEVEL'].isnull().all():
                    fig_sub_portfolio_alloc = px.pie(
                        filtered_df,
                        names='SUB_PORTFOLIO_OBS_LEVEL',
                        values='BUSINESS_ALLOCATION',
                        title='Allocation by Sub-Portfolio Level',
                        hole=0.3
                    )
                    st.plotly_chart(fig_sub_portfolio_alloc, use_container_width=True)
                else:
                    st.info("No 'SUB_PORTFOLIO_OBS_LEVEL' data available for allocation.")

            with col_alloc3:
                if not filtered_df['BRS_CLASSIFICATION'].isnull().all():
                    fig_brs_alloc = px.pie(
                        filtered_df,
                        names='BRS_CLASSIFICATION',
                        values='BUSINESS_ALLOCATION',
                        title='Allocation by BRS Classification',
                        hole=0.3
                    )
                    st.plotly_chart(fig_brs_alloc, use_container_width=True)
                else:
                    st.info("No 'BRS_CLASSIFICATION' data available for allocation.")
        else:
            st.info("Column 'BUSINESS_ALLOCATION' not found for allocation breakdown.")

        st.markdown("---")

        # --- Detailed Project Financials (on selection) ---
        st.subheader("Detailed Project Financials")
        project_names = ['Select a Project'] + filtered_df['PROJECT_NAME'].dropna().unique().tolist()
        selected_project_name = st.selectbox("Select a project for detailed view:", project_names)

        fig_project_monthly = None
        if selected_project_name != 'Select a Project':
            project_details = filtered_df[filtered_df['PROJECT_NAME'] == selected_project_name].iloc[0]

            st.write(f"### Details for: {project_details['PROJECT_NAME']}")

            # Display key financial metrics for the selected project
            col_d1, col_d2, col_d3 = st.columns(3)
            with col_d1:
                st.metric(label="Business Allocation", value=f"${project_details.get('BUSINESS_ALLOCATION', 0):,.2f}")
            with col_d2:
                st.metric(label="Current EAC", value=f"${project_details.get('CURRENT_EAC', 0):,.2f}")
            with col_d3:
                st.metric(label="All Prior Years Actuals", value=f"${project_details.get('ALL_PRIOR_YEARS_ACTUALS', 0):,.2f}")

            st.write("#### 2025 Monthly Breakdown:")
            monthly_breakdown_df = pd.DataFrame({
                'Month': [f"2025_{i:02d}" for i in range(1, 13)],
                'Actuals': [project_details.get(f'2025_{i:02d}_A', 0) for i in range(1, 13)],
                'Forecasts': [project_details.get(f'2025_{i:02d}_F', 0) for i in range(1, 13)],
                'Capital Plan': [project_details.get(f'2025_{i:02d}_CP', 0) for i in range(1, 13)]
            })

            # Format the monthly breakdown table
            st.dataframe(monthly_breakdown_df.style.format({
                'Actuals': "${:,.2f}",
                'Forecasts': "${:,.2f}",
                'Capital Plan': "${:,.2f}"
            }), use_container_width=True, hide_index=True)

            # Bar chart for monthly breakdown for the selected project
            monthly_project_melted = monthly_breakdown_df.melt(id_vars=['Month'], var_name='Type', value_name='Amount')

            fig_project_monthly = px.bar(
                monthly_project_melted,
                x='Month',
                y='Amount',
                color='Type',
                barmode='group',
                title=f'Monthly Financials for {selected_project_name}',
                labels={'Amount': 'Amount ($)'}
            )
            st.plotly_chart(fig_project_monthly, use_container_width=True)

        else:
            st.info("Select a project from the dropdown to see its detailed monthly financials.")

        st.markdown("---")

        # --- Report Generation Feature ---
        st.subheader("Generate Professional Report")
        st.markdown("Click the button below to generate a comprehensive HTML report of the current dashboard view.")

        # Function to generate the HTML report content
        def generate_html_report(filtered_df, total_business_allocation, total_current_eac, total_actuals_to_date, total_projects,
                                 monthly_combined_df, fig_monthly_trends, fig_qe_variance, fig_ba_variance,
                                 fig_portfolio_alloc, fig_sub_portfolio_alloc, fig_brs_alloc,
                                 selected_project_name, project_details, fig_project_monthly):

            report_html_content = f"""
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Capital Project Portfolio Report</title>
                <style>
                    body {{ font-family: sans-serif; line-height: 1.6; margin: 20px; color: #333; }}
                    h1, h2, h3 {{ color: #004d40; }}
                    .metric-container {{ display: flex; justify-content: space-around; flex-wrap: wrap; margin-bottom: 20px; }}
                    .metric-box {{ border: 1px solid #ddd; border-radius: 8px; padding: 15px; margin: 10px; flex: 1; min-width: 200px; text-align: center; background-color: #f9f9f9; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
                    .metric-label {{ font-size: 0.9em; color: #555; }}
                    .metric-value {{ font-size: 1.5em; font-weight: bold; color: #222; margin-top: 5px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                    th {{ background-color: #e6f2f0; }}
                    .chart-container {{ margin-top: 30px; border: 1px solid #eee; padding: 10px; border-radius: 8px; background-color: #fff; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
                    .section-title {{ margin-top: 40px; border-bottom: 2px solid #004d40; padding-bottom: 10px; }}
                    footer {{ text-align: center; margin-top: 50px; padding-top: 20px; border-top: 1px solid #eee; font-size: 0.8em; color: #777; }}
                </style>
            </head>
            <body>
                <h1>Capital Project Portfolio Report</h1>
                <p>Generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

                <h2 class="section-title">Key Metrics Overview</h2>
                <div class="metric-container">
                    <div class="metric-box">
                        <div class="metric-label">Total Business Allocation</div>
                        <div class="metric-value">${total_business_allocation:,.2f}</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-label">Total Current EAC</div>
                        <div class="metric-value">${total_current_eac:,.2f}</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-label">Total Actuals To Date</div>
                        <div class="metric-value">${total_actuals_to_date:,.2f}</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-label">Number of Projects</div>
                        <div class="metric-value">{total_projects}</div>
                    </div>
                </div>

                <h2 class="section-title">Filtered Project Details</h2>
                {filtered_df[[
                    'PORTFOLIO_OBS_LEVEL', 'SUB_PORTFOLIO_OBS_LEVEL', 'MASTER_PROJECT_ID',
                    'PROJECT_NAME', 'PROJECT_MANAGER', 'BRS_CLASSIFICATION',
                    'BUSINESS_ALLOCATION', 'CURRENT_EAC', 'ALL_PRIOR_YEARS_ACTUALS',
                    'TOTAL_2025_ACTUALS', 'TOTAL_2025_FORECASTS', 'TOTAL_2025_CAPITAL_PLAN',
                    'QE_FORECAST_VS_QE_PLAN', 'FORECAST_VS_BA'
                ]].style.format({
                    'BUSINESS_ALLOCATION': "${:,.2f}",
                    'CURRENT_EAC': "${:,.2f}",
                    'ALL_PRIOR_YEARS_ACTUALS': "${:,.2f}",
                    'TOTAL_2025_ACTUALS': "${:,.2f}",
                    'TOTAL_2025_FORECASTS': "${:,.2f}",
                    'TOTAL_2025_CAPITAL_PLAN': "${:,.2f}",
                    'QE_FORECAST_VS_QE_PLAN': "{:,.2f}",
                    'FORECAST_VS_BA': "{:,.2f}"
                }).to_html(index=False)}

                <h2 class="section-title">2025 Monthly Spend Trends</h2>
                <div class="chart-container">
                    {fig_monthly_trends.to_html(full_html=False, include_plotlyjs='cdn') if monthly_combined_df is not None and not monthly_combined_df.empty else ''}
                </div>

                <h2 class="section-title">Variance Analysis</h2>
                <div style="display: flex; justify-content: space-around; flex-wrap: wrap;">
                    <div class="chart-container" style="flex: 1; min-width: 45%;">
                        {fig_qe_variance.to_html(full_html=False, include_plotlyjs='cdn') if fig_qe_variance else ''}
                    </div>
                    <div class="chart-container" style="flex: 1; min-width: 45%;">
                        {fig_ba_variance.to_html(full_html=False, include_plotlyjs='cdn') if fig_ba_variance else ''}
                    </div>
                </div>

                <h2 class="section-title">Capital Allocation Breakdown</h2>
                <div style="display: flex; justify-content: space-around; flex-wrap: wrap;">
                    <div class="chart-container" style="flex: 1; min-width: 30%;">
                        {fig_portfolio_alloc.to_html(full_html=False, include_plotlyjs='cdn') if fig_portfolio_alloc else ''}
                    </div>
                    <div class="chart-container" style="flex: 1; min-width: 30%;">
                        {fig_sub_portfolio_alloc.to_html(full_html=False, include_plotlyjs='cdn') if fig_sub_portfolio_alloc else ''}
                    </div>
                    <div class="chart-container" style="flex: 1; min-width: 30%;">
                        {fig_brs_alloc.to_html(full_html=False, include_plotlyjs='cdn') if fig_brs_alloc else ''}
                    </div>
                </div>
            """
            if selected_project_name != 'Select a Project' and project_details is not None and fig_project_monthly is not None:
                # Re-create monthly breakdown DF for report to ensure it has correct formatting
                monthly_breakdown_df_html = pd.DataFrame({
                    'Month': [f"2025_{i:02d}" for i in range(1, 13)],
                    'Actuals': [project_details.get(f'2025_{i:02d}_A', 0) for i in range(1, 13)],
                    'Forecasts': [project_details.get(f'2025_{i:02d}_F', 0) for i in range(1, 13)],
                    'Capital Plan': [project_details.get(f'2025_{i:02d}_CP', 0) for i in range(1, 13)]
                }).style.format({
                    'Actuals': "${:,.2f}",
                    'Forecasts': "${:,.2f}",
                    'Capital Plan': "${:,.2f}"
                }).to_html(index=False) # Important to set index=False for HTML tables

                report_html_content += f"""
                <h2 class="section-title">Detailed Financials for {selected_project_name}</h2>
                <div class="metric-container">
                    <div class="metric-box">
                        <div class="metric-label">Business Allocation</div>
                        <div class="metric-value">${project_details.get('BUSINESS_ALLOCATION', 0):,.2f}</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-label">Current EAC</div>
                        <div class="metric-value">${project_details.get('CURRENT_EAC', 0):,.2f}</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-label">All Prior Years Actuals</div>
                        <div class="metric-value">${project_details.get('ALL_PRIOR_YEARS_ACTUALS', 0):,.2f}</div>
                    </div>
                </div>
                <h3>2025 Monthly Breakdown:</h3>
                {monthly_breakdown_df_html}
                <div class="chart-container">
                    {fig_project_monthly.to_html(full_html=False, include_plotlyjs='cdn')}
                </div>
                """
            report_html_content += """
                <footer>
                    <p>Generated by Capital Project Portfolio Dashboard Streamlit App.</p>
                </footer>
            </body>
            </html>
            """
            return report_html_content

        if st.button("Generate Report (HTML)"):
            # Prepare HTML components for the report
            # The project_details_table itself is already styled, just need to convert to HTML string
            # from the .style object
            # Access the underlying DataFrame for .to_html() and apply formatting within it
            # Ensure the correct columns are selected for the report HTML table
            project_details_df_for_report = filtered_df[[
                'PORTFOLIO_OBS_LEVEL', 'SUB_PORTFOLIO_OBS_LEVEL', 'MASTER_PROJECT_ID',
                'PROJECT_NAME', 'PROJECT_MANAGER', 'BRS_CLASSIFICATION',
                'BUSINESS_ALLOCATION', 'CURRENT_EAC', 'ALL_PRIOR_YEARS_ACTUALS',
                'TOTAL_2025_ACTUALS', 'TOTAL_2025_FORECASTS', 'TOTAL_2025_CAPITAL_PLAN',
                'QE_FORECAST_VS_QE_PLAN', 'FORECAST_VS_BA'
            ]]
            
            # The generate_html_report function now directly takes `filtered_df`
            # and performs the styling and HTML conversion internally for the main table.
            # So, we pass the df object here, not its .style object.

            report_content = generate_html_report(
                filtered_df, total_business_allocation, total_current_eac, total_actuals_to_date, total_projects,
                monthly_combined_df, fig_monthly_trends, fig_qe_variance, fig_ba_variance,
                fig_portfolio_alloc, fig_sub_portfolio_alloc, fig_brs_alloc,
                selected_project_name, project_details if selected_project_name != 'Select a Project' else None, fig_project_monthly
            )

            st.download_button(
                label="Download Report as HTML",
                data=report_content,
                file_name="capital_project_report.html",
                mime="text/html"
            )

    else:
        st.warning("Please upload a CSV file with valid data to proceed.")

else:
    st.info("Upload your Capital Project CSV file to get started!")

# --- View Source Code Feature ---
st.markdown("---")
with st.expander("View Application Source Code"):
    # Dynamically get the source code of the current file
    source_code = inspect.getsource(inspect.currentframe())
    st.code(source_code, language='python')
