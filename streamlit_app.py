import streamlit as st
import pandas as pd
import plotly.express as px
import io

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
        # Check for cleaned duplicate versions of monthly columns
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
        all_portfolio_levels = ['All'] + df['PORTFOLIO_OBS_LEVEL'].dropna().unique().tolist()
        selected_portfolio = st.sidebar.selectbox("Select Portfolio Level", all_portfolio_levels)

        all_sub_portfolio_levels = ['All'] + df['SUB_PORTFOLIO_OBS_LEVEL'].dropna().unique().tolist()
        selected_sub_portfolio = st.sidebar.selectbox("Select Sub-Portfolio Level", all_sub_portfolio_levels)

        all_managers = ['All'] + df['PROJECT_MANAGER'].dropna().unique().tolist()
        selected_manager = st.sidebar.selectbox("Select Project Manager", all_managers)

        all_brs_classifications = ['All'] + df['BRS_CLASSIFICATION'].dropna().unique().tolist()
        selected_brs_classification = st.sidebar.selectbox("Select BRS Classification", all_brs_classifications)

        # Apply filters
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
        st.dataframe(
            filtered_df[[
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
            }),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")

        # --- Monthly Spend Trends ---
        st.subheader("2025 Monthly Spend Trends")

        # Identify monthly columns for plotting, making sure they exist
        monthly_actuals_cols_present = [col for col in df.columns if col.startswith('2025_') and col.endswith('_A')]
        monthly_forecasts_cols_present = [col for col in df.columns if col.startswith('2025_') and col.endswith('_F')]
        monthly_plan_cols_present = [col for col in df.columns if col.startswith('2025_') and col.endswith('_CP')]

        # Create a melted DataFrame for time series plotting
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

            # Combine all monthly data
            monthly_combined_df = pd.concat([monthly_data_actuals, monthly_data_forecasts, monthly_data_plan])

            # Order months correctly for plotting
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
        st.warning("Please upload a CSV file with valid data to proceed.")

else:
    st.info("Upload your Capital Project CSV file to get started!")

