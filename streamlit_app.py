import streamlit as st
import pandas as pd
import plotly.express as px
import io

def app():
    st.set_page_config(layout="wide")
    st.title("Excel Data Extractor and Visualizer")

    st.write(
        """
        Upload your Excel file to extract specific data points and visualize the 2025
        monthly data as time series graphs, along with other relevant charts.
        """
    )

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        try:
            # Read the Excel file into a Pandas DataFrame
            # Use io.BytesIO to handle the uploaded file
            df = pd.read_excel(io.BytesIO(uploaded_file.getvalue()))
            st.success("File successfully uploaded!")

            # Standardized and consolidated list of columns
            # Clean up column names in DataFrame by stripping whitespace and replacing problematic characters
            df.columns = df.columns.str.strip().str.replace(' ', '_').str.replace('+', '_').str.upper()

            selected_columns = [
                'PORTFOLIO_OBS_LEVEL1',
                'SUB_PORTFOLIO_OBS_LEVEL',
                'MASTER_PROJECT_ID',
                'MASTER_PRJ_PROJ_NAME',
                'PROJECT_NAME',
                'PROJECT_MANAGER',
                'BRS_CLASSIFICATION',
                'INITIATIVE_PROGRAM',
                'ALL_PRIOR_YEARS_A',
                'BUSINESS_ALLOCATION',
                'CURRENT_EAC',
                'YE_RUN',
                'RATE',
                'QE_RUN',
                'QE_FORECAST_VS_QE_PLAN',
                'FORECAST_VS_BA',
            ]

            # Generate 2025 monthly columns dynamically
            month_suffixes = ['A', 'F', 'CP', 'AB']
            for month in range(1, 13):
                month_str = f"2025_{month:02d}"
                for suffix in month_suffixes:
                    selected_columns.append(f"{month_str}_{suffix}")

            # Check if all desired columns exist in the DataFrame
            # Normalize the case and remove extra spaces from user-provided column names for matching
            df_columns_normalized = [col.strip().replace(' ', '_').replace('+', '_').upper() for col in df.columns]

            # Filter selected_columns to only include those present in the DataFrame
            # Map normalized selected_columns back to actual DataFrame column names if necessary,
            # or just use the normalized ones for extraction if the DF columns are already normalized.
            # For simplicity, we assume df.columns has been normalized.
            existing_selected_columns = [col for col in selected_columns if col in df.columns]
            missing_columns = [col for col in selected_columns if col not in df.columns]


            if missing_columns:
                st.warning(
                    f"The following columns were not found in your Excel file with the exact names: {', '.join(missing_columns)}. "
                    "Please ensure the column headers in your file exactly match the expected names or adjust the script."
                )

            if not existing_selected_columns:
                st.error("No matching columns found in the uploaded file based on the required headers. Please check your Excel file's headers.")
                return

            # Extract the desired data
            extracted_df = df[existing_selected_columns].copy()

            st.header("Extracted Data (First 5 Rows)")
            st.dataframe(extracted_df.head())

            # --- Time Series Graphs for 2025 Data ---
            st.header("2025 Monthly Data Time Series")

            time_series_cols = [col for col in existing_selected_columns if col.startswith('2025_')]

            if time_series_cols:
                # Prepare data for time series plotting
                melted_df_list = []
                for col in time_series_cols:
                    if '_' in col:
                        parts = col.split('_')
                        if len(parts) >= 3:
                            month = parts[1]
                            data_type = parts[2]
                            temp_df = extracted_df[[col]].copy()
                            temp_df.columns = ['Value']
                            temp_df['Month'] = int(month)
                            temp_df['Type'] = data_type
                            melted_df_list.append(temp_df)

                if melted_df_list:
                    melted_df = pd.concat(melted_df_list)
                    # Drop rows where 'Value' might be non-numeric and coerce to numeric
                    melted_df['Value'] = pd.to_numeric(melted_df['Value'], errors='coerce')
                    melted_df.dropna(subset=['Value'], inplace=True)

                    if not melted_df.empty:
                        # Group by Month and Type to get the average or sum for plotting
                        agg_melted_df = melted_df.groupby(['Month', 'Type']).agg(
                            Mean_Value=('Value', 'mean'),
                            Sum_Value=('Value', 'sum')
                        ).reset_index()

                        st.subheader("Monthly Mean Values (by Type)")
                        fig_mean = px.line(agg_melted_df, x='Month', y='Mean_Value', color='Type',
                                         title='Average Monthly Values by Type (2025)',
                                         labels={'Mean_Value': 'Average Value', 'Month': 'Month (2025)'})
                        # Set month names for x-axis ticks
                        month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                        fig_mean.update_xaxes(tickmode='array', tickvals=list(range(1, 13)), ticktext=month_names)
                        st.plotly_chart(fig_mean, use_container_width=True)

                        st.subheader("Monthly Sum Values (by Type)")
                        fig_sum = px.line(agg_melted_df, x='Month', y='Sum_Value', color='Type',
                                        title='Sum of Monthly Values by Type (2025)',
                                        labels={'Sum_Value': 'Sum of Values', 'Month': 'Month (2025)'})
                        fig_sum.update_xaxes(tickmode='array', tickvals=list(range(1, 13)), ticktext=month_names)
                        st.plotly_chart(fig_sum, use_container_width=True)
                    else:
                        st.info("No numerical 2025 monthly data available after cleaning for plotting time series.")
                else:
                    st.info("No 2025 monthly data found in the selected columns for plotting.")
            else:
                st.info("No 2025 monthly columns found for time series analysis in the extracted data.")


            # --- Other Graphs ---
            st.header("Other Data Visualizations")

            # Example 1: Bar chart for 'PROJECT_MANAGER'
            if 'PROJECT_MANAGER' in extracted_df.columns:
                st.subheader("Projects per Project Manager")
                pm_counts = extracted_df['PROJECT_MANAGER'].value_counts().reset_index()
                pm_counts.columns = ['PROJECT_MANAGER', 'Count']
                fig_pm = px.bar(pm_counts, x='PROJECT_MANAGER', y='Count',
                                title='Number of Projects per Project Manager')
                st.plotly_chart(fig_pm, use_container_width=True)
            else:
                st.info("Column 'PROJECT_MANAGER' not found for visualization.")

            # Example 2: Bar chart for 'BUSINESS_ALLOCATION'
            if 'BUSINESS_ALLOCATION' in extracted_df.columns:
                st.subheader("Business Allocation Distribution")
                biz_alloc_counts = extracted_df['BUSINESS_ALLOCATION'].value_counts().reset_index()
                biz_alloc_counts.columns = ['BUSINESS_ALLOCATION', 'Count']
                fig_biz = px.bar(biz_alloc_counts, x='BUSINESS_ALLOCATION', y='Count',
                                 title='Distribution of Business Allocation')
                st.plotly_chart(fig_biz, use_container_width=True)
            else:
                st.info("Column 'BUSINESS_ALLOCATION' not found for visualization.")

            # Example 3: Histogram for 'CURRENT_EAC' (assuming numerical)
            if 'CURRENT_EAC' in extracted_df.columns:
                st.subheader("Distribution of Current EAC")
                # Ensure it's numeric, coerce errors to NaN and drop NaNs
                numeric_eac = pd.to_numeric(extracted_df['CURRENT_EAC'], errors='coerce').dropna()
                if not numeric_eac.empty:
                    fig_eac = px.histogram(numeric_eac, x='CURRENT_EAC',
                                           title='Distribution of Current EAC Values')
                    st.plotly_chart(fig_eac, use_container_width=True)
                else:
                    st.info("Column 'CURRENT_EAC' found but contains no numerical data for histogram.")
            else:
                st.info("Column 'CURRENT_EAC' not found for visualization.")


            # Example 4: Scatter plot for ALL_PRIOR_YEARS_A vs CURRENT_EAC
            relevant_cols = ['ALL_PRIOR_YEARS_A', 'CURRENT_EAC']
            if all(col in extracted_df.columns for col in relevant_cols):
                st.subheader("ALL_PRIOR_YEARS_A vs CURRENT_EAC")
                scatter_df = extracted_df[relevant_cols].copy()
                scatter_df['ALL_PRIOR_YEARS_A'] = pd.to_numeric(scatter_df['ALL_PRIOR_YEARS_A'], errors='coerce')
                scatter_df['CURRENT_EAC'] = pd.to_numeric(scatter_df['CURRENT_EAC'], errors='coerce')
                scatter_df.dropna(inplace=True)

                if not scatter_df.empty:
                    fig_scatter = px.scatter(scatter_df, x='ALL_PRIOR_YEARS_A', y='CURRENT_EAC',
                                             title='Prior Years Actuals vs Current EAC')
                    st.plotly_chart(fig_scatter, use_container_width=True)
                else:
                    st.info("No numerical data available for 'ALL_PRIOR_YEARS_A' and 'CURRENT_EAC' for scatter plot.")
            else:
                st.info("Columns 'ALL_PRIOR_YEARS_A' or 'CURRENT_EAC' not found for scatter plot.")


        except Exception as e:
            st.error(f"An error occurred: {e}. Please ensure the file is a valid Excel file and its contents are as expected.")

if __name__ == "__main__":
    app()
