import pandas as pd

# Define file paths
advertised_product_report_path = 'Sponsored Products Advertised product report.xlsx'
campaign_report_path = 'Sponsored Products Campaign report.csv'

# Load the data from the provided Excel files
advertised_product_report = pd.read_excel(advertised_product_report_path)
campaign_report = pd.read_csv(campaign_report_path)

# Convert Date columns to datetime
advertised_product_report['Date'] = pd.to_datetime(advertised_product_report['Date'])
campaign_report['Date'] = pd.to_datetime(campaign_report['Date'])

# Remove $ sign and convert Spend and Sales columns to numeric
campaign_report['Spend'] = campaign_report['Spend'].replace('[\$,]', '', regex=True).astype(float)
campaign_report['7 Day Total Sales '] = campaign_report['7 Day Total Sales '].replace('[\$,]', '', regex=True).astype(float)

# Convert percentage columns to numerical by removing '%' and converting to float
campaign_report['Total Advertising Cost of Sales (ACOS) '] = campaign_report['Total Advertising Cost of Sales (ACOS) '].str.rstrip('%').astype('float') / 100
campaign_report['Click-Thru Rate (CTR)'] = campaign_report['Click-Thru Rate (CTR)'].str.rstrip('%').astype('float') / 100

# Define input variables
comparison_period = 14
ASIN_list = ["ASIN"]

# Calculate the date ranges for the current and previous periods
end_date = campaign_report['Date'].max()
start_date_current = end_date - pd.Timedelta(days=comparison_period)
start_date_previous = start_date_current - pd.Timedelta(days=comparison_period)
end_date_previous = start_date_previous + pd.Timedelta(days=comparison_period)

# Filter the advertised product report based on ASINs
filtered_advertised_report = advertised_product_report[
 advertised_product_report['Advertised ASIN'].isin(ASIN_list)
]

# Extract relevant campaign names
relevant_campaigns = filtered_advertised_report['Campaign Name'].unique()

# Filter the campaign report using the extracted campaign names
filtered_campaign_report = campaign_report[
 campaign_report['Campaign Name'].isin(relevant_campaigns)
]

# Filter data for the current and previous periods
current_period_data = filtered_campaign_report[
 (filtered_campaign_report['Date'] >= start_date_current) &
 (filtered_campaign_report['Date'] <= end_date)
]

previous_period_data = filtered_campaign_report[
 (filtered_campaign_report['Date'] >= start_date_previous) &
 (filtered_campaign_report['Date'] <= end_date_previous)
]

def calculate_campaign_metrics(data):
 grouped_data = data.groupby('Campaign Name').agg({
 'Impressions': 'sum',
 'Clicks': 'sum',
 'Spend': 'sum',
 '7 Day Total Sales ': 'sum',
 '7 Day Total Orders (#)': 'sum'
 }).reset_index()

 # Calculate additional metrics
 grouped_data['CPC'] = grouped_data['Spend'] / grouped_data['Clicks']
 grouped_data['CVR'] = grouped_data['7 Day Total Orders (#)'] / grouped_data['Clicks']
 grouped_data['Total Sales Share'] = grouped_data['7 Day Total Sales '] / grouped_data['7 Day Total Sales '].sum()
 grouped_data['Total Spend Share'] = grouped_data['Spend'] / grouped_data['Spend'].sum()
 grouped_data['ACoS'] = (grouped_data['Spend'] / grouped_data['7 Day Total Sales ']) * 100
 grouped_data['CTR'] = grouped_data['Clicks'] / grouped_data['Impressions'] * 100

 return grouped_data

# Calculate metrics for the current period
current_campaign_metrics = calculate_campaign_metrics(current_period_data)

# Calculate metrics for the previous period
previous_campaign_metrics = calculate_campaign_metrics(previous_period_data)

# Merge current and previous metrics on Campaign Name to compare
campaign_comparison_df = current_campaign_metrics.merge(previous_campaign_metrics, on='Campaign Name', suffixes=('_current', '_previous'))

# Calculate the percentage differences
def calculate_percentage_difference(current, previous):
 if previous == 0 and current == 0:
 return 0
 elif previous == 0:
 return 99999
 elif current == 0:
 return -99999
 else:
 return ((current - previous) / previous) * 100

campaign_comparison_df['CPC Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['CPC_current'], row['CPC_previous']), axis=1
)
campaign_comparison_df['Total Spend Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['Total Spend Share_current'], row['Total Spend Share_previous']), axis=1
)
campaign_comparison_df['CVR Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['CVR_current'], row['CVR_previous']), axis=1
)
campaign_comparison_df['Total Impressions Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['Impressions_current'], row['Impressions_previous']), axis=1
)
campaign_comparison_df['Total Sales Share Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['Total Sales Share_current'], row['Total Sales Share_previous']), axis=1
)
campaign_comparison_df['Orders Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['7 Day Total Orders (#)_current'], row['7 Day Total Orders (#)_previous']), axis=1
)
campaign_comparison_df['Sales Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['7 Day Total Sales _current'], row['7 Day Total Sales _previous']), axis=1
)
campaign_comparison_df['ACoS Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['ACoS_current'], row['ACoS_previous']), axis=1
)
campaign_comparison_df['CTR Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['CTR_current'], row['CTR_previous']), axis=1
)
campaign_comparison_df['Clicks Difference (%)'] = campaign_comparison_df.apply(
 lambda row: calculate_percentage_difference(row['Clicks_current'], row['Clicks_previous']), axis=1
)

# Rearrange columns to put current and previous metrics side by side
campaign_comparison_df = campaign_comparison_df[[
 'Campaign Name',
 'CPC_current', 'CPC_previous', 'CPC Difference (%)',
 'Total Spend Share_current', 'Total Spend Share_previous', 'Total Spend Difference (%)',
 'CVR_current', 'CVR_previous', 'CVR Difference (%)',
 'Impressions_current', 'Impressions_previous', 'Total Impressions Difference (%)',
 'Total Sales Share_current', 'Total Sales Share_previous', 'Total Sales Share Difference (%)',
 '7 Day Total Orders (#)_current', '7 Day Total Orders (#)_previous', 'Orders Difference (%)',
 '7 Day Total Sales _current', '7 Day Total Sales _previous', 'Sales Difference (%)',
 'ACoS_current', 'ACoS_previous', 'ACoS Difference (%)',
 'CTR_current', 'CTR_previous', 'CTR Difference (%)',
 'Clicks_current', 'Clicks_previous', 'Clicks Difference (%)'
]]

# Display the tidy comparison dataframe
#print(campaign_comparison_df.head())

# Show the complete tidy comparison dataframe
campaign_comparison_df.head()