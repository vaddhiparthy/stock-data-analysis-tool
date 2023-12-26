import http.client
import json
import pandas as pd

# Define the stock symbols
stock_symbols = ["ipoof", "tga", "esea", "bgfv", "grin", "bbw"]  # Add your stock symbols here

# Initialize HTTP connection
conn = http.client.HTTPSConnection("seeking-alpha.p.rapidapi.com")
headers = {
    'X-RapidAPI-Key': ,
    'X-RapidAPI-Host': "seeking-alpha.p.rapidapi.com"
}

# Define fields for the request
fields = [
    'marketcap', 'pe_ratio', 'price_cf_ratio', 'pb_ratio', 'ev_ebitda',
    'revenueGrowth', 'gross_margin', 'net_margin', 'roe'
]
fields_url = "%2C".join(fields)

# Initialize an empty list to store data
data_list = []

# Make the request for each stock symbol
for symbol in stock_symbols:
    conn.request("GET", f"/symbols/get-metrics?symbols={symbol}&fields={fields_url}", headers=headers)
    res = conn.getresponse()
    data = res.read()

    # Parse the JSON response
    parsed_data = json.loads(data.decode("utf-8"))

    # Extracting metric types and their corresponding field names
    metric_types = {metric['id']: metric['attributes']['field'] for metric in parsed_data['included'] if metric['type'] == 'metric_type'}

    # Organizing data
    for item in parsed_data['data']:
        stock_id = item['relationships']['ticker']['data']['id']
        stock_slug = next(ticker['attributes']['slug'] for ticker in parsed_data['included'] if ticker['type'] == 'ticker' and ticker['id'] == stock_id)
        metric_id = item['relationships']['metric_type']['data']['id']
        metric_field = metric_types[metric_id]
        value = item['attributes']['value']
        data_list.append({'Stock': stock_slug, 'Metric': metric_field, 'Value': value})

# Creating DataFrame
df_metrics = pd.DataFrame(data_list)

# Pivot the DataFrame
pivot_df = df_metrics.pivot(index='Stock', columns='Metric', values='Value')

# Rename columns for clarity
pivot_df = pivot_df.rename(columns={
    'marketcap': 'Market Cap',
    'pe_ratio': 'P/E Ratio',
    'price_cf_ratio': 'Price/Cash Flow Ratio',
    'pb_ratio': 'Price/Book Ratio',
    'ev_ebitda': 'EV/EBITDA',
    'revenueGrowth': 'Revenue Growth',
    'gross_margin': 'Gross Margin',
    'net_margin': 'Net Margin',
    'roe': 'ROE'
})

# Define the file paths
output_csv_file = 'get-metrics.csv'
output_excel_file = 'get-metrics.xlsx'
output_text_file = 'get-metrics.txt'

# Save the DataFrame to different file formats
pivot_df.to_csv(output_csv_file)
pivot_df.to_excel(output_excel_file, sheet_name="Stock Metrics")
with open(output_text_file, 'w') as text_file:
    text_file.write(pivot_df.to_string())

# Print a message to confirm the file paths
print(f"DataFrame saved to CSV file: {output_csv_file}")
print(f"DataFrame saved to Excel file: {output_excel_file}")
print(f"DataFrame saved to text file: {output_text_file}")
