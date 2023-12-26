import http.client
import json
import pandas as pd

# Common parameters
api_key =
api_host = "seeking-alpha.p.rapidapi.com"

# Common HTTP connection initialization
def initialize_connection():
    conn = http.client.HTTPSConnection(api_host)
    headers = {
        'X-RapidAPI-Key': api_key,
        'X-RapidAPI-Host': api_host
    }
    return conn, headers

# Common function to make API requests and parse JSON response
def make_api_request(endpoint):
    conn, headers = initialize_connection()
    conn.request("GET", endpoint, headers=headers)
    res = conn.getresponse()
    data = res.read()
    conn.close()
    return json.loads(data.decode("utf-8"))

# Function to calculate Cash Flow Growth
def calculate_cash_flow_growth(symbol):
    endpoint = f"/symbols/get-financials?symbol={symbol}&target_currency=USD&period_type=annual&statement_type=cash-flow-statement"
    cash_flow_data = make_api_request(endpoint)

    for section in cash_flow_data:
        if section["title"] == "Cash Flow From Operating Activities":
            operating_activities = section["rows"][0]["cells"]
            year1 = operating_activities[6]["raw_value"]
            year2 = operating_activities[5]["raw_value"]
            cash_flow_growth = ((year2 - year1) / year1) * 100
            return cash_flow_growth
    return None

# Function to calculate EPS and Revenue Growth
def calculate_eps_revenue_growth(symbol):
    endpoint = f"/symbols/get-financials?symbol={symbol}&target_currency=USD&period_type=annual&statement_type=income-statement"
    response_data = make_api_request(endpoint)

    # Check if the response data is a list and contains data
    if isinstance(response_data, list) and response_data:
        # Extracting necessary data for each symbol
        eps_2022 = None
        eps_2023 = None
        eps_ttm = None
        revenue_2022 = None
        revenue_2023 = None

        for section in response_data:
            for row in section['rows']:
                if row['name'] == 'diluted_eps':
                    eps_2022 = row['cells'][8]['raw_value']  # Assuming 8th index is for Sep 2022
                    eps_2023 = row['cells'][9]['raw_value']  # Assuming 9th index is for Sep 2023
                    eps_ttm = row['cells'][10]['raw_value']  # Assuming TTM is the last cell
                elif row['name'] == 'total_revenue':
                    revenue_2022 = row['cells'][8]['raw_value']
                    revenue_2023 = row['cells'][9]['raw_value']

        # Calculating percentages
        eps_growth = ((eps_2023 - eps_2022) / eps_2022) * 100 if eps_2022 and eps_2023 else None
        revenue_growth = ((revenue_2023 - revenue_2022) / revenue_2022) * 100 if revenue_2022 and revenue_2023 else None

        # Creating a dictionary for each stock's data
        stock_data = {
            'Symbol': symbol,
            '2022 EPS': eps_2022,
            '2023 EPS': eps_2023,
            'EPS (TTM)': eps_ttm,
            '2022 Revenue': revenue_2022,
            '2023 Revenue': revenue_2023,
            '% Growth (EPS)': eps_growth,
            '% Growth (Revenue)': revenue_growth
        }

        return stock_data
    else:
        print(f"Error: No data found for {symbol}")
        return None

# Initialize the master list of stock symbols
master_stock_symbols = ["orcl", "googl", "tsla"]  # Add your stock symbols here

# API calls for stock metrics
fields = ['marketcap', 'pe_ratio', 'price_cf_ratio', 'pb_ratio', 'ev_ebitda',
          'revenueGrowth', 'gross_margin', 'net_margin', 'roe']
fields_url = "%2C".join(fields)

data_list = []

# Loop through each stock symbol in master_stock_symbols
for symbol in master_stock_symbols:
    endpoint = f"/symbols/get-metrics?symbols={symbol}&fields={fields_url}"
    parsed_data = make_api_request(endpoint)

    metric_types = {metric['id']: metric['attributes']['field'] for metric in parsed_data['included'] if metric['type'] == 'metric_type'}

    for item in parsed_data['data']:
        stock_id = item['relationships']['ticker']['data']['id']
        stock_slug = next(ticker['attributes']['slug'] for ticker in parsed_data['included'] if ticker['type'] == 'ticker' and ticker['id'] == stock_id)
        metric_id = item['relationships']['metric_type']['data']['id']
        metric_field = metric_types[metric_id]
        value = item['attributes']['value']
        data_list.append({'Stock': stock_slug, 'Metric': metric_field, 'Value': value})

# Create DataFrames from the collected data
metrics_df = pd.DataFrame(data_list)

# Pivot the DataFrame to have one row per symbol and columns for metrics
pivot_df = metrics_df.pivot(index='Stock', columns='Metric', values='Value')

# Retrieve valuation data
conn = http.client.HTTPSConnection("seeking-alpha.p.rapidapi.com")
headers = {
    'X-RapidAPI-Key': api_key,
    'X-RapidAPI-Host': api_host
}

# Making the API request for valuation data
symbols_for_valuation = "%2C".join(master_stock_symbols)
conn.request("GET", f"/symbols/get-valuation?symbols={symbols_for_valuation}", headers=headers)
res = conn.getresponse()
data = res.read()
conn.close()

# Parsing the JSON response for valuation data
valuation_data = json.loads(data.decode("utf-8"))

# Check if 'data' key is in valuation_data and process the data
if 'data' in valuation_data:
    # Create a list to store data
    valuation_data_list = []

    # Iterate through each item and extract values
    for item in valuation_data['data']:
        stock_id = item['id']
        attributes = item['attributes']

        # Extracting values
        price_sales = attributes.get('priceSales')
        peg_ratio = attributes.get('pegRatio')

        # Append data to the list
        valuation_data_list.append({'Stock': stock_id, 'Price/Sales (TTM)': price_sales, 'PEG (Price/Earnings to Growth) Ratio': peg_ratio})
else:
    print("Key 'data' not found in valuation_data")

# Create a DataFrame from the list of valuation dictionaries
valuation_df = pd.DataFrame(valuation_data_list)

# Convert the stock names to lowercase in valuation_df
valuation_df['Stock'] = valuation_df['Stock'].str.lower()

# Convert the stock names to lowercase in pivot_df
pivot_df.index = pivot_df.index.str.lower()

# Merge valuation_df with pivot_df using a left join on the "Stock" column
final_df = pivot_df.merge(valuation_df, on='Stock', how='left')

# Create a dictionary to store cash flow growth for each stock symbol
cf_growth_dict = {}

# Calculate and store Cash Flow Growth for each stock symbol
for symbol in master_stock_symbols:
    cash_flow_growth = calculate_cash_flow_growth(symbol)
    if cash_flow_growth is not None:
        cf_growth_dict[symbol.lower()] = cash_flow_growth
    else:
        print(f"Cash Flow From Operating Activities data not found for {symbol.upper()}.")

# Create a DataFrame from the cf_growth_dict
df_cf_growth = pd.DataFrame(list(cf_growth_dict.items()), columns=['Stock', 'Cash Flow Growth'])
df_cf_growth.set_index('Stock', inplace=True)  # Set 'Stock' as the index

# Merge df_cf_growth with final_df using a left join on the "Stock" column
final_df = final_df.merge(df_cf_growth, on='Stock', how='left')

# Create an empty list to store data for EPS and Revenue Growth
eps_revenue_growth_data = []

# Calculate EPS and Revenue Growth for each stock symbol
for symbol in master_stock_symbols:
    stock_data = calculate_eps_revenue_growth(symbol)
    if stock_data is not None:
        eps_revenue_growth_data.append(stock_data)
    else:
        print(f"No data found for {symbol}.")

# Create a DataFrame from the collected EPS and Revenue Growth data
eps_df = pd.DataFrame(eps_revenue_growth_data)

# Convert the stock names to lowercase in eps_df
eps_df['Symbol'] = eps_df['Symbol'].str.lower()

# Rename columns in eps_df to avoid conflicts during merging
eps_df.columns = [f"{col}_eps" if col != 'Symbol' else 'Stock' for col in eps_df.columns]

# Merge eps_df with final_df using a left join on the "Stock" column
final_df = final_df.merge(eps_df, on='Stock', how='left')

# Save the final DataFrame to an Excel file
excel_output_file = 'final_stock_metrics.xlsx'
final_df.to_excel(excel_output_file, index=False)

print(f'Data saved to {excel_output_file}')