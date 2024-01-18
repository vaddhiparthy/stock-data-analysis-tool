import http.client
import json
import pandas as pd
import certifi
import requests
from urllib.request import urlopen
import ssl
from pandas import ExcelWriter

# Prompting for user inputs
api_key_seeking_alpha = input("Enter Seeking Alpha API Key: ")
api_key_financial = input("Enter Financial Modeling Prep API Key: ")
api_key_alpha_vantage = input("Enter Alpha Vantage API Key: ")
excel_file_path = input("Enter the path to your Excel file: ")
sheet_name = input("Enter the Excel sheet name: ")
stock_symbols_input = input("Enter stock symbols separated by commas (e.g., AAPL, MSFT, AMZN): ")
master_stock_symbols = stock_symbols_input.split(',')

# Ensure stock symbols are stripped of spaces
master_stock_symbols = [symbol.strip() for symbol in master_stock_symbols]

# Common API host
api_host = "seeking-alpha.p.rapidapi.com"

# Function to initialize HTTP connection (from Code #1)
def initialize_connection():
    conn = http.client.HTTPSConnection(api_host)
    headers = {
        'X-RapidAPI-Key': api_key_seeking_alpha,
        'X-RapidAPI-Host': api_host
    }
    return conn, headers

# Function to make API requests and parse JSON response (from Code #1)
def make_api_request(endpoint):
    conn, headers = initialize_connection()
    conn.request("GET", endpoint, headers=headers)
    res = conn.getresponse()
    data = res.read()
    conn.close()
    return json.loads(data.decode("utf-8"))

# Function to fetch and parse JSON data from a URL (from Code #2)
def get_jsonparsed_data(url):
    context = ssl.create_default_context(cafile=certifi.where())
    response = urlopen(url, context=context)
    data = response.read().decode("utf-8")
    return json.loads(data)

# Combined function for Cash Flow Growth, EPS, and Revenue Growth (from Code #1)
def calculate_metrics(symbol):
    # Initialize the HTTP connection
    conn, headers = initialize_connection()

     # Calculate Cash Flow Growth
    cash_flow_endpoint = f"/symbols/get-financials?symbol={symbol}&target_currency=USD&period_type=annual&statement_type=cash-flow-statement"
    cash_flow_data = make_api_request(cash_flow_endpoint)
    cash_flow_growth = None
    if isinstance(cash_flow_data, list) and cash_flow_data:
        for section in cash_flow_data:
            if section["title"] == "Cash Flow From Operating Activities":
                if 'rows' in section and section['rows'] and 'cells' in section['rows'][0]:
                    operating_activities = section["rows"][0]["cells"]
                    # New check: ensure 'cells' has enough elements
                    if len(operating_activities) > 6:
                        year1 = operating_activities[6]["raw_value"] if 'raw_value' in operating_activities[6] else None
                        year2 = operating_activities[5]["raw_value"] if 'raw_value' in operating_activities[5] else None
                        cash_flow_growth = ((year2 - year1) / year1) * 100 if year1 and year2 else None
                    break

    # Calculate EPS and Revenue Growth
    eps_revenue_endpoint = f"/symbols/get-financials?symbol={symbol}&target_currency=USD&period_type=annual&statement_type=income-statement"
    response_data = make_api_request(eps_revenue_endpoint)
    eps_growth, revenue_growth, eps_2022, eps_2023, eps_ttm, revenue_2022, revenue_2023 = None, None, None, None, None, None, None
    if isinstance(response_data, list) and response_data:
        for section in response_data:
            for row in section['rows']:
                if row['name'] == 'diluted_eps':
                    if 'cells' in row and len(row['cells']) > 10:
                        eps_2022 = row['cells'][8]['raw_value'] if 'raw_value' in row['cells'][8] else None
                        eps_2023 = row['cells'][9]['raw_value'] if 'raw_value' in row['cells'][9] else None
                        eps_ttm = row['cells'][10]['raw_value'] if 'raw_value' in row['cells'][10] else None
                elif row['name'] == 'total_revenue':
                    if 'cells' in row and len(row['cells']) > 9:
                        revenue_2022 = row['cells'][8]['raw_value'] if 'raw_value' in row['cells'][8] else None
                        revenue_2023 = row['cells'][9]['raw_value'] if 'raw_value' in row['cells'][9] else None

        eps_growth = ((eps_2023 - eps_2022) / eps_2022) * 100 if eps_2022 and eps_2023 else None
        revenue_growth = ((revenue_2023 - revenue_2022) / revenue_2022) * 100 if revenue_2022 and revenue_2023 else None
         # API calls for stock metrics
    fields = ['marketcap', 'pe_ratio', 'price_cf_ratio', 'pb_ratio', 'ev_ebitda', 'revenueGrowth', 'gross_margin', 'net_margin', 'roe']
    fields_url = "%2C".join(fields)
    metrics_endpoint = f"/symbols/get-metrics?symbols={symbol}&fields={fields_url}"
    metrics_data = make_api_request(metrics_endpoint)

    # Parsing the metrics data
    metrics_dict = {}
    if 'data' in metrics_data and 'included' in metrics_data:
        metric_types = {metric['id']: metric['attributes']['field'] for metric in metrics_data['included'] if metric['type'] == 'metric_type'}
        for item in metrics_data['data']:
            metric_id = item['relationships']['metric_type']['data']['id']
            metric_field = metric_types[metric_id]
            value = item['attributes']['value']
            metrics_dict[metric_field] = value

    # Making the API request for valuation data
    valuation_endpoint = f"/symbols/get-valuation?symbols={symbol}"
    valuation_data = make_api_request(valuation_endpoint)
    price_sales, peg_ratio = None, None
    if 'data' in valuation_data:
        for item in valuation_data['data']:
            price_sales = item['attributes'].get('priceSales')
            peg_ratio = item['attributes'].get('pegRatio')

    # Creating a dictionary for each stock's data
    stock_data = {
        'Symbol': symbol,
        'Cash Flow Growth': cash_flow_growth,
        '2022 EPS': eps_2022,
        '2023 EPS': eps_2023,
        'EPS (TTM)': eps_ttm,
        '2022 Revenue': revenue_2022,
        '2023 Revenue': revenue_2023,
        '% Growth (EPS)': eps_growth,
        '% Growth (Revenue)': revenue_growth
    }

    # Merge metrics and valuation data
    stock_data.update(metrics_dict)
    stock_data['Price/Sales (TTM)'] = price_sales
    stock_data['PEG (Price/Earnings to Growth) Ratio'] = peg_ratio

    return stock_data


# Function to get combined financial data from different API endpoints (from Code #2)
def get_combined_data(symbols, api_key1, api_key2):
    key_metrics_data = []
    ratios_data = []
    eps_data = []

    # Fetching data from FinancialModelingPrep for Key Metrics
    for symbol in symbols:
        url = f"https://financialmodelingprep.com/api/v3/key-metrics/{symbol}?period=annual&apikey={api_key1}"
        response_data = get_jsonparsed_data(url)
        if response_data:
            latest_pe = response_data[0].get('peRatio', None)
            last_year_pe = response_data[1].get('peRatio', None)
            blended_avg = (latest_pe + last_year_pe) / 2 if latest_pe and last_year_pe else None
            key_metrics_data.append({
                'Symbol': symbol,
                'Latest_PE': latest_pe,
                'Last_Year_PE': last_year_pe,
                'Blended_Avg_PE': blended_avg
            })



    # Fetching data from FinancialModelingPrep for Financial Ratios
    for symbol in symbols:
        url = f"https://financialmodelingprep.com/api/v3/ratios-ttm/{symbol}?apikey={api_key1}"
        data = get_jsonparsed_data(url)
        if data:
            ratios_data.append({
                'Symbol': symbol,
                'P/E Ratio TTM': data[0]['peRatioTTM'],
                'Price to Cash Flow Ratio TTM': data[0]['priceCashFlowRatioTTM'],
                'Price to Sales Ratio TTM': data[0]['priceSalesRatioTTM'],
                'Return on Equity TTM': data[0]['returnOnEquityTTM']
            })

    # Fetching EPS data from Alpha Vantage
    for symbol in symbols:
        try:
            response = requests.get(f'https://www.alphavantage.co/query?function=EARNINGS&symbol={symbol}&apikey={api_key2}')
            data = response.json()
            if 'quarterlyEarnings' in data:
                eps_data.append({
                    'Symbol': symbol,
                    'current_eps': float(data['quarterlyEarnings'][0]['reportedEPS']),
                    'prev_eps_1': float(data['quarterlyEarnings'][1]['reportedEPS']),
                    'prev_eps_2': float(data['quarterlyEarnings'][2]['reportedEPS']),
                    'prev_eps_3': float(data['quarterlyEarnings'][3]['reportedEPS']),
                    'est_eps_1': float(data['quarterlyEarnings'][4]['reportedEPS']),
                    'est_eps_2': float(data['quarterlyEarnings'][5]['reportedEPS'])
                })
        except Exception as e:
            print(f"Error fetching EPS data for {symbol}: {e}")

    # Converting to DataFrames
    df_key_metrics = pd.DataFrame(key_metrics_data)
    df_ratios = pd.DataFrame(ratios_data)
    df_eps = pd.DataFrame(eps_data)

    # Merging the DataFrames
    combined_df = pd.merge(df_key_metrics, df_ratios, on='Symbol')
    print(combined_df)
    combined_df = pd.merge(combined_df, df_eps, on='Symbol')

    return combined_df

def calculate_pe_growth(row):
    if row['Last_Year_PE'] != 0:  # Check if 'Last_Year_PE' is not zero
        return (row['Latest_PE'] - row['Last_Year_PE']) / row['Last_Year_PE']
    else:
        return None  # Return None or a suitable default value if 'Last_Year_PE' is zero

# Main Execution
# Fetch data from Code #1
data_from_code1 = [calculate_metrics(symbol) for symbol in master_stock_symbols]
df_code1 = pd.DataFrame(data_from_code1)

# Fetch data from Code #2
df_code2 = get_combined_data(master_stock_symbols, api_key_financial, api_key_alpha_vantage)

# Merging the DataFrames on 'Symbol'
final_df = pd.merge(df_code1, df_code2, on='Symbol', how='left')

#Creating blank columns for data to be fetched manually
final_df['EPS-g TTM-manual'] = ''
final_df['Gross Margin TTM'] = ''
final_df['Operating_NI_margin'] = ''
final_df['est_eps_1'] = ''
final_df['est_eps_2'] = ''
final_df['seperator1-blank']=''
final_df['seperator2-blank']=''
final_df['seperator3-blank']=''
final_df['seperator3-blank']=''
final_df['P/E Growth'] = final_df.apply(calculate_pe_growth, axis=1)

final_df['Symbol2']=final_df['Symbol']
#Matching Columns order per excel sheet
#selecting columns
Final_columns = [
    'Symbol',
    'marketcap',
    'P/E Ratio TTM',
    'Blended_Avg_PE',
    'Price to Cash Flow Ratio TTM',
    'Price/Sales (TTM)',
    'pb_ratio',
    'ev_ebitda',
    'PEG (Price/Earnings to Growth) Ratio',
    'EPS-g TTM-manual',  # Blank
    'Cash Flow Growth',
    'seperator1-blank', #Blank
    'gross_margin', # Blank
    'Operating_NI_margin',  # Blank
    'seperator2-blank', #Blank
    'Return on Equity TTM',
    'seperator3-blank', #Blank
    'prev_eps_1',
    'prev_eps_2',
    'prev_eps_3',
    'current_eps',
    'est_eps_1',#Blank
    'est_eps_2',#Blank
    'Symbol2',
    '2022 EPS',
    '2023 EPS',
    '% Growth (EPS)',
    '2022 Revenue',
    '2023 Revenue',
    '% Growth (Revenue)',
    'Last_Year_PE',
    'Latest_PE',
    'Blended_Avg_PE',
    'P/E Growth']
final_df=final_df[Final_columns]

# Save to the specified Excel file and sheet, starting from the 5th row, without including headers
with ExcelWriter(excel_file_path, mode='a', if_sheet_exists='replace') as writer:
    final_df.to_excel(writer, sheet_name=sheet_name, startrow=3, index=False, header=True)

print(f'Data saved to {excel_file_path} in sheet {sheet_name} with data starting from row 5')