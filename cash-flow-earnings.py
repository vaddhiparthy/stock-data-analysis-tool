import http.client
import json
import pandas as pd

# List of stock symbols to fetch data for
stock_symbols = ["AAPL", "MSFT", "GOOGL"]  # Add more symbols as needed
output_file = "financial_data.xlsx"  # Specify the output Excel file name

# Initialize an empty list to store data for each stock
data_list = []

# Establishing the connection
conn = http.client.HTTPSConnection("seeking-alpha.p.rapidapi.com")
headers = {
    'X-RapidAPI-Key': ,  # Your API key
    'X-RapidAPI-Host': "seeking-alpha.p.rapidapi.com"
}

# Iterate through the stock symbols
for symbol in stock_symbols:
    # Making the API request for each symbol
    conn.request("GET",
                 f"/symbols/get-financials?symbol={symbol}&target_currency=USD&period_type=annual&statement_type=income-statement",
                 headers=headers)
    res = conn.getresponse()
    data = res.read()

    # Parsing the response
    response = json.loads(data.decode("utf-8"))

    # Extracting necessary data for each symbol
    eps_2022 = None
    eps_2023 = None
    eps_ttm = None
    revenue_2022 = None
    revenue_2023 = None

    for section in response:
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

    data_list.append(stock_data)

# Creating a DataFrame from the list of dictionaries
df = pd.DataFrame(data_list)

# Save the DataFrame to an Excel file
df.to_excel(output_file, index=False)

print(f"Financial data saved to {output_file}")
