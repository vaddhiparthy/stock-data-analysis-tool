import http.client
import json
import pandas as pd

# Initialize HTTP connection for valuation
conn = http.client.HTTPSConnection("seeking-alpha.p.rapidapi.com")
headers = {
    'X-RapidAPI-Key': ,
    'X-RapidAPI-Host': "seeking-alpha.p.rapidapi.com"
}

# Making the API request
conn.request("GET", "/symbols/get-valuation?symbols=aapl%2Ctsla", headers=headers)
res = conn.getresponse()
data = res.read()
conn.close()

# Parsing the JSON response
valuation_data = json.loads(data.decode("utf-8"))

# Check if 'data' key is in valuation_data and process the data
if 'data' in valuation_data:
    # Create a list to store data
    data_list = []

    # Iterate through each item and extract values
    for item in valuation_data['data']:
        stock_id = item['id']
        attributes = item['attributes']

        # Extracting values
        price_sales = attributes.get('priceSales')
        peg_ratio = attributes.get('pegRatio')

        # Append data to the list
        data_list.append({'Stock': stock_id, 'Price/Sales (TTM)': price_sales, 'PEG (Price/Earnings to Growth) Ratio': peg_ratio})
else:
    print("Key 'data' not found in valuation_data")

# Create a DataFrame from the list of dictionaries
df_valuation = pd.DataFrame(data_list)

# Save the DataFrame to an Excel file
output_file = 'valuation_data.xlsx'
df_valuation.to_excel(output_file, index=False)

# Display the DataFrame
print(df_valuation)
print(f"Data saved to Excel file: {output_file}")
