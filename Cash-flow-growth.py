import http.client
import json

# Define the API endpoint and headers
conn = http.client.HTTPSConnection("seeking-alpha.p.rapidapi.com")
headers = {
    'X-RapidAPI-Key': ,
    'X-RapidAPI-Host': "seeking-alpha.p.rapidapi.com"
}

# Make the API request
conn.request("GET", "/symbols/get-financials?symbol=orcl&target_currency=USD&period_type=annual&statement_type=cash-flow-statement", headers=headers)

# Get the API response
res = conn.getresponse()
data = res.read()

# Parse the JSON response
response_data = json.loads(data)

# Function to calculate Cash Flow Growth
def calculate_cash_flow_growth(response_data):
    for section in response_data:
        if section["title"] == "Cash Flow From Operating Activities":
            operating_activities = section["rows"][0]["cells"]
            year1 = operating_activities[6]["raw_value"]
            year2 = operating_activities[5]["raw_value"]
            cash_flow_growth = ((year2 - year1) / year1) * 100
            return cash_flow_growth
    return None

# Calculate and print Cash Flow Growth
cash_flow_growth = calculate_cash_flow_growth(response_data)
if cash_flow_growth is not None:
    print(f"Cash Flow Growth: {cash_flow_growth:.2f}%")
else:
    print("Cash Flow From Operating Activities data not found in the response.")
