import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

# List of tickers
tickers = ['ABSMARINE', 'ACCENTMIC', 'BIRLAMONEY', 'ADSL', 'ADVINFRA', 'ALICON', 'ANANTRAJ', 'ANNSWBL', 'ANUHPHRM', 'ATISHAY', 'AVG', 'BALUFORGE', 'BASILIC', 'BBOX', 'CPPL', 'CELLECOR', 'CREDITACC', 'CRISIL', 'DCG', 'DCI', 'DEEDEV', 'DEEM', 'DHARMAJ', 'DLINKINDIA', 'DODLA', 'DYNAMIC', 'ECOREC', 'EMKAYTOOLS', 'EMSLIMITED', 'EXICOM', 'GGBL', 'GPECO', 'HARIOMPIPE', 'INDRAMEDCO', 'IRMENERGY', 'K2INFRA', 'KALYANICAST', 'KCEIL', 'KILBURN', 'KRISHNADEF', 'KSOLVES', 'LACTOSE', 'LINCOLN', 'MADHUSUDAN', 'MANINDS', 'MEGATHERM', 'MMP', 'ORIANA', 'PGEL', 'PHANTOMFX', 'PPL', 'PRATHAM', 'PRLIND', 'RADHIKAJWE', 'RATNAVEER', 'RULKA', 'SJS', 'SALZERELEC', 'SHAREINDIA', 'SHERA', 'SHRIBALAJI', 'SHRIPISTON', 'SJLOGISTIC', 'SKYGOLD', 'SDBL', 'STARDELTA', 'STORAGE', 'SUPREMEPWR', 'SURYODAY', 'SUZLON', 'SWARAJ', 'SYSTANGO', 'TCL', 'VESUVIUS', 'VILAS', 'VPRPL', 'VIVIANA', 'WEALTH', 'WINSOL', 'YATHARTH','ZENTEC']


# Define base URL
base_url = "https://www.google.com/finance/quote/{ticker}:NSE"

# Set up headers to mimic a real browser request
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# List to hold the data
data = []

# Loop through each ticker and scrape the price
for ticker in tickers:
    try:
        url = base_url.format(ticker=ticker)
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raise an exception for HTTP errors

        soup = BeautifulSoup(response.text, 'html.parser')

        # Update class name according to the actual structure
        class_name = "YMlKec fxKbKc"  # Verify this class name
        price_element = soup.find(class_=class_name)

        if price_element:
            price = price_element.text.strip()[1:].replace(",", "")
            print(f"{ticker}: {price}")

            # Append data to list
            data.append({
                'Ticker': ticker,
                'Price': price,
                'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
        else:
            print(f"{ticker}: Price element not found. Check the class name or page structure.")

    except requests.exceptions.RequestException as e:
        print(f"{ticker}: An error occurred: {e}")

# Create a DataFrame with all the collected data
df = pd.DataFrame(data)

# Specify the filename and sheet name
filename = 'stock_prices.xlsx'
sheet_name = 'Prices'

# Append DataFrame to Excel file
try:
    # Load existing file if it exists
    existing_df = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
    updated_df = pd.concat([existing_df, df], ignore_index=True)
except FileNotFoundError:
    # If the file does not exist, create a new one
    updated_df = df

# Write DataFrame to Excel file
updated_df.to_excel(filename, sheet_name=sheet_name, index=False, engine='openpyxl')
print(f"Prices saved to {filename}")