import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import os

# Function to fetch live data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        data = response.json()
        crypto_data = []
        for coin in data:
            crypto_data.append({
                "Name": coin["name"],
                "Symbol": coin["symbol"].upper(),
                "Price (USD)": coin["current_price"],
                "Market Cap (USD)": coin["market_cap"],
                "24h Volume (USD)": coin["total_volume"],
                "24h Price Change (%)": coin["price_change_percentage_24h"]
            })
        return pd.DataFrame(crypto_data)
    else:
        print("Failed to fetch data")
        return pd.DataFrame()

# Function to analyze data
def analyze_data(df):
    top_5_by_market_cap = df.nlargest(5, "Market Cap (USD)")[["Name", "Market Cap (USD)"]]
    avg_price = df["Price (USD)"].mean()
    highest_24h_change = df.loc[df["24h Price Change (%)"].idxmax()]
    lowest_24h_change = df.loc[df["24h Price Change (%)"].idxmin()]
    
    analysis = {
        "Top 5 Cryptos by Market Cap": top_5_by_market_cap.to_dict(orient="records"),
        "Average Price of Top 50 Cryptos": avg_price,
        "Highest 24h Change": highest_24h_change.to_dict(),
        "Lowest 24h Change": lowest_24h_change.to_dict()
    }
    return analysis

# Function to update Excel file

def update_excel(df, file_name="crypto_data.xlsx"):
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Live Crypto Data"
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    wb.save(file_name)
    os.startfile(file_name)  # Automatically opens the file

# Main loop for live updates
def main():
    print("Starting live updates...")
    while True:
        df = fetch_crypto_data()
        if not df.empty:
            update_excel(df)
            analysis = analyze_data(df)
            print("Excel updated and analysis generated:")
            print(analysis)
        else:
            print("Failed to fetch data. Retrying...")
        time.sleep(300)  # Update every 5 minutes

if __name__ == "__main__":
    main()
