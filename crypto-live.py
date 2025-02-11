import pandas as pd
import requests
import time
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
def fetch_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data")
        return []
    

def process_data(data):
    df = pd.DataFrame(data)[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
    df.columns = ["Cryptocurrency Name", "Symbol", "Current Price (USD)", "Market Capitalization", "24h Trading Volume", "24h Price Change (%)"]
    return df

def update_excel(df, filename="crypto_data.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Crypto Data"
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    wb.save(filename)
    print(f"Excel file updated: {filename}")

def analyze_data(df):
    top_5 = df.nlargest(5, "Market Capitalization")
    avg_price = df["Current Price (USD)"].mean()
    highest_change = df.loc[df["24h Price Change (%)"].idxmax()]
    lowest_change = df.loc[df["24h Price Change (%)"].idxmin()]
    
    analysis = f"""
    Top 5 Cryptocurrencies by Market Cap:
    {top_5[['Cryptocurrency Name', 'Market Capitalization']]}

    Average Price of Top 50 Cryptocurrencies: ${avg_price:.2f}

    Highest 24h Change: {highest_change['Cryptocurrency Name']} ({highest_change['24h Price Change (%)']:.2f}%)
    Lowest 24h Change: {lowest_change['Cryptocurrency Name']} ({lowest_change['24h Price Change (%)']:.2f}%)
    """
    print(analysis)
    return analysis


def main():
    while True:
        data = fetch_data()
        if data:
            df = process_data(data)
            update_excel(df)
            analysis = analyze_data(df)
            with open("crypto_analysis.txt", "w") as f:
                f.write(analysis)
        time.sleep(300)  # Update every 5 minutes

if __name__ == "__main__":
    main()
