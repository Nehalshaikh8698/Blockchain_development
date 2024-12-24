import requests
import pandas as pd
from apscheduler.schedulers.blocking import BlockingScheduler

# here we fetch data from CoinGecko API
def fetch_crypto_data():
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
        print("Failed to fetch data from API")
        return []

# here is the process and save data to EXCLE
def process_and_save_to_excel():
    data = fetch_crypto_data()
    if not data:
        return

    # here we are Creating a DataFrame
    df = pd.DataFrame(data, columns=[
        "name", "symbol", "current_price", "market_cap",
        "total_volume", "price_change_percentage_24h"
    ])
    df.rename(columns={
        "name": "Cryptocurrency Name",
        "symbol": "Symbol",
        "current_price": "Current Price (USD)",
        "market_cap": "Market Capitalization",
        "total_volume": "24-hour Trading Volume",
        "price_change_percentage_24h": "Price Change (24-hour %)"
    }, inplace=True)

    # Perform all the mathmatical operations here
    top_5 = df.nlargest(5, "Market Capitalization")
    average_price = df["Current Price (USD)"].mean()
    highest_change = df["Price Change (24-hour %)"].max()
    lowest_change = df["Price Change (24-hour %)"].min()

    # Print above operations here
    print("Top 5 Cryptocurrencies by Market Cap:")
    print(top_5[["Cryptocurrency Name", "Market Capitalization"]])
    print(f"\nAverage Price of Top 50 Cryptocurrencies: ${average_price:.2f}")
    print(f"Highest 24-hour Percentage Change: {highest_change:.2f}%")
    print(f"Lowest 24-hour Percentage Change: {lowest_change:.2f}%")

    # here we Save data to Excel
    with pd.ExcelWriter("crypto_data.xlsx", engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name="Live Data")
        top_5.to_excel(writer, index=False, sheet_name="Top 5 Cryptos")

    print("Data saved to 'crypto_data.xlsx'")

# Schedule the script to run every 5 minutes
scheduler = BlockingScheduler()
scheduler.add_job(process_and_save_to_excel, "interval", minutes=5)

print("Script is running... Press Ctrl+C to stop.")
try:
    scheduler.start()
except (KeyboardInterrupt, SystemExit):
    print("Script stopped.")