import requests
import pandas as pd
import time
from datetime import datetime
import os

def fetch_crypto_data():
    
    print("\nFetching latest crypto data...")
    try:
        url = "https://api.coingecko.com/api/v3/coins/markets"
        params = {
            'vs_currency': 'usd',
            'order': 'market_cap_desc',
            'per_page': 50,
            'page': 1,
            'sparkline': False
        }
        response = requests.get(url, params=params)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Error fetching data: {e}")
        return None

def update_excel(data):
    
    if not data:
        return
    
    try:
        
        df = pd.DataFrame(data)
        df = df[[
            'name', 'symbol', 'current_price', 'market_cap',
            'total_volume', 'price_change_percentage_24h'
        ]]
        
        
        df.columns = [
            'Name', 'Symbol', 'Price (USD)', 'Market Cap',
            '24h Volume', '24h Change %'
        ]
        
        
        df['Price (USD)'] = df['Price (USD)'].round(2)
        df['24h Change %'] = df['24h Change %'].round(2)
        
        
        timestamp_df = pd.DataFrame([{
            'Name': f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        }])
        
        
        with pd.ExcelWriter('crypto_prices.xlsx') as writer:
            df.to_excel(writer, sheet_name='Live Data', index=False)
            timestamp_df.to_excel(writer, sheet_name='Live Data', startrow=len(df) + 2, index=False)
        
        print(f"Excel updated at {datetime.now().strftime('%H:%M:%S')}")
        print("\nTop 3 cryptocurrencies:")
        print(df[['Name', 'Price (USD)', '24h Change %']].head(3).to_string())
        
    except Exception as e:
        print(f"Error saving to Excel: {e}")

def main():
    print("=== Cryptocurrency Price Tracker ===")
    interval = input("Enter update interval in seconds (default 15): ").strip()
    interval = int(interval) if interval.isdigit() else 15
    
    print(f"\nTracker started! Updates every {interval} seconds")
    print("Excel file will be saved as 'crypto_prices.xlsx'")
    print("Press Ctrl+C to stop")
    
    while True:
        try:
            data = fetch_crypto_data()
            update_excel(data)
            print(f"\nNext update in {interval} seconds...")
            print("=" * 50)
            time.sleep(interval)
            
        except KeyboardInterrupt:
            print("\nStopping tracker...")
            break
        except Exception as e:
            print(f"Error occurred: {e}")
            time.sleep(interval)

if __name__ == "__main__":
    main()