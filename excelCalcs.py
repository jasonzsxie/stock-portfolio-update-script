import pandas as pd
import yfinance as yf
from pathlib import Path
import pymupdf 
import re

ROOT_DIR = Path(__file__).parent
# Read the Excel file
file_path = ROOT_DIR / 'currentPortfolio.xlsx'#MODIFY THE FILE NAME HERE AS NEEDED
df = pd.read_excel(file_path)

# Debug: Print the initial DataFrame
print("Initial DataFrame:")
print(df)

#I WANT TO EXTRACT THE MORNINGSTAR DATA AND ONLY HAVE THE TICKER SYMBOL AND SHARES HELD
ticker_and_shares = df[['Ticker', 'Shares\nHeld']]

#extra useless row that needs removal
ticker_and_shares_filtered = ticker_and_shares.dropna().iloc[:-1]
newPath = ROOT_DIR / 'modifiedData.xlsx'
ticker_and_shares_filtered.to_excel(newPath, index=False)



#NOW EXTRACT FUTURE ALLOCATION WITHIN THE DOC
pdf_path = ROOT_DIR / 'data.pdf ' #RENAME AS FIT
pdf_document = pymupdf.open(pdf_path)
start_page = pdf_document.page_count - 2 #change this based on number of pages (IF THE CURRENT PORTFOLIO IS ON THE LAST PAGE ONLY, SET TO 1. IF LAST 2 PAGES, SET TO 2)
end_page = pdf_document.page_count - 1
def extract_and_clean_text_from_pages(pdf_path, start_page, end_page):
    text_content = ""
    with pymupdf.open(pdf_path) as pdf_document:
        for page_num in range(start_page, end_page + 1):
            page = pdf_document[page_num]
            text_content += page.get_text()
    
    # Clean up random spaces in the extracted text
    cleaned_text = " ".join(text_content.split())
    return cleaned_text
cleaned_text_content = extract_and_clean_text_from_pages(pdf_path, start_page, end_page)
print(cleaned_text_content)
# Define a function to parse the text and extract ticker symbols and current allocations
def extract_ticker_allocations(text):
    # Regular expression to match the relevant lines in the table
    pattern = re.compile(r'([A-Za-z\s&]+)\s+([A-Z]+)\s+[\d.]+%\s+([\d.]+%)')
    
    # Find all matches
    matches = pattern.findall(text)
    
    # Extract ticker symbols and current allocations into a list
    ticker_allocations = [(match[1], match[2]) for match in matches]
    
    return ticker_allocations

# Define a function to clean and validate ticker symbols
def clean_ticker_symbols(text):
    # Pattern to identify ticker symbols with incorrect spaces (e.g., "QL YS" instead of "QLYS")
    pattern = re.compile(r'([A-Z])\s+([A-Z])')
    
    # Clean text by removing spaces between characters in ticker symbols
    cleaned_text = pattern.sub(r'\1\2', text)
    
    return cleaned_text

# Clean the specific known issues in the cleaned text
final_corrected_text_content = clean_ticker_symbols(cleaned_text_content)

# Extract ticker symbols and current allocations from the final corrected text
ticker_allocations = extract_ticker_allocations(final_corrected_text_content)

# Display the final corrected extracted list
print(ticker_allocations)
allocation_dict = dict(ticker_allocations)

df = pd.read_excel(newPath)
df['Future Distribution'] = df['Ticker'].map(allocation_dict)
for ticker, percent in allocation_dict.items():
    if ticker not in df['Ticker'].values:
        new_row = pd.DataFrame({'Ticker': [ticker], 'Shares\nHeld': [0], 'Future Distribution': [percent]})
        df = pd.concat([df, new_row])
output_path = ROOT_DIR / 'modifiedData.xlsx'
df.to_excel(output_path, index=False)

#NOW I WANT TO CLEAN THE EXCEL SHEET BEFORE I DO ANY MORE
df = pd.read_excel(output_path)
df['Future Distribution'] = df['Future Distribution'].str.rstrip('%').astype('float') / 100.0
df_filtered = df[(df['Shares\nHeld'] != 0) | (df['Future Distribution'].notna() & (df['Future Distribution'] != '0.0%'))]
cleanedData = ROOT_DIR / 'modifiedData.xlsx'
df_filtered.to_excel(cleanedData, index=False)





df = pd.read_excel(cleanedData)

# Define a function to get the closing price of a stock
def get_closing_price(ticker):
    stock = yf.Ticker(ticker)
    hist = stock.history(period="1d")
    closing_price = hist['Close'].iloc[0]
    return closing_price

# Add a new column for the closing prices
df['Closing Price'] = df['Ticker'].apply(get_closing_price)

# Debug: Print the DataFrame after fetching closing prices
print("\nDataFrame after fetching closing prices:")
print(df)

# Ensure 'Future Distribution' is treated as float directly, as it already represents a percentage
# No need to convert or strip percentage signs since it's read correctly as a float in decimal format

# Calculate the current value of each stock in the portfolio
df['Current Value'] = df['Shares\nHeld'] * df['Closing Price']

# Calculate the total current value of the portfolio
total_portfolio_value = df['Current Value'].sum()

# Debug: Print the total portfolio value
print("\nTotal Portfolio Value:", total_portfolio_value)

# Calculate the target value for each stock based on the new percentages
df['Target Value'] = df['Future Distribution'] * total_portfolio_value

# Calculate the target number of shares
df['Target Shares'] = round(df['Target Value'] / df['Closing Price'])

# Calculate the number of shares to buy or sell
df['Shares to Buy/Sell'] = df['Target Shares'] - df['Shares\nHeld']

df['Real Distribution'] = df['Target Shares'] * df["Closing Price"] /total_portfolio_value

# Debug: Print the final DataFrame with all calculations
print("\nFinal DataFrame with all calculations:")
print(df)

print("\nSorted")
df = df.sort_values(by='Real Distribution', ascending=False)
df['Real Distribution'] = (df['Real Distribution'] * 100).round(1).astype(str) + '%'
print(df)

# Save the updated dataframe to a new Excel file
output_file_path = ROOT_DIR /'updated_portfolio.xlsx'
df.to_excel(output_file_path, index=False)

#Look at updated_portfolio.xlsx and see if there is any issues that need to be fixed then go to buy/sell order file and execute it to create the orders
print("Updated portfolio saved to", output_file_path)

df['Closing Price (USD)'] = df['Closing Price'].apply(lambda x: f"${x:,.2f}")
df['Current Value (USD)'] = df['Current Value'].apply(lambda x: f"${x:,.2f}")
df['Target Value (USD)'] = df['Target Value'].apply(lambda x: f"${x:,.2f}")

# Add a summary row with total values
summary_row = pd.Series({
    'Ticker': 'Total',
    'Shares\nHeld': '',
    'Future Distribution': '',
    'Closing Price (USD)': '',
    'Current Value (USD)': f"${df['Current Value'].sum():,.2f}",
    'Target Value (USD)': f"${df['Target Value'].sum():,.2f}",
    'Target Shares': '',
    'Shares to Buy/Sell': '',
    'Real Distribution': ''
})
summary_df = pd.DataFrame([summary_row])
df = pd.concat([df, summary_df], ignore_index = True)

print(df)

# Save the updated dataframe with USD values and summary row to a new Excel file
final_output_path = ROOT_DIR / 'updated_portfolio_with_summary.xlsx'
df.to_excel(final_output_path, index=False)

print("Updated portfolio with summary saved to", final_output_path)