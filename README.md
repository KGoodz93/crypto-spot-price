# Crypto Spot Price

A webscraper which gathers information from CoinGecko and add the spot price to the relevant tab.

This has been coded to run twice a day, to make an entry in both the Open and Close column. This can be scheduled to run at the start (00:00) and end (23:59) of each day to achieve the most accurate results.

Once this is run, this will write the data to the spot-prices.xlsx file. The spot-prices.xlsx file must be in the same directory as the script by default. However, this can be changed in the script.

The data included in this web scrape include:

- Date
- Rank
- Open
- Close
- +/-
