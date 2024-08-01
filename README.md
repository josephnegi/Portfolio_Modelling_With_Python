# Portfolio Modelling with Python

This project demonstrates portfolio modelling for historical returns using the Constant Equally Weighted (CEW) and Non-Rebalanced (NRB) methods. The project leverages Python, yfinance for fetching stock data, and xlwings for automating the export of results to Excel.

## Features

- **Portfolio Modelling**: Calculate historical portfolio returns using the CEW and NRB methods.
- **Error Handling**: Rudimentary error handling for data entry and data fetching.
- **Excel Integration**: Automatically export the portfolio analysis results to Excel.

## Requirements

- Python 3.x
- pandas
- numpy
- yfinance
- xlwings

## Usage

1. Run the script:

    ```sh
    python portfolio_modelling.py
    ```

2. Enter the number of positions in your portfolio when prompted.

3. Follow the format for entering data:

    ```
    Stock ticker_<space>_no of stocks_<space>_average purchase price_<space>_Long(1) or Short(-1)
    ```

4. If an error occurs during data entry, re-enter the data for that position.

5. The script will fetch the historical data for the entered tickers and perform the CEW and NRB analyses.

6. The results will be exported to an Excel file named `Portfolio_Modelling_Py.xlsx`. If there are any errors saving the file, a timestamped filename will be used as a fallback.

## Example

Below is an example of how to enter data for three positions in the portfolio:
```
Enter data for position no. 1:
AAPL 50 150 1
Enter data for position no. 2:
MSFT 30 200 1
Enter data for position no. 3:
TSLA 20 600 -1
```

The script will then fetch the historical data for AAPL, MSFT, and TSLA, and perform the portfolio analysis.

## Error Handling

### Data Entry

- The script ensures that the entered data is in the correct format.
- If a ticker is already used, the user will be prompted to enter a different ticker.

### Data Fetching

- The script retries fetching data up to 3 times in case of an error.
- If data fetching fails after 3 attempts, the user can choose to enter a new ticker or skip the current ticker.

### File Saving

- The script attempts to save the results to an Excel file.
- If an error occurs during saving, a timestamped filename is used as a fallback.

## Example Output

The resulting Excel file will contain the following sheets:

1. **Summary**: Contains a summary of user positions.
2. **Individual Stock Data**: Each ticker has its own sheet with historical data.
3. **CEW Performance**: Contains the CEW performance analysis.
4. **NRB Performance**: Contains the NRB performance analysis.

## Contributing

Feel free to submit issues or pull requests if you have suggestions for improving the project.


## Contact

For any questions or suggestions, please contact [Joseph Negi](https://www.linkedin.com/in/joseph-negi/).


## Video Explanation

A video explanation of this project will be available soon.

[Placeholder for YouTube link]
