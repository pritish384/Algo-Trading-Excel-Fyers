# Algo-Trading-Excel-Fyers
Algo-Trading-Excel-Fyers is an algorithmic trading tool that integrates with the Fyers broker. This tool allows you to execute orders based on Excel formulas, enabling you to easily implement strategies like strangles, straddles, and many others.

## Features

- **Order Book**: Track all your executed orders within the tool.
- **Excel Integration**: Utilize Excel formulas to create custom trading strategies.
- **Flexibility**: Create and execute various trading strategies such as strangles, straddles, and more by simply defining the symbols and order parameters in Excel.

## How It Works

1. **Integration**: Connect the tool with your Fyers account using an API token and secret.
2. **Order Execution**: Define trading symbols and relevant order information in Excel. When the signal cell in Excel turns `TRUE`, the order is automatically executed.
3. **Customization**: Customize the strategy by modifying the Excel formulas as per your trading plan.

## Getting Started

To start using the Algo-Trading-Excel-Fyers tool, follow these steps:

### Prerequisites

- **Fyers Account**: You must have an account with Fyers.
- **Python**: Install Python on your system.
- **Excel**: Ensure that you have Excel installed on your system.

### Setup

1. **Generate Fyers API Credentials**:
   - Log in to your Fyers account.
   - Navigate to the [Fyers Developer Portal](https://myapi.fyers.in/).
   - Navigate to dashboard and create an app to get your API token and secret.
   - **Note:** You will need to provide the redirect URL as `http://127.0.0.1:5000`
   - Create an app to get your API token and secret.

2. **Setup Code**:
    - Clone this repository.
    - Install the required Python packages using the following command:
      ```bash
      pip install -r requirements.txt
      ```
    - Run `main.py` file
    - Enter your API token and secret when prompted.
    - Login to your Fyers account using the browser that opens up.
    - Once you are logged in, close the browser window.
    - The tool is now connected to your Fyers account.
    - **Note:** You will need to run the `main.py` file every time you want to use the tool.
    - If any error occurs, please run `reset.py` file to reset the tool.



3. **Set Up Your Strategy in Excel**:
    - Run the `main.py` file.
    - Open the Excel sheet named `Algo-Trading-Excel-Fyers.xlsx`.
    - Define the symbols and order parameters in the Excel sheet.
    - All the Live data will be updated in the Excel sheet.
    - Put down all your conditions in the Excel sheet.
    - When the signal cell turns `TRUE`, the order will be executed.
    - The order execution status will be displayed in the `Order Book` sheet.

4. **Customize Your Settings**:
    - You can customize the tool by modifying the value in the `settings` sheet.
    - Number of Rows: The number of rows to scan for signals. (Default: 10)
    - Time Delay: The time delay between each scan. (Default: 1 seconds)

    ### Note: More number of rows will increase load on the system. You can also increase the time delay to reduce the load. Low time delay will increase the load on the system but it will provide live data.


    ## Note: The terminal should be running while you are working on the Excel sheet.

## Contribution

Contributions are welcome! Open issues or pull requests for enhancements or bug fixes.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
