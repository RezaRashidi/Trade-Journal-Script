# Trading Journal Script

## Description
This Python script is designed to generate a comprehensive Excel-based trading journal. It helps traders meticulously record their trades, track performance, and analyze various aspects of their trading strategy. The script automates the creation of a structured journal with pre-defined columns, formulas, and conditional formatting to facilitate detailed analysis of trading activities.

## Features
- **Automated Excel Generation**: Creates a new `.xlsx` file with a pre-defined structure for daily, weekly, and monthly trading logs.
- **Detailed Trade Tracking**: Includes columns for entry/exit times, signals, status, momentum, order types, trade types, reasons, plan adherence, risk management, and trade results.
- **Performance Metrics**: Automatically calculates trade duration, result in R-multiples, cumulative results, balance, max reward balance, min/max balance, and drawdown percentages.
- **Conditional Formatting**: Applies visual cues (e.g., green/red fills) based on trade performance and adherence to plan.
- **Data Validation**: Provides dropdowns for various fields (e.g., Signal, Status, Momentum, Order Type, Trade Type, Plan Adherence) to ensure consistent data entry.
- **OS-Agnostic File Opening**: Attempts to open the generated Excel file automatically upon completion.

## Installation

### Prerequisites
- Python 3.x (It is recommended to use Python 3.8 or higher)

### Dependencies
This script requires the `openpyxl` library. You can install it using pip:
```bash
pip install openpyxl
```

### Steps
1. **Clone the repository (if applicable) or download the `Journal.py` file:**
   ```bash
   git clone https://github.com/your-username/trading-journal-script.git
   cd trading-journal-script
   ```
   (If you just have the `Journal.py` file, simply place it in your desired directory.)

2. **Install dependencies:**
   ```bash
   pip install openpyxl
   ```

## Usage

To run the trading journal script, navigate to the directory where `Journal.py` is located in your terminal and execute the following command:

```bash
python Journal.py
```

Upon running, the script will prompt you to enter the following information:
- **Start Date (YYYY-MM-DD)**: The beginning date for your journal. (Default: `2025-04-28`)
- **Initial Capital**: Your starting capital for trading. (Default: `25000`)
- **Number of Weeks (1-4)**: The duration for which the journal will be generated. (Default: `4`)

Example interaction:
```
Enter start date (YYYY-MM-DD) or press Enter for default (2025-04-28): 2025-06-01
Enter initial capital or press Enter for default (25000): 10000
Enter number of weeks (1-4) or press Enter for default (4): 2
```

After providing the inputs, the script will generate an Excel file named `trading_journal_YYYY-MM-DD_to_YYYY-MM-DD.xlsx` (e.g., `trading_journal_2025-06-01_to_2025-06-15.xlsx`) in an `output` directory within the script's location. The file will then attempt to open automatically.

## Excel File Columns

The generated Excel file includes the following columns:

| Column Header               | Description                                                                 |
| :-------------------------- | :-------------------------------------------------------------------------- |
| **Time Slot**               | Predefined time intervals for trade entries (e.g., 18:30, 19:30).           |
| **Entry Time**              | The exact time a trade was entered.                                         |
| **Exit Time**               | The exact time a trade was exited.                                          |
| **Signal**                  | Strength of the trading signal (Weak, Strong, Normal, V-Weak, V-Strong, V-Normal). |
| **Status**                  | Market condition at the time of trade (Range, Trend, Pullback, Undefined/Transition, V-Range, V-Trend, V-Pullback, V-Undefined/Transition). |
| **Momentum**                | Momentum of the market (Decreasing, Increasing, V-Decreasing, V-Increasing). |
| **Enter type**              | Type of order used for entry (e.g., 2R (Normal), 4R (Aggressive Stop)).     |
| **Trade Type**              | Classification of the trade (Reversal, Continuation).                       |
| **Reason / Flex / Mistakes**| Detailed notes on the trade's rationale, flexibility, or any errors made.   |
| **Plan Adherence**          | Assessment of adherence to the trading plan (Full Adherence, Entry Flaw, Exit/Flexibility Flaw, High Target Flaw, Fear of entering). |
| **Stop Loss (%)**           | Percentage of capital risked per trade for stop loss.                       |
| **Risk (%)**                | The percentage of initial capital risked on the trade.                      |
| **Target Reward (R)**       | The target reward in R-multiples (Risk multiples).                          |
| **Trade Duration (min)**    | Duration of the trade in minutes.                                           |
| **Max Reward (R)**          | The maximum potential reward achieved during the trade in R-multiples.      |
| **Result (R)**              | The actual result of the trade in R-multiples.                              |
| **Cumulative Result (R)**   | Running total of R-multiples.                                               |
| **Cumulative Max Reward (R)**| Running total of maximum potential R-multiples.                             |
| **Balance**                 | Account balance after each trade.                                           |
| **Max Reward Balance**      | Account balance if max reward was achieved.                                 |
| **Min Balance**             | Minimum account balance reached.                                            |
| **Max Balance**             | Maximum account balance reached.                                            |
| **Min Lows % Drawdown**     | Percentage drawdown from initial capital based on minimum balance.          |
| **Trade Screenshot**        | Space for embedding trade screenshots.                                      |

## Contributing
Contributions are welcome! If you have suggestions for improvements, new features, or bug fixes, please feel free to:
1. Fork the repository.
2. Create a new branch (`git checkout -b feature/YourFeature`).
3. Make your changes.
4. Commit your changes (`git commit -m 'Add some feature'`).
5. Push to the branch (`git push origin feature/YourFeature`).
6. Open a Pull Request.

## License
This project is licensed under the MIT License - see the LICENSE file for details (if applicable, otherwise state "No specific license is applied to this script.").

## Contact
For any questions or feedback, please open an issue on the GitHub repository or contact [Your Name/Email/GitHub Profile Link].
