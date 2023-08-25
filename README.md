# Apriori and Basket Penetration Analysis Tool

## Description

This project is a comprehensive tool for running Apriori and Basket Penetration analyses on sales data. It is designed to help businesses understand customer purchasing behavior and identify patterns in product purchases. The tool is built using Python and features a user-friendly GUI using Tkinter.

## Features

- Apriori Analysis: Identify frequent item sets and association rules in your sales data.
- Basket Penetration Analysis: Understand the penetration of products in customer baskets.
- Date Range Selection: Choose a custom date range for the analysis.
- Level of Analysis: Choose between SKU, Category, and SubCategory levels for your analysis.
- Progress Bar: Real-time progress updates.
- Export Results: Save the analysis results to CSV files.

## Installation

### Prerequisites

- Python 3.x
- Tkinter
- Pandas
- Numpy
- mlxtend
- pyodbc

### Steps

1. Clone the repository:

    ```bash
    git clone https://github.com/your-username/your-repo-name.git
    ```

2. Navigate to the project directory:

    ```bash
    cd your-repo-name
    ```

3. Install the required packages:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Run the main script:

    ```bash
    python main.py
    ```

2. The GUI will appear. Select the level of analysis, date range, and other parameters.

3. Click on "Run Apriori Analysis" or "Run Basket Analysis" to start the analysis.

4. Once the analysis is complete, the results will be saved as CSV files in the project directory.

## Contributing

If you'd like to contribute, please fork the repository and make changes as you'd like. Pull requests are warmly welcome.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

## Acknowledgments

- [mlxtend](http://rasbt.github.io/mlxtend/) for the Apriori and association rules algorithms.
- [Tkinter](https://docs.python.org/3/library/tkinter.html) for the GUI.
