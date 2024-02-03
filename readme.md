# back story
I needed to analyze a bunch of spreadsheets. The horror of opening each sheet individually encouraged me to write down each step. See prompt.

The first output from GPT4 kind of worked but looped endlessly.

I edited the prompt, then passed the prompt to a new instance of GPT4.

The second prompt with edits worked more or less out of the box!

10 minutes of thinking saved me an hour of mind numbing copy and paste!




# XLSX Data Extractor

This Python script is designed for data scientists and anyone who needs to extract and analyze data from multiple `.xlsx` files. It provides a user interface to navigate to the location of `.xlsx` files, select multiple files, and perform specific data analysis tasks.

## Features

- **File Navigation**: A simple UI to navigate and select `.xlsx` files.
- **Multiple File Selection**: Users can select multiple `.xlsx` files for analysis.
- **Sheet Handling**: The script loads the `.xlsx` files, opens them, and prints the names of each sheet, ignoring sheets named "Commission".
- **Sheet Selection**: Users can select specific sheets by number or choose all sheets for analysis.
- **Data Analysis**: The script analyzes the column names and sums the values of the "No of Nights" column in each selected sheet.
- **Results Notification**: Users are notified when the sums are completed.
- **Console Output**: The name of the sheet and the sum are printed to the console.
- **Save Results**: Users can prompt for a file name to save the results as a CSV file, including the sheet name, column name, and sum.

## Usage

1. Run the script ```main.py```
2. Use the UI to navigate to the `.xlsx` files you wish to analyze.
3. Select the files you want to work with.
4. The script will load the files and list the sheets, excluding any named "Commission".
5. When prompted, input the integers corresponding to the sheets you want to analyze, or enter "all" or "a" to select all sheets.
6. The script will then analyze the selected sheets and calculate the sums.
7. Once the analysis is complete, the results will be printed to the console.
8. You will be prompted to enter a file name to save the results as a CSV file.

## Requirements

- Python 3.x
- `openpyxl` library for handling `.xlsx` files.

## Installation

To set up the script, you need to have Python installed on your system. If you don't have Python installed, download and install it from [python.org](https://www.python.org/).

After installing Python, install the required `openpyxl` library by running the following command:


## Contributing

Contributions to this script are welcome. Please fork the repository, make your changes, and submit a pull request.

## License

This script is released under the MIT License. See the LICENSE file for more details.