# Email Parser

![Preview](preview.jpg)

## Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Usage](#usage)
- [File Structure](#file-structure)
- [Author](#author)
- [License](#license)

## Features
- Reads emails from a specific mail address from the past 12 hours.
- Extracts all relevant data from the email and creates a dataframe.
- Sends the dataframe as an Excel file to staff
- Parses emails in English or Japanese.

## Requirements
- [openpyxl](https://pypi.org/project/openpyxl/)
- [pandas](https://pypi.org/project/pandas/)

## Usage
1. Install the required Python packages:

    ```bash
    pip install openpyxl pandas
    ```

2. Update the following information in the Python script:

    - Email addresses
    - passwords (If using a standard email service provider such as gmail, you may need to setup an app password for this to view your inbox)
    - imap server

3. Place the Excel file you wish to translate into the same directory as the Python script.

4. Run the script:

    ```bash
    python email_parser_for_value_build.py
    ```

5. The `latest emails.xlsx` file will have the table of data extracted from emails from the last 12 hours from the specified address.

## File Structure
- `email_parser_for_value_build.py` Python script for parsing email.
- `latest emails.xlsx`: contains tabled data

## Special Acknowledgment
I would like to Thank Aya McKinley at Value-Build Ltd. for requesting this project.

## Author
Alex McKinley

## License
This project is licensed under the [MIT License](LICENSE).
