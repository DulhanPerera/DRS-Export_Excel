# Excel Export Tool for Case Details

This Python program exports case details from a MongoDB database into an Excel file. It generates multiple tables based on the data stored in the `case_details` collection, including case details, contact details, remarks, settlements, payments, and more.

## Features

- **Export Case Details**: Export case-related data into a structured Excel file.
- **Multiple Tables**: Generate tables for:
  - Case Details
  - Contact Details
  - Remarks
  - Settlements
  - Settlement Plans
  - Approvals
  - Case Status
  - Abnormal Stops
  - Debt Recovery Company (DRC)
  - Recovery Officers (RO)
  - Payments
  - Recovery Officer Negotiations
  - Recovery Officer Requests
- **Customizable Styles**: Apply predefined styles to Excel tables for better readability.
- **Dynamic Data Handling**: Handle arrays and nested data structures in MongoDB collections.

## Prerequisites

Before running the program, ensure you have the following installed:

- **Python 3.8 or higher**
- **MongoDB**: The program connects to a MongoDB database.
- **Required Python Packages**: Install the required packages using `pip`.

## Installation

### Clone the Repository:

```bash
git clone https://github.com/your-username/excel-export-tool.git
cd excel-export-tool
```

### Set Up a Virtual Environment (Optional but recommended):

```bash
python -m venv venv
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
```

### Install Dependencies:

```bash
pip install -r requirements.txt
```

### Configure the Program:

- Update the `Config/Config.ini` file with your MongoDB connection details and file paths.
- Ensure the `case_details` collection in your MongoDB database contains the required data.

## Usage

### Run the Program:

```bash
python export.py
```

### Output:

The program will generate an Excel file in the specified output directory (`Config/Config.ini`).
The file will contain multiple sheets with tables for case details, contacts, remarks, settlements, payments, etc.

## Configuration

The program uses a `Config.ini` file for configuration. Here’s an example:

```ini
[DATABASE]
MONGO_URI = mongodb://localhost:27017/
DB_NAME = DRS

[COLLECTIONS]
CASE_DETAIL_COLLECTION = Case_details
CASE_SETTLEMENTS_COLLECTION = Case_settlements
CASE_PAYMENTS_COLLECTION = Case_payments

[LOG_FILE_PATHS]
WIN_LOG = C:\ProgramData\Logs\application.log
LIN_LOG = /var/log/application.log

[EXCEL_EXPORT_FOLDER]
WIN_DB = D:\Exports\
LIN_DB = /var/database_exports/
```

## File Structure

```bash
excel-export-tool/
├── Config/
│   ├── Config.ini
│   ├── styles.ini
│   └── logger/
│       └── loggers.ini
├── exportExcel/
│   ├── __init__.py
│   ├── case_contact_tables.py
│   ├── config_loader.py
│   ├── data_fetcher.py
│   ├── excel_styles.py
│   ├── excel_writer.py
│   ├── payments.py
│   ├── settlements_remarks.py
│   ├── table_utils.py
│   ├── approve_table.py
│   ├── drc_ro_tables.py
│   ├── payments_table.py
│   └── ro_tables.py
├── export.py
├── requirements.txt
└── README.md
```

## Dependencies

The program uses the following Python packages:

- `pymongo`: For MongoDB interactions.
- `openpyxl`: For Excel file creation and formatting.
- `configparser`: For reading configuration files.
- `logging`: For logging errors and debugging information.

You can install all dependencies using:

```bash
pip install -r requirements.txt
```

## Contributing

Contributions are welcome! If you’d like to contribute, please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bugfix.
3. Commit your changes.
4. Submit a pull request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Example Output

The program generates an Excel file with the following structure:

### Case Details Sheet

| Case ID | Incident ID | Account No. | Customer Ref | Area | BSS Arrears Amount | Current Arrears Amount | ... |
|---------|------------|-------------|--------------|------|--------------------|----------------------|-----|

### Contact Details Sheet

| Mobile | Email | Home Phone | Address |
|--------|-------|-----------|---------|
| 0771234567 | customer@example.com | 020123456 | 123, Sample Street, Colombo |

### Payments Sheet

| Payment ID | Amount | Customer ID | Purchase Date | Promotion | ... |
|------------|--------|-------------|---------------|-----------|-----|

## Support

If you encounter any issues or have questions, please open an issue on the GitHub repository.
