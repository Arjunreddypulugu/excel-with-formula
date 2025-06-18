# Excel Formula Processor

A Streamlit application that processes Excel files containing equipment and spare parts data. The application connects to a SQL Server database to enrich the data with equipment information and reorganizes it based on equipment types.

## Features

- Processes multiple Excel sheets
- Connects to SQL Server for equipment data
- Reorganizes data by equipment type and parts
- Implements smart spare quantity calculation based on number of machines
- Handles column matching and data validation
- Outputs formatted Excel files

## Requirements

- Python 3.x
- Streamlit
- Pandas
- SQLAlchemy
- Openpyxl

## Installation

1. Clone the repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the Streamlit app:
```bash
streamlit run app.py
```

2. Upload your Excel file through the web interface
3. Download the processed Excel file

## Configuration

The application requires SQL Server connection details to be configured in Streamlit secrets. Create a `.streamlit/secrets.toml` file with the following structure:

```toml
[mssql]
connection_string = "your_connection_string_here"
``` 