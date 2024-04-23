from sqlalchemy import create_engine
import pandas as pd


def import_excel_to_sql(excel_path, server_name, database_name, username, password):
  """
  Imports data from an Excel file with multiple tabs into SQL Server.

  Args:
      excel_path (str): Path to the Excel file.
      server_name (str): Name of the SQL Server instance.
      database_name (str): Name of the database to import data into.
      username (str): Username for SQL Server authentication (optional).
      password (str): Password for SQL Server authentication (optional).
  """

  # Build connection string
  if username and password:
    conn_string = f"mssql+pyodbc://{username}:{password}@{server_name}/{database_name}"
  else:
    # Use Windows Authentication (if applicable)
    conn_string = f"mssql+pyodbc://DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes"

  # Create connection engine
  engine = create_engine(conn_string)

  # Read all sheets from the Excel file
  df_dict = pd.read_excel(excel_path, sheet_name=None)

  # Loop through each sheet and import data into SQL Server
  for sheet_name, df in df_dict.items():
    # Define table name (can be modified based on sheet name)
    table_name = sheet_name

    # Convert DataFrame to a SQL table creation statement
    df.to_sql(table_name, engine, index=False, if_exists='replace')

    print(f"Data from sheet '{sheet_name}' imported successfully!")


# Example usage
excel_path = "enter_excel_file_path"
server_name = "enter_server_name"
database_name = "enter_database_name"
username = "your_username"  # Optional, for SQL Server authentication
password = "your_password"  # Optional, for SQL Server authentication

import_excel_to_sql(excel_path, server_name, database_name, username, password)