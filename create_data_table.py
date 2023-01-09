'''
pip install XlsxWriter     # Python module to write the Microsoft Excel files in XLSX format
pip install xlrd           # Python module to read data from the excel files
pip install pandas
pip install SQLAlchemy     # SQLAlchemy is a library that facilitates the communication between Python programs and databases. Most of the times, this library is used as an Object Relational Mapper (ORM) tool that translates Python classes to tables on relational databases and automatically converts function calls to SQL statements.
pip install cx_Oracle
pip install openpyxl for tasting part
pip install python-dotenv

- pip3 freeze command will tell us the modules installed with their versions.
    * pip3 freeze command will tell us the modules installed with their versions.

Help :
    - https://www.sqlshack.com/python-scripts-to-format-data-in-microsoft-excel/
    - https://www.geeksforgeeks.org/oracle-database-connection-in-python/

'''
from dotenv.main import load_dotenv
import os

import xlsxwriter
from decouple import config
import pandas as pd
import cx_Oracle
import sqlalchemy
from sqlalchemy.exc import SQLAlchemyError

import ads
from sqlalchemy import text

load_dotenv()
username, password, host  = os.environ['username'], os.environ['password'], os.environ['host']

filepath = os.environ['filepath']
DBNAme = os.environ['DBNAme']
titles = os.environ['titles']
table_name = os.environ['table_name']


def fetch_data(username, password, host, table_name=table_name):
    """
    Fetch data from an Oracle database using the provided credentials and host.

    Parameters:
        username (str): The username to use when connecting to the database.
        password (str): The password to use when connecting to the database.
        host (str): The hostname or IP address of the machine hosting the database.

    Returns:
        pd.DataFrame: A dataframe containing the results of the query.
    """
    try:
        engine = sqlalchemy.create_engine(f"oracle+cx_oracle://{username}:{password}@{host}/xe")
        user_submitted_comments_sql = f"""SELECT * FROM {table_name}"""
        df_user_submitted_comments = pd.read_sql(user_submitted_comments_sql, con = engine)
        return df_user_submitted_comments

    except sqlalchemy.exc.SQLAlchemyError as e:
        print(e)
        return pd.DataFrame()


df_user_submitted_comments = fetch_data(username, password, host)

def write_data_to_excel(df, filepath = filepath, table_name = table_name, DBNAme = DBNAme, titles = titles):
    """
    Write the name and the type of columns from a dataframe to an Excel spreadsheet.

    Parameters:
        df (pd.DataFrame): The dataframe to write to the spreadsheet.
        filepath (str): The file path of the Excel spreadsheet.
    """
    with xlsxwriter.Workbook(filepath) as workbook:
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({'bold': True})
        cell_format.set_font_size(14)
        cell_format.set_align('center')

        worksheet.set_column('A:E', 50)

        worksheet.write(0, 0, DBNAme, cell_format)
        for row, title in enumerate(titles):
            worksheet.write(1, row, title, cell_format)

        for row, data in enumerate(df.columns, start=2):
            worksheet.write(row, 0, data)

        for row, data in enumerate(df.dtypes, start=2):
            worksheet.write(row, 1, str(data))

        for row, data in enumerate(df, start=2):
            worksheet.write(row, 2, table_name)


write_data_to_excel(df_user_submitted_comments, filepath, table_name, DBNAme, titles)
print(df_user_submitted_comments)





