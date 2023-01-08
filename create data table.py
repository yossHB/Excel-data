'''
pip install XlsxWriter     # Python module to write the Microsoft Excel files in XLSX format
pip install xlrd           # Python module to read data from the excel files
pip install pandas
pip install SQLAlchemy
pip install cx_Oracle

Help : 
    - https://www.sqlshack.com/python-scripts-to-format-data-in-microsoft-excel/

'''

import xlsxwriter

import pandas as pd
import cx_Oracle
import sqlalchemy
from sqlalchemy.exc import SQLAlchemyError


username, password, host = 'admin', 'admin', 'localhost'
filepath = 'C:/Users/Admin/Desktop/project/python/Excel/Welocme.xlsx'
DBNAme = 'Resources table'
titles = ['name', 'sources','type', 'transformations', 'RQ' ]

def fetch_data(username, password, host):
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
        engine = sqlalchemy.create_engine(f"oracle+cx_oracle://{username}:{password}@{host}")
        user_submitted_comments_sql = """SELECT * FROM USER_SUBMITTED_COMMENTS"""
        df_user_submitted_comments = pd.read_sql(user_submitted_comments_sql, engine)
        return df_user_submitted_comments
    except sqlalchemy.exc.SQLAlchemyError as e:
        print(e)
        return pd.DataFrame()


df_user_submitted_comments = fetch_data(username, password, host)

def write_data_to_excel(df, filepath, DBNAme, titles):
    """
    Write data from a dataframe to an Excel spreadsheet.

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
            worksheet.write(row, 2, str(data))



write_data_to_excel(df_user_submitted_comments, filepath, DBNAme, titles)






