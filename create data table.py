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

DBNAme = 'Resources table'
titles = ['name', 'sources','type', 'transformations', 'RQ' ]

try:

   engine = sqlalchemy.create_engine("oracle+cx_oracle://admin:admin@localhost")
   user_submitted_comments_sql = """SELECT * FROM USER_SUBMITTED_COMMENTS""";
   df_user_submitted_comments = pd.read_sql(user_submitted_comments_sql, engine)

   print(df_user_submitted_comments)
   print("success")
   print(df_user_submitted_comments.dtypes)
except SQLAlchemyError as e:
   print(e)

with xlsxwriter.Workbook('C:/Users/Admin/Desktop/project/python/Excel/Welocme.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    cell_format = workbook.add_format({'bold': True})
    cell_format.set_font_size(16)
    cell_format.set_align('center')

    worksheet.set_column('A:E', 50)

    worksheet.write(0,0,DBNAme,cell_format)
    for row, title in enumerate(titles):
        worksheet.write(1,row, title, cell_format)

    if not df_user_submitted_comments.empty:
        for row, data in enumerate(df_user_submitted_comments.columns, start=2):
            worksheet.write(row, 0, data)

        for row, data in enumerate(df_user_submitted_comments.dtypes, start=2):
            worksheet.write(row, 2, str(data))
    else:
        # Write a message to the worksheet indicating that no data was returned
        worksheet.write(2, 0, "No data returned")










