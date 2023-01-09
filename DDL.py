import pandas as pd
import cx_Oracle
import sqlalchemy
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.sql import text


import cx_Oracle


try:
    ''' to create a table in Oracle database
    with cx_Oracle.connect('admin/admin')as co:
        print("Connected")
        cur=co.cursor()
        cur.execute("""
            CREATE TABLE
            CodeSpeedy(employee_id number(10),employee_name varchar2(10))""")
        print("Table Created")
        cur.close()
    '''
    ''' to insert data into a table in Oracle database
    with cx_Oracle.connect('admin/admin')as co:
        print("Connected")
        cur=co.cursor()
        cur.execute("""INSERT INTO
                       CodeSpeedy values(101,'Ravi')""")
        cur.execute("""INSERT INTO
                       CodeSpeedy values(102,'Ramu')""")
        cur.execute("""INSERT INTO
                       CodeSpeedy values(103,'Rafi')""")
        co.commit()
        cur.close()
        print("Data Inserted")
    '''
    ''' to fetch data from a table in Oracle database
    
    with cx_Oracle.connect('admin/admin')as co:
        print("Connected")
        cur=co.cursor()
        
        cur.execute("SELECT * FROM CodeSpeedy")
        x=cur.fetchone()
        print(x)
        cur.execute("SELECT * FROM CodeSpeedy")
        y=cur.fetchmany(2)
        print(y)
        
        cur.execute("SELECT * FROM CodeSpeedy")
        z=cur.fetchall()
        print(z)
        
        print("Data Fetched")
        cur.close()
    '''
    ''' to Connect to the database using sqlalchemy lib
    engine = sqlalchemy.create_engine("oracle+cx_oracle://admin:admin@localhost")

    # Execute a SQL query and fetch the results
    sql = "SELECT * FROM CodeSpeedy"
    df = pd.read_sql(sql, engine)

    # Print the results
    print(df)
    '''

    print('conn')
    
except sqlalchemy.exc.SQLAlchemyError as e:
    print(e)