
from create_data_table import fetch_data, write_data_to_excel
import pandas as pd



def test_fetch_data()-> None:
    # Test valid credentials
    df = fetch_data("admin", "admin", "localhost", 'CodeSpeedy')
    assert not df.empty, "Dataframe should not be empty"

    # Test invalid credentials
    df = fetch_data("invalid_username", "invalid_password", "localhost")
    assert df.empty, "Dataframe should be empty"

    # Test invalid host

test_fetch_data()


titles = ['name', 'type']

def test_write_data_to_excel() -> None:
    # Generate some test data
    df = pd.DataFrame([[1, "a", 3.5], [4, "b", 6.3]], columns=['a', 'b', 'c'])

    # Write the data to an Excel file
    write_data_to_excel(df, 'C:/Users/Admin/Desktop/project/python/Excel/test.xlsx',titles=titles)

    # Read the data back from the file
    df_read = pd.read_excel('C:/Users/Admin/Desktop/project/python/Excel/test.xlsx',index_col=0, skiprows=[0,1])
    print(df_read)

test_write_data_to_excel()
