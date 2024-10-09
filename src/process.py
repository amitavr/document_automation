import pandas as pd
import openpyxl
from enum import Enum

# Class represent columns names to avoid using hardcoded Strings.
class ColumnNames(Enum):
    COLUMN_ONE = 'COLUMN_ONE_NAME'
    COLUMN_TWO = 'COLUMN_TWO_NAME'
    COLUMN_THREE = 'COLUMN_THREE_NAME'

# This method reads data from an Excel file and returns it as a DataFrame
def getDataFromExcel (path, sheetName) -> pd.DataFrame:
    df = pd.read_excel(f"{path}", sheet_name=sheetName)
    return df

# This Method recieve a DataFrame and validate it.
def validate_df(dataframe : pd.DataFrame) -> pd.DataFrame:
    # Ensure the dataframes have the necessary columns
    required_columns = [ColumnNames.COLUMN_ONE.value, ColumnNames.COLUMN_TWO.value, ColumnNames.COLUMN_THREE.value]
    for col in required_columns:
        if col not in dataframe.columns:
            raise ValueError(f"dataframe must contain the column: {col}")
    return dataframe

# This method recieve DataFrames and then create an excel file in below directory
def write_to(first_file : pd.DataFrame ,processed_file : pd.DataFrame, file1_name : str) -> None:
    with pd.ExcelWriter('src\\document_automation\\processed\\output.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        first_file.to_excel(writer, index=False, sheet_name=file1_name)
        processed_file.to_excel(writer, index=False, sheet_name=f'Processed {file1_name}')

def process_file(dataframe : pd.DataFrame) -> pd.DataFrame:
    pass

# The method recieve file paths, then processing all necessary methods.
def processing(file_path1 : str, file1_name : str) -> None:
    first_file = getDataFromExcel(f'src\\document_automation\\uploads\\{file_path1}', 'Sheet Name')
    processed_file = process_file(first_file)
    write_to(first_file=first_file, processed_file = processed_file ,file1_name=file1_name)






