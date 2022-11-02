from collections import defaultdict
from typing import Union, List

import pandas as pd
import openpyxl

HEADER_ROWS = range(4,8)
DIMENSIONS_SHEET_NAME = "Structure"
DIMENSIONS_CELL_RANGE = "B4:E1000"
DIMENSIONS_FIRST_ROW = 3
DIMENSIONS_COL_CAT = 0
DIMENSIONS_COL_CODE = 2
DIMENSIONS_COL_LABEL = 3


def get_header_codes_from_excel(excel_file: str):
    book = openpyxl.load_workbook(excel_file, data_only=True)
    
    codes = defaultdict(dict)
    sheet = book[DIMENSIONS_SHEET_NAME]
    cells = sheet[DIMENSIONS_CELL_RANGE]
    
    for row in cells:
        cat = row[DIMENSIONS_COL_CAT].value
        code = row[DIMENSIONS_COL_CODE].value
        label = row[DIMENSIONS_COL_LABEL].value
        if not cat:
            break

        codes[cat][code] = label

    return codes


def print_codes(excel_file: str):
    codes = get_header_codes_from_excel(excel_file)
    for k,v in codes.items():
        print("Category: ", k)
        print("---------")
        for k,v in v.items():
            print(f"{k}: {v}")
        print()


def get_data_from_excel(excel_file: str, headers: Union[tuple, List[tuple]] = None) -> pd.DataFrame:
    book = xlrd.open_workbook(excel_file, on_demand=True)
    list_of_df = []

    if isinstance(headers, tuple):
        headers = [headers]

    for sheet in book.sheet_names():
        sh = book.sheet_by_name(sheet)

        header_names = []
        header_values = []

        for row in range(6,10):
            name = sh.cell_value(row, 0)
            value = sh.cell_value(row, 1)

            if not name:
                break

            header_names.append(name)
            header_values.append(value.split(" - ")[0])

        header_names = tuple(header_names)
        header_values = tuple(header_values)

        # if we search for a specific header, skip sheet if no header matches the current sheet header
        # if we call the function with no headers, we take all sheets
        if headers:
            if not any(header_values == header for header in headers):
                continue

        # drop column GEO(L)/TIME
        # rename GEO to geo
        # melt dataframe from wide to long with columns year and value
        # add four columns for unit, hazard, waste and nace_r2
        # change columns year and geo to type category
        # set geo as index

        # find header row
        header_row = 0
        for row in range(20):
            if sh.cell_value(row, 0) == "GEO":
                header_row = row
                break

        nrows = header_row
        for row in range(header_row, 100):
            if sh.cell_value(row, 0) == "":
                nrows = row - header_row - 1
                break

        df_sheet = pd.read_excel(excel_file, sheet_name=sheet, header=header_row, nrows=nrows, na_values=":").drop(columns="GEO(L)/TIME").rename(columns={'GEO': 'geo'})
        df_sheet = df_sheet.melt(id_vars="geo", var_name="year", value_name="value")
        df_sheet = df_sheet.assign(**{x.lower():y for x,y in zip(header_names, header_values) if x})

        list_of_df.append(df_sheet)

    df = pd.concat(list_of_df)
    df.year = df.year.astype(int)
    df[df.select_dtypes("object").columns] = df.select_dtypes("object").astype("category")
    df = df.set_index("geo")

    return df
