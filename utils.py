from collections import defaultdict
from typing import Union, List

import pandas as pd
import xlrd

header_rows = range(6,10)


def get_header_codes_from_excel(excel_file: str):
    book = xlrd.open_workbook(excel_file, on_demand=True)
    header_rows = range(6,10)

    header = defaultdict(set)
    for sheet in book.sheet_names():
        sh = book.sheet_by_name(sheet)

        for row in header_rows:
            label, value = sh.cell_value(row, 0), sh.cell_value(row, 1)
            if not label:
                break

            header[label].add(value)

    codes = {k: {s.split(" - ")[0]: s.split(" - ", 1)[1] for s in v} for k,v in header.items()}
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
