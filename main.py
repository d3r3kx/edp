"""
Usage: python main.py -i Book1.xlsx -o output.xlsx, where -i specifies the input path of excel files, and -o specifies
the output path of the processed results
"""
import argparse

import pandas as pd


def main(input_path, output_path):
    print(f"reading from {input_path}")
    # specify which sheet to read from, 0 as the default value if sheet_name is not specified\
    # other values, like "Sheet1", are also supported
    # refer to: https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_excel.html#pandas-read-excel
    df = pd.read_excel(input_path, sheet_name=0)

    print(df.describe())  # display a summary of the dataframe, including count, mean and other stats

    # for all rows in the dataframe table, index_m is an integer, indicating the row number, and row_m is the data
    # values of the m-th row
    for index_m, row_m in df.iterrows():
        for index_n, row_n in df.iterrows():
            # if b(n) <= a(m) <= c(n)
            if row_n["B_DATA"] <= row_m['A_DATA'] <= row_n["C_DATA"]:
                # copy the values of c(n) and d(n) to e(m) and f(m)
                row_m["E_DATA"] = row_n["C_DATA"]
                row_m["F_DATA"] = row_n["D_DATA"]

    # write the updated values to the sheet named "Sheet1" of the specified output file path
    df.to_excel(excel_writer=output_path, sheet_name="Sheet1", index=False, engine="openpyxl")
    print(f"written to {output_path}")


if __name__ == '__main__':
    # argument parser to help manage multiple parameters and make it easy to call the program in command line
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", type=str, help="path of input file: xls, xlsx")
    parser.add_argument("-o", type=str, help="path of output file")

    # Usage: python main.py -i Book1.xlsx -o output.xlsx
    args = parser.parse_args()
    print(args.i, args.o)
    main(args.i, args.o)
