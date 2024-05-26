import argparse
import csv

import xlsxwriter

asset_folder = "./assets"

parser = argparse.ArgumentParser()

parser.add_argument("input_file", type=str)
parser.add_argument("output_file", type=str)
parser.add_argument("separator", type=str, default=";")

args = parser.parse_args()

data = []

with open(f"{asset_folder}/in/{args.input_file}", "r") as csv_file:
    reader = csv.reader(csv_file, delimiter=args.separator)

    workbook = xlsxwriter.Workbook(f"{asset_folder}/out/{args.output_file}")
    worksheet = workbook.add_worksheet()

    for row_index, row in enumerate(reader):
        worksheet.write_row(row_index, 0, row)
        data.append(row)

    workbook.close()

### python 4.py 45.csv 4.xlsx , 