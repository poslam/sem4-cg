import csv

import xlsxwriter

asset_folder = "./assets"

data = []

with open(f"{asset_folder}/in/45.csv", "r") as csv_file:
    reader = csv.reader(csv_file, delimiter=",")
    for row_index, row in enumerate(reader):
        data.append(row)

workbook = xlsxwriter.Workbook(f"{asset_folder}/out/5.xlsx")

sheets = [workbook.add_worksheet(group) for group in sorted({x[0] for x in data})]

for sheet in sheets:
    row_counter = 0
    grades = {2: 0, 3: 0, 4: 0, 5: 0}
    blocks = {x: 0 for x in range(1, 8)}

    for i in range(len(data)):
        if data[i][0] == sheet.name:
            rating = sum(float(x.replace(",", ".")) for x in data[i][2:]) / len(
                data[i][2:]
            )
            mark = (
                5
                if rating >= 80
                else (4 if rating >= 70 else (3 if rating >= 50 else 2))
            )
            data[i].extend([rating, mark])
            sheet.write_row(row_counter, 0, data[i])
            row_counter += 1

            if mark in grades:
                grades[mark] += 1

            for x in range(1, 8):
                val = data[i][x + 1]
                blocks[x] += (
                    float(val.replace(",", ".")) if isinstance(val, str) else val
                )

    row_counter_for_analytics = 0

    for i, grade in enumerate(grades.items()):
        sheet.write(i, 12, grade[0])
        sheet.write(i, 13, grade[1])
        row_counter_for_analytics += 1

    for i, block in enumerate(blocks.items(), start=row_counter_for_analytics + 1):
        sheet.write(i, 12, block[0])
        sheet.write(i, 13, block[1] / row_counter)

    mark_chart = workbook.add_chart({"type": "pie"})

    mark_chart.add_series(
        {
            "values": [sheet.name, 0, 13, 3, 13],
            "categories": [sheet.name, 0, 12, 3, 12],
        }
    )

    sheet.insert_chart("P1", mark_chart)

    block_chart = workbook.add_chart({"type": "column"})

    block_chart.add_series(
        {
            "values": [sheet.name, 5, 13, 11, 13],
            "categories": [sheet.name, 5, 12, 11, 12],
        }
    )
    
    block_chart.set_legend({"position": "none"})

    sheet.insert_chart("P17", block_chart)

workbook.close()
