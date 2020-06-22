import json
import xlsxwriter

if __name__ == "__main__":
    with open('item.json') as items:
        datas = json.loads(items.read())

    # for data in datas:
    #     print(data['title'])

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('scraped_data.xlsx')
    worksheet = workbook.add_worksheet()

    # Write the header
    for data in datas:
        header = 0
        for key, value in data.items():
            worksheet.write(0, header, key)
            header += 1
        break

    # Define row and col index
    row = 1
    col = 0
    
    # Iterate over the data and write it out row by row.
    for data in datas:
        for key, value in data.items():
            worksheet.write(row, col, value)
            col += 1
        row += 1
        col = 0

    # Write a total using a formula.
    # worksheet.write(row, 0, 'Total')
    # worksheet.write(row, 1, '=SUM(B1:B4)')

    workbook.close()