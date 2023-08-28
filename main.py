
import openpyxl

def rgb_to_hex(rgb):
    return "#{:02x}{:02x}{:02x}".format(rgb[0], rgb[1], rgb[2])

def get_cell_background_color(cell):
    if cell.fill.start_color.type == 'rgb':
        return cell.fill.start_color.rgb
    elif cell.fill.start_color.type == 'indexed':
        return cell.fill.start_color.indexed

try:
    file_path = 'OSS1.xlsx'
    workbook = openpyxl.load_workbook(file_path)
    sheet_names = ['General Information', 'Assets & Liabilities']

    target_bg_color = (153, 204, 255)

    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        print(f"Sheet: {sheet_name}")

        for row in sheet.iter_rows():
            cell_color = get_cell_background_color(row[4])

            if cell_color == "FF99CCFF":
                data = []
                for cell in row:
                    cell_value = cell.value
                    if cell_value and cell.column == 5 or cell.column == 4:
                        data.append(cell_value)
                print(data)

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    if 'workbook' in locals():
        workbook.close()

