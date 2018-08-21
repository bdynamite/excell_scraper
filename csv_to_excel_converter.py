import os
import csv
import argparse

import xlrd


def get_params():
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument('input', help='excel file path')
    parser.add_argument('-o', '--output', help='csv folder path')
    return parser.parse_args()


def convert_cell(cell, datemode):
    if cell.ctype == 3:
        return xlrd.xldate_as_datetime(cell.value, datemode)
    if cell.ctype == 1:
        return cell.value.strip()
    return cell.value


def write_csv_file(sheet, output_dir, datemode):
    with open(os.path.join(output_dir, '{}.csv'.format(sheet.name)), 'w') as csv_file:
        writer = csv.writer(csv_file, delimiter=';')
        for row_number in range(sheet.nrows):
            values_row = [convert_cell(sheet.cell(row_number, col_number), datemode)
                          for col_number in range(sheet.ncols)]
            writer.writerow(values_row)


def csv_from_excel(input_path, output_path):
    wb = xlrd.open_workbook(input_path)
    for sh in wb.sheets():
        write_csv_file(sh, output_path, wb.datemode)


def remove_accessory_quotes(path):
    if path[0] == path[-1] == '"':
        path = path[1:-1]
    return path


def check_paths(input, output):
    input = remove_accessory_quotes(input)
    if not os.path.isfile(input):
        raise FileExistsError('There is no file for this path ({})'.format(input))
    if output is None:
        head, tail = os.path.split(input)
        output = os.path.join(head, tail.split(".")[0])
    else:
        output = remove_accessory_quotes(output)
    os.makedirs(output)
    return input, output


if __name__ == "__main__":
    params = get_params()
    excel_path, dir_path = check_paths(params.input, params.output)
    csv_from_excel(excel_path, dir_path)
    print('Success! You can find your csv file(s) on this path {}'.format(os.path.abspath(dir_path)))



