import re
import openpyxl
from openpyxl import Workbook

path = 'C:\\Users\\FusRada\\Desktop\\contacts.xlsx'


def create_new_excel_file(data_list):
    wb = Workbook()
    wb.save("refined_contacts")

    ws = wb.active
    ws.title = "Contacts"
    ws['A1'] = "Names"
    ws['B1'] = "Phone Numbers"








def get_excel_data():
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    data = []

    for row in ws.iter_rows(ws.min_row + 1, ws.max_row):
        row_dict = {'name': None, 'n1': None, 'n2': None}
        for cell in row:
            if cell.column == 1:
                row_dict['name'] = cell.value

            if cell.column == 2:
                row_dict['n1'] = cell.value

            if cell.column == 3:
                row_dict['n2'] = cell.value

        data.append(row_dict)

    return data


def extract_numbers(phone_number):
    result = re.sub('[^0-9]', '', phone_number)

    return result


def refine_excel_data(raw_data):
    refined_list = []

    # determine texting number for name, check first for mobile and if not present then use home number
    for x in range(len(raw_data)):
        data_dict = {'name': raw_data[x]['name'], 'number': None}

        if raw_data[x]['n1'] != '':
            value = extract_numbers(raw_data[x]['n1'])
            data_dict['number'] = value
        else:
            value = extract_numbers(raw_data[x]['n2'])
            data_dict['number'] = value

        refined_list.append(data_dict)

    print(refined_list.__len__())

    # create list of landlines that need to be removed
    remove_list = []
    for x in range(len(refined_list)):
        if refined_list[x]['number'].__len__() > 11:
            remove_list.append(refined_list[x])

    # remove landlines from refined_list
    for x in range(len(remove_list)):
        refined_list.remove(remove_list[x])

    print(refined_list.__len__())

    # remove country code that is not USA/Canada
    remove_country = []
    for x in range(len(refined_list)):
        if refined_list[x]['number'].__len__() > 10 and refined_list[x]['number'][:1] != "1":

            # modify 11-digit numbers that start with 0, remove 0
            if refined_list[x]['number'][:1] == "0":
                val = re.sub("0", "", refined_list[x]['number'], 1)
                refined_list[x]['number'] = val
            else:
                remove_country.append(refined_list[x])

    for x in range(len(remove_country)):
        refined_list.remove(remove_country[x])

    print(refined_list.__len__())

    final_list = list({v['number']: v for v in refined_list}.values())

    print(final_list)


refine_excel_data(get_excel_data())
