import time
import random
import openpyxl
import pywinauto.application
from pywinauto.application import Application

# script used to mass text with grasshopper desktop app

# grasshopper desktop app must be launched and logged in

excel_file_path = 'C:\\Users\\FusRada\\Desktop\\test_sheet.xlsx'


def get_list_of_number_and_name():
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb.active

    column_a = sheet['C']
    column_b = sheet['B']

    phone_number_array = []
    first_name_array = []

    for x in range(1, len(column_a)):
        phone_number_array.append(column_a[x].value)

    for x in range(1, len(column_b)):
        first_name_array.append(column_b[x].value)

    full_list = []

    if len(phone_number_array) == len(first_name_array):
        for x in range(len(first_name_array)):
            number_and_name = {'number': phone_number_array[x], 'name': first_name_array[x]}
            full_list.append(number_and_name)
    else:
        print("Error check the excel file. All columns must be the same length")

    print(full_list)
    return full_list


def text_phone_number(application, number, name, text_message):

    message = str(name) + "!" + text_message

    send_message = application['Grasshopper App'].child_window(title="Send a Message",
                                                               control_type="DataItem").wrapper_object()
    send_message.click_input()

    time.sleep(2)

    input_number = application['Grasshopper App'].child_window(title="Type a phone number", auto_id="sms-dialed-num",
                                                               control_type="Edit").wrapper_object()
    input_number.click_input()

    time.sleep(2)

    input_number.type_keys(number)

    time.sleep(2)

    emoji = application['Grasshopper App'].child_window(title="Emoji Picker", control_type="Image").wrapper_object()

    emoji.click_input()

    time.sleep(2)

    input_message = application['Grasshopper App'].child_window(title="Type a message",
                                                                control_type="Edit").wrapper_object()
    input_message.click_input()

    time.sleep(2)

    input_message.type_keys(message + "{ENTER}", with_spaces=True)

    time.sleep(5)


def begin_mass_texting(dict_list):

    message_dict = {
        0: " Just a friendly reminder that prices for the AYP Convention go up on July 1 (this weekend). Visit "
           "AYP.me/Convention to register for just $104.99 today! Reach out if you have any questions!",
        1: " Wanted to let you know that prices for the AYP Convention increase on July 1 (this weekend). Visit "
           "AYP.me/Convention and register for just $104.99 today! Text back with any questions!",
        2: " Ticket prices for the AYP Convention go up on July 1 (this weekend). Visit AYP.me/Convention to register "
           "for just $104.99 today! Please share with friends and text back with any questions!",
    }

    try:
        # start up
        app = Application(backend='uia').connect(title='Grasshopper App')
        # app['Grasshopper App'].print_control_identifiers()
        messages_tab = app['Grasshopper App'].child_window(auto_id="messages", control_type="Custom").wrapper_object()
        messages_tab.click_input()

        num = 0
        for i in range(len(dict_list)):

            if num == 3:
                num = 0

            text_phone_number(app, dict_list[i]['number'], dict_list[i]['name'], message_dict[num])
            print("row " + str(i + 2) + " has been messaged")
            num += 1

    except pywinauto.application.ProcessNotFoundError:
        print("failed to connect to app, grasshopper desktop app must be launched and logged in")


begin_mass_texting(get_list_of_number_and_name())
