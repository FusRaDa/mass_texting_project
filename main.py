import time

import pywinauto.application
from pywinauto.application import Application

# script used to mass text with grasshopper desktop app

# grasshopper desktop app must be launched and logged in


def get_number_and_name():
    print('hi')


def text_phone_number(application):

    send_message = application['Grasshopper App'].child_window(title="Send a Message",
                                                               control_type="DataItem").wrapper_object()
    send_message.click_input()

    input_number = application['Grasshopper App'].child_window(title="Type a phone number", auto_id="sms-dialed-num",
                                                               control_type="Edit").wrapper_object()
    input_number.click_input()
    input_number.type_keys("2402102023")

    emoji = application['Grasshopper App'].child_window(title="Emoji Picker", control_type="Image").wrapper_object()
    emoji.click_input()

    input_message = application['Grasshopper App'].child_window(title="Type a message",
                                                                control_type="Edit").wrapper_object()
    input_message.click_input()
    input_message.type_keys("Hi there {ENTER}", with_spaces=True)

    time.sleep(2)


try:
    # start up
    app = Application(backend='uia').connect(title='Grasshopper App')
    # app['Grasshopper App'].print_control_identifiers()
    messages_tab = app['Grasshopper App'].child_window(auto_id="messages", control_type="Custom").wrapper_object()
    messages_tab.click_input()

    for i in range(0, 3):
        text_phone_number(app)





except pywinauto.application.ProcessNotFoundError:
    print("failed to connect to app, grasshopper desktop app must be launched and logged in")




