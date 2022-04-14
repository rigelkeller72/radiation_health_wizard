# Automated Keyboard and Mouse Demo
# Rigel Keller
# Best example inside either a google sheets or excel spreadsheet

from pynput.keyboard import Key, Controller as KeyboardController
from pynput.mouse import Button, Controller as MouseController
import time
import openpyxl

path = "random_data.xlsx"
time.sleep(1)

# workbook object is created
wb_obj = openpyxl.load_workbook(path)

sheet_obj = wb_obj.active
max_col = sheet_obj.max_column
# max_row = sheet_obj.max_row
max_row = 40
big_lst = []


# Loop will print all columns name
for x in range(1, max_row + 1):
    for i in range(1, max_col + 1):
        cell_obj = sheet_obj.cell(row=x, column=i)
        box = cell_obj.value
        if type(box) == int:
            box = str(box)
        # print(type(box))
        big_lst.append(box)
        time.sleep(0.1)

# Converting list into list of Tuples
lst_tuple = [x for x in zip(*[iter(big_lst)]*3)]

# Keyboard and mouse setup
keyboard = KeyboardController()
mouse = MouseController()
t = Key.tab
e = Key.enter
# firstname = 'Rigel'
# lastname = 'Keller'
# dob = '1111111'
#
# firstname1 = 'adam'
# lastname1 = 'smith'
# dob1 = '2222222'
#
# firstname2 = 'george'
# lastname2 = 'washington'
# dob2 = '3333333'
#
# firstname3 = 'michael'
# lastname3 = 'smithers'
# dob3 = '44444444'
#
# firstname4 = 'denzel'
# lastname4 = 'washington'
# dob4 = '555555555'
#
# thistuple = (firstname, lastname, dob)
# thistuple1 = (firstname1, lastname1, dob1)
# thistuple2 = (firstname2, lastname2, dob2)
# thistuple3 = (firstname3, lastname3, dob3)
# thistuple4 = (firstname4, lastname4, dob4)
#
# listotuples = [thistuple, thistuple1, thistuple2, thistuple3, thistuple4]

time.sleep(3)


def myfunction(first_name, last_name, dob_):

    keyboard.press(t)
    keyboard.release(t)

    for char in first_name:
        keyboard.press(char)
        keyboard.release(char)
        time.sleep(0.05)

    keyboard.press(t)
    keyboard.release(t)

    for char in last_name:
        keyboard.press(char)
        keyboard.release(char)
        time.sleep(0.05)

    keyboard.press(t)
    keyboard.release(t)

    for char in dob_:
        keyboard.press(char)
        keyboard.release(char)
        time.sleep(0.05)

    keyboard.press(e)
    keyboard.release(e)
    # # Move pointer relative to current position
    # mouse.move(0, 20)
    #
    # # Press and release
    # mouse.press(Button.left)
    # mouse.release(Button.left)


# Set pointer position
mouse.position = (220, 330)
time.sleep(0.05)

# Press and release
mouse.press(Button.left)
mouse.release(Button.left)
time.sleep(0.5)

# Start going through patient data
for index, tuples in enumerate(lst_tuple):
    myfunction(tuples[0], tuples[1], tuples[2])
