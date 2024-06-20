import datetime
import os
import time

import pandas as pd # https://pandas.pydata.org/
import PySimpleGUI as sg # https://pysimplegui.readthedocs.io/en/latest/


# Declares DB location
excelpath = 'NumsAndNames.xlsx'
wb = pd.read_excel(excelpath)

# Declares ticket location
filepath = 'ticket.txt'

# Declares record location
recordExcel = 'Log.xlsx'

# Declares destinations and defines the UI layout
sg.theme("Dark Teal 4")  # https://user-images.githubusercontent.com/46163555/71361827-2a01b880-2562-11ea-9af8-2c264c02c3e8.jpg
destinations = ['Bathroom', 'Guidance', 'Library', 'Another Teacher', 'Water', 'Other']
layout = [
    [sg.Text("", font = ("Roboto", 30))],
    [sg.Text("1. Please select your destination", font = ("Roboto", 40))],
    [sg.Text("", font=("Roboto", 7))],
    [sg.Radio("Bathroom", 'Radio1', k='Bathroom', font=("Roboto", 25), p=((0, 200), (0, 0))), 
        sg.Radio("Guidance", 'Radio1', k='Guidance', font=("Roboto", 25), p=((0, 75), (0, 0))), 
        sg.Radio("Library", 'Radio1', k='Library', font=("Roboto", 25), p=((0, 0), (0, 0)))], 
    [sg.Radio("Another Teacher", 'Radio1', k='Another Teacher', font=("Roboto", 25), p=((0, 100), (0, 0))), 
        sg.Radio("Water", 'Radio1', k='Water', font=("Roboto", 25), p=((0, 125), (0, 0))), 
        sg.Radio("Other", 'Radio1', k='Other', font=("Roboto", 25), p=((0, 200), (0, 0)))],
    [sg.Text(font=("Roboto", 20))],
    [sg.Text("2. Please scan your ID", font = ("Roboto", 40))], 
    [sg.InputText(font=("Roboto", 20), k='-id-')],
    [sg.Submit(font=("Roboto", 20))],
    [sg.Text("", font=("Roboto", 25))],
    [sg.Text('3. Take a notecard', font=('Roboto', 40))],
    [sg.Text('4. Follow the instructions on printer.', font=('Roboto', 40))],
]


def scan(id, loc):
    '''Scans the ID'''
    # Resets all variables in scan()
    x = 0
    idFound = False
    file = open(filepath, 'w')

    # Tests for ID in the database. If found, idFound = True for next step
    while x < len(wb['StudentID']) and idFound == False:
        if id == str(wb['StudentID'].iloc[x]):
            idFound = True
        else:
            x += 1
        
    # If ID was in the DB, move to printing stage. If not, popup    
    if idFound == True:
        endorse(id, x, loc, file)
    else:
        sg.popup_auto_close("ID not found in the database. Please try again.", auto_close_duration=4, keep_on_top=True)


def endorse(id, pos, loc, file):
    '''Sends the ticket to the printer and adds data to the excel record sheet'''
    # Retrieve name from DB, create teacher signature, get current date and time, and write to file
    name = wb['Names'].iloc[pos]
    loc = loc.lower()
    if loc in ['library', 'bathroom']:
        signature = '\n\nPlease permit \n' + name + ' \nto go to the\n' + loc + '. \n\n--Mr. TeacherNameHere \nRoom ####'
    elif loc in ['guidance', 'another teacher', 'water', 'other']:
        signature = '\n\nPlease permit \n' + name + ' \nto go to\n' + loc + '. \n\n--Mr. TeacherNameHere \nRoom ####'
    else:
        print("Error")  # Usually runs if the locations are not properly capitalized. Notice \\loc = loc.lower()\\
        return
    now = datetime.datetime.now()
    currentDate = now.strftime('%B ' + '%d, ' + '%Y')
    currentTime = now.strftime('%I:' + '%M' + ' %p')
    file.write(name + "\n" + id + "\n" + currentDate + "\n" + currentTime + "\n" + signature)
    file.flush()

    # Adds student data to excel record
    df1 = pd.read_excel(recordExcel)
    df2 = pd.DataFrame([[name, currentDate, currentTime, loc]], columns=['Name', 'Date', 'Time', 'Location'])
    data = pd.concat([df1, df2])
    data.to_excel(recordExcel, index=False)

    # Prints the hall pass and closes and deletes ticket.txt file
    #os.startfile(filepath, 'print')
    sg.popup_auto_close("Thank you! Your pass is now printing!", auto_close_duration=10, keep_on_top=True)
    file.close()
    os.remove(filepath)


# Defines the UI
win1 = sg.Window("Hall Pass", layout, size = (1024, 768), keep_on_top=True)

while True:
    # Necessary code to read and close window. Basically keeps it open and shuts it down if pressing exit or the x
    event, values = win1.read()
    if event in (sg.WIN_CLOSED, None, "Exit"):
        break

    # If the user presses the submit button, perform checks on the destination and id after removing the colon from scanning, if done with scanner
    if event == 'Submit':
        dest = ""
        y = values['-id-'].replace(":", "") 
        for item in destinations:
            if values[item] == True:
                dest = item
    if event == 'Submit' and len(y) != 6:
        sg.popup_auto_close("Please enter a valid ID.", auto_close_duration=4, keep_on_top=True)
    elif dest == "":
        sg.popup_auto_close("Please select a destination.", auto_close_duration=4, keep_on_top=True)
    elif event == 'Submit' and y != '':
        try:
            int(y)
        except:
            sg.popup_auto_close("Please enter a valid ID.", auto_close_duration=4, keep_on_top=True)
        else:
            scan(y, dest)
            time.sleep(.5)
    else:
        sg.popup_auto_close("Please enter a valid ID.", auto_close_duration=4, keep_on_top=True)
    
    # If the user presses submit, no matter the result, clear the -id- field and clear the radios
    if event == 'Submit':
        win1['-id-'].Update("")
        for item in destinations:
            win1[item].Update(False)


# Upon closing the UI, close necessary items, then ends the code
win1.close()

