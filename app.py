from distutils.command.config import config
import json
import os
from time import time
from tkinter import Button
import PySimpleGUI as sg
from file_parser.backend import checkClashes, viewFreeAndBusy, createNewTimetable, createPersonalTimetable
from os import listdir
from sys import exit

sg.theme('reddit')


def setup(string) -> str:
    # Window layout
    layout = [
        [sg.Text(string)],
        [sg.In(size=(25, 1), enable_events=True,
               key='-FOLDER-'), sg.FolderBrowse()],
        [sg.Button('Ok'), sg.Button('Quit')]

    ]

    # Window
    window = sg.Window('Folder Select', layout=layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Quit':
            exit()

        if values['-FOLDER-'] == '':
            sg.popup('Please select a folder', title='Error')
            continue

        if event == 'Ok':
            if values['-FOLDER-'] != None:
                window.close()
                break

            else:
                sg.popup('Select a folder', title='Error')
                continue
    return values['-FOLDER-']


def setup1():
    layout = [
        [sg.Text("Number of periods"), sg.InputText()],
        [sg.Text("Active working days"), sg.InputText()],
        [sg.Text("Repeat Count"), sg.InputText()],
        [sg.Text("Output Filename"), sg.InputText()],
        [sg.Button('Confirm'), sg.Button('Cancel')]
    ]

    window = sg.Window('Details', layout=layout)

    while True:
        event, values = window.read()

        if event == 'Cancel' or event == sg.WIN_CLOSED:
            break

        if event == 'Confirm':
            items = list(values.values())
            return items


def timetableGenerator(folder, periodNum, activeDays, repeatCount, fout):
    teacherCreds = {}
    layout = [
        [sg.Text("Teacher's Name"), sg.InputText()],
        [sg.Text("Teacher's Subject"), sg.InputText()],
        [sg.Button('Add Teacher'), sg.Button('Generate Timetable'),
         sg.Button('Cancel All Inputs'), sg.Button('Exit')]
    ]

    window = sg.Window('Generate Timetable', layout=layout)

    while True:  # Errors: Number of given periods less than repeat count

        event, values = window.read()

        if event == 'Exit':
            break

        tName, tSub = list(values.values())

        if event == 'Add Teacher':
            teacherCreds[tSub] = tName

        if event == 'Cancel All Inputs':
            teacherCreds.clear()
            sg.popup('Inputs Cleared Successfully', title='Success')

        try:
            if event == 'Generate Timetable':
                createNewTimetable(int(periodNum), int(activeDays),
                                   folder, fout, teacherCreds, int(repeatCount))
                sg.popup('Timetable generated Successfully', title='Success')
                break
        except IndexError:
            sg.popup('Number of periods too less. Please try again.',
                     title='Error')

        # sg.popup('An Error occured. Please try again.', title='Error')


def main_window(list_of_files, folder) -> None:
    days_of_the_week = ['Monday', 'Tuesday',
                        'Wednesday', 'Thursday', 'Friday', 'Saturday']

    # layout
    layout = [
        [sg.Text('Timetables to compare')],
        [sg.Combo(list_of_files, size=(25, 1))],
        [sg.Combo(list_of_files, size=(25, 1))],
        [sg.Text('              ')],
        [sg.Listbox(days_of_the_week, size=(30, 5))],
        [sg.Radio('Check Clashes', 'Function'), sg.Radio(
            'Get free teachers', 'Function'), sg.Radio('Get Busy Teachers', 'Function')],
        [sg.Text('Enter Period'), sg.InputText()],
        [sg.Text('              ')],
        [sg.Button('Ok', size=(4, 1)), sg.Button('Quit', size=(4, 1)),
         sg.Button('Generate Timetable', size=(15, 1))]
    ]

    window = sg.Window('App', layout=layout)

    # main loop
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Quit':
            break

        if event == 'Generate Timetable':
            periodNum, activeDays, repeatCount, fout = setup1()
            if not os.path.exists(rf'{folder}\class_timetables'):
                os.mkdir(rf'{folder}\class_timetables')
            timetableGenerator(rf'{folder}\class_timetables', periodNum,
                               activeDays, repeatCount, fout)

            with open(file='config.json', mode='r') as file:
                teachers = list(json.load(file).values())

            if not os.path.exists(rf'{folder}\personal_timetables'):
                os.mkdir(rf'{folder}\personal_timetables')

            for teacher in teachers:
                createPersonalTimetable(rf'{folder}\personal_timetables', teacher, rf'{folder}\class_timetables', int(
                    activeDays), int(periodNum))
            break

        # fn1 = Check clashes, fn2 = get free teachers, fn3 = get busy teachers
        f1, f2, day, fn1, fn2, fn3, period = list(values.values())

        if f1 == f2:
            sg.popup('Please select 2 different timetables', title='Error')
            continue

        if event == 'Ok':
            try:
                if fn1 == True:
                    res = checkClashes(
                        f'{folder}/{f1}', f'{folder}/{f2}', str(day[0]))
            except KeyError:
                sg.popup(
                    'No periods on this day', title='Error')
                continue

            if fn2:
                try:
                    res = viewFreeAndBusy(
                        folder=rf'{folder}\personal_timetables', day=day[0], period_number=int(period), view_busy=False)
                except ValueError:
                    sg.popup('Enter a value for the period', title='Error')
                    continue
                except IndexError:
                    sg.popup(
                        'Enter a value for the period within the correct number of periods', title='Error')
                    continue
                except KeyError:
                    sg.popup(
                        'No periods on this day', title='Error')
                    continue

            if fn3:
                try:
                    res = viewFreeAndBusy(
                        folder=rf'{folder}\personal_timetables', day=day[0], period_number=int(period), view_busy=True)
                except ValueError:
                    sg.popup('Enter a correct value for the period',
                             title='Error')
                    continue
                except IndexError:
                    sg.popup(
                        'Enter a value for the period within the correct number of periods', title='Error')
                    continue
                except KeyError:
                    sg.popup(
                        'No periods on this day', title='Error')
                    continue

        if len(res) != 0:
            sg.popup('\n'.join(i.replace('.xlsx', '')
                     for i in res), title='Result')
        else:
            sg.popup('None', title='Result')


def main():
    k = setup('Select the folder containing the timetables')
    if not os.path.exists(rf'{k}\personal_timetables'):
        fileList = listdir(k)
        main_window(fileList, k)
    else:
        fileList = listdir(rf'{k}\personal_timetables')
    main_window(fileList, rf'{k}\personal_timetables')

# Scope: Compare timetable of one teacher to another for checking clashes. Most manipulations will be performed only on 2 timetables. It is also possible to generate boilerplate timetables that can be edited as the user sees fit. Any clashes that arise can be checked using the app.


if __name__ == '__main__':
    main()
