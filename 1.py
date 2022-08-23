from openpyxl import load_workbook
import PySimpleGUI as sg


sg.theme('DarkAmber')

layout = [[sg.Text('Name'), sg.Push(), sg.Input(key='NAME')],
            [sg.Text('Father name'), sg.Push(), sg.Input(key='FATHER_NAME')],
            [sg.Text('PHONE:'), sg.Push(), sg.Input(key='P_NUMBER')],
            [sg.Text('AADHAR:'), sg.Push(), sg.Input(key='A_NUMBER')],
            [sg.Text('CLASS:'), sg.Push(), sg.Input(key='CLASS_NUMBER')],
            [sg.Text('SEC:'), sg.Push(), sg.Input(key='SEC')],
            [sg.Text('Q_MATHS:'), sg.Push(), sg.Input(key='QM_NUMBER')],
            [sg.Text('Q_SCIENCE:'), sg.Push(), sg.Input(key='QS_NUMBER')],
            [sg.Text('Q_SOCIAL:'), sg.Push(), sg.Input(key='QSS_NUMBER')],
            [sg.Text('Q_LANG:'), sg.Push(), sg.Input(key='QL_NUMBER')],
            [sg.Text('A_MATHS:'), sg.Push(), sg.Input(key='AM_NUMBER')],
            [sg.Text('A_SCIENCE:'), sg.Push(), sg.Input(key='AS_NUMBER')],
            [sg.Text('A_SOCIAL:'), sg.Push(), sg.Input(key='ASS_NUMBER')],
            [sg.Text('A_LANG:'), sg.Push(), sg.Input(key='AL_NUMBER')],
            [sg.Button('ADD'), sg.Button('VIEW'), sg.Button('UPDATE'),sg.Button('DELETE'),sg.Button('Close')]]

window = sg.Window('Data Entry', layout, element_justification='center')


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Close':
        break
    if event == 'ADD':
        try:
            wb = load_workbook('DB1.xlsx')
            sheet = wb['Sheet1']
            ID = len(sheet['ID']) + 100
           

            data = [ID, values['NAME'], values['FATHER_NAME'], values['P_NUMBER'], values['A_NUMBER'], values['CLASS_NUMBER'], values['SEC'], values['QM_NUMBER'], values['QS_NUMBER'], values['QSS_NUMBER'], values['QL_NUMBER'], values['AM_NUMBER'], values['AS_NUMBER'], values['ASS_NUMBER'], values['AL_NUMBER']]

            sheet.append(data)

            wb.save('DB1.xlsx')

            window['NAME'].update(value='')
            window['FATHER_NAME'].update(value='')
            window['P_NUMBER'].update(value='')
            window['CLASS_NUMBER'].update(value='')
            window['SEC'].update(value='')
            window['QM_NUMBER'].update(value='')
            window['QS_NUMBER'].update(value='')
            window['QSS_NUMBER'].update(value='')
            window['QL_NUMBER'].update(value='')
            window['AM_NUMBER'].update(value='')
            window['AS_NUMBER'].update(value='')
            window['ASS_NUMBER'].update(value='')
            window['AL_NUMBER'].update(value='')

            window['NAME'].set_focus()

            sg.popup('Success', 'Data Saved')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')

window.close()