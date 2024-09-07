import PySimpleGUI as sg
import pandas as pd

#Add some color to the window
sg.theme('DarkTeal')

EXCEL_FILE = 'Excel_data.xlsx'
df = pd.read_excel(EXCEL_FILE)
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15, 1)), sg.InputText(key='-NAME-')],
    [sg.Text('Favourite Colour', size=(15, 1)), sg.Combo(['Red', 'Green', 'Blue', 'Yellow', 'Orange', 'Purple', 'Pink', 'Brown', 'Gray', 'Black', 'White'], key='-COLOUR-', default_value='Blue', size=(20, 1))],
    [sg.Text('.I speak ', size=(15,1)),
     sg.Checkbox('English', key='-ENGLISH-'),
     sg.Checkbox('French', key='-FRENCH-'),
     sg.Checkbox('Spanish', key='-SPANISH-'),
     sg.Checkbox('German', key='-GERMAN-'),
      ],
      [sg.Text('No. of Children',size=(15,1)),sg.Spin([i for i in range(0,16)],initial_value=1,key='-CHILDREN-',size=(5,1))],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]  # Changed from sg.Exist() to sg.Exit()
]

window = sg.Window('Data Entry Form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':  
        clear_input()
    if event == 'Submit':
        df = df._append(values, ignore_index=True)
        try:
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Thanks for submitting the form. Data saved successfully.')
            clear_input()
        except PermissionError:
            sg.popup_error('Error: Unable to save data. The Excel file might be open. Please close it and try again.')
            clear_input()
        except Exception as e:
            sg.popup_error(f'An error occurred while saving data: {str(e)}')
            clear_input()
window.close()


