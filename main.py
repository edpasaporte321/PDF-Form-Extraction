import extractingFieldNames as ef
import FreeSimpleGUI as sg

pdfFileName = "TestFile.pdf"
ef.extractingFieldNames(pdfFileName)

# Define the window's contents
layout = [[sg.Text("Input file name (with .pdf extension):")],
          [sg.Input(key='formFileName')],
          [sg.Text(size=(40,1), key='output')],
          [sg.Button('Ok'), sg.Button('Quit')]]

# Create the window
window = sg.Window('Window Title', layout)

# Display and interact with the Window using an Event Loop
while True:
    event, values = window.read()
    # See if user wants to quit or window was closed
    if event == sg.WINDOW_CLOSED or event == 'Quit':
        break
    # Output a message to the window
    try:
        ef.extractingFieldNames(values['formFileName'])
    except:
        window['output'].update('Incorrect File Name, or file not present in directory.')
    else:
        window['output'].update('Field names extracted successfully! Check output.txt.')

# Finish up by removing from the screen
window.close()