from spire.pdf.common import *
from spire.pdf import *
import FreeSimpleGUI as sg

def creatingTheWindow():
    # Define the window's contents
    layout = [[sg.Text("Input file name (with .pdf extension):")],
            [sg.Input(key='formFileName')],
            [sg.Text(size=(40,1), key='output')],
            [sg.Button('Extract'), sg.Button('Cancel')]]

    # Create the window
    window = sg.Window('PDF Field Name Extractor', layout)

    # Display and interact with the Window using an Event Loop
    while True:
        event, values = window.read()
        # See if user wants to quit or window was closed
        if event == sg.WINDOW_CLOSED or event == 'Cancel':
            break
        # Output a message to the window
        try:
            extractingFieldNames(values['formFileName'])
        except:
            window['output'].update('Incorrect File Name, or file not present in directory.')
        else:
            window['output'].update('Field names extracted successfully! Check output.txt.')

    # Finish up by removing from the screen
    window.close()

def separateTextByAngleBrackets(text):
    """Separates text by angle brackets and returns a list of the separated items.
     Args:
         text (str): The input text containing items separated by angle brackets.
     Returns:
         list: A list of separated items containing the items separated by angle brackets. 
     """
    text1 = []
    for i in ["><"]:
        for a in range(0, 5):
            try:
                list(text1) == text1.append(text.split("><")[a].strip())
                a = a + 1
            except:
                pass
        return text1

def sanitizingText(text):
    """Sanitizes the input text by removing specific characters.
    Args:
        text (str): The input text to be sanitized.
    Returns:
        str: The sanitized text with specific characters removed.
    """
    for i in [">", "<", "END", "IF", "ELSE", " "]:
        text = text.split("=")[0].strip()
        text = text.replace(i, "")
        textResult1 = text
    return textResult1

def processTheList(nowItsAList):
    """Processes a list of items and returns each item.
    Args:
        list: The input list containing items to be processed.
    Returns: 
        Each item from the input list.
    """
    for ItemsList in nowItsAList:
        #print(f"Item in the list: {ItemsList}")
        return ItemsList

def extractingFieldNames(pdfFileName):
    """Extracts field names from a PDF form and writes them to an output file."""
    doc = PdfDocument()
    doc.LoadFromFile(f"{pdfFileName}")

    form = doc.Form
    formWidget = PdfFormWidget(form)
    field = formWidget.FieldsWidget.get_Item(0)
    textbox = PdfTextBoxFieldWidget(field.Ptr)

    file1 = open("output.txt", "w")
    
    for i in range(formWidget.FieldsWidget.Count):
        field = formWidget.FieldsWidget.get_Item(i)
        textbox = PdfTextBoxFieldWidget(field.Ptr) #declaring the textbox variable and field variable
        beforeBracketedText = textbox.Name #getting bracketed text
        toBeCleaned = beforeBracketedText #assigning bracketed text to a new variable for cleaning
        nowItsAList = separateTextByAngleBrackets(toBeCleaned) #removing angle brackets. output is a list

        if len(nowItsAList) == 1:
            resultOfSanitizingTextAlone = nowItsAList[0] #getting the only item in the list
            file1.write(f"[{sanitizingText(resultOfSanitizingTextAlone)}]\n")

        elif len(nowItsAList) > 1:
            for item in nowItsAList:
                resultOfSanitizingText = sanitizingText(item) #sanitizing each item in the list
                if resultOfSanitizingText != "":
                    file1.write(f"[{resultOfSanitizingText}]\n")

                elif resultOfSanitizingText == "":
                    pass

    file1.close()

creatingTheWindow()