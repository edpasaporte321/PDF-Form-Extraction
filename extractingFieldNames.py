from spire.pdf.common import *
from spire.pdf import *
import FreeSimpleGUI as sg
import pandas as pd
from PyPDF2 import PdfReader

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

def getFields(fileNamePDF):
    
    reader = PdfReader(fileNamePDF)
    fields = reader.get_fields()
    listOfKeys = list(fields.keys()) #output is a dict so getting the keys only which are the field names
    #note: the length of listOfKeys is the number of fields in the PDF form with UNIQUE field names. ensure that all field names are unique. (VarName_1, VarName_2, etc.)

    # print(len(listOfKeys)) 
    numberOfTE = 0
    numberOfMC = 0
    numberOfTF = 0
    numberOfNU = 0
    numberOfCO = 0
    numberOfDA = 0
    numberOfNoVariable = 0

    for i in range(len(listOfKeys)):
        newString = listOfKeys[i]
        newString = newString.split("=")[0].strip().split("_")[0].strip()
        if "TE" in newString:
            numberOfTE += 1
            pass
        if "MC" in newString:
            numberOfMC += 1
            pass
        if "TF" in newString:
            numberOfTF += 1
            pass
        if "NU" in newString:
            numberOfNU += 1
            pass
        if "CO" in newString:
            numberOfCO += 1
            pass
        if "DA" in newString:
            numberOfDA += 1
            pass
        if not ("TE" in newString or "MC" in newString or "TF" in newString or "NU" in newString or "CO" in newString or "DA" in newString):
            numberOfNoVariable += 1
            pass
        

    # print(f"Number of TE fields: {numberOfTE}")
    # print(f"Number of MC fields: {numberOfMC}")
    # print(f"Number of TF fields: {numberOfTF}")
    # print(f"Number of NU fields: {numberOfNU}")
    # print(f"Number of CO fields: {numberOfCO}")
    # print(f"Number of DA fields: {numberOfDA}")

    return len(listOfKeys), numberOfTE, numberOfMC, numberOfTF, numberOfNU, numberOfCO, numberOfDA, numberOfNoVariable


def readAndWriteToCSV(pdfFileName):
    data = pd.read_csv('Extracted.csv')
    df = pd.DataFrame(data)
    #print(df)

    df.loc[len(df)] = [pdfFileName, getFields(pdfFileName)[0], getFields(pdfFileName)[1], getFields(pdfFileName)[2], getFields(pdfFileName)[3], getFields(pdfFileName)[4], getFields(pdfFileName)[5], getFields(pdfFileName)[6], getFields(pdfFileName)[7]]

    df.to_csv('Extracted.csv', index=False)


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

def creatingTheWindow():
    # Define the window's contents
    layout = [[sg.Text("Input file name (with .pdf extension):")],
            [sg.Input(key='formFileName')],
            [sg.Text(size=(40,1), key='output')],
            [sg.Button('Get Variables'), sg.Button('Extract Report') , sg.Button('Cancel')]]

    # Create the window
    window = sg.Window('PDF Field Name Extractor', layout)

    # Display and interact with the Window using an Event Loop
    while True:
        event, values = window.read()
        # See if user wants to quit or window was closed
        if event == sg.WINDOW_CLOSED or event == 'Cancel':
            break
        # Output a message to the window
        if event == 'Get Variables':
            try:
                extractingFieldNames(values['formFileName'])
            except:
                window['output'].update('Incorrect File Name, or file not present in directory.')
            else:
                window['output'].update('Field names extracted successfully! Check output.txt.')
        
        if event == 'Extract Report':
            try:
                readAndWriteToCSV(values['formFileName'])
            except:
                window['output'].update('Incorrect File Name, or file not present in directory.')
            else:
                window['output'].update('Report extracted successfully! Check Extracted.csv.')

    # Finish up by removing from the screen
    window.close()

creatingTheWindow()

