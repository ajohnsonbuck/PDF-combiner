# -*- coding: utf-8 -*-
"""
PDF Combiner - Alex Johnson-Buck, 2023

Simple, lightweight GUI program that allows user to combine multiple PDFs into a single PDF.

"""

from PyPDF2 import PdfWriter
import os
import PySimpleGUI as sg

#%%--------------------------------------------------------------------------------
# Function Definitions

def writeMergedPDF(fileList,outFile):  
    # fileList = list of PDF input files
    # outFile = output file name
    merger = PdfWriter()
    for pdf in fileList:
        merger.append(pdf)   
    merger.write(outFile)
    merger.close()

def mainWindow():                       
    # Create PySimpleGUI window layout
    layout = [[sg.Button('  Add file...  ',key='-ADD-'), sg.Button(' Remove file ',key='-REMOVE-'),sg.Button('  ^ Move Up  ',key='-MOVEUP-'),sg.Button('  v Move Down   ',key='-MOVEDOWN-')],
              [sg.Listbox(values=[],size=(50,10),key='-LISTBOX-',expand_x=True)],
              [sg.Button('Save merged file...',key='-SAVE-')]
              ]
    window = sg.Window(title='AJB PDF Combiner', layout=layout, resizable=True, finalize=True)
    return window

def openFile(title,default_path,file_types,multiple_files):
    # Get path to one or multiple files using dialog box.
    # title = string representing window title
    # default_path = string representing default directory to save file to
    # file_types = tuple of tuples of (File Description, filter spec)
    # multiple_files = True or False depending on whether multiple files should be selectable
    if multiple_files == True:
        filename = sg.tk.filedialog.askopenfilenames(
            filetypes=file_types, # tuple of tuples of strings, such as (("All TXT Files", "*.txt"), ("All Files", "*.*"))
            initialdir=default_path,
            title=title
            )        
        filename = list(filename)
    else:
        filename = sg.tk.filedialog.askopenfilename(
            filetypes=file_types, # tuple of tuples of strings, such as (("All TXT Files", "*.txt"), ("All Files", "*.*"))
            initialdir=default_path,
            title=title
            )
    
    return filename

def saveAs(title,default_path,default_file_name,file_types):
    # Get path to one or multiple files using dialog box.
    # title = string representing window title
    # default_path = string representing default directory to save file to
    # file_types = tuple of tuples of (File Description, filter spec)
    filename = sg.tk.filedialog.asksaveasfilename(
        filetypes=file_types, # tuple of tuples of strings, such as (("All TXT Files", "*.txt"), ("All Files", "*.*"))
        initialdir=default_path,
        initialfile=default_file_name,
        title=title
        )
    
    return filename

def updateFilesList(fullFilesList):
    filesList = []
    for n in range(0,len(fullFilesList)):
        filesList = filesList + [os.path.split(fullFilesList[n])[1]]
    return filesList

#%%-----------------------------------------------------------------------------------------------------------------
# Main Body of Code

mainWindow = mainWindow()
path = os.getcwd()
fullFilesList = []
filesList = []


try:    # GUI Window Event Callback Loop
    while True:
        event, values = mainWindow.read()
        if event=='-ADD-': # Add PDF file to filesList
            files = openFile(title='Select your PDF files to combine', 
                          default_path=path,
                          file_types=(("Adobe PDF","*.pdf"),), 
                          multiple_files=True)
            if files is not None and len(files)>0: # check whether files were selected
                # files = files.split(';')
                fullFilesList = fullFilesList + files
                path = os.path.split(fullFilesList[0])[0]
                filesList = updateFilesList(fullFilesList)
                mainWindow['-LISTBOX-'].update(values=filesList)                    
        elif event=='-REMOVE-':
            try:  # If possible, remove selected file from list
                fileToRemove = mainWindow['-LISTBOX-'].get()
                fileToRemove = fileToRemove[0]
                fullFileToRemove = [i for i in fullFilesList if fileToRemove in i]
                fullFilesList = [i for i in fullFilesList if i not in fullFileToRemove]
                filesList = updateFilesList(fullFilesList)
                mainWindow['-LISTBOX-'].update(values=filesList)         
            except:
                print('Error: no file selected to remove.')
        elif event=='-SAVE-': # Save merged PDF in user-defined file
            outFile = saveAs(title='Choose where to save your merged PDF',
                             default_path=path,
                             default_file_name='Merged.pdf',
                             file_types=(("Adobe PDF","*.pdf"),))
            writeMergedPDF(fullFilesList,outFile)
        elif event=='-MOVEUP-': # Move selected item up in list
            try:
                fileToMove = mainWindow['-LISTBOX-'].get()[0]
                fileToMove = [i for i in fullFilesList if fileToMove in i]
                idx = fullFilesList.index(fileToMove[0])
                if idx > 0:
                    displacedItem = fullFilesList[idx-1]
                    fullFilesList[idx-1] = fileToMove[0]
                    fullFilesList[idx] = displacedItem
                    filesList = updateFilesList(fullFilesList)
                    mainWindow['-LISTBOX-'].update(values=filesList)   
            except:
                print('Error: no file selected to move.')
        elif event=='-MOVEDOWN-': # Move selected item down in list
            try:
                fileToMove = mainWindow['-LISTBOX-'].get()[0]
                fileToMove = [i for i in fullFilesList if fileToMove in i]
                idx = fullFilesList.index(fileToMove[0])
                if idx < len(fullFilesList)-1:
                    displacedItem = fullFilesList[idx+1]
                    fullFilesList[idx+1] = fileToMove[0]
                    fullFilesList[idx] = displacedItem
                    filesList = updateFilesList(fullFilesList)
                    mainWindow['-LISTBOX-'].update(values=filesList)
            except:
                print('Error: no file selected to move.')
        elif event==sg.WIN_CLOSED:
            break
except:
    print('Error; closing program') # If error encountered, close window.

mainWindow.close()