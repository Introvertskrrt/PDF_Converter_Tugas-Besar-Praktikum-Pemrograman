import os
import time
import shutil
import subprocess
import colorama

from colorama import Fore
from tkinter import filedialog
from win32com import client # MOST IMPORTANT LIBRARY

colorama.init()

def n_files(directory): # Count .pdf files in the directory
    total = 0

    for file in os.listdir(directory):
        if (file.endswith('.doc') or file.endswith('.docx') or file.endswith('.tmd')):
            total += 1
    return total

def createFolder(directory): # Create Output Folder
    try:
        if not os.path.exists(directory + 'Output Folder/Pdf to Docx'):
            os.makedirs(directory + '/Output Folder/Pdf to Docx')

        else:# If Folder already exists, then Skip/Pass
            pass
    except:
        pass

# open file dialog using Tkinter GUI
def selectpdffile(directory): # Select and Copy file to directory
    filepath = filedialog.askopenfilename(initialdir="Documents",
                                          title="Select File",
                                          filetypes= (("docx files","*.pdf"), # File Selection
                                          ))
    file = filepath
    file_dir = directory
    shutil.copy(file, file_dir) # copy selected file to directory

def remove_pdf(directory): # Remove docx file in directory after conversion
    files_in_directory = os.listdir(directory)
    filtered_files = [file for file in files_in_directory if file.endswith(".pdf")] # Filter format file (make sure program only remove .pptx files)

    for file in filtered_files:
        path_to_file = os.path.join(directory, file)
        os.remove(path_to_file)

# Main Program!!!
def pdftodocx_convert():
    directory = os.getcwd()
    createFolder(directory)
    selectpdffile(directory)
    n_files(directory)

    docx = client.Dispatch('Word.Application')
    docx.Visible = 0
    print(Fore.BLUE+"Converting PDF to Docx..."+Fore.WHITE)

    try:
        for file in os.listdir(directory):
            if file.endswith('.pdf'):
                ending = ".pdf"
        
                new_name = file.replace(ending,r".docx")
                input_file = os.path.abspath(directory + '\\' + file)
                pdf = docx.Documents.Open(input_file)
                output_file = os.path.abspath(directory + "\\Output Folder\\Pdf to Docx" + '\\' + new_name)
                print(f"Converted File: {new_name}")
                pdf.SaveAs(output_file, FileFormat=16)
                pdf.Close()               

    except Exception as e: # If there is an error during conversion, then show the error to user
        print(e)
        os.system('pause')
        os.system('cls')

    # Program Finished
    print(Fore.GREEN+"\nConversion Finished!"+Fore.WHITE)
    time.sleep(2)
    remove_pdf(directory)
    converted_dir = directory+"\\Output Folder\\Pdf to Docx"
    subprocess.Popen(f'explorer "{converted_dir}"')
    os.system("cls")
