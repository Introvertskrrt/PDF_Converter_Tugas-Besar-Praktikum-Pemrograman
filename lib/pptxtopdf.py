import os
import time
import shutil
import colorama
import subprocess

from colorama import Fore
from tkinter import filedialog
from win32com import client

colorama.init()

def n_files(directory): # Count .pptx files in the directory
    total = 0

    for file in os.listdir(directory):
        if file.endswith('.pptx'):
            total += 1
    return total

# open file dialog using Tkinter GUI
def selectpptxfile(directory): # Select and Copy file to directory
    filepath = filedialog.askopenfilename(initialdir="Documents",
                                          title="Select File",
                                          filetypes= (("pptx files","*.pptx"), # File Selection (Make sure user only input .pptx file)
                                          ))
    file = filepath
    file_dir = directory
    shutil.copy(file, file_dir) # Copy selected file into program directory

def remove_pptx(directory): # Remove cache/unnecessary pptx files from directory
    files_in_directory = os.listdir(directory)
    filtered_files = [file for file in files_in_directory if file.endswith('.pptx')] # Filter format file (make sure program only remove .pptx files)

    for file in filtered_files:
        path_to_file = os.path.join(directory,file)
        os.remove(path_to_file)

def createFolder(directory): # Create Output Folder
    try:
        if not os.path.exists(directory + 'Output Folder/Pptx to Pdf'):
            os.makedirs(directory + '/Output Folder/Pptx to Pdf')

        else: # If Folder already exists, then Skip/Pass
            pass
    except:
        pass

# Main Program!!!
def pptxtopdf_convert():
    directory = os.getcwd()

    createFolder(directory)
    selectpptxfile(directory)
    n_files(directory)

    pptx = client.Dispatch('PowerPoint.Application')
    
    print(Fore.BLUE+"Converting PPTX to PDF..."+Fore.WHITE)

    try:
        for file in os.listdir(directory):
            if file.endswith('.pptx'):
                ending = '.pptx'

                new_name = file.replace(ending, r".pdf")
                input_file = os.path.abspath(directory + '\\' + file)
                pdf = pptx.Presentations.Open(input_file, WithWindow=False)
                output_file = os.path.abspath(directory + "\\Output Folder\\Pptx to Pdf" + '\\' + new_name)
                print(f"Converted File: {new_name}")
                pdf.SaveAs(output_file, FileFormat = 32) # 32 for pptx
                pdf.Close()

    except Exception as e: # If there is an error during conversion, then show the error to user
        print(e)
        os.system('pause')
        os.system('cls')

    # Program Finished
    print(Fore.GREEN+'\nConversion Finished!'+Fore.WHITE)
    time.sleep(2)
    remove_pptx(directory)
    converted_dir = directory+"\\Output Folder\\Pptx to Pdf"
    subprocess.Popen(f'explorer "{converted_dir}"')
    os.system('cls')
