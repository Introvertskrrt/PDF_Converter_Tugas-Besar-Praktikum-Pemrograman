import os
import shutil
import colorama
import time
import subprocess

from win32com import client # MOST IMPORTANT LIBRARY
from colorama import Fore
from tkinter import filedialog
from sys import platform

colorama.init()

def n_files(directory): # Count .doc/.docx/.tmd files in the directory
    total = 0

    for file in os.listdir(directory):
        if (file.endswith('.doc') or file.endswith('.docx') or file.endswith('.tmd')):
            total += 1
    return total

def createFolder(directory): # Create Output Folder
    try:
        if not os.path.exists(directory + 'Output Folder/Docx to PDF'):
            os.makedirs(directory + '/Output Folder/Docx to PDF')

        else: # If Folder already exists, then Skip/Pass
            pass
    except:
        pass
    
def doc2pdf(doc, ending, newdic): # Platfrom Selector
    cmd = f"lowriter --convert-to pdf:writer_pdf_Export '{doc}'" # CMD Command to convert file to PDF
    os.system(cmd)
    new_file = doc.replace(ending, r".pdf")

    if platform =='linux': # For Linux
        cmdmove = f"mv '{new_file}' '{newdic}'"

    elif platform == 'win32': # For Windows
        new_file = new_file.replace("/", "\\")
        cmdmove = f"move '{new_file}' '{newdic}'"
    
    os.system(cmdmove)
    print(new_file)

def is_tool(name):
    #Check whether `name` is on PATH and marked as executable.
    from shutil import which

    return which(name) is not None

# open file dialog using Tkinter GUI
def selectdocxFile(directory): # Select and Copy file to directory
    filepath = filedialog.askopenfilename(initialdir="Documents",
                                          title="Select File",
                                          filetypes= (("docx files","*.docx"), # File Selection
                                          ("document files","*.doc"), ("text files", "*.tmd")))
    file = filepath
    file_dir = directory
    shutil.copy(file, file_dir) # copy selected file to directory

def remove_docx(directory): # Remove docx file in directory after conversion
    files_in_directory = os.listdir(directory)
    # Filter format file (make sure program only remove .docx/.doc/.tmd files)
    filtered_files = [file for file in files_in_directory if file.endswith(".docx") or file.endswith(".doc") or file.endswith(".tmd")]

    for file in filtered_files:
        path_to_file = os.path.join(directory, file)
        os.remove(path_to_file)


# Main Program !!!
def doctopdf_convert():
    directory = os.getcwd()

    selectdocxFile(directory)
    createFolder(directory)
	
    if n_files(directory) == 0:
        print(Fore.RED+'There are no files to convert'+Fore.WHITE)
        exit()
		
    print(Fore.BLUE+'Converting Docx to PDF... \n'+Fore.WHITE)

    # Opens each file with Microsoft Word and saves as a PDF
    try:
        if(is_tool('libreoffice') == False):
            word = client.DispatchEx('Word.Application')

        for file in os.listdir(directory):
            if (file.endswith('.doc') or file.endswith('.docx') or file.endswith('.tmd')):
                ending = ""

                if file.endswith('.doc'):
                    ending = '.doc'

                if file.endswith('.docx'):
                    ending = '.docx'

                if file.endswith('.tmd'):
                    ending = '.tmd'

                if is_tool('libreoffice'):
                    in_file = os.path.abspath(directory + '/' + file)
                    new_file = os.path.abspath(directory + 'Output Folder/Docx to PDF')
                    doc2pdf(in_file, ending, new_file)

                if(is_tool('libreoffice') == False):
                    new_name = file.replace(ending,r".pdf")
                    in_file = os.path.abspath(directory + '\\' + file)
                    new_file = os.path.abspath(directory + '\\Output Folder\\Docx to PDF' + '\\' + new_name)
                    doc = word.Documents.Open(in_file)
                    print(f"Converted File: {new_name}")
                    doc.SaveAs(new_file,FileFormat = 17)
                    doc.Close()
                    
    except Exception as e: # If there is an error during conversion, then show the error to user
        print(e)
        os.system('pause')
        os.system('cls')

    # Program Finished
    print(Fore.GREEN+'\nConversion Finished!'+Fore.WHITE)
    time.sleep(2)
    remove_docx(directory)
    converted_dir = directory+"\\Output Folder\\Docx to PDF"
    subprocess.Popen(f'explorer "{converted_dir}"')
    os.system('cls')
