import os
import shutil
import colorama
import tkinter
import subprocess

from colorama import Fore
from tkinter import filedialog
from pdf2docx import Converter
root = tkinter.Tk()
root.withdraw()
colorama.init()

def n_files(directory):
    total = 0

    for file in os.listdir(directory):
        if file.endswith('.pdf'):
            total += 1
    return total

def create_folder(directory): # Create a folder to put the converted files
    if not os.path.exists(directory + '/Pdf to Docx'):
        os.makedirs(directory + '/Pdf to Docx')

    else:
        pass

# File Selector using tkinter GUI
def selectpdfFile(directory): # Select and Copy file to directory
    filepath = filedialog.askopenfilename(initialdir="Documents",
                                          title="Select File",
                                          filetypes= (("pdf files","*.pdf"), # File Selection
                                          ))
    file = filepath
    file_dir = directory
    shutil.copy(file, file_dir)

# Main Program!!!
def pdftodocx_convert():
    directory = os.getcwd()
    create_folder(directory)
    selectpdfFile(directory)

    if n_files(directory) == 0:
        print(Fore.RED+"There are no files to convert"+Fore.WHITE)
        exit()

    print(Fore.BLUE+"Converting file to docx..."+Fore.WHITE)

    try:
        for file in os.listdir(directory):
            if file.endswith('.pdf'):
                pdf_file = file
                docx_file = pdf_file.replace(".pdf", r".docx")
                cv = Converter(pdf_file)
                cv.convert(docx_file, start = 0, end=None)
                cv.close()
                shutil.move(docx_file, "Pdf to Docx\\")
                os.remove(pdf_file) # Remove docx file after conversion finished
        
        converted_dir = directory+"\\Pdf to Docx"
        subprocess.Popen(f'explorer "{converted_dir}"')

    except Exception as e:
        print(e)
