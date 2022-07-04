import os
import time
import colorama

from colorama import Fore
from lib.devs import *
from lib.doctopdf import *
from lib.pdftodocx import *
from lib.pptxtopdf import *

colorama.init()

# null int var
command = 0

# Menu
def menu():
    print(Fore.RED+"""
░█▀█░█▀▄░█▀▀░░░█▀▀░█▀█░█▀█░█░█░█▀▀░█▀▄░▀█▀░█▀▀░█▀▄
░█▀▀░█░█░█▀▀░░░█░░░█░█░█░█░▀▄▀░█▀▀░█▀▄░░█░░█▀▀░█▀▄
░▀░░░▀▀░░▀░░░░░▀▀▀░▀▀▀░▀░▀░░▀░░▀▀▀░▀░▀░░▀░░▀▀▀░▀░▀
Please close your Office Application before converting!
    """+Fore.WHITE)
    print(Fore.CYAN+"Select Converter:"+Fore.WHITE)
    print("[1] Docx to PDF (Word to Pdf)\n[2] PDF to Docx (Pdf to Word)\n[3] PPTX to PDF (PowerPoint to Pdf)\n[4] PDF to PPTX (Pdf to PowerPoint)\n[5] Credits\n[6] Exit")

# Main Program
if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear') # to prevent "colorama" font color bug
    while command != 6:
        menu()
        try:
            command = int(input(Fore.GREEN+"\n>> "+Fore.WHITE))
            if command == 1:
                doctopdf_convert()              

            elif command == 2:
                pdftodocx_convert()    

            elif command == 3:
                pptxtopdf_convert()

            elif command == 4:
                print("Fitur Belum Tersedia")
                time.sleep(2)

            elif command == 4:
                os.system('cls')
                devs()
                print("\n[0] Menu")
                command = int(input(Fore.GREEN+">> "+Fore.WHITE))
                if command == 0:
                    os.system('cls')
                    continue

            elif command == 6: # Exit Program
                break

            else:
                print(Fore.RED+"Unknown Command!"+Fore.WHITE)
                time.sleep(2)
                os.system('cls')
                continue

        except ValueError:
            print(Fore.RED+"An Error Occured! Please input a number!"+Fore.WHITE)
            time.sleep(2)
            os.system('cls')
            continue

        except FileNotFoundError:
            os.system('cls')
            continue

        except Exception as e:
            print(e)
