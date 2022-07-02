import os
import time
import colorama

from colorama import Fore
from lib.docx2pdf import *
from lib.kelompok import *

colorama.init()

# null int var
command = 0

# Menu
def menu():
    print(Fore.RED+"File to PDF Converter"+Fore.WHITE)
    print("[1] Docx to PDF\n[2] PDF to Docx\n[3] Credits\n[4] Exit")

# Main Program
if __name__ == "__main__":
    while command != 4:
        menu()
        try:
            command = int(input(Fore.GREEN+"\n>> "+Fore.WHITE))
            if command == 1:
                doctopdf_convert()

            elif command == 2:
                print("Fitur Belum Tersedia!")
                time.sleep(2)
                os.system('cls')
                continue

            elif command == 3:
                os.system('cls')
                kelompok()
                print("\n[0] Menu")
                command = int(input(">> "))
                if command == 0:
                    os.system('cls')
                    continue

            elif command == 4:
                break

            else:
                print(Fore.RED+"Unknown Command!"+Fore.WHITE)
                time.sleep(2)
                os.system('cls')
                continue

        except ValueError:
            print(Fore.RED+"An Error Occured! Please input a number!")
            time.sleep(2)
            os.system('cls')
            continue

        except FileNotFoundError:
            os.system('cls')
            continue

        except:
            print(Fore.RED+"Unknown Error Occured!"+Fore.WHITE)
            time.sleep(2)
            os.system('cls')
            continue
