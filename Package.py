import openpyxl.worksheet
import openpyxl.worksheet.worksheet
import pandas
import numpy
import sys
import shutil
import os
import subprocess
import time


# main
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("[argv] [1: type]")
        exit(0)

    type = sys.argv[1]
    if type == "1":
        if os.path.exists("./pack"):
            shutil.rmtree("./pack") 
        os.mkdir("./pack")
        shutil.copy("./MakeTableRun.py","./pack/MakeTableRun.py")
    elif type == "2":
        os.chdir("./pack")
        subprocess.run("pyinstaller.exe -F -w ./MakeTableRun.py",capture_output=True)
    elif type == "3":
        os.mkdir("./release/config")
        shutil.copy("./pack/dist/MakeTableRun.exe","./release/MakeTableRun.exe")
        shutil.copy("./config/readme.txt","./release/config/readme.txt")
        shutil.copy("./config/config.xlsx","./release/config/config.xlsx")
    elif type == "4":
        shutil.make_archive("./release","zip","./release")

    print("== end ==")

