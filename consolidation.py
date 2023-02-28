import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os, sys
import subprocess
import ntpath
import xlsxwriter
import openpyxl
import glob
import pandas as pd
from os import path
import shutil

class EventWatcher():

    def __init__(self):
        self.observer = Observer()

    def execute(self, the_folder):
        DIRECTORY = "C:\\Users\\gabya\\OneDrive\\Documentos\\proyectos\\genpact\\python_exercise\\" + the_folder
        event = EventHandler()
        self.observer.schedule(event,DIRECTORY,recursive=True)
        try:
            self.observer.start()
        except FileNotFoundError as e:
            print(f"Folder not found\n")
            subprocess.call([sys.executable, os.path.realpath(__file__)] + sys.argv[1:])
        try:
            while True:
                time.sleep(15)
        except:
            self.observer.stop()

        self.observer.join()

    
    def movefiles(self, my_path):
        dire = "C:\\Users\\gabya\\OneDrive\\Documentos\\proyectos\\genpact\\python_exercise"
        head = dire + "\\" + my_path
        processed_path = dire + "\\Processed"
        not_path = dire + "\\Not applicable"

        if not path.exists(processed_path):
            os.mkdir(processed_path) 
        if not path.exists(not_path):
            os.mkdir(not_path) 

        for file in os.listdir(head):
            if file.endswith('.xlsx') or file.endswith('.xlsm') or file.endswith('.xlsb'):
                new_path = dire + '\\Processed\\' + file
                #shutil.move(file,new_path)
                os.rename(head + "\\" + file, new_path)
            else:
                new_path = dire + '\\Not applicable\\' + file
                #shutil.move(file,new_path)
                os.rename(head + "\\" + file, new_path)

class EventHandler(FileSystemEventHandler):

    @staticmethod
    def on_any_event(event):
        if event.is_directory:
            return None

        head, tail = ntpath.split(event.src_path)
        print("Received")
        wb = openpyxl.Workbook()  
        folder_name = head.split("\\")[-1]
        dest_filename = 'Master'+ folder_name +'.xlsx' 
        dire = "C:\\Users\\gabya\\OneDrive\\Documentos\\proyectos\\genpact\\python_exercise"
 
        wb.save(os.path.join(dire, dest_filename))
        wb.close()
        for file in os.listdir(head):
            if file.endswith('.xlsx') or file.endswith('.xlsm') or file.endswith('.xlsb'):
                excel_file = pd.ExcelFile(file)
                sheets = excel_file.sheet_names
                for sheet in sheets:
                    df = excel_file.parse(sheet_name = sheet)
                    with pd.ExcelWriter(dest_filename, engine='openpyxl', mode ='a') as writer:
                        df.to_excel(writer, sheet_name=f"{sheet}", header=None, index = False)


if __name__ == '__main__':
    print("\n Please specified the folder to watch: " , end = "")
    my_path = str(input())
    w = EventWatcher()
    w.execute(my_path)
    w.movefiles(my_path)