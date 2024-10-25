import os
from pathlib import Path

class Login:
    def __init__(self) -> None:
        pass
    
    def login(self):
        self.user = 'nichamot'
        self.password = 'Nummon@082024'
    
    def webdriver(self):
        self.dmr = "https://newdebttelesystem.ais.co.th/"
class find_file:
    def __init__(self,path):
        self.path = path
    def file_last_time(self):
        path_file = self.path
        file_excelfile = [file for file in os.listdir(path_file) if file.endswith('.xlsx')]
        if file_excelfile:
            file_ex_time = max(
                (os.path.join(path_file, file) for file in file_excelfile),
                key = os.path.getmtime
            )
            return file_ex_time

url = 'https://tha-crm.wiz.ai/#/login'
us = 'nichamot'
ps = 'Nummon@082024'

dir_path = os.path.dirname(os.path.abspath(__file__)) + os.sep
Data_path = dir_path + 'DATA ' + os.sep
Fill_rec_path = dir_path + 'DATA FILL REC' + os.sep
BC_path = dir_path + 'BC' + os.sep
Master_path = dir_path + 'Master' + os.sep

Path(Data_path).mkdir(parents=True, exist_ok= True)
Path(Fill_rec_path).mkdir(parents=True, exist_ok= True)
Path(BC_path).mkdir(parents=True, exist_ok= True)
Path(Master_path).mkdir(parents=True, exist_ok= True)