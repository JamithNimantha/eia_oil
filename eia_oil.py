import requests
from datetime import datetime
import os
import time
import xlrd
import shutil

error = '503 - Sorry! This file is temporarily unavailable.'
dir_path = os.path.dirname(os.path.realpath(__file__))
# current = datetime.now().strftime('OIL_%Y-%m-%d.xls')
current = datetime.now().strftime('OIL_%Y-%m-%d.csv')
start_time = time.time()

#########################
# CHANGE WAIT TIME BELOW
wait_sleep_time = 10
########################

path = f'{dir_path}{os.sep}Data{os.sep}Oil{os.sep}'
if not os.path.exists(path):
    os.makedirs(path)


def is_valid(content):
    temp_folder = f'{dir_path}{os.sep}.temp{os.sep}'
    if not os.path.exists(temp_folder):
        os.mkdir(temp_folder)
    open(temp_folder + current, 'wb').write(content)
    workbook = xlrd.open_workbook(temp_folder + current)
    worksheet = workbook.sheet_by_index(0)
    first_row_value = worksheet.row(1)[0].value
    if first_row_value == error:
        print(worksheet.row(1)[0].value)
        shutil.rmtree(temp_folder)
        return False
    shutil.rmtree(temp_folder)
    return True


while True:
    try:
        response = requests.get('https://ir.eia.gov/wpsr/table1.csv', allow_redirects=True)
        if response.status_code == 200:
            open(path + current, 'wb').write(response.content)
            print(f'Found {path + current} with Data !!!')
            break
        else:
            print(error)
        print(f'Trying again in {wait_sleep_time} seconds!')
        time.sleep(wait_sleep_time)
    except:
        print("Trying Again in 5 Seconds")
        if time.time() - start_time > 600:
            print('10 minutes have passed exiting')
            exit()
        time.sleep(wait_sleep_time)
        continue
exit()
