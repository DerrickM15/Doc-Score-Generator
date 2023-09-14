import pip
import csv
from pathlib import PurePath, Path
import os
import glob
from datetime import date
import random
import re
import tempfile
import webbrowser


def import_or_install(package):
    try:
        __import__(package)
        # print('imported', package)
    except ImportError:
        pip.main(['install', package]) 
        # print('installed',package)

import_or_install('pandas')
import_or_install('openpyxl')
import pandas as pd
import openpyxl


url = r'https://actsoft.crm.dynamics.com/main.aspx?appid=6433e0ff-9fe8-e911-be16-00155d00dabf&forceUCI=1&pagetype=entityrecord&etn=incident&id='
today = date.today()
filedate = "-".join(today.strftime("%#m %#d %Y").split(" "))
user = os.getlogin()
filename_pattern = r"Doc Scores "+filedate+r" \d{1,2}-\d{2}-\d{2} (AM|PM)\.xlsx"
directory = PurePath('C:/', 'Users', user, 'Downloads')

matching_files = glob.glob(os.path.join(directory, "*.xlsx"))
print(matching_files)
matching_files_list = [file for file in matching_files if re.match(filename_pattern, os.path.basename(file))]
print(matching_files_list[0])


read_file = pd.read_excel(matching_files_list[0],engine='openpyxl',)
read_file.to_csv(f"{directory}/temp.csv", index=None, header=True)

cases_dict = {'Case': []
              , 'Created On':[]
              , 'Case Title': [] 
              , 'Case Number': []
              , 'Customer': []
              , 'Owner': []
             }

with open(f"{directory}/temp.csv", newline='', encoding='UTF-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        cases_dict['Case'].append(row['(Do Not Modify) Case'])
        cases_dict['Created On'].append(row['Created On'])
        cases_dict['Case Title'].append(row['Case Title'])
        cases_dict['Case Number'].append(row['Case Number'])
        cases_dict['Customer'].append(row['Customer'])
        cases_dict['Owner'].append(row['Owner'])

owners = set()
for name in sorted(cases_dict['Owner']):
    owners.add(name)

owners = list(sorted(owners))

while True:
    print("Owners: ")
    for count, name in enumerate(owners):
        print(count, name)
    end = len(owners)
    user_input = int(input(f"Would you like to omit any Owners? {end} to End\n"))
    if user_input > len(owners) - 1:
        break
    else:
        del owners[user_input]
        
owners_dict = {name: [] for name in owners}
for name in owners_dict.keys():
    for i, value in enumerate(cases_dict['Owner']):
        if name == value: 
            owners_dict[name].append(i)

for name in owners_dict.keys():
    # first_case = cases_dict['Case'][random.choice(owners_dict[name])]
    webbrowser.open(url+cases_dict['Case'][random.choice(owners_dict[name])])
    # second_case = cases_dict['Case'][random.choice(owners_dict[name])]
    webbrowser.open(url+cases_dict['Case'][random.choice(owners_dict[name])])
    # print(f"{name}:\n{url+first_case}\n{url+second_case}", end = "\n\n")

os.remove(f"{directory}/temp.csv")
os.system('pause')