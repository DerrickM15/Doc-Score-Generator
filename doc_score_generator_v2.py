import csv
from pathlib import PurePath, Path
from os import getlogin
from datetime import date
import random

url = r'https://actsoft.crm.dynamics.com/main.aspx?appid=6433e0ff-9fe8-e911-be16-00155d00dabf&forceUCI=1&pagetype=entityrecord&etn=incident&id='
today = date.today()
filedate = "-".join(today.strftime("%#m %#d %Y").split(" "))
user = getlogin()
filename = f"Doc Scores {filedate}.csv"
path = PurePath('C:/', 'Users', user, 'Downloads', filename)

cases_dict = {'Case': []
              , 'Created On':[]
              , 'Case Title': [] 
              , 'Case Number': []
              , 'Customer': []
              , 'Owner': []
             }

with open(path, newline='', encoding='UTF-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        cases_dict['Case'].append(row['\ufeff(Do Not Modify) Case'])
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
    first_case = cases_dict['Case'][random.choice(owners_dict[name])]
    second_case = cases_dict['Case'][random.choice(owners_dict[name])]
    print(f"{name}:\n{url+first_case}\n{url+second_case}", end = "\n\n")