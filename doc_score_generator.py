import csv
from pathlib import PurePath, Path
from os import getlogin
from datetime import date
import pandas as pd
import random
import urllib.request

url = r'https://actsoft.crm.dynamics.com/main.aspx?appid=6433e0ff-9fe8-e911-be16-00155d00dabf&forceUCI=1&pagetype=entityrecord&etn=incident&id='

today = date.today()
filedate = "-".join(today.strftime("%#m %#d %Y").split(" "))
user = getlogin()

filename = f"Doc Scores {filedate} .csv"
path = PurePath('C:/', 'Users', user, 'Downloads', filename)

csvData = pd.read_csv(path)
headers = list(csvData.columns)
grouped_data = csvData.groupby('Owner')
row_count = csvData.value_counts('Owner')

random_cases = grouped_data.apply(lambda x: x.iloc[random.sample(range(0, len(x.index)), 2)])['(Do Not Modify) Case']


for case in random_cases:
    # html_page = urllib.request.urlopen(url+case)
    print(url+case)


# print(headers)
# print([item for item in grouped_data])



#TODO: Change filename to regex to allow any date and time
#TODO: Open links in chrome? Edge?
#TODO: Allow to select omitted owners
    # Maybe use enumberate on a list provided by the grouped_date['Owner'] parameter with Exit at 0

