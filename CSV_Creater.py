import pandas as pd
from test1 import FindDetails
import csv
import os

raw_arr = os.listdir()

filename = 'Book1.xlsx'
df = pd.read_excel(filename)

#Notice here we used two '\' backslashesh

def MyFun_CSV(test_string):
    f = open('csvfile.csv','a')
    f.write(test_string+'\n')
    f.close()


if 'csvfile.csv' in raw_arr:
    os.remove(os.getcwd()+'\\csvfile.csv')

MyFun_CSV('Street Address,Searched Address')

for i in range(len(df['Street Address'])):
    value = df['Street Address'][i]
    value = str(value)
    value = value.replace(',',' ')
    value = value.replace('  ',' ')
    try:
        exact_address = FindDetails.calculate(value)
        exact_address = exact_address.replace(',',' ')
        exact_address = exact_address.replace('  ',' ')
    except Exception as e:
        pass
##		exact_address = 'NA'
    csv_string = str(value)+','+str(exact_address)
    MyFun_CSV(csv_string)
    print('Extracted for ',value)



