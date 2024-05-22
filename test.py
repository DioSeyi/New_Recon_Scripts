# pip install openpyxl

#from openpyxl import workbook, load_workbook
#wb = load_workbook('name.csv')
#ws = wb.active
#print(ws)

# print('Hello World')
# x = 1
# print(x)
# x = x + 6
# print(x)
# exit()

# name = input('Enter file:')
# handle = open(name)

# counts = dict()
# for line in handle:
#     words = line.split()
#     for word in words:
#         counts[word] = counts.get(word,0) + 1
        
# bigcount = None
# bigword = None
# for word, count in counts.items():
#     if bigcount is None or count > bigcount:
#         bigword = word
#         bigcount = count
        
# print(bigword, bigcount)

# fruit = 'banana'
# for letter in fruit:
#     print(letter)
# if letter == 'n':
#     print(n)




# import openpyxl as xl
# import pprint

import pandas as pd
df = pd.read_csv('buk1.csv')
print(df)
