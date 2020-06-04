# -*- coding: utf-8 -*-
"""
Created on Fri May 29 21:07:52 2020

@author: MJ
"""

#from docx import Document
#from io import StringIO
from docx2python import docx2python
import os
import pandas as pd
#import numpy as np
#from collections import Iterable

ROOT_PATH = os.path.dirname(os.path.abspath('__file__'))
STS2016 = ROOT_PATH + "\Sample term sheet 2016"
#print("OS - ", STS2016)
# extract docx content
doc_result = docx2python(STS2016 + ".docx")
#print(doc_result)
partOneResult = []
#tradeAttributes = ["Bank ref", "FCM TERM SHEET", "Trade Date", "Initial Valuation Date", "Effective Date", "Notional Amount", "Fixed Rate", "(t)", "Fixed Rate Payer Payment Date(t)"]
#print(tradeAttributes)
doc_result = doc_result.body
df = pd.DataFrame(doc_result)
df.replace(to_replace=[None], value=[[""]], inplace=True)                             # To fill None with space in the list

col = len(df.columns)
row = len(df)



#print("row = ", row, " and col = ", col)

'''
This below method is to flatten the string arrays into single array
'''
def flatten(A):
    rt = []
    for i in A:
        if isinstance(i,list): rt.extend(flatten(i))
        else: rt.append(i)
    return rt

def findDateEvent(lst, searchString):
    if(lst[0] == searchString):
        partOneResult.append(str(lst[1]))
        

consolidateList = []
tempStr = ""
bankRefFlag = True
monthsFlag = True
for c in range(col):
    for r in range(0, row-2, 1):
        #print("row - ", r, "\n col - ", c)
        lst = flatten(df[c][r])                                                       # To flatten the list values
        filter_lst = list(filter(lambda x: x != "", lst))                             # To filter the space from the list
        if(len(filter_lst) != 0):                                                     # Consider list which is greater than 0
            #print("\n Values - ", filter_lst)
            #tempStr = ""
            if(monthsFlag == True or bankRefFlag == True):
                for i in filter_lst:
                    tempStr = str(i)
                    #print(tempStr)
                    if (tempStr.find("Bank ref: ") != -1): 
                        partOneResult.append((tempStr.split("Bank ref: ",1)[1]).strip())
                        bankRefFlag = False
                        
                    if (tempStr.find(" FCM TERM SHEET ") != -1): 
                        partOneResult.append((tempStr.split(" FCM TERM SHEET ",1)[0]).strip())
                        monthsFlag = False
                    
                    tempStr = ""

            findDateEvent(filter_lst, "Trade Date:")
            findDateEvent(filter_lst, "Initial Valuation Date:")
            findDateEvent(filter_lst, "Effective Date:")
            findDateEvent(filter_lst, "Notional Amount:")
            findDateEvent(filter_lst, "Fixed Rate:")
            consolidateList.append(filter_lst)

#consolidateList = flatten(consolidateList)
#print(consolidateList)
#After extracting all the required fields
#print(partOneResult)

revIndex = 0
revValue = ""
tempStr = ""
for i in range(len(partOneResult)):
    tempStr = str(partOneResult[i]) 
    #print(tempStr)
    if (tempStr.find("%") != -1):
        partOneResult.append(partOneResult[i])                                        # To add the fixed rate value in the last
        revIndex = i
        revValue = partOneResult[i]
        break
    tempStr = ""

#print(revIndex)
#print(revValue)
#Remove the fixed rate value using index
del(partOneResult[revIndex])
partOneResult.insert(0,partOneResult[1])
del(partOneResult[2])
print(partOneResult)

with open(ROOT_PATH + '/key_attributes.txt', 'w') as file:
    #print("test")
    for listitem in partOneResult:
        #print(listitem)
        if(listitem == partOneResult[len(partOneResult)-1]):
            file.write('%s' %listitem)
        else:
            file.write('%s|' %listitem)
file.close()



'''
Below code is to generate payment date files
using (t) and Fixed Rate Payer Payment Date(t)
'''

print("\n\n")
partOneResult2 = []
for c in range(col):
    for r in range(row-3, row, 1):
        #print("row - ", r, "\n col - ", c)
        lst = flatten(df[c][r])                                                       # To flatten the list values
        filter_lst = list(filter(lambda x: x != "", lst))                             # To filter the space from the list
        if(len(filter_lst) != 0):                                                     # Consider list which is greater than 0
            #print("\n Values - ", filter_lst)
            if (filter_lst[0].find("(t)") == -1 and filter_lst[1].find("(t)") == -1):
                print('|'.join(filter_lst))
                partOneResult2.append('|'.join(filter_lst))
            else:
                continue


with open(ROOT_PATH + '/payment_date.txt', 'w') as file:
    #print("test")
    for listitem in partOneResult2:
        file.write('%s\n' %listitem)
file.close()
