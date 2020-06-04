# -*- coding: utf-8 -*-
"""
Created on Fri May 29 21:07:52 2020

@author: MJ
"""

from docx2python import docx2python
import os
import pandas as pd

ROOT_PATH = os.path.dirname(os.path.abspath('__file__'))

#----- Below are the lists of methods which is being used for 1st part (Extracting data from Sample term sheet 2016.docx)

#Read the word document using docx library
def readWordDocument(documentName):
    return docx2python(documentName)


#This below method is to flatten the string arrays into single array
def flatten(A):
    rt = []
    for i in A:
        if isinstance(i,list): rt.extend(flatten(i))
        else: rt.append(i)
    return rt

# Fetch the corresponding text using matching text else return space
def findDateEvent(lst, searchString):
    string = ""
    if(lst[0] == searchString):
        string = str(lst[1])
    return string

#Extract All the key Attributies from the first table
def extractingTableTradeAttributies(df):
    result = []
    tempStr = ""
    bankRefFlag = True
    monthsFlag = True
    col = len(df.columns)
    row = len(df)
    for c in range(col):
        for r in range(0, row-2, 1):
            #print("row - ", r, "\n col - ", c)
            lst = flatten(df[c][r])                                                       # To flatten the list values
            filter_lst = list(filter(lambda x: x != "", lst))                             # To filter the space from the list
            if(len(filter_lst) != 0):                                                     # Consider list which is greater than 0
                if(monthsFlag == True or bankRefFlag == True):
                    for i in filter_lst:
                        tempStr = str(i)
                        if (tempStr.find("Bank ref: ") != -1): 
                            result.append((tempStr.split("Bank ref: ",1)[1]).strip())
                            bankRefFlag = False
                        
                        if (tempStr.find(" FCM TERM SHEET ") != -1): 
                            result.append((tempStr.split(" FCM TERM SHEET ",1)[0]).strip())
                            monthsFlag = False
                    
                        tempStr = ""

                result.append(findDateEvent(filter_lst, "Trade Date:"))
                result.append(findDateEvent(filter_lst, "Initial Valuation Date:"))
                result.append(findDateEvent(filter_lst, "Effective Date:"))
                result.append(findDateEvent(filter_lst, "Notional Amount:"))
                result.append(findDateEvent(filter_lst, "Fixed Rate:"))
    result = list(filter(lambda x: x != "", result))                                 # To filter the space from the list            
    return result

#Extract All the payment dates from the second table
def extractingTablePaymentDate(df):
    result = []
    col = len(df.columns)
    row = len(df)
    for c in range(col):
        for r in range(row-3, row, 1):
            lst = flatten(df[c][r])                                                       # To flatten the list values
            filter_lst = list(filter(lambda x: x != "", lst))                             # To filter the space from the list
            if(len(filter_lst) != 0):                                                     # Consider list which is greater than 0
                #print("\n Values - ", filter_lst)
                if (filter_lst[0].find("(t)") == -1 and filter_lst[1].find("(t)") == -1):
                    #print('|'.join(filter_lst))
                    result.append('|'.join(filter_lst))
                else:
                    continue
    return result

#Creating key_attributes file using list of key attributies
def writeKeyAttributiesFile(keyAttributiesResult, keyAttributiesFileName):
    with open(keyAttributiesFileName, 'w') as file:
    #print("test")
        for listitem in keyAttributiesResult:
            #print(listitem)
            if(listitem == keyAttributiesResult[len(keyAttributiesResult)-1]):
                file.write('%s' %listitem)
            else:
                file.write('%s|' %listitem)
    file.close()

#Creating payment_date file using list containing dates
def writePaymentDateFile(paymentDatesResult, paymentDateFileName):
    with open(paymentDateFileName, 'w') as file:
        for listitem in paymentDatesResult:
            file.write('%s\n' %listitem)
    file.close()



#---- Using the above methods, to complete the 1st part document extracting and storing the required results in the text files
    
STS2016 = ROOT_PATH + "\Sample term sheet 2016" + ".docx"
doc = readWordDocument(STS2016)
#print(doc)

df = pd.DataFrame(doc.body)
df.replace(to_replace=[None], value=[[""]], inplace=True)                             # To fill None with space in the list

#Extracting all ket attributies from first tables
keyAttributiesResult = extractingTableTradeAttributies(df)
#print(keyAttributiesResult)

#Appending the fixed rate value in the last
index = 0
value = ""
tempStr = ""
for i in range(len(keyAttributiesResult)):
    tempStr = str(keyAttributiesResult[i]) 
    #print(tempStr)
    if (tempStr.find("%") != -1):
        keyAttributiesResult.append(keyAttributiesResult[i])                          # To add the fixed rate value in the last
        index = i
        value = keyAttributiesResult[i]
        break
    tempStr = ""
        
#Rearranging the key attributies
del(keyAttributiesResult[index])
keyAttributiesResult.insert(0,keyAttributiesResult[1])
del(keyAttributiesResult[2])
print(keyAttributiesResult)

#Writing the key attributies in text file
keyAttributiesFileName = ROOT_PATH + '/key_attributes.txt'
writeKeyAttributiesFile(keyAttributiesResult, keyAttributiesFileName)
print("\n\n")

#Extracting the payment dates from second tables
paymentDatesResult = extractingTablePaymentDate(df)
print(paymentDatesResult)

#Writing the payment dates in text file
paymentDateFileName = ROOT_PATH + '/payment_date.txt'
writePaymentDateFile(paymentDatesResult, paymentDateFileName)


