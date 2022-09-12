import os 
import pandas as pd

# request 1 - create new directory
myPath = os.getcwd()
newfolder = myPath + '/Hassal_ Samantha Excel Assessment VBA'
if not os.path.exists(newfolder):
    os.mkdir(newfolder)

# request 2 - copy table and save it as a new workbook
myDF = pd.read_excel("NCoE Candidate Excel Assessment.v2.2 - updated.xlsm", "2. Formatting")
x = list(myDF.columns)
myDF = myDF[x[1:]]
myDF = myDF.loc[4:15]
filename = 'Hassal_ Samantha Excel Assessment VBA.xlsx'
myDF.to_excel(newfolder+'/'+filename, index=False, header=False)

# request 3 - return rows of data containing a keyword found in the engagement name
myDF2 = pd.read_excel(newfolder+'/'+filename)
word = raw_input("enter a keyword")
myDF2 = myDF2[myDF2['Engagement Name'].str.contains(word)]

myList = ["Client Name", "Engagement Name", "Engagement Partner", "Total Expense", "Charged Hours", "TER", "NER", "SER"]

print(myDF2[myList])

