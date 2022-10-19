#1. remember to install "python-docx" and "xlwings" first
#2. Import libraries
import xlwings as xw
import docx
from docx import Document
from tkinter import *
import re

excelFile = xw.Book()                #Creates an empty excel file
excelFile.save('report.xlsx')        #Saves that excel file as "data1"

ws1 = excelFile.sheets['Sheet1']

#Create a Docx file document object and pass the path to the Docx file
Text = docx.Document('SRS_ACE_Pump_X01.docx')

#Create an empty data dictionary
#data = {}

#Create a paragraph object out of the document object.
#This object can access all the paragraphs of the document
#paragraphs = Text.paragraphs

# iterate over all the paragraphs, access the text, and save them into a data dictionary
#for i in range(0, len(Text.paragraphs)):
    #data[i] = tuple(Text.paragraphs[i].text.split(':'))

#values of the dictionary (list)
#data_values = list(data.values())
#print(data_values)

#m = re.search("SRS:(\w+)", paragraphs)
#print m.groups()
fullText = []
for para in Text.paragraphs:
    fullText.append(para.text)

mystring =' '.join(map(str, fullText))
print(mystring)





product = "TARGEST"
name = "NAME"
title = "TITLE"
title2 = "Co-Founder"

#check if there is a keyword that you are looking for and if it is, it will replace with the name
def find_(paragraph_keyword,paragraph):

    if paragraph_keyword in paragraph.text:
        print("found tag:", paragraph_keyword)
        #prints out "found tag:" whenever a tag is found

#going in the document.paragraphs using for loop
for paragraph in Text.paragraphs:

    find_("SRS", paragraph)
    find_("PUMP",paragraph)
    find_("Jan", paragraph)
    find_("CSU", paragraph)

user = [re.findall('(?<=SRS: )\w+', s) for s in mystring]

print(user)
print('found tag', user)

