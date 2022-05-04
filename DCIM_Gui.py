
#Odhran Doherty DCIM File Creator February 2020

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from datetime import datetime
import time
import os

#Opens File dialog to browse for Excel File

root = tk.Tk()
root.title("DCIM Creator")
canvas1 = tk.Canvas(root, width = 500, height = 150)
canvas1.pack()
lbl = Label(root, text="Converts Excel File to .o2m File for Marcom.", fg="black", font =("Arial Bold",14))
lbl.place(x=250, y=20, anchor="center")

lb5 = Label(root, text="OD 2020", fg="black", font =("Arial Bold",8))
lb5.place(x=450, y=140, anchor="center")



def OutputApplication():

    path = "Output"
    path = os.path.realpath(path)
    os.startfile(path)

def TemplateApplication():

    path = "Template"
    path = os.path.realpath(path)
    os.startfile(path)




def RunApplication():
    #test = 4

    lb2 = Label(root, text="Converting", fg="red", font =("Arial Bold",14))
    lb2.place(x=250, y=75, anchor="center")
    filepath = filedialog.askopenfilename(title = "Select Excel File:", filetypes = (("Excel Files","*.xlsx"),("all files","*.*"))) 


    data = pd.read_excel(filepath,dtype={'Address':np.int32,'Bit':np.int32,'DataTypeNo':np.int32})#Prevents decimal places appearing

    df = pd.DataFrame(data, columns=['Tag'])
    df1 = pd.DataFrame(data, columns=['Address'])
    df2 = pd.DataFrame(data, columns=['Bit'])
    df3 = pd.DataFrame(data, columns=['DataTypeNo'])

    df4 = pd.DataFrame(data, columns=['DataType'])
    df5 = pd.DataFrame(data, columns=['ClientHandle'])
    df6 = pd.DataFrame(data, columns=['Sname'])
    df7 = pd.DataFrame(data, columns=['Node'])
    df8 = pd.DataFrame(data, columns=['Setup'])

    taglist = df.values.tolist()
    addlist = df1.values.tolist()
    bitlist = df2.values.tolist()
    typenolist = df3.values.tolist()
    typelist = df4.values.tolist()
    chandlelist = df5.values.tolist()
    snamelist = df6.values.tolist()
    nodelist = df7.values.tolist()
    setuplist = df8.values.tolist()

    nodename1 = (str(setuplist[0]).strip('[]')).replace("'", "")
    nodename2 = (str(setuplist[1]).strip('[]')).replace("'", "")
    node1word = (str(setuplist[2]).strip('[]')).replace("'", "")
    node2word = (str(setuplist[3]).strip('[]')).replace("'", "")


    node2 = 0  #Intially set to 0. Then if statement below checks if second node exists in excel sheet.

    #Open Text Files to place sections of XML
    file1= open("layoutfiles/item.txt","w+")
    file2= open("layoutfiles/variable.txt","w+")
    file3= open("layoutfiles/parameters.txt","w+")
    file4= open("layoutfiles/variablenode2.txt","w+")
    headerfile = open("layoutfiles/header.txt", "a+")
    afteritem = open("layoutfiles/afteritem.txt", "a+")
    aftermemarea = open("layoutfiles/aftermemoryarea.txt", "a+")
    endoffile = open ("layoutfiles/endoffile.txt", "a+")
    datefile = open("layoutfiles/datecreated.txt","w+")
    #TimeStamps Date of File Creation
    now = str(datetime.now())
    now = "<!-- Created: " + now + "-->"


    file1.write("<Items>\n")
    file2.write('<MemoryArea>\n<Name>'+ nodename1.replace("'", "") + '</Name>\n<Description />\n<Length>'+ node1word.replace("'", "")+'</Length>\n<Type>HoldingRegisters</Type>\n<Address>1</Address>\n<SaveValues>false</SaveValues>\n<BigEndianByte>false</BigEndianByte>\n<BigEndianWord>true</BigEndianWord>\n<BigEndianDWord>false</BigEndianDWord>\n<Variables>\n')
    file3.write("<Mappings>\n")
    file4.write('<MemoryArea>\n<Name>'+ nodename2.replace("'", "") + '</Name>\n<Description />\n<Length>'+ node2word.replace("'", "")+'</Length>\n<Type>HoldingRegisters</Type>\n<Address>2</Address>\n<SaveValues>false</SaveValues>\n<BigEndianByte>false</BigEndianByte>\n<BigEndianWord>true</BigEndianWord>\n<BigEndianDWord>false</BigEndianDWord>\n<Variables>\n')

    now = str(datetime.now())
    now = "\n<!-- Created: " + now + "-->\n"

    datefile.write(now)

    #Iterate through excel file and write to text files
    for i in range(0, len(taglist)):
        
        itemFormat = '<ImpostazioniItem>\n<ItemID>' + (str(taglist[i]).strip('[]')).replace("'", "") + '</ItemID>\n<Status>Stopped</Status>\n<AccessPath />\n<CanonicalDataType>' + str(typenolist[i]).strip('[]') +'</CanonicalDataType>\n<RequestedDataType>0</RequestedDataType>\n<ClientHandle>' + str(chandlelist[i]).strip('[]') + '</ClientHandle>\n<AccessRights>READWRITEABLE</AccessRights>\n</ImpostazioniItem>\n'
        
        file1.write(itemFormat)
        file1.write("\n")

        # If statements check if two nodes used
        if nodelist[i] == [1]:
            varFormat = '<Variable>\n<DataType>' + (str(typelist[i]).strip('[]')).replace("'", "") + '</DataType>\n<Address>'+ str(addlist[i]).strip('[]') + '</Address>\n<Bit>'+ str(bitlist[i]).strip('[]') + '</Bit>\n<Length>1</Length>\n<Description />\n<IsWritable>true</IsWritable>\n<Name>' + str(snamelist[i]).strip('[]')+'</Name>\n</Variable>\n'
            #print(varFormat)
            file2.write(varFormat)
            file2.write("\n")

            paramFormat = '<Mapping>\n<ParametersMap1>\n<string>5</string>\n<string>1</string>\n<string>'+ str(chandlelist[i]).strip('[]') + '</string>\n<string>Value</string>\n</ParametersMap1>\n<ParametersMap2>\n<string>DCIM</string>\n<string>'+ nodename1.replace("'", "") + '</string>\n <string>' + str(snamelist[i]).strip('[]')+'</string>\n</ParametersMap2>\n</Mapping>'
            file3.write(paramFormat)
            file3.write("\n")
        
        if nodelist[i] == [2]:
            node2 = 1
            varFormat = '<Variable>\n<DataType>' + (str(typelist[i]).strip('[]')).replace("'", "") + '</DataType>\n<Address>'+ str(addlist[i]).strip('[]') + '</Address>\n<Bit>'+ str(bitlist[i]).strip('[]') + '</Bit>\n<Length>1</Length>\n<Description />\n<IsWritable>true</IsWritable>\n<Name>' + str(snamelist[i]).strip('[]')+'</Name>\n</Variable>\n'
            #print(varFormat)
            file4.write(varFormat)
            file4.write("\n")

            paramFormat = '<Mapping>\n<ParametersMap1>\n<string>5</string>\n<string>1</string>\n<string>'+ str(chandlelist[i]).strip('[]') + '</string>\n<string>Value</string>\n</ParametersMap1>\n<ParametersMap2>\n<string>DCIM</string>\n<string>'+ nodename2.replace("'", "") + '</string>\n <string>' + str(snamelist[i]).strip('[]')+'</string>\n</ParametersMap2>\n</Mapping>'
            file3.write(paramFormat)
            file3.write("\n")
        


    file2.write("</Variables>\n</MemoryArea>\n")  #Node 1
    file3.write("</Mappings>\n")
    file4.write("</Variables>\n</MemoryArea>\n")  #Node 2

    #Close all files
    file1.close()
    file2.close()
    file3.close()
    file4.close()
    headerfile.close()
    afteritem.close() 
    aftermemarea.close()
    endoffile.close() 
    datefile.close()

    #Combines in to o2m file.
    #Checks if node 2 exists.
    if node2 == 0:
        filenames = ['layoutfiles/header.txt','layoutfiles/datecreated.txt', 'layoutfiles/item.txt', 'layoutfiles/afteritem.txt', 'layoutfiles/variable.txt', 'layoutfiles/aftermemoryarea.txt', 'layoutfiles/parameters.txt', 'layoutfiles/endoffile.txt']
        with open('output.o2m', 'w') as outfile:
            for fname in filenames:
                with open(fname) as infile:
                    outfile.write(infile.read())


    if node2 == 1:
        filenames = ['layoutfiles/header.txt','layoutfiles/datecreated.txt', 'layoutfiles/item.txt', 'layoutfiles/afteritem.txt', 'layoutfiles/variable.txt','layoutfiles/variablenode2.txt', 'layoutfiles/aftermemoryarea.txt', 'layoutfiles/parameters.txt', 'layoutfiles/endoffile.txt']
        with open('output.o2m', 'w') as outfile:
            for fname in filenames:
                with open(fname) as infile:
                    outfile.write(infile.read())


    f = open('output.o2m').read()
    f = f.replace("SLASH", "\\")  #\ is an escape character and causes error when reading from excel. So word SLASH used in import file then replaced in output file. 

    outputfile = open("Output/DCIM.o2m", "w")
    outputfile.write(f)
    outputfile.close()

    
    lb2 = Label(root, text="Complete", fg="green", font =("Arial Bold",16))
    lb2.place(x=250, y=75, anchor="center")
    #print( "*******************************\n o2m file created succesfully \n*******************************" )

    button2 = tk.Button(root, text='Open Output File Location',command=OutputApplication)
    canvas1.create_window(345, 120, window=button2)



button1 = tk.Button(root, text='Browse for Excel File',command=RunApplication)
canvas1.create_window(200, 120, window=button1)

button3 = tk.Button(root, text='Template',command=TemplateApplication)
canvas1.create_window(100, 120, window=button3)

mainloop()
