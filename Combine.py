#!/usr/bin/python3

#######################################################################################################
#1. EPLAN E/A-Funktionstexte und Kommentare in der SPS-Symboltabelle müssen übereinstimmen.
#
#2. Bei allen vorhandenen Ein-/Ausgangskarten sind nicht benutze Adressen mit „Reserve_Byte_Bitadresse“
#   zu betexten. Gleicher Eintrag ist in den Kommentaren der Symboltabelle in der Software gefordert.
#
#3. Alle programmierten Ein-/Ausgänge müssen in der Symboltabelle erfasst sein.
#
#4. Ein-/Ausgänge, welche als Platzhalter in nicht aktiven Programmteilen verwendet werden,
#   müssen in der Symboltabelle im Langtext mit dem Suffix „(Option)“ ergänzt werden.
#######################################################################################################

import os
import sys
import time

import PySimpleGUI as sg
import pandas as pd
import re




class RetVal():
    def __init__(self):
        self.eplan= {}
        self.plc=[]
        self.path_merged = ""


class LookUp():
    def __init__(self):
        # Dict for renaming address
        self.dictRen = {
            "A": "A", "AB": "PAB", "AW": "PAW",
            "DB": "DB", "E": "E", "EB": "PEB", "EW": "PEW",
            "FB": "FB", "FC": "FC", "M": "M",
            "MB": "MB", "MW": "MW", "MD": "MD",
            "OB": "OB", "PAB": "PAB", "PAW": "PAW", "PAD": "PAD",
            "PEB": "PEB", "PEW": "PEW", "PED": "PED",
            "SFB": "SFB", "SFC": "SFC", "T": "T",
            "UDT": "UDT", "VAT": "VAT", "Z": "Z"
        }
        #Dict for renaming symbol
        self.dictSymbol = {
            "A": "q", "PAB": "pqb", "PAW": "pqw", "PAD": "pqd",
            "E": "i", "PEB": "pib", "PEW": "piw", "PED": "pid"
        }
        #Dict for renaming datatype
        self.dictDataType = {
            "A": "BOOL", "PAB": "BYTE", "PAW": "WORD", "PAD": "DWORD",
            "E": "BOOL", "PEB": "BYTE", "PEW": "WORD", "PED": "DWORD"
        }



class Subhandler():

    def __init__(self):
        self.df = pd.DataFrame()
        self.path_symbols = ""
        self.path_eplan = ""
        self.path_merged = ""
        self.path_merged=""
        self.listEplan=[]
        #Optionshandling
        self.Option = True
        self.Reserve = True
        self.AddMissing = True
        #GUI Layout
        self.layout = [
            [sg.Text("Enter Path for the EPlan Export File below:")],
            [sg.Input(key='eplan'), sg.FileBrowse(file_types=(('EPlan Export','*.xls'),))],
            [sg.T(background_color='white')],  # spacer
            [sg.Text("Enter Path for the PLC Symbol Export below:")],
            [sg.Input(key='plc'), sg.FileBrowse(file_types=(('System Data Format','*.sdf'),))],
            [sg.Text("!! Only SDF supported so far !!", font='Default 8')],
            [sg.T(background_color='white')],  # spacer
            [sg.Text("Save merged File to:")],
            [sg.Input(key='save'), sg.FileSaveAs(file_types=(('System Data Format', '*.sdf'),))],
            [sg.T(background_color='white')], #spacer
            [sg.Checkbox("Inactive EAs get Suffix '_Option'", default=True,key='Option')],
            [sg.Checkbox("if Comment = 'Reserve' -> Symbols = 'Reserve_Kürzel_Byte_Bitadresse'",default=True,key='Reserve')],
            [sg.Checkbox("Add missing Values to File",default = True,key='Add')],
            [sg.T(background_color='white')],  # spacer
            [sg.OK(), sg.Cancel()]
        ]


    def filedialog(self):
        sg.theme('Reddit')
        sg.theme()
        self.window = sg.Window('EPlan Merge',self.layout)

        ##################################
        #Get User Entrys
        ##################################

        while True:
            event,values = self.window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                self.window.close()
                sys.exit(1996)
            else:
                if event == 'OK':
                    self.path_eplan=values['eplan']
                    self.path_symbols=values['plc']
                    self.path_merged=values['save']
                    self.Option=values['Option']
                    self.Reserve=values['Reserve']
                    self.AddMissing=values['Add']
                    self.bReturn = False
                    for strToCheck in [self.path_eplan,self.path_symbols,self.path_merged]:
                        if len(strToCheck) == 0:
                            self.bReturn = True
                    self.path_merged += '.sdf' if not self.path_merged[-4:] == '.sdf' else ''
                    if self.bReturn:
                        sg.popup_ok('Please enter a valid value for every Path',title='Error')
                        #!!! Dont build another for/while loop around that !!!
                        continue
                    else:
                        self.window.close()
                        break



        ##################################
        #Return RetVal Object
        ##################################

        #Eplan as Dict, PLC as List, Final Path as String
        self.retval=RetVal()
        self.retval.eplan=self.readEplanFile(self.path_eplan)
        self.retval.plc=self.readPLCFile(self.path_symbols)
        self.retval.path_merged=str(self.path_merged)
        return self.retval





    def readEplanFile(self,path):
        self.file_path=path
        title=['BMK', 'Address', 'Symbol']
        self.eplan=pd.read_excel(self.file_path,sheet_name=None,header=None,names=title)
        #change dict of dataframes to dataframe
        for keys in self.eplan:
            self.df=self.eplan[keys]
        #Remove all entrys without name
        self.df=self.df.dropna()
        self.df=self.df.reset_index(drop=True)


        ##################################
        #Cut leading zeros from EAs < 100
        #Rename Symbols to match Siemens
        ##################################

        #Split List items between Nondecimal digits
        rex=r'(^\D+)'
        listCut=[re.split(rex,i,maxsplit=1) for i in self.df.iloc[:,1]]
        #Delete empty Strings from lists
        listCut=[list(filter(None,i)) for i in listCut]
        #Cut Zeros from Bools + Rename EW/EB/etc
        for i in listCut:
            valAdd=""
            for j in i:
                if i.index(j) > 0 and len(i[0]) != 2:
                    valAdd+=(str(float(j)))
                else:
                    if i.index(j) == 0:
                        try:
                            j=LookUp().dictRen[j]
                        except:
                            print(i.index(j))
                    valAdd+=j
            self.listEplan.append(valAdd)
        #Drop old column, reinsert new column
        self.df=self.df.drop(['Address'],axis=1)
        self.df.insert(1,'Address',self.listEplan)
        self.df=self.df.reset_index(drop=True)
        #convert Dataframe to dict
        self.eplan_dict = dict(self.df.iloc[:, 1:3].values.tolist())
        return self.eplan_dict



    def readPLCFile(self,path):
        self.path=path
        with open(self.path) as file:
             self.ListSym=(list(zip(*(line.split('",') for line in file))) )
        #remove whitespace and quatation marks
        self.ListSym=[[element.replace('"','').strip() for element in lst] for lst in self.ListSym]
        #remove whitespaces in address column
        self.ListSym[1]=[n.replace(' ','') for n in self.ListSym[1]]
        return self.ListSym






if __name__ == "__main__":
    Sub=Subhandler()
    Data=Sub.filedialog()

    ################################################
    #initialise all E/As with Extension: ' (Option)'
    ################################################
    max_len=80
    regex_string='^[AEP]'
    ext = ' (Option)'

    if Sub.Option:
        listOpt=[i for i, item in enumerate(Data.plc[1]) if re.search(regex_string,item)]
        for i in listOpt:
            if len(Data.plc[3][i]) <= (max_len - (len(ext))):
                Data.plc[3][i]+=ext
            else:
                Data.plc[3][i]=Data.plc[3][i].replace(Data.plc[3][i][-len(ext):],str(ext))


    ################################################
    #Replace Comments with EPlan Values
    ################################################

    #Replace Comments of existing values
    Data.plc[3]=[Data.eplan[item] if item in Data.eplan else Data.plc[3][Data.plc[1].index(item)] for item in Data.plc[1]]

    #Add missing values
    if Sub.AddMissing:
        listAddplc1=[item for item in Data.eplan if not item in Data.plc[1]]
        listAddplc3=[Data.eplan[item] for item in listAddplc1]
        rex = r'(^\D*|\.)'
        listAddplc0 = [re.split(rex,item) for item in listAddplc1]
        #filter empty strings
        listAddplc0 = [list(filter(None, i)) for i in listAddplc0]
        listSplitAddr = [list(filter(lambda x: x!= ".", i)) for i in listAddplc0]
        listAddplc2 = [LookUp().dictDataType[j] for i in listSplitAddr for j in i if i.index(j) == 0]
        listAddplc0 = []
        for i in listSplitAddr:
            strSymb=""
            for j in i:
                if i.index(j) == 0:
                    strSymb += LookUp().dictSymbol[j]
                else:
                    strSymb += str("_" + j)
            listAddplc0.append(strSymb)
        #Append missing processed values to Data.plc
        plcAdd = [listAddplc0, listAddplc1, listAddplc2, listAddplc3]
        Data.plc = [a+b for a,b in zip(Data.plc,plcAdd) if len(listAddplc0) == len(listAddplc1) == len(listAddplc2) == len(listAddplc3)]


    #################################################################
    #Replace unused E/A with „Reserve_Byte_Bitadresse“
    #Addition: Differentiate between E/A to bypass double occupancy
    #################################################################

    #get list with index of "Reserve" Comments
    if Sub.Reserve:
        rex = r'Reserve'
        listResIndex = [x for x, i in enumerate(Data.plc[3]) if bool(re.search(rex,i,re.IGNORECASE))]
        #Split up Address to create new symbol name
        rex = r'(^\D*|\.)'
        listResSym  = [re.split(rex,item) for item in Data.plc[1]]
        #filter empty strings
        listResSym  = [list(filter(None, i)) for i in listResSym]
        listResSym  = [list(filter(lambda x: x!= ".", i)) for i in listResSym]
        for index in listResIndex:
            strResSym="Reserve"
            for item in listResSym[index]:
                if listResSym[index].index(item) > 0:
                    strResSym += "_" + item
                else:
                    strResSym += "_" + LookUp().dictSymbol[item.upper()]
            Data.plc[0][index]=strResSym


    #################################################################
    #SAVE FILE
    #################################################################
    with open(Data.path_merged,'w+') as file:
        for index in range(len(Data.plc[0])):
            strLines=""
            for i in range(len(Data.plc)):
                strLines += '"' + Data.plc[i][index] + '"' + ','
            strLines = strLines[:-1] + '\n'
            file.writelines(strLines)

    sg.popup_ok("Finished",title="Info")
    sys.exit()