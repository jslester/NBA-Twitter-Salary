import pandas as pd 
import xlrd as xl
import xlsxwriter
from pandas import ExcelWriter
from pandas import ExcelFile 

df=pd.read_excel("NBAStat.xlsx",sheet_name='Sheet3')

class Player:
    def __init__(self, name, PER, BPM, WS, vorp, salary):
        self.name = name
        self.PER = PER
        self.BPM = BPM
        self.WS = WS
        self.vorp = vorp
        self.salary = salary
        self.perDiff = ''
        self.BPMDiff = ''
        self.WSDiff = ''
        self.vorpDiff = ''
        self.averageDiff=''
        
        
playerList = df['Player'].values
perList = df['PER'].values
bpmList = df['BPM'].values
wsList = df['WS/48'].values
vorpList = df['VORP'].values
salaryList = df['Salary'].values
minuteList = df['MP'].values
fullObjectList = []
for i in range(len(playerList)):
    if(float(minuteList[i]) > 500):
        fullObjectList.append(Player(playerList[i],float(perList[i]),float(bpmList[i]),float(wsList[i]),float(vorpList[i]),float(salaryList[i])))

fullObjectList.sort(key = lambda c: c.salary, reverse=True)
testSortList = fullObjectList[:]

#Test for PER
testSortList.sort(key = lambda c: c.PER, reverse=True)
for i in range(len(testSortList)):
    for j in range(len(fullObjectList)):
        if(testSortList[i].name==fullObjectList[j].name):
            fullObjectList[j].perDiff = j-i
#Test for BPM
testSortList.sort(key = lambda c: c.BPM, reverse=True)
for i in range(len(testSortList)):
    for j in range(len(fullObjectList)):
        if(testSortList[i].name==fullObjectList[j].name):
            fullObjectList[j].BPMDiff = j-i

#Test for WS
testSortList.sort(key = lambda c: c.WS, reverse=True)
for i in range(len(testSortList)):
    for j in range(len(fullObjectList)):
        if(testSortList[i].name==fullObjectList[j].name):
            fullObjectList[j].WSDiff = j-i

#Test for VORP
testSortList.sort(key = lambda c: c.vorp, reverse=True)
for i in range(len(testSortList)):
    for j in range(len(fullObjectList)):
        if(testSortList[i].name==fullObjectList[j].name):
            fullObjectList[j].vorpDiff = j-i

exampleDocument = "playerStats.xlsx"
######
workbook = xlsxwriter.Workbook(exampleDocument)
worksheet = workbook.add_worksheet()
worksheet.write(0,0, 'Player Name')
worksheet.write(0,1, 'Salary')
worksheet.write(0,2, 'BPM')
worksheet.write(0,3, 'WS')
worksheet.write(0,4, 'VORP')
worksheet.write(0,5, 'PER')
worksheet.write(0,6, 'PER Difference')
worksheet.write(0,7, 'BPM Difference')
worksheet.write(0,8, 'WS Difference')
worksheet.write(0,9, 'VORP Difference')
worksheet.write(0,10, 'Average Difference')



fullObjectList.sort(key = lambda c: c.salary, reverse=True)
for i in fullObjectList:
    i.averageDiff = (i.perDiff + i.WSDiff + i.vorpDiff + i.BPMDiff)/4
    #print(i.name + " PER Dif:"+ str(i.perDiff)+" TS Dif:"+ str(i.TSDiff)+" WS/48 Dif:"+ str(i.WSDiff)+" VORP Dif:"+ str(i.vorpDiff))
    
fullObjectList.sort(key = lambda c: c.averageDiff, reverse=True)         
for i in range(len(fullObjectList)):
    worksheet.write(i+1,0, fullObjectList[i].name)
    worksheet.write(i+1,1, fullObjectList[i].salary)
    worksheet.write(i+1,2, fullObjectList[i].BPM)
    worksheet.write(i+1,3, fullObjectList[i].WS)
    worksheet.write(i+1,4, fullObjectList[i].vorp)
    worksheet.write(i+1,5, fullObjectList[i].PER)
    worksheet.write(i+1,6, fullObjectList[i].perDiff)
    worksheet.write(i+1,7, fullObjectList[i].BPMDiff)
    worksheet.write(i+1,8, fullObjectList[i].WSDiff)
    worksheet.write(i+1,9, fullObjectList[i].vorpDiff)
    worksheet.write(i+1,10, fullObjectList[i].averageDiff)
workbook.close() 
#print(DataF.iloc[1,6] )
#print(DataF.columns)

