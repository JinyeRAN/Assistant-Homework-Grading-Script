import os
import pandas as pd
import win32com
from win32com import client
import numpy as np 
import shutil

def DiffFile(DoneRoot,RawRoot):
    ContainerD = []
    ContainerR = []
    for itemD, itemR in os.listdir(DoneRoot), os.listdir(RawRoot):
        ContainerD.append(itemD)
        ContainerR.append(itemR)
    Diff = list(set(ContainerR) - set(ContainerD))
    data = list(map(lambda x: os.path.join(RawRoot, x), Diff))
    return data

def GetFileRoot(root):
    container = []
    for item in os.listdir(root):
        FileRoot = os.path.join(root, item)
        container.append(FileRoot)
    container.sort()
    container = list(map(lambda x: os.path.join(os.getcwd(),x),container))
    return container

def VisualDocx(WordRoot):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = 1
    doc = word.Documents.Open(WordRoot)
    return doc

def RegisterDegreeExcel(ExcelRoot,KeyName,KeyProjectNum,degree):
    excel = pd.read_excel(ExcelRoot)
    idex = excel[(excel.姓名==str(KeyName))].index.tolist()
    excel.loc[int(idex[0]), str(KeyProjectNum)] = str(degree)
    excel.to_excel(ExcelRoot,index=False)
    
def RegisterDegreeWord(WordRoot,degree):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = 1
    doc = word.Documents.Open(WordRoot)
    myRange1 = doc.Range(0,0)
    myRange1.InsertBefore(degree)
    myRange1.Font.Name = "宋体"
    myRange1.Font.Size = "48"
    myRange1.Font.Bold = 200
    myRange1.Font.Color = 255
    myRange1.ParagraphFormat.Alignment = 0
    doc.Save()
    doc.Close()
    
def main():
    InfoExcelRoot = './SystemStudentName.xlsx'
    while True:
        Model = input('what turn is it\nr: review turn  n:new turn\n')
        if Model == 'r':
            ProjectName = input('Enter Project Name')
            InfoDoneRoot, InfoRawRoot = os.path.join(os.getcwd(), 'Raw', ProjectName), os.path.join(os.getcwd(), 'Done',ProjectName)
            FileList = DiffFile(InfoDoneRoot, InfoRawRoot)
            break
        elif Model == 'n':
            ProjectName = input('project root without convert char\n')
            InfoRoot = os.path.join(os.getcwd(), 'Raw', ProjectName)
            InfoProjectNum = input('project numbers\n')
            InfoDoneRoot = os.path.join('./Done',InfoRoot)
            if not os.path.exists(InfoDoneRoot):
                os.mkdir(InfoDoneRoot)
            FileList = GetFileRoot(InfoRoot)
            break

    while not len(FileList) == 0:
        File = FileList[0]
        doc = VisualDocx(File)
        while True:
            degree = input('FileName: ' + File.split('\\')[-1] + '.\nEnter degree e：优秀  j：良好  m：中等\n')
            if degree=='e':
                degree = '优'
                break
            elif degree=='j':
                degree = '良'
                break
            elif degree=='m':
                degree = '中'
                break
        doc.Close()
        KeyName = File.split('\\')[-1].split('-')[1]
        RegisterDegreeExcel(InfoExcelRoot,KeyName,InfoProjectNum,degree)
        RegisterDegreeWord(File,degree)
        FileOldRoot = FileList.pop(0)
        FileNewRoot = os.path.join(InfoDoneRoot,FileOldRoot.split('\\')[-1])
        shutil.move(FileOldRoot,FileNewRoot)
    
    
if __name__ == '__main__':
    main()