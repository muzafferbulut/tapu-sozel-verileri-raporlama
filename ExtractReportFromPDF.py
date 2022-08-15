# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22 15:35:32 2022

@author: Muzaffer Bulut | Harita Mühendisi
"""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from os import listdir
from openpyxl import Workbook
import tabula
    
class Tapu(tk.Tk):
    
    def __init__(self):
        
        super().__init__()
        
        # pencere örneğinin geometrisi
        self.title("Tapu Raporlama")
        self.geometry("380x250+300+100")
        self.resizable(False, False)
        
        # pencere örneğindeki label nesnelerinin konumlandırılması
        self.openFileLabel = ttk.Label(self, text= "Open File Directory :")
        self.openFileLabel.place(x=20, y=25)
        
        self.saveToLabel = ttk.Label(self, text= "Extract and Save to :")
        self.saveToLabel.place(x=20, y=120)
        
        # kaynak ve hedef dosya yolarını içeren entry yapılarının konumlandırılması
        self.directoryNameEntry = ttk.Entry(self, width=40)
        self.directoryNameEntry.place(x=20, y=50)
        
        self.saveToEntry = ttk.Entry(self, width=40)
        self.saveToEntry.place(x=20, y=145)
        
        # buton yapılarının konumlandırılamsı
        self.askOpenFileButton = ttk.Button(self, text="Open")
        self.askOpenFileButton['command'] = self.askOpenFileDirectory
        self.askOpenFileButton.place(x=275, y=48)
        
        self.saveToButton = ttk.Button(self, text="Save")
        self.saveToButton['command'] = self.saveTo
        self.saveToButton.place(x=275, y=143)
        
        self.getReportButton = ttk.Button(self, text="Get Report")
        self.getReportButton['command'] = self.getReport
        self.getReportButton.place(x=275, y=190)
    
    def askOpenFileDirectory(self):
        """
        HELP:
            
        """
        global nameDir
        nameDir = filedialog.askdirectory()
        nameDir = nameDir + "/"
        self.directoryNameEntry.insert(0, nameDir)
    
    def saveTo(self):
        """
        HELP:
            
        """
        global saveTo
        saveTo = filedialog.asksaveasfilename(defaultextension="*.xlsx", filetypes=(("Excel Files", "*.xlsx"),("All Files", "*.*")))
        self.saveToEntry.insert(0, saveTo)
    
    def createEmptyReport(self):
        """
        HELP:
            NetCAD yazılımının NetMAP modülünde tapu sözel verileri okuma ve eşleştirme
            aşamasında istenilen verileri şablon olarak kabul eden ve bu şablona göre
            bir rapor örneği oluşturur.        
            """
            
        emptyReport = Workbook()
        reportPage = emptyReport.active
        reportPage.title = "Tapu Sözel Verileri Raporu"
        
        reportPage["A1"] = "TaşınmazID"
        reportPage["B1"] = "İl / İlçe"
        reportPage["C1"] = "Mahalle / Köy Adı"
        reportPage["D1"] = "Mevkii"
        reportPage["E1"] = "Cilt / Sayfa"
        reportPage["F1"] = "Kayıt Durum"
        reportPage["G1"] = "Ada / Parsel"
        reportPage["H1"] = "Yüzölçüm (m2)"
        reportPage["I1"] = "Ana Taşınmaz Nitelik"
        reportPage["J1"] = "Malik"
        reportPage["K1"] = "Pay / Payda"
        reportPage["L1"] = "SerhBeyanTip"
        reportPage["M1"] = "SerhBeyanMetin"
        return emptyReport
    
    def getFileName(self, path):
        """
        HELP:
            path değişkeni olarak verilen bir dosya yolundaki bütün pdf dosyalarını
            okur ve isim listesini geri döndürür.
        """
        
        allFiles = listdir(path)
        files = []
        
        for file in allFiles:
            
            if file.endswith("pdf"):
                files.append(file)
        return files
    
    def generateColumns(self, i):
        """
        HELP:
            Verileri excele raporlamada sütun listesi ile satır sayısını eşleştirir
            ve yazma işleminde kullanmak üzere yeni bir liste oluşturarak geri
            döndürür.
        """
        col_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M"] 
        generatedCol = []
        for let in col_list:
            
            generatedCol.append(let+str(i))
        
        return generatedCol
    
    def setRealEstateInfos(self, df, colList):
        """
        HELP:
            Dataframe yapısında dağınık bir şekilde tutulan taşınmaz bilgilerini
            ayrıştırarak excel raporuna ekler.
        """
        titleList = df.keys()
        report[colList[6]] = titleList[3]
        
        report[colList[0]] = df.values[0,1]
        report[colList[1]] = df.values[1,1]
        report[colList[2]] = df.values[3,1]
        report[colList[3]] = df.values[4,1]
        report[colList[4]] = df.values[5,1]
        report[colList[5]] = df.values[6,1]
        report[colList[7]] = df.values[0,3]
        report[colList[8]] = df.values[1,3]       
    
    def setOwnerInfos(self, df, colList):
        """
        HELP:
            Dataframe yapısında dağınık bir şekilde tutulan malik bilgilerini
            ayrıştırarak excel raporuna ekler.
        """
        report[colList[9]] = df['Malik'].values[0]
        report[colList[10]] = df['Pay / Payda'].values[0]
    
    def setSerhBeyanInfos(self, df, colList):
        """
        HELP:
            Dataframe yapısında dağınık bir şekilde tutulan serh beyan bilgilerini
            ayrıştırarak excel raporuna ekler.
        """
        report[colList[11]] = df['Tip'].values[0]
        report[colList[12]] = df['Ş.B.İ. Metin'].values[0]
    
    def getReport(self):
        """
        HELP:
            PDF olarak temin edilen tapu sözel verilerini excel formatına
            otomatik olarak dönüştürür. Dönüştürme işleminde 
        """
        global excelReport, report
        excelReport = self.createEmptyReport()
        report = excelReport.active
        
        files = self.getFileName(nameDir)
        
        ownerCounter = 2 # excele 2. satırdan yazmaya başlayacağım için 2 oldu
        
        for file in files:
            
            filePath = nameDir + file
            
            tapu = tabula.read_pdf(filePath)
            colList = self.generateColumns(ownerCounter)
            
            self.setRealEstateInfos(tapu[0], colList)
            
            if len(tapu) == 2:
                """
                Şerh ve beyan bilgisi bu tapu belgesi için yoktur. Sadece malik
                ve taşınmaz bilgileri yer almaktadır.
                """
                self.setOwnerInfos(tapu[1], colList)
            
            elif len(tapu) > 2:
                """
                Taşınmaz ve malik bilgileri yanında şerh beyan bilgileri içeren
                tapu belgesidir.
                """
                for i in range(1, len(tapu)):
                    
                    df = tapu[i]
                    
                    if df.keys()[0] == 'SistemNo':
                        self.setOwnerInfos(df, colList)
                    
                    elif df.keys()[0] == 'Tip':
                        self.setSerhBeyanInfos(df, colList)
                    
                    else:
                        messagebox.showerror("Error!","Hata Kodu 2")
            else:
                messagebox.showerror("Error!","Hata Kodu 1")
                
            ownerCounter += 1
        
        excelReport.save(saveTo)
        
        messagebox.showinfo("Successfully!", "Completed!")
    
if __name__ == "__main__":
    App = Tapu()
    App.mainloop()