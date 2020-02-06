"""Dependencies: pandas, pywin32, openpyxl
Uses pywin32 to make windows API calls to excel
Opens all excel files in specified location and searches tab for specified text
if worksheet matches search criteria, copy sheet to new workbook
save new workbook and reopen with pandas to concatenate all resulting data
"""

import tkinter as tk
import glob,os
from win32com.client import Dispatch #pip install pywin32
import pandas as pd
from tkinter import filedialog
import threading,queue

iQ = queue.Queue(1) #input path queue
oQ = queue.Queue(1) #output path
sQ =queue.Queue(1) #search term 
tQ = queue.Queue() #for updating status label

def main():
    ipath = iQ.get()
    opath = oQ.get()
    searchTerm = sQ.get()
    tQ.put("Working on it, don't press go again")
    xl = Dispatch("Excel.Application")
    try:
        xl.Visible=False   #if excel is already open, you'll see the macro running. hence, the warning in the gui 
        wb = xl.Workbooks.Add()
        files = glob.glob(str(ipath)+r'\*.xlsx')
        xlsfiles = glob.glob(str(ipath)+r'\*.xls')
        files.extend(xlsfiles)
        files_to_work = len(files) #will be used to feed how many files are left to look at to the gui
        xfiles=[] #list to identify the files that the tabs we are looking for came from
        n=1 #the number that gets appended to each tab 
        for i in files:
            tQ.put(f"looking through {files_to_work} files")
            wb1=xl.Workbooks.Open(Filename=i)
            for sh in wb1.Sheets:
                if searchTerm in sh.Name:
                    ws1 = wb1.Worksheets(sh.Name)
                    ws1.Name=f'{ws1.Name} - {str(n)}'
                    n+=1
                    ws1.Copy(Before=wb.Worksheets(1))
                    xfiles.append((os.path.basename(i),sh.Name))
                else:
                    continue
            wb1.Saved=True
            wb1.Close()
            files_to_work -=1
        wb.Worksheets('sheet1').Cells(1,1).Value='File Name'
        wb.Worksheets('sheet1').Cells(1,2).Value='Tab'
        row=2
        for i,j in xfiles:
            wb.Worksheets('sheet1').Cells(row,1).Value=i
            wb.Worksheets('sheet1').Cells(row,2).Value=j
            row+=1
        wb.SaveAs(f'{opath}\\{searchTerm}_Entity Analysis.xlsx')
        xl.Quit()
    except:
        xl.Quit()
        print('something went wrong...')
    tQ.put("Done finding the sheets, now building consolidated file")
    dflist=[]
    warnings=[]
    realcols = ['Line Number','Business Unit','Natural Account','Cost Center','Intercompany',
                'Product Line','Project','Branch','Growth Center','Reserve1','EnteredDR','EnteredCR',
                'Journal Entry Line Description','Context','Attribute1','Attribute2','Attribute3',
                'Attribute4','Attribute5','Attribute6']
    pdxl=pd.ExcelFile(opath+"\\"+searchTerm+"_Entity Analysis.xlsx")
    for i in pdxl.sheet_names: 
        df = pdxl.parse(sheet_name=i)
        for i, row in df.iterrows():
            if row.notnull().all():
                data = df.iloc[(i+1):].reset_index(drop=True)
                data.columns = list(df.iloc[i])
                break
        if len(list(df.columns)) != len(realcols):
            warn='WARNING The following sheet was not added to the consolidated file as it did not match the expected data: '+i
            print(warn)
            warnings.append(warn)
            continue
        elif list(df.columns) != realcols: #this was added for safety, incase someone accidentally mis-typed a column. may want to remove
            df.columns = realcols
            warn2 = f'WARNING The following sheet had its columns renamed: {i}'
            warnings.append(warn2)
        df.dropna(thresh=10,inplace=True)
        df.set_index(df.iloc[:,0],inplace=True)
        df.drop(index='Line Number',inplace=True)
        df['Tab'] = i
        dflist.append(df)
    df=pd.concat(dflist)
    
    refdf = pd.DataFrame(xfiles,columns = ['File Name','Tab'])
    df=df.merge(refdf,how='left',on='Tab')
    
    df2 = pd.DataFrame({'warnings' : warnings})
    
    writer = pd.ExcelWriter(path=f'{opath}\\{searchTerm}_Entity Analysis Consolidated.xlsx',engine='openpyxl')
    with writer as writer:
        df.to_excel(writer,sheet_name = 'Consolidation', index = False)
        df2.to_excel(writer,sheet_name = 'Warnings', index = False)
    #Use this code to insert the sheet into the first analysis file instead of generating a new workbook
    #with pd.ExcelWriter(opath+"\\"+searchTerm+"_Entity Analysis.xlsx",engine='openpyxl',mode='a') as writer:
    #	df.to_excel(writer,sheet_name='consol sheet')
    #status.config(text='Program Complete, Ready to go Again')
    tQ.put("Program complete, ready to go again")
    main()

#Code below handles GUI 
class App(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.start()

    def callback(self):
        self.root.quit()
    
    def run(self):
        def queup(i,o,s):
            q1 = i.get()
            q2 = o.get()
            q3 = s.get()
            iQ.put(q1)
            oQ.put(q2)
            sQ.put(q3)
        self.root = tk.Tk()
        self.root.protocol("WM_DELETE_WINDOW", self.callback)
        self.root.title('Entity Analysis Script')
        ipath_label = tk.Label(self.root,text='Input path')
        ipath_label.grid(row=1,column=1)
        self.ipath_entry = tk.Entry(self.root,width=50)
        self.ipath_entry.grid(row=1,column=2)
        st_label = tk.Label(self.root,text='Search')
        st_label.grid(row=3,column=1)
        self.st_entry = tk.Entry(self.root,width=50)
        self.st_entry.grid(row=3,column=2)
        opath_label = tk.Label(self.root,text='Output path')
        opath_label.grid(row=2,column=1)
        self.opath_entry = tk.Entry(self.root,width=50)
        self.opath_entry.grid(row=2,column=2)
        go = tk.Button(self.root,text='  Go  ',command=lambda:queup(self.ipath_entry,self.opath_entry,self.st_entry))
        go.grid(row=4,column=2)
        status = tk.Label(self.root,text='Not Started')
        status.grid(row=5,column=2)
        warntext2 ="""It is recommended that you close all excel files
        before running this program."""
        tk.Label(self.root,text=warntext2).grid(row=7,column=2)
        tk.Label(self.root,text='WARNING').grid(row=7,column=1)

        def path_browser(entrynum):
            entrynum.delete(0,'end')
            filepath = filedialog.askdirectory()
            filepath = os.path.normcase(filepath)
            entrynum.insert(0,filepath)
        ipath_button = tk.Button(self.root,text = 'browse',command = lambda x = self.ipath_entry:path_browser(x))
        ipath_button.grid(row = 1, column = 3)

        opath_button = tk.Button(self.root,text = 'browse',command = lambda x = self.opath_entry:path_browser(x))
        opath_button.grid(row = 2, column = 3)
        
        def getstatus():
            try:
                stat = tQ.get_nowait()
                status.config(text = stat)
            except:
                pass
            finally:
                self.root.after(100,getstatus)
        self.root.after(100,getstatus)
        self.root.mainloop()
app=App()
main()
