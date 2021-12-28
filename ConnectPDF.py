import time
import os
from PyPDF2 import PdfFileMerger
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox
import re
import glob

class Source(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.active_frame = None
        self.swich_frame(MainMenu)
        #Global variable
        self.outputFileName = ""
        self.outputFolderName = ""

    def swich_frame(self, _class):
        self.newFrame = _class(self)
        if self.active_frame is not None:
            self.active_frame.destroy()
        self.active_frame = self.newFrame
        self.active_frame.pack()
        self.title("Connect PDF")

    def exitProgram(self):
        tk.Tk.destroy(self)

    def resizeWindow(self, width, height):
        self.widthScreen = self.winfo_screenwidth()
        self.heightScreen = self.winfo_screenheight()

        self.x = (self.widthScreen / 2) - (width / 2)
        self.y = (self.heightScreen / 2) - (height / 2)

        self.geometry("{}x{}+{}+{}".format(width, height, int(self.x), int(self.y)))

class MainMenu(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        #size
        master.resizeWindow(300,200)
        #button
        self.connectPDF_button = tk.Button(self, text="ConnectPDF",
                                           command = lambda: master.swich_frame(CreatorMenu), padx=10, pady=10)
        self.exit_button = tk.Button(self, text="Exit", command=master.exitProgram, padx=10, pady=10)
        #label
        self.title_table = tk.Label(self, text="ConnectPDFConventer")
        #pack
        self.title_table.pack()
        self.connectPDF_button.pack(padx=10, pady=10, side="top", fill='x')
        self.exit_button.pack(padx=10, pady=10, side="top", fill='x')

class CreatorMenu(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        #size
        master.resizeWindow(1000,600)
        #globalVariable
        self.file_list = []
        self.countID = 0
        self.master.bind('<Control-a>', lambda *args: self.treeview.selection_add(self.treeview.get_children()))
        #Frame
        self.treeview_frame = tk.Frame(self)
        self.folderName_frame = tk.LabelFrame(self, text='Select a folder where you want to save')
        self.fileName_frame = tk.LabelFrame(self, text='Enter a filename of PDF')
        self.optionButton_frame = tk.LabelFrame(self, text='Options')
        #ScrollBar
        self.treeviewScrollBar = tk.Scrollbar(self.treeview_frame)
        #Style
        self.style = ttk.Style()
        self.style.theme_use("xpnative")
        self.style.configure('Treeview',
                             background='white',
                             foreground='black',
                             rowheight=25,
                             fieldbackgound='silver')
        self.style.map('Treeview',
                       background=[('selected','red')])
        #TreeView
        self.treeview = ttk.Treeview(self.treeview_frame, yscrollcommand=self.treeviewScrollBar.set)
        self.treeview['columns'] = ("FileName", "FilePath")
        self.treeview.column("#0", width=0, stretch="NO") #white space no visible ???
        self.treeview.column("FileName", anchor='w', width=300)
        self.treeview.column("FilePath", anchor='w', width=600)
        self.treeview.heading("FileName", text="FileName", anchor='w')
        self.treeview.heading("FilePath", text="FilePath/Status", anchor='w')
        self.treeviewScrollBar.config(command=self.treeview.yview)
        self.treeview.tag_configure('odd', background="white")
        self.treeview.tag_configure('even', background="lightblue")
        #Entry
        self.folderName_entry = tk.Entry(self.folderName_frame, width=50, borderwidth=3)
        self.fileName_entry = tk.Entry(self.fileName_frame, width=50, borderwidth=3)
        #Buttons
        #TODO: Add moving buttons
        # add functions to move file in tree view by drag ???
        self.add_button = tk.Button(self.optionButton_frame, text="Add File", command=self.addFile, padx=10, pady=10)
        self.delete_button = tk.Button(self.optionButton_frame, text="Delete File", command=self.remove, padx=10, pady=10)
        self.connectPDF_button = tk.Button(self.optionButton_frame, text="ConnectPDF", command=self.connectPDF,
                                           padx=10, pady=10, bg='red', fg='white')
        self.browse_button = tk.Button(self.folderName_frame, text="Browse", command=self.selectPath)
        #Progres bar
        self.progresBar = ttk.Progressbar(self.optionButton_frame, orient="horizontal", length=100, mode='determinate')
        #pack or grid
        self.treeview_frame.grid(column=0,row=0, columnspan=2, padx=10, pady=10)
        self.treeviewScrollBar.pack(side="right", fill="y")
        self.treeview.pack()
        self.folderName_frame.grid(column=0, row=1, padx=10, pady=10)
        self.fileName_frame.grid(column=0, row=2, padx=10, pady=20)
        self.folderName_entry.grid(row=0, column=0, padx=10, pady=15, ipadx=5)
        self.browse_button.grid(row=0, column=1, padx=10, pady=15)
        self.fileName_entry.pack(padx=10, pady=15, ipadx=38)
        self.optionButton_frame.grid(column=1, row=1, rowspan=2, padx=10, pady=10)
        self.add_button.pack(padx=10, pady=10, side="top", fill="x")
        self.delete_button.pack(padx=10, pady=10, side="top", fill="x")
        self.connectPDF_button.pack(padx=10, pady=10, side="top", fill="x")
        self.progresBar.pack(padx=10, pady=10, expand=True)

    def selectPath(self):
        self.folderName_entry.delete(0, "end")
        self.folderName_entry.insert(0, fd.askdirectory())
        self.master.outputFolderName = self.folderName_entry.get()

    def addFile(self):
        fileFullPath = fd.askopenfilenames(filetypes=[('PDF files', '.pdf')])
        for x in fileFullPath:
            fileName = x.split("/")[-1]
            if(len(self.treeview.get_children()) % 2 == 0):
                self.treeview.insert(parent='', index='end', iid = self.countID, text='',
                                 values=(fileName, x), tags=('even',))
            else:
                self.treeview.insert(parent='', index='end', iid = self.countID, text='',
                                     values=(fileName, x), tags=('odd',))
            self.countID += 1

    def remove(self):
        delete = self.treeview.selection()
        for x in delete:
            self.treeview.delete(x)

    def checkIfEntryFolderNameIsFilled(self):
        if not self.folderName_entry.get():
            messagebox.showerror("Erorr!", "Enter a folder path")
            return False
        return True

    def checkIfEntryFileNameIsFilled(self):
        if not self.fileName_entry.get():
            messagebox.showerror("Erorr!", "Enter a file name")
            return False
        return True

    def checkIfAtLeastTwoPDFFileSelected(self, curItems):
        if (len(curItems) < 2):
            messagebox.showerror("Erorr!", "Add at least two PDF file")
            return False
        return True

    def checkIfFileNameAlreadyExists(self):
        for PDFfilename in glob.glob("{}/*.pdf".format(self.folderName_entry.get())):
            if PDFfilename == "{}\{}.pdf".format(self.folderName_entry.get(), self.fileName_entry.get()):
                return True
        return False

    def checkIfCanConnectPDF(self, curItems):
        if not self.checkIfEntryFolderNameIsFilled():
            return False
        if not self.checkIfEntryFileNameIsFilled():
            return False

        if not self.checkIfAtLeastTwoPDFFileSelected(curItems):
            return False

        if self.checkIfFileNameAlreadyExists():
            answer = messagebox.askyesno(title="Warning!",
                                         message="This filename already exists. Do you want to overwrite?")
            if not answer:
                return False

        return True

    def updateProgessBar(self, progresBarrCounter):
        self.update_idletasks()
        self.progresBar['value'] += progresBarrCounter
        time.sleep(0.2)

    def connectPDF(self):
        #Select Section
        self.treeview.selection_add(self.treeview.get_children()) # in order to know witch files convert (method to save filename in list)
        curItems = self.treeview.selection()

        #Check Condition Section
        if not self.checkIfCanConnectPDF(curItems):
            for x in curItems:
                self.treeview.selection_remove(x)
            return

        # ConnectPDF
        self.file_list = [] # in order to be 100% sure that when right before convert the list is empty
        for x in curItems:
            self.file_list.append(self.treeview.item(x)) #append to list path of PDF file

        merger = PdfFileMerger()
        progresBarrCounter = 100/(len(self.file_list)+1) # plus 1 in order to left one pice of progresbar to
        # merger.write() process

        for count, pdf in enumerate(self.file_list):
            try:
                merger.append(open(pdf['values'][1], 'rb'))
            except:
                self.treeview.item(count, values=(pdf['values'][0], "ERROR")) # in order to show which file is incorret
                messagebox.showerror(title="Oh no!", message="Your PDF file are not convertable!")
                self.progresBar['value'] = 0
                return
            self.updateProgessBar(progresBarrCounter)

        merger.write("{}/{}.pdf".format(self.folderName_entry.get(), self.fileName_entry.get()))
        merger.close()
        self.updateProgessBar(progresBarrCounter)

        self.master.swich_frame(AfterConnectMenu)


class AfterConnectMenu(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        #size
        master.resizeWindow(400, 150)
        #Label
        self.title_label = tk.Label(self, text="The PDF file merging finish!")
        #Button
        self.showFolder_button = tk.Button(self, text='Show Folder', command=self.showFolder)
        self.returnMainMenu_button = tk.Button(self, text='Exit to MainMenu',
                                               command=lambda: master.swich_frame(MainMenu))
        #grid
        self.title_label.pack(padx=10, pady=10)
        self.showFolder_button.pack(padx=10, pady=10, side="top", fill="x")
        self.returnMainMenu_button.pack(padx=10, pady=10, side="top", fill="x")

    def showFolder(self):
        path = "{}".format(self.master.outputFolderName)
        path = os.path.realpath(path)
        os.startfile(path)

#Main
app = Source()
app.mainloop()