import configparser
import os
import platform
import time
import io
import csv
import traceback
import serial
#import uuid
import tkinter
import threading

from   tkinter import ttk
from   tkinter import messagebox
from   tkinter import filedialog
from   os.path import normpath
from   tksheet import Sheet

#---------------------- Fix local.DE -------------------
class locale:
    
    @staticmethod
    def atof(s):
        return float(s.strip().replace(',','.'))

    @staticmethod
    def format_string(fmt, *args):
        return  (fmt % args).replace('.',',')
    
#-------------------------------------------------------------------
# Define the GUI
#-------------------------------------------------------------------
class sheetapp_tk(tkinter.Tk):
    
    def __init__(self,parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()
        self.run        =  0
        self.xRow       = -1
        self.xCol       = -1
        self.xVal       = ''
        self.pending    = -1

    def showError(self, *args):
        err = traceback.format_exception(*args)
        messagebox.showerror('Exception',err)
        
        # but this works too
        tkinter.Tk.report_callback_exception = self.showError

    def initialize(self):
        noteStyle = ttk.Style()
        noteStyle.theme_use('default')
        noteStyle.configure("TNotebook", background='lightgray')
        noteStyle.configure("TNotebook.Tab", background='#eeeeee')
        noteStyle.map("TNotebook.Tab", background=[("selected", '#005fd7')],foreground=[("selected", 'white')])
        
        self.geometry("600x400")
        
        self.menuBar = tkinter.Menu(self)
        self.config(menu = self.menuBar)
        
        self.menuFile = tkinter.Menu(self.menuBar, tearoff=False)
        self.menuFile.add_command(command = self.saveSheet, label="Blatt speichern")
        self.menuFile.add_command(command = self.loadSheet, label="Blatt laden")
        self.menuFile.add_command(command = self.clearSheet, label="Blatt löschen")
        
        self.menuBar.add_cascade(label="Datei",menu=self.menuFile)
        
        self.pageHeader = tkinter.Label(self,text="Startnummer Eingabe",
                                        font=("Arial", 18),
                                        bg='#D3E3FD')
        self.pageHeader.pack(expand = 0, fill ="x") 
        
        self.tabControl = ttk.Notebook(self) 
        self.tabControl
          
        self.startTab   = ttk.Frame(self.tabControl) 
        self.tabControl.add(self.startTab, text ='Start') 
        
        self.tabControl.pack(expand = 1, fill ="both") 
         
        #----- Start Page -------
                 
        self.startTab.grid_columnconfigure(0, weight = 1)
        self.startTab.grid_rowconfigure(0, weight = 1)
        self.startSheet = Sheet(self.startTab,
                               name = 'startSheet',
                               #data = [['00:00:00','0,00','',''] for r in range(2)],
                               header = ['Uhrzeit','Zeitstempel','Startnummer','Slot'],
                               header_bg = "azure",
                               header_fg = "black",
                               index_bg  = "azure",
                               index_fg  = "gray",
                               font = ("Calibri", 12, "bold")
                            )
        self.startSheet.grid(column = 0, row = 0)
        self.startSheet.grid(row = 0, column = 0, sticky = "nswe")
        self.startSheet.span('A:').align('right')
        self.startSheet.span('A').readonly()
        if config.getboolean('view','hide_slots'):
            self.startSheet.hide_columns(3)
        self.startSheet.span('D').readonly()
        #self.startSheet.hide_columns(3)
        
        self.startSheet.disable_bindings("All")
        self.startSheet.enable_bindings("edit_cell","single_select","right_click_popup_menu",
                                        "drag_select","row_select","copy")
        self.startSheet.extra_bindings("end_edit_cell", func=self.startEndEditCell)
        
    def startEndEditCell(self, event):
        print("Start EndEditCell: ")
        
        for cell, value in event.cells.table.items():
            row = cell[0]
            col = cell[1]
            print(row,col,value)
            slot  = self.startSheet[row,3].data
            self.pending = row    
            self.startSheet.after_idle(self.sendStartMsg,"${:},{:}\r".format(value,slot))
                        
    def sendStartMsg(self,*args):
        if messagebox.askyesno("MODBUS", "Sende Startnummer zur Basis"):
            if len(args) == 1:
                print("Send: ",args[0])
                serialPort.write(args[0].encode())
                
    def getSelectedSheet(self):
        tab = self.tabControl.tab(self.tabControl.select(),"text")
        if tab == "Start":
            return self.startSheet
        
    def saveSheet(self):
        saveSheet = self.getSelectedSheet()
        print("Save: "+saveSheet.name)
        # create a span which encompasses the table, header and index
        # all data values, no displayed values
        sheet_span = saveSheet.span(
            header=True,
            index=True,
            hdisp=False,
            idisp=False,
        )
        
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Save sheet as",
            filetypes=[("CSV File", ".csv"), ("TSV File", ".tsv")],
            defaultextension=".csv",
            confirmoverwrite=True,
        )
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(
                    fh,
                    dialect=csv.excel if filepath.lower().endswith(".csv") else csv.excel_tab,
                    lineterminator="\n",
                )
                writer.writerows(sheet_span.data)
        except Exception as error:
            print(error)
            return

    def loadSheet(self):
        loadSheet = self.getSelectedSheet()
        print("Load: "+loadSheet.name)
        
        sheet_span = loadSheet.span(
            header=True,
            index=True,
            hdisp=False,
            idisp=False,
        )
        
        filepath = filedialog.askopenfilename(parent=self, title="Select a csv file")
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "r") as filehandle:
                filedata = filehandle.read()
            loadSheet.reset()
            sheet_span.data = [
                r
                for r in csv.reader(
                    io.StringIO(filedata),
                    dialect=csv.Sniffer().sniff(filedata),
                    skipinitialspace=False,
                )
            ]
        except Exception as error:
            print(error)
            return
        
    def clearSheet(self):
        tab = self.tabControl.index(self.tabControl.select())  
        if messagebox.askyesno("Start/Ziel", "Alle Daten löschen ?"):
            print("Clear sheet:",tab)
            if tab == 0:
                self.startSheet.deselect()
                self.startSheet.data = []
    
    def resendData(self):
        row = self.startSheet.get_currently_selected().row
        print("Resend: ",row)
        num  = self.startSheet[row,2].data.strip()
        slot = self.startSheet[row,3].data.strip()
        self.pending = row    
        self.startSheet.after_idle(self.sendStartMsg,"${:},{:}\r".format(num,slot))
        
#-------------------------------------------------------------------

    
#-------------------------------------------------------------------
# Main program
#-------------------------------------------------------------------

if __name__ == '__main__':    
   
    myPlatform = platform.system()
    print("OS in my system : ", myPlatform)
    myArch = platform.machine()
    print("ARCH in my system : ", myArch)

    config = configparser.ConfigParser()
   
    # Defaults Linux
    config['serial'] = {'enabled':'no',
                        'port':'/dev/ttyUSB0',
                        'baud':'115200',
                        'timeout':'10'}
    
    config['view'] = {'hide_slots':'no'}
     
    # Platform specific
    if myPlatform == 'Windows':
        # Platform defaults
        config.read('windows.ini') 
    if myPlatform == 'Linux':
        config.read('linux.ini')

    # ---------- setup and start GUI --------------
    app = sheetapp_tk(None)
    
    # Initialize the port    
    serialPort = serial.Serial(config.get('serial', 'port'),
                               config.getint('serial', 'baud'), 
                               timeout=config.getint('serial', 'timeout'))
    
    # Function to call whenever there is data to be read
    def readFunc(port):
        while True:
            try:
                line = port.readline().decode("utf-8").strip() 
                if (len(line) > 0) and line[0] == '$':
                    app.startSheet.after_idle(insertStamp,line[1:])
                elif (len(line) > 0) and line[0] == '!':
                    app.startSheet.after_idle(processMessage,line[1:])
                elif (len(line) > 0) and line[0] == '?':
                    print("Read msg: ", line[1:])
            except Exception as e:
                print("EXCEPTION in readline: ",e) 
        
        print("DONE readline")
     
    def insertStamp(line):  
        print("Insert data: ",line)
        data = line.split(',')
        if len(data) == 2:
            stamp = int(data[0])
            slot  = data[1].strip()
            app.startSheet.insert_row([time.strftime('%H:%M:%S'),
                                       "{:0.2f}".format(stamp/100).replace('.',','),
                                       '0',slot]) 
            row = app.startSheet.get_currently_selected().row
            app.startSheet[row].highlight(bg='#D3E3FD')
            app.startSheet.see(row)   
     
    def processMessage(line): 
        if app.pending >= 0:
            print("Read msg: ", line)
            if line == "AKN":
                app.startSheet[app.pending].highlight(bg='aquamarine')
                app.pending = -1
            if line == "NAK":
                app.startSheet[app.pending].highlight(bg="pink")
                app.pending = -1    
                          
    # Configure threading
    usbReader = threading.Thread(target = readFunc, args=[serialPort])
    usbReader.start()
    
    app.title("MODBUS Start/Ziel Tabelle Elz-Zeit")
    
    app.startSheet.popup_menu_add_command(
        "Clear sheet data",
        app.clearSheet,
    )
    
    app.startSheet.popup_menu_add_command(
        "Resend data",
        app.resendData,
    )
    
    # run
    app.mainloop()
    print(time.asctime(), "GUI done")
         
    # Stop all dangling threads
    os.abort()