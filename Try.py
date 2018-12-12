#=====================
"""Imports"""
#=====================
import openpyxl
import datetime as dt
import tkinter as tk
from tkinter import ttk
import re

#======================
"""The initial class"""
#======================
class main_class():
    def __init__(self):         # Initializer method
        # Create instance
        self.win = tk.Tk()
        
        #Add Title
        self.win.title("Python GUI")
        
        # Getting today's Date
        now = dt.datetime.now()
        get_date = str(now.day)
        get_today = str(get_date)                           #Converting today's Day to string
        if len(get_today)<2:
            self.get_today_date = str('0'+get_today)         #If today's date is single digit like "8" convert to "08"
        else:
            self.get_today_date=get_today         
        # defining column variable
        self.column = str()
        print(self.get_today_date)
        
        #Lists to store all the column data of particular shifts
        self.s1_shift = []
        self.s2_shift = []
        self.s3_shift = []
        
        #Lists to store the Team members name
        self.s1=[]
        self.s2=[]
        self.s3=[]
        
        #Calling out the methods
        self.create_widgets()                
        self.get_shift_person()

    #Get's shift person for today
    def get_shift_person(self):
        #=============================================================
        """Opening up the Roster file and selecting the work sheet"""
        #=============================================================
        roster_file = openpyxl.load_workbook("Roster.xlsx")
        rost = roster_file['Sheet2']

        #============================================
        """Scanning the work sheet for today's date"""
        #============================================
        for index, row in enumerate(rost.iter_rows()):
            for cell in row:
                counter = "False"
                cell_data=str(cell.value)
                c=re.search("-(\d{1,2}) 00:",cell_data)
                if c:
                    x = c.group(1)
                    if self.get_today_date in x:        #Matching today's date to cell date
                        self.date_column = str(cell)
                        counter = "True"
                if counter=='True':                     #Breaking out of the inner for loop
                    break
            if counter=='True':                         #Breaking out of the outer for loop
                break
        #==========================================================
        """Scanning all the shifts in the matched today's column"""
        #==========================================================
        column_header = re.search("'Sheet2'.(.+?)\d>", self.date_column)    #Extracting the particular column alphabet from matched cell
        row = str()
        if column_header:
            column_alphabet = column_header.group(1)
        for a in rost[column_alphabet]:
            x = a.value
            if x=='S1':                             # checking with the shift S1
                self.s1_shift.append(str(a))
            if x=='S2':                             # checking with the shift S2
                self.s2_shift.append(str(a))
            if x=='S3':                             # checking with the shift S3
                self.s3_shift.append(str(a))
        shift = 1
        #===========================================================       
        """Loop Through each list and fetch out the shift members"""
        #===========================================================
        for d in self.s1_shift, self.s2_shift, self.s3_shift:
            found3=[]
            m2=str()                                    #Variable to hold the value of the Search function
            #print("\nShift S"+str(shift)+" members are:")
            for x3 in d:
                m2 = re.search("(\d{1,2})>$", x3)       #matches any two 1 or 2 digit number before '>'
                if m2:
                    found3.insert(0,m2.group(1))        #converted string to integer for comparison ahead
            count = int(1)
            ab=len(found3)
            for x4 in range(ab):
                abc=int(found3[x4])
                count=1
                for a in rost['A']:               
                    if count==abc:
                        if shift is 1:
                            self.s1.append(a.value)
                        elif shift is 2:
                            self.s2.append(a.value)
                        else:
                            self.s3.append(a.value)
                    count=count+1
            shift = shift+1

    def click_me(self):
        #action.configure(text="Hello"+name.get())
        #====================================
        """loads the GUI when event occurs"""
        #====================================
        head = ['Shift1', 'Shift2', 'Shift3']
        for col in range(3):                                                    #Creating the Header labels in loop
            self.lab = tk.Label(self.win, text=head[col])
            self.lab.grid(column = col, row = 1,sticky = tk.W + tk.E)
        
        timings = ['[06:00 to 15:00]', '[14:00 to 23:00]', '[22:00 to 07:00]']
        for col in range(3):
            self.lab2 = tk.Label(self.win, text=timings[col])
            self.lab2.grid(column = col, row = 2, sticky = tk.W + tk.E)
        
        column_number=0
        for d in self.s1, self.s2, self.s3:
            for row_number in range(len(d)):                                    #Creating the Data labels in loop
                a=str(d[row_number])
                self.data_label = tk.Label(self.win, text=a)
                self.data_label.grid(column = column_number, row = row_number+3,sticky = tk.W + tk.E)
            column_number = column_number+1
            
    def create_widgets(self):
           
        #Adding a Button
        self.action = ttk.Button(self.win, text="CLick here to load the data", command=self.click_me)
        self.action.grid(columnspan = 3,sticky = tk.W)
        
        self.action.focus()        
#======================
# Start GUI
#======================
oop = main_class()
oop.win.mainloop()