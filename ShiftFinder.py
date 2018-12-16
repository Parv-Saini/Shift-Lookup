#=====================
"""Imports"""
#=====================
import openpyxl
import datetime as dt
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as msg
import re

#======================
"""The initial class"""
#======================
class main_class():
    def __init__(self):         # Initializer method
        # Create instance
        self.win = tk.Tk()
        #Add Title
        self.win.title("Shift Finder")

        self.firstclick = True
        

        # defining Label Style in Frames      
        self.my_frame = ttk.Style()
        self.my_frame.configure('my.TLabel', font = ('Helvetica', 12, 'italic'))
        # defining Label Style outside the Frames
        self.my_label = ttk.Style()
        self.my_label.configure('my1.TLabel', font = ('Helvetica', 12))
        #Calling out the GUI method
        self.create_widgets()

    #---------------------------------------------------------------------------------
    #Method used to define global variables    
    def global_variables(self):
        #Lists to store all the column data of particular shifts
        self.s1_shift = []
        self.s2_shift = []
        self.s3_shift = []
        
        #Lists to store the Team members name
        self.s1=[]
        self.s2=[]
        self.s3=[]
        
    #---------------------------------------------------------------------------------        
    #Method used to let the user browse for Roster file
    def get_roster_file(self):
        
        self.root = filedialog.askopenfilename(initialdir = '/tmp', title = 'Select file', filetypes = [("Excel files","*.xlsx")])
        
    #---------------------------------------------------------------------------------
    #Method used to get System date        
    def get_today_date(self):
        self.today_date = str()        
        
        # Getting today's Date
        now = dt.datetime.now()
        get_date = str(now.day)
        get_today = str(get_date)                           #Converting today's Day to string
        if len(get_today)<2:
            self.today_date = str('0'+get_today)         #If today's date is single digit like "8" convert to "08"
        else:
            self.today_date=get_today

        self.get_shift_person()                                     #calling the method for loading the shift members data'
        row_number = 2
        self.click_me(row_number)
        self.button2.configure(state='disabled')        

    #---------------------------------------------------------------------------------        
    #Gets shift person for today
    def get_shift_person(self):
        self.global_variables()
        #=============================================================
        """Opening up the Roster file and selecting the work sheet"""
        #=============================================================
        try:
            roster_file = openpyxl.load_workbook(self.root)  
        except FileNotFoundError:
            msg.showerror('File not found error!', 'Please select the roster file first')
        except AttributeError:
            msg.showerror('File not found error!', 'Please select the roster file first')

        #Fetching the workbook provided by the user
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
                    if self.today_date in x:        #Matching today's date to cell date
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
            found=[]
            m2=str()                                        #Variable to hold the value of the Search function
            
            #Extracting the row numbers of all the elements in list
            for counter in d:
                m2 = re.search("(\d{1,2})>$", counter)      #matches any two 1 or 2 digit number before '>'
                if m2:
                    found.insert(0,m2.group(1))             #converted string to integer for comparison ahead

            count = int(1)
            
            #Extracting the shift member names and putting them in lists
            for counter in range(len(found)):
                row_number=int(found[counter])
                count=1
                for a in rost['A']:               
                    if count == row_number:
                        if shift is 1:
                            self.s1.append(a.value)
                        elif shift is 2:
                            self.s2.append(a.value)
                        else:
                            self.s3.append(a.value)
                    count = count + 1
            shift = shift + 1
            
        #====================================================
        """Getting the Shift Leads"""
        #====================================================
        self.shift_leads = []                           #List to store the name of Shift Leads
        
        for counter in range(3):                        #As last 3 rows of a column holds shift lead data
            lead_cell_location = len(rost[column_alphabet]) - counter
            lead_cell = column_alphabet + str(lead_cell_location)
            self.shift_leads.append(rost[lead_cell].value)          #Adding the shift lead name to the list
        
    #---------------------------------------------------------------------------------
    #Method used for displaying the data in GUI
    def click_me(self, row_variable):
        counter = 0
        #===============================================
        """Displaying Shift Members"""
        #===============================================

        for d in self.s1, self.s2, self.s3:                         #Iterating all the shift lists to get Shift Members
            
            frame_header = 'Shift - ' + str(counter + 1)
            new_frame = ttk.LabelFrame(self.tab1, text = frame_header)          #Creating shift frames dynamically
            new_frame.grid(column = counter, row = row_variable, padx = 10, sticky = tk.W)
            column_number = 1
            
            for row_number in range(len(d)):                                    #Creating the Data labels dynamically
                shift_member = str(d[row_number])
                self.data_label = ttk.Label(new_frame, text = shift_member, style = 'my.TLabel')
                self.data_label.grid(column = 0, row = column_number, sticky = 'W', padx = 15, pady = 4)
                column_number = column_number + 1

            counter = counter + 1

        counter = 0
        self.shift_leads.reverse()
        #========================================
        """Displaying Shift Leads"""
        #========================================
        for a in self.shift_leads:                                  #Iterating Shift Leads list to display all the shift leads
            frame_header = 'Shift - ' + str(counter+1)
            #Crating Frames
            new_frame = ttk.LabelFrame(self.tab2, text = frame_header)
            new_frame.grid(column = counter, row = row_variable, padx = 10, sticky = tk.W)
            #Adding labels to those Frames
            new_label = ttk.Label(new_frame, text = a, style = 'my.TLabel')
            new_label.grid(column = 0, row = counter)
            
            counter = counter + 1

    #---------------------------------------------------------------------------------
    #Method used to clear the Text Box when it is clicked            
    def on_entry_click(self, event):
            """function that gets called whenever entry1 is clicked"""        
            global firstclick

            if self.firstclick: # if this is the first time they clicked it
                self.firstclick = False
                self.text_box.delete(0, "end") # delete all the text in the entry
    
    #---------------------------------------------------------------------------------
    #Method used to display data as per the entered date            
    def get_custom_date(self):
        self.today_date = str()        
        
        x = self.date_entered.get()                      #Converting today's Day to string
        try:
            if len(x)==0:
                raise Exception("Null value in Textbox")
        except Exception:
            msg.showerror("Text Box Empty!", 'Please fill the Text Box')
            self.today_date = 'Null'
        if len(x)<2 and len(x)>=1:
            self.today_date = str('0'+x)         #If today's date is single digit like "8" convert to "08"
        elif len(x)==2:
            self.today_date=x
        row_number = 6
        self.get_shift_person()                                     #calling the method for loading the shift members data'
        self.click_me(row_number)     

    #---------------------------------------------------------------------------------
    #Method used to hold all initial GUI Widgets        
    def create_widgets(self):
        
        #Adding Tabs
        self.tab_control = ttk.Notebook(self.win)                              #Create tab control
        
        #Creating Shift Members tab
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab1, text = "Shift Members")
        
        #Creating Shift Leads tab
        self.tab2 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab2, text = "Shift Leads")
        
        self.tab_control.pack(expand = 1, fill = "both")                            #Packing the tabs to make them visible
        
        #Adding a Button for selecting roster file
        self.button1 = ttk.Button(self.tab1, text = "First select the Roster file", command = self.get_roster_file)
        self.button1.grid(column = 0, row = 0, padx = 20, pady = 20, sticky = 'WE', rowspan = 2, columnspan = 3)
        
        #Adding a Button for loading shift data
        self.button2 = ttk.Button(self.tab1, text = "Team Members in shift Today", command = self.get_today_date)
        self.button2.grid(column = 0, row = 3, padx = 20, pady = 20, sticky = 'WE', columnspan = 3)

        #Adding a Text Box for entering Date
        self.date_entered = tk.StringVar()
        self.text_box = ttk.Entry(self.tab1, textvariable = self.date_entered)
        #self.text_box.insert(0, 'Enter a date here eg: 18')
        self.text_box.bind('<FocusIn>', self.on_entry_click)
        self.text_box.grid(column = 0, row = 4, padx = 5, pady = 20, sticky = 'WE', columnspan = 2)
        
        #Adding a Button for loading shift data
        self.button3 = ttk.Button(self.tab1, text = "Get Data for Specified Date", command = self.get_custom_date)
        self.button3.grid(column = 2, row = 4, padx = 5, pady = 20, sticky = 'WE')
        
        #Adding warning Label
        self.warn = ttk.Label(self.tab1, text = "!!!Attention!!! - If no data is displayed then that Date is not present in Roster File", style = 'my1.TLabel')
        self.warn.grid(column = 0, row = 10, columnspan = 3, padx = 20, pady = 5)
        #self.action.focus()

#======================
# Start GUI
#======================
oop = main_class()
oop.win.mainloop()