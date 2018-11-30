#this script is used to create the user interface
import tkinter as tk
from tkinter import messagebox as mBox
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
import processor, os, threading

class Application(Frame):

    def __init__(self,master):

        #setting up all the initial variables

        Frame.__init__(self,master)
        self.master = master

        self.orgName = tk.StringVar()
        self.year = tk.StringVar()
        self.month = tk.StringVar()
        self.fileName = tk.StringVar()

        self.open_file_path1 = tk.StringVar()
        self.open_file_path2 = tk.StringVar()
        self.open_file_path3 = tk.StringVar()

        #this app will savae converted file automatically to the ADD_FILES folder
        self.save_file_path = 'B:\\HARDI\\SQL\\ADD_Files\\'

        self.stat_message = tk.StringVar()
        self.result_message = tk.StringVar()

        self.file_names = []

        #the following is used to create the search bar,
        #now it's easier to find distributors

        self.distributorLabel = tk.Label(self.master,text = 'Distributor\'s Name',font = 14).grid(row = 0, column = 2,columnspan = 3)

        self.distributors = sorted(['2J Supply','ABR Wholesalers','AC Pro','ACR Supply Company','API of NH','APR Supply',   \
        'Aireco Supply','Airefco','Associated Equipment Company','Behler-Young','Benoist Brothers','CFM Equipment Distributors',   \
        'Capitol Group','Century AC Supply','Charles D Jones Company','Comfort Supply','Corken Steel Products Company',   \
        'Duncan Supply Company','Ferguson','G.W. Berkheimer Company','Geary Pacific Supply','Gustave A Larson Company',   \
        'HVAC Distributors','HVAC Sales & Supply Company','Illco','Interline Brands','Johnson Supply','Johnstone Supply - Popma',   \
        'Key Refrigeration Supply','Luce Schwab & Kase','McCall\'s Supply Company','Meier Supply','Mingledorff\'s',   \
        'Minnesota Air','Morsco','NB Handy','O\'Connor Company','RE Michel','Refrigeration Sales Corp','Sid Harvey\'s',   \
        'Sigler Wholesale Distributors','South Central Company','Standard Supply','Star Supply Company','Team Air Distributing',   \
        'The Granite Group','Thos. Somerville Company','Total Air Supply','Weathertech Distributing Co','cfm Distributors Inc',   \
        'Munch Supply Company','Famous Supply','Crescent Parts','American Air Distributing','Lohmiller & Company','M&A Supply',   \
        'Auer Steel','Temperature Systems','Best Choice','Design Air','Williams Distributing','Robert Madden','Peirce Phelps',   \
        'Winsupply','Ed\'s Supply','Carr Supply','Capco Energy Supply','Robertson Heating Supply','Comfort Air Distributing',   \
        'Hercules Industries','Koch Air','Emerson','Weil-Mclain','Monroe Equipment','Locke Supply','Heating and Cooling Supply Co. Inc.',   \
        'Republic Companies','Dunphey and Associates Supply Co.','Shearer Supply','Distributor Corporation of New England',   \
        'P&N Distribution Corp','Standard Air & Lite','Owens Corning'],key=str.lower)

        #this is where you enter year and month
        self.yearLabel = tk.Label(self.master,text = 'Year',font = 14).grid(row = 3, column = 2)
        self.entryYear = tk.Entry(self.master)
        self.entryYear.grid(row = 4, column = 2)

        self.monthLabel = tk.Label(self.master,text = 'Month',font = 14).grid(row = 5, column = 2)
        self.entryMonth = tk.Entry(self.master)
        self.entryMonth.grid(row = 6, column = 2)

        #call create_widgets function to create other widgets
        self.create_widgets()

    #this function is used to open the Submission file
    #you can select up to 3 files
    #i hard coded this part because 3 is the most we have
    def open_file(self,*args):

        self.pathLabel1.config(text=' ')
        self.pathLabel2.config(text=' ')
        self.pathLabel3.config(text=' ')

        self.file_names = filedialog.askopenfilenames(initialdir='B:\\HARDI\\Participants\\Submissions')

        if len(self.file_names)==1:
            self.open_file_path1=self.file_names[0]
            self.pathLabel1.config(text=os.path.split(self.open_file_path1)[1])

        elif len(self.file_names)==2:
            self.open_file_path1=self.file_names[0]
            self.pathLabel1.config(text=os.path.split(self.open_file_path1)[1])

            self.open_file_path2=self.file_names[1]
            self.pathLabel2.config(text=os.path.split(self.open_file_path2)[1])

        elif len(self.file_names)==3:
            self.open_file_path1=self.file_names[0]
            self.pathLabel1.config(text=os.path.split(self.open_file_path1)[1])

            self.open_file_path2=self.file_names[1]
            self.pathLabel2.config(text=os.path.split(self.open_file_path2)[1])

            self.open_file_path3=self.file_names[2]
            self.pathLabel3.config(text=os.path.split(self.open_file_path3)[1])

        else:
            return

    #this function is used to create more widgets like the search bar on the
    #top of the UI
    def create_widgets(self):

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.update_list)
        self.entry = Entry(self.master, textvariable = self.search_var, width = 45)
        self.lbox = Listbox(self.master, width = 45, height = 10,exportselection = 0)

        self.entry.grid(row = 1, column = 2, padx = 10, pady = 3)
        self.lbox.grid(row = 2, column = 2, padx = 10, pady = 3)

        self.openFileButton = tk.Button(text='Open File', command=self.open_file).grid(row = 7, column = 2)
        self.pathLabel1 = tk.Label(self.master)
        self.pathLabel1.grid(row = 8, column = 1,columnspan = 3)
        self.pathLabel2 = tk.Label(self.master)
        self.pathLabel2.grid(row = 9, column = 1,columnspan = 3)
        self.pathLabel3 = tk.Label(self.master)
        self.pathLabel3.grid(row = 10, column = 1,columnspan = 3)

        #the progress bar that got displayed under file names
        self.progress_bar = ttk.Progressbar(self.master,orient = tk.HORIZONTAL,
                            mode = 'indeterminate',
                            takefocus = True)
        self.progress_bar.grid(row = 11, column = 1,columnspan = 3)

        self.startButton = Button(text = 'Start', command = self.convert).grid(row = 12, column = 2)

        self.update_list()

    #the function that startButton uses
    def convert(self,*args):

        try:
            self.orgName = self.lbox.get(self.lbox.curselection())
        except:
            mBox._show('ERROR','Please Select a Distributor')
            return

        self.year = self.entryYear.get()
        if self.year == '':
            mBox._show('ERROR','Please Enter Year')
            return

        self.month = self.entryMonth.get()
        if self.month == '':
            mBox._show('ERROR','Please Enter Month')
            return

        #in order to have the progress bar running, we have to create thread
        self.show_progress(True)

        self.thread = threading.Thread(target = self.convert_worker)
        self.thread.daemon = True
        self.thread.start()

        #start the thread and check it
        self.convert_check()

    #this is where all heavy lifting begins
    def convert_worker(self,*args):

        #construct the saving file path
        save_path = self.save_file_path+'c'+self.orgName+' '+self.year+'-'+self.month+'.csv'

        self.year=int(self.year)
        self.month=int(self.month)

        # call the process function to start converting
        self.stat_message, self.result_message=processor.process(self.orgName,self.year,self.month,self.file_names,save_path)

        #if everything runs well, you should be able to see the message box
        mBox._show('Result','This is the result \n: %s \n %s' %(self.stat_message,self.result_message))

    #this is used to check the thread
    #if its still running, check after 10 milliseconds
    def convert_check(self):
        if self.thread.is_alive():
            self.after(10,self.convert_check)
        else:
            self.show_progress(False)

    #this function is used to update the search function
    def update_list(self,*args):

        search_term = self.search_var.get()
        lbox_list = self.distributors

        self.lbox.delete(0, END)

        for item in lbox_list:
            if search_term.lower() in item.lower():
                self.lbox.insert(END, item)

    #this is how the progress bar moves
    def show_progress(self,start):

        if start:
            self.progress_bar.start()
        else:
            self.progress_bar.stop()


root = tk.Tk()
root.geometry('300x500')
root.title('HARDI Data Submission')
app = Application(master = root)
app.mainloop()
