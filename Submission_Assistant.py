import tkinter as tk
from tkinter import messagebox as mBox
from tkinter import filedialog
import processor
import os
import sys

root = tk.Tk()
root.iconbitmap('B:\\\HARDI\\\Administration\\logo2.ico')
root.geometry('400x300')

orgName = tk.StringVar()
year = tk.StringVar()
month = tk.StringVar()
fileName = tk.StringVar()

open_file_path1 = tk.StringVar()
open_file_path2 = tk.StringVar()
save_file_path = tk.StringVar()

stat_message = tk.StringVar()
result_message = tk.StringVar()

file_names=[]

root.title('HARDI Data Submission')

w3=tk.Label(root,justify=tk.CENTER,text='Distributor\' Name:\n(Copy From the Website)').pack()
e3=tk.Entry(root)
e3.pack()

w4=tk.Label(root,justify=tk.CENTER,text='Year:\n(Like 2018)').pack()
e4=tk.Entry(root)
e4.pack()


w5=tk.Label(root,justify=tk.CENTER,text='Month:\n(Like 6 for June)').pack()
e5=tk.Entry(root)
e5.pack()


def open_file():

    global open_file_path
    global file_names

    #open_file_path=filedialog.askopenfilename(initialdir='B:\\HARDI\\Participants\\Submissions')
    file_names=filedialog.askopenfilenames(initialdir='B:\\HARDI\\Participants\\Submissions')

    #pathLabel.config(text=os.path.split(open_file_path)[1])
    if len(file_names)==1:
        open_file_path1=file_names[0]
        pathLabel1.config(text=os.path.split(open_file_path1)[1])
    elif len(file_names)==2:
        open_file_path1=file_names[0]
        pathLabel1.config(text=os.path.split(open_file_path1)[1])
        open_file_path2=file_names[1]
        pathLabel2.config(text=os.path.split(open_file_path2)[1])
    elif len(file_names)==3:
        open_file_path1=file_names[0]
        pathLabel1.config(text=os.path.split(open_file_path1)[1])
        open_file_path2=file_names[1]
        pathLabel2.config(text=os.path.split(open_file_path2)[1])
        open_file_path3=file_names[2]
        pathLabel3.config(text=os.path.split(open_file_path3)[1])
    else:
        return

errmsg = 'Error!'
wopen=tk.Button(text='Open File', command=open_file)
wopen.pack()

# pathLabel=tk.Label(root)
# pathLabel.pack()

pathLabel1=tk.Label(root)
pathLabel1.pack()

pathLabel2=tk.Label(root)
pathLabel2.pack()

pathLabel3=tk.Label(root)
pathLabel3.pack()


def convert():


    global stat_message
    global result_message
    global year
    global orgName
    global month
    global save_file_path
    global open_file_path

    orgName=e3.get()

    # open_file_path1=file_names=[0]
    # names=os.path.basename(open_file_path1)
    # file_name=names.split('.')[0]
    # file_name=file_name.split(' ')[0]
    #
    # if file_name not in orgName:
    #     mBox.showwarning('Warning', 'Distributor\'s name doesn\'t match file name')
    #     return

    save_file_path = filedialog.asksaveasfile(mode='w', defaultextension='.csv')

    year=e4.get()
    month=e5.get()

    year=int(year)
    month=int(month)

    #stat_message, result_message=processor.process(orgName,year,month,open_file_path,save_file_path)
    stat_message, result_message=processor.process(orgName,year,month,file_names,save_file_path)

    mBox._show('Result','This is the result \n: %s \n %s' %(stat_message,result_message))

wprocess=tk.Button(text='Start Processing',command=convert)
wprocess.pack()

root.mainloop()
