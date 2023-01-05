"""
This app allows user to mailmerge infos from the user interface into a word docx.

Composed of two classes (one Data // one Tkinter) and functions to pre-load the data into fields, based on user's choice

The Tkinter class is about the design, interface, entry fields and buttons. Few variables are declared inside it.
The Data class is about getting the data from the fields into the mailmerge function.


coded by kiki @VCH

"""
from mailmerge import *

from tkinter import *
from tkinter import messagebox
import tkinter as tk

#import win32com.client as win32

from pathlib import Path
import shutil
import datetime
import os

# class 
class data():

    # button output
    def open():
        os.system("explorer output")

    def get_data():
        global name
        global first_name
        global age
        global location
        global mail

        name=name.get()
        first_name = first_name.get()
        age = age.get()
        location = location.get()
        mail = mail.get()

    def send_data():
        print(name, first_name, age, location, mail)
        return (name, first_name, age, location, mail)

    def publish():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            if len(os.listdir(path)) == 0:
                data.get_data()
                print("data copied into the word file")


                template_1 = "temp/file.docx"
                document_1 = MailMerge(template_1)
                document_1.merge(name_p = name, first_n_p = first_name, age_p = age, location_p = location)
                document_1.write('file_output.docx')

            else:
                print("Output not empty. Erase file in it and try again.")
                messagebox.showinfo("Warning", "Output folder not empty")
        else:
            print("Output folder not present anymore")
            messagebox.showinfo("Warning", "Folder not present anymore")
  
                



# class for the front office aka the visual app 
class app(tk.Frame):
    def __init__(self):
        self.tk.Frame.__init()
   
if __name__ == "__main__":
    window=tk.Tk()
    window.title("Word Deployer")
    window.geometry("500x280")
    window.resizable(width=True, height=True)
    # 1 these global variables are how we can have a value in the field, and then still use the user input to be what's going to be mailmerged



    #Labels
    l1=Label(window, text= "Name")
    l1.grid(row=0,column=0, pady=2)
    l2=Label(window, text= "First Name")
    l2.grid(row=2,column=0, pady=2)
    l3=Label(window, text= "Age")
    l3.grid(row=4,column=0, pady=2)
    l4=Label(window, text= "E-Mail")
    l4.grid(row=6,column=0, pady=2)
    l5=Label(window, text= "Location")
    l5.grid(row=8,column=0, pady=4)

    #Entry fields
    name=StringVar()
    e1=Entry(window, width=40, bd=1, textvariable=name)
    e1.grid(row=0, column=1, pady=2)
    first_name=StringVar()
    e2=Entry(window, width=40, bd=1, textvariable=first_name)
    e2.grid(row=2, column=1, pady=2)
    age=StringVar()
    e3=Entry(window, width=40, bd=1, textvariable=age)
    e3.grid(row=4, column=1, pady=2)
    mail=StringVar()
    e4=Entry(window, width=40, bd=1, textvariable=mail)
    e4.grid(row=6, column=1, pady=2)
    location=StringVar()
    e5=Entry(window, width=40, bd=1, textvariable=location)
    e5.grid(row=8, column=1, pady=2)
    


    
    #Buttons
    b1=Button(window, text="Publish", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command = data.get_data)
    b1.grid(row=11, column=1, pady=2)
    b2=Button(window,text="Output", width=16, borderwidth=1, relief="raised", fg="blue", activebackground="green", overrelief="sunken", command=data.open)
    b2.grid(row=12, column=1, pady=2)
    b3=Button(window,text="Fermer", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=window.quit)
    b3.grid(row=13, column=1, pady=2)
    
    window.mainloop()  