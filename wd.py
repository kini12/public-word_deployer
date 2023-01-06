"""
This app allows user to mailmerge infos from the user interface into a word docx.

Composed of two classes (one Data // one Tkinter) and functions to pre-load the data into fields, based on user's choice

The Tkinter class is about the design, interface, entry fields and buttons. Few variables are declared inside it.
The Data class is about getting the data from the fields into the mailmerge function.

coded by kiki @VCH

"""
from mailmerge import MailMerge
from tkinter import *
from tkinter import messagebox
import tkinter as tk
import os

# class 
class data():

    # open the output folder from within the app
    def open():
        os.system("explorer output")

    # assign var to the fields of the app
    def get_data():
        global name_form
        global first_name_form
        global age_form
        global location_form
        global mail_form
        name_form=name_p.get()
        first_name_form = first_name_p.get()
        age_form = age_p.get()
        location_form = location_p.get()
        #print(name_form, first_name_form, age_form, location_form, mail_form)
        return (name_form, first_name_form, age_form, location_form)

    # function for the publish button
    # use the output path to store the file with the data inserted from the app
    def publish():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            data.get_data()
            print("data is about to be copied into the word file")
            template_1 = "temp/file_1.docx"
            document_1 = MailMerge(template_1)
            document_1.merge(name_w = name_form, first_n_w = first_name_form, age_w = age_form, location_w = location_form)
            document_1.write('output/file_output.docx')
            os.rename('output/file_output.docx', "output/%s_%s.docx" %(first_name_form, name_form))
            print("data copied, it'all good") 
        else:
            print("Output not empty. Erase file in it and try again.")
            messagebox.showinfo("Warning", "Output folder not empty")
            
# class for the front office aka the visual app 
class app(tk.Frame):
    def __init__(self):
        self.tk.Frame.__init()
   
if __name__ == "__main__":
    window=tk.Tk()
    window.title("Word Deployer")
    window.geometry("350x200")
    window.resizable(width=True, height=True)

    # theses global variables are the variables used to capture what is in the fields
    global name_p
    name = tk.StringVar(window)
    name.set(name)
    global first_name_p
    first_name = tk.StringVar(window)
    first_name.set(first_name)
    global age_p
    age = tk.StringVar(window)
    age.set(age)
    global location_p
    location = tk.StringVar(window)
    location.set(location)

    # design of the UI 
    #Labels
    l1=Label(window, text= "Name")
    l1.grid(row=0,column=0, pady=2)
    l2=Label(window, text= "First Name")
    l2.grid(row=2,column=0, pady=2)
    l3=Label(window, text= "Age")
    l3.grid(row=4,column=0, pady=2)
    l4=Label(window, text= "Location")
    l4.grid(row=6,column=0, pady=4)

    #Entry fields
    name_p=StringVar()
    e1=Entry(window, width=40, bd=1, textvariable=name_p)
    e1.grid(row=0, column=1, pady=2)
    first_name_p=StringVar()
    e2=Entry(window, width=40, bd=1, textvariable=first_name_p)
    e2.grid(row=2, column=1, pady=2)
    age_p=StringVar()
    e3=Entry(window, width=40, bd=1, textvariable=age_p)
    e3.grid(row=4, column=1, pady=2)
    location_p=StringVar()
    e4=Entry(window, width=40, bd=1, textvariable=location_p)
    e4.grid(row=6, column=1, pady=2)
    
    #Buttons
    b1=Button(window, text="Publish", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command = data.publish)
    b1.grid(row=11, column=1, pady=2)
    b2=Button(window,text="Output", width=16, borderwidth=1, relief="raised", fg="blue", activebackground="green", overrelief="sunken", command=data.open)
    b2.grid(row=12, column=1, pady=2)
    b3=Button(window,text="Fermer", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=window.quit)
    b3.grid(row=13, column=1, pady=2)
    
    window.mainloop()  