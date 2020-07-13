import ModifyEffCell
from ModifyEffCell import main
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import *

def select_path():
    global path
    curr_directory = os.getcwd()
    filename = filedialog.askdirectory(initialdir=curr_directory, title="Select Folder")
    path.set(filename)

def dummy():
    main(path.get())

root = Tk()
root.title('Efficiency')
root.geometry('250x200')

path = StringVar()

label = tk.Label(root, text="File Path:")
label.place(x=0, y=5)

entry =  tk.Entry(root, width=20, text=path)
entry.place(x=52, y=7)

button1 = tk.Button(root, text="select",  command=select_path)
button1.place(x=180, y=0)

button2 = tk.Button(root, text="GO",  command=dummy)
button2.place(x=100, y=30)


root.mainloop()

