import GUI_Functions
import Define_variable
import tkinter as tk
from functools import partial
from tkmacosx import Button

Define_variable.window_label.pack(pady=20);


normal_run_btn = Button(Define_variable.window, text='Run', bg='black', fg='green', borderless=1,
                        command=partial(GUI_Functions.normal_run))
normal_run_btn.pack(ipadx=100, ipady=30, pady=10)

Define_variable.result_label.pack(ipadx=100,pady=10)
Define_variable.Separate1.pack(ipadx=100)

Define_variable.start_index_frame.pack()
Define_variable.start_index_label.pack()
Define_variable.start_index_entry.insert(0, '0')

Define_variable.start_index_entry.pack()

start_index_run_btn = Button(Define_variable.window,
                             text='Run from index', bg='black', fg='green', borderless=1,
                             command=partial(GUI_Functions.normal_run))
start_index_run_btn.pack()


Define_variable.Separate2.pack(ipadx=100)

Define_variable.checkbutton3.pack()
Define_variable.checkbutton2.pack()
Define_variable.checkbutton1.pack()

Define_variable.window.mainloop()
