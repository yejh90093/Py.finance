import tkinter as tk

#  tkinter part   #
window = tk.Tk()
window.title('TW Finance')
window.geometry('500x400')
window.configure(background='white')

result_label = tk.Label(window)

window_label = tk.Label(window, text='TW Finance', fg='blue')

Separate1 = tk.Label(window, text="===========================================================", bg="gray", fg="white")
Separate2 = tk.Label(window, text="===========================================================", bg="gray", fg="white")
start_index_frame = tk.Frame(window)
start_index_label = tk.Label(start_index_frame, text='Run from index')
start_index_entry = tk.Entry(window)

var_debugMode = tk.IntVar()
var_saveLocal = tk.IntVar()
var_overwrite = tk.IntVar()
checkbutton1 = tk.Checkbutton(window, text='Debug Mode', variable=var_debugMode, onvalue=1, offvalue=0)
checkbutton2 = tk.Checkbutton(window, text='Save Local', variable=var_saveLocal, onvalue=1, offvalue=0)
checkbutton3 = tk.Checkbutton(window, text='Overwrite File', variable=var_overwrite, onvalue=1, offvalue=0)


#  tkinter part   #
