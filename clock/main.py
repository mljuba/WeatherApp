from tkinter import *
from tkinter.ttk import *

from time import strftime

root = Tk()
root.title("Clock")


def time():
    string = strftime('%H:%M:%S')
    label.config(text=string)
    label.after(1000, time)


label = Label(root, font=("ds-digital", 80), background="black", foreground="red")
label.pack(anchor='center')
if __name__ == '__main__':
    time()
    mainloop()
