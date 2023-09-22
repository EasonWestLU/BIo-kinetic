import Kinetic
import os
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
from PIL import ImageTk, Image

parent = tk.Tk()
image1 = Image.open("pig.gif")
image2 = Image.open("pig2.gif")
image_tk = ImageTk.PhotoImage(image1)
#image_tk = ImageTk.PhotoImage(image2)



def TAKE():
    if str(combo.get()) == 'DAEM':
        Kinetic.DAEM()
    if str(combo.get()) == 'Friedman':
        os.system('python Friedman.py')
    if str(combo.get()) == 'Redfern':
        os.system('python CR.py')



parent.title("选择你的英雄")
combo = ttk.Combobox(parent)
combo.place(x=50, y=100)
combo['values'] = ('DAEM', 'Friedman', 'Redfern')
combo.current(2)
button = tk.Button(parent, command=TAKE, text='确定')
button.place(x=80, y=150)
parent.geometry('300x300')
parent.mainloop()