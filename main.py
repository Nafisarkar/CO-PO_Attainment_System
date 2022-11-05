from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import os


def openfile():
    filename = filedialog.askopenfilename(
        initialdir="/",
        title="Select CSV",
        filetypes=(("csv files", "*.csv"), ("all files", "*.*")),
    )
    print(filename)
    if filename == "":
        status["text"] = "No File Selected"
    else:
        status["text"] = "File Selected Successfully"
        file_csv_data = pd.read_csv(filename,skiprows=0)
        file_csv_data["Total"] = [i for i in range(len(file_csv_data))]
        file_csv_data.to_csv(filename,index=False)
        print(file_csv_data)
        
        for i in range(len(file_csv_data)):
            print(file_csv_data.iloc[i][1])
            print(file_csv_data.iloc[i][2])
            print(file_csv_data.iloc[i][3])
            
            





def submit():
    print("->Submit button pressed")



window = Tk()
os.system("cls")
window.title("Co Po Attenment System")
window.geometry("260x300")
window.option_add("*font", "Helvetica 8 bold")

name_lable = Label(
    window,
    text="Name",
)


submit_btn = Button(window, width=10, text="Submit", command=submit)
submit_btn.grid(column=0, row=4, padx=5, pady=5, sticky=E)

status_lable = LabelFrame(window, text="debug:")
status_lable.grid(column=0, row=5, columnspan=2, padx=5, pady=5, sticky=W + E)

status = Label(status_lable, text="idle")
status.grid(padx=5, pady=5, sticky=N + S)

file_selecter = Button(window, width=10, text="Select File", command=openfile)
file_selecter.grid(column=1, row=4, padx=5, pady=5, sticky=W)

window.mainloop()
