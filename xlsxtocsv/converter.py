import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from time import sleep


def statidle():
    status["text"] = "Idle"


def openfile():
    filename = filedialog.askopenfilename(
        initialdir="/",
        title="Select Excel",
        filetypes=(("xlsx files", "*.xlsx"),("csv files", "*.csv")),
        # filetypes=(("csv files", "*.csv"), ("xlsx files", "*.xlsx")),
    )
    print(filename)
    if filename == "":
        status["text"] = "No File Selected"
    else:
        status["text"] = "File Selected Successfully"
        if filename.endswith(".csv"):
            file_csv_data = pd.read_csv(filename)
            file_csv_data.to_excel("output_xl.xlsx", index=False, header=True)
            status["text"] = "CSV to Excel Conversion Successful"
            print(file_csv_data)

        else:
            file_xlsx_data = pd.read_excel(filename)
            file_xlsx_data.to_csv("output_csv.csv", index=False, header=True)
            status["text"] = "Excel to CSV Conversion Successful"
            print(file_xlsx_data)
        statidle()


root = Tk()
root.title("Xlsx to Csv Converter")
root.geometry("300x100")

status_lable = LabelFrame(root, text="debug:")
status_lable.grid(column=0, row=0, padx=5, pady=5, sticky=W + E)

status = Label(status_lable, text="Not Selected", width=39)
status.grid(column=0, row=0, padx=5, pady=5)

select_btn = Button(root, text="Convert", command=openfile)
select_btn.grid(column=0, row=2, padx=5, pady=5, sticky=W + E)


root.mainloop()
