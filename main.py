from tkinter import *
from openpyxl import *
from tkinter import filedialog
from openpyxl import load_workbook
import pandas as pd
import os

box= []

def starting_row():
    return 22

def remove_junk_value(all_val):
    fress_box = []
    for val in all_val:
        if str(val)[0].isdigit():
            #debugging
            #print(val)
            fress_box.append(val)           
    return fress_box

def get_total_student_count(df2):
    for i in range(len(df2)):
            if str(df2[i]).strip().isdigit():
                ts_count = int(df2[i])
    return ts_count

#takes list as input and returns last row number
def get_last_row (df2):
    ts_count = get_total_student_count(df2)
    return 22+int(ts_count)

def get_all_student_id(df2):
    counter = 0 
    student_id = remove_junk_value(df2)
    for student_id in student_id:
        counter+=1
        print(str(counter) + " " +str(student_id))

def print_name_id_serial_ (df):
    df2 = list(df["SL"])
    
    start = starting_row()
    ending = get_total_student_count(df2)
    print("starting row: ", start)
    print("ending row: ", ending)
    for i in range(0, ending):
        temp_lis = list(df[['SL' , 'Name', 'ID']].iloc[i])
        print(temp_lis)
        box.append(temp_lis)

#gets the exect value from the cell
def get_num(filename, column="O", row=21):
    """Read a single cell value from an Excel file"""
    return pd.read_excel(filename, skiprows=row - 1, usecols=column, nrows=1, header=None, names=["Value"]).iloc[0]["Value"]

#getting the sum of a student 
def student_sum(filename, row=23, total=0):
    co_index = ['E','F','G','H','J','K','L','M']
    po_index = ['E','F','G','H','J','K','L','M']
    co1_fillmark = get_num(filename, 'U', 21)
    co2_fillmark = get_num(filename, 'V', 21)
    co3_fillmark = get_num(filename, 'W', 21)
    co4_fillmark = get_num(filename, 'X', 21)
    co5_fillmark = get_num(filename, 'Y', 21)
    
    #co_index = ['E','F','G','H','J','K','L']
    sum = 0
    df = pd.read_excel(filename, skiprows=21)
    wb = load_workbook(filename)
    ws = wb.worksheets[0]
    status["text"] = "Processing : " + "Data"  
    for j in range(total):
        
        co1 = get_num(filename,co_index[0], row)
        ws.cell(row, column=16).value = co1
        ws.cell(row, column=21).value = co1
        co1_percent = co1/co1_fillmark
        print("co1: ", co1_percent)
        ws.cell(row, column=26).value = co1_percent
        ws.cell(row, column=31).value = co1_percent
        if(co1_percent >= 0.5):
            ws.cell(row, column=36).value = "Yes"
            ws.cell(row, column=41).value = "Yes"
        else:
            ws.cell(row, column=36).value = "No"
            ws.cell(row, column=41).value = "No"
        
        
        co2 = get_num(filename,co_index[2], row)
        ws.cell(row, column=17).value = co2
        ws.cell(row, column=22).value = co2
        co2_percent = co2/co2_fillmark
        print("co2: ", co2_percent)
        ws.cell(row, column=27).value = co2_percent
        ws.cell(row, column=32).value = co2_percent
        if(co2_percent >= 0.5):
            ws.cell(row, column=37).value = "Yes"
            ws.cell(row, column=42).value = "Yes"
        else:
            ws.cell(row, column=37).value = "No"
            ws.cell(row, column=42).value = "No"
        
        
        co3 = get_num(filename,co_index[1], row)
        co3 += get_num(filename,co_index[3], row)
        co3 += get_num(filename,co_index[4], row)
        co3 += get_num(filename,co_index[5], row)
        ws.cell(row, column=18).value = co3
        ws.cell(row, column=23).value = co3
        co3_percent = co3/co3_fillmark
        print("co3: ", co3_percent)
        ws.cell(row, column=28).value = co3_percent
        ws.cell(row, column=33).value = co3_percent
        if(co3_percent >= 0.5):
            ws.cell(row, column=38).value = "Yes"
            ws.cell(row, column=43).value = "Yes"
        else:
            ws.cell(row, column=38).value = "No"
            ws.cell(row, column=43).value = "No"
        
        co4 = get_num(filename,co_index[6], row)
        ws.cell(row, column=19).value = co4
        ws.cell(row, column=24).value = co4
        co4_percent = co4/co4_fillmark
        print("co4: ", co4_percent)
        ws.cell(row, column=29).value = co4_percent
        ws.cell(row, column=34).value = co4_percent
        if(co4_percent >= 0.5):
            ws.cell(row, column=39).value = "Yes"
            ws.cell(row, column=44).value = "Yes"
        else:
            ws.cell(row, column=39).value = "No"
            ws.cell(row, column=44).value = "No"
        
        
        co5 = get_num(filename,co_index[7], row)
        ws.cell(row, column=20).value = co5
        ws.cell(row, column=25).value = co5
        co5_percent = co5/co5_fillmark
        print("co5: ", co5_percent)
        ws.cell(row, column=30).value = co5_percent
        ws.cell(row, column=35).value = co5_percent
        if(co5_percent >= 0.5):
            ws.cell(row, column=40).value = "Yes"
            ws.cell(row, column=45).value = "Yes"
        else:
            ws.cell(row, column=40).value = "No"
            ws.cell(row, column=45).value = "No"
        
        
        
        #for i in range(8):
            #print(get_num(filename, co_index[i], row))
            #sum += get_num(filename, co_index[i], row)
        sum = co1 + co2 + co3 + co4 + co5    
        ws.cell(row, column=15).value = sum
            
        print(f"{box[j]} got numbers in total  {j}: ", sum)
        row= row+1
        sum = 0
        status["text"] = "Processing : " + str(j+1) + " of " + str(total) + " Data"
        status.update()
    wb.save(filename)
    status["text"] = "Processing : " + "Done"


#handles file selection and file reading and branching
def openfile():
    ts_count = 0
    filename = filedialog.askopenfilename(
        initialdir="/",
        title="Select CSV",
        filetypes=(("csv files", "*.csv"),("xlsx files", "*.xlsx")),
    )
    print(filename)
    if filename == "":
        status["text"] = "No File Selected"
    else:
        status["text"] = "File Selected Successfully"
        df = pd.read_excel(filename, skiprows=21)
        print_name_id_serial_(df)
        df2 = list(df["SL"])
        #get_last_row(df2)
        total = get_total_student_count(df2)
        
        
        
        print("total student: ", get_total_student_count(df2))
        
        student_sum(filename, 23, total)
        #status["text"] = "Total Student: " + str(ts_count)
        #print("last_row number: ", get_last_row(df2))
        
        #student_id =  list(df["ID"])
        #get_all_student_id(student_id)
        
            
#debugging
def submit():
    print("->Submit button pressed")



#main window
window = Tk()

#clears the terminal
os.system("cls")
#window title
window.title("Co Po Attenment System")
#initial window size
window.geometry("180x200")
#globally sets the font
window.option_add("*font", "Helvetica 8 bold")


submit_btn = Button(window, width=10, text="Submit", command=submit)
submit_btn.grid(column=0, row=4, padx=5, pady=5, sticky=W)

file_selecter = Button(window, width=10, text="Select File", command=openfile)
file_selecter.grid(column=1, row=4, padx=5, pady=5, sticky=E)

status_lable = LabelFrame(window, text="debug:")
status_lable.grid(column=0, row=5, columnspan=2, padx=5, pady=5, sticky=W + E)

status = Label(status_lable, text="idle")
status.grid(padx=5, pady=5, sticky=N + S)



window.mainloop()
