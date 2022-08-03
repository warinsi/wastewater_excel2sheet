# Import the required libraries
from cmath import nan
from tkinter import *
from tkinter import ttk, filedialog
import pandas as pd
import gspread
from datetime import datetime
from tkinter import messagebox
import math
import numpy as np

sa = gspread.service_account(filename='./skcn_wastewater.json')
sh = sa.open("SKCN_waste_water_IoT")

wks = sh.worksheet("data")
now = datetime.now()
Date = now.strftime("%Y-%m-%d")
Month = now.strftime("%B")
Year = now.strftime("%Y")

sampling_date = ""
sampling_date_form = ""
dummy_date = [Year + "-01-05", Year + "-02-05", Year + "-03-05", Year + "-04-05", Year + "-05-05", Year + "-06-05", Year + "-07-05", Year + "-08-05", Year + "-09-05", Year + "-10-05", Year + "-11-05", Year + "-12-05" ]

# last_row = len(wks.get_all_values()) + 1
# print("last_row :",last_row)
# last_write = int(last_row)+53
# Write_row = str("A"+str(last_row)+":"+"J"+str(last_write))
# print("Write_row :",Write_row)

excel_data = []
win = Tk()

# Set the size of the tkinter window
win.geometry("1200x500")

# Create an object of Style widget
style = ttk.Style()
style.theme_use('clam')

# Create a Frame
frame = Frame(win)
frame.pack(pady=20)
# Define a function for opening the file

# send data to google spreadsheet "SKCN_waste_water_IoT"
def send_data():
    global match_form_,sampling_date_form, wks
    last_row = len(wks.get_all_values()) + 1

    if match_form_ > 0:
        messagebox.showerror("Error", "Your form incorrect please check your form first")
        
    else:
        global excel_data, Write_sheet, Date, year_, month_
        month_remains = 12 - month_

        # remove dummy date
        amount_delete_row = (month_remains+1)*54
        start_delete_row = last_row - amount_delete_row
        print("start_delete_row: ", start_delete_row)
        print("amount_delete_row: ", amount_delete_row)
        wks.delete_rows(start_delete_row, last_row)

        # define row to write
        last_row = len(wks.get_all_values()) + 1
        print("last_row :",last_row)
        last_write = int(last_row)+53
        Write_row = str("A"+str(last_row)+":"+"J"+str(last_write))
        print("Write_row :",Write_row)
        
        #change month_/year_ to name
        datetime_object = datetime.strptime(str(month_), "%m")
        month_name = datetime_object.strftime("%B")

        # add value from excel to google sheet
        wks.update(Write_row, [
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[12][0], float(450.00), float(276.00), "", excel_data[12][2]],   #BOD
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[13][0], float(600.00), float(523.00), "", excel_data[13][2]],   #COD
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[19][0], float(1.00), float(0.13), "", excel_data[19][2]],       #Copper
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[20][0], float(1.00), float(0.12), "", excel_data[20][2]],       #Lead
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[21][0], float(1.00), float(0.15), "", excel_data[21][2]],       #Nickel
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[16][0], float(100.00), float(65.00), "", excel_data[16][2]],    #Oil & Grease
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[9][0],  float(9.00), float(6.00), "", excel_data[9][2]],        #pH
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[18][0], float(10.00), float(4.00), "", excel_data[18][2]],      #Sulfide
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[15][0], float(3000.00), float(1424.00), "", excel_data[15][2]], #TDS
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[17][0], float(100.00), float(57.00), "", excel_data[17][2]],    #TKN
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", "TSS", float(500.00), float(391.00), "", excel_data[14][2]],   #TSS
            [sampling_date_form, year_, month_name , "Head Office", "Navanakorn", excel_data[22][0], float(5.00), float(2.89), "", excel_data[22][2]],       #Zinc

            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[12][0], float(450.00), float(53.00), excel_data[12][3], excel_data[12][4]],     #BOD 
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[13][0], float(600.00), float(221.00), excel_data[13][3], excel_data[13][4]],    #COD
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[19][0], float(1.00), float(0.14), excel_data[19][3], excel_data[19][4]],        #Copper
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[20][0], float(1.00), float(0.10), excel_data[20][3], excel_data[20][4]],        #Lead
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[21][0], float(1.00), float(0.21), excel_data[21][3], excel_data[21][4]],        #Nickel
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[16][0], float(100.00), float(8.00), excel_data[16][3], excel_data[16][4]],      #Oil & Grease
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[9][0],  float(9.00), float(6.00), excel_data[9][3], excel_data[9][4]],          #pH
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[18][0], float(10.00), float(2.00), excel_data[18][3], excel_data[18][4]],       #Sulfide
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[15][0], float(3000.00), float(3000.00), excel_data[15][3], excel_data[15][4]],  #TDS
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[17][0], float(100.00), float(16.00), excel_data[17][3], excel_data[17][4]],     #TKN
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", "TSS", float(500.00), float(217.00), excel_data[14][3], excel_data[14][4]],    #TSS
            [sampling_date_form, year_, month_name , "Head Office", "WWTP", excel_data[22][0], float(5.00), float(5.00), excel_data[22][3], excel_data[22][4]],        #Zinc
             
            [sampling_date_form, year_, month_name , "Head Office", "Public", excel_data[12][0], float(20.00), float(10.00), "", excel_data[12][5]],                 #BOD 
            [sampling_date_form, year_, month_name , "Head Office", "Public", excel_data[13][0], float(120.00), float(75.00), "", excel_data[13][5]],                #COD
            [sampling_date_form, year_, month_name , "Head Office", "Public", excel_data[16][0], float(5.00), float(2.50), "", excel_data[16][5]],                   #Oil & Grease
            [sampling_date_form, year_, month_name , "Head Office", "Public", excel_data[9][0],  float(9.00), float(6.00), "", excel_data[9][5]],                    #pH
            [sampling_date_form, year_, month_name , "Head Office", "Public", excel_data[18][0], float(1.00), float(1.00), "", excel_data[18][5]],                   #Sulfide
            [sampling_date_form, year_, month_name , "Head Office", "Public", excel_data[15][0], float(3000.00), float(926.00), "", excel_data[15][5]],              #TDS
            [sampling_date_form, year_, month_name , "Head Office", "Public", excel_data[17][0], float(100.00), float(11.00), "", excel_data[17][5]],                 #TKN
            [sampling_date_form, year_, month_name , "Head Office", "Public", "TSS", float(50.00), float(28.00), "", excel_data[14][5]],                  #TSS

            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", excel_data[12][0], float(450.00), float(163.00), excel_data[12][7], excel_data[12][6]],
            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", excel_data[13][0], float(600.00), float(114.00), excel_data[13][7], excel_data[13][6]],
            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", excel_data[16][0], float(100.00), float(12.00), excel_data[16][7], excel_data[16][6]],
            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", excel_data[9][0],  float(9.00), float(6.00), excel_data[9][7], excel_data[9][6]],
            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", excel_data[18][0], float(10.00), float(6.00), excel_data[18][7], excel_data[18][6]],
            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", excel_data[15][0], float(3000.00), float(1057.00), excel_data[15][7], excel_data[15][6]],
            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", excel_data[17][0], float(100.00), float(83.00), excel_data[17][7], excel_data[17][6]],
            [sampling_date_form, year_, month_name , "Parts Center", "WWTP", "TSS", float(500.00), float(293.00), excel_data[14][7], excel_data[14][6]],
             
            [sampling_date_form, year_, month_name , "Parts Center", "Public", excel_data[12][0], float(20.00), float(12.00), "", excel_data[12][8]],
            [sampling_date_form, year_, month_name , "Parts Center", "Public", excel_data[13][0], float(120.00), float(65.00), "", excel_data[13][8]],
            [sampling_date_form, year_, month_name , "Parts Center", "Public", excel_data[16][0], float(5.00), float(4.00), "", excel_data[16][8]],
            [sampling_date_form, year_, month_name , "Parts Center", "Public", excel_data[9][0],  float(9.00), float(6.00), "", excel_data[9] [8]],
            [sampling_date_form, year_, month_name , "Parts Center", "Public", excel_data[18][0], float(1.00), float(1.00), "", excel_data[18][8]],
            [sampling_date_form, year_, month_name , "Parts Center", "Public", excel_data[15][0], float(3000.00), float(1019.00), "", excel_data[15][8]],
            [sampling_date_form, year_, month_name , "Parts Center", "Public", excel_data[17][0], float(100.00), float(51.00), "", excel_data[17][8]],
            [sampling_date_form, year_, month_name , "Parts Center", "Public", "TSS", float(50.00), float(31.00), "", excel_data[14][8]],

            [sampling_date_form, year_, month_name , "K-Max3", "WWTP", excel_data[12][0], float(450.00), float(13.00), "", excel_data[12][9]],
            [sampling_date_form, year_, month_name , "K-Max3", "WWTP", excel_data[13][0], float(600.00), float(81.00), "", excel_data[13][9]],
            [sampling_date_form, year_, month_name , "K-Max3", "WWTP", excel_data[16][0], float(100.00), float(8.00), "", excel_data[16][9]],
            [sampling_date_form, year_, month_name , "K-Max3", "WWTP", excel_data[9][0],  float(9.00), float(6.00), "", excel_data[9][9]],
            [sampling_date_form, year_, month_name , "K-Max3", "WWTP", excel_data[15][0], float(3000.00), float(1375.00), "", excel_data[15][9]],
            [sampling_date_form, year_, month_name , "K-Max3", "WWTP", "TSS", float(500.00), float(83.00), "", excel_data[14][9]],
            
            ])

        # add date until end of the year
        for i in range(month_remains): 
            month_index = month_ + i
            d_date = math.ceil(Time2GSS(dummy_date[month_index]))

            last_row = len(wks.get_all_values()) + 1
            print("last_row :",last_row)
            last_write = int(last_row)+53
            Write_row = str("A"+str(last_row)+":"+"J"+str(last_write))
            print("Write_row :",Write_row)

            #change month_+i+1/year_+i+1 to name
            datetime_object = datetime.strptime(str(month_index+1), "%m")
            month_name = datetime_object.strftime("%B")

            wks.update(Write_row, [
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[12][0], "", "", "", ""],   
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[13][0], "", "", "", ""],   
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[19][0], "", "", "", ""],       
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[20][0], "", "", "", ""],     
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[21][0], "", "", "", ""],      
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[16][0], "", "", "", ""],   
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[9][0],  "", "", "", ""],     
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[18][0], "", "", "", ""],     
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[15][0], "", "", "", ""], 
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[17][0], "", "", "", ""],    
                [d_date, year_, month_name , "Head Office", "Navanakorn", "TSS", "", "", "", ""],   
                [d_date, year_, month_name , "Head Office", "Navanakorn", excel_data[22][0],  "", "", "", ""],       

                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[12][0], "", "", "", ""],    
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[13][0], "", "", "", ""],    
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[19][0], "", "", "", ""],       
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[20][0], "", "", "", ""],        
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[21][0], "", "", "", ""], 
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[16][0], "", "", "", ""],  
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[9][0],  "", "", "", ""], 
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[18][0], "", "", "", ""], 
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[15][0], "", "", "", ""], 
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[17][0], "", "", "", ""], 
                [d_date, year_, month_name , "Head Office", "WWTP", "TSS", "", "", "", ""],  
                [d_date, year_, month_name , "Head Office", "WWTP", excel_data[22][0], "", "", "", ""],  

                [d_date, year_, month_name , "Head Office", "Public", excel_data[12][0], "", "", "", ""],
                [d_date, year_, month_name , "Head Office", "Public", excel_data[13][0], "", "", "", ""],
                [d_date, year_, month_name , "Head Office", "Public", excel_data[16][0], "", "", "", ""],
                [d_date, year_, month_name , "Head Office", "Public", excel_data[9][0],  "", "", "", ""],
                [d_date, year_, month_name , "Head Office", "Public", excel_data[18][0], "", "", "", ""],
                [d_date, year_, month_name , "Head Office", "Public", excel_data[15][0], "", "", "", ""],
                [d_date, year_, month_name , "Head Office", "Public", excel_data[17][0], "", "", "", ""],
                [d_date, year_, month_name , "Head Office", "Public", "TSS", "", "", "", ""],            

                [d_date, year_, month_name , "Parts Center", "WWTP", excel_data[12][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "WWTP", excel_data[13][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "WWTP", excel_data[16][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "WWTP", excel_data[9][0],  "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "WWTP", excel_data[18][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "WWTP", excel_data[15][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "WWTP", excel_data[17][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "WWTP", "TSS", "", "", "", ""],

                [d_date, year_, month_name , "Parts Center", "Public", excel_data[12][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "Public", excel_data[13][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "Public", excel_data[16][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "Public", excel_data[9][0],  "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "Public", excel_data[18][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "Public", excel_data[15][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "Public", excel_data[17][0], "", "", "", ""],
                [d_date, year_, month_name , "Parts Center", "Public", "TSS", "", "", "", ""],

                [d_date, year_, month_name , "K-Max3", "WWTP", excel_data[12][0], "", "", "", ""],
                [d_date, year_, month_name , "K-Max3", "WWTP", excel_data[13][0], "", "", "", ""],
                [d_date, year_, month_name , "K-Max3", "WWTP", excel_data[16][0], "", "", "", ""],
                [d_date, year_, month_name , "K-Max3", "WWTP", excel_data[9][0],  "", "", "", ""],
                [d_date, year_, month_name , "K-Max3", "WWTP", excel_data[15][0], "", "", "", ""],
                [d_date, year_, month_name , "K-Max3", "WWTP", "TSS", "", "", "", ""],
            ])

        print("send data to google sheet")

def Time2GSS(dt):
    #2022-05-05
    day = str(dt[8:10])
    month = str(dt[5:7])
    year_ = str(dt[0:4])
    dt2 = year_ + "/" + month + "/" + day
    datetime_object = datetime.strptime(dt2 , '%Y/%m/%d')
    print("datetime_object: ", datetime_object)
    #datetime_object = datetime.strptime(self.datetime_str, '%d/%m/%Y') # Si quieres puede iniciar en un string
    secs = datetime.timestamp(datetime_object) # Tiempo en segundos
    Difpy_ex_h = 25568 # Días de diferencia entre el punto cero de excel y python (Excel inicia el 0/1/1900 a las 00 y Python el 31/12/1969 a las 19)
    Difpy_ex_s = (Difpy_ex_h * 24 + 19) * 60 * 60 # La diferencia en segundos incluyendo las 19 horas
    Day_ex = (secs + Difpy_ex_s)/(60 * 60 *24) # El número de dias Excel style obtenido en python
    return Day_ex

def open_file():
    global excel_data,sampling_date_form,year_,month_,day_
    month_ = 0
    year_ = 0
    filename = filedialog.askopenfilename(title="Open a File", filetype=(
        ("xlxs files", ".*xlsx"), ("All Files", "*.")))

    if filename:
        try:
            filename = r"{}".format(filename)
            df = pd.read_excel(filename)
            label.config(text="File OK")
        except ValueError:
            label.config(text="File could not be opened")
        except FileNotFoundError:
            label.config(text="File Not Found")

    # Clear all the previous data in tree
    clear_treeview()

    # Add new data in Treeview widget
    tree["column"] = list(df.columns)
    tree["show"] = "headings"

    # For Headings iterate over the columns
    for col in tree["column"]:
        tree.heading(col, text=col)

    df = df.fillna(0)

    # Put Data in Rows
    df_rows = df.to_numpy().tolist()
    excel_data = df_rows

    # ตำแหน่งแรกคือ row ตำแหน่งสองคือ
    # print("excel_data[11][2]: ", excel_data[11][2])

    for row in df_rows:
        tree.insert("", "end", values=row)

    tree.pack()

    if isinstance(excel_data[6][2], datetime):
        excel_date = math.ceil(Time2GSS(str(excel_data[6][2])))
    else:
        excel_date = excel_data[6][2]

    sampling_date = str(datetime.fromordinal(datetime(1900, 1, 1).toordinal() + excel_date - 2))
    print("sampling date: ", sampling_date)
    sampling_date_short = sampling_date[0:10]
    print("sampling_date_short: ", sampling_date_short)

    # dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + excel_data[6][2] - 2)
    # tt = dt.timetuple()
    # print("dt: ", dt)
    # print("tt: ", tt)
    
    year_ = int(sampling_date_short[0:4])
    month_ = int(sampling_date_short[5:7])
    day_ = int(sampling_date_short[8:10])

    print("year_: ",year_) 
    print("month_: ",month_) 
    print("date_: ",day_) 
    
    # date_convert = str(year_ + "/" + month_ +"/" +date_)
    # print("date convert : ",date_convert)
    #print("date : ",sampling_date_short)
    # datetime_object = datetime.strptime(date_convert , '%Y/%m/%d')
    # sampling_date_form = math.ceil(Time2GSS(datetime_object))
    
    sampling_date_form = excel_date
    print("date : ",sampling_date_form)

    check_form_EXCEL()

# Clear the Treeview Widget


def clear_treeview():
    tree.delete(*tree.get_children())


# Create a Treeview widget
tree = ttk.Treeview(frame)

def check_form_EXCEL():
     global excel_data
     global match_form_
     check_list = ["pH",
                   "Color - ADMI (Adjust)",
                   "Color - ADMI (Normal)",
                   "BOD",
                   "COD",
                   "SS",
                   "TDS ",
                   "Oil & Grease",
                   "TKN",
                   "Sulfide",
                   "Copper (Cu)",
                   "Lead (Pb)",
                   "Nickel (Ni)",
                   "Zinc (Zn)"
                   ]
     head_list = ["1. น้ำทิ้งลงสู่ระบบบำบัดน้ำเสียของ\nนวนคร",
                  "2. น้ำที่เข้าระบบบำบัดโลหะหนัก Paint 2",
                  "3. น้ำที่ปล่อยออกจากระบบบำบัดโลหะหนัก Paint 2",
                  "4. น้ำที่ปล่อยลงสู่\nลำรางสาธารณะ",
                  "1. น้ำที่ออกจากระบบบำบัดน้ำเสีย",
                  "2. น้ำที่เข้าระบบบำบัดน้ำเสีย",
                  "3. น้ำที่ปล่อยลงสู่\nลำรางสาธารณะ",
                  "1. น้ำที่ออกจากระบบบำบัดน้ำเสีย",
                  ]
     check_list_cell = [9,10,11,12,13,14,15,16,17,18,19,20,21,22]
     head_list_cell = [2,3,4,5,6,7,8,9]
     print("data_test",excel_data[5][head_list_cell[0]])
     i = 0
     match_form_ = 0
     for x in check_list :
         
         if x == excel_data[check_list_cell[i]][0]:
             i = i+1
             print("data match",excel_data[check_list_cell[i-1]][0])
             match_form_ += 0
         else :
             i = i+1
             print("not match form",excel_data[check_list_cell[i-1]][0])
             match_form_ += 2
     i = 0
     for x in head_list:

          if x == excel_data[5][head_list_cell[i]]:
              i = i+1
              print("data match",excel_data[5][head_list_cell[i-1]])
              match_form_ += 0
          else :
              i = i+1
              print("not match form",excel_data[5][head_list_cell[i-1]])
              match_form_ += 1
             

btn = Button(win, text='open excel file', command=open_file).pack()
btn2 = Button(win, text='Send data',
              command=send_data).place(x=570, y=400)
# Add a Label widget to display the file content
label = Label(win, text='')
label.pack(pady=20)

win.mainloop()
