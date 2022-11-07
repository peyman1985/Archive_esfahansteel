#1401-08-15
#peyman-ramezani
from datetime import date
import jdatetime
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter.filedialog import *
import pyodbc
import os
from pathlib import Path
import shutil
import sys

TODAY = str(jdatetime.date.today())
cursor = pyodbc.Cursor
# path for Access file address
path_a = ""
path_m = Path()
path_f = ""


def get_cods():
    path_maps = str(askdirectory())
    file_name = os.listdir(path_maps)
    for each in file_name:
      txt_in.insert('1.0', each.split('.')[0]+'\n')

# *******************************************************************************************************************************************
# find Access file path
def browser():
    path_a = str(askopenfilename())  # ask path
    lb_Access["text"] = path_a  # replease path to lebal Access file

# *******************************************************************************************************************************************
def readTextIn():
    dwgno1 = txt_in.get("1.0", END)  # get numbers from txt_in as list
    dwgno = []  # list without \n and whitespace
    tem = dwgno1.split("\n")
    for i in tem:
        if i != "":
            dwgno.append(i)
    return set(dwgno)
# *******************************************************************************************************************************************
def today_slash():
    part1 = TODAY
    date_today_slash = ""
    for i in part1:
        if i == "-":
            i = "/"
            date_today_slash += i
        else:
            date_today_slash += i
    return date_today_slash
# *******************************************************************************************************************************************
def rest():
    txt_out.delete("1.0", END)
    txt_in.delete("1.0",END)  
    txt_folder.delete("1.0",END) 
# *******************************************************************************************************************************************
# function for copy entered maps on folder
def ad_file():
    path_m = str(askdirectory())
    lb_file["text"] = path_m

# *******************************************************************************************************************************************
def connet_Access():
    conn = pyodbc.connect(
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        + f"DBQ={lb_Access['text']};"
    )
    cursor = conn.cursor()
    return cursor

# *******************************************************************************************************************************************
# for quit of App
def quit():
    win.destroy()
# *******************************************************************************************************************************************
def copy():
    de = str(askdirectory())  # direcctory for copy maps
    in_maps = readTextIn()  # get input values for text_in
    list_maps = []  # list to hold all maps address like str
    path_maps = Path(lb_file["text"]).glob("*")
    for each in path_maps:  # get all map path as list
        list_maps.append(each)
    for each_in in in_maps:
        for each_list in list_maps:
            if each_in == each_list.stem:  # basename without .tif or .pdf
                des = Path(de) / os.path.basename(str(each_list))
                shutil.copy(str(each_list), des)


# *******************************************************************************************************************************************
# function for move entered maps on folder
def delete_maps():
    txt_out.delete('1.0',END)
    if messagebox.askokcancel("هشدار", "از حذف نقشه های وارد شده مطمین هستید؟"):
        in_maps = readTextIn()  # get input values for text_in
        list_maps = []  # list to hold all maps address like str
        path_maps = Path(lb_file["text"]).glob("*")
        for each in path_maps:
            list_maps.append(each)
        for each_in in in_maps:
            for each_list in list_maps:
                if each_in == each_list.stem:
                    txt_out.insert(END, f"شماره نقشه {each_in} پاک شد"+'\n')
                    os.remove(str(each_list))
# *******************************************************************************************************************************************
# function to cheak maps in archive or export
def check_maps():
    txt_out.delete("1.0", END)
    folder = lb_Access["text"]
    if folder == "address File Access":
        messagebox.askokcancel("FIEL ERROR", "لطفا آدرس فایل اکسس را وارد کنید")
    else:
        d_now = txt_date.get("1.0", END)  # get date from user
        cursor = connet_Access()  # connect to Access file
        dwgno = readTextIn() # list without \n and whitespace 
        l_table1 = list(cursor.execute("select * from table1"))
        l_table4 = list(cursor.execute("select * from table4"))
        for dwg in dwgno:
            for row_t1 in l_table1:
                if dwg ==  row_t1[3]:
                    for row_t4 in l_table4:
                        if  row_t1[0] == row_t4[0]:
                            txt_out.insert(
                                END,
                                    '{0:<7}'.format(str(row_t1[4])[0:6])
                                +  '{0:<12}'.format(str(row_t1[3]))
                                +  '{0:<11}'.format(str(row_t4[1]))
                                +  '{0:<11}'.format(str(row_t4[2]))
                                +  '{0:<10}'.format(str(row_t4[4]))
                                + "\n",
                            )
                            cursor.commit()
        txt_out.insert(END, "checked")
# *******************************************************************************************************************************************
def checkCod():
    txt_out.delete("1.0", END)
    folder = lb_Access["text"]
    if folder == "address File Access":
        messagebox.askokcancel("FIEL ERROR", "لطفا آدرس فایل اکسس را وارد کنید")
    else:
        cursor = connet_Access()  # connect to Access file
        dwgno = readTextIn() # list without \n and whitespace
        l_table1 = list(cursor.execute("select * from table1"))
        l_table4 = list(cursor.execute("select * from table4"))
        for dwg in dwgno:
            for row_t1 in l_table1:
                if dwg ==  row_t1[3]:
                            txt_out.insert(END,
                            '{0:<12}'.format(str( row_t1[3]))+
                            '{0:<8}'.format(str( row_t1[4])[0:6])+
                            '{0:<5}'.format(str( row_t1[2]))+
                            '{0:<5}'.format(str( row_t1[5]))+
                            "\n")
                            cursor.commit()
        txt_out.insert(END, "checked")
# *******************************************************************************************************************************************
# function for maps that back to Archive
def importMaps():
    txt_out.delete('1.0',END)
    # get date from user
    maps =[]
    setcodes={}
    cursor = connet_Access()  # connect to Access file
    maps = readTextIn()
    table1 = list(cursor.execute("select * from table1"))
    table4 = list(cursor.execute("select * from table4"))
    for each in table1:
        setcodes[each[3]]=each[0]#each[0]= values and each[3]=keys
    for each in maps:
        if each in setcodes.keys():
            ID= setcodes[each]
            for row in table4:
                if row[0] ==ID:
                    if row[1]!=None and row[2]==None:
                        cursor.execute(f"update table4 set  TAREKH_VOROD= '{today_slash()}' , etmamtahvil=True where ID={row[0]} AND TAREKH_KHO='{row[1]}' AND KARGAH='{row[4]}' ")
                        cursor.commit()
                        txt_out.insert(END,F"ورود نقشه {each} در سیستم ثبت گردید."+'\n')
    txt_out.insert(END, "اتمام")
    
# *******************************************************************************************************************************************
# function to export maps
def exportMaps():
    F_NAME= txt_folder.get("1.0", END)
    folder = lb_Access["text"]
    if F_NAME =='\n':
        messagebox.askokcancel("FIEL ERROR", "لطفا نام  کاربر را وارد کنید")

    elif folder == "address File Access":
        messagebox.askokcancel("FIEL ERROR", "لطفا آدرس فایل اکسس را وارد کنید")
        
    else:
        setcodes={}
        cursor = connet_Access()  # connect to Access file
        maps = readTextIn()
        table1 = list(cursor.execute("select * from table1"))
        for each in table1:
            setcodes[each[3]]=each[0]#each[0]= values and each[3]=keys
        for each in maps:
            if each in setcodes.keys():
                ID= setcodes[each]
                cursor.execute(f"insert into table4 (ID,TAREKH_KHO,KARGAH) values('{ID}','{today_slash()}','{F_NAME.strip()}')")
                cursor.commit()
            else:
                txt_out.insert('1.0','this number isnt in access file')
# *******************************************************************************************************************************************
def getDetails():
    txt_out.delete('1.0',END)
    details =[]
    code_Archive=[]
    maps = readTextIn()
    cursor =connet_Access()
    table1 = list(cursor.execute("select * from table1"))
    for map in maps:
        for row_t1 in table1:
            if map == row_t1[3]:
                if row_t1[4] !=None:
                    code_Archive.append(row_t1[4][0:6])
    if code_Archive !=[]:
        for each_code in code_Archive:
            for row_t1 in table1:
                if row_t1[4] !=None:
                    if each_code == row_t1[4][0:6]:
                        details.append(row_t1[3])
                        txt_out.insert(END, row_t1[3]+'\n')
# *******************************************************************************************************************************************
def getDetails_add():
    txt_out.delete('1.0',END)
    details =[]
    code_Archive=[]
    maps = readTextIn()
    cursor =connet_Access()
    table1 = list(cursor.execute("select * from table1"))
    for map in maps:
        for row_t1 in table1:
            if map == row_t1[3]:
                if row_t1[4] !=None:
                    code_Archive.append(row_t1[4][0:6])
    if code_Archive !=[]:
        for each_code in code_Archive:
            for row_t1 in table1:
                if row_t1[4] !=None:
                    if each_code == row_t1[4][0:6]:
                        maps.add(row_t1[3])
    for map in maps:
        txt_in.insert(END,'\n'+map)



#configure the window thinter
# passWord()
# if password:
win = Tk()

win.geometry("1150x500")
win.configure(bg=("#ddd"))
win.title("آرشیو مهندسی نورد")
win.resizable = ("False", "False")
#label title
lb_welcome = Label(win, text="آرشیو مهندسی نورد", font=("aril 20"))
lb_welcome.grid(row=0, columnspan=6, sticky=EW, pady=5)
#
lb_folder = ttk.Label(win, text="درخواست کننده:", font=("aril 11"))
lb_folder.grid(row=1, column=0, sticky=EW, padx=5)
#
txt_folder = Text(win, height=1, width=18)
txt_folder.grid(row=1, column=1, columnspan=2, padx=2, sticky=EW)
#
lb_date = ttk.Label(win, text="تاریخ:", font=("aril 11"))
lb_date.grid(row=1, column=4, sticky=EW)
#
txt_date = Text(win, height=1, width=40)
txt_date.grid(row=1, column=5, padx=2, sticky=EW)
##btns
btn_Access = ttk.Button(win, text="جستجو فایل اکسس", command=browser)
btn_Access.grid(row=2, column=0, sticky=EW, padx=5)
lb_Access = ttk.Label(win, text="آدرس فایل اکسس", relief="sunken", borderwidth=0.01)
lb_Access.grid(row=2, column=1, columnspan=3, sticky=EW, ipadx=2, ipady=2)
# button for addres of maps
btn_file = ttk.Button(win, text="فایل نقشه ها", command=ad_file)
btn_file.grid(row=2, column=4, sticky=EW, padx=5)
#
lb_file = ttk.Label(
    win, text="محل ذخیره اسکن نقشه ها", relief="sunken", borderwidth=0.01
)
lb_file.grid(row=2, column=5, columnspan=2, sticky=EW, ipadx=2, ipady=2)
#
btn_createF = ttk.Button(win, text="گرفتن شماره نقشه ", command=get_cods)
btn_createF.grid(row=3, column=0, sticky=(W,N,E,S), padx=5)
#
btn_copy = ttk.Button(win, text="کپی", command=copy)
btn_copy.grid(row=4, column=0, sticky=(W,N,E,S), padx=5)
#
btn_cute = ttk.Button(win, text="حذف نقشه" ,command=delete_maps)
btn_cute.grid(row=5, column=0, sticky=(W,N,E,S), padx=5)
#
btn_check = ttk.Button(win, text="چک کردن",command=check_maps)
btn_check.grid(row=6, column=0, sticky=EW, padx=5)
#
btn_check = ttk.Button(win, text="شماره بایگانی ",command=checkCod)
btn_check.grid(row=7, column=0, sticky=(W,N,E,S), padx=5)
#
btn_import = ttk.Button(win, text="تحویل دادن ", command=exportMaps)
btn_import.grid(row=8, column=0, sticky=(W,N,E,S), padx=5)
#
btn_export = ttk.Button(win, text="تحویل گرفتن", command=importMaps)
btn_export.grid(row=9, column=0, sticky=(W,N,E,S), padx=5)
#
#
btn_check1 = ttk.Button(win, text=" ریست کردن", command=rest)
btn_check1.grid(row=10, column=0, sticky=(W,N,E,S), padx=5)
#
btn_quit = ttk.Button(win, text=" دیتایل ها ", command=getDetails)
btn_quit.grid(row=11, column=0, sticky=(W,N,E,S), padx=5)
btn_quit = ttk.Button(win, text=" اضاف کردن دیتایل به لیست", command=getDetails_add)
btn_quit.grid(row=13, column=0, sticky=(W,N,E,S), padx=5)
btn_quit = ttk.Button(win, text="خروج از برنامه", command=win.destroy)
btn_quit.grid(row=14, column=0, sticky=(W,N,E,S), padx=5)

#
txt_in = Text(win, width=12, height=18, padx=5)
txt_in.grid(row=3, column=1, rowspan=11, sticky=EW)
#
txt_out = Text(win, width=80, height=18, padx=5)
txt_out.grid(row=3, column=2, rowspan=11, columnspan=4, sticky=EW)
#
lb_end = ttk.Label(
    win,
    text="تهیه کننده : پیمان رمضانی هفشجانی",
    relief="sunken",
    borderwidth=0.01,
    anchor=W,
)
lb_end.grid(row=16, columnspan=6, sticky=EW)



txt_date.insert("1.0",TODAY)

win.mainloop()
