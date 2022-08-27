# Import Widgets
from tkinter import * 
import ttkbootstrap as ttk
import pandas as pd
import sqlite3
from time import strftime
import os
import threading

# Import Custom Widgets
from kanwar.Form import MyFrame
from kanwar.ToolTip import CreateToolTip
from kanwar.MessageBox import MessageBox

# Import Security
from kanwar.authenticate import *
from kanwar.customDateEntry import CustomDateEntry

####################### CONSTANTS ###############################

FILE_EXPORT_RESULT = 'Results.xlsx'
FILE_EXPORT_CLASS = 'Classes.xlsx'
FILE_EXPORT_STUDENT = 'Students.xlsx'
FILE_EXPORT_USER = 'Users.xlsx'
FILE_EXPORT_TEST = 'Tests.xlsx'

Authenticate = False

TITLE = 'Kanwar Adnan'

DATABASE = "Database.db"
DATABASEATTENDNACE = "Database2.db"

DEFAULT_ROLE = 'Admin'
DEFAULT_ADMIN = 'admin'
DEFAULT_PASSWORD = '123'
DEFAULT_EMAIL = 'admin'
DEFAULT_CLASS = 'IT-G2'
DEFAULT_LECTURE = 'DLD'

DEFAULT_ROLE_USER = 'Teacher'

LOGGER = 'Admin'

########################################################################

themes = ['united', 'cosmo','yeti', 'clam', 'solar', 'pulse', 'vista',
         'winnative', 'xpnative', 'default', 'litera',
          'darkly', 'lumen', 'alt', 'journal',
          'classic', 'superhero', 'minty', 'flatly',
           'cyborg', 'sandstone','vapor','morph']

style = ttk.Style(theme=themes[0])
font_lbl=  ['Helvetica',12]

root = style.master
root.focus()
root.title("Kanwar Adnan")
padx,pady = 10,100
root.geometry("940x600+100+80")
root.minsize(940,600)
style.configure('New.TNotebook',tabposition='nw')
style.configure('TNotebook.Tab',background='white')

#####################
#VARS
dcourse = StringVar()
dcourse_code = StringVar()
dlecture = StringVar()
dname_test = StringVar()
dtotal_marks_test = StringVar()
dpassing_marks_test = StringVar()
dname_test_desc = StringVar()
ddate_test = StringVar()
#####################
classes = [] 
lectures= []
tests = []
rollnumbers = []

mainframe = ttk.Frame(root)
mainframe.pack(fill=BOTH, expand=1)

nb = ttk.Notebook(mainframe)
nb.pack(fill=BOTH, expand=1,side='left')

page1 = ttk.Frame(nb)
page2_ = ttk.Frame(nb)

nb.add(page1 , text= 'Take Attendance ')
nb.add(page2_ , text= 'Data Management')

nb2 = ttk.Notebook(page2_,style='New.TNotebook')
nb2.pack(fill='both',expand=1)

page2 = ttk.Frame(nb2)  # CLASS
page3 = ttk.Frame(nb2)  # STUDENTS
page5 = ttk.Frame(nb2)  # USERS
page6 = ttk.Frame(nb2)  # USERS
page7 = ttk.Frame(nb2)  # RESULTS
page8 = ttk.Frame(nb2)  # REPORTS

nb2.add(page2, text = 'Classes')
nb2.add(page3, text = 'Students')
nb2.add(page7, text = 'Results')
nb2.add(page5, text = 'Staff')
nb2.add(page8, text = 'Reports')
nb2.add(page6, text = 'Sessions')

def reorder(event):
    try:
        index = nb.index(f"@{event.x},{event.y}")
        nb.insert(index, child=nb.select())

    except TclError:
        pass

nb2.bind("<B1-Motion>", reorder)

def reorder2(event):
    try:
        index = nb2.index(f"@{event.x},{event.y}")
        nb2.insert(index, child=nb2.select())

    except TclError:
        pass

nb2.bind("<B1-Motion>", reorder2)

nb.bind("<B1-Motion>", reorder)


def createdb():
    con = sqlite3.connect(database=DATABASE)
    cur = con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
    cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text,rollno text)")
    cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text,subjects text)")
    cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
    con.commit()
    cur.execute("select * from USERS where email=?",(DEFAULT_EMAIL,))
    row=cur.fetchone()
    if row==None:
        cur.execute("insert into USERS (name,email,password,role,course_code,subjects) values(?,?,?,?,?,?)",(
            DEFAULT_ADMIN,
            DEFAULT_EMAIL,
            DEFAULT_PASSWORD,
            DEFAULT_ROLE,
            DEFAULT_CLASS,
            DEFAULT_LECTURE
        ))
        con.commit()
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,email text,pass text,role text,date text,time text,rem text)")
        con.commit()
        cdate = strftime("%m/%d/%Y")
        ctime =strftime("%H:%M")
        cur.execute("insert into SESSIONS (email,pass,date,time,rem,role) values(?,?,?,?,?,?)",(
            DEFAULT_EMAIL,
            DEFAULT_PASSWORD,
            cdate,
            ctime,
            '1',
            DEFAULT_ROLE
        ))
        con.commit()

createdb()    

################################# PAGE 2

class_Frame = ttk.LabelFrame(page2 , text = 'Class Management')
class_Frame.pack(fill = BOTH, expand=1)

left_column = ttk.Frame(class_Frame)
left_column.pack(expand=1,fill = BOTH)

#### Row 1

left_column_row1 = ttk.Frame(left_column)
left_column_row1.pack(side = 'top' , fill = 'x', padx = 10 , pady = 10)

lbl_course_code = ttk.Label(left_column_row1 , text = 'Class Code :   ')
lbl_course_code.pack(side = 'left' , padx = 10)

txt_course_code = ttk.Entry(left_column_row1 , textvariable=dcourse_code)
txt_course_code.pack(side='left' , padx = 18)

lbl_course = ttk.Label(left_column_row1 , text = 'Description : ')
lbl_course.pack(side = 'left' , padx = 10)

txt_course = ttk.Entry(left_column_row1 , textvariable=dcourse)
txt_course.pack(side = 'left' , padx = 10,fill='x',expand=1)

left_column_row1__1 = ttk.Frame(left_column)
left_column_row1__1.pack(side = 'top' , fill = 'x', padx = 10 , pady = 10)

lbl_lecture_page2 = ttk.Label(left_column_row1__1 , text = 'Lectures : ')
lbl_lecture_page2.pack(side = 'left' , padx = 10)

txt_lecture_page2 = ttk.Entry(left_column_row1__1 , textvariable=dlecture , width = 45)
txt_lecture_page2.pack(side = 'left',fill='x',padx=[35,10],expand=1)

#### Row 2

left_column_row2 = ttk.Frame(left_column)
left_column_row2.pack(fill = 'both' , padx = 10 , pady = 10, expand=1)

scrollx = ttk.Scrollbar(left_column_row2 , orient='horizontal')
scrolly = ttk.Scrollbar(left_column_row2 , orient='vertical')
scrolly.pack(side = 'right',fill='y')
scrollx.pack(side = 'bottom',fill='x')

class_columns = ('Sr','Class','Code','lecture') 

class_tree = ttk.Treeview(left_column_row2 , columns=class_columns,
                xscrollcommand=scrollx.set,yscrollcommand=scrolly.set)
class_tree.pack(fill=BOTH,expand=1)

scrollx.config(command=class_tree.xview)
scrolly.config(command=class_tree.yview)

class_tree.column("#0" , width = 0 , stretch=NO)

class_tree.heading("Sr" , text = 'Sr.' , anchor='w')
class_tree.heading("Class" , text = 'Code' , anchor='w')
class_tree.heading("Code" , text = 'Description' , anchor='w')
class_tree.heading("lecture" , text = 'Lectures' , anchor='w')

class_tree.column('Sr' , width= 200 , anchor='w')
class_tree.column('Class' , width= 200 , anchor='w')
class_tree.column('Code' , width= 200 , anchor='w')
class_tree.column('lecture' , width=200 , anchor='w')
####################################### DATA BASE START ##########################################
def show_class(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text)")
        con.commit()
        cur.execute("select * from CLASSES")
        rows=cur.fetchall()
        class_tree.delete(*class_tree.get_children())
        classes.clear()
        for row in rows:
            class_tree.insert('',END,values=row)
            classes.append(row[1])
        con.close()
        if len(rows)>0:
            btn_export.configure(state='active')
        else:
            btn_export.configure(state='disabled')
        update_classes()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x1",parent=root)
        con.close()
from tkinter import filedialog

def search_class(event=None):
    try:
        con = sqlite3.connect(database=DATABASE,check_same_thread=False)
        cur = con.cursor()
        cur.execute(f"select * from CLASSES where course_code LIKE '%{txt_search.get()}%'")
        rows=cur.fetchall()
        if rows!=None and len(rows)!=0:
            class_tree.delete(*class_tree.get_children())
            for row in rows:
                class_tree.insert('',END,values=row)
            con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x16")


def export_class(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text)")
        con.commit()
        data = class_tree.selection()
        if len(data)>0:
            rows = [class_tree.item(select_item)['values'] for select_item in data]
        else:
            cur.execute("select * from CLASSES")
            rows=cur.fetchall()
        FILE_EXPORT_CLASS = filedialog.asksaveasfilename(defaultextension='.xlsx',filetype=[('Excel File','.xlsx')])
        if FILE_EXPORT_CLASS =="":
            return
        else:
            rows = pd.DataFrame(rows)
            rows.columns = ['Sr.','Class','Description','Lectures']
            os.system(f'del {FILE_EXPORT_CLASS}')
            # writing to Excel
            datatoexcel = pd.ExcelWriter(FILE_EXPORT_CLASS)
            
            # write DataFrame to excel
            rows.to_excel(datatoexcel)
            
            # save the excel
            datatoexcel.save()
            #rows.to_excel(FILE_EXPORT_CLASS)
            MessageBox("Success",f"{FILE_EXPORT_CLASS} has been exported !")
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x1",parent=root)
        con.close()


###############################################################################################
def getdata_class(event=None):
    r=class_tree.focus()
    content=class_tree.item(r)
    row=content["values"]
    dcourse_code.set(row[1])
    dcourse.set(row[2])
    dlecture.set(row[3])
    if not len(class_tree.selection())>1:
        btn_update.configure(state='active')
    else:
        btn_update.configure(state='disabled')
    btn_add.configure(state='disabled')
    btn_delete.configure(state='active')
    txt_course_code.configure(state='disabled')
       
def reset_class(event=None):
    root.title(TITLE)
    dcourse_code.set("")
    dcourse.set("")
    dlecture.set("")
    txt_course_code.configure(state='enabled')
    btn_add.configure(state='active')
    btn_update.configure(state='disabled')
    btn_delete.configure(state='disabled')
    txt_search.delete(0,END)
    show_class()

class_tree.bind("<ButtonRelease-1>",getdata_class)
####################################### FUNCTIONS ##############################################

def delete_class(event=None):
    con = sqlite3.connect(database=DATABASE,check_same_thread=False)
    cur = con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text)")
    con.commit()
    try:
        if dcourse_code.get()=="":
            MessageBox("Error","Class Code is required ! ",parent=root)
            txt_course_code.focus()
        else:
            cur.execute("select * from CLASSES where course_code=?",(dcourse_code.get(),))
            row=cur.fetchone()
            if row==None:
                MessageBox("Error","Fetch class from records !",parent=root)
            else:
                op = MessageBox("Confirm","Do you really want to delete class? ",b1='Yes',b2='No',parent=root)
                if op.choice=='Yes':
                    def progress():
                        selected_items = class_tree.selection()
                        frame = ttk.Frame(left_column)
                        frame.pack(fill = 'both' , padx = [0,0] , pady = 10)
                        frame.pack(fill = 'x' , padx = [10,10] , pady = [10,10])
                        pg = ttk.Progressbar(frame,maximum=len(selected_items),style='primary.Striped.Horizontal.TProgressbar')
                        pg.grid(row=0,column=0,sticky='ew',columnspan=2)
                        lbl = ttk.Label(frame, text='',font=font_lbl,style='primaary.TLabel')
                        lbl.grid(row=1,column=0)
                        frame.columnconfigure(1,weight=1)
                        for index,select_item in enumerate(selected_items):
                            percent = round(index/len(selected_items)*100,2)
                            current_val = class_tree.item(select_item)['values']
                            class_tree.delete(select_item)
                            cur.execute("delete from CLASSES where course_code=?",(current_val[1],))
                            con.commit()
                            pg['value'] += 1
                            lbl.configure(text=f'Please Do not click any widget of this page, Press Reset Button after its done !\nCompleted : {percent} %')
                            root.title(f"Deleting Classes : {percent} %")
                        lbl.configure(text=f'Please wait till the process finishes')
                        lbl.destroy()
                        pg.destroy()
                        frame.destroy()
                        root.title('Completed : 100 %')
                        MessageBox("Success","Class(s) has been deleted !",parent=root)
                        root.title('Reset Needed for Class Page !')
                        con.close()
                    a=threading.Thread(target=progress).start()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x2",parent=root)
        con.close()

def update_class(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text)")
        con.commit()
        cur.execute("select * from CLASSES where course_code=?",(dcourse_code.get(),))
        row=cur.fetchone()

        for i in range(1):
            if dcourse_code.get().strip()=='':
                MessageBox("Error","Enter Class Code first ! ", parent = root)
                txt_course_code.focus()
                break

            if dcourse.get().strip()=='':
                MessageBox("Error","Enter course description/name first ! ", parent = root)
                txt_course.focus()
                break

            if dlecture.get().strip()=='':
                MessageBox("Error","Enter Class lectures with comma (,) separation \nExample : A,B,C ! ", parent = root)
                txt_lecture_page2.focus()
                break

            else:
                if row==None:
                    MessageBox("Error","Fetch Class from records !",parent=root)
                else:
                    selected_items = class_tree.selection()
                    if len(selected_items)==1:
                        cur.execute("update CLASSES set degree=? , subjects=? where course_code=?",(
                        dcourse.get().strip(),
                        dlecture.get().strip(),
                        dcourse_code.get().strip()
                        ))
                        con.commit()
                        reset_class()
                        MessageBox("Success","Class has been updated ! ",parent=root)
                        con.close()
                    else:
                        MessageBox("Information","Currently Not Supported !")
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x3",parent=root)
        con.close()

def add_class(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text)")
        con.commit()
        cur.execute("select * from CLASSES where course_code=?",(dcourse_code.get(),))
        row=cur.fetchone()

        for _ in range(1):
            if dcourse_code.get().strip()=='':
                MessageBox("Error","Enter Class Code first ! ", parent = root)
                txt_course_code.focus()
                break

            if dcourse.get().strip()=='':
                MessageBox("Error","Enter course description/name first ! ", parent = root)
                txt_course.focus()
                break

            if dlecture.get().strip()=='':
                MessageBox("Error","Enter Class lectures with comma (,) separation \nExample : A,B,C ! ", parent = root)
                txt_lecture_page2.focus()
                break
            else:
                if row!=None:
                    MessageBox("Error","Class already exits !",parent=root)
                    txt_course_code.focus()
                else:
                    cur.execute("insert into CLASSES (course_code,degree,subjects) values(?,?,?)",(
                    dcourse_code.get().strip(),
                    dcourse.get().strip(),
                    dlecture.get().strip()
                    ))
                    con.commit()
                    MessageBox("Succes","The class was added successfully ! ",parent=root)
                    reset_class()
                    con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x9",parent=root)
        con.close()

####################################### DATA BASE END ##########################################
#### Row 3
######## BUTTONS 
left_column_row3 = ttk.Frame(left_column)
left_column_row3.pack(fill = 'both' , padx = [0,0] , pady = 10)

btn_add = ttk.Button(left_column_row3, text = 'Save' , width = 12 , command=add_class)
btn_add.pack(side = 'left' , padx = 10)

btn_clear = ttk.Button(left_column_row3 , text = 'Reset' , width = 12 , command=reset_class)
btn_clear.pack(side = 'left' , padx  = 10)

btn_update = ttk.Button(left_column_row3 , text = 'Update' , width = 12 , command=update_class)
btn_update.pack(side = 'left' , padx = 10)

btn_delete = ttk.Button(left_column_row3 , text = 'Delete' , width = 12 , command=delete_class)
btn_delete.pack(side = 'left' , padx = 10)

btn_sort = ttk.Button(left_column_row3 , text = 'Sort' , width = 12 , state='disabled')
btn_sort.pack(side = 'left' , padx = 10)

btn_export = ttk.Button(left_column_row3 , text = 'Export' , width = 12 , command=export_class)
btn_export.pack(side = 'left' , padx = 10)

txt_search = ttk.Entry(left_column_row3)
txt_search.pack(side='left',fill='x',expand=1,padx=10)

txt_search.bind("<KeyPress>",search_class)

btn_update.configure(state='disabled')
btn_delete.configure(state='disabled')

###########################################################################

student_Frame = ttk.LabelFrame(page3 , text='Student Managment')
student_Frame.pack(expand=1, fill=BOTH)
left_column_page3 = ttk.Frame(student_Frame)
left_column_page3.pack(expand=1,fill = BOTH)

#### Row 1

left_column_page3_row1 = ttk.Frame(left_column_page3)
left_column_page3_row1.pack(side = 'top' , fill = 'x', padx = 10 , pady = 10)

lbl_course_code_page3 = ttk.Label(left_column_page3_row1 , text='Class Code :   ')
lbl_course_code_page3.pack(side = 'left' , padx = 10)

com_codes = ()

com_course_code_page3 = ttk.Combobox(left_column_page3_row1 , state='readonly' , values=com_codes , width=18)
com_course_code_page3.pack(side = 'left' , padx = 18)

#### Row 1.1

left_column_page3_row1_1 = ttk.Frame(left_column_page3)
left_column_page3_row1_1.pack(side = 'top' , fill = 'x', padx = 10 , pady = 10)

lbl_rollno = ttk.Label(left_column_page3_row1_1 , text = 'Roll No : ')
lbl_rollno.pack(side = 'left' , padx = 11)

dname = StringVar()
drollno = StringVar()

txt_rollno = ttk.Entry(left_column_page3_row1_1 , textvariable=drollno)
txt_rollno.pack(side ='left' , padx = [42,10])

lbl_student_name = ttk.Label(left_column_page3_row1_1,text = 'Student Name : ')
lbl_student_name.pack(side = 'left' , padx = 10)

txt_student_name = ttk.Entry(left_column_page3_row1_1 , textvariable=dname , width = 45)
txt_student_name.pack(side = 'left' , padx = 10)

#### Row 2
left_column_page3_row2 = ttk.Frame(left_column_page3)
left_column_page3_row2.pack(fill = 'both' , padx = 10 , pady = 10, expand=1)

scrollx_page3 = ttk.Scrollbar(left_column_page3_row2 , orient='horizontal')
scrolly_page3 = ttk.Scrollbar(left_column_page3_row2 , orient='vertical')
scrolly_page3.pack(side = 'right',fill='y')
scrollx_page3.pack(side = 'bottom',fill='x')

class_columns_page3 = ('Sr','code','roll no','name') 

class_tree_page3 = ttk.Treeview(left_column_page3_row2 , columns=class_columns_page3,
                xscrollcommand=scrollx_page3.set,yscrollcommand=scrolly_page3.set)
class_tree_page3.pack(fill=BOTH,expand=1)

scrollx_page3.config(command=class_tree_page3.xview)
scrolly_page3.config(command=class_tree_page3.yview)

class_tree_page3.column("#0" , width = 0 , stretch=NO)

class_tree_page3.heading("Sr" , text = 'Sr.' , anchor='w')
class_tree_page3.heading("code" , text = 'Code' , anchor='w')
class_tree_page3.heading("roll no" , text = 'Name',anchor='w')
class_tree_page3.heading("name" , text = 'Roll No.' , anchor='w')

class_tree_page3.column('Sr' , width=200 , anchor='w')
class_tree_page3.column('code' , width=200 , anchor='w')
class_tree_page3.column('roll no' , width=200 , anchor='w')
class_tree_page3.column('name' , width=200 , anchor='w')

####################################### DATA BASE START ##########################################

def show_student(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text , rollno text)")
        con.commit()
        cur.execute("select * from STUDENTS")
        rows=cur.fetchall()
        class_tree_page3.delete(*class_tree_page3.get_children())
        for row in rows:
            class_tree_page3.insert('',END,values=row)
        if len(rows)>=1:
            btn_export_page3.configure(state='active')
        else:
            btn_export_page3.configure(state='disabled')
        com_course_code_page3.config(values=classes)
        con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x10",parent=root)
        con.close()

def show_student2(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text , rollno text)")
        con.commit()
        cur.execute("select * from STUDENTS where rollno=? AND course_code=?",(txt_rollno.get().strip(),com_course_code_page3.get().strip(),))
        row=cur.fetchone()
        class_tree_page3.insert('',END,values=row)
        con.close()
        txt_rollno.delete(0,END)
        txt_student_name.delete(0,END)
        txt_rollno.focus()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x10",parent=root)
        con.close()

def sort_student(event=None):
    if com_course_code_page3.get().strip()=='':
        MessageBox("Information","Select class first !")
        com_course_code_page3.focus()
    else:
        try:
            con = sqlite3.connect(database=DATABASE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text , rollno text)")
            con.commit()
            if com_course_code_page3.get().strip()=="":
                MessageBox("Error","Select Class First !")
                com_course_code_page3.focus()
            else:
                cur.execute("select * from STUDENTS where course_code=?",(com_course_code_page3.get().strip(),))
                rows=cur.fetchall()
                if rows!=None and len(rows)!=0:
                    class_tree_page3.delete(*class_tree_page3.get_children())
                    for row in rows:
                        class_tree_page3.insert('',END,values=row)
                    com_course_code_page3.config(values=classes)
                    con.close()
                else:
                    MessageBox("Information",f"There are no students in class {com_course_code_page3.get().strip()}")
        except Exception as ex:
            MessageBox("Error",f"Error due to {str(ex)} 0x10",parent=root)
            con.close()

def export_student(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text , rollno text)")
        con.commit()
        data = class_tree_page3.selection()
        if len(data)>0:
            rows = [class_tree_page3.item(select_item)['values'] for select_item in data]
        else:
            cur.execute("select * from STUDENTS")
            rows=cur.fetchall()
        rows = pd.DataFrame(rows)
        file = filedialog.asksaveasfilename(defaultextension='.xlsx',filetype=[('Excel File','.xlsx')])
        if file =="":
            return
        else:
            rows = pd.DataFrame(rows)
            rows.columns = ['Sr.','Class','Student Name','Roll No.']
            os.system(f'del {file}')
            # writing to Excel
            datatoexcel = pd.ExcelWriter(file)
            
            # write DataFrame to excel
            rows.to_excel(datatoexcel)
            
            # save the excel
            a = datatoexcel.save()
            #rows.to_excel(FILE_EXPORT_CLASS)
            MessageBox("Success",f"Student(s) has been exported !",parent=root)
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x10",parent=root)
        con.close()

###############################################################################################
def getdata_student(event=None):
    btn_sort_page3.configure(state='active')
    txt_rollno.configure(state='disabled')
    r=class_tree_page3.focus()
    content=class_tree_page3.item(r)
    row=content["values"]
    com_course_code_page3.set(row[1])
    drollno.set(row[3])
    dname.set(row[2])
    btn_add_page3.configure(state='disabled')
    btn_delete_page3.configure(state='active')
    btn_update_page3.configure(state='active')

def reset_student(event=None):
    root.title(TITLE)
    txt_rollno.configure(state='active')
    btn_add_page3.configure(state='active')
    btn_delete_page3.configure(state='disabled')
    btn_update_page3.configure(state='disabled')
    btn_sort_page3.configure(state='disabled')
    txt_search_page3.delete(0,END)
    drollno.set("")
    dname.set("")
    com_course_code_page3.set("")
    show_student()

def search_student(event=None):
    try:
        con = sqlite3.connect(database=DATABASE,check_same_thread=False)
        cur = con.cursor()
        cur.execute(f"select * from STUDENTS where name LIKE '%{txt_search_page3.get()}%'")
        rows=cur.fetchall()
        if rows!=None and len(rows)!=0:
            class_tree_page3.delete(*class_tree_page3.get_children())
            for row in rows:
                class_tree_page3.insert('',END,values=row)
            con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x16")

com_course_code_page3.bind("<<ComboboxSelected>>",lambda e, : btn_sort_page3.configure(state='active'))
class_tree_page3.bind("<ButtonRelease-1>",getdata_student)
####################################### FUNCTIONS ##############################################

def delete_student(event=None):
    con = sqlite3.connect(database=DATABASE,check_same_thread=False)
    cur = con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text , rollno text)")
    con.commit()
    try:
        if drollno.get()=="":

            MessageBox("Error","Student Roll number is required ! ",parent=root)
            txt_rollno.focus()
        else:
            cur.execute("select * from STUDENTS where rollno=?",(drollno.get(),))
            row=cur.fetchone()
            if row==None:
                MessageBox("Error","Fetch record from records !",parent=root)
            else:
                op = MessageBox("Confirm","Do you really want to delete this student ? ",parent=root,b1='Yes',b2='No')
                if op.choice=='Yes':
                    def progress():
                        selected_items = class_tree_page3.selection()
                        left_column_page3_row3.pack_forget()
                        frame = ttk.Frame(left_column_page3)
                        frame.pack(fill = 'both' , padx = [10,10] , pady = 10)
                        pg = ttk.Progressbar(frame,maximum=len(selected_items),style='primary.Striped.Horizontal.TProgressbar')
                        pg.grid(row=0,column=0,sticky='ew',columnspan=2)
                        lbl = ttk.Label(frame, text='',font=font_lbl,style='primaary.TLabel')
                        lbl.grid(row=1,column=0)
                        frame.columnconfigure(1,weight=1)
                        for index,select_item in enumerate(selected_items):
                            percent = round(index/len(selected_items)*100,2)
                            current_val = class_tree_page3.item(select_item)['values']
                            class_tree_page3.delete(select_item)
                            cur.execute("delete from STUDENTS where rollno=? AND course_code=? and cid=?",(current_val[3],current_val[1],str(current_val[0]),))
                            con.commit()
                            pg['value'] += 1
                            lbl.configure(text=f'Please Do not click any widget of this page, Press Reset Button after its done !\nCompleted : {percent} %')
                            root.title(f"Deleting Students : {percent} %")
                        lbl.configure(text=f'Please wait till the process finishes')
                        lbl.destroy()
                        pg.destroy()
                        frame.destroy()
                        root.title('Completed : 100 %')
                        left_column_page3_row3.pack(fill = 'both' , padx = [0,0] , pady = 10)
                        MessageBox("Success","Student(s) has been deleted !",parent=root)
                        root.title(TITLE)
                        con.close()
                    a=threading.Thread(target=progress).start()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x4",parent=root)
        con.close()

def update_student(event=None):
    try:
        con = sqlite3.connect(database=DATABASE,check_same_thread=False)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text , rollno text)")
        con.commit()
        cur.execute("select * from STUDENTS where rollno=?",(drollno.get(),))
        row=cur.fetchone()

        for _ in range(1):

            if com_course_code_page3.get().strip()=='':
                MessageBox("Error","Select course (code) first ! ", parent = root)
                com_course_code_page3.focus()
                break

            if drollno.get().strip()=='':
                MessageBox("Error","Enter Student Roll Number ! ", parent = root)
                txt_rollno.focus()
                break

            if dname.get().strip()=='':
                MessageBox("Error","Enter Student name ! ", parent = root)
                txt_student_name.focus()
                break
            else:

                if row==None:
                    MessageBox("Error","Fetch record from records !",parent=root)
                else:
                    selected_items = class_tree_page3.selection()
                    if len(selected_items)>1:
                        def progress():
                            selected_items = class_tree_page3.selection()
                            left_column_page3_row3.pack_forget()
                            frame = ttk.Frame(left_column_page3)
                            frame.pack(fill = 'both' , padx = [10,10] , pady = 10)
                            pg = ttk.Progressbar(frame,maximum=len(selected_items),style='primary.Striped.Horizontal.TProgressbar')
                            pg.grid(row=0,column=0,sticky='ew',columnspan=2)
                            lbl = ttk.Label(frame, text='',font=font_lbl,style='primaary.TLabel')
                            lbl.grid(row=1,column=0)
                            frame.columnconfigure(1,weight=1)
                            for index,select_item in enumerate(selected_items):
                                percent = round(index/len(selected_items)*100,2)
                                current_val = class_tree_page3.item(select_item)['values']
                                cur.execute("update STUDENTS set course_code=? where name=? AND rollno=? AND cid=?",(
                                com_course_code_page3.get(),
                                current_val[2],
                                current_val[3],
                                str(current_val[0])
                                ))
                                con.commit()
                                pg['value'] += 1
                                lbl.configure(text=f'Please Do not click any widget of this page, Press Reset Button after its done !\nCompleted : {percent} %')
                                root.title(f"Updating Students : {percent} %")
                            lbl.configure(text=f'Please wait till the process finishes')
                            lbl.destroy()
                            pg.destroy()
                            frame.destroy()
                            root.title('Completed : 100 %')
                            left_column_page3_row3.pack(fill = 'both' , padx = [0,0] , pady = 10)
                            MessageBox("Success","Student(s) has been updated ! ",parent=root)
                            root.title("Reset Needed for Student Page !")
                            con.close()
                        a=threading.Thread(target=progress).start()
                    else:
                        selected_items = class_tree_page3.selection()
                        current_val = class_tree_page3.item(selected_items)['values']
                        cur.execute("update STUDENTS set course_code=?,name=? where rollno=? AND cid=?",(
                        com_course_code_page3.get(),
                        dname.get().strip(),
                        drollno.get().strip(),
                        current_val[0]
                        ))
                        con.commit()
                        reset_student()
                        MessageBox("Success","Student(s) has been updated ! ",parent=root)
                        con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x5",parent=root)
        con.close()

def add_student(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text ,rollno text)")
        con.commit()
        cur.execute("select * from STUDENTS where rollno=? AND course_code=?",(drollno.get(),com_course_code_page3.get(),))
        row=cur.fetchone()

        for _ in range(1):

            if com_course_code_page3.get().strip()=='':
                MessageBox("Error","Select course (code) first ! ", parent = root)
                com_course_code_page3.focus()
                break

            if drollno.get().strip()=='':
                MessageBox("Error","Enter Student Roll Number ! ", parent = root)
                txt_rollno.focus()
                break

            if dname.get().strip()=='':
                MessageBox("Error","Enter Student name ! ", parent = root)
                txt_student_name.focus()
                break

            else:

                if row!=None:
                    MessageBox("Error","Student already exits !",parent=root)
                    txt_rollno.focus()
                else:
                    cur.execute("insert into STUDENTS (course_code,name,rollno) values(?,?,?)",(
                    com_course_code_page3.get(),
                    dname.get().strip(),
                    drollno.get().strip(),
                    ))
                    con.commit()
                    MessageBox("Succes","Student added successfully ! ",parent=root)
                    show_student2()
                con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x11",parent=root)
        con.close()
####################################### DATA BASE END ##########################################
#### Row 3
######## BUTTONS 

left_column_page3_row3 = ttk.Frame(left_column_page3)
left_column_page3_row3.pack(fill = 'x' , padx = [0,0] , pady = 10)

btn_add_page3 = ttk.Button(left_column_page3_row3, text = 'Save' , width = 12 , command = add_student)
btn_add_page3.pack(side = 'left' , padx = 10)

btn_clear_page3 = ttk.Button(left_column_page3_row3 , text = 'Reset' , width = 12 , command = reset_student)
btn_clear_page3.pack(side = 'left' , padx  = 10)

btn_update_page3 = ttk.Button(left_column_page3_row3 , text = 'Update' , width = 12 , command = update_student)
btn_update_page3.pack(side = 'left' , padx = 10)

btn_delete_page3 = ttk.Button(left_column_page3_row3 , text = 'Delete' , width = 12 , command = delete_student)
btn_delete_page3.pack(side = 'left' , padx = 10)

btn_sort_page3 = ttk.Button(left_column_page3_row3 , text = 'Sort' , width = 12 , command = sort_student)
btn_sort_page3.pack(side = 'left' , padx  = 10)

btn_export_page3 = ttk.Button(left_column_page3_row3 , text = 'Export' , width = 12 , command = export_student)
btn_export_page3.pack(side = 'left' , padx = 10)

txt_search_page3 = ttk.Entry(left_column_page3_row3)
txt_search_page3.pack(padx=10,side='left',fill='x',expand=1)

txt_search_page3.bind("<KeyPress>",search_student)

btn_delete_page3.configure(state='disabled')
btn_update_page3.configure(state='disabled')
btn_sort_page3.configure(state='disabled')

###########################################################################

################################# PAGE 1 ###############################

attendance_Frame = ttk.LabelFrame(page1 , text = 'Attendance Management')
attendance_Frame.pack(expand= 1 , fill =BOTH)

left_column_page1 = ttk.Frame(attendance_Frame)

left_column_page1.pack(expand=1,fill = BOTH)
attendance_Frame_row1 = ttk.LabelFrame(left_column_page1 , text='Configure')
attendance_Frame_row1.grid(row = 0 , column = 0 , sticky=W+E , pady = [5,0],padx=10)

lbl_date = ttk.Label(attendance_Frame_row1 , text = 'Date : ')
lbl_date.grid(row=0,column=0, padx = [10,0] , pady = 5)

ddate = StringVar()

# #2780e3
txt_date = ttk.DateEntry(attendance_Frame_row1)
txt_date.grid(row=0,column=1, padx = 10 , pady = 10,sticky=W+E)
"""
txt_date = ttk.DateEntry(attendance_Frame_row1,
                textvariable=ddate,state='readonly')
txt_date._set_text(txt_date._date.strftime('%m/%d/%Y'))
txt_date.grid(row=0,column=1, padx = 10 , pady = 10,sticky=W+E)

color = style.colors.primary

txt_date.config(background=color,selectbackground=color,
        headersforeground='white',headersbackground=color)
"""
lbl_class_page1 = ttk.Label(attendance_Frame_row1 , text = 'Class :')
lbl_class_page1.grid(row=0, padx = 10 , pady = 10,column=2)

com_class_page1 = ttk.Combobox(attendance_Frame_row1 , state='readonly' , values=classes)
com_class_page1.grid(row=0, padx = 10 , pady = 10,sticky=W+E,column=3)

lbl_lecture = ttk.Label(attendance_Frame_row1 , text = 'Lecture : ')
lbl_lecture.grid(row=0, padx = 10 , pady = 10,column=4)

com_lecture_page1 = ttk.Combobox(attendance_Frame_row1 , state='readonly' , values=lectures)
com_lecture_page1.grid(row=0, padx = 10 , pady = 10,sticky=E+W,column=5)

attendance_Frame_row1.columnconfigure(tuple(range(6)), weight=1)

attendance_Frame_row1_1 = ttk.LabelFrame(left_column_page1 , text='Record : ')
attendance_Frame_row1_1.grid(row = 1 , column = 0 , sticky = W+E)

left_column_page1.columnconfigure(0,weight=1,uniform=1)

lbl_total = ttk.Label(attendance_Frame_row1_1 , text='Total Students : ' , font=font_lbl)
lbl_total.grid(row = 0 , column=0 , pady=10 , padx = 10 )

lbl_total_val = ttk.Label(attendance_Frame_row1_1 , text='#### ' , font=("Helvetica",12,'bold'))
lbl_total_val.grid(row = 0 , column=1 , pady=10 , padx = 10 )

lbl_present = ttk.Label(attendance_Frame_row1_1 , text='Present : '  , font=font_lbl)
lbl_present.grid(row = 0 , column=2 , pady=10 , padx = 10 )

lbl_present_val = ttk.Label(attendance_Frame_row1_1 , text='#### ' , font=("Helvetica",12,'bold'))
lbl_present_val.grid(row = 0 , column=3 , pady=10 , padx = 10 )

lbl_absent = ttk.Label(attendance_Frame_row1_1 , text='Absent : ' , font=font_lbl)
lbl_absent.grid(row = 0 , column=4 , pady=10 , padx = 10 )

lbl_absent_val = ttk.Label(attendance_Frame_row1_1 , text='#### ' , font=("Helvetica",12,'bold'))
lbl_absent_val.grid(row = 0 , column=5 , pady=10 , padx = 10 )

lbl_leave = ttk.Label(attendance_Frame_row1_1 , text='Leave : ' , font=font_lbl)
lbl_leave.grid(row = 0 , column=6 , pady=10 , padx = 10 )

lbl_leave_val = ttk.Label(attendance_Frame_row1_1 , text='#### ' , font=("Helvetica",12,'bold'))
lbl_leave_val.grid(row = 0 , column=7 , pady=10 , padx = 10 )

attendance_Frame_row1_1.columnconfigure(tuple(range(8)),weight=1)

lbls_information = [
    lbl_total,
    lbl_total_val,
    lbl_present,
    lbl_present_val,
    lbl_absent,
    lbl_absent_val,
    lbl_leave,
    lbl_leave_val,
]

for i in lbls_information:
    i.grid_forget()

attendance_Frame_row2 = ttk.LabelFrame(left_column_page1 , text = 'Attendance : ')
#attendance_Frame_row2.grid(row = 2 , column = 0 ,padx=10, pady = 5 , sticky = W+E+N+S)

attendance_Frame_row3 = ttk.Frame(left_column_page1)

attendance_Frame_row_2 = ttk.Frame(left_column_page1)
attendance_Frame_row_2.grid(row = 3 , column = 0,pady=[68,10] , sticky = S+N+W+E,padx=10)

attendance_Frame_row_3_3 = ttk.Frame(left_column_page1)
attendance_Frame_row_3_3.grid(row = 4 , column = 0 , sticky = W+E,padx=2)

btn_save_prime = ttk.Button(attendance_Frame_row_3_3,text='Save' , width=12)
btn_save_prime.pack(side='left',padx=10,pady=10)

btn_reset_prime = ttk.Button(attendance_Frame_row_3_3,text='Reset' , width=12)
btn_reset_prime.pack(side='left',padx=10,pady=10)

btn_update_prime = ttk.Button(attendance_Frame_row_3_3,text='Update' , width=12)
btn_update_prime.pack(side='left',padx=10,pady=10)

btn_delete_prime = ttk.Button(attendance_Frame_row_3_3,text='Delete' , width=12)
btn_delete_prime.pack(side='left',padx=10,pady=10)

#attendance_Frame_row_3 = ttk.Frame(left_column_page1)
#attendance_Frame_row_3.grid(row = 4 , column = 0 , pady = 10 , sticky = W+E)
#########################################################################################

scrollx = ttk.Scrollbar(attendance_Frame_row_2 , orient='horizontal')
scrolly = ttk.Scrollbar(attendance_Frame_row_2 , orient='vertical')
scrolly.pack(side = 'right',fill='y')
scrollx.pack(side = 'bottom',fill='x')

class_columns = ('Sr','code','lecture','date','students','status') 

class_tree_page1 = ttk.Treeview(attendance_Frame_row_2 , columns=class_columns,
                xscrollcommand=scrollx.set,yscrollcommand=scrolly.set)
class_tree_page1.pack(fill=BOTH,expand=1)

scrollx.config(command=class_tree_page1.xview)
scrolly.config(command=class_tree_page1.yview)

class_tree_page1.column("#0" , width = 0 , stretch=NO)

class_tree_page1.heading("Sr" , text = 'Sr.' , anchor='w')
class_tree_page1.heading("code" , text = 'Class' , anchor='w')
class_tree_page1.heading("lecture" , text = 'Lecture' , anchor='w')
class_tree_page1.heading("date" , text = 'Date' , anchor='w')
class_tree_page1.heading("students" , text = 'Students' , anchor='w')
class_tree_page1.heading("status" , text = 'Status' , anchor='w')

class_tree_page1.column('Sr' , width= 200 , anchor='w')
class_tree_page1.column('code' , width= 200 , anchor='w')
class_tree_page1.column('lecture' , width=200 , anchor='w')
class_tree_page1.column('date' , width=200 , anchor='w')
class_tree_page1.column('students' , width=200 , anchor='w')
class_tree_page1.column('status' , width=200 , anchor='w')

def show_attendance(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
        # Course , Subject , date , status
        con.commit()
        cur.execute("""
        select * from ATTENDANCE
        """)
        rows=cur.fetchall()
        if rows!=None:
            class_tree_page1.delete(*class_tree_page1.get_children())
            for row in rows:
                class_tree_page1.insert('',END,values=row)
    except Exception as ex:
        print(str(ex))
show_attendance()
#########################################################################################
legend = ttk.Label(attendance_Frame_row3 , text='' , font=14)
legend.pack(side = 'right' , padx = 10)

def clear_classes():
    for i in attendance_Frame_row2.winfo_children():
        i.destroy()
    attendance_Frame_row2.grid_forget()
    attendance_Frame_row3.grid_forget()

def make_tuples_to_list(data):
    newData =  [list(i) for i in data]
    return newData

def add_item(data,item):
    li = []
    for j,i in enumerate(data):
        i.append(item[j])
        li.append(i)
    return li

def pop_first(data):
    li = []
    for i in data:
        i.pop(0)
        li.append(i)
    return li

def get_lectures(event=None,clear=True):
    try:
        if clear:
            clear_classes()
            attendance_Frame_row1_1.grid_forget()
        com_class_page1.config(values=classes)
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text,subjects text)")
        con.commit()
        cur.execute("select * from CLASSES where course_code=?",(com_class_page1.get().strip(),))
        rows=cur.fetchall()
        if rows==None:
            pass
        else:
            rows = make_tuples_to_list(rows)
            com_lecture_page1.config(values=rows[0][3].split(','))
    except Exception as ex:
        #print('Error due to ',ex,' 0x13')
        com_lecture_page1.config(values=())
        com_lecture_page1.set("")

def get_lectures_for_staff(event=None):
    btn_sort_page5.configure(state='active')

    if com_class_page5.get().strip()=="":
        MessageBox("Information","Select class first !")
        com_class_page5.focus()
    else:
        try:
            com_class_page5.config(values=classes)
            con = sqlite3.connect(database=DATABASE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text,subjects text)")
            con.commit()
            cur.execute("select * from CLASSES where course_code=?",(com_class_page5.get().strip(),))
            rows=cur.fetchall()
            if rows==None:
                pass
            else:
                rows = make_tuples_to_list(rows)
                com_lecture_page5.config(values=rows[0][3].split(','))
        except Exception as ex:
            #print('Error due to ',ex,' 0x13')
            com_lecture_page5.config(values=())
            com_lecture_page5.set("")

def get_classes(event=None):
    com_class_page1.config(values=classes)
    get_lectures(clear=True)

obj = object
students_information = []
boxes_infomration = None
took = None
def get_students(event=None):
    global students_information
    global boxes_infomration
    global obj
    try:
        clear_classes()
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text,rollno text)")
        con.commit()
        cur.execute("select * from STUDENTS where course_code=?",[com_class_page1.get()])
        rows=cur.fetchall()
        rows = make_tuples_to_list(rows)
        rows = pop_first(rows)
        rows = pop_first(rows)
        attendance_Frame_row_3_3.grid_forget()
        attendance_Frame_row_2.grid_forget()
        attendance_Frame_row2.grid(row = 3 , column = 0 , pady = 10 , sticky = S+N+W+E,padx=10)
        attendance_Frame_row3.grid(row = 4 , column = 0 , pady = 10 , sticky = W+E)
        MyFrameObject = MyFrame(attendance_Frame_row2,window=root)
        obj = MyFrameObject
        students_information = rows.copy()
        MyFrameObject.make_rows(rows)
        boxes_infomration = MyFrameObject.Status_boxes()
        con.close()
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
        # Course , Subject , date , status
        con.commit()
        cur.execute("select * from ATTENDANCE where date=? AND subjects=? AND course_code=?",(txt_date.entry.get(),com_lecture_page1.get(),com_class_page1.get(),))
        row=cur.fetchone()
        absent = 0
        leave = 0
        if row!=None:
            #print('record already present')
            if row[4]!=None:
                volcano = [i for i,j in enumerate(students_information) if j[1] in row[4].split(',')]
                volcano2 = row[5].split(',')
                for index,value in enumerate(volcano):
                    boxes_infomration[value].set(volcano2[index])
                    if volcano2[index]=='Absent':
                        absent += 1
                        obj.col_0[value].config(fg='red')
                        obj.col_1[value].config(fg='red')
                        obj.col_2[value].config(fg='red')
                    elif volcano2[index]=='Leave':
                        leave += 1
                        obj.col_0[value].config(fg='blue')
                        obj.col_1[value].config(fg='blue')
                        obj.col_2[value].config(fg='blue')
                legend.configure(text='Attedance Recorded ! ')
                attendance_Frame_row1_1.grid(row = 1 , column = 0 , padx = 10 , pady = 10 , sticky=W+E)
                lbl_total_val.configure(text=len(obj.col_0))
                lbl_present_val.configure(text = len(obj.col_0)-len(volcano))
                lbl_absent_val.configure(text=absent)
                lbl_leave_val.configure(text=leave)

                for index,label in enumerate(lbls_information):
                    label.grid(row = 0 , column = index , padx = 10 , pady = 10)
                return True
            else:
                return False
    except Exception as ex:
        MessageBox("Error",f"Error due to {ex}. 0x14")
        con.close()

def get_studentsCsv(row):
    global students_information
    global boxes_infomration
    global obj
    try:
        clear_classes()
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text,rollno text)")
        con.commit()
        cur.execute("select * from STUDENTS where course_code=?",[row[1]])
        rows=cur.fetchall()
        rows = make_tuples_to_list(rows)
        rows = pop_first(rows)
        rows = pop_first(rows)
        attendance_Frame_row_3_3.grid_forget()
        attendance_Frame_row_2.grid_forget()
        attendance_Frame_row2.grid(row = 3 , column = 0 , pady = 10 , sticky = S+N+W+E,padx=10)
        attendance_Frame_row3.grid(row = 4 , column = 0 , pady = 10 , sticky = W+E)
        MyFrameObject = MyFrame(attendance_Frame_row2,window=root)
        obj = MyFrameObject
        students_information = rows.copy()
        MyFrameObject.make_rows(rows)
        boxes_infomration = MyFrameObject.Status_boxes()
        con.close()
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
        # Course , Subject , date , status
        con.commit()
        cur.execute("select * from ATTENDANCE where date=? AND subjects=? AND course_code=?",(row[3],row[2],row[1],))
        row=cur.fetchone()
        absent = 0
        leave = 0
        if row!=None:
            #print('record already present')
            if row[4]!=None:
                volcano = [i for i,j in enumerate(students_information) if j[1] in row[4].split(',')]
                volcano2 = row[5].split(',')
                for index,value in enumerate(volcano):
                    boxes_infomration[value].set(volcano2[index])
                    if volcano2[index]=='Absent':
                        absent += 1
                        obj.col_0[value].config(fg='red')
                        obj.col_1[value].config(fg='red')
                        obj.col_2[value].config(fg='red')
                    elif volcano2[index]=='Leave':
                        leave += 1
                        obj.col_0[value].config(fg='blue')
                        obj.col_1[value].config(fg='blue')
                        obj.col_2[value].config(fg='blue')
                legend.configure(text='Attedance Recorded ! ')
                attendance_Frame_row1_1.grid(row = 1 , column = 0 , padx = 10 , pady = 10 , sticky=W+E)
                lbl_total_val.configure(text=len(obj.col_0))
                lbl_present_val.configure(text = len(obj.col_0)-len(volcano))
                lbl_absent_val.configure(text=absent)
                lbl_leave_val.configure(text=leave)

                for index,label in enumerate(lbls_information):
                    label.grid(row = 0 , column = index , padx = 10 , pady = 10)
                return True
            else:
                return False
    except Exception as ex:
        MessageBox("Error",f"Error due to {ex}. 0x14")
        con.close()


def update_attendance(event=None):
    try:
        global obj
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
        # Course , Subject , date , status
        con.commit()
        cur.execute("select * from ATTENDANCE where date=? AND subjects=? AND course_code=?",(txt_date.entry.get(),com_lecture_page1.get(),com_class_page1.get(),))
        row=cur.fetchone()
        for i in range(1):
            if com_class_page1.get()=='':
                MessageBox("Error","Select a class first ! ")
                com_class_page1.focus()
                break

            if com_lecture_page1.get()=='':
                MessageBox("Error","Select a lecture first ! ")
                com_lecture_page1.focus()
                break

            if txt_date.entry.get()=='':
                MessageBox("Error","Select date first ! ")
                txt_date.focus()
                break
            else:
                if row==None:
                    MessageBox("Error","Attendance record does not exit !. Please provide required information or re-take attendance.",parent=root)
                else:
                    boxes_infomration = obj.Status_boxes()

                    statuses = [i.get() for i in boxes_infomration]
                    string = ""
                    string2 = ""
                    statuses = [i.get() for i in boxes_infomration]
                    roll_nums = [i[1] for index,i in enumerate(students_information) if statuses[index]!='Present']
                    roll_nums = tuple(roll_nums)
                    reasons = tuple([statuses[index] for index,i in enumerate(students_information) if statuses[index]!='Present'])
                    if len(roll_nums)==0:
                        string = None
                        string2 = None
                    else:
                        for i in roll_nums:
                            string += f'{i},'
                        string = string[0:len(string)-1]

                        for i in reasons:
                            string2 += f'{i},'
                        string2 = string2[0:len(string2)-1]
                    cur.execute("update ATTENDANCE set students=?,status=? where course_code=? AND date=? AND subjects=?",(
                    string,
                    string2,
                    com_class_page1.get().strip(),
                    txt_date.entry.get().strip(),
                    com_lecture_page1.get().strip(),
                    ))
                    con.commit()
                    record()
                    MessageBox("Success",'Attendance Records have been updated ! ',parent=root)
                    con.close()
    except Exception as ex:
        print(str(ex))
        con.close()

def add_attendance(event=None):
    try:
        global obj
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,email text,pass text,role text,date text,time text,rem text)")
        # email , pass , role , date , time , rem
        cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
        # Course , Subject , RollNo , date , status
        con.commit()
        cur.execute("select * from ATTENDANCE where date=? AND subjects=? AND course_code=?",(txt_date.entry.get(),com_lecture_page1.get(),com_class_page1.get(),))
        row=cur.fetchone()
        for i in range(1):

            if com_class_page1.get()=='':
                MessageBox("Error","Select a class first ! ")
                com_class_page1.focus()
                break

            if com_lecture_page1.get()=='':
                MessageBox("Error","Select a lecture first ! ")
                com_lecture_page1.focus()
                break

            if txt_date.entry.get()=='':
                MessageBox("Error","Select date first ! ")
                txt_date.focus()
                break
            else:

                if row!=None:
                    MessageBox("Error","Attendance record already exits !",parent=root)
                    legend.configure(text='Attedance Recorded ! ')
                else:
                    string = ""
                    string2 = ""
                    boxes_infomration = obj.Status_boxes()
                    statuses = [i.get() for i in boxes_infomration]
                    roll_nums = [i[1] for index,i in enumerate(students_information) if statuses[index]!='Present']
                    roll_nums = tuple(roll_nums)
                    reasons = tuple([statuses[index] for index,i in enumerate(students_information) if statuses[index]!='Present'])
                    if len(roll_nums)<0:
                        string = None
                        string2 = None
                    else:
                        for i in roll_nums:
                            string += f'{i},'
                        string = string[0:len(string)-1]

                        for i in reasons:
                            string2 += f'{i},'
                        string2 = string2[0:len(string2)-1]
                    a = txt_date.entry.get().strip().split('/')
                    a.pop(1)
                    a="/".join(a)
                    cur.execute("insert into ATTENDANCE (course_code,subjects,date,students,status,identifier) values(?,?,?,?,?,?)",(
                    com_class_page1.get().strip(),
                    com_lecture_page1.get().strip(),
                    txt_date.entry.get().strip(),
                    string,
                    string2,
                    a
                    ))
                    con.commit()
                    record()
                    MessageBox("Succes","Attendance records were added successfully ! ",parent=root)
                    legend.configure(text='Attedance Recorded !')
                con.close()
    except Exception as ex:
        print(str(ex))
        con.close()

def reset_page1(event=None):
    attendance_Frame_row1_1.grid_forget()
    attendance_Frame_row3.grid_forget()
    attendance_Frame_row2.grid_forget()
    attendance_Frame_row_2.grid(row = 3 , column = 0 ,pady=[68,10] , sticky = S+N+W+E,padx=10)
    attendance_Frame_row_3_3.grid(row = 4 , column = 0 , sticky = W+E)
    show_attendance()

    btn_save.configure(state='active')
    btn_delete_page1.configure(state='disabled')
    btn_update_page1.configure(state='disabled')

    com_lecture_page1.set("")
    com_class_page1.set("")
    com_lecture_page1.config(values=())
    for i in lbls_information:
        i.grid_forget()
    legend.configure(text='')

def delete_attendance(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
        # Course , Subject , date , status
        con.commit()
        cur.execute("select * from ATTENDANCE where date=? AND subjects=? AND course_code=?",(txt_date,com_lecture_page1.get(),com_class_page1.get(),))
        row=cur.fetchone()
        if row==None:
            MessageBox("Error","There's no such record ! ",parent = root)
        else:
            confirm = MessageBox("Confirm",'Do you really want to delete these records? ',b1='Yes',b2='No',parent=root)
            if confirm.choice=='Yes':
                cur.execute("delete from ATTENDANCE where course_code=? AND date=? AND subjects=?",(com_class_page1.get(),txt_date,com_lecture_page1.get(),))
                con.commit()
                reset_page1()
                MessageBox("Success",'Records have been deleted ! ')
                con.close()
    except Exception as ex:
        print(str(ex))
        con.close()

def record(event=None,get=True,*args):
    global obj
    for i in range(1):

        if com_class_page1.get()=='':
            MessageBox("Error","Select a class first ! ")
            com_class_page1.focus()
            break

        if com_lecture_page1.get()=='':
            MessageBox("Error","Select a lecture first ! ")
            com_lecture_page1.focus()
            break

        if txt_date=='':
            MessageBox("Error","Select date first ! ")
            txt_date.focus()
            break
        else:
            if get:
                attendance_Frame_row1_1.grid_forget()
                for i in lbls_information:
                    i.grid_forget()
                get_students()
            try:
                con = sqlite3.connect(database=DATABASEATTENDNACE)
                cur = con.cursor()
                cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
                con.commit()
                cur.execute("select * from ATTENDANCE where date=? AND subjects=? AND course_code=?",(txt_date,com_lecture_page1.get(),com_class_page1.get(),))
                row=cur.fetchone()
                if row!=None:
                    legend.configure(text='Attendance Recorded ! ')
                    btn_save.configure(state='disabled')
                    btn_update_page1.configure(state='active')
                    btn_delete_page1.configure(state='active')
                else:
                    legend.configure(text='')
                    btn_save.configure(state='active')
                    btn_update_page1.configure(state='disabled')
                    btn_delete_page1.configure(state='disabled')
                con.close()
            except Exception as ex:
                print(f'Error due to {ex}. ox16')
                con.close()

def recordCsv(event=None,*args):
    r=class_tree_page1.focus()
    content=class_tree_page1.item(r)
    row=content["values"]
    print(row)
    global obj
    attendance_Frame_row1_1.grid_forget()
    for i in lbls_information:
        i.grid_forget()
    get_studentsCsv(row)
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,subjects text,date text,students text,status blob,identifier text)")
        con.commit()
        cur.execute("select * from ATTENDANCE where date=? AND subjects=? AND course_code=?",(row[3],row[2],row[1],))
        row=cur.fetchone()
        if row!=None:
            legend.configure(text='Attendance Recorded ! ')
            btn_save.configure(state='disabled')
            btn_update_page1.configure(state='active')
            btn_delete_page1.configure(state='active')
        else:
            legend.configure(text='')
            btn_save.configure(state='active')
            btn_update_page1.configure(state='disabled')
            btn_delete_page1.configure(state='disabled')
        con.close()
    except Exception as ex:
        print(f'Error due to {ex}. ox16')
        con.close()
btn_save = ttk.Button(attendance_Frame_row3 , text = 'Save' , width = 12 , command = add_attendance)
btn_save.pack(side = 'left' , padx = 10 , pady = 5)

btn_clear_page1 = ttk.Button(attendance_Frame_row3 , text = 'Reset' , width = 12 , command = reset_page1)
btn_clear_page1.pack(side = 'left' , padx = 10 , pady = 5)

btn_update_page1 = ttk.Button(attendance_Frame_row3 , text = 'Update' , width = 12 , command = update_attendance)
btn_update_page1.pack(side = 'left' , padx = 10 , pady = 5)

btn_delete_page1 = ttk.Button(attendance_Frame_row3 , text = 'Delete' , width = 12 , command=delete_attendance)
btn_delete_page1.pack(side = 'left' , padx = 10 , pady = 5)

btn_save.configure(state='disabled')
btn_delete_page1.configure(state='disabled')

get_lectures(clear=True)

get_lectures()

################################## END PAGE 1 ##########################

################
#nb.select(1)
#nb2.select(3)
################################# PAGE 2

page5_Frame = ttk.LabelFrame(page5 , text = 'Staff Management')
page5_Frame.pack(fill = BOTH, expand=1)

dteacher_name = StringVar()
demail = StringVar()
dpassword = StringVar()

left_column_page5 = ttk.Frame(page5_Frame)
left_column_page5.pack(expand=1,fill = BOTH)

#### Row 1
left_column_page5_row1 = ttk.Frame(left_column_page5)
left_column_page5_row1.pack(side = 'top' , fill = 'x', padx = 10)

lbl_class_page5 = ttk.Label(left_column_page5_row1 , text='Class Code :   ')
lbl_class_page5.grid(row=0,column=0,pady=10)

com_class_page5 = ttk.Combobox(left_column_page5_row1 , state='readonly',values=classes , width=23)
com_class_page5.grid(row=0,column=1,pady=10)

lbl_lecture_page5 = ttk.Label(left_column_page5_row1, text='Lecture : ')
lbl_lecture_page5.grid(row=0,column=2,pady=10)

com_lecture_page5 = ttk.Combobox(left_column_page5_row1 , state='readonly' , values=lectures, width = 23)
com_lecture_page5.grid(row=0,column=3,pady=10)

lbl_role = ttk.Label(left_column_page5_row1 , text = 'Role : ')
lbl_role.grid(row=0,column=4)

com_roles = ('Admin','Teacher','Assistant')

com_role = ttk.Combobox(left_column_page5_row1 , state='readonly',values=com_roles , width = 23)
com_role.grid(row=0,column=5)

com_class_page5.bind('<<ComboboxSelected>>', lambda e :get_lectures_for_staff(e))

lbl_teacher_name = ttk.Label(left_column_page5_row1 , text = 'Teacher Name : ')
lbl_teacher_name.grid(row=1,column=0,pady=10)

txt_teacher_name = ttk.Entry(left_column_page5_row1 , width = 25 , textvariable=dteacher_name)
txt_teacher_name.grid(row=1,column=1,pady=10)

lbl_email = ttk.Label(left_column_page5_row1 , text = 'Email Address : ')
lbl_email.grid(row=1,column=2,pady=10)

txt_email = ttk.Entry(left_column_page5_row1 , width = 25 , textvariable=demail)
txt_email.grid(row=1,column=3,pady=10)

lbl_password = ttk.Label(left_column_page5_row1 , text = 'Password : ')
lbl_password.grid(row=1,column=4,pady=10)

txt_password = ttk.Entry(left_column_page5_row1 , width = 25 , textvariable=dpassword)
txt_password.grid(row=1,column=5,pady=10)

left_column_page5_row1.columnconfigure((0,1,2,3,4,5),weight=1)

left_column_page5_row2 = ttk.Frame(left_column_page5)
left_column_page5_row2.pack(fill = 'both' , padx = 10 , pady = [10,20], expand=1)

scrollx_page5 = ttk.Scrollbar(left_column_page5_row2 , orient='horizontal')
scrolly_page5 = ttk.Scrollbar(left_column_page5_row2 , orient='vertical')
scrolly_page5.pack(side = 'right',fill='y')
scrollx_page5.pack(side = 'bottom',fill='x')

class_columns_page5 = ('Sr','name','role','code','lecture','email','pass') 

class_tree_page5 = ttk.Treeview(left_column_page5_row2 , columns=class_columns_page5,
                xscrollcommand=scrollx_page5.set,yscrollcommand=scrolly_page5.set)
class_tree_page5.pack(fill=BOTH,expand=1)

scrollx_page5.config(command=class_tree_page5.xview)
scrolly_page5.config(command=class_tree_page5.yview)

class_tree_page5.column("#0" , width = 0 , stretch=NO)

class_tree_page5.heading("Sr" , text = 'Sr.' , anchor='w')
class_tree_page5.heading("name" , text = 'Name' , anchor='w')
class_tree_page5.heading("role" , text = 'Role' , anchor='w')
class_tree_page5.heading("code" , text = 'Class',anchor='w')
class_tree_page5.heading("lecture" , text = 'Lecture' , anchor='w')
class_tree_page5.heading("email" , text = 'Email' , anchor='w')
class_tree_page5.heading("pass" , text = 'Password' , anchor='w')

class_tree_page5.column('Sr' , width=200 , anchor='w')
class_tree_page5.column('code' , width=200 , anchor='w')
class_tree_page5.column('lecture' , width=200 , anchor='w')
class_tree_page5.column('name' , width=200 , anchor='w')
class_tree_page5.column('role' , width=200 , anchor='w')
class_tree_page5.column('email' , width=200 , anchor='w')
class_tree_page5.column('pass' , width=200 , anchor='w')

class_tree_page5["displaycolumns"]=('Sr', 'code' , 'lecture' , 'name' , 'role' , 'email')

###############################################################################################

def show_user(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
        con.commit()
        cur.execute("select * from USERS")
        rows=cur.fetchall()
        if rows!=None:
            class_tree_page5.delete(*class_tree_page5.get_children())
            for row in rows:
                class_tree_page5.insert('',END,values=row)
            if len(rows)>=1:
                btn_export_page5.configure(state='active')
            else:
                btn_export_page5.configure(state='disabled')
        con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x14",parent=root)
        con.close()

def show_user2(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
        con.commit()
        cur.execute("select * from USERS where email=?",(demail.get().strip(),))
        row=cur.fetchone()
        class_tree_page5.insert('',END,values=row)
        con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x14",parent=root)
        con.close()

def sort_user(event=None):
    if com_class_page5.get().strip()=="":
        MessageBox("Information","Select class first !")
        com_class_page5.focus()
    else:
        try:
            con = sqlite3.connect(database=DATABASE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
            con.commit()
            cur.execute("select * from USERS where course_code=?",(com_class_page5.get().strip(),))
            rows=cur.fetchall()
            if rows!=None and len(rows)!=0:
                class_tree_page5.delete(*class_tree_page5.get_children())
                for row in rows:
                    class_tree_page5.insert('',END,values=row)
                con.close()
            else:
                MessageBox("Information",'There are no records having this class !')
        except Exception as ex:
            MessageBox("Error",f"Error due to {str(ex)}. 0x14",parent=root)
            con.close()

def export_user(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
        con.commit()
        data = class_tree_page5.selection()
        if len(data)>0:
            rows = [class_tree_page5.item(select_item)['values'] for select_item in data]
        else:
            cur.execute("select * from USERS")
            rows=cur.fetchall()
        file = filedialog.asksaveasfilename(defaultextension='.xlsx',filetype=[('Excel File','.xlsx')])
        if file =="":
            return
        else:
            rows = pd.DataFrame(rows)
            rows.columns = ['Sr.','User Name','Role','Class','Subject','Email','Password']
            os.system(f'del {file}')
            # writing to Excel
            datatoexcel = pd.ExcelWriter(file)
            
            # write DataFrame to excel
            rows.to_excel(datatoexcel)
            
            # save the excel
            a = datatoexcel.save()
            #rows.to_excel(FILE_EXPORT_CLASS)
            MessageBox("Success",f"Users(s) has been exported !",parent=root)
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x14",parent=root)
        con.close()


def getdata_user(event=None):
    btn_sort_page5.configure(state='active')
    txt_email.configure(state='disabled')
    r=class_tree_page5.focus()
    content=class_tree_page5.item(r)
    row=content["values"]
    dteacher_name.set(row[1])
    demail.set(row[5])
    dpassword.set(row[6])
    com_class_page5.set(row[3])
    com_lecture_page5.set(row[4])
    com_role.set(row[2])
    get_lectures_for_staff()
    btn_delete_page5.configure(state='active')
    if not len(class_tree_page5.selection())>1:
        btn_update_page5.configure(state='active')
    else:
        btn_update_page5.configure(state='disabled')
    btn_add_page5.configure(state='disabled')

    #com_course_code_page3.set(row[1])

def reset_user(event=None):
    btn_sort_page5.configure(state='disabled')
    txt_email.configure(state='enabled')
    dteacher_name.set("")
    demail.set("")
    dpassword.set("")
    com_lecture_page5.set("")
    com_class_page5.set("")
    com_role.set("")
    txt_search_page5.delete(0,END)
    com_lecture_page5.configure(values=())
    btn_delete_page5.configure(state='disabled')
    btn_update_page5.configure(state='disabled')
    btn_add_page5.configure(state='active')
    show_user()

def reset_user2(event=None):
    show_user2()
    txt_email.configure(state='enabled')
    dteacher_name.set("")
    demail.set("")
    dpassword.set("")
    com_lecture_page5.set("")
    com_class_page5.set("")
    com_role.set("")
    com_lecture_page5.configure(values=())
    btn_delete_page5.configure(state='disabled')
    btn_update_page5.configure(state='disabled')
    btn_add_page5.configure(state='active')


def search_user(event=None):
    try:
        con = sqlite3.connect(database=DATABASE,check_same_thread=False)
        cur = con.cursor()
        cur.execute(f"select * from USERS where name LIKE '%{txt_search_page5.get()}%'")
        rows=cur.fetchall()
        if rows!=None and len(rows)!=0:
            class_tree_page5.delete(*class_tree_page5.get_children())
            for row in rows:
                class_tree_page5.insert('',END,values=row)
            con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x16")

class_tree_page5.bind("<ButtonRelease-1>",getdata_user)
####################################### FUNCTIONS ##############################################
###############################################################################

def delete_user(event=None):
    try:
        con = sqlite3.connect(database=DATABASE,check_same_thread=False)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
        con.commit()
        cur.execute("Select * from USERS where email=?",(demail.get(),))
        row = cur.fetchone()
        if row==None:
            MessageBox("Error",'Select user from records !')
        else:
            op = MessageBox("Confirm","Do you really wish to delete ? ",b1='Yes',b2='No')
            if op.choice=='Yes':
                def progress():
                    selected_items = class_tree_page5.selection()
                    left_column_page5_row3.pack_forget()
                    frame = ttk.Frame(left_column_page5)
                    frame.pack(fill = 'x' , padx = [10,10] , pady = [10,10])
                    pg = ttk.Progressbar(frame,maximum=len(selected_items),style='primary.Striped.Horizontal.TProgressbar')
                    pg.grid(row=0,column=0,sticky='ew',columnspan=2)
                    lbl = ttk.Label(frame, text='',font=font_lbl,style='primaary.TLabel')
                    lbl.grid(row=1,column=0)
                    frame.columnconfigure(1,weight=1)
                    for index,select_item in enumerate(selected_items):
                        percent = round(index/len(selected_items)*100,2)
                        current_val = class_tree_page5.item(select_item)['values']
                        class_tree_page5.delete(select_item)
                        cur.execute("delete from USERS where cid=?",(current_val[0],))
                        con.commit()
                        pg['value'] += 1
                        lbl.configure(text=f'Please Do not click any widget of this page, Press Reset Button after its done !\nCompleted : {percent} %')
                        root.title(f"Deleting Users : {percent} %")
                    lbl.configure(text=f'Please wait till the process finishes')
                    root.title(f"Reset Needed for Staff Page")
                    lbl.destroy()
                    pg.destroy()
                    frame.destroy()
                    left_column_page5_row3.pack(fill = 'both' , padx = [0,0] , pady = [0,10])
                    root.title('Completed : 100 %')
                    MessageBox("Success","User(s) has been deleted !",parent=root)
                    root.title(TITLE)
                a=threading.Thread(target=progress).start()
    except Exception as ex:
        MessageBox("Erorr",f'Error due to {str(ex)}')

def update_user(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
        con.commit()
        cur.execute("Select * from USERS where email=?",(demail.get(),))
        row = cur.fetchone()
        if row==None:
            MessageBox("Error",'Select user from records !')
        else:
            selected_items = class_tree_page5.selection()
            for _ in range(1):
                if len(selected_items)>1:
                    MessageBox("Error","You cannot update multiple users at once\nCurrently Not Supported !")
                else:
                    if com_class_page5.get().strip()=='' or com_class_page5.get().strip()=='None' :
                        MessageBox("Error","Select class first ! ")
                        com_class_page5.focus()
                        break
                    if dteacher_name.get().strip()=='':
                        MessageBox("Error","Enter User Name ! ")
                        txt_teacher_name.focus()
                        break
                    if com_lecture_page5.get().strip()=='' or com_lecture_page5.get().strip()=='None':
                        MessageBox("Error","Select lecture ! ")
                        com_lecture_page5.focus()
                        break
                    if demail.get().strip()=='':
                        MessageBox("Error","Enter Email Address ! ")
                        txt_email.focus()
                        break
                    if com_role.get().strip()=='' or com_role.get().strip()=='None':
                        MessageBox("Error","Select role for user ! ")
                        com_role.focus()
                        break
                    if dpassword.get().strip()=='':
                        MessageBox("Error","Enter Password ! ")
                        txt_password.focus()
                        break
                    else:
                        selected_items = class_tree_page5.selection()
                        for select_item in selected_items:
                            current_val = class_tree_page5.item(select_item)['values']
                            cur.execute("update USERS set name=? , role=? , course_code=? , subjects=? , email=?, password=? where cid=?",(
                            dteacher_name.get().strip(),
                            com_role.get().strip(),
                            com_class_page5.get().strip(),
                            com_lecture_page5.get().strip(),
                            demail.get().strip(),
                            dpassword.get().strip(),
                            str(current_val[0])
                            ))
                            con.commit()
                        reset_user()
                        MessageBox("Success","User(s) has been updated !")
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}")
        con.close()

def add_user(event=None):
    try:
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
        con.commit()
        cur.execute("Select * from USERS where email=?",(demail.get(),))
        row = cur.fetchone()
        if row!=None:
            MessageBox("Error",'User already exits !')
        else:
            for _ in range(1):
                if com_class_page5.get().strip()=='':
                    MessageBox("Error","Select class first ! ")
                    com_class_page5.focus()
                    break
                if dteacher_name.get().strip()=='':
                    MessageBox("Error","Enter User Name ! ")
                    txt_teacher_name.focus()
                    break
                if com_lecture_page5.get().strip()=='':
                    MessageBox("Error","Select lecture ! ")
                    com_lecture_page5.focus()
                    break
                if demail.get().strip()=='':
                    MessageBox("Error","Enter Email Address ! ")
                    txt_email.focus()
                    break
                if com_role.get().strip()=='':
                    MessageBox("Error","Select role for user ! ")
                    com_role.focus()
                    break
                if dpassword.get().strip()=='':
                    MessageBox("Error","Enter Password ! ")
                    txt_password.focus()
                    break
                else:
                    cur.execute("Insert into USERS (name,role,course_code,subjects,email,password) values(?,?,?,?,?,?)",(
                        dteacher_name.get().strip(),
                        com_role.get().strip(),
                        com_class_page5.get().strip(),
                        com_lecture_page5.get().strip(),
                        demail.get().strip(),
                        dpassword.get().strip()
                    ))
                    con.commit()
                    MessageBox("Sucess","User has been added !")
                    reset_user2()
                    con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}")
        con.close()

###############################################################################

#### Row 3
######## BUTTONS 

left_column_page5_row3 = ttk.Frame(left_column_page5)
left_column_page5_row3.pack(fill = 'x' , padx = [0,0] , pady = [0,10])

btn_add_page5 = ttk.Button(left_column_page5_row3, text = 'Save' , width = 12 , command=add_user)
btn_add_page5.pack(side = 'left' , padx = 10)

btn_clear_page5 = ttk.Button(left_column_page5_row3 , text = 'Reset' , width = 12 , command=reset_user)
btn_clear_page5.pack(side = 'left' , padx  = 10)

btn_update_page5 = ttk.Button(left_column_page5_row3 , text = 'Update' , width = 12 , command=update_user)
btn_update_page5.pack(side = 'left' , padx = 10)

btn_delete_page5 = ttk.Button(left_column_page5_row3 , text = 'Delete' , width = 12 , command=delete_user)
btn_delete_page5.pack(side = 'left' , padx = 10)

btn_sort_page5 = ttk.Button(left_column_page5_row3 , text = 'Sort' , width = 12 , command=sort_user)
btn_sort_page5.pack(side = 'left' , padx  = 10)

btn_export_page5 = ttk.Button(left_column_page5_row3 , text = 'Export' , width = 12 , command=export_user)
btn_export_page5.pack(side = 'left' , padx  = 10)

txt_search_page5= ttk.Entry(left_column_page5_row3)
txt_search_page5.pack(side='left',padx=10,fill='x',expand=1)

txt_search_page5.bind('<KeyPress>',search_user)

def Size(event):
    print(
    root.winfo_height()
    ,root.winfo_width()
    )

root.bind('<Alt-Return>',Size)
btn_update_page5.configure(state='disabled')
btn_delete_page5.configure(state='disabled')
btn_sort_page5.configure(state='disabled')

#################################################################################################

########### PAGE 6 ############################
page6_Frame = ttk.LabelFrame(page6 , text = 'Sessions')
page6_Frame.pack(fill = BOTH, expand=1)

left_column_page6_row2 = ttk.Frame(page6_Frame)
left_column_page6_row2.pack(expand=1,fill = BOTH,padx=10)

#### Row 1

scrollx_page6 = ttk.Scrollbar(left_column_page6_row2 , orient='horizontal')
scrolly_page6 = ttk.Scrollbar(left_column_page6_row2 , orient='vertical')
scrolly_page6.pack(side = 'right',fill='y')
scrollx_page6.pack(side = 'bottom',fill='x')

class_columns_page6 = ('sr', 'email' , 'pass' , 'role' , 'date' , 'time' , 'rem') 

class_tree_page6 = ttk.Treeview(left_column_page6_row2 , columns=class_columns_page6,
                xscrollcommand=scrollx_page6.set,yscrollcommand=scrolly_page6.set)
class_tree_page6.pack(fill=BOTH,expand=1)

scrollx_page6.config(command=class_tree_page6.xview)
scrolly_page6.config(command=class_tree_page6.yview)

class_tree_page6.column("#0" , width = 0 , stretch=NO)

class_tree_page6.heading("sr" , text = 'Sr.' , anchor='w')
class_tree_page6.heading("email" , text = 'Email' , anchor='w')
class_tree_page6.heading("pass" , text = 'Password' , anchor='w')
class_tree_page6.heading("role" , text = 'Role',anchor='w')
class_tree_page6.heading("date" , text = 'Date' , anchor='w')
class_tree_page6.heading("time" , text = 'Time' , anchor='w')
class_tree_page6.heading("rem" , text = 'Keep Logged In' , anchor='w')

class_tree_page6.column('sr' , width=200 , anchor='w')
class_tree_page6.column('email' , width=200 , anchor='w')
class_tree_page6.column('pass' , width=200 , anchor='w')
class_tree_page6.column('role' , width=200 , anchor='w')
class_tree_page6.column('date' , width=200 , anchor='w')
class_tree_page6.column('time' , width=200 , anchor='w')
class_tree_page6.column('rem' , width=200 , anchor='w')
class_tree_page6["displaycolumns"]=('sr', 'email' , 'role' , 'date' , 'time' , 'rem') 
####################################### FUNCTIONS ##############################################
def show_sessions(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,email text,pass text,role text,date text,time text,rem text)")
        con.commit()
        cur.execute("select * from SESSIONS")
        rows=cur.fetchall()
        class_tree_page6.delete(*class_tree_page6.get_children())
        for row in rows:
            class_tree_page6.insert('',END,values=row)
        con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}. 0x15",parent=root)
        con.close()
show_sessions()
###############################################################################

########### PAGE 7 ############################
page7_Frame = ttk.LabelFrame(page7 , text = 'Results Management')
page7_Frame.pack(expand=1, fill='both')

left_column_page7 = ttk.Frame(page7_Frame)
left_column_page7.pack(side='left',fill='both')

left_column_page7_row1 = ttk.Frame(left_column_page7)
left_column_page7_row1.pack(side = 'top' , fill = 'x',padx=10)

lbl_test_page7 = ttk.Label(left_column_page7_row1 , text = 'Test Code : ')
lbl_test_page7.pack(side = 'left' , padx=[10,26] , pady=10)

com_test_page7 = ttk.Combobox(left_column_page7_row1 , values=() , width=18,state='readonly')
com_test_page7.pack(side='left' , padx=15 , pady =10)

lbl_conf_page7 = ttk.Label(left_column_page7_row1,text='Manage Test : ')
lbl_conf_page7.pack(side='left' , padx=10 , pady=10)

def show_test(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
        con.commit()
        cur.execute("select * from TESTS")
        rows=cur.fetchall()
        if rows!=None:
            tests.clear()
            for row in rows:
                tests.append(row[1])
            com_test_page7.configure(values=tests)
            con.close()
    except Exception as ex:
        MessageBox("Error",f'Erro due to {str(ex)}')


def get_rollnumbers(event=None,clear=True):
    btn_sort_page7.configure(state='active')
    com_rollno_page7.configure(values=())
    if clear:
        com_rollno_page7.set("")
    try:
        con2 = sqlite3.connect(database=DATABASE)
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
        con.commit()
        cur2 = con2.cursor()
        cur2.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text,rollno text)")
        con2.commit()
        cur.execute("select * from TESTS where test_name=?",(com_test_page7.get().strip(),))
        rows=cur.fetchone()
        if rows!=None:
            cur2.execute("select * from STUDENTS where course_code=?",(rows[3],))
            rows2 = cur2.fetchall()
            if rows2!=None:
                rollnumbers.clear()
                for row in rows2:
                    rollnumbers.append(row[3])
                com_rollno_page7.configure(values=rollnumbers)
        con.close()
    except Exception as ex:
        MessageBox("Error",f'Erro due to {str(ex)}')

com_test_page7.bind("<<ComboboxSelected>>",get_rollnumbers)


show_test()

def test_topup(event=None):
    test_root = Toplevel(root)
    test_root.minsize(940,600)

    test_root.geometry("940x600+50+80")
    test_root.focus()

    main_frame  = ttk.LabelFrame(test_root , text='Test Configurations')
    main_frame.pack(expand=1,fill='both')

    first_frame = ttk.Frame(main_frame)
    first_frame.pack(side='top',fill='x')

    lbl_test_name = ttk.Label(first_frame , text= 'Test Name :')
    lbl_test_name.grid(row=0 , column = 0 , padx=10 , pady=10)

    txt_test_name = ttk.Entry(first_frame , textvariable=dname_test)
    txt_test_name.grid(row=0 , column = 1 , padx=10 , pady=10 , sticky='ew')

    lbl_test_name_desc = ttk.Label(first_frame , text= 'Description :')
    lbl_test_name_desc.grid(row=0 , column = 2 , padx=10 , pady=10)

    txt_test_name_desc = ttk.Entry(first_frame , textvariable=dname_test_desc)
    txt_test_name_desc.grid(row=0 , column = 3 , padx=10 , pady=10 , sticky='ew',columnspan=3)

    lbl_date = ttk.Label(first_frame , text= 'Date :')
    lbl_date.grid(row=1 , column = 0 , padx=10 , pady=10)


    ent_date = ttk.DateEntry(first_frame , width=18)

    #ent_date = ttk.DateEntry(first_frame , width=18,state='readonly' , textvariable=ddate_test)
    #ent_date._set_text(ent_date._date.strftime('%m/%d/%Y'))
    ent_date.grid(row=1 , column = 1 , padx=10 , pady=10 , sticky='ew')

    color = style.colors.primary

    #ent_date.config(background=color,selectbackground=color,
    #        headersforeground='white',headersbackground=color,disabledbackground='white')

    lbl_class = ttk.Label(first_frame, text='Class Code  :   ')
    lbl_class.grid(row=1 , column = 2 , padx=10 , pady=10)

    com_class_test = ttk.Combobox(first_frame , values=classes, width=18,state='readonly')
    com_class_test.grid(row=1 , column = 3 , padx=10 , pady=10 , sticky='ew')

    lbl_lecture = ttk.Label(first_frame, text='Lecture : ')
    lbl_lecture.grid(row=1 , column = 4 , padx=10 , pady=10)

    com_lecture_test = ttk.Combobox(first_frame , values=lectures, width=18,state='readonly')
    com_lecture_test.grid(row=1 , column = 5 , padx=10 , pady=10 , sticky='ew')

    def get_lectures_for_test(event=None):
        try:
            com_class_test.config(values=classes)
            con = sqlite3.connect(database=DATABASE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text,subjects text)")
            con.commit()
            cur.execute("select * from CLASSES where course_code=?",(com_class_test.get().strip(),))
            rows=cur.fetchall()
            if rows==None:
                pass
            else:
                rows = make_tuples_to_list(rows)
                com_lecture_test.config(values=rows[0][3].split(','))
        except Exception as ex:
            #print('Error due to ',ex,' 0x13')
            com_lecture_test.config(values=())
            com_lecture_test.set("")

    com_class_test.bind("<<ComboboxSelected>>",get_lectures_for_test)

    lbl_total_marks = ttk.Label(first_frame,text='Total Marks : ')
    lbl_total_marks.grid(row=2 , column = 0 , padx=10 , pady=10)

    txt_total_marks = ttk.Entry(first_frame , textvariable=dtotal_marks_test)
    txt_total_marks.grid(row=2 , column = 1 , padx=10 , pady=10 , sticky='ew')

    lbl_passing_marks = ttk.Label(first_frame,text='Passing Marks : ')
    lbl_passing_marks.grid(row=2 , column = 2 , padx=10 , pady=10)

    txt_passing_marks = ttk.Entry(first_frame,textvariable=dpassing_marks_test)
    txt_passing_marks.grid(row=2 , column = 3 , padx=10 , pady=10 , sticky='ew')

    first_frame.columnconfigure((0,1,2,3,4,5) , weight=1)

    second_frame = ttk.Frame(main_frame)
    second_frame.pack(expand=1,fill='both')

    #### Row 1

    scrollx_test = ttk.Scrollbar(second_frame , orient='horizontal')
    scrolly_test = ttk.Scrollbar(second_frame , orient='vertical')
    scrolly_test.pack(side = 'right',fill='y')
    scrollx_test.pack(side = 'bottom',fill='x')

    class_columns_test = ('sr','test','test_desc' , 'code','subject','date','total_marks','passing_marks') 

    class_tree_test = ttk.Treeview(second_frame , columns=class_columns_test,
                    xscrollcommand=scrollx_test.set,yscrollcommand=scrolly_test.set)
    class_tree_test.pack(fill=BOTH,expand=1)

    scrollx_test.config(command=class_tree_test.xview)
    scrolly_test.config(command=class_tree_test.yview)

    class_tree_test.column("#0" , width = 0 , stretch=NO)

    class_tree_test.heading("sr" , text = 'Sr.' , anchor='w')
    class_tree_test.heading("test" , text = 'Test Title' , anchor='w')
    class_tree_test.heading("test_desc" , text = 'Test' , anchor='w')
    class_tree_test.heading("code" , text = 'Class' , anchor='w')
    class_tree_test.heading("subject" , text = 'Lecture' , anchor='w')
    class_tree_test.heading("date" , text = 'Date' , anchor='w')
    class_tree_test.heading("total_marks" , text = 'Total Marks' , anchor='w')
    class_tree_test.heading("passing_marks" , text = 'Passing Marks' , anchor='w')

    class_tree_test.column('sr' , width=200 , anchor='w')
    class_tree_test.column('test' , width=200 , anchor='w')
    class_tree_test.column('test_desc' , width=200 , anchor='w')
    class_tree_test.column('code' , width=200 , anchor='w')
    class_tree_test.column('subject' , width=200 , anchor='w')
    class_tree_test.column('date' , width=200 , anchor='w')
    class_tree_test.column('total_marks' , width=200 , anchor='w')
    class_tree_test.column('passing_marks' , width=200 , anchor='w')
    ###############################################################################

    def show_test(event=None):
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
            con.commit()
            cur.execute("select * from TESTS")
            rows=cur.fetchall()
            if rows!=None:
                tests.clear()
                class_tree_test.delete(*class_tree_test.get_children())
                for row in rows:
                    class_tree_test.insert('',END,values=row)
                    tests.append(row[1])
                com_test_page7.configure(values=tests)
                con.close()
                if len(rows)>=1:
                    btn_export_test.configure(state='active')
                else:
                    btn_export_test.configure(state='disabled')
        except Exception as ex:
            MessageBox("Error",f'Erro due to {str(ex)}')

    def search_test(event=None):
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
            con.commit()
            cur.execute(f"select * from TESTS where test_name LIKE '%{txt_search_test.get()}%'")
            rows=cur.fetchall()
            if rows!=None:
                class_tree_test.delete(*class_tree_test.get_children())
                for row in rows:
                    class_tree_test.insert('',END,values=row)
                con.close()
        except Exception as ex:
            MessageBox("Error",f'Erro due to {str(ex)} x 17')


    def export_test(event=None):
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
            con.commit()
            data = class_tree_test.selection()
            if len(data)>0:
                rows = [class_tree_test.item(select_item)['values'] for select_item in data]
            else:
                cur.execute("select * from TESTS")
                rows=cur.fetchall()
            file = filedialog.asksaveasfilename(defaultextension='.xlsx',filetype=[('Excel File','.xlsx')],parent=test_root)
            if file =="":
                return
            else:
                rows = pd.DataFrame(rows)
                rows.columns = ['Sr.','Test','Description','Class','Subject','Date','Total Marks','Passing Marks']
                os.system(f'del {file}')
                # writing to Excel
                datatoexcel = pd.ExcelWriter(file)
                
                # write DataFrame to excel
                rows.to_excel(datatoexcel)
                
                # save the excel
                a = datatoexcel.save()
                #rows.to_excel(FILE_EXPORT_CLASS)
                MessageBox("Success",f"Users(s) has been exported !",parent=root)
        except Exception as ex:
            MessageBox("Error",f'Erro due to {str(ex)}')

    ###############################################################################################
    def getdata_test(event=None):
        r=class_tree_test.focus()
        content=class_tree_test.item(r)
        row=content["values"]
        dname_test.set(row[1])
        dname_test_desc.set(row[2])
        ddate_test.set(row[5])
        com_class_test.set(row[3])
        com_lecture_test.set(row[4])
        dtotal_marks_test.set(row[6])
        dpassing_marks_test.set(row[7])
        txt_test_name.configure(state='disabled')
        txt_test_name_desc.configure(state='disabled')
        btn_add_test.configure(state='disabled')
        btn_delete_test.configure(state='active')
        if not len(class_tree_test.selection())>1:
            btn_update_test.configure(state='active')
        else:
            btn_update_test.configure(state='disabled')
        get_lectures_for_test()

    def reset_test(event=None):
        test_root.title(TITLE)
        dname_test.set("")
        dname_test_desc.set("")
        ddate_test.set("")
        com_class_test.set("")
        com_lecture_test.set("")
        dtotal_marks_test.set("")
        dpassing_marks_test.set("")
        btn_add_test.configure(state='active')
        btn_update_test.configure(state='disabled')
        txt_test_name.configure(state='enabled')
        txt_test_name_desc.configure(state='enabled')
        btn_delete_test.configure(state='disabled')
        com_lecture_test.configure(values=())
        txt_search_test.delete(0,END)
        show_test()

    class_tree_test.bind("<ButtonRelease-1>",getdata_test)
    ####################################### FUNCTIONS ##############################################
    def update_data(event=None):
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE,check_same_thread=False)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
            con.commit()
            cur.execute("select * from TESTS where test_name=?",(dname_test.get().strip(),))
            row=cur.fetchone()
            if row!=None:
                con2 = sqlite3.connect(database=DATABASEATTENDNACE,check_same_thread=False)
                cur2 = con2.cursor()
                cur2.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
                con2.commit()
                cur2.execute("select * from RESULTS where test_title=?",(dname_test.get().strip(),))
                rows = cur2.fetchall()
                if rows!=None:
                    def progress():
                        third_frame.pack_forget() #WAKAO
                        frame = ttk.Frame(main_frame)
                        frame.pack(fill = 'x' , padx = [10,10] , pady = [10,10])
                        pg = ttk.Progressbar(frame,maximum=len(rows),style='primary.Striped.Horizontal.TProgressbar')
                        pg.grid(row=0,column=0,sticky='ew',columnspan=2)
                        lbl = ttk.Label(frame, text='',font=font_lbl,style='primaary.TLabel')
                        lbl.grid(row=1,column=0)
                        frame.columnconfigure(1,weight=1)
                        for index,row2 in enumerate(rows):
                            percent = round(index/len(rows)*100,2)
                            total_marks = int(row2[8])
                            obtained_marks = int(row2[9])
                            passing_marks = int(dpassing_marks_test.get().strip())
                            rollno = row2[4]
                            Class = row2[5]
                            status=""
                            if obtained_marks>=passing_marks:
                                status='PASS'
                            else:
                                status="FAIL"
                            percentage = f'{str(round(obtained_marks/int(dtotal_marks_test.get().strip())*100,3))}%'
                            cur2.execute('update RESULTS set course_code=? , subjects=? , total_marks = ? , passing_marks=? , date=? , grade=? , status=? where test_title=? AND course_code=? AND rollno=?',(
                                com_class_test.get().strip(),
                                com_lecture_test.get().strip(),
                                dtotal_marks_test.get().strip(),
                                dpassing_marks_test.get().strip(),
                                ent_date.entry.get().strip(),
                                percentage.strip(),
                                status.strip(),
                                dname_test.get().strip(),
                                Class.strip(),
                                rollno.strip()
                            ))
                            con2.commit()
                            pg['value'] += 1
                            lbl.configure(text=f'Please Do not click any widget of this page, Press Reset Button after its done !\nCompleted : {percent} %')
                            test_root.title(f"Updating Results : {percent} %")
                        lbl.configure(text=f'Please wait till the process finishes')
                        root.title(f"Reset Needed of Result Page")
                        lbl.destroy()
                        pg.destroy()
                        frame.destroy()
                        third_frame.pack(fill = 'x' , padx = [0,0] , pady = [10,10])
                        test_root.title('Completed : 100 %')
                        MessageBox("Success","All records have been updated, Reset the Table of Result Page !")
                        test_root.title("Reset Needed ! ")
                        con.close()
                    a=threading.Thread(target=progress).start()
        except Exception as ex:
            MessageBox("Error",f'Erro due to {str(ex)}')

    def update_test(event=None):
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
            con.commit()
            cur.execute("select * from TESTS where test_name=?",(dname_test.get().strip(),))
            row=cur.fetchone()
            if dname_test.get().strip() == '' or dname_test.get().strip() == 'None':
                MessageBox("Error","Select recrod from table !")
            else:
                if row==None:
                    MessageBox("Error","No Such record !")
                else:
                    selected_items = class_tree_test.selection()
                    if len(selected_items) == 1:
                        for select_item in selected_items:
                            current_val = class_tree_test.item(select_item)['values']
                            cur.execute("update TESTS set date=? , course_code=? , subjects=? , total_marks=? , passing_marks=? where cid=?",(
                                ent_date.entry.get().strip(),
                                com_class_test.get().strip(),
                                com_lecture_test.get().strip(),
                                dtotal_marks_test.get().strip(),
                                dpassing_marks_test.get().strip(),
                                str(current_val[0])
                            ))
                            con.commit()
                        update_data()
                    else:
                        MessageBox('Information',"Currently Not supported In next updates it will be covered !")
        except Exception as ex:
            MessageBox("Error",f'Erro due to {str(ex)}')

    def delete_test(event=None):
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
            con.commit()
            cur.execute("select * from TESTS where test_name=?",(dname_test.get().strip(),))
            row=cur.fetchone()
            if dname_test.get().strip() == '' or dname_test.get().strip() == 'None':
                MessageBox("Error","Select recrod from table !")
            else:
                if row==None:
                    MessageBox("Error","No Such record !")
                else:
                    op = MessageBox("Confirm","Do you really want to delete Test(s)? ",b1='Yes',b2='No',parent=root)
                    if op.choice=='Yes':
                        selected_items = class_tree_test.selection()
                        for select_item in selected_items:
                            current_val = class_tree_test.item(select_item)['values']
                            cur.execute("delete from TESTS where cid=?",(current_val[0],))
                            con.commit()
                        MessageBox("Sucess","Test(s) has been deleted !")
                        reset_test()
        except Exception as ex:
            MessageBox("Error",f'Erro due to {str(ex)}')

    def add_test(event=None):
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS TESTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_name text,test_desc text,course_code text,subjects text,date text,total_marks text,passing_marks text)")
            con.commit()
            cur.execute("select * from TESTS where test_name=?",(dname_test.get().strip(),))
            row=cur.fetchone()
            if row!=None:
                MessageBox("Error","Test already exists !")
            else:
                for _ in range(1):
                    if dname_test.get().strip()=='' or dname_test.get().strip()=='None':
                        MessageBox("Error","Enter Name/Title for Test !")
                        txt_test_name.focus()
                        break
                    if dname_test_desc.get().strip()=='' or dname_test_desc.get().strip()=='None':
                        MessageBox("Error","Enter Description !")
                        txt_test_name_desc.focus()
                        break
                    if ent_date.entry.get().strip()=='' or ent_date.entry.get().strip()=='None':
                        MessageBox("Error","Select Date !")
                        ent_date.focus()
                        break
                    if com_class_test.get().strip()=='' or com_class_test.get().strip()=='None':
                        MessageBox("Error","Select class !")
                        com_class_test.focus()
                        break
                    if com_lecture_test.get().strip()=='' or com_lecture_test.get().strip()=='None':
                        MessageBox("Error","Select lecture for class !")
                        com_lecture_test.focus()
                        break
                    if dtotal_marks_test.get().strip()=='' or dtotal_marks_test.get().strip()=='None':
                        MessageBox("Error","Enter Total Marks !")
                        txt_total_marks.focus()
                        break
                    try:
                        int(dtotal_marks_test.get().strip())
                    except:
                        MessageBox("Error","Only Numerics are allowed !")
                        txt_total_marks.focus()
                        break
                    if dpassing_marks_test.get().strip()=='' or dpassing_marks_test.get().strip()=='None':
                        MessageBox("Error","Enter Passing Marks !")
                        txt_passing_marks.focus()
                        break
                    try:
                        int(dpassing_marks_test.get().strip())
                    except:
                        MessageBox("Error","Only Numerics are allowed !")
                        txt_passing_marks.focus()
                        break
                    if int(dpassing_marks_test.get().strip())>=int(dtotal_marks_test.get().strip()):
                        MessageBox("Error","Passing Marks cannot be more than total marks ! ")
                        txt_passing_marks.focus()
                        break
                    else:
                        cur.execute("Insert into TESTS (test_name,test_desc,course_code,subjects,date,total_marks,passing_marks) values(?,?,?,?,?,?,?)",(
                            dname_test.get().strip(),
                            dname_test_desc.get().strip(),
                            com_class_test.get().strip(),
                            com_lecture_test.get().strip(),
                            ent_date.entry.get().strip(),
                            dtotal_marks_test.get().strip(),
                            dpassing_marks_test.get().strip()
                        ))
                        con.commit()
                        MessageBox("Success",'Test has been added !')
                        reset_test()
        except Exception as ex:
            MessageBox("Error",f'Erro due to {str(ex)}')

    ###############################################################################

    #### Row 3
    ######## BUTTONS 
    third_frame = ttk.Frame(main_frame)
    third_frame.pack(fill = 'x' , padx = [0,0] , pady = [10,10])

    btn_add_test = ttk.Button(third_frame, text = 'Save' , width = 12,command=add_test)
    btn_add_test.pack(side = 'left' , padx = 10)

    btn_clear_test = ttk.Button(third_frame , text = 'Reset' , width = 12,command=reset_test)
    btn_clear_test.pack(side = 'left' , padx  = 10)

    btn_update_test = ttk.Button(third_frame , text = 'Update' , width = 12,command=update_test)
    btn_update_test.pack(side = 'left' , padx = 10)

    btn_delete_test = ttk.Button(third_frame , text = 'Delete' , width = 12,command=delete_test)
    btn_delete_test.pack(side = 'left' , padx = 10)

    btn_sort_test = ttk.Button(third_frame , text = 'Sort' , width = 12,state='disabled')
    btn_sort_test.pack(side = 'left' , padx = 10)

    btn_export_test = ttk.Button(third_frame , text = 'Export' , width = 12,command=export_test)
    btn_export_test.pack(side = 'left' , padx = 10)

    txt_search_test = ttk.Entry(third_frame)
    txt_search_test.pack(side='left',padx=10,pady=10,fill='x',expand=1)

    txt_search_test.bind("<KeyPress>",search_test)

    btn_update_test.configure(state='disabled')
    btn_delete_test.configure(state='disabled')

    show_test()

    def quit_me(e=None):
        op = MessageBox('Confirm',"Do you really wish to close this tab?",b1='Yes',b2='No')
        if op.choice=='Yes':
            test_root.destroy()

    test_root.bind("<Return>",add_test)
    test_root.bind("<Control-u>",update_test)
    test_root.bind("<Control-d>",delete_test)
    test_root.bind("<Control-r>",reset_test)
    test_root.bind("<Control-s>",export_test)

    test_root.protocol("WM_DELETE_WINDOW", quit_me)
    test_root.wait_window()

btn_conf_page7 = ttk.Button(left_column_page7_row1 , text ='Configure' , width=18 , style='Outline.TButton' , command=test_topup)
btn_conf_page7.pack(side = 'left' , padx =5 , pady = 10)

left_column_page7_row2 = ttk.Frame(left_column_page7)
left_column_page7_row2.pack(side = 'top' , fill = 'x',padx=12)

lbl_rollno_page7 = ttk.Label(left_column_page7_row2 , text = "Roll No :" )
lbl_rollno_page7.pack(side = 'left' ,padx=10, pady = 10)

com_rollno_page7_values = ()

com_rollno_page7 = ttk.Combobox(left_column_page7_row2 , width = 18 , values=com_rollno_page7_values,state='readonly')
com_rollno_page7.pack(side = 'left' , padx = [45,10] , pady = 10)

def get_test_info(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("select * from TESTS where test_name=?",(com_test_page7.get().strip(),))
        row = cur.fetchone()
        if row!=None:
            con.close()
            CreateToolTip(com_test_page7,row)
    except Exception as ex:
        MessageBox("Error",f'Erro due to {str(ex)}')

com_test_page7.bind("<Enter>",get_test_info)


def get_student_name(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("select * from TESTS where test_name=?",(com_test_page7.get().strip(),))
        row = cur.fetchone()
        if row!=None:
            degree =row[3]
            con2 = sqlite3.connect(database=DATABASE)
            cur2 = con2.cursor()
            cur2.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text,rollno text)")
            con2.commit()
            cur2.execute("select * from STUDENTS where rollno=? AND course_code=?",(com_rollno_page7.get().strip(),degree,))
            rows2 = cur2.fetchone()
            con.close()
            CreateToolTip(com_rollno_page7,rows2)
    except Exception as ex:
        MessageBox("Error",f'Erro due to {str(ex)}')

com_rollno_page7.bind("<Enter>",get_student_name)

lbl_obtained_marks_page7 = ttk.Label(left_column_page7_row2 , text = "Obtained Marks :" )
lbl_obtained_marks_page7.pack(side = 'left' , padx = [10,0] , pady = 10)

txt_obtained_marks_page7 = ttk.Entry(left_column_page7_row2)
txt_obtained_marks_page7.pack(side = 'left' , padx =5 , pady =10)

left_column_page7_row3 = ttk.Frame(left_column_page7)
left_column_page7_row3.pack(expand=1,fill = BOTH,pady=10,padx=10)

#### Row 1

scrollx_page7 = ttk.Scrollbar(left_column_page7_row3 , orient='horizontal')
scrolly_page7 = ttk.Scrollbar(left_column_page7_row3 , orient='vertical')
scrolly_page7.pack(side = 'right',fill='y')
scrollx_page7.pack(side = 'bottom',fill='x')

class_columns_page7 = ('sr' ,'testname', 'date' ,'student_name' ,'rollno' ,  'code' , 'lecture' ,'grade', 'total_marks', 'obtained_marks', 'passing_marks','status') 

class_tree_page7 = ttk.Treeview(left_column_page7_row3 , columns=class_columns_page7,
                xscrollcommand=scrollx_page7.set,yscrollcommand=scrolly_page7.set)
class_tree_page7.pack(fill='both',expand=1)

scrollx_page7.config(command=class_tree_page7.xview)
scrolly_page7.config(command=class_tree_page7.yview)

class_tree_page7.column("#0" , width = 0 , stretch=NO)

class_tree_page7.heading("sr" , text = 'Sr.' , anchor='w')
class_tree_page7.heading("student_name" , text = 'Student Name' , anchor='w')
class_tree_page7.heading("rollno" , text = 'Roll No' , anchor='w')
class_tree_page7.heading("code" , text = 'Class' , anchor='w')
class_tree_page7.heading("lecture" , text = 'Subject' , anchor='w')
class_tree_page7.heading("testname" , text = 'Test Title' , anchor='w')
class_tree_page7.heading("date" , text = 'Test Date' , anchor='w')
class_tree_page7.heading("total_marks" , text = 'Total Marks' , anchor='w')
class_tree_page7.heading("obtained_marks" , text = 'Obtained Marks' , anchor='w')
class_tree_page7.heading("passing_marks" , text = 'Passing Marks' , anchor='w')
class_tree_page7.heading("status" , text = 'Status' , anchor='w')
class_tree_page7.heading("grade" , text = 'Percentage' , anchor='w')

class_tree_page7.column('sr' , width=200 , anchor='w')
class_tree_page7.column('student_name' , width=200 , anchor='w')
class_tree_page7.column('rollno' , width=200 , anchor='w')
class_tree_page7.column('code' , width=200 , anchor='w')
class_tree_page7.column('lecture' , width=200 , anchor='w')
class_tree_page7.column('testname' , width=200 , anchor='w')
class_tree_page7.column('date' , width=200 , anchor='w')
class_tree_page7.column('total_marks' , width=200 , anchor='w')
class_tree_page7.column('obtained_marks' , width=200 , anchor='w')
class_tree_page7.column('passing_marks' , width=200 , anchor='w')
class_tree_page7.column('status' , width=200 , anchor='w')
class_tree_page7.column('grade' , width=200 , anchor='w')

class_tree_page7["displaycolumns"]=('sr', 'student_name' , 'rollno', 'testname','obtained_marks','status') 

##############################################################################################
def show_result(event=None):
    try:
        root.title(TITLE)
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
        con.commit()
        cur.execute("select * from RESULTS")
        rows = cur.fetchall()
        if rows!=None:
            class_tree_page7.delete(*class_tree_page7.get_children())
            for row in rows:
                class_tree_page7.insert('',END,values=row)
            if len(rows)>=1:
                btn_export_page7.configure(state='active')
            else:
                btn_export_page7.configure(state='disabled')
    except Exception as ex:
        MessageBox("Error",f'Error due to {str(ex)}')

def show_result2(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
        con.commit()
        cur.execute("select * from RESULTS where test_title=? AND rollno=?",(com_test_page7.get().strip(),com_rollno_page7.get().strip(),))
        row = cur.fetchone()
        if row!=None:
            class_tree_page7.insert('',END,values=row)
    except Exception as ex:
        MessageBox("Error",f'Error due to {str(ex)}')

def search_result(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE,check_same_thread=False)
        cur = con.cursor()
        cur.execute(f"select * from RESULTS where student_name LIKE '%{txt_search_page7.get()}%'")
        rows=cur.fetchall()
        if rows!=None and len(rows)!=0:
            class_tree_page7.delete(*class_tree_page7.get_children())
            for row in rows:
                class_tree_page7.insert('',END,values=row)
            con.close()
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)} 0x16")


def sort_result(event=None):
    if com_test_page7.get().strip()=="":
        MessageBox("Information","Select Test first !")
        com_test_page7.focus()
    else:
        try:
            con = sqlite3.connect(database=DATABASEATTENDNACE)
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
            con.commit()
            cur.execute("select * from RESULTS where test_title=?",(com_test_page7.get().strip(),))
            rows = cur.fetchall()
            if rows!=None and len(rows)!=0:
                class_tree_page7.delete(*class_tree_page7.get_children())
                for row in rows:
                    class_tree_page7.insert('',END,values=row)
            else:
                MessageBox("Information","There are no students having this test !")
        except Exception as ex:
            MessageBox("Error",f'Error due to {str(ex)}')

##############################################################################################
def export_result(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
        con.commit()
        data = class_tree_page7.selection()
        if len(data)>0:
            rows = [class_tree_page7.item(select_item)['values'] for select_item in data]
        else:
            cur.execute("select * from RESULTS")
            rows = cur.fetchall()
        file = filedialog.asksaveasfilename(defaultextension='.xlsx',filetype=[('Excel File','.xlsx')])
        if file =="":
            return
        else:
            rows = pd.DataFrame(rows)
            rows.columns = ['Sr.','Test','Date','Student Name','Roll No.','Class','Subject','Percentage %','Total Marks','Obtained Marks','Passing Marks','Status']
            os.system(f'del {file}')
            # writing to Excel
            datatoexcel = pd.ExcelWriter(file)
            
            # write DataFrame to excel
            rows.to_excel(datatoexcel)
            
            # save the excel
            a = datatoexcel.save()
            #rows.to_excel(FILE_EXPORT_CLASS)
            MessageBox("Success",f"Result(s) has been exported !",parent=root)

    except Exception as ex:
        MessageBox("Error",f'Error due to {str(ex)}')

###############################################################################################
def getdata_result(event=None):
    btn_sort_page7.configure(state='active')
    r=class_tree_page7.focus()
    content=class_tree_page7.item(r)
    row=content["values"]
    com_test_page7.set(row[1])
    com_rollno_page7.set(row[4])
    txt_obtained_marks_page7.delete(0,END)
    txt_obtained_marks_page7.insert(0,row[9])
    btn_update_page7.configure(state='active')
    btn_add_page7.configure(state='disabled')
    com_rollno_page7.configure(state='disabled')
    btn_delete_page7.configure(state='active')
    txt_course_code.configure(state='disabled')
    get_rollnumbers(clear=False)

def reset_result(event=None):
    btn_sort_page7.configure(state='disabled')
    btn_add_page7.configure(state='active')
    com_rollno_page7.configure(state='readonly')
    btn_update_page7.configure(state='disabled')
    btn_delete_page7.configure(state='disabled')
    txt_search_page7.delete(0,END)
    com_test_page7.set("")
    com_rollno_page7.set("")
    com_rollno_page7.configure(values="")
    txt_obtained_marks_page7.delete(0,END)
    show_result()

def reset_result3(event=None):
    btn_add_page7.configure(state='disabled')
    btn_update_page7.configure(state='disabled')
    btn_delete_page7.configure(state='disabled')
    txt_obtained_marks_page7.delete(0,END)
    show_result()
    txt_obtained_marks_page7.focus()


def reset_result2(event=None):
    btn_add_page7.configure(state='active')
    btn_update_page7.configure(state='disabled')
    btn_delete_page7.configure(state='disabled')
    next = rollnumbers.index(com_rollno_page7.get().strip())
    show_result2()
    if rollnumbers[next]==rollnumbers[-1]:
        next = 0
        com_rollno_page7.set(rollnumbers[next])
    else:
        com_rollno_page7.set(rollnumbers[next+1])
    txt_obtained_marks_page7.delete(0,END)
    txt_obtained_marks_page7.focus()

class_tree_page7.bind("<ButtonRelease-1>",getdata_result)
####################################### FUNCTIONS ##############################################

def update_result(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE,check_same_thread=False)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
        con.commit()
        #cur.execute("select * from RESULTS where rollno=?",(com_test_page7.get().strip(),))
        #row2 = cur.fetchone()
        if com_test_page7.get().strip()=="":
            MessageBox("Error","No Such record !")
        else:
            for _ in range(1):
                if com_test_page7.get().strip() == '':
                    MessageBox("Error","Select Test first !")
                    com_test_page7.focus()
                    break
                if txt_obtained_marks_page7.get().strip() == '':
                    MessageBox("Error","Enter Obtained Marks !")
                    txt_obtained_marks_page7.focus()
                    break
                else:
                    op = MessageBox("Confirm","Do you really want to update Result(s)? ",b1='Yes',b2='No',parent=root)
                    if op.choice=='Yes':
                        selected_items = class_tree_page7.selection()
                        if len(selected_items)>1:
                            def progress2():
                                selected_items = class_tree_page7.selection()
                                left_column_page7_row4.pack_forget()
                                frame = ttk.Frame(left_column_page7)
                                frame.pack(side='bottom',fill='x',padx=10,pady=10,anchor='s')
                                pg = ttk.Progressbar(frame,maximum=len(selected_items),style='primary.Striped.Horizontal.TProgressbar')
                                pg.grid(row=0,column=0,sticky='ew',columnspan=2)
                                lbl = ttk.Label(frame, text='',font=font_lbl,style='primaary.TLabel')
                                lbl.grid(row=1,column=0)
                                frame.columnconfigure(1,weight=1)
                                selected_items = class_tree_page7.selection()
                                for index,select_item in enumerate(selected_items):
                                    percent = round(index/len(selected_items)*100,2)
                                    current_val = class_tree_page7.item(select_item)['values']
                                    cur.execute("update RESULTS set test_title=? where cid=?",(
                                        com_test_page7.get().strip(),
                                        str(current_val[0])
                                    ))
                                    con.commit()
                                    pg['value'] += 1
                                    lbl.configure(text=f'Please Do not click any widget of this page, Press Reset Button after its done !\nCompleted : {percent} %')
                                    root.title(f"Updating Results : {percent} %")
                                lbl.configure(text=f'Completed : 100 %')
                                lbl.configure(text=f'Please wait till the process finishes')
                                root.title(f"Reset Needed of Result Page")
                                lbl.destroy()
                                pg.destroy()
                                frame.destroy()
                                left_column_page7_row4.pack(fill = 'x' , padx = [0,0] , pady = 10)
                                MessageBox("Success","Result(s) has been updated ! ",parent=root)
                                con.close()
                            b = threading.Thread(target=progress2).start()
                        else:
                            con = sqlite3.connect(database=DATABASEATTENDNACE)
                            cur = con.cursor()
                            cur.execute("select * from TESTS where test_name=?",(com_test_page7.get().strip(),))
                            row = cur.fetchone()
                            if row!=None:
                                total_marks = row[6]
                                passing_marks = int(row[7])
                            if int(txt_obtained_marks_page7.get().strip())>int(total_marks):
                                MessageBox("Error",f"Obtained Marks Can't be more than total marks ({total_marks})")
                                txt_obtained_marks_page7.focus()
                                break
                            else:
                                status = ""
                                if int(txt_obtained_marks_page7.get().strip())>=passing_marks:
                                    status='PASS'
                                else:
                                    status="FAIL"
                                grade = f'{str(round(int(txt_obtained_marks_page7.get().strip())/int(total_marks)*100,3))}%'
                                for select_item in selected_items:
                                    current_val = class_tree_page7.item(select_item)['values']
                                    cur.execute("update RESULTS set test_title=? , obtained_marks=? ,grade=?, status=? where cid=?",(
                                        com_test_page7.get().strip(),
                                        txt_obtained_marks_page7.get().strip(),
                                        grade.strip(),
                                        status.strip(),
                                        str(current_val[0])
                                    ))
                                    con.commit()
                                reset_result3()
                                MessageBox("Success","Result(s) has been updated ! ",parent=root)
                                con.close()
    except Exception as ex:
        MessageBox("Error",f'Error due to {str(ex)}.')


def delete_result(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE,check_same_thread=False)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
        con.commit()
        cur.execute("select * from RESULTS where test_title=?",(com_test_page7.get().strip(),))
        row=cur.fetchone()
        if row==None:
            MessageBox("Error","Fetch record from Table !",parent=root)
        else:
            op = MessageBox("Confirm","Do you really want to delete Result(s)? ",b1='Yes',b2='No',parent=root)
            if op.choice=='Yes':
                selected_items = class_tree_page7.selection()
                left_column_page7_row4.pack_forget()
                frame = ttk.Frame(left_column_page7)
                frame.pack(side='bottom',fill='x',padx=10,pady=10,anchor='s')
                pg = ttk.Progressbar(frame,maximum=len(selected_items),style='primary.Striped.Horizontal.TProgressbar')
                pg.grid(row=0,column=0,sticky='ew',columnspan=2)
                lbl = ttk.Label(frame, text='',font=font_lbl,style='primaary.TLabel')
                lbl.grid(row=1,column=0)
                frame.columnconfigure(1,weight=1)
        
                def progress():
                    selected_items = class_tree_page7.selection()
                    for index,select_item in enumerate(selected_items):
                        percent = round(index/len(selected_items)*100,2)
                        current_val = class_tree_page7.item(select_item)['values']
                        class_tree_page7.delete(select_item)
                        cur.execute("delete from RESULTS where cid=?",(current_val[0],))
                        con.commit()
                        pg['value'] += 1
                        lbl.configure(text=f'Completed : {percent} %')
                        root.title(f"Deleting Results : {percent} %")
                    lbl.configure(text=f'Completed : 100 %')
                    lbl.configure(text=f'Please wait till the process finishes')
                    root.title("Reset Needed of Result Page !")
                    lbl.destroy()
                    pg.destroy()
                    frame.destroy()
                    left_column_page7_row4.pack(fill = 'x' , padx = [0,0] , pady = 10)
                    MessageBox("Success","Result(s) has been deleted !",parent=root)
                    con.close()
                a = threading.Thread(target=progress).start()
    except Exception as ex:
        MessageBox("Error",f'Error due to {str(ex)}.')

def add_result(event=None):
    try:
        con2 = sqlite3.connect(database=DATABASE)
        cur2 = con2.cursor()
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RESULTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,test_title text , date text ,student_name text ,rollno text , course_code text , subjects text , grade text , total_marks text ,obtained_marks text , passing_marks text,status text)")
        con.commit()
        cur.execute("select * from RESULTS where test_title=? AND rollno=?",(com_test_page7.get().strip(),com_rollno_page7.get().strip(),))
        row2 = cur.fetchone()
        if row2!=None:
            MessageBox("Error","Record already exists !")
        else:
            for _ in range(1):
                if com_test_page7.get().strip() == '':
                    MessageBox("Error","Select Test first !")
                    com_test_page7.focus()
                    break
                if com_rollno_page7.get().strip() == '':
                    MessageBox("Error","Select Roll number first !")
                    com_rollno_page7.focus()
                    break
                if txt_obtained_marks_page7.get().strip() == '':
                    MessageBox("Error","Enter Obtained Marks !")
                    txt_obtained_marks_page7.focus()
                    break
                else:
                    cur.execute("select * from TESTS where test_name=?",(com_test_page7.get().strip(),))
                    row = cur.fetchone()
                    if row!=None:
                        title = row[1]
                        course_code = row[3]
                        subjects = row[4]
                        total_marks = row[6]
                        passing_marks = row[7]
                        date = row[5]
                        status =""
                        if int(txt_obtained_marks_page7.get().strip())>=int(passing_marks):
                            status='PASS'
                        else:
                            status="FAIL"
                        grade = f'{str(round(int(txt_obtained_marks_page7.get().strip())/int(total_marks)*100,3))}%'
                        cur2.execute("select * from STUDENTS where rollno=? AND course_code=?",(com_rollno_page7.get().strip(),course_code.strip(),))
                        student = cur2.fetchone()
                        student = student[2]
                        if int(txt_obtained_marks_page7.get().strip())>int(total_marks):
                            MessageBox("Error",f"Obtained Marks Can't be more than total marks ({total_marks})")
                            txt_obtained_marks_page7.focus()
                            break
                        else:
                            cur.execute("insert into RESULTS(student_name,rollno,course_code,subjects,test_title,grade,obtained_marks,passing_marks,total_marks,date,status) values(?,?,?,?,?,?,?,?,?,?,?)",(
                                student.strip(),
                                com_rollno_page7.get().strip(),
                                course_code.strip(),
                                subjects.strip(),
                                title.strip(),
                                grade.strip(),
                                txt_obtained_marks_page7.get().strip(),
                                passing_marks.strip(),
                                total_marks.strip(),
                                date.strip(),
                                status.strip()
                            ))
                            con.commit()
                            reset_result2()
                            MessageBox('Success',"Record has been added !")
    except Exception as ex:
        MessageBox("Error",f'Error due to {str(ex)}.')

###################################################################################
#### Row 3
######## BUTTONS 

left_column_page7_row4 = ttk.Frame(left_column_page7)
left_column_page7_row4.pack(fill = 'x' , padx = [0,0] , pady = 10)

btn_add_page7 = ttk.Button(left_column_page7_row4, text = 'Save' , width = 12 , command=add_result)
btn_add_page7.pack(side = 'left' , padx = 10)

btn_clear_page7 = ttk.Button(left_column_page7_row4 , text = 'Reset' , width = 12 , command=reset_result)
btn_clear_page7.pack(side = 'left' , padx  = 10)

btn_update_page7 = ttk.Button(left_column_page7_row4 , text = 'Update' , width = 12 , command=update_result)
btn_update_page7.pack(side = 'left' , padx = 10)

btn_delete_page7 = ttk.Button(left_column_page7_row4 , text = 'Delete' , width = 12 , command=delete_result)
btn_delete_page7.pack(side = 'left' , padx = 10)

btn_sort_page7 = ttk.Button(left_column_page7_row4 , text = 'Sort' , width = 12 , command=sort_result)
btn_sort_page7.pack(side = 'left' , padx = 10)

btn_export_page7 = ttk.Button(left_column_page7_row4 , text = 'Export' , width = 12 , command=export_result)
btn_export_page7.pack(side = 'left' , padx = 10)

txt_search_page7 = ttk.Entry(left_column_page7_row4)
txt_search_page7.pack(side='left',fill='x',expand=1,padx=10)

txt_search_page7.bind("<KeyPress>",search_result)
btn_sort_page7.configure(state='disabled')
btn_delete_page7.configure(state='disabled')
btn_update_page7.configure(state='disabled')
#################################################################################################

show_result()
#nb.select(1)
#nb2.select(4)
#################### REPORTS ######################
report_frame = ttk.LabelFrame(page8  , text ='Report')
report_frame.pack(expand=1,fill='both')

report_frame_left_column = ttk.Frame(report_frame)
report_frame_left_column.pack(expand=1,fill='both')

report_frame_left_column_row1 = ttk.Frame(report_frame_left_column)
report_frame_left_column_row1.pack(side='top',fill='x',padx=10)

lbl_class_report = ttk.Label(report_frame_left_column_row1 , text='Class Code :')
lbl_class_report.grid(row=0,column=0)

com_class_report = ttk.Combobox(report_frame_left_column_row1,state='readonly',values=classes , width=18)
com_class_report.grid(row=0,column=1,pady=10 , sticky='ew')

lbl_rollno_report = ttk.Label(report_frame_left_column_row1 , text='Roll No :')
lbl_rollno_report.grid(row=0,column=2)

com_rollno_report = ttk.Combobox(report_frame_left_column_row1,values=rollnumbers,state='readonly' , width=18)
com_rollno_report.grid(row=0,column=3,pady=10 , sticky='ew')

lbl_lecture_report = ttk.Label(report_frame_left_column_row1 , text='Lecture :')
lbl_lecture_report.grid(row=0,column=4)

com_lecture_report = ttk.Combobox(report_frame_left_column_row1,values=() , width=18,state='readonly')
com_lecture_report.grid(row=0,column=5,pady=10 , sticky='ew')

lbl_month_report = ttk.Label(report_frame_left_column_row1 , text='Month :  ')
lbl_month_report.grid(row=1,column=0)

com_month_report = ttk.Combobox(report_frame_left_column_row1,values=() , width=18,state='readonly')
com_month_report.grid(row=1,column=1,pady=10 , sticky='ew')

report_frame_left_column_row1.columnconfigure((0,1,2,3,4,5),weight=1)

###################################################

def get_lectures_for_report(event=None,clear=True):
    try:
        com_lecture_report.config(values=classes)
        con = sqlite3.connect(database=DATABASE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS CLASSES(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,degree text,subjects text)")
        con.commit()
        cur.execute("select * from CLASSES where course_code=?",(com_class_report.get().strip(),))
        rows=cur.fetchall()
        if rows==None:
            pass
        else:
            rows = make_tuples_to_list(rows)
            com_lecture_report.config(values=rows[0][3].split(','))
    except Exception as ex:
        #print('Error due to ',ex,' 0x13')
        com_lecture_report.config(values=())
        com_lecture_report.set("")
dates = []

def convert_to_tree(list_):
    pass
    """
    for i in range(1,31+1):
        if i<=len(list_):
            print(list_[i-1])
    """

def get_details(event=None):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,
            course_code text,
            subjects text,
            date text,
            students text,
            status blob,
            identifier text)
            """)
        # Course , Subject , date , status
        con.commit()
        cur.execute("""
        select * from ATTENDANCE where course_code=? AND subjects=? AND identifier=?
        """,(com_class_report.get().strip(),com_lecture_report.get().strip(),com_month_report.get().strip(),))
        rows=cur.fetchall()
        if rows!=None:
            class_tree_page8.delete(*class_tree_page8.get_children())
            for row in rows:
                class_tree_page8.insert('',END,values=row)
        convert_to_tree(rows)
    except Exception as ex:
        MessageBox("Error",f'Error due to {ex}')

def get_months_report(event=None):
    global dates
    try:
        dates.clear()
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ATTENDANCE(cid INTEGER PRIMARY KEY AUTOINCREMENT,
            course_code text,
            subjects text,
            date text,
            students text,
            status blob,
            identifier text)
            """)
        # Course , Subject , date , status
        con.commit()
        cur.execute("""
        select * from ATTENDANCE where course_code=?
        """,(com_class_report.get().strip(),))
        rows=cur.fetchall()
        if rows!=None:
            for row in rows:
                dates.append(row[3])
            dates_ = []
            for i in dates:
                i = i.split('/')
                popped = i.pop(1)
                i.insert(0,popped)
                dates_.append("/".join(i))
            month_year_parts = [x[x.index('/') + 1:] for x in dates_]
            dates = list(set(month_year_parts))
            com_month_report.configure(values=dates)
        else:
            com_month_report.configure(values=[])
            com_month_report.set("")
    except Exception as ex:
        MessageBox("Error",f'Error due to {ex}')

def get_rollnumbers_report(event=None,clear=True):
    if clear:
        com_month_report.configure(values=[])
        com_rollno_report.set("")
        com_rollno_report.configure(values=[])
        com_lecture_report.set("")
        com_lecture_report.config(values=[])
        com_month_report.set("")
    try:
        con2 = sqlite3.connect(database=DATABASE)
        cur2 = con2.cursor()
        cur2.execute("CREATE TABLE IF NOT EXISTS STUDENTS(cid INTEGER PRIMARY KEY AUTOINCREMENT,course_code text,name text,rollno text)")
        con2.commit()
        cur2.execute("select * from STUDENTS where course_code=?",(com_class_report.get().strip(),))
        rows=cur2.fetchall()
        if rows!=None:
            rollnumbers.clear()
            for row in rows:
                rollnumbers.append(row[3])
            com_rollno_report.configure(values=rollnumbers)
            get_lectures_for_report()
            get_months_report()
        con2.close()
    except Exception as ex:
        MessageBox("Error",f'Erro due to {str(ex)}')
com_class_report.bind("<<ComboboxSelected>>",get_rollnumbers_report)
com_month_report.bind("<<ComboboxSelected>>",get_details)
###################################################
report_frame_left_column_row2 = ttk.Frame(report_frame_left_column)
report_frame_left_column_row2.pack(expand=1,fill = BOTH,pady=10,padx=10)

#### Row 1

scrollx_page8 = ttk.Scrollbar(report_frame_left_column_row2 , orient='horizontal')
scrolly_page8 = ttk.Scrollbar(report_frame_left_column_row2 , orient='vertical')
scrolly_page8.pack(side = 'right',fill='y')
scrollx_page8.pack(side = 'bottom',fill='x')

class_columns_page8 = ('sr' ,'lecture','a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','aa','ab','ac','ad') 

class_tree_page8 = ttk.Treeview(report_frame_left_column_row2 , columns=class_columns_page8,
                xscrollcommand=scrollx_page8.set,yscrollcommand=scrolly_page8.set)
class_tree_page8.pack(fill='both',expand=1)

scrollx_page8.config(command=class_tree_page8.xview)
scrolly_page8.config(command=class_tree_page8.yview)

class_tree_page8.column("#0" , width = 0 , stretch=NO)

class_tree_page8.heading("sr" , text = 'Sr.' , anchor='w')
class_tree_page8.heading("lecture" , text = 'Lecture' , anchor='w')
class_tree_page8.heading("a" , text = '1' , anchor='w')
class_tree_page8.heading("b" , text = '2' , anchor='w')
class_tree_page8.heading("c" , text = '3' , anchor='w')
class_tree_page8.heading("d" , text = '4' , anchor='w')
class_tree_page8.heading("e" , text = '5' , anchor='w')
class_tree_page8.heading("f" , text = '6' , anchor='w')
class_tree_page8.heading("g" , text = '7' , anchor='w')
class_tree_page8.heading("h" , text = '9' , anchor='w')
class_tree_page8.heading("i" , text = '10' , anchor='w')
class_tree_page8.heading("j" , text = '11' , anchor='w')
class_tree_page8.heading("k" , text = '12' , anchor='w')
class_tree_page8.heading("l" , text = '13' , anchor='w')
class_tree_page8.heading("m" , text = '14' , anchor='w')
class_tree_page8.heading("n" , text = '15' , anchor='w')
class_tree_page8.heading("o" , text = '16' , anchor='w')
class_tree_page8.heading("p" , text = '17' , anchor='w')
class_tree_page8.heading("q" , text = '18' , anchor='w')
class_tree_page8.heading("r" , text = '19' , anchor='w')
class_tree_page8.heading("s" , text = '20' , anchor='w')
class_tree_page8.heading("t" , text = '21' , anchor='w')
class_tree_page8.heading("u" , text = '22' , anchor='w')
class_tree_page8.heading("v" , text = '23' , anchor='w')
class_tree_page8.heading("w" , text = '24' , anchor='w')
class_tree_page8.heading("x" , text = '25' , anchor='w')
class_tree_page8.heading("y" , text = '26' , anchor='w')
class_tree_page8.heading("z" , text = '27' , anchor='w')
class_tree_page8.heading("aa" , text = '28' , anchor='w')
class_tree_page8.heading("ab" , text = '29' , anchor='w')
class_tree_page8.heading("ac" , text = '30' , anchor='w')
class_tree_page8.heading("ad" , text = '31' , anchor='w')

for i in class_columns_page8:
    class_tree_page8.column(i , width=50 , anchor='w')

class_tree_page8.column('sr' , width=200 , anchor='w')
class_tree_page8.column('lecture' , width=200 , anchor='w')


###############################################
#nb.select(1)
#nb2.select(2)
###############################################

def quit_me():
    op = MessageBox("Confirm",'Press the desired button ! ',b1='Exit',b2='Cancel',b3='Login')
    if op.choice=='Login':
        reset_page1()
        Login_System(root=root,nb=nb)
    elif op.choice=='Exit':
        root.quit()
        root.destroy()

def quit_me_2():
    op = MessageBox("Confirm",'Do you really wish to exit?',b1='Yes',b2='No')
    if op.choice=='Yes':
        root.quit()
        root.destroy()


root.protocol("WM_DELETE_WINDOW", quit_me)

txt_date.bind("<<DateEntrySelected>>", record) 
com_class_page1.bind('<<ComboboxSelected>>', lambda e : get_lectures(e,clear=True))
com_lecture_page1.bind('<<ComboboxSelected>>', record)

def unbinder():
    try:
        root.unbind('<Return>')
        root.unbind('<Control-o>')
        root.unbind('<Control-u>')
        root.unbind('<Control-d>')
        root.unbind('<Control-r>')
        root.unbind('<Control-s>')
    except:
        pass

def assign_commands_nb(event=None):
    tab = nb.index(nb.select())
    unbinder()
    if tab==0:
        unbinder()
        root.bind("<Return>",add_attendance)
        root.bind("<Control-u>",update_attendance)
        root.bind("<Control-d>",delete_attendance)
        root.bind("<Control-r>",reset_page1)
    elif tab==1:
        unbinder()
        root.bind("<Return>",add_class)
        root.bind("<Control-u>",update_class)
        root.bind("<Control-d>",delete_class)
        root.bind("<Control-r>",reset_class)
        root.bind("<Control-s>",export_class)

def assign_commands_nb2(event=None):
    tab = nb2.index(nb2.select())
    if tab==0 or tab=="":
        root.bind("<Return>",add_class)
        root.bind("<Control-u>",update_class)
        root.bind("<Control-d>",delete_class)
        root.bind("<Control-r>",reset_class)
        root.bind("<Control-s>",export_class)
    if tab==1:
        root.bind("<Return>",add_student)
        root.bind("<Control-u>",update_student)
        root.bind("<Control-d>",delete_student)
        root.bind("<Control-r>",reset_student)
        root.bind("<Control-s>",export_student)
        root.bind("<Control-o>",sort_student)
    if tab==2:
        root.bind("<Return>",add_result)
        root.bind("<Control-u>",update_result)
        root.bind("<Control-d>",delete_result)
        root.bind("<Control-r>",reset_result)
        root.bind("<Control-s>",export_result)
        root.bind("<Control-o>",sort_result)
    if tab==3:
        root.bind("<Return>",add_user)
        root.bind("<Control-u>",update_user)
        root.bind("<Control-d>",delete_user)
        root.bind("<Control-r>",reset_user)
        root.bind("<Control-s>",export_user)
        root.bind("<Control-o>",sort_user)
    if tab==4:
        unbinder()

class_tree_page1.bind("<ButtonRelease-1>",recordCsv)

nb.bind("<<NotebookTabChanged>>",lambda e : assign_commands_nb(e))
nb2.bind("<<NotebookTabChanged>>",lambda e : assign_commands_nb2(e))


def update_classes(event=None):
    com_course_code_page3.configure(values=classes)
    com_class_page1.configure(values=classes)
    com_class_page5.configure(values=classes)
    com_class_report.configure(values=classes)

left_column_page1.rowconfigure((3),weight=1)

check_previous_Login(root,nb)
reset_class()
reset_student()
reset_user()
reset_result()
unbinder()
root.bind("<Return>",add_attendance)
root.bind("<Control-u>",update_attendance)
root.bind("<Control-d>",delete_attendance)
root.bind("<Control-r>",reset_page1)

def delete_whole_word(event):
    val = event.widget.get()
    index = (event.widget.index(INSERT))
    current = val.find(index)
    print(current)
    return "break"

root.bind_class("ttk.Entry", "Control-i",lambda e: delete_whole_word(e))
txt_lecture_page2.bind("<Control-Delete>",lambda e: delete_whole_word(e))
#nb2.hide(4)
#nb2.hide(5) 
if Authenticate:
    Login_System(root,nb)
    if LOGGER !='Admin':
        nb.tab(1,state='disabled')
    else:
        try:
            nb.tab(1,state='normal')
        except:
            pass
root.mainloop()