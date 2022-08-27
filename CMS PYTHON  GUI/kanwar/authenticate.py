from kanwar.MessageBox import MessageBox
from kanwar.Login import Login_Page
from kanwar.SignUp import SignUp_Page
from tkinter import * 
import sqlite3
from time import strftime

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

def Login_System(root,nb):
    root.withdraw()

    page_login = Toplevel(root)
    page_login.geometry("660x200+200+200")
    page_login.resizable(False,False)

    page_login_object = Login_Page(page_login)
    page_login_object.pack(expand=1,fill='both')
    page_login_object.txt_pass.configure(show="*")

    def Sign_up_page(event=None):
        page_login.withdraw()

        page_signup = Toplevel(page_login)
        page_signup.focus()
        page_signup.geometry("660x240+200+200")
        page_signup.resizable(False,False)

        page_signup_object = SignUp_Page(page_signup)
        page_signup_object.pack(expand=1,fill='both',padx= 10 , pady =10)
        page_signup_object.txt_name.focus()

        page_signup_object.txt_pass.configure(show="*")
        page_signup_object.txt_cpass.configure(show="*")
        def create_account():
            Name = page_signup_object.Name.get().strip()
            Email = page_signup_object.Email.get().strip()
            Pass = page_signup_object.Pass.get().strip()
            CPass = page_signup_object.CPass.get().strip()
            for _ in range(1):
                if Name == '':
                    page_signup_object.Msg.set("Enter Name !")
                    page_signup_object.txt_name.focus()
                    break
                if Email == '':
                    page_signup_object.Msg.set("Enter Email Address !")
                    page_signup_object.txt_email.focus()
                    break
                if not "@" in Email:
                    page_signup_object.Msg.set("Not a valid email format. Must inlcude '@'.")
                    page_signup_object.txt_email.focus()
                    break
                if Pass == '':
                    page_signup_object.Msg.set("Enter Password !")
                    page_signup_object.txt_pass.focus()
                    break
                if CPass == '':
                    page_signup_object.Msg.set("Confirm your Password !")
                    page_signup_object.txt_cpass.focus()
                    break
                if CPass != Pass:
                    page_signup_object.Msg.set("Password and Confirm Password should be same !")
                    page_signup_object.txt_cpass.focus()
                    break
                else:
                    try:
                        con = sqlite3.connect(database=DATABASE)
                        cur = con.cursor()
                        cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
                        cur.execute("select * from USERS where email=?",(Email,))
                        row=cur.fetchone()
                        if row!=None:
                            page_signup_object.Msg.set("User already exists !")
                            break
                        else:
                            cur.execute("insert into USERS (name,email,password,role) values(?,?,?,?)",(
                                Name,
                                Email,
                                Pass,
                                DEFAULT_ROLE_USER
                            ))
                            con.commit()
                            MessageBox("Success","Your Account has been created !\nWait for the Admin to assign you a role Till then you can take attendance only. ")
                            page_login.deiconify()
                            page_login_object.txt_id.focus()
                            page_login_object.Id.set("")
                            page_login_object.Pass.set("")
                            page_signup.destroy()
                    except Exception as ex:
                        MessageBox("Error",f"Error due to {str(ex)}")

        page_signup_object.btn_create.configure(command=create_account)
        def del_window(event=None):
            page_login.deiconify()
            page_signup.destroy()
        page_signup.protocol("WM_DELETE_WINDOW",del_window)
        page_signup.wait_window()

    def confirm_login(event=None):
        Id = page_login_object.Id.get().strip()
        Pass = page_login_object.Pass.get().strip()

        for _ in range(1):
            if Id=='':
                page_login_object.Msg.set("Enter User Id !")
                page_login_object.txt_id.focus()
                break
            if Pass=='':
                page_login_object.Msg.set("Enter Password !")
                page_login_object.txt_pass.focus()
                break
            else:
                try:
                    con = sqlite3.connect(database=DATABASE)
                    cur = con.cursor()
                    cur.execute("CREATE TABLE IF NOT EXISTS USERS(cid INTEGER PRIMARY KEY AUTOINCREMENT,name text,role text ,course_code text,subjects text,email text, password text )")
                    cur.execute("select * from USERS where email=?",(Id,))
                    row=cur.fetchone()
                    print(row)
                    if row==None:
                        page_login_object.Msg.set("No User found !")
                        break
                    else:
                        if Pass!=row[6]:
                            page_login_object.Msg.set("Wrong Password !")
                        else:
                            try:
                                con.close()
                                con = sqlite3.connect(database=DATABASEATTENDNACE)
                                cur = con.cursor()
                                cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,email text,pass text,role text,date text,time text,rem text)")
                                con.commit()
                                cdate = strftime("%m/%d/%Y")
                                ctime =strftime("%H:%M")
                                cur.execute("insert into SESSIONS (email,pass,date,time,rem,role) values(?,?,?,?,?,?)",(
                                    Id,
                                    Pass,
                                    cdate,
                                    ctime,
                                    page_login_object.Keep_login.get(),
                                    row[2]
                                ))
                                con.commit()
                                global LOGGER
                                LOGGER = row[2]
                                page_login_object.txt_id.focus()
                                root.deiconify()
                                page_login.destroy()
                                if LOGGER !='Admin':
                                    nb.tab(1,state='disabled')
                                else:
                                    nb.tab(1,state='normal')
                            except Exception as ex:
                                MessageBox("Error",f"Error due to {str(ex)}")
                except Exception as ex:
                    MessageBox("Error",f"Error due to {str(ex)}")

    # ASSIGNING CUSTOM FUNCTIONS TO BUTTONS
    page_login_object.btn_sign_up.configure(command=Sign_up_page)
    page_login_object.btn_sign_in.configure(command=confirm_login)

    # LAST THING TO BIND
    page_login.protocol("WM_DELETE_WINDOW",lambda : root.destroy())
    page_login.after(100,lambda : page_login.focus_set())
    page_login.wait_window()
    return LOGGER

def check_previous_Login(root,nb):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,email text,pass text,role text,date text,time text,rem text)")
        con.commit()
        cur.execute("Select * from sqlite_sequence where name=?",('SESSIONS',))
        row = cur.fetchone()
        print(row)
        if row!=None:
            cur.execute("Select * from SESSIONS where cid=?",(row[1],))
            row =cur.fetchone()
            user = row[1]
            password = row[2]
            #print(user,password)
            confirm = redirect(user,password,root,nb)
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}")

def redirect(user,password,root,nb):
    try:
        con = sqlite3.connect(database=DATABASEATTENDNACE)
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,email text,pass text,role text,date text,time text,rem text)")
        con.commit()
        cur.execute("select * from SESSIONS where email=?",(user,))
        row=cur.fetchall()
        cur.execute("delete from SESSIONS where email=?",(user,))
        con.commit()
        _user = row[-1][1]
        _password = row[-1][2]
        if user == _user and password == _password and row[-1][6]=='1':
            MessageBox("Information",f"You're currently Logged in as {_user.title()}")
            #print(f"You're currently Logged in as {_user.title()}")
            try:
                con = sqlite3.connect(database=DATABASEATTENDNACE)
                cur = con.cursor()
                cur.execute("CREATE TABLE IF NOT EXISTS SESSIONS(cid INTEGER PRIMARY KEY AUTOINCREMENT,email text,pass text,role text,date text,time text,rem text)")
                con.commit()
                cdate = strftime("%m/%d/%Y")
                ctime =strftime("%H:%M")
                cur.execute("insert into SESSIONS (email,pass,role,date,time,rem) values(?,?,?,?,?,?)",(
                    _user,
                    _password,
                    row[-1][3],
                    cdate,
                    ctime,
                    row[-1][6]
                ))
                con.commit()
                root.deiconify()
                if row[-1][3] !='Admin':
                    nb.tab(1,state='disabled')
                else:
                    try:
                        nb.tab(1,state='normal')
                    except:
                        pass
                return 'logged in'
            except Exception as ex:
                MessageBox("Error",f"Error due to {str(ex)}")
        else:
            con.close()
            Login_System(root,nb)
            return 'error'
    except Exception as ex:
        MessageBox("Error",f"Error due to {str(ex)}")
if __name__ == '__main__':
     # IMPORT MODERN TTK LIB
    from ttkbootstrap import Style

    # MAKING IT'S OBJECT SETTINGS
    style = Style('united')

    # MAKING ROOT OBJECT    
    root = style.master
    root.withdraw()
    root.title("Kanwar Adnan")
    MessageBox("Information","Authentication system of Kanwar Adnan (Kanwaradnanrajput@gmail.com)",b2='Close')