from tkinter import ttk , StringVar, BooleanVar


class Login_Page(ttk.Frame):

    labels = {'id' : 'Email : ','pass':'Password : '}
    buttons = {'signin' : 'Login' , 'signup' : 'Sign Up'}
    paddings = {'padx' : [10,0] , 'pady' : [20,0]}

    button_width = 10
    button_properties = {'sticky' : 'w'}
    entry_properties = {'sticky' : 'ew'}
    label_properties = {'sticky' : 'e'}

    entry_properties.update(**paddings)
    button_properties.update(**paddings)
    label_properties.update(**paddings)

    def __init__(self,root, *args,**kwargs):
        self.root = root
        ttk.Frame.__init__(self,self.root,*args,**kwargs)
        self.Pass = StringVar()
        self.Id = StringVar()
        self.Msg = StringVar()
        self.Keep_login = BooleanVar()

        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(row = 0 , column=0 , sticky='ew')
        self.main_frame.focus()

        self.lbl_id = ttk.Label(self.main_frame, text = Login_Page.labels['id'])
        self.lbl_id.grid(row=0,column=0,**Login_Page.label_properties)

        self.txt_id = ttk.Entry(self.main_frame,textvariable=self.Id)
        self.txt_id.grid(row=0,column=1,**Login_Page.entry_properties)

        self.lbl_pass = ttk.Label(self.main_frame, text = Login_Page.labels['pass'])
        self.lbl_pass.grid(row=1,column=0,**Login_Page.label_properties)

        self.txt_pass = ttk.Entry(self.main_frame,textvariable=self.Pass)
        self.txt_pass.grid(row=1,column=1,**Login_Page.entry_properties)

        self.chk_keep_login = ttk.Checkbutton(self.main_frame, text='Keep me Logged in' , variable=self.Keep_login)
        self.chk_keep_login.grid(row=2,column=1,sticky='w',**Login_Page.paddings)

        self.lbl_message = ttk.Label(self.main_frame, textvariable=self.Msg,style='danger.TLabel')
        self.lbl_message.grid(row=2,column=0,columnspan=3,sticky='e' , **Login_Page.paddings)

        self.bottom_frame = ttk.Frame(self)
        self.bottom_frame.grid(row = 1 , column=0)

        self.btn_sign_in = ttk.Button(self.bottom_frame , text = Login_Page.buttons['signin'] , width=Login_Page.button_width)
        self.btn_sign_in.grid(row = 0 , column = 1,**Login_Page.button_properties)

        self.btn_sign_up = ttk.Button(self.bottom_frame , text = Login_Page.buttons['signup'],style='primary.Outline.TButton' , width=Login_Page.button_width)
        self.btn_sign_up.grid(row = 0 , column = 2,**Login_Page.button_properties)

        self.main_frame.columnconfigure((0,1,2),weight=1)
        self.bottom_frame.columnconfigure((0,1,2),weight=1)

        self.columnconfigure(0,weight=1)

if __name__ == '__main__':
     # IMPORT MODERN TTK LIB
    from ttkbootstrap import Style

    # MAKING IT'S OBJECT SETTINGS
    style = Style('united')

    # MAKING ROOT OBJECT    
    root = style.master
    root.title("Kanwar Adnan")
    root.geometry("660x240+100+100")
    root.resizable(False,False) 

    # MAKING LOGIN PAGE
    page_login_object = Login_Page(root)
    page_login_object.pack(expand=1,fill='both',padx=10,pady=10)
    page_login_object.txt_pass.configure(show='*')
    
    root.mainloop()
