from tkinter import ttk , StringVar

class SignUp_Page(ttk.Frame):
    labels = {'name' : 'Name : ' , 'email' : 'Email : ' , 'pass' : 'Password : ' , 'cpass' : 'Confirm : '}
    buttons = {'create' : 'Create' , 'cancel' : 'Cancel'}

    paddings = {'padx' : [10,0] , 'pady' : [20,0]}

    button_width = 10

    button_properties = {'sticky' : 'w'}
    entry_properties = {'sticky' : 'ew'}
    label_properties = {'sticky':'e'}

    entry_properties.update(**paddings)
    button_properties.update(**paddings)
    label_properties.update(**paddings)

    def __init__(self,root, *args,**kwargs):
        self.root = root
        ttk.Frame.__init__(self,self.root,*args,**kwargs)
        self.Name = StringVar()
        self.Email = StringVar()
        self.Pass = StringVar()
        self.CPass = StringVar()
        self.Msg = StringVar()

        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(row = 0 , column=0 , sticky='ew')
        self.main_frame.focus()

        self.lbl_name = ttk.Label(self.main_frame , text = SignUp_Page.labels['name'])
        self.lbl_name.grid(row = 0 , column= 0 , **SignUp_Page.label_properties)

        self.txt_name = ttk.Entry(self.main_frame , textvariable=self.Name)
        self.txt_name.grid(row = 0 , column= 1 , **SignUp_Page.entry_properties)

        self.lbl_email = ttk.Label(self.main_frame , text = SignUp_Page.labels['email'])
        self.lbl_email.grid(row = 0 , column= 2 , **SignUp_Page.label_properties)

        self.txt_email = ttk.Entry(self.main_frame , textvariable=self.Email)
        self.txt_email.grid(row = 0 , column= 3 , **SignUp_Page.entry_properties)

        self.lbl_pass = ttk.Label(self.main_frame , text = SignUp_Page.labels['pass'])
        self.lbl_pass.grid(row = 1 , column= 0 , **SignUp_Page.label_properties)

        self.txt_pass = ttk.Entry(self.main_frame , textvariable=self.Pass)
        self.txt_pass.grid(row = 1 , column= 1 , **SignUp_Page.entry_properties)

        self.lbl_cpass = ttk.Label(self.main_frame , text = SignUp_Page.labels['cpass'])
        self.lbl_cpass.grid(row = 1 , column= 2 , **SignUp_Page.label_properties)

        self.txt_cpass = ttk.Entry(self.main_frame , textvariable=self.CPass)
        self.txt_cpass.grid(row = 1 , column= 3 , **SignUp_Page.entry_properties)

        self.lbl_message = ttk.Label(self.main_frame, textvariable=self.Msg,style='danger.TLabel')
        self.lbl_message.grid(row=2,column=1,columnspan=3 , **SignUp_Page.paddings,sticky='w')

        self.bottom_frame = ttk.Frame(self)
        self.bottom_frame.grid(row = 1 , column=0 , sticky='ew')

        self.btn_create = ttk.Button(self.bottom_frame , text = SignUp_Page.buttons['create'] , width=SignUp_Page.button_width)
        self.btn_create.pack(side='left',padx = [130,0] , pady = [40,0])

        self.btn_cancel = ttk.Button(self.bottom_frame , text = SignUp_Page.buttons['cancel'],style='primary.Outline.TButton' , width=SignUp_Page.button_width)
        self.btn_cancel.pack(side='left',padx = [10] , pady = [40,0])
        self.btn_cancel.configure(command=lambda : self.root.destroy())

        self.bottom_frame.columnconfigure((0,1),weight=1)

        self.main_frame.columnconfigure((0,1,2,3),weight=1)

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
    page_signin_object = SignUp_Page(root)
    page_signin_object.pack(expand=1,fill='both',padx=10,pady=10)
    page_signin_object.txt_pass.configure(show='*')
    
    root.mainloop()
