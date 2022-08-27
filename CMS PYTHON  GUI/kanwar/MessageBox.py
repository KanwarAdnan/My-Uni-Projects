from tkinter import ttk,Toplevel

class MessageBox:

    def __init__(self,title='MessageBox', msg='Hello I will be your message',parent='', b1='Ok', b2='',b3=''):
        self.title = title
        self.msg = msg
        self.choice = None
        self.b1 = b1
        self.b2 = b2
        self.b3 = b3
        self.parent = parent

        self.root = Toplevel()
        self.root.focus()
        self.root.title(self.title)
        self.root.resizable(False,False)

        self.root.update_idletasks()
        width = self.root.winfo_width()
        frm_width = self.root.winfo_rootx() - self.root.winfo_x()
        win_width = width + 2 * frm_width
        height = self.root.winfo_height()
        titlebar_height = self.root.winfo_rooty() - self.root.winfo_y()
        win_height = height + titlebar_height + frm_width
        x = self.root.winfo_screenwidth() // 2 - win_width // 2
        y = self.root.winfo_screenheight() // 2 - win_height // 2
        self.root.geometry('+{}+{}'.format(x, y))

        self.main_frame= ttk.Frame(self.root)
        self.main_frame.pack(fill='both',expand=1)

        self.message_frame= ttk.Frame(self.main_frame)
        self.message_frame.pack(side='top',fill='both',expand=1,pady=15, padx = 10)

        self.lbl_message = ttk.Label(self.message_frame , text=self.msg , wraplength=128*3 , font=("Helvetica",12))
        self.lbl_message.pack(side = 'left' , padx = 5 , pady = 1)

        self.footer_frame= ttk.Frame(self.main_frame)
        self.footer_frame.pack(side='bottom',fill='x',expand=1)

        self.button_1 = ttk.Button(self.footer_frame , text = self.b1 , default='active')
        self.button_1.pack(side = 'right' , padx = 5 , pady = 5)
        self.button_1.configure(command=self.click1)
        self.button_1.focus()

        if self.b2:

            self.button_2 = ttk.Button(self.footer_frame , text = self.b2, style='primary.Outline.TButton')
            self.button_2.pack(side = 'right' , padx = 5 , pady = 5)
            self.button_2.configure(command=self.click2)

        if self.b3:

            self.button_2 = ttk.Button(self.footer_frame , text = self.b3, style='primary.Outline.TButton')
            self.button_2.pack(side = 'right' , padx = 5 , pady = 5)
            self.button_2.configure(command=self.click3)

        self.button_1.bind("<Return>",lambda e: self.click1())
        self.root.bind("<Escape>",lambda e: self.root.destroy())

        #self.root.after(5000,lambda : self.root.destroy()) # UNCOMMENT IT IF YOU WANT AUTOMATIC DESTROY MESSAGE BOXES
        self.root.wait_window()

    def closed(self):
        self.root.destroy()
        self.choice='closed'
        
    def click1(self,event=None):
        self.root.destroy()
        self.choice=self.b1

    def click2(self,event=None):
        self.root.destroy()
        self.choice=self.b2

    def click3(self,event=None):
        self.root.destroy()
        self.choice=self.b3

if __name__ == '__main__':
     # IMPORT MODERN TTK LIB
    from ttkbootstrap import Style

    # MAKING IT'S OBJECT SETTINGS
    style = Style('united')

    # MAKING ROOT OBJECT    
    root = style.master
    root.withdraw()
    root.title("Kanwar Adnan")
    MessageBox("Information","Hello It's Kanwar Adnan (kanwaradnanrajput@gmail.com)",b2='Close')