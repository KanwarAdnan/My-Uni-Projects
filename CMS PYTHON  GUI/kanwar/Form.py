from tkinter import *
from tkinter import ttk
from ttkbootstrap import Style

class MyFrame(ttk.Frame):
    vals = ['Present','Absent','Leave']
    def __init__ (self,root,window,*args,**kwargs):
        self.root = root
        self.window = window
        self.data = []
        self.boxes = []
        self.padx = 10

        self.col_0 = [] # Labels
        self.col_1 = [] # Labels
        self.col_2 = [] # Labels

        self.pady = 10

        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(expand=1,fill=BOTH)

        self.right_frame = ttk.Frame(self.main_frame)
        self.right_frame.pack(side = 'right' , fill = 'y')

        self.mid_frame = ttk.Frame(self.main_frame)
        self.mid_frame.pack(expand=1 , fill = 'both')

        self.bottom_frame = ttk.Frame(self.main_frame)
        self.bottom_frame.pack(side = 'bottom' , fill = 'x')

        self.scrollx = ttk.Scrollbar(self.bottom_frame,orient='horizontal')

        self.scrolly = ttk.Scrollbar(self.right_frame,orient='vertical')
        self.scrolly.pack(side = 'right' , fill = 'y')
        self.scrollx.pack(side = 'bottom' , fill = 'x')

        self.canvas = Canvas(self.mid_frame,yscrollcommand=self.scrolly.set,xscrollcommand=self.scrollx.set, highlightthickness=0, relief='flat' , border=0)
        self.canvas.pack(expand=True,fill=BOTH , side=LEFT)

        self.scrolly.config(command=self.canvas.yview)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion = self.canvas.bbox("all")))
        self.scrollx.config(command=self.canvas.xview)

        ttk.Frame.__init__(self,self.canvas,*args,**kwargs)
        self.pack(expand=1,fill=BOTH)

        self.canvas.create_window((0,0),window=self,anchor='nw')

        self.bind("<Enter>", self.entered)
        self.bind("<Leave>", self.left)

    def _on_mouse_wheel(self,event):
        self.canvas.update_idletasks()
        self.canvas.yview_scroll(-1 * int((event.delta / 60)), "units")

    def entered(self,event):
        self.canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind("<1>",     lambda event: self.canvas.focus_set())
        self.canvas.bind("<Up>",    lambda event: self.canvas.yview_scroll(-5, "units"))
        self.canvas.bind("<Down>",  lambda event: self.canvas.yview_scroll( 5, "units"))
        self.canvas.bind("<Left>",    lambda event: self.canvas.xview_scroll(-5, "units"))
        self.canvas.bind("<Right>",  lambda event: self.canvas.xview_scroll( 5, "units"))

#        self.canvas.focus_set()
        
    def left(self,event):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind("<1>")
        self.canvas.unbind("<Up>")
        self.canvas.unbind("<Down>")
        self.canvas.unbind("<Left>")
        self.canvas.unbind("<Right>")

    def even_odd(self,data,order=None):
        if order=='even':
            return [j for i,j in enumerate(data) if i%2==0]
        elif order=='odd':
            return [j for i,j in enumerate(data) if i%2!=0]
        else:
            return data

    def make_rows(self,data,order=None):
        self.data = data.copy()
        data = self.even_odd(data,order=order)
        self.add_data(data)

    def add_data(self,data):
        column = 0
        prime_index = 0
        font =  ('Helvetica',12)
        self.Label = Label(self , text= 'Sr # ' , font=("Helvetica",12,'bold'))
        self.Label.grid(row = 0 , column=0)

        self.Label = Label(self , text= 'Name ' , font=("Helvetica",12,'bold'))
        self.Label.grid(row = 0 , column=1 , sticky=W)

        self.Label = Label(self, text= 'Roll No. ' , font=("Helvetica",12,'bold'))
        self.Label.grid(row = 0 , column=2)

        #self.Label = Label(self , text= 'Status ' , font=("Helvetica",12,'bold'))
        #self.Label.grid(row = 0 , column=3 , sticky=W)

        self.com_mode = ttk.Combobox(self, state='readonly',values=MyFrame.vals , width=10)
        self.com_mode.grid(row = 0 , column = 3 , sticky=W , pady=10)
        self.com_mode.current(0)

        self.com_mode.bind("<<ComboboxSelected>>",self.assign)

        for index in range(len(data)):
            prime_index += 1
            val = Label(self , text=prime_index , font=font)
            val.grid(row = index+1 , column=0 , pady = self.pady)
            self.col_0.append(val)

            for col,value in enumerate(data[index]):
                if col == 1:
                    val = Label(self,text=value,anchor=CENTER,justify='left',font=font)
                    val.grid(row=index+1,column=col+1 , sticky='ns')
                else:
                    val = Label(self,text=value.upper(),anchor='w',justify='left',font=font)
                    val.grid(row=index+1,column=col+1 , sticky=W)
                if col == 0:
                    self.col_1.append(val)
                elif col == 1:
                    self.col_2.append(val)

                column = col+1

            self.grid_columnconfigure([i for i in range(len(data[index])+2)], weight=1,uniform="column",pad=self.padx)

            comb = ttk.Combobox(self ,state='readonly', values=MyFrame.vals , width=10 , justify='left')
            comb.current(0)
            comb.grid(row=index+1,column=column+1 , pady = self.pady , sticky=W)
            comb.bind("<Tab>",lambda e : self.next(e))
            comb.bind("<Shift-Tab>",lambda e : self.prev(e))

            self.boxes.append(comb)

    def assign(self,event=None):
        info = self.Status_boxes()
        for i in info:
            i.set(self.com_mode.get())

    def prev(self,event=None):
        self.canvas.yview_scroll(-1,'units')
        event.widget.tk_focusPrev.focus()

    def next(self,event=None):
        self.canvas.yview_scroll(1, "units")
        event.widget.tk_focusNext().focus()

    def Status_boxes(self,event=None):
        return self.boxes.copy()

    def Data(self,event=None):
        return self.data.copy()

    def Boxes(self):
        return [i.get() for i in self.boxes.copy()]

    def add_item(self,data,item):
        li = []
        li.clear()
        for j,i in enumerate(data):
            i.append(item[j])
            li.append(i)
        return li

    def returninfo(self,event=None):
        info = self.data.copy()
        vals = self.boxes.copy()
        #remainder = self.add_item(info,vals)
        return None

if __name__ == '__main__':
    root = Style(theme='united')
    root = root.master
    root.geometry('900x600+0+0')
    root.title("Kanwar Adnan")
    f2 = ttk.Frame(root)
    f2.pack(fill=BOTH,expand=1)
    f1 = MyFrame(f2,root)

    #data = [(str(i),str(i**2)) for i in range(101)]
    data = [
            ['Made by Kanwar Adnan','1'],
            ['kanwaradnanrajput@gmail.com','2']
            ]
    f1.make_rows(data*10)
    f3 = ttk.Frame(root)
    f3.pack(side ='bottom' , fill='x')

    def add_item(data,item):
        a = list(zip(data,item))

    a = f1.Data()
    b = f1.Boxes()
    add_item(a,b)
    root.mainloop()