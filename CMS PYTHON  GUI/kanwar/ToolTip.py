from tkinter import ttk,Toplevel,Tk
class CreateToolTip(object):
    """
    create a tooltip for a given widget
    """
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 512   #pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None
        self.enter()

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        if not self.text==None:
            self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        if not self.widget['state']=='disabled':
            x = y = 0
            x, y, cx, cy = self.widget.bbox("insert")
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 20
            # creates a toplevel window
            self.tw = Toplevel(self.widget)
            # Leaves only the label and removes the app window
            self.tw.wm_overrideredirect(True)
            self.tw.wm_geometry("+%d+%d" % (x, y))
            label = ttk.Label(self.tw,text=self.text, justify='left',
                        wraplength = self.wraplength,font=("Helvetica",12),style='info.TLabel')
            label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw= None
        if tw:
            tw.destroy()

if __name__=='__main__':
    from tkinter import ttk
    from ttkbootstrap import Style
    style = Style("united")
    root = style.master
    root.title("Kanwar Adnan")
    button = ttk.Button(root , text='Button',style='Outline.TButton')
    button.pack(fill='both' , expand=1)
    CreateToolTip(button , 'Made by Kanwar Adnan (kanwaradnanrajput@gmail.com)')
    root.mainloop()