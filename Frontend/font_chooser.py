from tkinter import *
from tkinter import ttk
from tkinter import font
import re


class FontChooser(Toplevel):
    def __init__(self):
        Toplevel.__init__(self)
        self.FontInit()
        self.window()

    def FontInit(self):
        self.weight = False
        self.slant= False
        self.underline = False
        self.strike = False
        self.fonts = {}
        self.selected_font = ''
        self.fonts_style_list=[
            "Arial",
            12,
            False,
            False,
            False,
            False
        ]


    def window(self):
        self.title("Select Font")
        self.geometry("250x100")
        self.selected_font = StringVar()
        available_fonts=list(font.families())
        available_fonts.sort()
        self.choose_font=ttk.Combobox(self, width=20, textvariable=self.selected_font, values=available_fonts,)
        self.choose_font.current(31)
        self.choose_font.place(x=10, y=20)
        self.choose_font.bind("<Key>",lambda f:self.FontSearch())
        text_sizes = [i for i in range(0,24,2)]+[i for i in range(24,80,4)] + [9, 11, 13, 15, 17]
        text_sizes.sort()
        self.selected_size = IntVar()
        self.default_size  = 5
        self.choose_size=ttk.Combobox(self, width=2, values= text_sizes, textvariable=self.selected_size)
        self.choose_size.current(self.default_size)
        self.choose_size.place(x=153, y=20)
        self.bold_button=Button(self, text='B', font=font.Font(weight = 'bold', size=9), width=2, command= self.Bold)
        self.bold_button.place(x=10, y=50)
        self.italic_button=Button(self, text='I', font=font.Font(slant = 'italic', size=9), width=2, command= self.Italic)
        self.italic_button.place(x=40, y=50)
        self.underline_button=Button(self, text='U', width=2, underline=False, command= self.Underline)
        self.underline_button.place(x=70, y=50)
        self.strike_button=Button(self, text='ab', font=font.Font(overstrike=True, size=9), command = self.Strike)
        self.strike_button.place(x=100, y=50)
        Button(self, text='A+', command= self.SizeInc).place(x=140, y=50)
        Button(self, text='A-', command = self.SizeDec).place(x=170, y=50)
        Button(self, text= 'Ok', width= 5, command= self.Ok).place(x=200, y= 20)
        Button(self, text= 'Cancel', width= 5, command = self.destroy).place(x=200, y= 50)
        self.tk_focusFollowsMouse()

    def FontSearch(self):
        values = []
        for i in font.families():
            s=re.match(r"^%s"%(self.selected_font.get()),i, re.I)
            if s:
                values.append(i)
        self.choose_font["values"]=values
        self.choose_font.tk_focusFollowsMouse()

    def Bold(self):
        if self.weight:
            self.bold_button.config(bg="whitesmoke")
            self.weight = False
        else:
            self.bold_button.config(bg="darkgrey")
            self.weight = True

    def Italic(self):
        if self.slant:
            self.italic_button.config(bg="whitesmoke")
            self.slant = False
        else:
            self.italic_button.config(bg="darkgrey")
            self.slant = True
    def Underline(self):
        if self.underline:
            self.underline_button.config(bg="whitesmoke")
            self.underline = False
        else:
            self.underline_button.config(bg="darkgrey")
            self.underline = True
    def Strike(self):
        if self.strike:
            self.strike_button.config(bg="whitesmoke")
            self.strike = False
        else:
            self.strike_button.config(bg="darkgrey")
            self.strike = True
    def SizeInc(self):
        if self.default_size <25:
            self.default_size+=1
            self.choose_size.current(self.default_size)
    def SizeDec(self):
        if self.default_size >0:
            self.default_size-=1
            self.choose_size.current(self.default_size)

    def Ok(self):
        self.fonts = {"font":self.selected_font.get(),
                      "size":self.selected_size.get(),
                      "weight":self.weight,
                      "slant":self.slant,
                      "underline":self.underline,
                      "strike":self.strike,
                      }
        self.fonts_style_list = [
            self.selected_font.get(),
            self.selected_size.get(),
            self.weight,
            self.slant,
            self.underline,
            self.strike
        ]
        self.destroy()
        

def FontDialog():
    f=FontChooser()
    f.wait_window()
    return f.fonts_style_list