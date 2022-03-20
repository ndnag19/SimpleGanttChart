from tkinter import *
from tkinter import filedialog,Toplevel
from tkinter.messagebox import showinfo, showerror, showwarning
import subprocess,os,platform
from pptx.text.fonts import FontFiles
from pptx import Presentation

# Generates a global list of fonts available on the system
def fontList():
    font_keys = list(FontFiles._installed_fonts().keys())
    fonts = []
    for i in font_keys:
        fonts.append(str(i[0]))
    return fonts

# Generates a list of slide_layouts from any given presentation
def slideLayoutList(prs: Presentation):
    slide_layouts = prs.slide_layouts
    names = []
    for slide in slide_layouts:
        names.append(slide.name)
    return names

# Opens dialog dialog boxes for PowerPoint Presentation files
def browsePowerPointButton(root: Tk,textbox: Entry):
    root.filename = filedialog.askopenfilename(title="Select a powerpoint presentation",
    filetypes=(("PowerPoint Files","*.pptx"),
    ("PowerPoint Macro Enabled Files","*.pptm"),("All Files","*.*")))
    textbox.delete(0,END)
    textbox.insert(0,root.filename)

# Opens dialog boxes for Excel/Csv data files
def browseDataButton(root: Tk,textbox: Entry):
    root.filename = filedialog.askopenfilename(title="Select Data",
    filetypes=(("Excel Files","*.xlsx"),("Comma Seperated Values","*.csv"),
    ("Excel Macro Enabled Files","*.xlsm"),("All Files","*.*")))
    textbox.delete(0,END)
    textbox.insert(0,root.filename)

# Save dialog boxes for Presentation files
def browseSaveButton(root: Tk,textbox: Entry, prs):
    root.filename = filedialog.asksaveasfilename(defaultextension=".*",title="Save Presentation",
    filetypes=(("PowerPoint Files","*.pptx"),
    ("PowerPoint Macro Enabled Files","*.pptm"),("All Files","*.*")))
    textbox.delete(0,END)
    textbox.insert(0,root.filename)
    try:
        prs.save(textbox.get())
        showinfo("Presentation saved successfully",f"Your file is saved successfully at {textbox.get()}")
    except:
        showwarning("Saving Unsuccessful!",
        "Please close the powerpoint file before saving.")

# Changes the frame useful for navigation
def showFrame(frame: Frame):
    frame.tkraise()

# Open links using the web browser
def openLinks(link: str):
    if platform.system()=="Darwin":
        subprocess.Popen(("open",link))
    elif platform.system()=="Windows":
        os.startfile(link)
    else:
        try:
            subprocess.call(("xdg-open",link))
        except OSError:
            showerror("OS Error",f"Please visit the page at {link}")