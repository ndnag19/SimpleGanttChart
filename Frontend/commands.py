from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showinfo,showerror
import subprocess,os,platform
from pptx.text.fonts import FontFiles

# Generates a global list of fonts available on the system
def fontList():
    font_keys = list(FontFiles._installed_fonts().keys())
    fonts = []
    for i in font_keys:
        fonts.append(str(i[0]))
    return fonts


# Opens dialog dialog boxes for PowerPoint Presentation files
def browsePowerPointButton(root: Tk,textbox: Entry):
    root.filename = filedialog.askopenfilename(title="Select a powerpoint presentation",
    filetypes=(("PowerPoint Files","*.pptx"),("PowerPoint Legacy Files","*.ppt"),
    ("PowerPoint Macro Enabled Files","*.pptm"),("All Files","*.*")))
    textbox.delete(0,END)
    textbox.insert(0,root.filename)

# Opens dialog boxes for Excel/Csv data files
def browseDataButton(root: Tk,textbox: Entry):
    root.filename = filedialog.askopenfilename(title="Select Data",
    filetypes=(("Excel Files","*.xlsx"),("Comma Seperated Values","*.csv"),("Excel Legacy Files","*.xls"),
    ("Excel Macro Enabled Files","*.xlsm"),("All Files","*.*")))
    textbox.delete(0,END)
    textbox.insert(0,root.filename)

# Save dialog boxes for Presentation files
def browseSaveButton(root: Tk,textbox: Entry, prs):
    root.filename = filedialog.asksaveasfilename(defaultextension=".*",title="Save Presentation",
    filetypes=(("PowerPoint Files","*.pptx"),("PowerPoint Legacy Files","*.ppt"),
    ("PowerPoint Macro Enabled Files","*.pptm"),("All Files","*.*")))
    textbox.delete(0,END)
    textbox.insert(0,root.filename)
    prs.save(textbox.get())
    showinfo("Presentation saved successfully",f"Your file is saved successfully at {textbox.get()}")

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
