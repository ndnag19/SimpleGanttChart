# Class for saving the file

from tkinter.ttk import Frame, Entry, Button
from tkinter import Tk, PhotoImage, Canvas
from Frontend.commands import *
from Frontend.strings import *
from pathlib import Path
import PIL.Image
import PIL.ImageTk
import PIL.ImageOps
import subprocess, os, platform

class SaveScreenFrame():
    def __init__(self, root: Tk, controller):
        super().__init__()

        # Locates the assests folder and stores it in ASSETS_PATH
        self.OUTPUT_PATH = Path(__file__).parent
        self.ASSETS_PATH = self.OUTPUT_PATH / Path(str_assets)

        # Initialises the root and controller classes
        self.controller = controller
        self.root=root
        
        # Creates a frame in the root window and places it such that it covers whole window
        self.frame = Frame(root)
        self.frame.grid(row=0,column=0,sticky="nsew")
        self.frame.configure(bg = "#938E8E")
        
        # Creates a canvas with desired properties
        self.canvas = Canvas(
            self.frame,
            bg = "#938E8E",
            height = 700,
            width = 984,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )
        # Places a rectange with desired color to cover whole canvas
        self.canvas.place(x = 0, y = 0)
        self.canvas.create_rectangle(
            0.0,
            0.0,
            984.0,
            700.0,
            fill="#000000",
            outline="")

        # Title
        self.canvas.create_text(
            283.0,
            42.0,
            anchor="nw",
            text=str_project_title,
            fill="#FFFFFF",
            font=("RobotoRoman SemiBold", 48 * -1)
        )

        #Label for selecting file name
        self.canvas.create_text(
            54.0,
            204.0,
            anchor="nw",
            text="Select path to save your presentation:",
            fill="#FFFFFF",
            font=("RobotoRoman Medium", 30 * -1)
        )

        # Entry of saving ppt path
        self.entry_presentation_path = Entry(self.frame,
            bd=0,
            bg="#ECECEC",
            highlightthickness=2,
            width=39,
            font=("RobotoRoman",24)
        )
        self.canvas.create_window(54,250,anchor="nw",window=self.entry_presentation_path)
        self.entry_presentation_path.insert(0,"Open file navigator by clicking browse")

        # Button to browse ppt path
        self.button_browse_img = PhotoImage(
            file=self.relative_to_assets(str_browse_button_img_name))
        self.button_ppt_browse = Button(self.canvas,
            image=self.button_browse_img,
            borderwidth=0,
            highlightthickness=2,
            command=lambda: self.selfBrowseSaveButton(),
            relief="raised"
        )
        self.button_ppt_browse.place(
            x=778.0,
            y=243.0,
            width=150.0,
            height=50.0
        )
        # Button to go to previous frame
        self.button_previous_img = PhotoImage(
            file=self.relative_to_assets(str_previous_button_img_name))

        self.button_previous = Button(self.canvas,
            image=self.button_previous_img,
            borderwidth=0,
            highlightthickness=2,
            command=lambda: showFrame(controller.myGanttProperties.frame),
            relief="raised"
        )
        self.button_previous.place(
            x=34.0,
            y=620.0,
            width=150.0,
            height=50.0
        )

        # Github profile link
        self.github_img = PIL.Image.open(self.relative_to_assets(str_githublogo))
        self.github_img = self.github_img.resize((50,50),PIL.Image.ANTIALIAS)
        self.github_img_tk = PIL.ImageTk.PhotoImage(self.github_img)

        self.github_button = Button(self.canvas,
            image=self.github_img_tk,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: openLinks(str_github),
            bg="black"
        )
        self.github_button.place(
            x=730.0,
            y=620.0,
            width=50.0,
            height=50.0
        )

        # LinkedIn profile link
        self.linkedIn_img = PIL.Image.open(self.relative_to_assets(str_linkedInlogo))
        self.linkedIn_img = self.linkedIn_img.resize((60,50),PIL.Image.ANTIALIAS)
        self.linkedIn_img_tk = PIL.ImageTk.PhotoImage(self.linkedIn_img)

        self.linkedIn_button = Button(self.canvas,
            image=self.linkedIn_img_tk,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: openLinks(str_linkedIn),
            bg="black"
        )
        self.linkedIn_button.place(
            x=785.0,
            y=620.0,
            width=60.0,
            height=50.0
        )

        # Twitter profile link
        self.twitter_img = PIL.Image.open(self.relative_to_assets(str_twitterlogo))
        self.twitter_img = self.twitter_img.resize((50,50),PIL.Image.ANTIALIAS)
        self.twitter_img_tk = PIL.ImageTk.PhotoImage(self.twitter_img)

        self.twitter_button = Button(self.canvas,
            image=self.twitter_img_tk,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: openLinks(str_twitter),
            bg="black"
        )
        self.twitter_button.place(
            x=850.0,
            y=620.0,
            width=50.0,
            height=50.0
        )

        # View Presenetation Button
        self.button_view_presentation = Button(
            self.canvas,
            borderwidth=0,
            highlightthickness=0,
            relief="flat",
            text="View Presentation",
            command=lambda: self.presentationViewer(self.entry_presentation_path.get()),
            font=("Roboto",18)
        )
        self.button_view_presentation.place(
            x=367,
            y=380,
            width=250,
            height=50
        )
        # Image signature
        
        self.signature_img = PIL.Image.open(self.relative_to_assets(str_signature))
        self.signature_img = self.signature_img.resize((227,150),PIL.Image.ANTIALIAS)
        self.signature_img_tk = PIL.ImageTk.PhotoImage(self.signature_img)

        self.canvas.create_image(378.5,450.0, anchor="nw",image=self.signature_img_tk)

# Function to return relative path
    def relative_to_assets(self,path: str) -> Path:
        return self.ASSETS_PATH / Path(path)

    def presentationViewer(self,path: str):
        try:
            if platform.system()=="Darwin":
                subprocess.call(("open",path))
            elif platform.system()=="Windows":
                os.startfile(path)
            else:
                subprocess.call(("xdg-open",path))
        except:
            showerror("Save Presentation before viewing","Please click on browse to save presentation then view the presentation")
    
    def selfBrowseSaveButton(self):
        browseSaveButton(self.root,self.entry_presentation_path,
        self.controller.myGanttProperties.final_presentation)

