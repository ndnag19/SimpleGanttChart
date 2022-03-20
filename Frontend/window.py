# Class which creates the whole application

from tkinter import *
from Frontend.object_save_screen import SaveScreenFrame
from Frontend.object_start_screen import *
from Frontend.object_data_setup import *
from Frontend.object_gantt_properties import *
from Frontend.object_timeline_properties import *
from Frontend.strings import *
from Frontend.commands import *
from Frontend.object_start_screen import *

class MainWindow:
    def __init__(self,font_list: list):
        super().__init__()
        # Locates the assests folder and stores it in ASSETS_PATH
        self.OUTPUT_PATH = Path(__file__).parent
        self.ASSETS_PATH = self.OUTPUT_PATH / Path(str_assets)

        # Creating a window using Tkinter
        self.root = Tk()
        self.root.wm_title(str_project_title)
        self.root.iconbitmap(self.relative_to_assets("icon_white.ico"))
        self.root.geometry("984x700")
        self.root.rowconfigure(0,weight=1)
        self.root.columnconfigure(0,weight=1)

        #Adding Frames to the window
        self.myStartScreen = StartScreenFrame(self.root,controller=self)
        self.myDataSetup = DataSetupFrame(self.root,controller=self)
        self.myTimelineProperties = TimelinePropertiesFrame(self.root,controller=self)
        self.myGanttProperties = GanttPropertiesFrame(self.root,controller=self,font_list=font_list)
        self.mySaveScreen = SaveScreenFrame(self.root,controller=self)
        # Setting Start Screen as the first framee
        showFrame(self.myStartScreen.frame)

        # Making the window not resizable
        self.root.resizable(False,False)

    def start_window(self):
        self.root.mainloop()
    
    def relative_to_assets(self,path: str) -> Path:
        return self.ASSETS_PATH / Path(path)