# Class for setting up the imported data
from pathlib import Path
from tkinter.ttk import Frame, Button, Combobox, Treeview
from tkinter import Tk, Canvas, PhotoImage, Scrollbar
from tkinter.messagebox import showerror
from Frontend.commands import *
from Frontend.strings import *
from Backend.Root import *


# When class is initialized, it creates a frame and canvas with relevant widgets
class DataSetupFrame:
    def __init__(self,root: Tk,controller):
        super().__init__()
        
        # Locates the assests folder and stores it in ASSETS_PATH
        self.OUTPUT_PATH = Path(__file__).parent
        self.ASSETS_PATH = self.OUTPUT_PATH / Path(str_assets)
        
        # Initialises the root and controller classes
        self.root = root
        self.controller = controller
        
        # Creates a frame in the root window and places it such that it covers whole window
        self.frame = Frame(self.root)
        self.frame.grid(row=0,column=0,sticky="nsew")
        self.frame.configure(bg = "#FFFFFF")

        # Creates a canvas with desired properties
        self.canvas = Canvas(
            self.frame,
            bg = "#FFFFFF",
            height = 700,
            width = 984,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )

        # Title for the frame
        self.canvas.place(x = 0, y = 0)
        self.canvas.create_text(
            68.0,
            30.0,
            anchor="nw",
            text="Assign Column Names:",
            fill="#000000",
            font=("Roboto", 36 * -1)
        )

        # Headers for the data setup
        self.canvas.create_text(
            485.0,
            40.0,
            anchor="nw",
            text="Slide Title\t:",
            fill="#000000",
            font=("Roboto", 24 * -1)
        )
        
        self.canvas.create_text(
            68.0,
            99.0,
            anchor="nw",
            text="Task Name\t:",
            fill="#000000",
            font=("Roboto", 24 * -1)
        )
                
        self.canvas.create_text(
            485.0,
            99.0,
            anchor="nw",
            text="Task Duration\t:",
            fill="#000000",
            font=("Roboto", 24 * -1)
        )

        self.canvas.create_text(
            68.0,
            184.0,
            anchor="nw",
            text="Task Level\t:",
            fill="#000000",
            font=("Roboto", 24 * -1)
        )

        self.canvas.create_text(
            485.0,
            184.0,
            anchor="nw",
            text="Gantt Start Date\t:",
            fill="#000000",
            font=("Roboto", 24 * -1)
        )

        self.canvas.create_text(
            68.0,
            269.0,
            anchor="nw",
            text="Task Start Date\t:",
            fill="#000000",
            font=("Roboto", 24 * -1)
        )

        self.canvas.create_text(
            485.0,
            269.0,
            anchor="nw",
            text="Gantt Duration\t:",
            fill="#000000",
            font=("Roboto", 24 * -1)
        )
        
        # Button to go to next frame
        self.button_next_img = PhotoImage(
            file=self.relative_to_assets(str_next_button_img_name))

        self.button_next = Button(self.canvas,
            image=self.button_next_img,
            borderwidth=0,
            highlightthickness=2,
            command=lambda: self.self_next_frame(),
            relief="raised"
        )
        self.button_next.place(
            x=800.0,
            y=620.0,
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
            command=lambda: showFrame(controller.myStartScreen.frame),
            relief="raised"
        )
        self.button_previous.place(
            x=34.0,
            y=620.0,
            width=150.0,
            height=50.0
        )

# Function to return relative path    
    def relative_to_assets(self,path: str) -> Path:
        return self.ASSETS_PATH / Path(path)

# Creates the drop down selection to assign column names    
    def drop_down_columns(self,combobox_values_list,ganttData):
        
        self.data_preview(ganttData)
        
        # Task name dropdown list
        self.comboBox_task_name_selected = StringVar()
        self.comboBox_task_name = Combobox(
            self.canvas,
            value=combobox_values_list,
            textvariable=self.comboBox_task_name_selected,
            width=30
        )
        self.comboBox_task_name.state(['readonly'])
        try:
            self.comboBox_task_name.current(0)
        except:
            showerror("Data Not Selected","Please select data in the previous page")
        self.canvas.create_window(260,102,anchor="nw",window=self.comboBox_task_name)

        # Task duration dropdown list
        self.comboBox_task_duration_selected = StringVar()
        self.comboBox_task_duration = Combobox(
            self.canvas,
            value=combobox_values_list,
            textvariable=self.comboBox_task_duration_selected,
            width=30
        )
        self.comboBox_task_duration.state(['readonly'])
        try:
            self.comboBox_task_duration.current(1)
        except:
            showerror("Data Error","Please select data in the previous page or make sure your data has atleast 6 columns")
        self.canvas.create_window(670,102,anchor="nw",window=self.comboBox_task_duration)

        # Task level dropdown list
        self.comboBox_task_level_selected = StringVar()
        self.comboBox_task_level = Combobox(
            self.canvas,
            value=combobox_values_list,
            textvariable=self.comboBox_task_level_selected,
            width=30
        )
        self.comboBox_task_level.state(['readonly'])
        try:
            self.comboBox_task_level.current(2)
        except:
            showerror("Data Error","Please select data in the previous page or make sure your data has atleast 6 columns")
        self.canvas.create_window(260,188,anchor="nw",window=self.comboBox_task_level)

        # Gantt Start date dropdown list
        self.comboBox_gantt_start_date_selected = StringVar()
        self.comboBox_gantt_start_date = Combobox(
            self.canvas,
            value=combobox_values_list,
            textvariable=self.comboBox_gantt_start_date_selected,
            width=30
        )
        self.comboBox_gantt_start_date.state(['readonly'])
        try:
            self.comboBox_gantt_start_date.current(3)
        except:
            showerror("Data Error","Please select data in the previous page or make sure your data has atleast 6 columns")
        self.canvas.create_window(670,188,anchor="nw",window=self.comboBox_gantt_start_date)

        # Task start date dropdown list
        self.comboBox_task_start_date_selected = StringVar()
        self.comboBox_task_start_date = Combobox(
            self.canvas,
            value=combobox_values_list,
            textvariable=self.comboBox_task_start_date_selected,
            width=30
        )
        self.comboBox_task_start_date.state(['readonly'])
        try:
            self.comboBox_task_start_date.current(4)
        except:
            showerror("Data Error","Please select data in the previous page or make sure your data has atleast 6 columns")
        self.canvas.create_window(260,273,anchor="nw",window=self.comboBox_task_start_date)

        # Gantt duration dropdown list
        self.comboBox_gantt_duration_selected = StringVar()
        self.comboBox_gantt_duration = Combobox(
            self.canvas,
            value=combobox_values_list,
            textvariable=self.comboBox_gantt_duration_selected,
            width=30
        )
        self.comboBox_gantt_duration.state(['readonly'])
        try:
            self.comboBox_gantt_duration.current(5)
        except:
            showerror("Data Error","Please select data in the previous page or make sure your data has atleast 6 columns")
        self.canvas.create_window(670,273,anchor="nw",window=self.comboBox_gantt_duration)

        # Gantt Title dropdown list
        self.comboBox_slide_title_selected = StringVar()
        self.comboBox_slide_title = Combobox(
            self.canvas,
            value=combobox_values_list,
            textvariable=self.comboBox_slide_title_selected,
            width=30
        )
        self.comboBox_slide_title.state(['readonly'])
        try:
            self.comboBox_slide_title.current(6)
        except:
            showerror("Data Error","Please select data in the previous page or make sure your data has atleast 6 columns")
        self.canvas.create_window(670,43,anchor="nw",window=self.comboBox_slide_title)

# Creates treeview with horizontal and vertical scroll bars
    def data_preview(self,data):
        self.tree_frame = Frame(self.frame,height=200)
        
        self.tree_x_scroll = Scrollbar(self.tree_frame,orient="horizontal")
        self.tree_x_scroll.pack(side=BOTTOM,fill=X)
        self.tree_y_scroll = Scrollbar(self.tree_frame)
        self.tree_y_scroll.pack(side=RIGHT,fill=Y)

        self.data_treeview = Treeview(
            self.tree_frame,
            yscrollcommand=self.tree_y_scroll.set,
            selectmode="extended",
            xscrollcommand=self.tree_x_scroll.set,
        )

        self.tree_x_scroll.config(command=self.data_treeview.xview)
        self.tree_y_scroll.config(command=self.data_treeview.yview)
        
        self.data_treeview['column']=list(data.columns)
        self.data_treeview['show']="headings"
        self.canvas.create_window(50,310,anchor="nw",window=self.tree_frame)        
        column_width = int(884/len(data.columns))
        self.show_table(column_width,data)

# Using pandas dataframe creates tabular visual on screen
    def show_table(self,column_width,data):
        for col in self.data_treeview['column']:
            self.data_treeview.column(col,width=126,stretch=NO)
            self.data_treeview.heading(col,text=col)
        data_rows = data.to_numpy().tolist()
        for row in data_rows:
            self.data_treeview.insert("","end",values=row)
        self.data_treeview.pack()
        for col in self.data_treeview['column']:
            self.data_treeview.column(col,width=column_width,stretch=YES,minwidth=50)
        
# Setups the data for further processing and switches to gantt properties frame
    def self_next_frame(self):
        showFrame(self.controller.myTimelineProperties.frame)
        data = self.controller.myStartScreen.ganttData
        self.final_data=data_setup_for_gantt(
            data,
            task_name=self.comboBox_task_name_selected.get(),
            task_duration=self.comboBox_task_duration_selected.get(),
            task_level=self.comboBox_task_level_selected.get(),
            task_start_date=self.comboBox_task_start_date_selected.get(),
            gantt_start_date=self.comboBox_gantt_start_date_selected.get(),
            gantt_duration=self.comboBox_gantt_duration_selected.get(),
            slide_title=self.comboBox_slide_title_selected.get()
        )