# Class for setting up the Gantt Properties
from pathlib import Path
from tkinter.ttk import Frame, Button, Combobox, Entry
from tkinter import Tk, Canvas, PhotoImage, colorchooser
import PIL.Image
import PIL.ImageTk
import PIL.ImageOps
from tkinter.messagebox import showerror
from Frontend.commands import *
from Frontend.strings import *
from Backend.Root import *

class GanttPropertiesFrame():
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
            text="Select Gantt Properties:",
            fill="#000000",
            font=("Roboto", 36 * -1)
        )

        # Style preview
        self.gantt_image = PIL.Image.open(self.relative_to_assets(str_gantt_chart_style_1))
        self.gantt_image_resized=self.gantt_image.resize((600,337),PIL.Image.ANTIALIAS)
        self.gantt_image_bordered=PIL.ImageOps.expand(self.gantt_image_resized,border=(2,2,2,2),fill="black")
        self.gantt_image_new = PIL.ImageTk.PhotoImage(self.gantt_image_bordered)
        self.canvas.create_image(320,80,anchor="nw",image=self.gantt_image_new)

        # Fill, font color for all params
        self.font_properties = ("Roboto",12 * -1)
        self.fill_color = "#000000"
        # Timeline Table Properties- Left of reference image
        self.canvas.create_text(68,80,anchor="nw",text="Timeline Properties",fill="#008080",font=("Roboto",14 * -1,))
        self.canvas.create_text(68,110,anchor="nw",text="Timeline Left\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,140,anchor="nw",text="Timeline Top\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,170,anchor="nw",text="Timeline Width\t:",fill=self.fill_color,font=self.font_properties)

        # Entryboxes and Option menu for entering timeline properties
        self.units_list=['inch','cm']
        
        # Timeline Left
        self.entry_left = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(190,108,anchor="nw",window=self.entry_left)
        self.entry_left.insert(0,"2.0")
        
        self.option_left_selection = StringVar()
        self.option_left = Combobox(self.canvas,textvariable= self.option_left_selection,value=self.units_list, width=5)
        self.option_left.current(0)
        self.option_left.state(["readonly"])
        self.canvas.create_window(250,108,anchor="nw",window=self.option_left)
        
        # Timeline Top
        self.entry_top = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(190,138,anchor="nw",window=self.entry_top)
        self.entry_top.insert(0,"1.0")
        
        self.option_top_selection = StringVar()
        self.option_top = Combobox(self.canvas,textvariable= self.option_top_selection,value=self.units_list, width=5)
        self.option_top.current(0)
        self.option_top.state(["readonly"])
        self.canvas.create_window(250,138,anchor="nw",window=self.option_top)
        
        # Timeline width
        self.entry_width = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(190,168,anchor="nw",window=self.entry_width)
        self.entry_width.insert(0,"10.0")
        
        self.option_width_selection = StringVar()
        self.option_width = Combobox(self.canvas,textvariable= self.option_width_selection,value=self.units_list, width=5)
        self.option_width.current(0)
        self.option_width.state(["readonly"])
        self.canvas.create_window(250,168,anchor="nw",window=self.option_width)

        # Task alignment with respect to shape - Left of reference image 
        self.canvas.create_text(68,220,anchor="nw",text="Task Level Alignment",fill="#008080",font=("Roboto",14 * -1,))
        self.canvas.create_text(68,250,anchor="nw",text="Task Level 1\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,280,anchor="nw",text="Task Level 2\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,310,anchor="nw",text="Task Level 3\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,340,anchor="nw",text="Task Level 4\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,370,anchor="nw",text="Task Level 5\t:",fill=self.fill_color,font=self.font_properties)
        
        self.create_alignment_selectors()
        
        # Task Text and shape properties - bottom of the reference image
        self.canvas.create_text(68,430,anchor="nw",text="Task Level\t",fill="#008080",font=("Roboto",14 * -1,))
        self.canvas.create_text(68,455,anchor="nw",text="Task Level 1\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,485,anchor="nw",text="Task Level 2\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,515,anchor="nw",text="Task Level 3\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,545,anchor="nw",text="Task Level 4\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,575,anchor="nw",text="Task Level 5\t:",fill=self.fill_color,font=self.font_properties)
        
        # Task Text and shape properties - bottom of the reference image headers
        self.canvas.create_text(190,430,anchor="nw",text="Shape Type\t",fill="#800000",font=("Roboto",14 * -1,))
        self.canvas.create_text(320,430,anchor="nw",text="Shape Color\t",fill="#800000",font=("Roboto",14 * -1,))
        self.canvas.create_text(450,430,anchor="nw",text="Font Properties\t",fill="#800000",font=("Roboto",14 * -1,))
        self.canvas.create_text(580,430,anchor="nw",text="Font Style\t",fill="#800000",font=("Roboto",14 * -1,))
        self.canvas.create_text(710,430,anchor="nw",text="Font Color\t",fill="#800000",font=("Roboto",14 * -1,))
        self.canvas.create_text(840,430,anchor="nw",text="Font Size\t",fill="#800000",font=("Roboto",14 * -1,))
        
        self.create_shape_type_selectors()
        self.create_shape_color_selectors()
        self.create_font_color_selectors()
        self.create_font_size_selectors()
        self.create_font_property_selectors()
        self.create_font_style_selectors()

        # Set default shape color
        self.str_task1_shape_color="#5D5C5B"
        self.str_task2_shape_color="#045D5D"
        self.str_task3_shape_color="#689E9E"
        self.str_task4_shape_color="#000000"
        self.str_task5_shape_color="#C22B2E"

        # Set default font color
        self.str_task1_font_color="#000000"
        self.str_task2_font_color="#000000"
        self.str_task3_font_color="#000000"
        self.str_task4_font_color="#000000"
        self.str_task5_font_color="#000000"

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
            command=lambda: showFrame(controller.myDataSetup.frame),
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
    
    # Combo boxes for alignment
    def create_alignment_selectors(self):
        # Comboboxes for alignment for each task level
        self.alignment_list = [
            "Left of Slide",
            "Right of Slide",
            "Left of Shape",
            "On the Shape",
            "Right of Shape",
            "Bottom of Shape"
        ]

        # Task 1 alignment
        self.alignment_task_1_selection = StringVar()
        self.alignment_task_1 = Combobox(self.canvas,textvariable= self.alignment_task_1_selection, value=self.alignment_list, width= 15)
        self.alignment_task_1.current(0)
        self.alignment_task_1.state(['readonly'])
        self.canvas.create_window(190,248,anchor="nw",window=self.alignment_task_1)

        # Task 2 alignment
        self.alignment_task_2_selection = StringVar()
        self.alignment_task_2 = Combobox(self.canvas,textvariable= self.alignment_task_2_selection, value=self.alignment_list, width= 15)
        self.alignment_task_2.current(0)
        self.alignment_task_2.state(['readonly'])
        self.canvas.create_window(190,278,anchor="nw",window=self.alignment_task_2)

        # Task 3 alignment
        self.alignment_task_3_selection = StringVar()
        self.alignment_task_3 = Combobox(self.canvas,textvariable= self.alignment_task_3_selection, value=self.alignment_list, width= 15)
        self.alignment_task_3.current(0)
        self.alignment_task_3.state(['readonly'])
        self.canvas.create_window(190,308,anchor="nw",window=self.alignment_task_3)

        # Task 4 alignment
        self.alignment_task_4_selection = StringVar()
        self.alignment_task_4 = Combobox(self.canvas,textvariable= self.alignment_task_4_selection, value=self.alignment_list, width= 15)
        self.alignment_task_4.current(5)
        self.alignment_task_4.state(['readonly'])
        self.canvas.create_window(190,338,anchor="nw",window=self.alignment_task_4)

        # Task 5 alignment
        self.alignment_task_5_selection = StringVar()
        self.alignment_task_5 = Combobox(self.canvas,textvariable= self.alignment_task_5_selection, value=self.alignment_list, width= 15)
        self.alignment_task_5.current(5)
        self.alignment_task_5.state(['readonly'])
        self.canvas.create_window(190,368,anchor="nw",window=self.alignment_task_5)

    # Combo boxes for shape type
    def create_shape_type_selectors(self):
        # Task shape type combo boxes
        self.main_shapes_list=[
            "Chevron",
            "Pentagon",
            "Rectangle"
        ]
        # Task 1 shape type selection
        self.shape_type_task_1_selection = StringVar()
        self.shape_type_task_1 = Combobox(self.canvas,textvariable= self.shape_type_task_1_selection, value=self.main_shapes_list, width= 15)
        self.shape_type_task_1.current(1)
        self.shape_type_task_1.state(['readonly'])
        self.canvas.create_window(190,455,anchor="nw",window=self.shape_type_task_1)
        # Task 2 shape type selection
        self.shape_type_task_2_selection = StringVar()
        self.shape_type_task_2 = Combobox(self.canvas,textvariable= self.shape_type_task_2_selection, value=self.main_shapes_list, width= 15)
        self.shape_type_task_2.current(1)
        self.shape_type_task_2.state(['readonly'])
        self.canvas.create_window(190,485,anchor="nw",window=self.shape_type_task_2)
        # Task 3 shape type selection
        self.shape_type_task_3_selection = StringVar()
        self.shape_type_task_3 = Combobox(self.canvas,textvariable= self.shape_type_task_3_selection, value=self.main_shapes_list, width= 15)
        self.shape_type_task_3.current(1)
        self.shape_type_task_3.state(['readonly'])
        self.canvas.create_window(190,515,anchor="nw",window=self.shape_type_task_3)

        self.secondary_shapes_list=[
            "Diamond",
            "Star",
            "Circle",
            "Triangle",
            "Square"
        ]
        # Task 4 shape type selection
        self.shape_type_task_4_selection = StringVar()
        self.shape_type_task_4 = Combobox(self.canvas,textvariable= self.shape_type_task_4_selection, value=self.secondary_shapes_list, width= 15)
        self.shape_type_task_4.current(0)
        self.shape_type_task_4.state(['readonly'])
        self.canvas.create_window(190,545,anchor="nw",window=self.shape_type_task_4)
        # Task 5 shape type selection
        self.shape_type_task_5_selection = StringVar()
        self.shape_type_task_5 = Combobox(self.canvas,textvariable= self.shape_type_task_5_selection, value=self.secondary_shapes_list, width= 15)
        self.shape_type_task_5.current(0)
        self.shape_type_task_5.state(['readonly'])
        self.canvas.create_window(190,575,anchor="nw",window=self.shape_type_task_5)
    
    # Create shape color selectors
    def create_shape_color_selectors(self):
        # Task 1
        self.colorchooser_button_task_1 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task1_shape_color(self.colorchooser_button_task_1), font=self.font_properties)
        self.colorchooser_button_task_1.place(x=320,y=455,height=20,width=100)
        # Task 2
        self.colorchooser_button_task_2 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task2_shape_color(self.colorchooser_button_task_2), font=self.font_properties)
        self.colorchooser_button_task_2.place(x=320,y=485,height=20,width=100)
        # Task 3
        self.colorchooser_button_task_3 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task3_shape_color(self.colorchooser_button_task_3), font=self.font_properties)
        self.colorchooser_button_task_3.place(x=320,y=515,height=20,width=100)
        # Task 4
        self.colorchooser_button_task_4 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task4_shape_color(self.colorchooser_button_task_4), font=self.font_properties)
        self.colorchooser_button_task_4.place(x=320,y=545,height=20,width=100)
        # Task 5
        self.colorchooser_button_task_5 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task5_shape_color(self.colorchooser_button_task_5), font=self.font_properties)
        self.colorchooser_button_task_5.place(x=320,y=575,height=20,width=100)
    
     # Create font color selectors
    def create_font_color_selectors(self):
        # Task 1
        self.font_colorchooser_button_task_1 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task1_font_color(self.font_colorchooser_button_task_1), font=self.font_properties)
        self.font_colorchooser_button_task_1.place(x=710,y=455,height=20,width=100)
        # Task 2
        self.font_colorchooser_button_task_2 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task2_font_color(self.font_colorchooser_button_task_2), font=self.font_properties)
        self.font_colorchooser_button_task_2.place(x=710,y=485,height=20,width=100)
        # Task 3
        self.font_colorchooser_button_task_3 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task3_font_color(self.font_colorchooser_button_task_3), font=self.font_properties)
        self.font_colorchooser_button_task_3.place(x=710,y=515,height=20,width=100)
        # Task 4
        self.font_colorchooser_button_task_4 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task4_font_color(self.font_colorchooser_button_task_4), font=self.font_properties)
        self.font_colorchooser_button_task_4.place(x=710,y=545,height=20,width=100)
        # Task 5
        self.font_colorchooser_button_task_5 = Button(self.canvas,text="Choose Color",borderwidth=2,
        highlightthickness=2,relief="raised", command=lambda: self.task5_font_color(self.font_colorchooser_button_task_5), font=self.font_properties)
        self.font_colorchooser_button_task_5.place(x=710,y=575,height=20,width=100)
    
    # Function which activates color picker background
    def pick_bg_color(self,button: Button):
        selected_color = colorchooser.askcolor()[1]
        button.config(bg=selected_color)
        return selected_color
    # Task1 Shape color
    def task1_shape_color(self,button: Button):
        self.str_task1_shape_color = self.pick_bg_color(button)
    # task2 Shape color
    def task2_shape_color(self,button: Button):
        self.str_task2_shape_color = self.pick_bg_color(button)
    # task3 Shape color
    def task3_shape_color(self,button: Button):
        self.str_task3_shape_color = self.pick_bg_color(button)
    # task4 Shape color
    def task4_shape_color(self,button: Button):
        self.str_task4_shape_color = self.pick_bg_color(button)
    # task5 Shape color
    def task5_shape_color(self,button: Button):
        self.str_task5_shape_color = self.pick_bg_color(button)

    # Function which activates color picker foreground
    def pick_fg_color(self,button: Button):
        selected_color = colorchooser.askcolor()[1]
        button.config(fg=selected_color)
        return selected_color
    # task1 font color
    def task1_font_color(self,button: Button):
        self.str_task1_font_color = self.pick_fg_color(button)
    # task2 font color
    def task2_font_color(self,button: Button):
        self.str_task2_font_color = self.pick_fg_color(button)
    # task3 font color
    def task3_font_color(self,button: Button):
        self.str_task3_font_color = self.pick_fg_color(button)
    # task4 font color
    def task4_font_color(self,button: Button):
        self.str_task4_font_color = self.pick_fg_color(button)
    # task5 font color
    def task5_font_color(self,button: Button):
        self.str_task5_font_color = self.pick_fg_color(button)
    
    # Create font size selectors
    def create_font_size_selectors(self):
        # Task 1
        self.entry_font_size_task_1 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(850,455,anchor="nw",window=self.entry_font_size_task_1)
        self.entry_font_size_task_1.insert(0,"11")
        # Task 2
        self.entry_font_size_task_2 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(850,485,anchor="nw",window=self.entry_font_size_task_2)
        self.entry_font_size_task_2.insert(0,"11")
        # Task 3
        self.entry_font_size_task_3 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(850,515,anchor="nw",window=self.entry_font_size_task_3)
        self.entry_font_size_task_3.insert(0,"11")
        # Task 4
        self.entry_font_size_task_4 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(850,545,anchor="nw",window=self.entry_font_size_task_4)
        self.entry_font_size_task_4.insert(0,"10")
        # Task 5
        self.entry_font_size_task_5 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(850,575,anchor="nw",window=self.entry_font_size_task_5)
        self.entry_font_size_task_5.insert(0,"10")
    
    def create_font_style_selectors(self):
        # Task 1
        self.entry_font_name_task_1 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=15,font=("RobotoRoman",10))
        self.canvas.create_window(580,455,anchor="nw",window=self.entry_font_name_task_1)
        self.entry_font_name_task_1.insert(0,"Arial")
        # Task 2
        self.entry_font_name_task_2 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=15,font=("RobotoRoman",10))
        self.canvas.create_window(580,485,anchor="nw",window=self.entry_font_name_task_2)
        self.entry_font_name_task_2.insert(0,"Arial")
        # Task 3
        self.entry_font_name_task_3 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=15,font=("RobotoRoman",10))
        self.canvas.create_window(580,515,anchor="nw",window=self.entry_font_name_task_3)
        self.entry_font_name_task_3.insert(0,"Arial")
        # Task 4
        self.entry_font_name_task_4 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=15,font=("RobotoRoman",10))
        self.canvas.create_window(580,545,anchor="nw",window=self.entry_font_name_task_4)
        self.entry_font_name_task_4.insert(0,"Arial")
        # Task 5
        self.entry_font_name_task_5 = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=15,font=("RobotoRoman",10))
        self.canvas.create_window(580,575,anchor="nw",window=self.entry_font_name_task_5)
        self.entry_font_name_task_5.insert(0,"Arial")

    def create_font_property_selectors(self):
        self.font_property_list=[
            "Regular",
            "Bold",
            "Italic",
            "Underline",
            "Bold Italic",
            "Bold Underline",
            "Underline Italic"
        ]
        # Task 1 font property
        self.font_property_task_1_selection = StringVar()
        self.font_property_task_1 = Combobox(self.canvas,textvariable= self.font_property_task_1_selection, value=self.font_property_list, width= 15)
        self.font_property_task_1.current(0)
        self.font_property_task_1.state(['readonly'])
        self.canvas.create_window(450,455,anchor="nw",window=self.font_property_task_1)
        # Task 2 font property
        self.font_property_task_2_selection = StringVar()
        self.font_property_task_2 = Combobox(self.canvas,textvariable= self.font_property_task_2_selection, value=self.font_property_list, width= 15)
        self.font_property_task_2.current(0)
        self.font_property_task_2.state(['readonly'])
        self.canvas.create_window(450,485,anchor="nw",window=self.font_property_task_2)
        # Task 3 font property
        self.font_property_task_3_selection = StringVar()
        self.font_property_task_3 = Combobox(self.canvas,textvariable= self.font_property_task_3_selection, value=self.font_property_list, width= 15)
        self.font_property_task_3.current(0)
        self.font_property_task_3.state(['readonly'])
        self.canvas.create_window(450,515,anchor="nw",window=self.font_property_task_3)
        # Task 4 font property
        self.font_property_task_4_selection = StringVar()
        self.font_property_task_4 = Combobox(self.canvas,textvariable= self.font_property_task_4_selection, value=self.font_property_list, width= 15)
        self.font_property_task_4.current(0)
        self.font_property_task_4.state(['readonly'])
        self.canvas.create_window(450,545,anchor="nw",window=self.font_property_task_4)
        # Task 5 font property
        self.font_property_task_5_selection = StringVar()
        self.font_property_task_5 = Combobox(self.canvas,textvariable= self.font_property_task_5_selection, value=self.font_property_list, width= 15)
        self.font_property_task_5.current(0)
        self.font_property_task_5.state(['readonly'])
        self.canvas.create_window(450,575,anchor="nw",window=self.font_property_task_5)
    
    # Creates the slides and runs all the presentation making in the background and switches to save screen
    def self_next_frame(self,data=None):
        # Finalizing Inputs for creating the gantt chart
        self.final_data = self.controller.myDataSetup.final_data
        self.gantt_prs = self.controller.myStartScreen.ganttPrs
        self.align_tl = [
            self.alignment_task_1_selection.get(),
            self.alignment_task_2_selection.get(),
            self.alignment_task_3_selection.get(),
            self.alignment_task_4_selection.get(),
            self.alignment_task_5_selection.get()
        ]
        self.shape_type_tl = [
            self.shape_type_task_1_selection.get(),
            self.shape_type_task_2_selection.get(),
            self.shape_type_task_3_selection.get(),
            self.shape_type_task_4_selection.get(),
            self.shape_type_task_5_selection.get()
        ]
        self.shape_color_tl = [
            self.str_task1_shape_color,
            self.str_task2_shape_color,
            self.str_task3_shape_color,
            self.str_task4_shape_color,
            self.str_task5_shape_color
        ]
        self.font_prop_tl = [
            self.font_property_task_1_selection.get(),
            self.font_property_task_2_selection.get(),
            self.font_property_task_3_selection.get(),
            self.font_property_task_4_selection.get(),
            self.font_property_task_5_selection.get()
        ]
        self.font_style_tl = [
            self.entry_font_size_task_1.get(),
            self.entry_font_size_task_2.get(),
            self.entry_font_size_task_3.get(),
            self.entry_font_size_task_4.get(),
            self.entry_font_size_task_5.get()
        ]
        self.font_color_tl = [
            self.str_task1_font_color,
            self.str_task2_font_color,
            self.str_task3_font_color,
            self.str_task4_font_color,
            self.str_task5_font_color
        ]
        self.font_size_tl = [
            self.entry_font_size_task_1.get(),
            self.entry_font_size_task_2.get(),
            self.entry_font_size_task_3.get(),
            self.entry_font_size_task_4.get(),
            self.entry_font_size_task_5.get()
        ]
        # Creating the Gantt Chart with finalized inputs
        self.final_presentation = create_gantt_slides(
            prs=self.gantt_prs,
            gantt_data=self.final_data,
            table_left=self.entry_left.get(),
            table_top=self.entry_top.get(),
            table_width=self.entry_width.get(),
            font_color_tl=self.font_color_tl,
            font_size_tl=self.font_size_tl,
            font_prop_tl=self.font_prop_tl,
            font_style_tl=self.font_style_tl,
            shape_type_tl=self.shape_type_tl,
            shape_color_tl=self.shape_color_tl,
            align_tl=self.align_tl,
            table_left_metric=self.option_left_selection.get(),
            table_top_metric=self.option_top_selection.get(),
            table_width_metric=self.option_width_selection.get()
        )
        showFrame(self.controller.mySaveScreen.frame)