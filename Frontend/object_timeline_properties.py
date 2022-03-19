# Creates Timeline Properties Frame with all the table properties
from pathlib import Path
from tkinter import *
from tkinter import Tk, Canvas, Frame, colorchooser, PhotoImage, Entry, Label
from tkinter.messagebox import showinfo,showerror,showwarning
from tkinter import ttk
from font_chooser import FontDialog
from ttkwidgets.frames import Tooltip
from pptx import Presentation
import PIL.Image
import PIL.ImageTk
import PIL.ImageOps
from strings import *
from commands import showFrame,slideLayoutList

class TimelinePropertiesFrame():
    def __init__(self, root: Tk,controller,slide_layout_list):
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
            text="Select Timeline Properties:",
            fill="#000000",
            font=("Roboto", 36 * -1)
        )

        # Timeline Preview
        self.timeline_image = PIL.Image.open(self.relative_to_assets(str_timeline_style_1))
        self.timeline_image_resized=self.timeline_image.resize((848,91),PIL.Image.ANTIALIAS)
        self.timeline_image_bordered=PIL.ImageOps.expand(self.timeline_image_resized,border=(1,1,1,1),fill="black")
        self.timeline_image_new = PIL.ImageTk.PhotoImage(self.timeline_image_bordered)
        self.canvas.create_image(68,80,anchor="nw",image=self.timeline_image_new)

        # Fill, font color for all params
        self.font_properties = ("Roboto",12 * -1)
        self.fill_color = "#000000"
        
        # Timeline Table Properties- Left of reference image
        self.canvas.create_text(68,200,anchor="nw",text="Timeline Left\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,230,anchor="nw",text="Timeline Top\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,260,anchor="nw",text="Timeline Width\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,290,anchor="nw",text="Timeline Height\t:",fill=self.fill_color,font=self.font_properties)

        # Entryboxes and Option menu for entering timeline properties
        self.units_list=['inch','cm']
        
        # Timeline Left
        self.entry_left = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(190,198,anchor="nw",window=self.entry_left)
        self.entry_left.insert(0,"2.0")
        
        self.option_left_selection = StringVar()
        self.option_left = ttk.Combobox(self.canvas,textvariable= self.option_left_selection,value=self.units_list, width=5)
        self.option_left.current(0)
        self.option_left.state(["readonly"])
        self.canvas.create_window(250,198,anchor="nw",window=self.option_left)
        
        # Timeline Top
        self.entry_top = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(190,228,anchor="nw",window=self.entry_top)
        self.entry_top.insert(0,"1.0")
        
        self.option_top_selection = StringVar()
        self.option_top = ttk.Combobox(self.canvas,textvariable= self.option_top_selection,value=self.units_list, width=5)
        self.option_top.current(0)
        self.option_top.state(["readonly"])
        self.canvas.create_window(250,228,anchor="nw",window=self.option_top)
        
        # Timeline width
        self.entry_width = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(190,258,anchor="nw",window=self.entry_width)
        self.entry_width.insert(0,"10.0")
        
        self.option_width_selection = StringVar()
        self.option_width = ttk.Combobox(self.canvas,textvariable= self.option_width_selection,value=self.units_list, width=5)
        self.option_width.current(0)
        self.option_width.state(["readonly"])
        self.canvas.create_window(250,258,anchor="nw",window=self.option_width)

        # Timeline height
        self.entry_height = Entry(self.frame, bd=0, bg="#ECECEC", highlightthickness=2,width=5,font=("RobotoRoman",10))
        self.canvas.create_window(190,288,anchor="nw",window=self.entry_height)
        self.entry_height.insert(0,"0.50")
        
        self.option_height_selection = StringVar()
        self.option_height = ttk.Combobox(self.canvas,textvariable= self.option_height_selection,value=self.units_list, width=5)
        self.option_height.current(0)
        self.option_height.state(["readonly"])
        self.canvas.create_window(250,288,anchor="nw",window=self.option_height)

        # Labels for Rows, Timeline Milestones and Title Properties
        self.canvas.create_text(445,180,anchor="nw",text="Slide Title\t",fill="#008080",font=("Roboto",14 * -1,))
        self.canvas.create_text(555,180,anchor="nw",text="Row 1\t",fill="#008080",font=("Roboto",14 * -1,))
        self.canvas.create_text(665,180,anchor="nw",text="Row 2\t",fill="#008080",font=("Roboto",14 * -1,))
        self.canvas.create_text(775,180,anchor="nw",text="Milestone\t",fill="#008080",font=("Roboto",14 * -1,))
        self.canvas.create_text(325,200,anchor="nw",text="Font Properties\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(325,230,anchor="nw",text="Font Color     \t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(325,260,anchor="nw",text="Fill Color     \t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(325,290,anchor="nw",text="Text Align     \t:",fill=self.fill_color,font=self.font_properties)
        
        # Title Fonts
        self.title_font_button = ttk.Button(self.canvas, text="Title", command= lambda: self.title_font_chooser(), padding=-1)
        self.title_font_button.place(x=445, y=200, height=20, width=100)

        # Font Color
        self.title_font_color_button = Button(self.canvas,text="Title", command= lambda: self.title_font_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6")
        self.title_font_color_button.place(x=445, y=230, height=20, width=100)

        # Fill Color
        self.title_fill_color_button = Button(self.canvas, text="Title", command= lambda: self.title_fill_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6")
        self.title_fill_color_button.place(x=445, y=260, height=20, width=100)

        # Text Align Combo Boxes
        self.text_align_list = [
            "Left",
            "Right",
            "Center"
        ]

        self.title_text_align_selection = StringVar()
        self.title_text_align = ttk.Combobox(self.canvas,textvariable=self.title_text_align_selection,values=self.text_align_list,width=13)
        self.title_text_align.current(0)
        self.title_text_align.state(['readonly'])
        self.canvas.create_window(445,288,anchor="nw",window=self.title_text_align)
        
        # Create Timeline CheckButton
        self.create_timeline_checkbutton_selection = BooleanVar()
        self.create_timeline_checkbutton = ttk.Checkbutton(self.canvas,command=self.timeline_checkbutton_command,
        text="Check to create timeline",variable=self.create_timeline_checkbutton_selection, width=65)
        self.canvas.create_window(68,320,anchor="nw",window=self.create_timeline_checkbutton)

        # Create Milestone CheckButton
        self.create_milestone_checkbutton_selection = BooleanVar()
        self.create_milestone_checkbutton = ttk.Checkbutton(self.canvas,command=self.milestone_checkbutton_command,
        text="Check to create milestones in timeline",variable=self.create_milestone_checkbutton_selection,width=65)
        self.canvas.create_window(500,320,anchor="nw",window=self.create_milestone_checkbutton)

        # Create Slide-Layout Dropdown
        self.slide_layout_selection = StringVar()
        self.slide_layout_dropdown = ttk.Combobox(self.canvas,textvariable=self.slide_layout_selection, values=slide_layout_list, width=40)
        self.slide_layout_dropdown.current(6)
        self.slide_layout_dropdown.state(['readonly'])
        self.canvas.create_window(218,348,anchor="nw",window=self.slide_layout_dropdown)

        self.canvas.create_text(68,350,anchor="nw",text="Select Slide Layout\t:",fill="#008080",font=("Roboto",14 * -1,))
        
        # Timeline Granularity Dropdown
        self.timeline_granularity_list = [
            "Month",
            "Quarter",
            "Semi-Annual"
        ]

        self.timeline_granularity_selected = StringVar()
        self.timeline_granularity_dropdown = ttk.Combobox(self.canvas,textvariable=self.timeline_granularity_selected,values=self.timeline_granularity_list,width=30)
        self.timeline_granularity_dropdown.current(0)
        self.timeline_granularity_dropdown.state(['readonly'])
        self.canvas.create_window(710,348,anchor="nw",window=self.timeline_granularity_dropdown)

        self.canvas.create_text(500,350,anchor="nw",text="Select Timeline Granularity\t:",fill="#008080",font=("Roboto",14 * -1,))
        
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
        # Tool tips
        # self.font_choose_tooltip = Tooltip(
        #     self.font_choose_button,
        #     headertext="Select Font Properties:",
        #     text="Click button to select font properties for the object"
        # )
    # Function to package all the collected inputs
    def __selected_inputs(self)-> dict:
        pass

    # Function to set defaults
    def __set_default_values(self):
        pass

    # Function to Go to Next Frame
    def self_next_frame(self):
        self.final_inputs = self.__selected_inputs()
        showFrame(self.controller.myGanttProperties.frame)

    # Function to select font properties
    def row_font_chooser(self):
        list_font_selected = FontDialog()
        show_text = str(list_font_selected[0]) + " | " + str(list_font_selected[1])
        if list_font_selected[2]:
            show_text = show_text + " | Bold"
        if list_font_selected[3]:
            show_text = show_text + " | Italic"
        if list_font_selected[4]:
            show_text = show_text + " | Underline"
        if list_font_selected[5]:
            show_text = show_text + " | Strike"
        showinfo("Selected Font Properties:", show_text)
        return list_font_selected
    
    # Stores Font Properties in a list
    # [Font Name, Font Size, bold, italic, is_underline, is_overstrike]
    def row_1_font_chooser(self):
        self.list_row_1_font_selected = self.row_font_chooser()
    def row_2_font_chooser(self):
        self.list_row_2_font_selected = self.row_font_chooser()
    def title_font_chooser(self):
        self.list_title_font_selected = self.row_font_chooser()
    def milestone_font_chooser(self):
        self.list_milestone_font_selected = self.row_font_chooser()

    # Function to select font color
    def pick_font_color(self,button: Button):
        selected_color = colorchooser.askcolor()[1]
        button.config(fg=selected_color)
        return selected_color
    
    # Funtion to store selected colors
    def row_1_font_color(self):
        self.str_row_1_font_color_selected = self.pick_font_color(self.row_1_font_color_button)
    def row_2_font_color(self):
        self.str_row_2_font_color_selected = self.pick_font_color(self.row_2_font_color_button)
    def title_font_color(self):
        self.str_title_font_color_selected = self.pick_font_color(self.title_font_color_button)
    def milestone_font_color(self):
        self.str_milestone_font_color_selected = self.pick_font_color(self.milestone_font_color_button)
    
    # Function to select fill color
    def pick_fill_color(self,button: Button):
        selected_color = colorchooser.askcolor()[1]
        button.config(bg=selected_color)
        return selected_color
    
    # Funtion to store selected colors
    def row_1_fill_color(self):
        self.str_row_1_fill_color_selected = self.pick_fill_color(self.row_1_fill_color_button)
    def row_2_fill_color(self):
        self.str_row_2_fill_color_selected = self.pick_fill_color(self.row_2_fill_color_button)
    def title_fill_color(self):
        self.str_title_fill_color_selected = self.pick_fill_color(self.title_fill_color_button)
    def milestone_fill_color(self):
        self.str_milestone_fill_color_selected = self.pick_fill_color(self.milestone_fill_color_button)
    def __milestone_1_shape_color(self):
        self.str_shape_color_milestone_1_selected = self.pick_fill_color(self.shape_color_milestone_1)
    def __milestone_2_shape_color(self):
        self.str_shape_color_milestone_2_selected = self.pick_fill_color(self.shape_color_milestone_2)
    def __milestone_3_shape_color(self):
        self.str_shape_color_milestone_3_selected = self.pick_fill_color(self.shape_color_milestone_3)
    def __milestone_4_shape_color(self):
        self.str_shape_color_milestone_4_selected = self.pick_fill_color(self.shape_color_milestone_4)
    def __milestone_5_shape_color(self):
        self.str_shape_color_milestone_5_selected = self.pick_fill_color(self.shape_color_milestone_5)
    def __milestone_1_outline_color(self):
        self.str_outline_color_milestone_1_selected = self.pick_fill_color(self.outline_color_milestone_1)
    def __milestone_2_outline_color(self):
        self.str_outline_color_milestone_2_selected = self.pick_fill_color(self.outline_color_milestone_2)
    def __milestone_3_outline_color(self):
        self.str_outline_color_milestone_3_selected = self.pick_fill_color(self.outline_color_milestone_3)
    def __milestone_4_outline_color(self):
        self.str_outline_color_milestone_4_selected = self.pick_fill_color(self.outline_color_milestone_4)
    def __milestone_5_outline_color(self):
        self.str_outline_color_milestone_5_selected = self.pick_fill_color(self.outline_color_milestone_5)
    
    
    # Function to Create Timeline Properties
    def create_timeline_properties(self):
        # Font Style/Properties
        self.row_1_font_button = ttk.Button(self.canvas, text="Row 1", command= lambda: self.row_1_font_chooser(), padding=-1)
        self.row_1_font_button.place(x=555, y=200, height=20, width=100)

        self.row_2_font_button = ttk.Button(self.canvas, text="Row 2", command= lambda: self.row_2_font_chooser(), padding=-1)
        self.row_2_font_button.place(x=665, y=200, height=20, width=100)

        # Font Color
        self.row_1_font_color_button = Button(self.canvas, text="Row 1", command= lambda: self.row_1_font_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6")
        self.row_1_font_color_button.place(x=555, y=230, height=20, width=100)

        self.row_2_font_color_button = Button(self.canvas, text="Row 2", command= lambda: self.row_2_font_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6")
        self.row_2_font_color_button.place(x=665, y=230, height=20, width=100)

        # Fill Color
        self.row_1_fill_color_button = Button(self.canvas, text="Row 1", command= lambda: self.row_1_fill_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.row_1_fill_color_button.place(x=555, y=260, height=20, width=100)

        self.row_2_fill_color_button = Button(self.canvas, text="Row 2", command= lambda: self.row_2_fill_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6")
        self.row_2_fill_color_button.place(x=665, y=260, height=20, width=100)

        # Text Align
        self.row_1_text_align_selection = StringVar()
        self.row_1_text_align = ttk.Combobox(self.canvas,textvariable=self.row_1_text_align_selection,values=self.text_align_list,width=13)
        self.row_1_text_align.current(2)
        self.row_1_text_align.state(['readonly'])
        self.canvas.create_window(555,288,anchor="nw",window=self.row_1_text_align)

        self.row_2_text_align_selection = StringVar()
        self.row_2_text_align = ttk.Combobox(self.canvas,textvariable=self.row_2_text_align_selection,values=self.text_align_list,width=13)
        self.row_2_text_align.current(2)
        self.row_2_text_align.state(['readonly'])
        self.canvas.create_window(665,288,anchor="nw",window=self.row_2_text_align)
    
    # Destroy timline properties
    def destroy_timeline_properties(self):
        self.row_1_font_button.destroy()
        self.row_2_font_button.destroy()
        self.row_1_font_color_button.destroy()
        self.row_2_font_color_button.destroy()
        self.row_1_fill_color_button.destroy()
        self.row_2_fill_color_button.destroy()
        self.row_1_text_align.destroy()
        self.row_2_text_align.destroy()

    # Depending on the timlinecheckbutton creates/destroys timeline properties
    def timeline_checkbutton_command(self):
        if self.create_timeline_checkbutton_selection.get():
            self.create_timeline_properties()
        else:
            self.destroy_timeline_properties()
    
    # Creates Milestone Properties
    def create_milestone_properties(self):
        # Font Style/Properties
        self.milestone_font_button = ttk.Button(self.canvas, text="Milestone", command= lambda: self.milestone_font_chooser(), padding=-1)
        self.milestone_font_button.place(x=775, y=200, height=20, width=100)
        
        # Font Color
        self.milestone_font_color_button = Button(self.canvas, text="Milestone", command= lambda: self.milestone_font_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6")
        self.milestone_font_color_button.place(x=775, y=230, height=20, width=100)

        # Fill Color
        self.milestone_fill_color_button = Button(self.canvas, text="Milestone", command= lambda: self.milestone_fill_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.milestone_fill_color_button.place(x=775, y=260, height=20, width=100)

        # Text Align
        self.milestone_text_align_selection = StringVar()
        self.milestone_text_align = ttk.Combobox(self.canvas,textvariable=self.milestone_text_align_selection,values=self.text_align_list,width=13)
        self.milestone_text_align.current(0)
        self.milestone_text_align.state(['readonly'])
        self.canvas.create_window(775,288,anchor="nw",window=self.milestone_text_align)

        # Lists
        self.secondary_shapes_list=[
            "Diamond",
            "Star",
            "Circle",
            "Triangle",
            "Square"
        ]

        self.shape_alignment_list=[
            "Left of Shape",
            "Right of Shape",
            "Bottom of Shape"
        ]

        # Milestone Labels/Checkbuttons
        self.milestone_1_selection = BooleanVar()
        self.milestone_1 = ttk.Checkbutton(self.canvas,command=self.__milestone_1_command,
        text="Milestone 1",variable=self.milestone_1_selection,width=16)
        self.canvas.create_window(200,375,anchor="nw",window=self.milestone_1)

        self.milestone_2_selection = BooleanVar()
        self.milestone_2 = ttk.Checkbutton(self.canvas,command=self.__milestone_2_command,
        text="Milestone 2",variable=self.milestone_2_selection,width=16)
        self.canvas.create_window(350,375,anchor="nw",window=self.milestone_2)

        self.milestone_3_selection = BooleanVar()
        self.milestone_3 = ttk.Checkbutton(self.canvas,command=self.__milestone_3_command,
        text="Milestone 3",variable=self.milestone_3_selection,width=16)
        self.canvas.create_window(500,375,anchor="nw",window=self.milestone_3)

        self.milestone_4_selection = BooleanVar()
        self.milestone_4 = ttk.Checkbutton(self.canvas,command=self.__milestone_4_command,
        text="Milestone 4",variable=self.milestone_4_selection,width=16)
        self.canvas.create_window(650,375,anchor="nw",window=self.milestone_4)

        self.milestone_5_selection = BooleanVar()
        self.milestone_5 = ttk.Checkbutton(self.canvas,command=self.__milestone_5_command,
        text="Milestone 5",variable=self.milestone_5_selection,width=16)
        self.canvas.create_window(800,375,anchor="nw",window=self.milestone_5)

        # Create Paramaters Labels
        self.canvas.create_text(68,400,anchor="nw",text="Shape Type\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,430,anchor="nw",text="Shape Color\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,460,anchor="nw",text="Shape Outline Color\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,490,anchor="nw",text="Textbox Alignment\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,520,anchor="nw",text="Milestone Name\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,550,anchor="nw",text="Milestone Month\t:",fill=self.fill_color,font=self.font_properties)
        self.canvas.create_text(68,580,anchor="nw",text="Milestone Year\t:",fill=self.fill_color,font=self.font_properties)
        

    def __create_milestone_1_properties(self):    
        
        # Milestone shape type selection
        self.shape_type_milestone_1_selection = StringVar()
        self.shape_type_milestone_1 = ttk.Combobox(self.canvas,textvariable= self.shape_type_milestone_1_selection, value=self.secondary_shapes_list, width= 16)
        self.shape_type_milestone_1.current(0)
        self.shape_type_milestone_1.state(['readonly'])
        self.canvas.create_window(200, 400,anchor="nw",window=self.shape_type_milestone_1)
        # Shape Color
        self.shape_color_milestone_1 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_1_shape_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.shape_color_milestone_1.place(x=200, y=430, height=20, width=120)
        # outline Color
        self.outline_color_milestone_1 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_1_outline_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.outline_color_milestone_1.place(x=200, y=460, height=20, width=120)
        # Milestone Alignment
        self.shape_align_milestone_1_selection = StringVar()
        self.shape_align_milestone_1 = ttk.Combobox(self.canvas,textvariable= self.shape_align_milestone_1_selection, value=self.shape_alignment_list, width= 16)
        self.shape_align_milestone_1.current(0)
        self.shape_align_milestone_1.state(['readonly'])
        self.canvas.create_window(200, 490,anchor="nw",window=self.shape_align_milestone_1)
        # Milestone Name
        self.milestone_name_milestone_1_selection = StringVar()
        self.milestone_name_milestone_1 = ttk.Entry(self.canvas,textvariable=self.milestone_name_milestone_1_selection, width=19)
        self.milestone_name_milestone_1.insert(0,"Milestone 1")
        self.canvas.create_window(200, 520, anchor="nw", window=self.milestone_name_milestone_1)
        # Milestone Month
        self.month_milestone_1_selection = StringVar()
        self.month_milestone_1 = ttk.Combobox(self.canvas,textvariable= self.month_milestone_1_selection, value=list(months_dict.keys()), width= 16)
        self.month_milestone_1.current(0)
        self.month_milestone_1.state(['readonly'])
        self.canvas.create_window(200, 550,anchor="nw",window=self.month_milestone_1)
        # Milestone Year
        self.year_milestone_1_selection = StringVar()
        self.year_milestone_1 = ttk.Entry(self.canvas,textvariable=self.year_milestone_1_selection, width=19)
        self.year_milestone_1.insert(0,"2022")
        self.canvas.create_window(200, 580, anchor="nw", window=self.year_milestone_1)

    
    def __destroy_milestone_1_properties(self):
        self.shape_type_milestone_1.destroy()
        self.shape_align_milestone_1.destroy()
        self.shape_color_milestone_1.destroy()
        self.outline_color_milestone_1.destroy()
        self.milestone_name_milestone_1.destroy()
        self.month_milestone_1.destroy()
        self.year_milestone_1.destroy()
    
    def __milestone_1_command(self):
        if self.milestone_1_selection.get():
            self.__create_milestone_1_properties()
        else:
            self.__destroy_milestone_1_properties()

    def __create_milestone_2_properties(self):    
        # Milestone shape type selection
        self.shape_type_milestone_2_selection = StringVar()
        self.shape_type_milestone_2 = ttk.Combobox(self.canvas,textvariable= self.shape_type_milestone_2_selection, value=self.secondary_shapes_list, width= 16)
        self.shape_type_milestone_2.current(0)
        self.shape_type_milestone_2.state(['readonly'])
        self.canvas.create_window(350, 400,anchor="nw",window=self.shape_type_milestone_2)
        # Shape Color
        self.shape_color_milestone_2 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_2_shape_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.shape_color_milestone_2.place(x=350, y=430, height=20, width=120)
        # outline Color
        self.outline_color_milestone_2 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_2_outline_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.outline_color_milestone_2.place(x=350, y=460, height=20, width=120)
        # Milestone Alignment
        self.shape_align_milestone_2_selection = StringVar()
        self.shape_align_milestone_2 = ttk.Combobox(self.canvas,textvariable= self.shape_align_milestone_2_selection, value=self.shape_alignment_list, width= 16)
        self.shape_align_milestone_2.current(0)
        self.shape_align_milestone_2.state(['readonly'])
        self.canvas.create_window(350, 490,anchor="nw",window=self.shape_align_milestone_2)
        # Milestone Name
        self.milestone_name_milestone_2_selection = StringVar()
        self.milestone_name_milestone_2 = ttk.Entry(self.canvas,textvariable=self.milestone_name_milestone_2_selection, width=19)
        self.milestone_name_milestone_2.insert(0,"Milestone 2")
        self.canvas.create_window(350, 520, anchor="nw", window=self.milestone_name_milestone_2)
        # Milestone Month
        self.month_milestone_2_selection = StringVar()
        self.month_milestone_2 = ttk.Combobox(self.canvas,textvariable= self.month_milestone_2_selection, value=list(months_dict.keys()), width= 16)
        self.month_milestone_2.current(0)
        self.month_milestone_2.state(['readonly'])
        self.canvas.create_window(350, 550,anchor="nw",window=self.month_milestone_2)
        # Milestone Year
        self.year_milestone_2_selection = StringVar()
        self.year_milestone_2 = ttk.Entry(self.canvas,textvariable=self.year_milestone_2_selection, width=19)
        self.year_milestone_2.insert(0,"2022")
        self.canvas.create_window(350, 580, anchor="nw", window=self.year_milestone_2)
    
    def __destroy_milestone_2_properties(self):
        self.shape_type_milestone_2.destroy()
        self.shape_align_milestone_2.destroy()
        self.shape_color_milestone_2.destroy()
        self.outline_color_milestone_2.destroy()
        self.milestone_name_milestone_2.destroy()
        self.month_milestone_2.destroy()
        self.year_milestone_2.destroy()
    
    def __milestone_2_command(self):
        if self.milestone_2_selection.get():
            self.__create_milestone_2_properties()
        else:
            self.__destroy_milestone_2_properties()

    def __create_milestone_3_properties(self):    
        # Milestone shape type selection
        self.shape_type_milestone_3_selection = StringVar()
        self.shape_type_milestone_3 = ttk.Combobox(self.canvas,textvariable= self.shape_type_milestone_3_selection, value=self.secondary_shapes_list, width= 16)
        self.shape_type_milestone_3.current(0)
        self.shape_type_milestone_3.state(['readonly'])
        self.canvas.create_window(500, 400,anchor="nw",window=self.shape_type_milestone_3)
        # Shape Color
        self.shape_color_milestone_3 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_3_shape_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.shape_color_milestone_3.place(x=500, y=430, height=20, width=120)
        # outline Color
        self.outline_color_milestone_3 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_3_outline_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.outline_color_milestone_3.place(x=500, y=460, height=20, width=120)
        # Milestone Alignment
        self.shape_align_milestone_3_selection = StringVar()
        self.shape_align_milestone_3 = ttk.Combobox(self.canvas,textvariable= self.shape_align_milestone_3_selection, value=self.shape_alignment_list, width= 16)
        self.shape_align_milestone_3.current(0)
        self.shape_align_milestone_3.state(['readonly'])
        self.canvas.create_window(500, 490,anchor="nw",window=self.shape_align_milestone_3)
        # Milestone Name
        self.milestone_name_milestone_3_selection = StringVar()
        self.milestone_name_milestone_3 = ttk.Entry(self.canvas,textvariable=self.milestone_name_milestone_3_selection, width=19)
        self.milestone_name_milestone_3.insert(0,"Milestone 3")
        self.canvas.create_window(500, 520, anchor="nw", window=self.milestone_name_milestone_3)
        # Milestone Month
        self.month_milestone_3_selection = StringVar()
        self.month_milestone_3 = ttk.Combobox(self.canvas,textvariable= self.month_milestone_3_selection, value=list(months_dict.keys()), width= 16)
        self.month_milestone_3.current(0)
        self.month_milestone_3.state(['readonly'])
        self.canvas.create_window(500, 550,anchor="nw",window=self.month_milestone_3)
        # Milestone Year
        self.year_milestone_3_selection = StringVar()
        self.year_milestone_3 = ttk.Entry(self.canvas,textvariable=self.year_milestone_3_selection, width=19)
        self.year_milestone_3.insert(0,"2022")
        self.canvas.create_window(500, 580, anchor="nw", window=self.year_milestone_3)
    
    def __destroy_milestone_3_properties(self):
        self.shape_type_milestone_3.destroy()
        self.shape_align_milestone_3.destroy()
        self.shape_color_milestone_3.destroy()
        self.outline_color_milestone_3.destroy()
        self.milestone_name_milestone_3.destroy()
        self.month_milestone_3.destroy()
        self.year_milestone_3.destroy()
    
    def __milestone_3_command(self):
        if self.milestone_3_selection.get():
            self.__create_milestone_3_properties()
        else:
            self.__destroy_milestone_3_properties()
            
    def __create_milestone_4_properties(self):    
        # Milestone shape type selection
        self.shape_type_milestone_4_selection = StringVar()
        self.shape_type_milestone_4 = ttk.Combobox(self.canvas,textvariable= self.shape_type_milestone_4_selection, value=self.secondary_shapes_list, width= 16)
        self.shape_type_milestone_4.current(0)
        self.shape_type_milestone_4.state(['readonly'])
        self.canvas.create_window(650, 400,anchor="nw",window=self.shape_type_milestone_4)
        # Shape Color
        self.shape_color_milestone_4 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_4_shape_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.shape_color_milestone_4.place(x=650, y=430, height=20, width=120)
         # outline Color
        self.outline_color_milestone_4 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_4_outline_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.outline_color_milestone_4.place(x=650, y=460, height=20, width=120)
        # Milestone Alignment
        self.shape_align_milestone_4_selection = StringVar()
        self.shape_align_milestone_4 = ttk.Combobox(self.canvas,textvariable= self.shape_align_milestone_4_selection, value=self.shape_alignment_list, width= 16)
        self.shape_align_milestone_4.current(0)
        self.shape_align_milestone_4.state(['readonly'])
        self.canvas.create_window(650, 490,anchor="nw",window=self.shape_align_milestone_4)
        # Milestone Name
        self.milestone_name_milestone_4_selection = StringVar()
        self.milestone_name_milestone_4 = ttk.Entry(self.canvas,textvariable=self.milestone_name_milestone_4_selection, width=19)
        self.milestone_name_milestone_4.insert(0,"Milestone 4")
        self.canvas.create_window(650, 520, anchor="nw", window=self.milestone_name_milestone_4)
        # Milestone Month
        self.month_milestone_4_selection = StringVar()
        self.month_milestone_4 = ttk.Combobox(self.canvas,textvariable= self.month_milestone_4_selection, value=list(months_dict.keys()), width= 16)
        self.month_milestone_4.current(0)
        self.month_milestone_4.state(['readonly'])
        self.canvas.create_window(650, 550,anchor="nw",window=self.month_milestone_4)
        # Milestone Year
        self.year_milestone_4_selection = StringVar()
        self.year_milestone_4 = ttk.Entry(self.canvas,textvariable=self.year_milestone_4_selection, width=19)
        self.year_milestone_4.insert(0,"2022")
        self.canvas.create_window(650, 580, anchor="nw", window=self.year_milestone_4)
        
    
    def __destroy_milestone_4_properties(self):
        self.shape_type_milestone_4.destroy()
        self.shape_align_milestone_4.destroy()
        self.shape_color_milestone_4.destroy()
        self.outline_color_milestone_4.destroy()
        self.milestone_name_milestone_4.destroy()
        self.month_milestone_4.destroy()
        self.year_milestone_4.destroy()
    
    def __milestone_4_command(self):
        if self.milestone_4_selection.get():
            self.__create_milestone_4_properties()
        else:
            self.__destroy_milestone_4_properties()

    def __create_milestone_5_properties(self):    
        # Milestone shape type selection
        self.shape_type_milestone_5_selection = StringVar()
        self.shape_type_milestone_5 = ttk.Combobox(self.canvas,textvariable= self.shape_type_milestone_5_selection, value=self.secondary_shapes_list, width= 16)
        self.shape_type_milestone_5.current(0)
        self.shape_type_milestone_5.state(['readonly'])
        self.canvas.create_window(800, 400,anchor="nw",window=self.shape_type_milestone_5)
        # Shape Color
        self.shape_color_milestone_5 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_5_shape_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.shape_color_milestone_5.place(x=800, y=430, height=20, width=120)
        # outline Color
        self.outline_color_milestone_5 = Button(self.canvas, text="Select Color", command= lambda: self.__milestone_5_outline_color(),
            relief="flat", borderwidth=2, bd=1, activebackground="#ADD8E6",)
        self.outline_color_milestone_5.place(x=800, y=460, height=20, width=120)
        # Milestone Alignment
        self.shape_align_milestone_5_selection = StringVar()
        self.shape_align_milestone_5 = ttk.Combobox(self.canvas,textvariable= self.shape_align_milestone_5_selection, value=self.shape_alignment_list, width= 16)
        self.shape_align_milestone_5.current(0)
        self.shape_align_milestone_5.state(['readonly'])
        self.canvas.create_window(800, 490,anchor="nw",window=self.shape_align_milestone_5)
        # Milestone Name
        self.milestone_name_milestone_5_selection = StringVar()
        self.milestone_name_milestone_5 = ttk.Entry(self.canvas,textvariable=self.milestone_name_milestone_5_selection, width=19)
        self.milestone_name_milestone_5.insert(0,"Milestone 5")
        self.canvas.create_window(800, 520, anchor="nw", window=self.milestone_name_milestone_5)
        # Milestone Month
        self.month_milestone_5_selection = StringVar()
        self.month_milestone_5 = ttk.Combobox(self.canvas,textvariable= self.month_milestone_5_selection, value=list(months_dict.keys()), width= 16)
        self.month_milestone_5.current(0)
        self.month_milestone_5.state(['readonly'])
        self.canvas.create_window(800, 550,anchor="nw",window=self.month_milestone_5)
        # Milestone Year
        self.year_milestone_5_selection = StringVar()
        self.year_milestone_5 = ttk.Entry(self.canvas,textvariable=self.year_milestone_5_selection, width=19)
        self.year_milestone_5.insert(0,"2022")
        self.canvas.create_window(800, 580, anchor="nw", window=self.year_milestone_5)
    
    def __destroy_milestone_5_properties(self):
        self.shape_type_milestone_5.destroy()
        self.shape_align_milestone_5.destroy()
        self.shape_color_milestone_5.destroy()
        self.outline_color_milestone_5.destroy()
        self.milestone_name_milestone_5.destroy()
        self.month_milestone_5.destroy()
        self.year_milestone_5.destroy()
    
    def __milestone_5_command(self):
        if self.milestone_5_selection.get():
            self.__create_milestone_5_properties()
        else:
            self.__destroy_milestone_5_properties()

    # Destroy Milestone Properties
    def destroy_milestone_properties(self):
        self.milestone_fill_color_button.destroy()
        self.milestone_font_color_button.destroy()
        self.milestone_text_align.destroy()
        self.milestone_font_button.destroy()
        self.milestone_1.destroy()
        self.milestone_2.destroy()
        self.milestone_3.destroy()
        self.milestone_4.destroy()
        self.milestone_5.destroy()

        try:
            self.__destroy_milestone_1_properties()
        except AttributeError:
            pass
        try:
            self.__destroy_milestone_2_properties()
        except AttributeError:
            pass    
        try:    
            self.__destroy_milestone_3_properties()
        except AttributeError:
            pass
        try:    
            self.__destroy_milestone_4_properties()
        except AttributeError:
            pass
        try:    
            self.__destroy_milestone_5_properties()
        except AttributeError:
            pass
        

    # Depending on the milestone checkbutton creates/destroys timeline properties
    def milestone_checkbutton_command(self):
        if self.create_milestone_checkbutton_selection.get():
            self.create_milestone_properties()
        else:
            self.destroy_milestone_properties()

    # Function to return relative path    
    def relative_to_assets(self,path: str) -> Path:
        return self.ASSETS_PATH / Path(path)

if __name__=="__main__":
    root = Tk()
    root.wm_title(str_project_title)
    root.geometry("984x700")
    root.rowconfigure(0,weight=1)
    root.columnconfigure(0,weight=1)
    frame = TimelinePropertiesFrame(root,None,slideLayoutList(Presentation()))
    showFrame(frame.frame)
    root.mainloop()
