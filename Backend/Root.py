#imports for the application file
from Backend.Gantt_Shapes import *
from Backend.Gantt_Tables import *
from Backend.Util_Functions import *
from tkinter.messagebox import showerror
import pandas as pd
import numpy as np
import pptx

# Creates a dataframe with required columns for creating gantt charts
def data_setup_for_gantt(data: pd.DataFrame,task_name: str,task_duration: str,
task_level: str,gantt_start_date: str,task_start_date: str,gantt_duration: str):
    # Creates a new data frame
    gantt_data = data

    # Renames the column to prepare further operations
    gantt_data.rename(columns={
        task_name:"task_name",
        task_duration:"task_duration",
        task_level:"task_level",
        gantt_start_date:"gantt_start_date",
        task_start_date:"task_start_date",
        gantt_duration:"gantt_duration"
    },inplace=True)

    # Changes the type of the data incase the type is not good
    try:
        gantt_data['task_duration']=gantt_data['task_duration'].astype(float)
        gantt_data['gantt_duration']=gantt_data['gantt_duration'].astype(float)
        gantt_data['task_level']=gantt_data['task_level'].astype(int)
        gantt_data['task_name']=gantt_data['task_name'].astype("string")
        gantt_data['gantt_start_date']=pd.to_datetime(gantt_data['gantt_start_date'])
        gantt_data['task_start_date']=pd.to_datetime(gantt_data['task_start_date'])
    except:
        showerror("Incorrect Column Selection",
        "Incorrect column selection or data type. Please check your column selection and data types of columns used in the original data")
    # Create helper columns and other mathematical calculations
    gantt_data['Start Calibration']=gantt_data['task_start_date']-gantt_data['gantt_start_date']
    gantt_data['Start Calibration']=gantt_data['Start Calibration']/np.timedelta64(1,"M")
    gantt_data['Start Calibration']=gantt_data['Start Calibration'].astype(int)
    return gantt_data

# Creates the gantt slides and presentation
def create_gantt_slides(prs: pptx.Presentation,gantt_data: pd.DataFrame,table_left: str,table_left_metric: str, table_top: str,table_top_metric: str, 
table_width: str, table_width_metric: str,align_tl: list,shape_type_tl: list,shape_color_tl:list,font_prop_tl: list,font_style_tl:list,
font_color_tl: list,font_size_tl: list):
    gantt_duration = int(gantt_data['gantt_duration'][0])
    gantt_start_date = gantt_data['gantt_start_date'][0]
    # Changing the data type for table properties
    try:
        table_left = getInches(float(table_left),table_left_metric)
        table_top = getInches(float(table_top),table_top_metric)
        table_width = getInches(float(table_width),table_width_metric)
    except:
        showerror("Enter a Valid Value",f"Timeline properties must be numeric. Cross check these inputs \n{table_left}\n{table_top}\n{table_width}")
    
    # Defines the slide layout
    slide_layout = prs.slide_layouts[6]
    # All the variables required for Gantt Automation Operation
    width_per_col = table_width/gantt_duration
    max_slide_height = 7.5
    table_height = 0.5
    table_columns = gantt_duration
    table_rows = 2
    height_measure = table_top
    slide_counter = 0
    task_top = table_top+table_height*table_rows-0.2
    task_height = 0.2
    # Creates the first slide with table
    slide = [prs.slides.add_slide(slide_layout)]
    table = [createTable(table_rows,int(table_columns),table_left,table_top,table_width,table_height,slide[0])]
    simpleMonthlyTable(table[slide_counter],gantt_start_date,int(gantt_duration))
    # Create the main loop to do all the operation
    for i in range(len(gantt_data)):
        height_measure=task_top+task_height
        if height_measure>=(max_slide_height-0.5):
            slide_counter=slide_counter+1
            new_slide = prs.slides.add_slide(slide_layout)
            slide.append(new_slide)
            height_measure=table_top
            new_table = createTable(table_rows,int(table_columns),table_left,table_top,table_width,table_height,slide[slide_counter])
            table.append(new_table)
            simpleMonthlyTable(table[slide_counter],gantt_start_date,int(gantt_duration))
            task_top = table_top+table_height*table_rows-0.2
        if gantt_data['task_level'][i]==1:
            shape_left=table_left+gantt_data['Start Calibration'][i]*width_per_col
            shape_width = gantt_data['task_duration'][i]*width_per_col
            ganttShape(
                top=task_top,
                left=shape_left,
                width=shape_width,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                shape_name=shape_type_tl[0],
                fill_color=shape_color_tl[0]
            )
            text_left = set_textbox_shape_alignment(align_tl[0],shape_left=shape_left,shape_width=shape_width)
            prop=set_font_properties(font_prop_tl[0])

            if align_tl[0].lower()=="bottom of shape":
                task_top=task_top+task_height+0.05
            
            textBox(
                top=task_top,
                left=text_left,
                width=table_left+1,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                is_bold=prop[0],
                is_italics=prop[1],
                is_underline=prop[2],
                font_name=font_style_tl[0],
                font_color=font_color_tl[0],
                font_size=int(font_size_tl[0]),
                text=str(gantt_data['task_name'][i]),
                text_align="Left"
            )
            task_top=task_top+task_height+0.1
        elif gantt_data['task_level'][i]==2:
            shape_left=table_left+gantt_data['Start Calibration'][i]*width_per_col
            shape_width = gantt_data['task_duration'][i]*width_per_col
            ganttShape(
                top=task_top,
                left=shape_left,
                width=shape_width,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                shape_name=shape_type_tl[1],
                fill_color=shape_color_tl[1]
            )
            text_left = set_textbox_shape_alignment(align_tl[1],shape_left=shape_left,shape_width=shape_width)
            prop=set_font_properties(font_prop_tl[1])

            if align_tl[1].lower()=="bottom of shape":
                task_top=task_top+task_height+0.05
            
            textBox(
                top=task_top,
                left=text_left,
                width=table_left+1,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                is_bold=prop[0],
                is_italics=prop[1],
                is_underline=prop[2],
                font_name=font_style_tl[1],
                font_color=font_color_tl[1],
                font_size=int(font_size_tl[1]),
                text=str(gantt_data['task_name'][i]),
                text_align="Left"
            )
            task_top=task_top+task_height+0.1
        elif gantt_data['task_level'][i]==3:
            shape_left=table_left+gantt_data['Start Calibration'][i]*width_per_col
            shape_width = gantt_data['task_duration'][i]*width_per_col
            ganttShape(
                top=task_top,
                left=shape_left,
                width=shape_width,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                shape_name=shape_type_tl[2],
                fill_color=shape_color_tl[2]
            )
            text_left = set_textbox_shape_alignment(align_tl[2],shape_left=shape_left,shape_width=shape_width)
            prop=set_font_properties(font_prop_tl[2])

            if align_tl[2].lower()=="bottom of shape":
                task_top=task_top+task_height+0.05
            
            textBox(
                top=task_top,
                left=text_left,
                width=table_left+1,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                is_bold=prop[0],
                is_italics=prop[1],
                is_underline=prop[2],
                font_name=font_style_tl[2],
                font_color=font_color_tl[2],
                font_size=int(font_size_tl[2]),
                text=str(gantt_data['task_name'][i]),
                text_align="Left"
            )
            task_top=task_top+task_height+0.1
        elif gantt_data['task_level'][i]==4:
            shape_left=table_left+gantt_data['Start Calibration'][i]*width_per_col
            shape_width = gantt_data['task_duration'][i]*width_per_col
            milestoneShape(
                top=task_top,
                left=shape_left,
                width=task_height,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                shape_name=shape_type_tl[3],
                fill_color=shape_color_tl[3]
            )
            text_left = set_textbox_shape_alignment(align_tl[3],shape_left=shape_left,shape_width=shape_width)
            prop=set_font_properties(font_prop_tl[3])

            if align_tl[3].lower()=="bottom of shape":
                task_top=task_top+task_height+0.05
            
            textBox(
                top=task_top,
                left=text_left,
                width=table_left+1,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                is_bold=prop[0],
                is_italics=prop[1],
                is_underline=prop[2],
                font_name=font_style_tl[3],
                font_color=font_color_tl[3],
                font_size=int(font_size_tl[3]),
                text=str(gantt_data['task_name'][i]),
                text_align="Left"
            )
            task_top=task_top+task_height+0.1
        else:
            shape_left=table_left+gantt_data['Start Calibration'][i]*width_per_col
            shape_width = gantt_data['task_duration'][i]*width_per_col
            milestoneShape(
                top=task_top,
                left=shape_left,
                width=task_height,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                shape_name=shape_type_tl[4],
                fill_color=shape_color_tl[4]
            )
            text_left = set_textbox_shape_alignment(align_tl[4],shape_left=shape_left,shape_width=shape_width)
            prop=set_font_properties(font_prop_tl[4])

            if align_tl[4].lower()=="bottom of shape":
                task_top=task_top+task_height+0.05
            
            textBox(
                top=task_top,
                left=text_left,
                width=table_left+1,
                height=task_height,
                shapes=slide[slide_counter].shapes,
                is_bold=prop[0],
                is_italics=prop[1],
                is_underline=prop[2],
                font_name=font_style_tl[4],
                font_color=font_color_tl[4],
                font_size=int(font_size_tl[4]),
                text=str(gantt_data['task_name'][i]),
                text_align="Left"
            )
            task_top=task_top+task_height+0.1    
    return prs

# Function to align textboxes based on the selection
def set_textbox_shape_alignment(align: str, shape_left: float, shape_width: float):
    max_width = 13.3
    if align.lower()=="left of slide":
        return 0.3
    elif align.lower()=="right of slide":
        return 12.0
    elif align.lower()=="left of shape":
        return shape_left-2.0
    elif align.lower()=="right of shape":
        if max_width - shape_left - shape_width>2.0:
            return shape_width+shape_left+0.2
        return shape_width+shape_left-2.0
    else:
        return shape_left

# Function to set font properties in text box based on the selection
def set_font_properties(prop: str)->list():
    if prop.lower()=="regular":
        return [False,False,False]
    elif prop.lower()=="bold":
        return [True,False,False]
    elif prop.lower()=="italic":
        return [False,True,False]
    elif prop.lower()=="underline":
        return [False,False,True]
    elif prop.lower()=="bold italic":
        return [True,True,False]
    elif prop.lower()=="bold underline":
        return [True,False,True]
    elif prop.lower()=="underline italic":
        return [False,True,True]
    else:
        return [False,False,False]
    