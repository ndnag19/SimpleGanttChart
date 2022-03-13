#imports for the functions defined here
import pandas as pd
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import *
from dateutil.relativedelta import relativedelta
import pptx
from tkinter.messagebox import showerror

#Create a Table
def createTable(rows: int,cols: int, left: float,top: float,width: float,height: float, slide):
    shape=slide.shapes.add_table(rows,cols,Inches(left), Inches(top), Inches(width), Inches(height))
    return shape.table


#Create a Monthly Table
def simpleMonthlyTable(table: pptx.table.Table,start_month: pd._libs.tslibs.timestamps.Timestamp, no_of_months: int,
font_size=12,font_color="#FFFFFF",font_name="Arial",fill_color_year=None,fill_color_month=None,text_align="Center",is_bold=False,is_italics=False):
    current_month = start_month
    current_year=current_month.year
    index1=0
    for i in range(no_of_months):
        new_current_year=current_month.year
        if i>0 and new_current_year>current_year and i<no_of_months-1:
            cell = table.cell(0,index1)
            index2=i
            other_cell = table.cell(0,index2-1)
            cell.merge(other_cell)
            index1=index2
            
            cell = table.cell(0,index1)

            if fill_color_year==None:
                cell.fill.solid()
                cell.fill.fore_color.theme_color=MSO_THEME_COLOR.ACCENT_1
            else:
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(fill_color_year.lstrip('#'))
                except:
                    showerror("Invalid Fill Color",
                    "Enter a valid hex color code. Ex. #FFFF00")
            
            cell.text_frame.clear()

            p=cell.text_frame.paragraphs[0]

            if text_align.lower()=="right":
                p.alignment = PP_ALIGN.RIGHT
            elif text_align.lower()=="center":
                p.alignment = PP_ALIGN.CENTER
            else:
                p.alignment = PP_ALIGN.LEFT
            
            run = p.add_run()
            run.text= str(current_month.year)
            run.font.bold = is_bold
            run.font.italic = is_italics
            run.font.name = font_name
            run.font.size = Pt(font_size+1)
            
            try: 
                run.font.color.rgb = RGBColor.from_string(font_color.lstrip('#'))
            except:
                showerror("Invalid Font Color",
                    "Enter a valid hex color code. Ex. #FFFF00")
                        
            current_year=new_current_year   
        if i==0:
            cell = table.cell(0,i)
            cell.text_frame.clear()

            if fill_color_year==None:
                cell.fill.solid()
                cell.fill.fore_color.theme_color=MSO_THEME_COLOR.ACCENT_1
            else:
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(fill_color_year.lstrip('#'))
                except:
                    showerror("Invalid Fill Color",
                    "Enter a valid hex color code. Ex. #FFFF00")

            p=cell.text_frame.paragraphs[0]

            if text_align.lower()=="right":
                p.alignment = PP_ALIGN.RIGHT
            elif text_align.lower()=="center":
                p.alignment = PP_ALIGN.CENTER
            else:
                p.alignment = PP_ALIGN.LEFT   

            run = p.add_run()
            run.text= str(current_month.year)
            run.font.bold = is_bold
            run.font.italic = is_italics
            run.font.name = font_name
            run.font.size = Pt(font_size+1)
            try:
                cell.fill.solid() 
                run.font.color.rgb = RGBColor.from_string(font_color.lstrip('#'))
            except:
                showerror("Invalid Font Color",
                    "Enter a valid hex color code. Ex. #FFFF00")
                       
        if i==no_of_months-1:
            cell = table.cell(0,index1)
            other_cell = table.cell(0,i)
            cell.merge(other_cell)

        cell = table.cell(1,i)
        if fill_color_month==None:
            cell.fill.solid()
            cell.fill.fore_color.theme_color=MSO_THEME_COLOR.ACCENT_1
            cell.fill.fore_color.brightness = -0.5
        else:
            try:
                cell.fill.fore_color.rgb = RGBColor.from_string(fill_color_month.lstrip('#'))
            except:
                showerror("Invalid Fill Color",
                    "Enter a valid hex color code. Ex. #FFFF00")
        
        cell.text_frame.clear()
        p=cell.text_frame.paragraphs[0]

        if text_align.lower()=="right":
            p.alignment = PP_ALIGN.RIGHT
        elif text_align.lower()=="center":
            p.alignment = PP_ALIGN.CENTER
        else:
            p.alignment = PP_ALIGN.LEFT

        run = p.add_run()
        run.font.bold = is_bold
        run.font.italic = is_italics
        run.font.name = font_name
        run.text = str(current_month.month)
        # run.text = str(current_month.month_name())[0:3]
        run.font.size = Pt(font_size)
        try: 
            run.font.color.rgb = RGBColor.from_string(font_color.lstrip('#'))
        except:
            showerror("Invalid Font Color",
                    "Enter a valid hex color code. Ex. #FFFF00")
        
        current_month=current_month+relativedelta(months=1)
