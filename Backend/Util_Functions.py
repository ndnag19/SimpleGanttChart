#Library with all the small functions to be used

#imports for the application
import pandas as pd
from dateutil.relativedelta import relativedelta
import pptx
from pptx.util import Cm, Inches


#Import data from excel file
def importDataExcel(path: str,sheet: str):
    data = pd.read_excel(path,sheet)
    return data

#Import data from csv file
def importDataCsv(path: str):
    data = pd.read_csv(path)
    return data

#Add months to start date
def addMonths(start_date: pd._libs.tslibs.timestamps.Timestamp,delta_period: int):
    end_date = start_date + relativedelta(months=delta_period)
    return end_date

#Save Gantt Chart
def saveGantt(prs,path: str,name: str):
    prs.save(path+name+".pptx")
    return f"{name}.pptx is successfully saved at {path}"

#Convert Cms to Inches
def cmToInches(length: float):
    length = Cm(length)
    return length.inchespip

#Based on the input string converts cm to inches or returns inches
def getInches(length: float,metric: str):
    if metric.lower()=="cm":
        return length/2.54
    return length