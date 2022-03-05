### Introduction

SimpleGanttChart is a Python GUI tool built to automate creation of beautiful Gantt Chart Slides in PowerPoint using PowerPoint Autoshapes. 

### Using SimpleGanttChart

Gantt Charts are excellent for project management. A wide range of project management tools is available in the market, but these tools lack the visual appeal in business presentations. Create great  slides by following these three simple steps:

1. Download the data from these tools as an excel file. 
2. Prepare the data with simple and repeatable steps (even automate it using Power Query!)
3. Choose your shape and font properties.

And that's it!

### Requirements and Installations

SimpleGanttChart is built using Python, hence you need to install python to run the tool. Installing Python is a one-time process. Head over to [python website](https://www.python.org/downloads/) to download and install python. Make sure to check "Add to Path" when installing python.

SimpleGanttChart primarily uses [pandas](https://pandas.pydata.org/) to import and process excel data and automates Gantt Chart preparation in PowerPoint using [python-pptx](https://python-pptx.readthedocs.io/en/latest/index.html). In additon to pandas and python-pptx, SimpleGanttChart uses [openpyxl](https://openpyxl.readthedocs.io/en/stable/) to read Excel/csv data. Installing these dependencies is also open-time process

To install these dependencies. Open terminal and paste the following lines.

1. Install pandas
```markdown
    pip install pandas
```
2. Install python-pptx
```markdown
    pip install python-pptx
```
3. Install openpyxl
```markdown
    pip install openpyxl
```
### Get started with SimpleGanttChart

To run SimpleGanttChart tool first you need to download it from [SimpleGanttChart Repository](https://github.com/ndnag19/SimpleGanttChart). Follow these steps to download code as .zip file.
1. Visit [SimpleGanttChart Repository](https://github.com/ndnag19/SimpleGanttChart) and click on Code.

2. Download the code as .zip file

3. Extract the .zip file to desired location


Running SimpleGanttChart is as simple double clicking the **main.py** file in the SimpleGanttChart folder. 

Alternatively, you can follow these two steps to run the application
1. Open terminal in the SimpleGanttChart folder (make sure you can see **main.py** in this folder). On Windows, hold _SHIFT_ key and _RIGHT MOUSE CLICK_. A window will pop up, click on _Open PowerShell window here_
2. Run the command ``` python main.py```
