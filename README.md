### Introduction

SimpleGanttChart is a Python GUI tool built to automate creation of beautiful Gantt Chart Slides in PowerPoint using PowerPoint Autoshapes. SimpleGanttChart was created to automate a boring task I used to do in my first job. I hope you can automate this boring task by using this simple tool.

### Using SimpleGanttChart

Gantt Charts are excellent for project management. A wide range of project management tools is available in the market, but these tools lack the visual appeal in business presentations. Create great  slides by following these three simple steps:

1. Download the data from these tools as an excel file. 
2. Prepare the data with simple and repeatable steps (even automate it using Power Query!)
3. Choose your shape and font properties.

And that's it!

### How to install the application?

Windows users can directly download the latest .exe file by [using this link](https://github.com/ndnag19/SimpleGanttChart/releases/download/v2.0/SimpleGanttChart.-.setup.exe).
SimpleGanttChart installs like any other windows application. Follow the instructs on screen to install and run the software.

For users using other operating systems. The setup file is on the way. In the meantime you can follow next few sections instructions to get SimpleGanttChart running.

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
![code_github](https://user-images.githubusercontent.com/39146876/156887577-1460a2b6-7d17-4dc0-ae31-7cf29c01da39.png)
2. Download the code as .zip file
![download_github](https://user-images.githubusercontent.com/39146876/156887589-4d3bfdf4-c9b3-498e-b8d2-66838f65af54.png)
3. Extract the .zip file to desired location

Running SimpleGanttChart is as simple double clicking the **main.py** file in the SimpleGanttChart folder. 

Alternatively, you can follow these two steps to run the application
- Open terminal in the SimpleGanttChart folder (make sure you can see **main.py** in this folder). On Windows, hold _SHIFT_ key and _RIGHT MOUSE CLICK_. A window will pop up, click on _Open PowerShell window here_
![open_powershell](https://user-images.githubusercontent.com/39146876/156887568-f964d2fe-e0ef-4bce-9dea-1d82aef1c1e9.png)
- Run the following command on terminal ``` python main.py```
