import sys

from cx_Freeze import Executable, setup

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("main.py", base=base)]

setup(
    name="Simple Gantt Chart",
    version="0.1",
    description="Creates gantt charts in powerpoint using excel/csv data",
    executables=executables,
)