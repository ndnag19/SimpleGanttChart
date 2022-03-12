from Frontend.window import MainWindow
from Frontend.commands import fontList
from collections import OrderedDict

master_fonts = list(OrderedDict.fromkeys(fontList()))
if __name__=="__main__":    
    mywindow = MainWindow(font_list=master_fonts)
    mywindow.start_window()