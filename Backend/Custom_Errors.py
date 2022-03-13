# Creates custom error handling for Simple Gantt Chart Presentation creation loop

class ColorSelectionError(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self) -> str:
        return "Error: %s" %self.value
