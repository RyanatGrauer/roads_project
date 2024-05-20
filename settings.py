"""This file contains the settings and classes for the map and the game."""
import math
import openpyxl


class Cell:
    """This class represents a cell in the map."""

    def __init__(self, x, y, color: str):
        """Initialize the cell."""
        self.x = x
        self.y = y
        self.multiplier = 1
        self.color = color
        self.cost = 0
        self.speed = 0
        self.neighbors = []
        self.color_key = {
            "FFD9EAD3": "grass",
            "FF4A86E8": "water",
            "FFFFF2CC": "desert",
            "FF93C47D": "forest",
            "FFF9CB9C": "hills",
            "FF000000": "city",
            "FF7F6000": "mountain",
        }
        self._type = self.color_key.get(color)

        if self.type is None:
            raise ValueError(f"Color {color} not recognized")
        # assign multiplier based on cell type
        match self.type:
            case "grass":
                self.multiplier = 1
            case "water":
                self.multiplier = math.inf
            case "desert":
                self.multiplier = 1.5
            case "forest":
                self.multiplier = 2
            case "hills":
                self.multiplier = 2
            case "city":
                self.multiplier = 1
            case "mountain":
                self.multiplier = math.inf
            case _:
                self.multiplier = 1

    @property
    def type(self):
        """Return the type of the cell."""
        if self._type is None:
            raise ValueError(f"Color {self.color} not recognized")
        if self._type == "road":
            self.cost = 20
            self.speed = 50
        elif self._type == "highway":
            self.cost = 100
            self.speed = 70

        return self._type

    @type.setter
    def type(self, value):
        """Set the type of the cell."""
        self._type = value

    def __str__(self):
        """Return the cell as a string."""
        return f"({self.x}, {self.y}): {self.type}"


class Map:
    """This class contains the map of the game."""

    def __init__(self, size):
        """Initialize the map."""
        self.size = size
        self.grid = [[None for _ in range(size)] for _ in range(size)]

    def __str__(self):
        """Return the map as a string."""
        return "\n".join([" ".join([str(cell) for cell in row]) for row in self.grid])

    def read_data(self, filename: str, sheet: int = 2):
        """Read the data from the Excel file, only on cells after row 5 that are not colored white, based on
        the size of the map"""
        wb = openpyxl.load_workbook(filename)
        active_map = wb.worksheets[sheet]
        # loop through each cell starting from row 5 and ending at the size of the map
        for row in active_map.iter_rows(min_row=6, max_row=self.size + 5, min_col=2, max_col=self.size + 1):
            for cell in row:
                # if the cell is not colored white, add it to the map
                color = cell.fill.bgColor.index
                if color == 1:
                    color = "FF000000"
                if color != "FFFFFFFF":
                    # x is row, y is column
                    x = cell.row - 6
                    y = cell.column - 2
                    # Pick cell type based on color
                    self.grid[x][y] = Cell(x, y, color)
                    # print(cell.fill.start_color.index)


if __name__ == "__main__":
    new_map = Map(20)
    new_map.read_data("test_maps.xlsx", 0)
    print(new_map)
