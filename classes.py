"""The code is divided into three main classes that are the base for the
implementation of the spreadsheets: Cell, Worksheet, Workbook"""

# Import packages
from typing import *
import re
import math
import statistics
import json


# Cell class
class Cell:
    """
    Cell is the basic unit of storage in a worksheet, each cell can hold
    individual data, text or formulas (to use formula - the user must write
    the equals sign "=" at the beginning of writing in the cell). cells are
    referenced by their column letter and row number (e.g. A1, B10).
    """
    def __init__(self, worksheet) -> None:
        self.value: Optional[float] = None
        self.txt: Optional[str] = None
        # Implementing the observer design pattern to manage dependencies
        # between cells. This allows a cell to notify other dependent cells
        # (subscribers) to update themselves when its value changes.
        self.subscribers: List['Cell'] = []
        self.formula: Optional[str] = None
        self.owner_worksheet = worksheet
    def subscribe(self, cell: 'Cell') -> None:
        """
        Adds a new cell to the list of subscribers that will be notified when this cell's value changes.
        """
        if cell not in self.subscribers:
            self.subscribers.append(cell)

    def set_value(self, value: Union[float, str]) -> None:
        """
        Sets the numerical value of the cell and notifies subscribers about the change.
        """
        if isinstance(value, str) and value.startswith('='):
            self.formula = value[1:]
            self.__calculate_expression()
            cell_names = re.findall(r'[A-Za-z]+\d+', self.formula) # fix it to work also for A10
            for cell_name in cell_names:
                cell = self.owner_worksheet.get_cell_by_reference(cell_name)
                cell.subscribe(self)

        elif self.value != value:
                self.value = value

        self.notify_subscribers()


    def get_display_value(self):
        return str(self.value) if self.value is not None else ""

    def get_value(self) -> float:
        """
        Returns the current numerical value of the cell.
        """
        return self.value

    def __calculate_expression(self) -> None:
        # Find all cell names in the expression
        if self.formula is None:
            return
        expression = self.formula
        cell_names = re.findall(r'[A-Za-z]+\d+', expression) # fix it to work also for A10
        modified_expression = expression
        
        # Replace cell names with their current values from the worksheet
        for cell_name in cell_names:
            if self.owner_worksheet.cell_exists(cell_name):
                cell = self.owner_worksheet.get_cell_by_reference(cell_name)
                cell.subscribe(self)
                cell_value = cell.get_value()
                # Use 0 if the cell is empty
                if cell_value is None:
                    cell_value = 0
                modified_expression = modified_expression.replace(cell_name, str(cell_value))
            else:
                raise ValueError(f"Cell {cell_name} does not exist in the worksheet.")

        # Prepare the safe environment with necessary math functions
        custom_functions = {"sqrt": math.sqrt, "pow": math.pow}  # Add more functions as needed

        # Evaluate the expression
        try:
            result = eval(modified_expression, custom_functions)
        except Exception as e:
            raise ValueError(f"Failed to evaluate expression '{modified_expression}'. Error: {str(e)}")

        # Set the result in the target cell and trigger updates to subscribers
        self.set_value(result)

    def notify_subscribers(self) -> None:
        """
        Notifies all subscribed cells to update based on the new value of this cell.
        """
        for subscriber in self.subscribers:
            subscriber.update()

    def update(self):
        self.__calculate_expression()



class Worksheet:
    """
    Sheet is a single page or tab within a workbook.
    it consists of a grid of cells organized in rows and columns where users
    can enter, calculate, manipulate, and analyze data.
    """
    def __init__(self) -> None:
        # Setting default dimensions
        self.num_rows = 10
        self.num_columns = 10

        # Creating a sheet as a 2D list consisting of Cell objects
        self.table = []
        for row in range(self.num_rows):
            new_row = []
            for column in range(self.num_columns):
                # Adding a new Cell object to each column in the row
                new_row.append(Cell(self))
            self.table.append(new_row)

    def expand_rows(self) -> None:
        # Creating a new row with a new Cell in each column
        new_row = [Cell(self) for _ in range(self.num_columns)]
        # Adding the new row to the table
        self.table.append(new_row)
        # Updating the row count
        self.num_rows += 1

    def expand_columns(self) -> None:
        # Adding a new Cell to each existing row
        for row in self.table:
            row.append(Cell(self))
        # Updating the column count
        self.num_columns += 1

    def set_cell_value(self, row: int, column: int, value) -> None: # TODO: union type str or float
        if 0 <= row < self.num_rows and 0 <= column < self.num_columns:
            self.table[row][column].set_value(value)
        else:
            print(f"Attempted to set value in cell at ({row}, {column}) which is not in the worksheet.")


    def get_cell(self, row: int, column: int) -> Cell:
        if 0 <= row < self.num_rows and 0 <= column < self.num_columns:
            return self.table[row][column]
        else:
            raise ValueError(f"Attempted to access cell at ({row}, {column}) which is not in the worksheet.")

    def get_cell_value(self, row: int, column: int) -> float:
        cell = self.get_cell(row, column)
        if isinstance(cell, Cell):
            return cell.get_value()
        else:
            raise ValueError("The cell at the specified location is not a valid Cell object")

    def get_cell_indices(self, cell_name: str) -> tuple[int, int]:
        """
        Convert a cell name (e.g., 'A1', 'B3') to row and column indices.
        """
        match = re.match(r"([A-Za-z]+)(\d+)", cell_name)
        if not match:
            raise ValueError("Invalid cell name format")

        column_letters, row_number = match.groups()
        column_index = column_letter_to_index(column_letters)  # Assuming this method is defined correctly
        row_index = int(row_number) - 1
        return row_index, column_index

    def get_cell_by_reference(self, cell_name: str) -> Cell:
        if self.cell_exists(cell_name):
            row_index, column_index = self.get_cell_indices(cell_name)
            return self.get_cell(row_index, column_index)
        else:
            raise ValueError("Cell does not exist")

    def cell_exists(self, cell_name: str) -> bool:
        """
        Checks if a cell exists in the worksheet by its name.
        """
        try:
            row_index, column_index = self.get_cell_indices(cell_name)
            return 0 <= row_index < self.num_rows and 0 <= column_index < self.num_columns
        except ValueError:
            return False



# General functions that the worksheet class uses


def column_letter_to_index(letter: str) -> int:
    """Convert column letter to index (e.g., 'A' -> 0, 'B' -> 1, ...)"""
    return ord(letter.upper()) - ord('A')


"""
def compute_if(condition: bool, true_val: Any, false_val: Any) -> Any:
    #Evaluates a condition and returns the corresponding value.
    if condition:
        return true_val
    else:
        return false_val


def count_if(self, start_row: int, end_row: int, start_col: int, end_col: int, criterion: Callable[[Any], bool]) -> int:
    #Counts cells in a specified range that meet the given criterion.
    count = 0
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = self.get_cell(row, col)
            if criterion(cell.get_value()):
                count += 1
    return count
"""


class Workbook:
    """
    Workbook is a file that contains one or more sheets.
    """

    def __init__(self) -> None:
        self.sheets: Dict[str, Worksheet] = {}

    def add_sheet(self, sheet_name: str) -> None:
        """Add a new sheet with a given name if it doesn't already exist."""
        if sheet_name not in self.sheets:
            self.sheets[sheet_name] = Worksheet()
        else:
            print(f"Sheet '{sheet_name}' already exists.")

    def get_sheet(self, sheet_name: str) -> Optional[Worksheet]:
        """Retrieve a sheet by name."""
        return self.sheets.get(sheet_name, None)

    def remove_sheet(self, sheet_name: str) -> None:
        """Remove a sheet by name, if it exists."""
        if sheet_name in self.sheets:
            del self.sheets[sheet_name]
        else:
            print(f"Sheet '{sheet_name}' does not exist.")

    def list_sheets(self) -> List[str]:
        """List all sheet names in the workbook."""
        return list(self.sheets.keys())

    def expand_sheet(self, sheet_name: str, rows: bool = False, columns: bool = False) -> None:
        sheet = self.get_sheet(sheet_name)
        if sheet:
            if rows:
                sheet.expand_rows()
            if columns:
                sheet.expand_columns()
        else:
            print(f"Sheet '{sheet_name}' does not exist.")

    @staticmethod
    def load_from_json(data):
        workbook = Workbook()
        for sheet_name, sheet_data in data.items():
            worksheet = Worksheet()
            for row_data in sheet_data:
                row = []
                for cell_data in row_data:
                    cell = Cell()
                    cell.set_value(cell_data.get('value', None))
                    row.append(cell)
                worksheet.table.append(row)
            workbook.sheets[sheet_name] = worksheet
        return workbook

    def to_json(self):
        # Convert the entire workbook to a JSON-serializable dictionary
        workbook_data = {}
        for sheet_name, worksheet in self.sheets.items():
            sheet_data = []
            for row in worksheet.table:
                row_data = []
                for cell in row:
                    cell_data = {
                        'value': cell.value,
                    }
                    row_data.append(cell_data)
                sheet_data.append(row_data)
            workbook_data[sheet_name] = sheet_data
        return workbook_data

# Functions that operate on a range of cells


def calculate_on_range(worksheet, cell_range: List[str], function: str) -> Union[str, tuple[None, str]]:
    """
    Calculates specified statistics (MAX, MIN, SUM, AVERAGE) of values in the specified cell range within the worksheet.
    Returns None if all cells are empty or the values are None.
    """
    functions = {'max': max, 'min': min, 'sum': sum, 'average': statistics.mean}
    values = []
    for cell_name in cell_range:
        row_index, column_index = worksheet.get_cell_indices(cell_name)
        cell = worksheet.get_cell(row_index, column_index)
        value = cell.get_value()
        if value is not None:
            values.append(value)
    if values:
        result = functions.get(function, lambda x: None)(values)
        if result is not None:
            message = f"Computed {function} for the cell range {cell_range}: {result}"
            return message
        else:
            return None, f"Function {function} is not supported."
    else:
        return None, f"No valid values found in the cell range {cell_range}. Cannot compute {function}."


def save_workbook_as(workbook, filename):
    if filename:
        with open(filename, 'w') as file:
            json_data = workbook.to_json()
            json.dump(json_data, file, indent=4)
            print("Workbook saved to", filename)
