import sys
import json
import win32com.client
import win32gui
import os
import clipboard
from datetime import datetime, timedelta, date
import pandas as pd
import numpy as np
import time
import io
import base64
from PIL import ImageGrab
import PIL.Image
import sys
import json
import subprocess
import pyautogui
import re
from PIL.ImageQt import ImageQt
from pymsgbox import *
import pygetwindow as gw
import keyboard
import openpyxl
import datetime
from datetime import datetime
from my_lib.shared_context import ExecutionContext as Context
from typing import List, Dict, Any, Optional, Union
class handle_excel():
    """A class for handling Excel file operations using the openpyxl library."""

    def __init__(self,context: Context):
        """Initializes the File_handle class."""
        self.context = context
        pass
    def read_excel(self,file_link,sheet_name=None):
        """Opens an Excel file and returns the workbook and a specific sheet object.

        Args:
            file_link (str): The full path to the Excel file.
            sheet_name (str, optional): The name of the sheet to access. If None or not
                                        found, the active sheet is returned. Defaults to None.

        Returns:
            list or None: A list containing [workbook, sheet] objects on success,
                          or None if an error occurs (e.g., file not found).
        """
        print (file_link)
        try:
            workbook = openpyxl.load_workbook(file_link)
            sheet = workbook.active
            if sheet_name is not None and sheet_name !='None'  and sheet_name !='':
                if sheet_name in workbook.sheetnames:
                    sheet = workbook[sheet_name]
                    return [workbook, sheet]
                else:
                    print(f"Error: Sheet '{sheet_name}' not found in the workbook.")
                    sheet = workbook.active
                    return [workbook, sheet]
            else:
                sheet = workbook.active  # Get the active (first) sheet
                return [workbook, sheet]
    
    
        except FileNotFoundError:
            print(f"Error: File not found at '{file_link}'")
            return None
        except Exception as e:
            print(f"An error occurred while reading the Excel file: {e}")
            return None
            
        return None
    def check_Length_String (self,String):
        """ Check lenght of a String """
        try:
            str_len = len(str(String))
        except:
            str_len = 0
        return str_len

    def read_excel_cells(self,workbook,col,row : int):
        """Reads the value from a single cell in an Excel worksheet.

        Args:
            workbook (list): The [workbook, sheet] list object returned by `read_excel`.
            col (str): The column letter of the cell (e.g., 'A', 'B').
            row (int): The 0-indexed row number. The method adds 1 to match Excel's
                       1-based row indexing.

        Returns:
            any: The value contained in the specified cell.
        """
        
        col=col.replace("'","")
        print (f'"{col}{int(row)}"')
        sheet = workbook[1]
        
        return sheet[f"{col}{int(row)}"].value
        
    def write_excel_cells(self,workbook,col,row,value):
        """Writes a value to a single cell in an Excel worksheet.

        Args:
            workbook (list): The [workbook, sheet] list object returned by `read_excel`.
            col (str): The column letter of the cell.
            row (int): The 0-indexed row number. The method adds 1 to match Excel's
                       1-based row indexing.
            value (any): The value to write into the cell.

        Returns:
            str: A confirmation message: "assigned Value".
        """
        workbook[1][f"{col}{int(row+1)}"] = value
        return "assigned Value"
        
    def save_excel(self,workbook,file_name):
        """Saves and closes the Excel workbook to a specified file.

        Args:
            workbook (list): The [workbook, sheet] list object to be saved.
            file_name (str): The file path (including name) to save the workbook to.

        Returns:
            str: A confirmation message: "saved excel".
        """
        workbook[0].save(file_name)
        workbook[0].close()
        return f"saved excel"  
        
    def close_excel(self,workbook):
        """Closes the Excel workbook without saving changes.

        Args:
            workbook (openpyxl.workbook.workbook.Workbook): The workbook object to close.

        Returns:
            str: A confirmation message: "closed excel".
        """
        workbook[0].close()
        return f"closed excel" 
    
    def count_non_empty_rows_in_column(self,workbook, column_identifier):
        """Counts the number of non-empty cells in a specific column of a worksheet.

        Args:
            workbook (list): The [workbook, sheet] list object.
            column_identifier (str or int): The column to check. Can be a column letter
                                            (e.g., 'A') or a 1-based integer index (e.g., 1).

        Returns:
            int: The number of non-empty cells in the specified column, or -1 if an error occurs.
        """
        worksheet_object =workbook[1]
        if not isinstance(worksheet_object, openpyxl.worksheet.worksheet.Worksheet):
            print("Error: The provided object is not an openpyxl Worksheet.")
            return -1
    
        count = 0
        try:
            # Get the max row of the entire sheet to avoid iterating endlessly
            max_sheet_row = worksheet_object.max_row
    
            # Convert column identifier to column letter if it's an integer
            if isinstance(column_identifier, int):
                if column_identifier <= 0:
                    print("Error: Column index must be 1-based.")
                    return -1
                column_letter = openpyxl.utils.get_column_letter(column_identifier)
            elif isinstance(column_identifier, str):
                column_letter = column_identifier.upper() # Ensure uppercase
                # Optional: Validate column letter
                if not column_letter.isalpha():
                    print(f"Error: Invalid column letter '{column_identifier}'.")
                    return -1
            else:
                print("Error: Column identifier must be a string (letter) or an integer (1-based index).")
                return -1
    
            # Iterate through cells in the specified column up to the max row of the sheet
            for row_num in range(1, max_sheet_row + 1):
                cell = worksheet_object[f"{column_letter}{row_num}"]
                if cell.value is not None:
                    count += 1
    
            print(f"Sheet '{worksheet_object.title}', Column '{column_letter}': Non-empty cells = {count}")
            return count
    
        except Exception as e:
            print(f"An error occurred while counting column '{column_identifier}': {e}")
            return -1

class CSVHandler:
    def __init__(self):
        self.df = None
        self.file_link = None
    
    def read_csv(self, file_link: str, **kwargs) -> pd.DataFrame:
        """
        Read CSV file into a pandas DataFrame
        
        Args:
            file_link (str): Path to the CSV file
            **kwargs: Additional arguments for pd.read_csv()
        
        Returns:
            pd.DataFrame: Loaded DataFrame
        """
        try:
            self.file_link = file_link
            self.df = pd.read_csv(file_link, **kwargs)
            print(f"Successfully loaded CSV: {file_link}")
            print(f"Shape: {self.df.shape}")
            return self.df
        except FileNotFoundError:
            print(f"Error: File '{file_link}' not found.")
            return None
        except Exception as e:
            print(f"Error reading CSV: {str(e)}")
            return None
    
    def write_csv(self, df: pd.DataFrame, file_link: str, **kwargs) -> bool:
        """
        Write DataFrame to CSV file
        
        Args:
            df (pd.DataFrame): DataFrame to save
            file_link (str): Output file path
            **kwargs: Additional arguments for df.to_csv()
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Set default parameters
            default_kwargs = {'index': False}
            default_kwargs.update(kwargs)
            
            df.to_csv(file_link, **default_kwargs)
            print(f"Successfully saved CSV: {file_link}")
            return True
        except Exception as e:
            print(f"Error writing CSV: {str(e)}")
            return False
    
    def get_info(self) -> None:
        """Display basic information about the DataFrame"""
        if self.df is not None:
            print("=== DataFrame Info ===")
            print(f"Shape: {self.df.shape}")
            print(f"Columns: {list(self.df.columns)}")
            print(f"Data types:\n{self.df.dtypes}")
            print(f"Memory usage: {self.df.memory_usage(deep=True).sum()} bytes")
            print(f"Missing values:\n{self.df.isnull().sum()}")
        else:
            print("No DataFrame loaded.")
    
    def preview_data(self, n: int = 5) -> None:
        """
        Preview first and last n rows of the DataFrame
        
        Args:
            n (int): Number of rows to display
        """
        if self.df is not None:
            print(f"=== First {n} rows ===")
            print(self.df.head(n))
            print(f"\n=== Last {n} rows ===")
            print(self.df.tail(n))
        else:
            print("No DataFrame loaded.")
    
    def filter_data(self, conditions: Dict[str, Any]) -> pd.DataFrame:
        """
        Filter DataFrame based on conditions
        
        Args:
            conditions (dict): Dictionary with column names as keys and filter values
        
        Returns:
            pd.DataFrame: Filtered DataFrame
        """
        if self.df is None:
            print("No DataFrame loaded.")
            return None
        
        filtered_df = self.df.copy()
        
        for column, value in conditions.items():
            if column in filtered_df.columns:
                if isinstance(value, list):
                    filtered_df = filtered_df[filtered_df[column].isin(value)]
                else:
                    filtered_df = filtered_df[filtered_df[column] == value]
            else:
                print(f"Warning: Column '{column}' not found in DataFrame.")
        
        print(f"Filtered data shape: {filtered_df.shape}")
        return filtered_df
    
    def select_columns(self, columns: List[str]) -> pd.DataFrame:
        """
        Select specific columns from DataFrame
        
        Args:
            columns (list): List of column names to select
        
        Returns:
            pd.DataFrame: DataFrame with selected columns
        """
        if self.df is None:
            print("No DataFrame loaded.")
            return None
        
        try:
            selected_df = self.df[columns]
            print(f"Selected columns: {columns}")
            return selected_df
        except KeyError as e:
            print(f"Error: Column(s) not found - {str(e)}")
            return None
    
    def add_column(self, column_name: str, values: Union[List, pd.Series, Any]) -> bool:
        """
        Add a new column to the DataFrame
        
        Args:
            column_name (str): Name of the new column
            values: Values for the new column
        
        Returns:
            bool: True if successful, False otherwise
        """
        if self.df is None:
            print("No DataFrame loaded.")
            return False
        
        try:
            self.df[column_name] = values
            print(f"Successfully added column: {column_name}")
            return True
        except Exception as e:
            print(f"Error adding column: {str(e)}")
            return False
    
    def remove_columns(self, columns: List[str]) -> bool:
        """
        Remove columns from DataFrame
        
        Args:
            columns (list): List of column names to remove
        
        Returns:
            bool: True if successful, False otherwise
        """
        if self.df is None:
            print("No DataFrame loaded.")
            return False
        
        try:
            self.df = self.df.drop(columns=columns)
            print(f"Successfully removed columns: {columns}")
            return True
        except Exception as e:
            print(f"Error removing columns: {str(e)}")
            return False
    
    def clean_data(self) -> pd.DataFrame:
        """
        Basic data cleaning operations
        
        Returns:
            pd.DataFrame: Cleaned DataFrame
        """
        if self.df is None:
            print("No DataFrame loaded.")
            return None
        
        cleaned_df = self.df.copy()
        
        # Remove duplicates
        initial_rows = len(cleaned_df)
        cleaned_df = cleaned_df.drop_duplicates()
        duplicates_removed = initial_rows - len(cleaned_df)
        
        # Remove rows with all NaN values
        cleaned_df = cleaned_df.dropna(how='all')
        
        # Strip whitespace from string columns
        string_columns = cleaned_df.select_dtypes(include=['object']).columns
        for col in string_columns:
            cleaned_df[col] = cleaned_df[col].astype(str).str.strip()
        
        print(f"Data cleaning completed:")
        print(f"- Removed {duplicates_removed} duplicate rows")
        print(f"- Final shape: {cleaned_df.shape}")
        
        return cleaned_df
    
    def get_statistics(self) -> None:
        """Display statistical summary of the DataFrame"""
        if self.df is not None:
            print("=== Statistical Summary ===")
            print(self.df.describe(include='all'))
        else:
            print("No DataFrame loaded.")
    
    def search_data(self, column: str, search_term: str, case_sensitive: bool = False) -> pd.DataFrame:
        """
        Search for data containing a specific term
        
        Args:
            column (str): Column name to search in
            search_term (str): Term to search for
            case_sensitive (bool): Whether search should be case sensitive
        
        Returns:
            pd.DataFrame: Filtered DataFrame with matching rows
        """
        if self.df is None:
            print("No DataFrame loaded.")
            return None
        
        if column not in self.df.columns:
            print(f"Error: Column '{column}' not found.")
            return None
        
        try:
            if case_sensitive:
                mask = self.df[column].astype(str).str.contains(search_term, na=False)
            else:
                mask = self.df[column].astype(str).str.contains(search_term, case=False, na=False)
            
            result_df = self.df[mask]
            print(f"Found {len(result_df)} rows containing '{search_term}' in column '{column}'")
            return result_df
        except Exception as e:
            print(f"Error during search: {str(e)}")
            return None


class ExcelHandler:
    def __init__(self):
        self.workbook = None
        self.file_path = None
        self.active_sheet = None
    
    def create_workbook(self):
        """Create a new Excel workbook"""
        try:
            self.workbook = Workbook()
            self.active_sheet = self.workbook.active
            print("New workbook created successfully")
            return True
        except Exception as e:
            print(f"Error creating workbook: {str(e)}")
            return False
    
    def load_workbook(self, file_path: str):
        """
        Load an existing Excel workbook
        
        Args:
            file_path (str): Path to the Excel file
        """
        try:
            if not os.path.exists(file_path):
                print(f"Error: File '{file_path}' not found.")
                return False
            
            self.workbook = load_workbook(file_path)
            self.file_path = file_path
            self.active_sheet = self.workbook.active
            print(f"Successfully loaded workbook: {file_path}")
            print(f"Available sheets: {self.workbook.sheetnames}")
            return True
        except Exception as e:
            print(f"Error loading workbook: {str(e)}")
            return False
    
    def save_workbook(self, file_path: Optional[str] = None):
        """
        Save the workbook
        
        Args:
            file_path (str, optional): Path to save the file. If None, uses current file_path
        """
        try:
            if file_path:
                self.file_path = file_path
            elif not self.file_path:
                self.file_path = "output.xlsx"
            
            self.workbook.save(self.file_path)
            print(f"Workbook saved successfully: {self.file_path}")
            return True
        except Exception as e:
            print(f"Error saving workbook: {str(e)}")
            return False
    
    def get_sheet_names(self) -> List[str]:
        """Get list of all sheet names"""
        if self.workbook:
            return self.workbook.sheetnames
        else:
            print("No workbook loaded")
            return []
    
    def create_sheet(self, sheet_name: str, index: Optional[int] = None):
        """
        Create a new worksheet
        
        Args:
            sheet_name (str): Name of the new sheet
            index (int, optional): Position to insert the sheet
        """
        try:
            if index is not None:
                sheet = self.workbook.create_sheet(sheet_name, index)
            else:
                sheet = self.workbook.create_sheet(sheet_name)
            print(f"Sheet '{sheet_name}' created successfully")
            return sheet
        except Exception as e:
            print(f"Error creating sheet: {str(e)}")
            return None
    
    def select_sheet(self, sheet_name: str):
        """
        Select a worksheet to work with
        
        Args:
            sheet_name (str): Name of the sheet to select
        """
        try:
            if sheet_name in self.workbook.sheetnames:
                self.active_sheet = self.workbook[sheet_name]
                print(f"Selected sheet: {sheet_name}")
                return True
            else:
                print(f"Sheet '{sheet_name}' not found")
                return False
        except Exception as e:
            print(f"Error selecting sheet: {str(e)}")
            return False
    
    def delete_sheet(self, sheet_name: str):
        """
        Delete a worksheet
        
        Args:
            sheet_name (str): Name of the sheet to delete
        """
        try:
            if sheet_name in self.workbook.sheetnames:
                del self.workbook[sheet_name]
                print(f"Sheet '{sheet_name}' deleted successfully")
                return True
            else:
                print(f"Sheet '{sheet_name}' not found")
                return False
        except Exception as e:
            print(f"Error deleting sheet: {str(e)}")
            return False
    
    def write_cell(self, row: int, col: int, value: Any):
        """
        Write value to a specific cell
        
        Args:
            row (int): Row number (1-indexed)
            col (int): Column number (1-indexed)
            value: Value to write
        """
        try:
            self.active_sheet.cell(row=row, column=col, value=value)
            return True
        except Exception as e:
            print(f"Error writing to cell: {str(e)}")
            return False
    
    def write_cell_by_reference(self, cell_reference: str, value: Any):
        """
        Write value to a cell using Excel reference (e.g., 'A1')
        
        Args:
            cell_reference (str): Excel cell reference (e.g., 'A1', 'B2')
            value: Value to write
        """
        try:
            self.active_sheet[cell_reference] = value
            return True
        except Exception as e:
            print(f"Error writing to cell {cell_reference}: {str(e)}")
            return False
    
    def read_cell(self, row: int, col: int):
        """
        Read value from a specific cell
        
        Args:
            row (int): Row number (1-indexed)
            col (int): Column number (1-indexed)
        
        Returns:
            Cell value
        """
        try:
            return self.active_sheet.cell(row=row, column=col).value
        except Exception as e:
            print(f"Error reading cell: {str(e)}")
            return None
    
    def read_cell_by_reference(self, cell_reference: str):
        """
        Read value from a cell using Excel reference
        
        Args:
            cell_reference (str): Excel cell reference (e.g., 'A1', 'B2')
        
        Returns:
            Cell value
        """
        try:
            return self.active_sheet[cell_reference].value
        except Exception as e:
            print(f"Error reading cell {cell_reference}: {str(e)}")
            return None
    
    def write_row(self, row_num: int, data: List[Any], start_col: int = 1):
        """
        Write data to a row
        
        Args:
            row_num (int): Row number to write to
            data (list): List of values to write
            start_col (int): Starting column number
        """
        try:
            for i, value in enumerate(data):
                self.active_sheet.cell(row=row_num, column=start_col + i, value=value)
            return True
        except Exception as e:
            print(f"Error writing row: {str(e)}")
            return False
    
    def write_column(self, col_num: int, data: List[Any], start_row: int = 1):
        """
        Write data to a column
        
        Args:
            col_num (int): Column number to write to
            data (list): List of values to write
            start_row (int): Starting row number
        """
        try:
            for i, value in enumerate(data):
                self.active_sheet.cell(row=start_row + i, column=col_num, value=value)
            return True
        except Exception as e:
            print(f"Error writing column: {str(e)}")
            return False
    
    def write_data_from_dict(self, data: List[Dict], start_row: int = 1, start_col: int = 1, include_headers: bool = True):
        """
        Write data from a list of dictionaries
        
        Args:
            data (list): List of dictionaries
            start_row (int): Starting row
            start_col (int): Starting column
            include_headers (bool): Whether to include headers
        """
        try:
            if not data:
                print("No data to write")
                return False
            
            current_row = start_row
            
            # Write headers
            if include_headers:
                headers = list(data[0].keys())
                for i, header in enumerate(headers):
                    self.active_sheet.cell(row=current_row, column=start_col + i, value=header)
                current_row += 1
            
            # Write data
            for row_data in data:
                for i, (key, value) in enumerate(row_data.items()):
                    self.active_sheet.cell(row=current_row, column=start_col + i, value=value)
                current_row += 1
            
            print(f"Successfully wrote {len(data)} rows of data")
            return True
        except Exception as e:
            print(f"Error writing data: {str(e)}")
            return False
    
    def read_range(self, start_cell: str, end_cell: str) -> List[List]:
        """
        Read a range of cells
        
        Args:
            start_cell (str): Starting cell reference (e.g., 'A1')
            end_cell (str): Ending cell reference (e.g., 'C3')
        
        Returns:
            List of lists containing cell values
        """
        try:
            cell_range = self.active_sheet[f"{start_cell}:{end_cell}"]
            data = []
            for row in cell_range:
                row_data = [cell.value for cell in row]
                data.append(row_data)
            return data
        except Exception as e:
            print(f"Error reading range: {str(e)}")
            return []
    
    def get_sheet_dimensions(self) -> tuple:
        """
        Get the dimensions of the active sheet
        
        Returns:
            tuple: (max_row, max_column)
        """
        if self.active_sheet:
            return (self.active_sheet.max_row, self.active_sheet.max_column)
        return (0, 0)
    
    def format_cell(self, cell_reference: str, font_name: str = None, font_size: int = None, 
                   bold: bool = False, italic: bool = False, font_color: str = None,
                   bg_color: str = None, border: bool = False):
        """
        Format a cell with various styling options
        
        Args:
            cell_reference (str): Cell reference (e.g., 'A1')
            font_name (str): Font name
            font_size (int): Font size
            bold (bool): Bold text
            italic (bool): Italic text
            font_color (str): Font color (hex code)
            bg_color (str): Background color (hex code)
            border (bool): Add border
        """
        try:
            cell = self.active_sheet[cell_reference]
            
            # Font formatting
            font_kwargs = {}
            if font_name:
                font_kwargs['name'] = font_name
            if font_size:
                font_kwargs['size'] = font_size
            if bold:
                font_kwargs['bold'] = True
            if italic:
                font_kwargs['italic'] = True
            if font_color:
                font_kwargs['color'] = font_color
            
            if font_kwargs:
                cell.font = Font(**font_kwargs)
            
            # Background color
            if bg_color:
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            
            # Border
            if border:
                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                cell.border = thin_border
            
            return True
        except Exception as e:
            print(f"Error formatting cell: {str(e)}")
            return False
    
    def format_range(self, start_cell: str, end_cell: str, **format_kwargs):
        """
        Format a range of cells
        
        Args:
            start_cell (str): Starting cell reference
            end_cell (str): Ending cell reference
            **format_kwargs: Formatting options (same as format_cell)
        """
        try:
            cell_range = self.active_sheet[f"{start_cell}:{end_cell}"]
            for row in cell_range:
                for cell in row:
                    self.format_cell(cell.coordinate, **format_kwargs)
            return True
        except Exception as e:
            print(f"Error formatting range: {str(e)}")
            return False
    
    def auto_adjust_column_width(self, columns: List[str] = None):
        """
        Auto-adjust column widths
        
        Args:
            columns (list): List of column letters to adjust. If None, adjusts all columns.
        """
        try:
            if columns is None:
                columns = [get_column_letter(col) for col in range(1, self.active_sheet.max_column + 1)]
            
            for column in columns:
                max_length = 0
                for cell in self.active_sheet[column]:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                self.active_sheet.column_dimensions[column].width = adjusted_width
            
            print(f"Adjusted width for columns: {columns}")
            return True
        except Exception as e:
            print(f"Error adjusting column width: {str(e)}")
            return False
    
    def add_formula(self, cell_reference: str, formula: str):
        """
        Add a formula to a cell
        
        Args:
            cell_reference (str): Cell reference (e.g., 'A1')
            formula (str): Excel formula (e.g., '=SUM(A1:A10)')
        """
        try:
            self.active_sheet[cell_reference] = formula
            return True
        except Exception as e:
            print(f"Error adding formula: {str(e)}")
            return False
    
    def insert_image(self, image_path: str, cell_reference: str):
        """
        Insert an image into the worksheet
        
        Args:
            image_path (str): Path to the image file
            cell_reference (str): Cell reference where to place the image
        """
        try:
            from openpyxl.drawing.image import Image
            
            if not os.path.exists(image_path):
                print(f"Image file not found: {image_path}")
                return False
            
            img = Image(image_path)
            self.active_sheet.add_image(img, cell_reference)
            print(f"Image inserted at {cell_reference}")
            return True
        except ImportError:
            print("Pillow library is required for image insertion. Install it with: pip install Pillow")
            return False
        except Exception as e:
            print(f"Error inserting image: {str(e)}")
            return False
    
    def create_table(self, data: List[List], start_cell: str = "A1", table_name: str = "Table1", 
                    style: str = "TableStyleMedium9"):
        """
        Create a formatted table
        
        Args:
            data (list): List of lists containing table data
            start_cell (str): Starting cell for the table
            table_name (str): Name of the table
            style (str): Table style
        """
        try:
            from openpyxl.worksheet.table import Table, TableStyleInfo
            
            # Write data
            start_row, start_col = openpyxl.utils.cell.coordinate_to_tuple(start_cell)
            for i, row in enumerate(data):
                for j, value in enumerate(row):
                    self.active_sheet.cell(row=start_row + i, column=start_col + j, value=value)
            
            # Create table
            end_cell = get_column_letter(start_col + len(data[0]) - 1) + str(start_row + len(data) - 1)
            table_range = f"{start_cell}:{end_cell}"
            
            table = Table(displayName=table_name, ref=table_range)
            style = TableStyleInfo(name=style, showFirstColumn=False,
                                 showLastColumn=False, showRowStripes=True, showColumnStripes=True)
            table.tableStyleInfo = style
            
            self.active_sheet.add_table(table)
            print(f"Table '{table_name}' created successfully")
            return True
        except Exception as e:
            print(f"Error creating table: {str(e)}")
            return False
    
    def get_workbook_info(self):
        """Display information about the workbook"""
        if self.workbook:
            print("=== Workbook Information ===")
            print(f"File path: {self.file_path}")
            print(f"Active sheet: {self.active_sheet.title}")
            print(f"Total sheets: {len(self.workbook.sheetnames)}")
            print(f"Sheet names: {self.workbook.sheetnames}")
            
            if self.active_sheet:
                max_row, max_col = self.get_sheet_dimensions()
                print(f"Active sheet dimensions: {max_row} rows x {max_col} columns")
        else:
            print("No workbook loaded")