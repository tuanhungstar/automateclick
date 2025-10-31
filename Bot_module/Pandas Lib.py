import pandas as pd
import clipboard
import os # Used for the example

class Excel:
    """
    A class to read and interact with Excel files using a pandas DataFrame.
    """
    
    def __init__(self):
        """
        Initializes the ExcelHelper.
        """
        df = None
        self.original_file_link = None
        self.original_sheetname = None

    def read_excel(self, file_link: str, sheetname: str | int) -> pd.DataFrame:
        """
        Reads a specific sheet from an Excel file into a pandas DataFrame.

        Args:
            file_link: The file path or URL to the Excel file.
            sheetname: The name (str) or index (int) of the sheet to read.

        Returns:
            The pandas DataFrame containing the sheet's data.
        """
        try:
            df = pd.read_excel(file_link, sheet_name=sheetname)
            self.original_file_link = file_link
            self.original_sheetname = sheetname
            print(f"Successfully read '{sheetname}' from '{file_link}'.")
            return df
        except FileNotFoundError:
            print(f"Error: File not found at '{file_link}'")
            return df
        except Exception as e:
            # Catches other errors like 'sheet not found'
            print(f"An error occurred: {e}")
            return None

    def _check_df(self,df) -> bool:
        """Internal helper to check if a DataFrame is loaded."""
        if df is None:
            print("Error: No DataFrame loaded. Please call 'read_excel' first.")
            return False
        return True
        
    def _col_letter_to_index(self, letter: str) -> int:
            """
            Converts an Excel column letter (e.g., 'A', 'B', 'AA') to a 0-based index.
            """
            index = 0
            for char in letter.upper():
                if not 'A' <= char <= 'Z':
                    raise ValueError(f"Invalid column letter '{letter}'")
                index = index * 26 + (ord(char) - ord('A')) + 1
            return index - 1 # Adjust to 0-based index
            
    def put_range_to_clipboard(self, df,
                                 start_row: int, 
                                 end_row: int, 
                                 start_col_letter: str="A", 
                                 end_col_letter: str="A",
                                 index: bool = False,
                                 header: bool = False):
            """
            Copies a specified range of the DataFrame to the clipboard.
            Row indices are 0-based and [start_row:end_row] (exclusive of end_row).
            Column letters are 1-based (e.g., 'A', 'B') and inclusive.

            Args:
                start_row: The starting row index (0-based).
                end_row: The ending row index (exclusive).
                start_col_letter: The starting column letter (e.g., 'A').
                end_col_letter: The ending column letter (e.g., 'C'). Inclusive.
                index: Whether to include the DataFrame index in the clipboard.
                header: Whether to include the column headers in the clipboard.
            """
            if not self._check_df(df):
                return

            try:
                # Convert column letters to 0-based integer indices
                start_col_idx = self._col_letter_to_index(start_col_letter)
                end_col_idx = self._col_letter_to_index(end_col_letter)

                # Ensure start_col is not after end_col
                if start_col_idx > end_col_idx:
                    print(f"Error: Start column '{start_col_letter}' is after end column '{end_col_letter}'.")
                    return

                # Select the sub-dataframe using iloc
                # +1 on end_col_idx because Excel ranges are inclusive, 
                # but Python slicing is exclusive.
                sub_df = df.iloc[start_row:end_row, start_col_idx : end_col_idx + 1]
                
                # Use the built-in pandas to_clipboard method
                sub_df.to_clipboard(index=index, header=header)
                
                print(f"Copied range [rows {start_row}-{end_row-1}, cols {start_col_letter}-{end_col_letter}] to clipboard.")
            
            except (IndexError, ValueError) as e:
                print(f"Error: The specified range is invalid. Details: {e}")
            except Exception as e:
                print(f"An error occurred while copying to clipboard: {e}")

    def read_cell(self,df, row_idx: int, col_idx: int) -> any:
        """
        Reads the value of a single cell by its integer index.

        Args:
            row_idx: The row index (0-based).
            col_idx: The column index (0-based).

        Returns:
            The value of the cell, or None if an error occurs.
        """
        if not self._check_df(df):
            return None
        
        try:
            # .iat is the fastest way to access a single cell by integer location
            value = df.iat[row_idx, col_idx]
            return value
        except IndexError:
            print(f"Error: Cell at ({row_idx}, {col_idx}) is out of bounds.")
            return None

    def write_cell(self,df, row_idx: int, col_idx: int, value: any):
        """
        Writes a value to a single cell by its integer index.

        Args:
            row_idx: The row index (0-based).
            col_idx: The column index (0-based).
            value: The value to write to the cell.
        """
        if not self._check_df(df):
            return
            
        try:
            # .iat is the fastest way to write to a single cell by integer location
            df.iat[row_idx, col_idx] = value
            print(f"Set cell ({row_idx}, {col_idx}) to '{value}'.")
            return df
        except IndexError:
            print(f"Error: Cell at ({row_idx}, {col_idx}) is out of bounds.")


            
    def save_excel(self, df, save_link: str = None, sheet_name: str = None, index: bool = False):
        """
        Saves the current DataFrame (self.df) to an Excel file.

        Args:
            save_path: The file path to save to. **If None, overwrites the
                       original file.**
            sheet_name: The sheet name to save to. **If None, uses the
                        original sheet name.**
            index: Whether to write the DataFrame's index as a column.
                   Defaults to False.
        """
        if not self._check_df():
            return

        # Determine save path
        path_to_save = save_link if save_link else self.original_file_link
        if not path_to_save:
            print("Error: No save path specified and no original file path available.")
            return

        # Determine sheet name
        sheet_to_save = sheet_name if sheet_name else self.original_sheetname
        if not sheet_to_save:
            sheet_to_save = 'Sheet1' # Default fallback

        try:
            self.df.to_excel(path_to_save, sheet_name=sheet_to_save, index=index)
            print(f"DataFrame successfully saved to '{path_to_save}' on sheet '{sheet_to_save}'.")
        except Exception as e:
            print(f"An error occurred while saving: {e}")

    # --- NEW METHOD ---
    def write_df_back_to_original(self, index: bool = False):
        """
        Saves the current DataFrame back to the original file and sheet.

        This is a convenience method that calls save_excel() with
        no 'save_path' or 'sheet_name', forcing it to use the
        original file details.

        Args:
            index: Whether to write the DataFrame's index as a column.
                   Defaults to False.
        """
        print("Attempting to save DataFrame back to original file...")
        # Calls save_excel with save_path=None and sheet_name=None
        self.save_excel(index=index)
        
    def filter_by_value(self,df, column_name: str, value_to_match: object) -> pd.DataFrame:
            """
            Filters the DataFrame to find rows where a specific column matches a given value.
            
            Args:
                column_name (str): The name of the column to filter.
                value_to_match (object): The value to match in the column.
            
            Returns:
                pd.DataFrame: A new DataFrame containing only the matching rows.
            """
            try:
                # This is the core filtering logic
                filtered_df = df[df[column_name] == value_to_match].copy()
                
                if filtered_df.empty:
                    print(f"Warning: No rows found where '{column_name}' == {value_to_match}")
                    
                return filtered_df
            
            except KeyError:
                print(f"Error: Column '{column_name}' not found in the DataFrame.")
                return pd.DataFrame() # Return an empty DataFrame on error
            except Exception as e:
                print(f"An error occurred: {e}")
                return pd.DataFrame()
                
    def read_csv(self, file_link: str) -> pd.DataFrame:
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
            df = pd.read_csv(file_link)
            return df
        except FileNotFoundError:
            print(f"Error: File '{file_link}' not found.")
            return None
        except Exception as e:
            print(f"Error reading CSV: {str(e)}")
            return None
            
    def df_len(self, df) -> int:
        return len(df)