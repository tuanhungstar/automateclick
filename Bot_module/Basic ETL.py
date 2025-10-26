import pandas as pd
import numpy as np
import os
from typing import List, Dict, Any, Optional, Union, Tuple
import json
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

class SimpleETL:
    """
    A simple ETL (Extract, Transform, Load) class for basic data operations.
    
    This class provides methods to extract data from various sources, transform it
    using common data cleaning and manipulation operations, and load it to different
    output formats.
    
    Attributes:
        data (pd.DataFrame): Current working dataset
        original_data (pd.DataFrame): Original dataset backup
        transformations_log (List[str]): Log of all transformations performed
    
    Example:
        >>> etl = SimpleETL()
        >>> df = etl.extract_from_csv('data.csv')
        >>> etl.clean_missing_values('fill')
        >>> etl.load_to_excel('output.xlsx')
    """
    
    def __init__(self):
        """
        Initialize the SimpleETL instance.
        
        Sets up empty data containers and transformation log.
        """
        self.data = None
        self.original_data = None
        self.transformations_log = []
    
    # EXTRACT Methods
    def extract_from_csv(self, file_link: str, **kwargs) -> Optional[pd.DataFrame]:
        """
        Extract data from a CSV file.
        
        Args:
            file_link (str): Path to the CSV file to read
            **kwargs: Additional keyword arguments passed to pandas.read_csv()
                     Common options include:
                     - sep: Field delimiter (default ',')
                     - header: Row number to use as column names (default 0)
                     - encoding: File encoding (e.g., 'utf-8', 'latin-1')
                     - skiprows: Number of rows to skip at the beginning
        
        Returns:
            pd.DataFrame: Extracted DataFrame if successful, None if failed
        
        Example:
            >>> df = etl.extract_from_csv('data.csv', encoding='utf-8', skiprows=1)
            >>> print(df.shape)
            (1000, 5)
        """
        try:
            self.data = pd.read_csv(file_link, **kwargs)
            self.original_data = self.data.copy()
            self._log_operation(f"Extracted data from CSV: {file_link}")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error extracting from CSV: {str(e)}")
            return None
    
    def extract_from_excel(self, file_link: str, sheet_name: str = None, **kwargs) -> Optional[pd.DataFrame]:
        """
        Extract data from an Excel file.
        
        Args:
            file_link (str): Path to the Excel file to read
            sheet_name (str, optional): Name of the sheet to read. If None, reads the first sheet
            **kwargs: Additional keyword arguments passed to pandas.read_excel()
                     Common options include:
                     - header: Row number to use as column names
                     - skiprows: Number of rows to skip
                     - usecols: Columns to read (list of column names or indices)
        
        Returns:
            pd.DataFrame: Extracted DataFrame if successful, None if failed
        
        Example:
            >>> df = etl.extract_from_excel('data.xlsx', sheet_name='Sales', header=1)
            >>> print(df.columns.tolist())
            ['Product', 'Sales', 'Date', 'Region']
        """
        try:
            self.data = pd.read_excel(file_link, sheet_name=sheet_name, **kwargs)
            self.original_data = self.data.copy()
            self._log_operation(f"Extracted data from Excel: {file_link}, sheet: {sheet_name}")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error extracting from Excel: {str(e)}")
            return None
    
    def extract_from_json(self, file_link: str, **kwargs) -> Optional[pd.DataFrame]:
        """
        Extract data from a JSON file.
        
        Args:
            file_link (str): Path to the JSON file to read
            **kwargs: Additional keyword arguments passed to pandas.read_json()
                     Common options include:
                     - orient: Expected JSON format ('records', 'index', 'values', etc.)
                     - lines: Read file as line-delimited JSON
        
        Returns:
            pd.DataFrame: Extracted DataFrame if successful, None if failed
        
        Example:
            >>> df = etl.extract_from_json('data.json', orient='records')
            >>> print(df.info())
        """
        try:
            self.data = pd.read_json(file_link, **kwargs)
            self.original_data = self.data.copy()
            self._log_operation(f"Extracted data from JSON: {file_link}")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error extracting from JSON: {str(e)}")
            return None
    
    def extract_from_database(self, connection_string: str, query: str) -> Optional[pd.DataFrame]:
        """
        Extract data from a database using SQL query.
        
        Currently supports SQLite databases. Can be extended for other database types.
        
        Args:
            connection_string (str): Database connection string (SQLite file path)
            query (str): SQL query to execute
        
        Returns:
            pd.DataFrame: Extracted DataFrame if successful, None if failed
        
        Example:
            >>> df = etl.extract_from_database('database.db', 'SELECT * FROM customers WHERE age > 25')
            >>> print(f"Retrieved {len(df)} records")
        """
        try:
            import sqlite3
            # Simple SQLite example - can be extended for other databases
            conn = sqlite3.connect(connection_string)
            self.data = pd.read_sql_query(query, conn)
            conn.close()
            self.original_data = self.data.copy()
            self._log_operation(f"Extracted data from database with query: {query[:50]}...")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error extracting from database: {str(e)}")
            return None
    
    def extract_sample_data(self, rows: int = 100) -> pd.DataFrame:
        """
        Generate sample data for testing purposes.
        
        Creates a dataset with typical business data including ID, name, age, 
        salary, department, and join date fields.
        
        Args:
            rows (int): Number of rows to generate (default: 100)
        
        Returns:
            pd.DataFrame: Generated sample DataFrame
        
        Example:
            >>> df = etl.extract_sample_data(500)
            >>> print(df.head())
            >>> print(f"Generated {len(df)} rows with {len(df.columns)} columns")
        """
        try:
            np.random.seed(42)
            self.data = pd.DataFrame({
                'id': range(1, rows + 1),
                'name': [f'Person_{i}' for i in range(1, rows + 1)],
                'age': np.random.randint(18, 65, rows),
                'salary': np.random.randint(30000, 100000, rows),
                'department': np.random.choice(['IT', 'HR', 'Finance', 'Marketing'], rows),
                'join_date': pd.date_range('2020-01-01', periods=rows, freq='D')
            })
            self.original_data = self.data.copy()
            self._log_operation(f"Generated sample data with {rows} rows")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error generating sample data: {str(e)}")
            return pd.DataFrame()
    
    def extract_from_url(self, url: str, file_type: str = 'csv', **kwargs) -> Optional[pd.DataFrame]:
        """
        Extract data from a URL.
        
        Args:
            url (str): URL to download data from
            file_type (str): Type of file to expect ('csv', 'json', 'excel')
            **kwargs: Additional keyword arguments passed to the respective pandas read function
        
        Returns:
            pd.DataFrame: Extracted DataFrame if successful, None if failed
        
        Example:
            >>> df = etl.extract_from_url('https://example.com/data.csv', file_type='csv')
            >>> print(df.shape)
        """
        try:
            if file_type.lower() == 'csv':
                self.data = pd.read_csv(url, **kwargs)
            elif file_type.lower() == 'json':
                self.data = pd.read_json(url, **kwargs)
            elif file_type.lower() in ['excel', 'xlsx', 'xls']:
                self.data = pd.read_excel(url, **kwargs)
            else:
                print(f"✗ Unsupported file type: {file_type}")
                return None
            
            self.original_data = self.data.copy()
            self._log_operation(f"Extracted data from URL: {url}")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error extracting from URL: {str(e)}")
            return None
    
    def extract_from_clipboard(self) -> Optional[pd.DataFrame]:
        """
        Extract data from clipboard (copied data).
        
        Useful for quickly importing data copied from Excel, web pages, etc.
        
        Returns:
            pd.DataFrame: Extracted DataFrame if successful, None if failed
        
        Example:
            >>> # Copy some tabular data to clipboard first
            >>> df = etl.extract_from_clipboard()
            >>> print(df.head())
        """
        try:
            self.data = pd.read_clipboard()
            self.original_data = self.data.copy()
            self._log_operation("Extracted data from clipboard")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error extracting from clipboard: {str(e)}")
            return None
    
    def extract_from_multiple_csv(self, file_links: List[str], **kwargs) -> Optional[pd.DataFrame]:
        """
        Extract and combine data from multiple CSV files.
        
        Args:
            file_links (List[str]): List of CSV file paths to read and combine
            **kwargs: Additional keyword arguments passed to pandas.read_csv()
        
        Returns:
            pd.DataFrame: Combined DataFrame if successful, None if failed
        
        Example:
            >>> files = ['data1.csv', 'data2.csv', 'data3.csv']
            >>> df = etl.extract_from_multiple_csv(files)
            >>> print(f"Combined {len(files)} files into {len(df)} rows")
        """
        try:
            dataframes = []
            for file_link in file_links:
                if os.path.exists(file_link):
                    df = pd.read_csv(file_link, **kwargs)
                    # Add source file column
                    df['source_file'] = os.path.basename(file_link)
                    dataframes.append(df)
                else:
                    print(f"⚠ File not found: {file_link}")
            
            if not dataframes:
                print("✗ No valid files found")
                return None
            
            self.data = pd.concat(dataframes, ignore_index=True)
            self.original_data = self.data.copy()
            self._log_operation(f"Extracted and combined data from {len(dataframes)} CSV files")
            return self.data.copy()
        except Exception as e:
            print(f"✗ Error extracting from multiple CSV files: {str(e)}")
            return None
    
    # Additional utility method for extract operations
    def get_current_data(self) -> Optional[pd.DataFrame]:
        """
        Get a copy of the current dataset.
        
        Returns:
            pd.DataFrame: Copy of current dataset if available, None otherwise
        
        Example:
            >>> df = etl.get_current_data()
            >>> if df is not None:
            ...     print(f"Current data shape: {df.shape}")
        """
        if self.data is not None:
            return self.data.copy()
        return None
    
    # TRANSFORM Methods (keeping existing implementation)
    def clean_missing_values(self, strategy: str = 'drop', fill_value: Any = None) -> bool:
        """
        Handle missing values in the dataset.
        
        Args:
            strategy (str): Method to handle missing values:
                          - 'drop': Remove rows with any missing values
                          - 'fill': Fill missing values with specified or default values
            fill_value (Any, optional): Value to use for filling. If None and strategy is 'fill',
                                       uses mean for numeric columns and 'Unknown' for others
        
        Returns:
            bool: True if operation successful, False otherwise
        
        Example:
            >>> etl.clean_missing_values('fill', fill_value=0)
            ✓ Cleaned missing values: 1000 → 1000 rows
            True
        """
        if self.data is None:
            print("✗ No data loaded")
            return False
        
        try:
            initial_rows = len(self.data)
            
            if strategy == 'drop':
                self.data = self.data.dropna()
            elif strategy == 'fill':
                if fill_value is not None:
                    self.data = self.data.fillna(fill_value)
                else:
                    # Fill with appropriate defaults
                    for col in self.data.columns:
                        if self.data[col].dtype in ['int64', 'float64']:
                            self.data[col] = self.data[col].fillna(self.data[col].mean())
                        else:
                            self.data[col] = self.data[col].fillna('Unknown')
            
            final_rows = len(self.data)
            self._log_operation(f"Cleaned missing values using '{strategy}' strategy")
            print(f"✓ Cleaned missing values: {initial_rows} → {final_rows} rows")
            return True
        except Exception as e:
            print(f"✗ Error cleaning missing values: {str(e)}")
            return False
    
    # ... (rest of the transform and load methods remain the same as in the previous version)
    
    def _log_operation(self, operation: str):
        """
        Internal method to log operations with timestamps.
        
        Args:
            operation (str): Description of the operation performed
        
        Note:
            This is an internal method used by other methods to maintain
            a log of transformations. Users typically don't call this directly.
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.transformations_log.append(f"{timestamp} - {operation}")




class DataFrameMerger:
    """
    A comprehensive class for merging, joining, and concatenating pandas DataFrames.
    
    This class provides various methods to combine DataFrames using different strategies
    including SQL-style joins, concatenation, and advanced merging operations with
    data validation and conflict resolution.
    
    Attributes:
        merge_log (List[str]): Log of all merge operations performed
        last_operation (Dict): Details of the last operation performed
        
    Example:
        >>> merger = DataFrameMerger()
        >>> df1 = pd.DataFrame({'id': [1, 2, 3], 'name': ['A', 'B', 'C']})
        >>> df2 = pd.DataFrame({'id': [1, 2, 4], 'value': [10, 20, 30]})
        >>> result = merger.inner_join(df1, df2, on='id')
        >>> print(result.shape)
        (2, 3)
    """
    
    def __init__(self):
        """
        Initialize the DataFrameMerger instance.
        
        Sets up logging for tracking merge operations and stores operation details.
        """
        self.merge_log = []
        self.last_operation = {}
    
    def inner_join(self, left_df: pd.DataFrame, right_df: pd.DataFrame, 
                   on: Union[str, List[str]] = None, 
                   left_on: Union[str, List[str]] = None,
                   right_on: Union[str, List[str]] = None,
                   suffixes: Tuple[str, str] = ('_left', '_right')) -> pd.DataFrame:
        """
        Perform an inner join between two DataFrames.
        
        Returns only rows that have matching keys in both DataFrames.
        
        Args:
            left_df (pd.DataFrame): Left DataFrame to join
            right_df (pd.DataFrame): Right DataFrame to join
            on (Union[str, List[str]], optional): Column name(s) to join on (must exist in both DataFrames)
            left_on (Union[str, List[str]], optional): Column name(s) to join on in left DataFrame
            right_on (Union[str, List[str]], optional): Column name(s) to join on in right DataFrame
            suffixes (Tuple[str, str]): Suffixes to add to overlapping column names (default: ('_left', '_right'))
        
        Returns:
            pd.DataFrame: Result of the inner join operation
            
        Raises:
            ValueError: If neither 'on' nor both 'left_on' and 'right_on' are specified
            
        Example:
            >>> df1 = pd.DataFrame({'id': [1, 2, 3], 'name': ['Alice', 'Bob', 'Charlie']})
            >>> df2 = pd.DataFrame({'id': [1, 2, 4], 'salary': [50000, 60000, 70000]})
            >>> result = merger.inner_join(df1, df2, on='id')
            >>> print(result)
               id    name  salary
            0   1   Alice   50000
            1   2     Bob   60000
        """
        try:
            if on is None and (left_on is None or right_on is None):
                raise ValueError("Must specify either 'on' or both 'left_on' and 'right_on'")
            
            result = pd.merge(left_df, right_df, how='inner', on=on, 
                            left_on=left_on, right_on=right_on, suffixes=suffixes)
            
            operation_details = {
                'operation': 'inner_join',
                'left_shape': left_df.shape,
                'right_shape': right_df.shape,
                'result_shape': result.shape,
                'join_keys': on or f"{left_on} = {right_on}",
                'rows_matched': len(result)
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in inner join: {str(e)}")
            return pd.DataFrame()
    
    def left_join(self, left_df: pd.DataFrame, right_df: pd.DataFrame,
                  on: Union[str, List[str]] = None,
                  left_on: Union[str, List[str]] = None,
                  right_on: Union[str, List[str]] = None,
                  suffixes: Tuple[str, str] = ('_left', '_right')) -> pd.DataFrame:
        """
        Perform a left join between two DataFrames.
        
        Returns all rows from the left DataFrame and matching rows from the right DataFrame.
        Non-matching rows from the right DataFrame are excluded.
        
        Args:
            left_df (pd.DataFrame): Left DataFrame to join (all rows will be preserved)
            right_df (pd.DataFrame): Right DataFrame to join
            on (Union[str, List[str]], optional): Column name(s) to join on (must exist in both DataFrames)
            left_on (Union[str, List[str]], optional): Column name(s) to join on in left DataFrame
            right_on (Union[str, List[str]], optional): Column name(s) to join on in right DataFrame
            suffixes (Tuple[str, str]): Suffixes to add to overlapping column names (default: ('_left', '_right'))
        
        Returns:
            pd.DataFrame: Result of the left join operation
            
        Example:
            >>> df1 = pd.DataFrame({'id': [1, 2, 3], 'name': ['Alice', 'Bob', 'Charlie']})
            >>> df2 = pd.DataFrame({'id': [1, 2, 4], 'salary': [50000, 60000, 70000]})
            >>> result = merger.left_join(df1, df2, on='id')
            >>> print(result)
               id     name   salary
            0   1    Alice  50000.0
            1   2      Bob  60000.0
            2   3  Charlie      NaN
        """
        try:
            if on is None and (left_on is None or right_on is None):
                raise ValueError("Must specify either 'on' or both 'left_on' and 'right_on'")
            
            result = pd.merge(left_df, right_df, how='left', on=on,
                            left_on=left_on, right_on=right_on, suffixes=suffixes)
            
            operation_details = {
                'operation': 'left_join',
                'left_shape': left_df.shape,
                'right_shape': right_df.shape,
                'result_shape': result.shape,
                'join_keys': on or f"{left_on} = {right_on}",
                'null_values_added': result.isnull().sum().sum() - left_df.isnull().sum().sum()
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in left join: {str(e)}")
            return pd.DataFrame()
    
    def right_join(self, left_df: pd.DataFrame, right_df: pd.DataFrame,
                   on: Union[str, List[str]] = None,
                   left_on: Union[str, List[str]] = None,
                   right_on: Union[str, List[str]] = None,
                   suffixes: Tuple[str, str] = ('_left', '_right')) -> pd.DataFrame:
        """
        Perform a right join between two DataFrames.
        
        Returns all rows from the right DataFrame and matching rows from the left DataFrame.
        Non-matching rows from the left DataFrame are excluded.
        
        Args:
            left_df (pd.DataFrame): Left DataFrame to join
            right_df (pd.DataFrame): Right DataFrame to join (all rows will be preserved)
            on (Union[str, List[str]], optional): Column name(s) to join on (must exist in both DataFrames)
            left_on (Union[str, List[str]], optional): Column name(s) to join on in left DataFrame
            right_on (Union[str, List[str]], optional): Column name(s) to join on in right DataFrame
            suffixes (Tuple[str, str]): Suffixes to add to overlapping column names (default: ('_left', '_right'))
        
        Returns:
            pd.DataFrame: Result of the right join operation
            
        Example:
            >>> df1 = pd.DataFrame({'id': [1, 2, 3], 'name': ['Alice', 'Bob', 'Charlie']})
            >>> df2 = pd.DataFrame({'id': [1, 2, 4], 'salary': [50000, 60000, 70000]})
            >>> result = merger.right_join(df1, df2, on='id')
            >>> print(result)
               id     name  salary
            0   1    Alice   50000
            1   2      Bob   60000
            2   4      NaN   70000
        """
        try:
            if on is None and (left_on is None or right_on is None):
                raise ValueError("Must specify either 'on' or both 'left_on' and 'right_on'")
            
            result = pd.merge(left_df, right_df, how='right', on=on,
                            left_on=left_on, right_on=right_on, suffixes=suffixes)
            
            operation_details = {
                'operation': 'right_join',
                'left_shape': left_df.shape,
                'right_shape': right_df.shape,
                'result_shape': result.shape,
                'join_keys': on or f"{left_on} = {right_on}",
                'null_values_added': result.isnull().sum().sum() - right_df.isnull().sum().sum()
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in right join: {str(e)}")
            return pd.DataFrame()
    
    def outer_join(self, left_df: pd.DataFrame, right_df: pd.DataFrame,
                   on: Union[str, List[str]] = None,
                   left_on: Union[str, List[str]] = None,
                   right_on: Union[str, List[str]] = None,
                   suffixes: Tuple[str, str] = ('_left', '_right')) -> pd.DataFrame:
        """
        Perform an outer (full) join between two DataFrames.
        
        Returns all rows from both DataFrames. Missing values are filled with NaN
        where no match is found.
        
        Args:
            left_df (pd.DataFrame): Left DataFrame to join
            right_df (pd.DataFrame): Right DataFrame to join
            on (Union[str, List[str]], optional): Column name(s) to join on (must exist in both DataFrames)
            left_on (Union[str, List[str]], optional): Column name(s) to join on in left DataFrame
            right_on (Union[str, List[str]], optional): Column name(s) to join on in right DataFrame
            suffixes (Tuple[str, str]): Suffixes to add to overlapping column names (default: ('_left', '_right'))
        
        Returns:
            pd.DataFrame: Result of the outer join operation
            
        Example:
            >>> df1 = pd.DataFrame({'id': [1, 2, 3], 'name': ['Alice', 'Bob', 'Charlie']})
            >>> df2 = pd.DataFrame({'id': [1, 2, 4], 'salary': [50000, 60000, 70000]})
            >>> result = merger.outer_join(df1, df2, on='id')
            >>> print(result)
               id     name   salary
            0   1    Alice  50000.0
            1   2      Bob  60000.0
            2   3  Charlie      NaN
            3   4      NaN  70000.0
        """
        try:
            if on is None and (left_on is None or right_on is None):
                raise ValueError("Must specify either 'on' or both 'left_on' and 'right_on'")
            
            result = pd.merge(left_df, right_df, how='outer', on=on,
                            left_on=left_on, right_on=right_on, suffixes=suffixes)
            
            operation_details = {
                'operation': 'outer_join',
                'left_shape': left_df.shape,
                'right_shape': right_df.shape,
                'result_shape': result.shape,
                'join_keys': on or f"{left_on} = {right_on}",
                'total_unique_keys': len(result)
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in outer join: {str(e)}")
            return pd.DataFrame()
    
    def concat_vertical(self, dataframes: List[pd.DataFrame], 
                       ignore_index: bool = True,
                       keys: List[str] = None,
                       sort: bool = False) -> pd.DataFrame:
        """
        Concatenate DataFrames vertically (row-wise).
        
        Stacks DataFrames on top of each other. Columns must have the same names
        or they will be aligned by name.
        
        Args:
            dataframes (List[pd.DataFrame]): List of DataFrames to concatenate
            ignore_index (bool): Whether to ignore the original indices and create a new one (default: True)
            keys (List[str], optional): List of keys to add a hierarchical index identifying each DataFrame
            sort (bool): Whether to sort the column names (default: False)
        
        Returns:
            pd.DataFrame: Vertically concatenated DataFrame
            
        Example:
            >>> df1 = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [25, 30]})
            >>> df2 = pd.DataFrame({'name': ['Charlie', 'David'], 'age': [35, 40]})
            >>> result = merger.concat_vertical([df1, df2])
            >>> print(result)
                  name  age
            0    Alice   25
            1      Bob   30
            2  Charlie   35
            3    David   40
        """
        try:
            if not dataframes:
                raise ValueError("List of DataFrames cannot be empty")
            
            result = pd.concat(dataframes, axis=0, ignore_index=ignore_index, 
                             keys=keys, sort=sort)
            
            total_rows = sum(df.shape[0] for df in dataframes)
            operation_details = {
                'operation': 'concat_vertical',
                'num_dataframes': len(dataframes),
                'input_shapes': [df.shape for df in dataframes],
                'result_shape': result.shape,
                'total_input_rows': total_rows,
                'keys_used': keys is not None
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in vertical concatenation: {str(e)}")
            return pd.DataFrame()
    
    def concat_horizontal(self, dataframes: List[pd.DataFrame],
                         ignore_index: bool = False,
                         keys: List[str] = None,
                         sort: bool = False) -> pd.DataFrame:
        """
        Concatenate DataFrames horizontally (column-wise).
        
        Places DataFrames side by side. Rows are aligned by index.
        
        Args:
            dataframes (List[pd.DataFrame]): List of DataFrames to concatenate
            ignore_index (bool): Whether to ignore the original column names (default: False)
            keys (List[str], optional): List of keys to add a hierarchical column index
            sort (bool): Whether to sort the row indices (default: False)
        
        Returns:
            pd.DataFrame: Horizontally concatenated DataFrame
            
        Example:
            >>> df1 = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [25, 30]})
            >>> df2 = pd.DataFrame({'salary': [50000, 60000], 'dept': ['IT', 'HR']})
            >>> result = merger.concat_horizontal([df1, df2])
            >>> print(result)
                name  age  salary dept
            0  Alice   25   50000   IT
            1    Bob   30   60000   HR
        """
        try:
            if not dataframes:
                raise ValueError("List of DataFrames cannot be empty")
            
            result = pd.concat(dataframes, axis=1, ignore_index=ignore_index,
                             keys=keys, sort=sort)
            
            total_cols = sum(df.shape[1] for df in dataframes)
            operation_details = {
                'operation': 'concat_horizontal',
                'num_dataframes': len(dataframes),
                'input_shapes': [df.shape for df in dataframes],
                'result_shape': result.shape,
                'total_input_cols': total_cols,
                'keys_used': keys is not None
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in horizontal concatenation: {str(e)}")
            return pd.DataFrame()
    
    def merge_on_index(self, left_df: pd.DataFrame, right_df: pd.DataFrame,
                      how: str = 'inner',
                      suffixes: Tuple[str, str] = ('_left', '_right')) -> pd.DataFrame:
        """
        Merge DataFrames using their indices as the join key.
        
        Args:
            left_df (pd.DataFrame): Left DataFrame to merge
            right_df (pd.DataFrame): Right DataFrame to merge
            how (str): Type of merge ('inner', 'left', 'right', 'outer') (default: 'inner')
            suffixes (Tuple[str, str]): Suffixes for overlapping column names (default: ('_left', '_right'))
        
        Returns:
            pd.DataFrame: Merged DataFrame based on indices
            
        Example:
            >>> df1 = pd.DataFrame({'name': ['Alice', 'Bob']}, index=[1, 2])
            >>> df2 = pd.DataFrame({'salary': [50000, 60000]}, index=[1, 3])
            >>> result = merger.merge_on_index(df1, df2, how='left')
            >>> print(result)
                name   salary
            1  Alice  50000.0
            2    Bob      NaN
        """
        try:
            result = pd.merge(left_df, right_df, how=how, left_index=True, 
                            right_index=True, suffixes=suffixes)
            
            operation_details = {
                'operation': 'merge_on_index',
                'merge_type': how,
                'left_shape': left_df.shape,
                'right_shape': right_df.shape,
                'result_shape': result.shape,
                'left_index_name': left_df.index.name,
                'right_index_name': right_df.index.name
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in index merge: {str(e)}")
            return pd.DataFrame()
    
    def cross_join(self, left_df: pd.DataFrame, right_df: pd.DataFrame,
                   suffixes: Tuple[str, str] = ('_left', '_right')) -> pd.DataFrame:
        """
        Perform a cross join (Cartesian product) between two DataFrames.
        
        Returns the Cartesian product of both DataFrames, where each row from the left
        DataFrame is combined with each row from the right DataFrame.
        
        Args:
            left_df (pd.DataFrame): Left DataFrame
            right_df (pd.DataFrame): Right DataFrame
            suffixes (Tuple[str, str]): Suffixes for overlapping column names (default: ('_left', '_right'))
        
        Returns:
            pd.DataFrame: Cross join result
            
        Warning:
            Can produce very large results. Use with caution on large DataFrames.
            
        Example:
            >>> df1 = pd.DataFrame({'A': [1, 2]})
            >>> df2 = pd.DataFrame({'B': ['X', 'Y']})
            >>> result = merger.cross_join(df1, df2)
            >>> print(result)
               A  B
            0  1  X
            1  1  Y
            2  2  X
            3  2  Y
        """
        try:
            # Add temporary key for cross join
            left_temp = left_df.copy()
            right_temp = right_df.copy()
            left_temp['_temp_key'] = 1
            right_temp['_temp_key'] = 1
            
            result = pd.merge(left_temp, right_temp, on='_temp_key', suffixes=suffixes)
            result = result.drop('_temp_key', axis=1)
            
            operation_details = {
                'operation': 'cross_join',
                'left_shape': left_df.shape,
                'right_shape': right_df.shape,
                'result_shape': result.shape,
                'cartesian_product_size': left_df.shape[0] * right_df.shape[0]
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in cross join: {str(e)}")
            return pd.DataFrame()
    
    def merge_with_validation(self, left_df: pd.DataFrame, right_df: pd.DataFrame,
                             on: Union[str, List[str]] = None,
                             how: str = 'inner',
                             validate: str = None,
                             indicator: bool = False,
                             suffixes: Tuple[str, str] = ('_left', '_right')) -> pd.DataFrame:
        """
        Merge DataFrames with validation of merge keys.
        
        Args:
            left_df (pd.DataFrame): Left DataFrame to merge
            right_df (pd.DataFrame): Right DataFrame to merge
            on (Union[str, List[str]]): Column name(s) to join on
            how (str): Type of merge ('inner', 'left', 'right', 'outer') (default: 'inner')
            validate (str): Validation type ('one_to_one', 'one_to_many', 'many_to_one', 'many_to_many')
            indicator (bool): Add a column indicating source of each row (default: False)
            suffixes (Tuple[str, str]): Suffixes for overlapping column names (default: ('_left', '_right'))
        
        Returns:
            pd.DataFrame: Merged DataFrame with validation
            
        Raises:
            ValueError: If validation fails
            
        Example:
            >>> df1 = pd.DataFrame({'id': [1, 2, 3], 'name': ['Alice', 'Bob', 'Charlie']})
            >>> df2 = pd.DataFrame({'id': [1, 2, 4], 'salary': [50000, 60000, 70000]})
            >>> result = merger.merge_with_validation(df1, df2, on='id', validate='one_to_one', indicator=True)
        """
        try:
            result = pd.merge(left_df, right_df, on=on, how=how, validate=validate,
                            indicator=indicator, suffixes=suffixes)
            
            operation_details = {
                'operation': 'merge_with_validation',
                'merge_type': how,
                'validation': validate,
                'left_shape': left_df.shape,
                'right_shape': right_df.shape,
                'result_shape': result.shape,
                'join_keys': on,
                'indicator_added': indicator
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in validated merge: {str(e)}")
            return pd.DataFrame()
    
    def union_dataframes(self, dataframes: List[pd.DataFrame],
                        ignore_index: bool = True,
                        drop_duplicates: bool = True) -> pd.DataFrame:
        """
        Union multiple DataFrames (similar to SQL UNION).
        
        Combines DataFrames vertically and optionally removes duplicates.
        
        Args:
            dataframes (List[pd.DataFrame]): List of DataFrames to union
            ignore_index (bool): Whether to create a new index (default: True)
            drop_duplicates (bool): Whether to remove duplicate rows (default: True)
        
        Returns:
            pd.DataFrame: Union of all DataFrames
            
        Example:
            >>> df1 = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [25, 30]})
            >>> df2 = pd.DataFrame({'name': ['Bob', 'Charlie'], 'age': [30, 35]})
            >>> result = merger.union_dataframes([df1, df2])
            >>> print(result)  # Bob appears only once
                  name  age
            0    Alice   25
            1      Bob   30
            2  Charlie   35
        """
        try:
            if not dataframes:
                raise ValueError("List of DataFrames cannot be empty")
            
            # First concatenate all DataFrames
            result = pd.concat(dataframes, axis=0, ignore_index=ignore_index)
            
            initial_rows = len(result)
            
            # Remove duplicates if requested
            if drop_duplicates:
                result = result.drop_duplicates()
            
            final_rows = len(result)
            duplicates_removed = initial_rows - final_rows
            
            operation_details = {
                'operation': 'union_dataframes',
                'num_dataframes': len(dataframes),
                'input_shapes': [df.shape for df in dataframes],
                'result_shape': result.shape,
                'duplicates_removed': duplicates_removed,
                'drop_duplicates': drop_duplicates
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in union operation: {str(e)}")
            return pd.DataFrame()
    
    def compare_dataframes(self, df1: pd.DataFrame, df2: pd.DataFrame,
                          on: Union[str, List[str]]) -> Dict[str, pd.DataFrame]:
        """
        Compare two DataFrames and return differences.
        
        Args:
            df1 (pd.DataFrame): First DataFrame
            df2 (pd.DataFrame): Second DataFrame
            on (Union[str, List[str]]): Key column(s) to compare on
        
        Returns:
            Dict[str, pd.DataFrame]: Dictionary containing:
                - 'only_in_df1': Rows only in first DataFrame
                - 'only_in_df2': Rows only in second DataFrame
                - 'common': Rows common to both DataFrames
                - 'differences': Rows with same keys but different values
                
        Example:
            >>> df1 = pd.DataFrame({'id': [1, 2, 3], 'value': ['A', 'B', 'C']})
            >>> df2 = pd.DataFrame({'id': [1, 2, 4], 'value': ['A', 'X', 'D']})
            >>> comparison = merger.compare_dataframes(df1, df2, on='id')
            >>> print(len(comparison['only_in_df1']))  # Records only in df1
            1
        """
        try:
            # Perform outer join with indicator
            merged = pd.merge(df1, df2, on=on, how='outer', indicator=True, suffixes=('_df1', '_df2'))
            
            # Split based on indicator
            only_in_df1 = merged[merged['_merge'] == 'left_only'].drop('_merge', axis=1)
            only_in_df2 = merged[merged['_merge'] == 'right_only'].drop('_merge', axis=1)
            both = merged[merged['_merge'] == 'both'].drop('_merge', axis=1)
            
            # Find differences in common rows
            differences = pd.DataFrame()
            if not both.empty:
                # Compare values for rows with same keys
                df1_cols = [col for col in both.columns if col.endswith('_df1')]
                df2_cols = [col for col in both.columns if col.endswith('_df2')]
                
                if df1_cols and df2_cols:
                    # Check for differences
                    diff_mask = pd.Series([False] * len(both))
                    for col1, col2 in zip(df1_cols, df2_cols):
                        if col1.replace('_df1', '') == col2.replace('_df2', ''):
                            diff_mask |= (both[col1] != both[col2])
                    
                    differences = both[diff_mask]
            
            # Common rows (same keys and values)
            common = both[~both.index.isin(differences.index)] if not both.empty else pd.DataFrame()
            
            result = {
                'only_in_df1': only_in_df1,
                'only_in_df2': only_in_df2,
                'common': common,
                'differences': differences
            }
            
            operation_details = {
                'operation': 'compare_dataframes',
                'df1_shape': df1.shape,
                'df2_shape': df2.shape,
                'comparison_key': on,
                'only_in_df1_count': len(only_in_df1),
                'only_in_df2_count': len(only_in_df2),
                'common_count': len(common),
                'differences_count': len(differences)
            }
            
            self._log_operation(operation_details)
            return result
            
        except Exception as e:
            print(f"✗ Error in DataFrame comparison: {str(e)}")
            return {}
    
    def _get_merge_statistics(self) -> Dict[str, Any]:
        """
        Get statistics about the last merge operation.
        
        Returns:
            Dict[str, Any]: Statistics about the last operation performed
            
        Example:
            >>> merger.inner_join(df1, df2, on='id')
            >>> stats = merger._get_merge_statistics()
            >>> print(f"Rows matched: {stats['rows_matched']}")
        """
        return self.last_operation.copy() if self.last_operation else {}
    
    def _get_merge_log(self) -> List[str]:
        """
        Get the complete log of all merge operations.
        
        Returns:
            List[str]: List of all operations performed with timestamps
            
        Example:
            >>> log = merger._get_merge_log()
            >>> for entry in log:
            ...     print(entry)
        """
        return self.merge_log.copy()
    
    def _clear_log(self):
        """
        Clear the merge operations log.
        
        Example:
            >>> merger._clear_log()
            >>> print(len(merger._get_merge_log()))
            0
        """
        self.merge_log = []
        self.last_operation = {}
        print("✓ Merge log cleared")
    
    def _print_merge_summary(self):
        """
        Print a summary of the last merge operation.
        
        Displays key statistics and information about the most recent operation.
        
        Example:
            >>> merger.inner_join(df1, df2, on='id')
            >>> merger._print_merge_summary()
            ========================================
            MERGE OPERATION SUMMARY
            ========================================
            Operation: inner_join
            Left DataFrame: (3, 2)
            Right DataFrame: (3, 2)
            Result: (2, 3)
            Join Keys: id
            Rows Matched: 2
        """
        if not self.last_operation:
            print("No merge operations performed yet.")
            return
        
        print("=" * 40)
        print("MERGE OPERATION SUMMARY")
        print("=" * 40)
        
        op = self.last_operation
        print(f"Operation: {op.get('operation', 'Unknown')}")
        print(f"Left DataFrame: {op.get('left_shape', 'Unknown')}")
        print(f"Right DataFrame: {op.get('right_shape', 'Unknown')}")
        print(f"Result: {op.get('result_shape', 'Unknown')}")
        
        if 'join_keys' in op:
            print(f"Join Keys: {op['join_keys']}")
        if 'rows_matched' in op:
            print(f"Rows Matched: {op['rows_matched']}")
        if 'duplicates_removed' in op:
            print(f"Duplicates Removed: {op['duplicates_removed']}")
        if 'null_values_added' in op:
            print(f"Null Values Added: {op['null_values_added']}")
    
    def _log_operation(self, operation_details: Dict[str, Any]):
        """
        Internal method to log merge operations.
        
        Args:
            operation_details (Dict[str, Any]): Details of the operation to log
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        operation_name = operation_details.get('operation', 'unknown')
        
        log_entry = f"{timestamp} - {operation_name}: {operation_details.get('result_shape', 'unknown shape')}"
        self.merge_log.append(log_entry)
        self.last_operation = operation_details.copy()
# Even simpler version without error checking
    def simple_rename_column(df: pd.DataFrame, old_name: str, new_name: str) -> pd.DataFrame:
        """
        Simple method to rename one column.
        
        Args:
            df (pd.DataFrame): DataFrame to modify
            old_name (str): Current column name
            new_name (str): New column name
            
        Returns:
            pd.DataFrame: DataFrame with renamed column
            
        Example:
            >>> df = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [25, 30]})
            >>> result = simple_rename_column(df, 'name', 'full_name')
            >>> print(result.columns.tolist())
            ['full_name', 'age']
        """
        return df.rename(columns={old_name: new_name})
# Example usage and demonstrations
if __name__ == "__main__":
    """
    Comprehensive examples demonstrating all merge/join/concatenation operations.
    """
    
    # Create sample DataFrames for demonstration
    df_employees = pd.DataFrame({
        'emp_id': [1, 2, 3, 4, 5],
        'name': ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'],
        'department_id': [10, 20, 10, 30, 20]
    })
    
    df_departments = pd.DataFrame({
        'dept_id': [10, 20, 30, 40],
        'dept_name': ['IT', 'HR', 'Finance', 'Marketing'],
        'budget': [100000, 80000, 120000, 90000]
    })
    
    df_salaries = pd.DataFrame({
        'emp_id': [1, 2, 3, 6],
        'salary': [70000, 65000, 75000, 68000],
        'bonus': [5000, 4000, 6000, 4500]
    })
    
    # Initialize merger
    merger = DataFrameMerger()
    
    print("Sample DataFrames:")
    print("\nEmployees:")
    print(df_employees)
    print("\nDepartments:")
    print(df_departments)
    print("\nSalaries:")
    print(df_salaries)
    
    # Demonstrate different join types
    print("\n" + "="*50)
    print("JOIN OPERATIONS")
    print("="*50)
    
    # Inner join
    print("\n1. Inner Join (Employees with Salaries):")
    inner_result = merger.inner_join(df_employees, df_salaries, on='emp_id')
    print(inner_result)
    
    # Left join
    print("\n2. Left Join (All Employees with Departments):")
    left_result = merger.left_join(df_employees, df_departments, 
                                  left_on='department_id', right_on='dept_id')
    print(left_result)
    
    # Concatenation examples
    print("\n" + "="*50)
    print("CONCATENATION OPERATIONS")
    print("="*50)
    
    # Vertical concatenation
    df_new_employees = pd.DataFrame({
        'emp_id': [6, 7],
        'name': ['Frank', 'Grace'],
        'department_id': [10, 30]
    })
    
    print("\n3. Vertical Concatenation:")
    vertical_result = merger.concat_vertical([df_employees, df_new_employees])
    print(vertical_result)
    
    # Union operation
    print("\n4. Union (with duplicate removal):")
    df_duplicates = pd.DataFrame({
        'emp_id': [1, 8],
        'name': ['Alice', 'Henry'],  # Alice is duplicate
        'department_id': [10, 20]
    })
    
    union_result = merger.union_dataframes([df_employees, df_duplicates])
    print(union_result)
    
    # Show operation log
    print("\n" + "="*50)
    print("OPERATION LOG")
    print("="*50)
    
    merger._print_merge_summary()
    
    print("\nComplete log:")
    for entry in merger._get_merge_log():
        print(entry)
# Example usage showing the new return behavior
if __name__ == "__main__":
    """
    Example demonstrating the modified extract methods that return DataFrames:
    """
    # Create ETL instance
    etl = SimpleETL()
    
    # Extract methods now return DataFrames
    
    # Example 1: Extract sample data
    df = etl.extract_sample_data(100)
    print(f"Sample data shape: {df.shape}")
    print(f"Columns: {df.columns.tolist()}")
    
    # Example 2: Extract from CSV (if file exists)
    # df_csv = etl.extract_from_csv('data.csv')
    # if df_csv is not None:
    #     print(f"CSV data shape: {df_csv.shape}")
    
    # Example 3: Extract from multiple sources and work with the DataFrames
    # df1 = etl.extract_sample_data(50)
    # df2 = etl.extract_sample_data(75)
    # 
    # # You can now work directly with the returned DataFrames
    # combined = pd.concat([df1, df2], ignore_index=True)
    # print(f"Combined shape: {combined.shape}")
    
    # Example 4: Extract and immediately analyze without storing in class
    # df_temp = etl.extract_from_clipboard()
    # if df_temp is not None:
    #     print(df_temp.describe())
    
    # The class still maintains internal state for transform/load operations
    etl.clean_missing_values('fill')
    current_data = etl.get_current_data()
    print(f"Current data in class: {current_data.shape if current_data is not None else 'None'}")