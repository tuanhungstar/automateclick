from my_lib.shared_context import ExecutionContext as Context

class Global_Variable:
    """A module for configuring global variables during task execution."""
    
    def __init__(self, context: Context):
        """Initializes the Global_Variable class."""
        self.context = context

    def set_global_variable(self, global_variable_list_name: str, value: any):
        """
        Sets a global variable by name.

        Args:
            global_variable_list_name (str): The name of the global variable to set.
            value (any): The value to assign to the variable.
            
        Returns:
            any: The value that was set.
        """
        if not global_variable_list_name or global_variable_list_name == "-- Select Variable --":
            if self.context:
                self.context.add_log("Error: Valid variable name not selected.")
            return None
            
        if self.context:
            self.context.set_variable(global_variable_list_name, value)
            # self.context.add_log(f"Global variable '{global_variable_list_name}' set to '{value}'.")
        return value

    def get_global_variable(self, global_variable_list_name: str):
        """
        Gets the value of a global variable by name.

        Args:
            global_variable_list_name (str): The name of the global variable to retrieve.

        Returns:
            any: The value of the global variable, or None if it doesn't exist.
        """
        if not global_variable_list_name or global_variable_list_name == "-- Select Variable --":
            if self.context:
                self.context.add_log("Error: Valid variable name not selected.")
            return None
            
    def preset_multiple_global_variables(self, dynamic_variables_list=None):
        """
        Sets values for an arbitrary number of global variables at once.

        Args:
            dynamic_variables_list (list or dict): Mapping or list of global variable names to values.
            
        Returns:
            bool: True when complete.
        """
        if not self.context:
            return False
            
        if not dynamic_variables_list:
            self.context.add_log("Error: Invalid or empty dynamic variable list provided.")
            return False
            
        count = 0
        
        items_to_set = []
        if isinstance(dynamic_variables_list, dict):
            items_to_set = list(dynamic_variables_list.items())
        elif isinstance(dynamic_variables_list, list):
            for item in dynamic_variables_list:
                if isinstance(item, dict) and "var" in item and "val" in item:
                    items_to_set.append((item["var"], item["val"]))
        
        for name, val in items_to_set:
            if name and name != "-- Select Variable --":
                self.context.set_variable(name, val)
                count += 1
                
        self.context.add_log(f"Successfully preset {count} global variables dynamically.")
        return True
