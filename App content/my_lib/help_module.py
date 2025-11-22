import ast

class Converting:
    def __init__(self):
        
        pass
    def convert_string_to_object(self,text):
    
        if isinstance(text, str):
            try:
                text = ast.literal_eval(text)  # Safely parse string to dict
            except (ValueError, SyntaxError):
        
                text = None
                
        return text