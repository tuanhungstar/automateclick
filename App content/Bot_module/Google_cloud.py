import urllib.request
import urllib.parse
import json
from typing import Dict, Any, Optional
from my_lib.shared_context import ExecutionContext as Context
import ast
import pandas as pd
class Google_cloud:
    def __init__(self, context: Context):
        self.context = context # Store the shared context
        self.log_prefix = f"[{self.__class__.__name__}]"
        self.context.add_log(f"{self.log_prefix} Initialized with Google_cloud.")
        pass

    def search_google(self, product_name,api_key,search_engine_id,proxy_config:(Dict[str, Any]) =None) ->str:
        """
        proxy = {'https': 'rb-proxy-de.bosch.com:8080'}
        results = Google_cloud.search_google("laptop", proxy_config=proxy)
        api_key = "AIzaSyDcI4Dya15MwEeprPsM28PHbKq_eBZwzB8"
        search_engine_id = "024aceea33b694611"       
        
        """

        query = urllib.parse.quote(product_name)
        
        url = f"https://www.googleapis.com/customsearch/v1?q={query}&key={api_key}&cx={search_engine_id}"
        if isinstance(proxy_config, str):
            try:
                proxy_config = ast.literal_eval(proxy_config)  # Safely parse string to dict
            except (ValueError, SyntaxError):
                print("Invalid proxy_config string format")
                proxy_config = None
                
                
        # Set up opener based on proxy configuration
        if proxy_config:
            proxy_handler = urllib.request.ProxyHandler(proxy_config)
            opener = urllib.request.build_opener(proxy_handler)
        else:
            # No proxy
            opener = urllib.request.build_opener()
        
        try:
            # Make the request
            request = urllib.request.Request(url)
            response = opener.open(request)
            
            # Read and decode the response
            data = response.read().decode('utf-8')
            results = json.loads(data)
            search_results = results['items']

            # Create a list to store the processed data
            processed_data = []

            for item in search_results:
                processed_data.append({
                    'title': item.get('title', ''),
                    'link': item.get('link', ''),
                    'snippet': item.get('snippet', ''),
                    'displayLink': item.get('displayLink', ''),
                    'formattedUrl': item.get('formattedUrl', '')
                })

            # Create DataFrame
            df = pd.DataFrame(processed_data)
            return df
            
        except urllib.error.URLError as e:
            print(f"Error occurred: {e}")
            return None

