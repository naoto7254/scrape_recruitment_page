import requests
from lxml import html
import re

class Scraper:
    def __init__(self, urls_array, elements):
        self.urls = urls_array
        self.elements = elements
        
    def scrape_urls(self):
        all_result = []
        keys_list = list(self.elements[0].keys())
        xpath_list = [value[0] if isinstance(value, tuple) else value for value in self.elements[0].values()]

        all_result.append(keys_list)
        
        for url in self.urls:
            response = requests.get(url)
            response.raise_for_status()
            
            tree = html.fromstring(response.content)
            
            result = []
            for xpath in xpath_list:
                element = tree.xpath(xpath)
                result.append(element[0].text_content().strip() if element else None)
            
            all_result.append(result)
            
        
        return all_result
    